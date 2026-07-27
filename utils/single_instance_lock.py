import errno
import os
import time
from typing import Optional


class SingleInstanceLock:
    """
    Lock por archivo compatible con Windows y Linux.

    Reglas:
    - La creación del lock es atómica para evitar que dos procesos lo adquieran.
    - Si el PID registrado sigue activo, el lock se conserva aunque supere el TTL.
    - Si el PID ya no existe, el lock se considera huérfano y puede retirarse.
    - Si el contenido no contiene un PID válido, solo se retira al superar el TTL.
    - release() únicamente elimina el lock si todavía pertenece a este proceso.
    """

    def __init__(self, lock_path: str, ttl_seconds: int = 1800):
        self.lock_path = os.path.abspath(lock_path)
        self.ttl_seconds = max(0, int(ttl_seconds))
        self.pid = os.getpid()
        self.acquired = False

    def _read_pid(self) -> Optional[int]:
        try:
            with open(self.lock_path, "r", encoding="utf-8") as file:
                raw = file.read().strip()

            pid = int(raw)
            return pid if pid > 0 else None
        except (OSError, TypeError, ValueError):
            return None

    @staticmethod
    def _pid_is_alive_posix(pid: int) -> bool:
        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
        except OSError as exc:
            if exc.errno == errno.ESRCH:
                return False
            if exc.errno == errno.EPERM:
                return True
            return False

    @staticmethod
    def _pid_is_alive_windows(pid: int) -> bool:
        try:
            import ctypes
            from ctypes import wintypes

            process_query_limited_information = 0x1000
            synchronize = 0x00100000
            wait_timeout = 0x00000102

            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.OpenProcess.argtypes = [
                wintypes.DWORD,
                wintypes.BOOL,
                wintypes.DWORD,
            ]
            kernel32.OpenProcess.restype = wintypes.HANDLE
            kernel32.WaitForSingleObject.argtypes = [
                wintypes.HANDLE,
                wintypes.DWORD,
            ]
            kernel32.WaitForSingleObject.restype = wintypes.DWORD
            kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
            kernel32.CloseHandle.restype = wintypes.BOOL

            handle = kernel32.OpenProcess(
                process_query_limited_information | synchronize,
                False,
                pid,
            )

            if not handle:
                # ERROR_ACCESS_DENIED: el proceso existe, pero no puede consultarse.
                return ctypes.get_last_error() == 5

            try:
                return kernel32.WaitForSingleObject(handle, 0) == wait_timeout
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return False

    def _pid_is_alive(self, pid: Optional[int]) -> bool:
        if pid is None:
            return False

        if pid == self.pid:
            return True

        if os.name == "nt":
            return self._pid_is_alive_windows(pid)

        return self._pid_is_alive_posix(pid)

    @staticmethod
    def _same_snapshot(before: os.stat_result, after: os.stat_result) -> bool:
        inode_available = bool(getattr(before, "st_ino", 0)) and bool(
            getattr(after, "st_ino", 0)
        )

        if inode_available:
            return (
                before.st_dev == after.st_dev
                and before.st_ino == after.st_ino
            )

        return (
            before.st_mtime_ns == after.st_mtime_ns
            and before.st_size == after.st_size
        )

    def _remove_if_stale(self) -> bool:
        try:
            snapshot = os.stat(self.lock_path)
        except FileNotFoundError:
            return True
        except OSError:
            return False

        existing_pid = self._read_pid()
        age_seconds = max(0.0, time.time() - snapshot.st_mtime)

        if existing_pid is not None:
            if self._pid_is_alive(existing_pid):
                return False
        elif age_seconds <= self.ttl_seconds:
            return False

        try:
            current = os.stat(self.lock_path)
        except FileNotFoundError:
            return True
        except OSError:
            return False

        if not self._same_snapshot(snapshot, current):
            return False

        try:
            os.remove(self.lock_path)
            return True
        except FileNotFoundError:
            return True
        except OSError:
            return False

    def _create_atomic(self) -> bool:
        flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY

        try:
            fd = os.open(self.lock_path, flags, 0o640)
        except FileExistsError:
            return False
        except OSError:
            return False

        try:
            with os.fdopen(fd, "w", encoding="utf-8") as file:
                file.write(f"{self.pid}\n")
                file.flush()
                os.fsync(file.fileno())
        except Exception:
            try:
                os.remove(self.lock_path)
            except OSError:
                pass
            return False

        self.acquired = True
        return True

    def acquire(self) -> bool:
        if self.acquired:
            return True

        parent = os.path.dirname(self.lock_path) or "."
        try:
            os.makedirs(parent, exist_ok=True)
        except OSError:
            return False

        # Dos intentos permiten retirar un lock huérfano y crear el nuevo
        # sin abrir una ventana de carrera entre procesos.
        for _ in range(2):
            if self._create_atomic():
                return True

            if not self._remove_if_stale():
                return False

        return self._create_atomic()

    def release(self) -> None:
        if not self.acquired:
            return

        try:
            if self._read_pid() == self.pid:
                os.remove(self.lock_path)
        except FileNotFoundError:
            pass
        except OSError:
            pass
        finally:
            self.acquired = False

    def __enter__(self) -> "SingleInstanceLock":
        if not self.acquire():
            raise RuntimeError(
                f"No fue posible adquirir el lock: {self.lock_path}"
            )
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> bool:
        self.release()
        return False
