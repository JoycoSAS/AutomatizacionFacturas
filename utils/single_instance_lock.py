import os
import time

class SingleInstanceLock:
    """
    Lock por archivo (Windows-friendly).
    Evita que el Programador de tareas dispare otra instancia mientras una sigue corriendo.
    """
    def __init__(self, lock_path: str, ttl_seconds: int = 1800):
        self.lock_path = lock_path
        self.ttl_seconds = ttl_seconds
        self.acquired = False

    def acquire(self) -> bool:
        os.makedirs(os.path.dirname(self.lock_path), exist_ok=True)

        # si existe, revisa TTL
        if os.path.exists(self.lock_path):
            try:
                mtime = os.path.getmtime(self.lock_path)
                age = time.time() - mtime
                if age > self.ttl_seconds:
                    # lock viejo → se considera muerto
                    try:
                        os.remove(self.lock_path)
                    except Exception:
                        return False
                else:
                    return False
            except Exception:
                return False

        try:
            with open(self.lock_path, "w", encoding="utf-8") as f:
                f.write(str(os.getpid()))
            self.acquired = True
            return True
        except Exception:
            return False

    def release(self):
        if not self.acquired:
            return
        try:
            os.remove(self.lock_path)
        except Exception:
            pass
        self.acquired = False