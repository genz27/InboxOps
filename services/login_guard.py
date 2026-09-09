from __future__ import annotations

from dataclasses import dataclass, field
from threading import Lock
from time import monotonic

from services.outlook_manager import MailboxError


@dataclass(slots=True)
class LoginGuard:
    max_failures: int = 5
    window_seconds: float = 900
    lockout_seconds: float = 900
    _lock: Lock = field(default_factory=Lock, init=False, repr=False)
    _failures: dict[str, list[float]] = field(default_factory=dict, init=False, repr=False)
    _locked_until: dict[str, float] = field(default_factory=dict, init=False, repr=False)

    def assert_allowed(self, key: str) -> None:
        now = monotonic()
        with self._lock:
            locked_until = self._locked_until.get(key, 0)
            if locked_until > now:
                retry_after = int(locked_until - now) + 1
                raise MailboxError(
                    f"登录失败次数过多，请 {retry_after} 秒后重试",
                    code="login_rate_limited",
                    status_code=429,
                )
            if locked_until:
                self._locked_until.pop(key, None)
            self._prune_failures(key, now)

    def record_failure(self, key: str) -> None:
        now = monotonic()
        with self._lock:
            stamps = self._prune_failures(key, now)
            stamps.append(now)
            self._failures[key] = stamps
            if len(stamps) >= self.max_failures:
                self._locked_until[key] = now + self.lockout_seconds
                self._failures.pop(key, None)

    def record_success(self, key: str) -> None:
        with self._lock:
            self._failures.pop(key, None)
            self._locked_until.pop(key, None)

    def _prune_failures(self, key: str, now: float) -> list[float]:
        cutoff = now - self.window_seconds
        stamps = [stamp for stamp in self._failures.get(key, []) if stamp >= cutoff]
        if stamps:
            self._failures[key] = stamps
        else:
            self._failures.pop(key, None)
        return stamps
