"""Answers to "Is this the same company?", kept in one JSON file.

{"<name as written in the Excel>": "<customer code>"}  same company as that customer
{"<name as written in the Excel>": null}               a new company, stop asking
{"__pairs__": [{"a": name, "b": name, "same": bool}]}  two spellings: one company or two

The file path comes from COMPANY_DECISIONS_FILE (point it at shared storage when the importer
runs on a server). Everything else goes through DecisionStore, so a CRM-backed store can replace
the file later without touching the matcher. Writes are serialised and re-read the file first, so
two answers arriving together (the web server is threaded) both survive.
"""
from __future__ import annotations

import json
import os
import tempfile
import threading
from pathlib import Path

ENV_VAR = "COMPANY_DECISIONS_FILE"
PAIRS = "__pairs__"
_LOCK = threading.RLock()


def default_path(importer_dir) -> Path:
    return Path(os.environ.get(ENV_VAR) or Path(importer_dir) / "excel" / "company_decisions.json")


def _key(name) -> str:
    from .companies import norm
    return norm(name) or " ".join(str(name).split()).lower()


class DecisionStore:
    def __init__(self, path):
        self.path = Path(path)
        self._data = self._read()

    def _read(self) -> dict:
        if not self.path.exists():
            return {}
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                raise ValueError("top level is not an object")
        except ValueError as exc:  # JSONDecodeError is a ValueError
            raise ValueError(f"{self.path}: unreadable company decisions file ({exc}); "
                             "fix or remove it, answers are not silently dropped") from exc
        return data

    def as_dict(self) -> dict:
        """The "same as customer X" / "new company" answers (pair answers are kept apart)."""
        return {k: v for k, v in self._data.items() if k != PAIRS}

    @staticmethod
    def _pairs(data: dict) -> list[dict]:
        raw = data.get(PAIRS, [])
        if not isinstance(raw, list):
            return []
        return [p for p in raw if isinstance(p, dict) and isinstance(p.get("a"), str) and isinstance(p.get("b"), str)]

    def pair_decision(self, a: str, b: str):
        """True = same company, False = different companies, None = nobody answered yet."""
        key = frozenset((_key(a), _key(b)))
        for p in self._pairs(self._data):
            if frozenset((_key(p["a"]), _key(p["b"]))) == key:
                return bool(p.get("same"))
        return None

    def set_pair(self, a: str, b: str, same: bool) -> None:
        key = frozenset((_key(a), _key(b)))
        with _LOCK:
            data = self._read()  # another writer may have saved since this store was opened
            kept = [p for p in self._pairs(data) if frozenset((_key(p["a"]), _key(p["b"]))) != key]
            data[PAIRS] = kept + [{"a": a, "b": b, "same": bool(same)}]
            self._save(data)

    def set_same(self, name: str, code: str) -> None:
        self._set(name, code)

    def set_new(self, name: str) -> None:
        self._set(name, None)

    def _set(self, name: str, code) -> None:
        if name == PAIRS:
            raise ValueError(f"{PAIRS!r} is reserved and cannot be a company name")
        with _LOCK:
            data = self._read()
            data[name] = code
            self._save(data)

    def _save(self, data: dict) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, tmp = tempfile.mkstemp(dir=self.path.parent, prefix=self.path.name + ".", suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=1, sort_keys=True)
            os.replace(tmp, self.path)  # a crash never leaves a half-written file
        except BaseException:
            try:
                os.unlink(tmp)
            except OSError:
                pass
            raise
        self._data = data
