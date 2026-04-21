"""LibreOffice UNO bridge.

Starts a headless soffice process that accepts UNO connections over a
socket, and resolves a component context the rest of the app can use.

Design notes:
  - One long-lived soffice per pipeline run, torn down on exit.
  - Every run gets an isolated UserInstallation so concurrent runs and
    a running desktop LibreOffice don't collide on profile locks.
  - The port is chosen at random to avoid conflicts.
"""

from __future__ import annotations

import logging
import os
import random
import shutil
import socket
import subprocess
import tempfile
import time
from pathlib import Path
from typing import Any
import uno  # type: ignore
from com.sun.star.beans import PropertyValue  # type: ignore
from com.sun.star.connection import NoConnectException  # type: ignore

log = logging.getLogger(__name__)


def make_prop(name: str, value: Any) -> PropertyValue:
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def props(**kwargs: Any) -> tuple:
    return tuple(make_prop(k, v) for k, v in kwargs.items())


def path_to_url(path: str | Path) -> str:
    """Convert a filesystem path to a `file://` URL.

    Uses `Path.as_uri()` so Windows paths come out correctly as
    `file:///C:/path/to/dir` (three slashes, forward slashes, drive
    letter un-escaped). Our own `quote()`-based version produced
    `file://C%3A%5C...` which LibreOffice silently rejects for
    `-env:UserInstallation=...`, causing it to fall back to the
    system profile — where a corrupt `bootstrap.ini` can surface.
    """
    p = Path(path).resolve()
    return p.as_uri()


def _find_free_port() -> int:
    for _ in range(20):
        port = random.randint(20000, 60000)
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    raise RuntimeError("no free port")


def _find_soffice() -> str:
    """Locate the soffice executable.

    `shutil.which` handles $PATH (and %PATH% on Windows), but a user who
    launches us via LibreOffice's bundled Python without the Windows
    launcher hasn't necessarily added `C:\\Program Files\\LibreOffice\\program`
    to PATH. Check a few known install locations so we still work.
    """
    for name in ("soffice", "libreoffice", "soffice.exe"):
        hit = shutil.which(name)
        if hit:
            return hit
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
        ]
        for c in candidates:
            if os.path.isfile(c):
                return c
    return "soffice"


class UnoBridge:
    """Owns the soffice subprocess and the UNO desktop handle."""

    def __init__(self, soffice: str = "soffice") -> None:
        explicit = shutil.which(soffice)
        self.soffice_bin = explicit or _find_soffice()
        self.port = _find_free_port()
        self.profile_dir = Path(tempfile.mkdtemp(prefix="pdfua-profile-"))
        self.proc: subprocess.Popen | None = None
        self.ctx = None
        self.desktop = None
        # Capture soffice stderr so that a silent crash (corrupt bootstrap.ini,
        # missing profile, licence dialog, etc.) surfaces in our error message
        # instead of leaving the pipeline hanging forever.
        self._stderr_path = self.profile_dir / "soffice.stderr.log"
        self._stdout_path = self.profile_dir / "soffice.stdout.log"
        self._stderr_fh = None
        self._stdout_fh = None

    def __enter__(self) -> "UnoBridge":
        self.start()
        return self

    def __exit__(self, *_: Any) -> None:
        self.stop()

    def start(self) -> None:
        accept = f"socket,host=127.0.0.1,port={self.port};urp;"
        user_url = path_to_url(self.profile_dir)
        cmd = [
            self.soffice_bin,
            "--headless",
            "--invisible",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            "--nodefault",
            "--nolockcheck",
            f"--accept={accept}",
            f"-env:UserInstallation={user_url}",
        ]
        log.info("starting soffice on port %d", self.port)
        log.info("soffice bin: %s", self.soffice_bin)
        log.info("profile dir: %s", self.profile_dir)
        self._stdout_fh = open(self._stdout_path, "wb")
        self._stderr_fh = open(self._stderr_path, "wb")
        self.proc = subprocess.Popen(
            cmd,
            stdout=self._stdout_fh,
            stderr=self._stderr_fh,
            env={**os.environ, "HOME": str(self.profile_dir)},
        )
        self._connect()

    def _read_soffice_log_tail(self, limit: int = 4000) -> str:
        parts: list[str] = []
        for label, path in (("stderr", self._stderr_path), ("stdout", self._stdout_path)):
            try:
                data = path.read_bytes()
            except Exception:
                continue
            if not data:
                continue
            text = data.decode("utf-8", errors="replace")
            if len(text) > limit:
                text = "…" + text[-limit:]
            parts.append(f"--- soffice {label} ---\n{text.rstrip()}")
        return "\n\n".join(parts) if parts else "(no soffice output captured)"

    def _connect(self, timeout: float = 45.0) -> None:
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        conn_str = (
            f"uno:socket,host=127.0.0.1,port={self.port};"
            f"urp;StarOffice.ComponentContext"
        )
        deadline = time.time() + timeout
        last_err: Exception | None = None
        while time.time() < deadline:
            if self.proc and self.proc.poll() is not None:
                tail = self._read_soffice_log_tail()
                raise RuntimeError(
                    f"soffice exited early (rc={self.proc.returncode}) "
                    f"before accepting UNO connections.\n{tail}"
                )
            try:
                self.ctx = resolver.resolve(conn_str)
                break
            except NoConnectException as e:
                last_err = e
                time.sleep(0.25)
            except Exception as e:
                last_err = e
                time.sleep(0.25)
        if self.ctx is None:
            tail = self._read_soffice_log_tail()
            raise RuntimeError(
                f"could not connect to soffice within {timeout:.0f}s "
                f"(last error: {last_err}).\n{tail}"
            )
        smgr = self.ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )
        log.info("UNO connected")

    def stop(self) -> None:
        try:
            if self.desktop is not None:
                try:
                    self.desktop.terminate()
                except Exception:
                    pass
        finally:
            if self.proc is not None:
                try:
                    self.proc.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    self.proc.kill()
                    self.proc.wait(timeout=5)
            for fh in (self._stdout_fh, self._stderr_fh):
                try:
                    if fh is not None:
                        fh.close()
                except Exception:
                    pass
            shutil.rmtree(self.profile_dir, ignore_errors=True)

    def load(self, path: str | Path, hidden: bool = True, **filter_props: Any):
        p = Path(path).resolve()
        # Clean up stale .~lock.<name># that a crashed soffice may leave behind.
        lock = p.with_name(f".~lock.{p.name}#")
        if lock.exists():
            try:
                lock.unlink()
            except Exception:
                pass
        url = path_to_url(p)
        base = {"Hidden": hidden, "ReadOnly": False}
        base.update(filter_props)
        doc = self.desktop.loadComponentFromURL(url, "_blank", 0, props(**base))
        if doc is None:
            raise RuntimeError(f"loadComponentFromURL returned None for {p}")
        return doc

    def new_writer(self):
        return self.desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, props(Hidden=True)
        )

    def save_as(self, doc, path: str | Path, filter_name: str, **filter_data: Any) -> None:
        url = path_to_url(path)
        if filter_data:
            fd = tuple(make_prop(k, v) for k, v in filter_data.items())
            p = props(FilterName=filter_name, Overwrite=True)
            p += (make_prop("FilterData", uno.Any("[]com.sun.star.beans.PropertyValue", fd)),)
        else:
            p = props(FilterName=filter_name, Overwrite=True)
        doc.storeToURL(url, p)
