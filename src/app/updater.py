from __future__ import annotations

import dataclasses
import hashlib
import os
import re
import subprocess
import tempfile
from typing import Callable, Optional

import requests


def parse_version(tag: str) -> tuple[int, int, int]:
    tag = tag.strip().lstrip("vV")
    m = re.match(r"^(\d+)\.(\d+)\.(\d+)$", tag)
    if not m:
        return (0, 0, 0)
    return (int(m.group(1)), int(m.group(2)), int(m.group(3)))


@dataclasses.dataclass(frozen=True)
class UpdateInfo:
    latest_tag: str
    latest_version: tuple[int, int, int]
    current_version: tuple[int, int, int]
    notes: str
    installer_url: str
    sha256_url: str


class GitHubReleaseUpdater:
    """
    Updater robusto:
    - controlla latest release
    - scarica sha256 + installer
    - verifica hash
    - lancia installer e chiude app (il chiamante deve poi chiudersi)
    """

    def __init__(self, repo_slug: str, installer_asset_name: str, sha256_asset_name: str, timeout_s: int = 20):
        self.repo_slug = repo_slug
        self.installer_asset_name = installer_asset_name
        self.sha256_asset_name = sha256_asset_name
        self.timeout_s = timeout_s

    def _latest_release_json(self) -> dict:
        url = f"https://api.github.com/repos/{self.repo_slug}/releases/latest"
        r = requests.get(url, timeout=self.timeout_s)
        # Private repo: qui tipicamente torna 404 senza token
        r.raise_for_status()
        return r.json()

    @staticmethod
    def _pick_asset_url(release_json: dict, name: str) -> str:
        for a in release_json.get("assets", []):
            if a.get("name") == name:
                return a.get("browser_download_url")
        raise RuntimeError(f"Asset non trovato nella release: {name}")

    def check(self, current_version_str: str) -> Optional[UpdateInfo]:
        release = self._latest_release_json()
        tag = release.get("tag_name") or ""
        notes = release.get("body") or ""

        latest_v = parse_version(tag)
        curr_v = parse_version(current_version_str)

        if latest_v <= curr_v:
            return None

        installer_url = self._pick_asset_url(release, self.installer_asset_name)
        sha256_url = self._pick_asset_url(release, self.sha256_asset_name)

        return UpdateInfo(
            latest_tag=tag,
            latest_version=latest_v,
            current_version=curr_v,
            notes=notes,
            installer_url=installer_url,
            sha256_url=sha256_url,
        )

    @staticmethod
    def _sha256_file(path: str) -> str:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return h.hexdigest()

    @staticmethod
    def _download(url: str, dst_path: str, progress_cb: Optional[Callable[[int, int], None]] = None) -> None:
        with requests.get(url, stream=True, timeout=60) as r:
            r.raise_for_status()
            total = int(r.headers.get("Content-Length") or 0)
            done = 0
            with open(dst_path, "wb") as f:
                for b in r.iter_content(chunk_size=1024 * 1024):
                    if not b:
                        continue
                    f.write(b)
                    done += len(b)
                    if progress_cb:
                        progress_cb(done, total)

    @staticmethod
    def _read_sha256_from_file(path: str) -> str:
        """
        Atteso formato: "<sha256>  <filename>" oppure solo "<sha256>"
        """
        txt = open(path, "r", encoding="utf-8", errors="replace").read().strip()
        if not txt:
            raise RuntimeError("File sha256 vuoto")
        first = txt.splitlines()[0].strip()
        token = first.split()[0].strip()
        if not re.fullmatch(r"[0-9a-fA-F]{64}", token):
            raise RuntimeError("Formato sha256 non valido")
        return token.lower()

    def download_and_verify(
        self,
        info: UpdateInfo,
        progress_cb: Optional[Callable[[str, int, int], None]] = None,
    ) -> str:
        """
        Ritorna path dell'installer scaricato e verificato.
        """
        tmpdir = tempfile.mkdtemp(prefix="vainieri_update_")
        sha_path = os.path.join(tmpdir, os.path.basename(self.sha256_asset_name))
        exe_path = os.path.join(tmpdir, os.path.basename(self.installer_asset_name))

        if progress_cb:
            progress_cb("Scarico firma SHA256...", 0, 0)
        self._download(info.sha256_url, sha_path, None)
        expected = self._read_sha256_from_file(sha_path)

        if progress_cb:
            progress_cb("Scarico installer...", 0, 0)

        def _pcb(done: int, total: int):
            if progress_cb:
                progress_cb("Scarico installer...", done, total)

        self._download(info.installer_url, exe_path, _pcb)

        actual = self._sha256_file(exe_path)
        if actual.lower() != expected.lower():
            raise RuntimeError("Verifica SHA256 fallita: installer corrotto o manomesso.")

        return exe_path

    @staticmethod
    def run_installer(installer_path: str, silent: bool = False) -> None:
        args = [installer_path]
        if silent:
            args += ["/VERYSILENT", "/NORESTART", "/CLOSEAPPLICATIONS", "/RESTARTAPPLICATIONS"]

        # Avvia e torna subito: l'app chiamante deve chiudersi dopo
        subprocess.Popen(args, close_fds=True)
