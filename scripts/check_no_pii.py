"""Fail if a tracked file looks like it contains real customer or client data.

Checks every file git tracks (or every file under the repo when git is not
available) for:
- email addresses outside the reserved example.com domain
- IPv4 addresses outside the documentation range 203.0.113.0/24
- API-key shaped strings
- client brand names, compared by SHA-256 so the names never appear here

Run: python scripts/check_no_pii.py
"""
from __future__ import annotations

import hashlib
import re
import subprocess
import sys
import unicodedata
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKIP_DIRS = {".git", ".venv", "venv", "__pycache__", ".pytest_cache"}
TEXT_SUFFIXES = {".py", ".md", ".csv", ".json", ".txt", ".html", ".toml", ".yml", ".yaml", ".cfg", ".example"}

BLOCKED_NAME_HASHES = {
    "451a99b9470331a54c7becaea605c3c8fe1df9b0185ced70ea3864ce3ffb5207",
    "326279d6b85281a947f5c4b3484439628762bb2e7d0d8f014c1f750ec4b296f1",
    "b0fbddb4123cbb2344fa3cd6d78588ff6d24914059bb2175aacaa0c40ae36c68",
    "53ea88be007bfb105b2d12ea7d3ace0de2bbbe9e545498c0b965209126c3c5ad",
    "a80f63fffc27306769f100f20979b4ca69fc5b93542831ebd4c654610b8d596a",
    "1bb2f908013368b28cf6e3520b95ff380cb2948f07abdbe9b7aadfae47fe16ae",
    "9746e385a7718b17763b74d7af039fa254facfc5c6af51e82493320289264429",
    "87e1b086b3f52d91e04e9216438113ecde867c986ae69503f8b03402702f4e77",
    "fbcf1ca078ec30468512237bb42455b9b78f845b977cdc5a0ce764b4399971ce",
}

EMAIL = re.compile(r"[\w.+-]+@([\w-]+\.)+[a-z]{2,}", re.I)
ALLOWED_EMAIL = re.compile(r"@(example\.(com|org|net)|users\.noreply\.github\.com)$|^noreply@anthropic\.com$", re.I)
IPV4 = re.compile(r"\b(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})\b")
SECRET = re.compile(r"sk-(proj-|ant-)?[A-Za-z0-9_-]{20,}|esecret_[A-Za-z0-9]{10,}|ghp_[A-Za-z0-9]{30,}|AKIA[0-9A-Z]{16}")
WORD = re.compile(r"[a-z0-9]+")


def tracked_files() -> list[Path]:
    try:
        out = subprocess.run(["git", "ls-files", "-co", "--exclude-standard"], cwd=ROOT,
                             capture_output=True, text=True, check=True).stdout
        paths = [ROOT / p for p in out.splitlines()]
    except (OSError, subprocess.CalledProcessError):
        paths = [p for p in ROOT.rglob("*") if p.is_file()]
    return [p for p in paths if p.is_file() and not SKIP_DIRS & set(p.relative_to(ROOT).parts)]


def read_text(path: Path) -> str | None:
    if path.suffix.lower() == ".xlsx":
        with zipfile.ZipFile(path) as zf:
            return " ".join(zf.read(n).decode("utf-8", "ignore") for n in zf.namelist() if n.endswith(".xml"))
    if path.suffix.lower() in TEXT_SUFFIXES or path.name in {"LICENSE", ".gitignore"}:
        return path.read_text(encoding="utf-8", errors="ignore")
    return None


def fold(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text.lower()) if not unicodedata.combining(c))


def problems(path: Path, text: str) -> list[str]:
    found = []
    for m in EMAIL.finditer(text):
        if not ALLOWED_EMAIL.search(m.group(0)):
            found.append(f"email address ({m.group(0)[:3]}...)")
    for m in IPV4.finditer(text):
        octets = [int(x) for x in m.groups()]
        if all(o <= 255 for o in octets) and octets[:3] != [203, 0, 113] and octets != [127, 0, 0, 1] and octets[0] != 0:
            # version strings like 3.1.1.0 are rare; flag and let a human decide
            found.append(f"IPv4 address ({octets[0]}.x.x.x)")
    if SECRET.search(text):
        found.append("API-key shaped string")
    words = set(WORD.findall(fold(text)))
    words |= {a + b for a, b in zip(WORD.findall(fold(text)), WORD.findall(fold(text))[1:])}
    hits = {w for w in words if hashlib.sha256(w.encode()).hexdigest() in BLOCKED_NAME_HASHES}
    if hits:
        found.append(f"{len(hits)} blocked client name(s)")
    return found


def main() -> int:
    failures = 0
    for path in tracked_files():
        if path.name == Path(__file__).name:
            continue
        text = read_text(path)
        if text is None:
            continue
        for issue in problems(path, text):
            failures += 1
            print(f"[pii] {path.relative_to(ROOT)}: {issue}")
    if failures:
        print(f"[pii] {failures} problem(s). Real client data must stay out of this repository.")
        return 1
    print("[pii] ok: no customer or client data found in tracked files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
