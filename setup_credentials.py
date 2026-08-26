from __future__ import annotations

import json
import os
import re
import subprocess
from pathlib import Path

APP_DIR = Path.home() / ".doc24_sender"
CONFIG_PATH = APP_DIR / "config.json"


def extract_credentials(source: str) -> tuple[str, str] | None:
    id_patterns = [
        r'page\.fill\(\s*["\']#id["\']\s*,\s*["\']([^"\']+)["\']\s*\)',
        r'locator\(\s*["\']#id["\']\s*\)\.fill\(\s*["\']([^"\']+)["\']\s*\)',
    ]
    pw_patterns = [
        r'page\.keyboard\.type\(\s*["\']([^"\']+)["\']',
        r'locator\(\s*["\']#password["\']\s*\)\.fill\(\s*["\']([^"\']+)["\']\s*\)',
        r'page\.fill\(\s*["\']#password["\']\s*,\s*["\']([^"\']+)["\']\s*\)',
    ]

    username = next(
        (match.group(1) for pattern in id_patterns if (match := re.search(pattern, source))),
        None,
    )
    password = next(
        (match.group(1) for pattern in pw_patterns if (match := re.search(pattern, source))),
        None,
    )

    if username and password:
        return username, password
    return None


def legacy_sources() -> list[tuple[str, str]]:
    sources: list[tuple[str, str]] = []
    candidates = [
        Path.home() / "Desktop" / "doc24-main-backup.py",
        Path.home() / "Desktop" / "doc24-autosender" / "legacy_main.py",
        Path.home() / "Desktop" / "doc24-autosender" / "main_old.py",
        Path("legacy_main.py"),
        Path("main_old.py"),
        Path("old_main.py"),
    ]

    for path in candidates:
        if not path.exists():
            continue
        try:
            sources.append((str(path), path.read_text(encoding="utf-8")))
        except Exception:
            pass

    try:
        completed = subprocess.run(
            ["git", "show", "main:main.py"],
            check=True,
            capture_output=True,
            text=True,
            timeout=10,
        )
        if completed.stdout:
            sources.append(("git main:main.py", completed.stdout))
    except Exception:
        pass

    return sources


def save_credentials(username: str, password: str) -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps({"username": username, "password": password}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass


def main() -> int:
    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            if data.get("username") and data.get("password"):
                print(f"이미 문서24 로그인정보가 로컬에 저장되어 있습니다: {CONFIG_PATH}")
                return 0
        except Exception:
            pass

    for label, source in legacy_sources():
        credentials = extract_credentials(source)
        if credentials:
            save_credentials(*credentials)
            print(f"기존 코드에서 문서24 로그인정보를 찾아 로컬에 저장했습니다: {CONFIG_PATH}")
            print(f"가져온 위치: {label}")
            print("아이디/비밀번호 값 자체는 출력하지 않았습니다.")
            return 0

    print("기존 코드에서 문서24 로그인정보를 찾지 못했습니다.")
    print("백업 파일이 있다면 ~/Desktop/doc24-main-backup.py 위치를 확인해주세요.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
