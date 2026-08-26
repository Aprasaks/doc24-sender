from __future__ import annotations

import csv
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook
from playwright.sync_api import BrowserContext, Page, Playwright, sync_playwright

DOC24_HOME = "https://docu.gdoc.go.kr/index.do"
SENT_DOCS_URL = "https://docu.gdoc.go.kr/doc/snd/sendDocList.do"

APP_DIR = Path.home() / ".doc24_sender"
PROFILE_DIR = APP_DIR / "chrome-profile"
CONFIG_PATH = APP_DIR / "config.json"
RESULT_DIR = APP_DIR / "results"

CHROME_PATHS = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    str(Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
]


@dataclass
class RecipientResult:
    recipient: str
    status: str
    reason: str = ""


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def ensure_app_dirs() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)


def parse_recipients(text: str) -> list[str]:
    recipients: list[str] = []
    seen: set[str] = set()
    for raw in text.replace(",", "\n").splitlines():
        value = raw.strip()
        if value and value not in seen:
            recipients.append(value)
            seen.add(value)
    return recipients


def load_recipients() -> list[str]:
    candidates = [
        Path("school_list.xlsx"),
        Path("recipients.xlsx"),
        Path("recipients.csv"),
        Path("recipients.txt"),
    ]
    source = next((path for path in candidates if path.exists()), None)
    if source is None:
        raise FileNotFoundError(
            "수신기관 파일이 없습니다.\n"
            "프로젝트 폴더에 school_list.xlsx, recipients.xlsx, recipients.csv, recipients.txt 중 하나를 넣어주세요."
        )

    if source.suffix.lower() == ".txt":
        names = parse_recipients(source.read_text(encoding="utf-8-sig"))
    elif source.suffix.lower() == ".csv":
        values: list[str] = []
        with source.open("r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.reader(handle):
                if row and str(row[0]).strip():
                    values.append(str(row[0]).strip())
        names = parse_recipients("\n".join(values))
    else:
        workbook = load_workbook(source, read_only=True, data_only=True)
        sheet = workbook.active
        values = []
        for row in sheet.iter_rows(values_only=True):
            if row and row[0] is not None:
                values.append(str(row[0]).strip())
        workbook.close()
        names = parse_recipients("\n".join(values))

    if names and names[0].replace(" ", "") in {"학교명", "기관명", "수신기관", "수신자"}:
        names = names[1:]

    if not names:
        raise RuntimeError(f"{source.name}에서 수신기관을 찾지 못했습니다.")

    log(f"수신기관 {len(names)}개 로드: {source.name}")
    return names


def save_results(results: list[RecipientResult]) -> Path:
    ensure_app_dirs()
    output = RESULT_DIR / f"문서24_발송결과_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관", "상태", "사유"])
        for item in results:
            writer.writerow([item.recipient, item.status, item.reason])
    return output


def load_local_credentials() -> tuple[str, str] | None:
    if not CONFIG_PATH.exists():
        return None
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        username = str(data.get("username", "")).strip()
        password = str(data.get("password", ""))
        if username and password:
            return username, password
    except Exception:
        return None
    return None


def save_local_credentials(username: str, password: str) -> None:
    ensure_app_dirs()
    CONFIG_PATH.write_text(
        json.dumps({"username": username, "password": password}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass


def extract_legacy_credentials(source: str) -> tuple[str, str] | None:
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


def migrate_legacy_credentials() -> tuple[str, str] | None:
    if CONFIG_PATH.exists():
        return load_local_credentials()

    legacy_sources: list[str] = []

    for path in [Path("legacy_main.py"), Path("main_old.py"), Path("old_main.py")]:
        if path.exists():
            try:
                legacy_sources.append(path.read_text(encoding="utf-8"))
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
            legacy_sources.append(completed.stdout)
    except Exception:
        pass

    for source in legacy_sources:
        credentials = extract_legacy_credentials(source)
        if credentials:
            save_local_credentials(*credentials)
            log("기존 코드의 로그인 정보를 맥 로컬 설정으로 이전했습니다.")
            return credentials

    return None


class PreventSleep:
    def __init__(self) -> None:
        self.process: subprocess.Popen | None = None

    def __enter__(self):
        try:
            self.process = subprocess.Popen(
                ["caffeinate", "-dimsu"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception:
            self.process = None
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.process is not None:
            self.process.terminate()
            try:
                self.process.wait(timeout=2)
            except Exception:
                self.process.kill()


class Doc24Automation:
    def __init__(self, logger: Callable[[str], None] = log):
        self.log = logger
        self.playwright: Playwright | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None

    def __enter__(self):
        ensure_app_dirs()
        self.playwright = sync_playwright().start()
        chrome_path = next((path for path in CHROME_PATHS if Path(path).exists()), None)
        if not chrome_path:
            raise RuntimeError("Google Chrome을 찾지 못했습니다. /Applications에 Chrome을 설치해주세요.")

        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            executable_path=chrome_path,
            headless=False,
            viewport={"width": 1280, "height": 900},
            args=["--disable-notifications"],
        )
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.context is not None:
                self.context.close()
        finally:
            if self.playwright is not None:
                self.playwright.stop()

    def _is_logged_in(self) -> bool:
        assert self.page is not None
        page = self.page
        try:
            page.goto(SENT_DOCS_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1200)
        except Exception:
            return False
        content = page.content()
        return "로그아웃" in content and "보낸 문서" in content

    def ensure_login(self) -> None:
        assert self.page is not None
        if self._is_logged_in():
            self.log("저장된 문서24 로그인 세션 사용")
            return

        credentials = load_local_credentials() or migrate_legacy_credentials()
        page = self.page

        if credentials:
            username, password = credentials
            self.log("저장된 로컬 계정정보로 문서24 로그인 시도")
            page.goto(DOC24_HOME, wait_until="domcontentloaded", timeout=30000)
            page.get_by_text("로그인", exact=True).click()
            page.wait_for_timeout(1000)
            page.locator("#entrprsHref").click()
            page.locator("#id").fill(username)
            page.locator("#password").fill(password)
            page.locator("#password").press("Enter")
            page.wait_for_timeout(3000)

            if self._is_logged_in():
                self.log("자동 로그인 성공")
                return
            self.log("자동 로그인 실패. 전용 Chrome에서 직접 로그인해주세요.")

        page.goto(DOC24_HOME, wait_until="domcontentloaded", timeout=30000)
        print("\n문서24 전용 Chrome에서 로그인해주세요.")
        print("로그인 완료 후 이 터미널로 돌아와 Enter를 누르세요.\n")
        input()
        if not self._is_logged_in():
            raise RuntimeError("문서24 로그인 상태를 확인하지 못했습니다.")
        self.log("로그인 확인 완료. 다음 실행부터 이 Chrome 세션을 재사용합니다.")

    def _save_last_document_debug(self, reason: str) -> Path:
        assert self.page is not None
        ensure_app_dirs()
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        prefix = RESULT_DIR / f"last_document_debug_{stamp}"
        try:
            self.page.screenshot(path=str(prefix.with_suffix(".png")), full_page=True)
        except Exception:
            pass
        try:
            prefix.with_suffix(".html").write_text(self.page.content(), encoding="utf-8")
        except Exception:
            pass
        try:
            prefix.with_suffix(".txt").write_text(reason, encoding="utf-8")
        except Exception:
            pass
        return prefix

    def _find_latest_document(self):
        assert self.page is not None
        page = self.page
        rows = page.locator("tbody tr")
        row_count = rows.count()
        if row_count == 0:
            debug = self._save_last_document_debug("tbody tr이 없습니다.")
            raise RuntimeError(f"보낸 문서함에서 문서 행을 찾지 못했습니다. 디버그: {debug}")

        row_logs: list[str] = []
        for index in range(row_count):
            row = rows.nth(index)
            try:
                row_text = " ".join(row.inner_text().split())
            except Exception:
                row_text = ""

            if row_text:
                row_logs.append(f"row {index}: {row_text[:300]}")

            if not row_text:
                continue
            if any(message in row_text for message in ["조회된 데이터가 없습니다", "검색 결과가 없습니다"]):
                continue

            selectors = ["a", "button", "[onclick]", "[role='button']"]
            for selector in selectors:
                candidates = row.locator(selector)
                for candidate_index in range(candidates.count()):
                    candidate = candidates.nth(candidate_index)
                    try:
                        text = " ".join(candidate.inner_text().split())
                    except Exception:
                        text = ""
                    try:
                        visible = candidate.is_visible()
                    except Exception:
                        visible = False

                    # 문서 제목은 보통 텍스트가 있는 클릭 요소다.
                    if visible and text and text not in {"보기", "상세", "다운로드", "삭제"}:
                        return row, candidate, text

            # 행 자체에 onclick이 걸린 구조도 허용한다.
            try:
                if row.get_attribute("onclick"):
                    return row, row, row_text
            except Exception:
                pass

        # 클릭요소를 못 찾더라도 첫 번째 실제 데이터 행 텍스트는 진단에 남긴다.
        reason = "\n".join(row_logs[:10]) or "행 텍스트도 읽지 못했습니다."
        debug = self._save_last_document_debug(reason)
        raise RuntimeError(
            "마지막 전송문서의 클릭 요소를 찾지 못했습니다. "
            f"디버그 파일이 저장되었습니다: {debug}"
        )

    def get_last_document_title(self) -> str:
        assert self.page is not None
        page = self.page
        page.goto(SENT_DOCS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1800)

        row, target, title = self._find_latest_document()
        if title:
            return title

        row_text = " ".join(row.inner_text().split())
        if row_text:
            return row_text
        raise RuntimeError("마지막 전송문서 제목을 읽지 못했습니다.")

    def _click_dialog_button(self, label: str) -> None:
        assert self.page is not None
        page = self.page
        candidates = [
            page.locator(f".jconfirm-buttons button:has-text('{label}')"),
            page.locator(f"button.btnSkyBlue:has-text('{label}')"),
            page.get_by_role("button", name=label),
        ]
        for locator in candidates:
            try:
                if locator.count() and locator.last.is_visible(timeout=1200):
                    locator.last.click(force=True)
                    return
            except Exception:
                continue
        raise RuntimeError(f"'{label}' 버튼을 찾지 못했습니다.")

    def _open_last_document_for_rewrite(self) -> None:
        assert self.page is not None
        page = self.page

        page.goto(SENT_DOCS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1800)

        _, target, _ = self._find_latest_document()
        target.click(force=True)
        page.wait_for_timeout(1600)

        rewrite_button = page.locator("button:has-text('재작성')")
        if rewrite_button.count() == 0:
            debug = self._save_last_document_debug("문서 상세 진입 후 재작성 버튼이 없습니다.")
            raise RuntimeError(f"재작성 버튼을 찾지 못했습니다. 디버그: {debug}")
        rewrite_button.first.click(force=True)
        page.wait_for_timeout(700)
        self._click_dialog_button("예")
        page.wait_for_timeout(1400)

        for index in range(1, 5):
            checkbox = page.locator(f"label[for='wteChk{index}']")
            try:
                if checkbox.count() and checkbox.is_visible():
                    checkbox.click()
                    page.wait_for_timeout(120)
            except Exception:
                pass

        confirm = page.get_by_role("button", name="확인")
        try:
            if confirm.count() and confirm.last.is_visible(timeout=700):
                confirm.last.click()
                page.wait_for_timeout(600)
        except Exception:
            pass

    def _select_recipient(self, recipient: str) -> None:
        assert self.page is not None
        page = self.page

        page.locator("#ldapSearch").click()
        page.wait_for_timeout(1000)

        search_box = page.locator("#searchOrgNm")
        search_box.fill(recipient)

        search_button = page.get_by_role("button", name="검색")
        if search_button.count():
            search_button.last.click()
        else:
            page.locator("button[type='submit']").last.click()

        page.wait_for_timeout(1400)

        rows = page.locator("tr:has(button:has-text('선택'))")
        normalized_target = re.sub(r"\s+", "", recipient)
        exact_match = None
        loose_match = None

        for index in range(rows.count()):
            row = rows.nth(index)
            try:
                raw = row.inner_text()
            except Exception:
                continue
            normalized_row = re.sub(r"\s+", "", raw)

            if normalized_target == normalized_row or normalized_row.startswith(normalized_target):
                exact_match = row
                break
            if normalized_target in normalized_row and loose_match is None:
                loose_match = row

        target = exact_match or loose_match
        if target is None:
            raise RuntimeError("수신기관 검색 결과에서 일치하는 기관을 찾지 못했습니다.")

        target.get_by_role("button", name="선택").click()
        page.wait_for_timeout(600)

    def resend_to(self, recipient: str, dry_run: bool) -> RecipientResult:
        assert self.page is not None
        page = self.page
        try:
            self._open_last_document_for_rewrite()
            self._select_recipient(recipient)

            if dry_run:
                return RecipientResult(recipient, "테스트", "수신기관 검색/선택 성공")

            page.locator("#sendDoc").click()
            page.wait_for_timeout(700)
            self._click_dialog_button("보내기")
            page.wait_for_timeout(700)
            self._click_dialog_button("예")
            page.wait_for_timeout(1600)
            return RecipientResult(recipient, "완료")
        except Exception as exc:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            return RecipientResult(recipient, "실패", str(exc))


def run() -> int:
    ensure_app_dirs()
    recipients = load_recipients()
    dry_run = os.getenv("DOC24_DRY_RUN", "").strip().lower() in {"1", "true", "yes", "y"}

    with PreventSleep(), Doc24Automation() as automation:
        automation.ensure_login()
        last_title = automation.get_last_document_title()

        print("\n" + "=" * 70)
        print("마지막 전송문서")
        print(last_title)
        print(f"수신기관: {len(recipients)}개")
        print("모드:", "테스트(실제 발송 안 함)" if dry_run else "실제 발송")
        print("=" * 70)

        if not dry_run:
            answer = input("이 문서를 위 수신기관에 재발송할까요? [YES 입력]: ").strip()
            if answer != "YES":
                print("취소했습니다.")
                return 0
        else:
            input("테스트를 시작하려면 Enter를 누르세요: ")

        results: list[RecipientResult] = []
        total = len(recipients)

        for index, recipient in enumerate(recipients, start=1):
            log(f"[{index}/{total}] {recipient} 처리 시작")
            result = automation.resend_to(recipient, dry_run=dry_run)
            results.append(result)

            if result.status == "실패":
                log(f"실패: {recipient} - {result.reason}")
            elif result.status == "테스트":
                log(f"테스트 성공: {recipient}")
            else:
                log(f"발송 완료: {recipient}")

        output = save_results(results)
        success = sum(1 for item in results if item.status in {"완료", "테스트"})
        failed = sum(1 for item in results if item.status == "실패")

        print("\n" + "=" * 70)
        print(f"처리 완료: {success} / 실패: {failed}")
        print(f"결과 파일: {output}")
        print("=" * 70)

    return 0


def smoke_test() -> int:
    assert parse_recipients("기관A\n기관B\n기관A") == ["기관A", "기관B"]
    assert parse_recipients("기관A, 기관B") == ["기관A", "기관B"]

    sample = 'page.fill("#id", "example-user")\npage.keyboard.type("example-password", delay=100)'
    assert extract_legacy_credentials(sample) == ("example-user", "example-password")
    return 0


def main() -> int:
    if "--smoke-test" in sys.argv:
        return smoke_test()
    return run()


if __name__ == "__main__":
    raise SystemExit(main())
