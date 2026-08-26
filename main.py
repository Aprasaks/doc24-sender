from __future__ import annotations

import csv
import json
import os
import re
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from getpass import getpass
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import Browser, Page, Playwright, sync_playwright

DOC24_HOME = "https://docu.gdoc.go.kr/index.do"
SENT_DOCS_URL = "https://docu.gdoc.go.kr/doc/snd/sendDocList.do"

APP_DIR = Path.home() / ".doc24_sender"
CONFIG_PATH = APP_DIR / "config.json"
RESULT_DIR = APP_DIR / "results"

CHROME_PATHS = [
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    str(Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
]

REGION_TERMS = ["전북특별자치도", "전라북도", "전북"]


@dataclass
class RecipientResult:
    recipient: str
    status: str
    reason: str = ""


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


def ensure_dirs() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)


def load_recipients() -> list[str]:
    source = Path("school_list.xlsx")
    if not source.exists():
        raise FileNotFoundError("school_list.xlsx 파일이 없습니다.")

    workbook = load_workbook(source, read_only=True, data_only=True)
    sheet = workbook.active

    values: list[str] = []
    for row in sheet.iter_rows(values_only=True):
        if row and row[0] is not None:
            value = str(row[0]).strip()
            if value:
                values.append(value)
    workbook.close()

    if values and values[0].replace(" ", "") in {"학교명", "기관명", "수신기관", "수신자"}:
        values = values[1:]

    recipients: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value not in seen:
            recipients.append(value)
            seen.add(value)

    if not recipients:
        raise RuntimeError("school_list.xlsx에서 학교명을 찾지 못했습니다.")

    log(f"수신기관 {len(recipients)}개 로드: school_list.xlsx")
    return recipients


def save_results(results: list[RecipientResult]) -> tuple[Path, Path]:
    ensure_dirs()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result_path = RESULT_DIR / f"문서24_발송결과_{stamp}.csv"
    failed_path = RESULT_DIR / f"문서24_반송목록_{stamp}.csv"

    with result_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관", "상태", "사유"])
        for result in results:
            writer.writerow([result.recipient, result.status, result.reason])

    failures = [result for result in results if result.status == "실패"]
    with failed_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관", "사유"])
        for result in failures:
            writer.writerow([result.recipient, result.reason])

    return result_path, failed_path


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
        pass
    return None


def save_local_credentials(username: str, password: str) -> None:
    ensure_dirs()
    CONFIG_PATH.write_text(
        json.dumps({"username": username, "password": password}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    try:
        os.chmod(CONFIG_PATH, 0o600)
    except OSError:
        pass


def extract_legacy_credentials(source: str) -> tuple[str, str] | None:
    id_match = re.search(r'page\.fill\(\s*["\']#id["\']\s*,\s*["\']([^"\']+)["\']', source)
    pw_match = re.search(r'page\.keyboard\.type\(\s*["\']([^"\']+)["\']', source)
    if id_match and pw_match:
        return id_match.group(1), pw_match.group(1)
    return None


def migrate_legacy_credentials() -> tuple[str, str] | None:
    sources: list[str] = []

    backup = Path.home() / "Desktop" / "doc24-main-backup.py"
    if backup.exists():
        try:
            sources.append(backup.read_text(encoding="utf-8"))
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
            sources.append(completed.stdout)
    except Exception:
        pass

    for source in sources:
        credentials = extract_legacy_credentials(source)
        if credentials:
            save_local_credentials(*credentials)
            return credentials
    return None


def get_credentials() -> tuple[str, str]:
    credentials = load_local_credentials() or migrate_legacy_credentials()
    if credentials:
        return credentials

    print("문서24 로그인 정보를 한 번만 입력해주세요.")
    username = input("아이디: ").strip()
    password = getpass("비밀번호: ")
    if not username or not password:
        raise RuntimeError("아이디 또는 비밀번호가 비어 있습니다.")
    save_local_credentials(username, password)
    return username, password


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
            pass
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.process is not None:
            self.process.terminate()


class Doc24Automation:
    def __init__(self) -> None:
        self.playwright: Playwright | None = None
        self.browser: Browser | None = None
        self.page: Page | None = None

    def __enter__(self):
        self.playwright = sync_playwright().start()
        chrome_path = next((path for path in CHROME_PATHS if Path(path).exists()), None)
        if not chrome_path:
            raise RuntimeError("Google Chrome을 찾지 못했습니다.")

        self.browser = self.playwright.chromium.launch(
            executable_path=chrome_path,
            headless=False,
        )
        self.page = self.browser.new_page(viewport={"width": 1280, "height": 900})
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.browser is not None:
                self.browser.close()
        finally:
            if self.playwright is not None:
                self.playwright.stop()

    def ensure_login(self) -> None:
        assert self.page is not None
        page = self.page
        username, password = get_credentials()

        log("문서24 로그인 페이지 이동")
        page.goto(DOC24_HOME)
        log("로그인 메뉴 선택")
        page.get_by_text("로그인", exact=True).click()
        page.wait_for_timeout(2000)
        log("법인 계정 선택")
        page.click("#entrprsHref")
        log("아이디 입력")
        page.fill("#id", username)
        log("비밀번호 입력")
        page.click("#password")
        page.keyboard.type(password, delay=100)
        log("로그인 실행")
        page.press("#password", "Enter")
        page.wait_for_timeout(5000)

        if "로그아웃" not in page.content():
            raise RuntimeError("문서24 로그인 실패")
        log("로그인 성공")

    def _find_latest_document_link(self):
        assert self.page is not None
        rows = self.page.locator("tbody tr")
        for index in range(rows.count()):
            links = rows.nth(index).locator("a")
            if links.count() == 0:
                continue
            link = links.first
            try:
                title = link.inner_text().strip()
            except Exception:
                title = ""
            if title:
                return link, title
        raise RuntimeError("최신 전송문서를 찾지 못했습니다.")

    def get_last_document_title(self) -> str:
        assert self.page is not None
        page = self.page
        log("보낸 문서함 이동")
        page.goto(SENT_DOCS_URL)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        page.wait_for_timeout(3000)
        _, title = self._find_latest_document_link()
        return title

    def _open_latest_document_for_rewrite(self) -> None:
        assert self.page is not None
        page = self.page

        page.goto(SENT_DOCS_URL)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        page.wait_for_timeout(3000)

        link, _ = self._find_latest_document_link()
        link.click(force=True)
        page.wait_for_timeout(4000)

        page.locator("button:has-text('재작성')").click(force=True)
        page.wait_for_timeout(2000)
        page.evaluate("""
            const btn = Array.from(document.querySelectorAll('button.btnSkyBlue'))
                .find(b => b.innerText.trim() === '예');
            if (btn) btn.click();
        """)
        page.wait_for_timeout(4000)

        for index in range(1, 5):
            page.click(f"label[for='wteChk{index}']")
            page.wait_for_timeout(300)

        try:
            page.get_by_role("button", name="확인").click()
            page.wait_for_timeout(2000)
        except Exception:
            pass

    def _select_recipient(self, school_name: str) -> None:
        assert self.page is not None
        page = self.page

        page.click("#ldapSearch")
        page.wait_for_timeout(2500)
        search_box = page.locator("#searchOrgNm")
        search_box.click()
        search_box.fill("")
        page.keyboard.type(school_name, delay=150)
        page.click("button[type='submit']")
        page.wait_for_timeout(3000)

        rows = page.locator("tr:has(button:has-text('선택'))")
        row_count = rows.count()
        if row_count == 0:
            raise RuntimeError("검색 결과 없음")

        if row_count == 1:
            rows.first.get_by_role("button", name="선택").click()
            log(f"수신처 지정 완료: {school_name}")
            return

        for region in REGION_TERMS:
            target_row = page.locator(f"tr:has-text('{region}'):has(button:has-text('선택'))").first
            try:
                if target_row.is_visible(timeout=1000):
                    target_row.get_by_role("button", name="선택").click()
                    log(f"수신처 지정 완료: {school_name} ({region})")
                    return
            except Exception:
                pass

        raise RuntimeError("검색 결과에서 전북 기관을 찾지 못했습니다.")

    def resend_to(self, school_name: str, dry_run: bool) -> RecipientResult:
        assert self.page is not None
        page = self.page
        try:
            self._open_latest_document_for_rewrite()
            self._select_recipient(school_name)

            if dry_run:
                return RecipientResult(school_name, "테스트", "수신처 선택 성공")

            log(f"최종 발송 중: {school_name}")

            # 오늘 수정 전 마지막 코드(ea675a6)의 최종 발송 3단계 방식 복원
            page.click("#sendDoc")
            log("1단계 전송요청 클릭 완료")
            page.wait_for_timeout(2000)

            try:
                page.wait_for_selector(".jconfirm-buttons button", state="visible", timeout=5000)
                page.evaluate("""
                    const modalBtns = Array.from(document.querySelectorAll('.jconfirm-buttons button'));
                    const sendBtn = modalBtns.find(b => b.innerText.trim() === '보내기');
                    if (sendBtn) sendBtn.click();
                """)
                log("2단계 보내기 클릭 완료")
            except Exception:
                page.keyboard.press("Enter")
                log("2단계 보내기 Enter 대체 실행")
            page.wait_for_timeout(2000)

            try:
                page.wait_for_selector("button.btnSkyBlue", state="visible", timeout=5000)
                page.evaluate("""
                    const finalBtns = Array.from(document.querySelectorAll('button.btnSkyBlue'));
                    const yesBtn = finalBtns.find(b => b.innerText.trim() === '예');
                    if (yesBtn) yesBtn.click();
                """)
                log("3단계 최종 예 클릭 완료")
            except Exception:
                page.keyboard.press("Enter")
                log("3단계 최종 예 Enter 대체 실행")

            page.wait_for_timeout(3000)
            return RecipientResult(school_name, "완료")
        except Exception as exc:
            try:
                page.keyboard.press("Escape")
                page.wait_for_timeout(500)
            except Exception:
                pass
            return RecipientResult(school_name, "실패", str(exc))


def run() -> int:
    recipients = load_recipients()
    dry_run = os.getenv("DOC24_DRY_RUN", "").strip().lower() in {"1", "true", "yes", "y"}

    with PreventSleep(), Doc24Automation() as automation:
        automation.ensure_login()
        title = automation.get_last_document_title()

        print("\n" + "=" * 70)
        print(f"마지막 전송문서: {title}")
        print(f"수신기관: {len(recipients)}개")
        print("모드:", "테스트" if dry_run else "실제 발송")
        print("=" * 70)

        if not dry_run:
            if input("실제 발송하려면 YES 입력: ").strip() != "YES":
                print("취소했습니다.")
                return 0

        results: list[RecipientResult] = []
        total = len(recipients)

        for index, school_name in enumerate(recipients, start=1):
            log(f"[{index}/{total}] {school_name} 작업 시작")
            result = automation.resend_to(school_name, dry_run)
            results.append(result)

            if result.status == "실패":
                log(f"반송/실패 기록: {school_name} - {result.reason}")
                log("다음 기관으로 이동")
            elif result.status == "테스트":
                log(f"테스트 성공: {school_name}")
            else:
                log(f"발송 완료: {school_name}")

        result_path, failed_path = save_results(results)
        success = sum(1 for result in results if result.status != "실패")
        failed = sum(1 for result in results if result.status == "실패")

        print("\n" + "=" * 70)
        print(f"처리 완료: 성공 {success} / 반송·실패 {failed}")
        print(f"전체 결과: {result_path}")
        print(f"반송 목록: {failed_path}")
        print("=" * 70)

    return 0


def smoke_test() -> int:
    assert REGION_TERMS[0] == "전북특별자치도"
    sample = [
        RecipientResult("학교A", "완료"),
        RecipientResult("학교B", "실패", "검색 결과 없음"),
    ]
    assert [item.recipient for item in sample if item.status == "실패"] == ["학교B"]
    return 0


def main() -> int:
    if "--smoke-test" in sys.argv:
        return smoke_test()
    return run()


if __name__ == "__main__":
    raise SystemExit(main())
