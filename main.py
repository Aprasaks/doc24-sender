from __future__ import annotations

import csv
import os
import queue
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from playwright.sync_api import BrowserContext, Page, Playwright, sync_playwright

DOC24_HOME = "https://docu.gdoc.go.kr/index.do"
SENT_DOCS_URL = "https://docu.gdoc.go.kr/doc/snd/sendDocList.do"
PROFILE_DIR = Path.home() / ".doc24_sender" / "chrome-profile"
LOGIN_WAIT_SECONDS = 300


@dataclass
class RecipientResult:
    recipient: str
    status: str
    reason: str = ""


def parse_recipients(text: str) -> list[str]:
    recipients: list[str] = []
    seen: set[str] = set()
    for raw in text.replace(",", "\n").splitlines():
        name = raw.strip()
        if name and name not in seen:
            recipients.append(name)
            seen.add(name)
    return recipients


def load_recipient_file(path: str) -> list[str]:
    suffix = Path(path).suffix.lower()
    values: list[str] = []

    if suffix == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as handle:
            for row in csv.reader(handle):
                if row and row[0].strip():
                    values.append(row[0].strip())
    elif suffix == ".xlsx":
        workbook = load_workbook(path, read_only=True, data_only=True)
        sheet = workbook.active
        for row in sheet.iter_rows(values_only=True):
            if row and row[0] is not None:
                value = str(row[0]).strip()
                if value:
                    values.append(value)
        workbook.close()
    else:
        raise ValueError("CSV 또는 XLSX 파일만 불러올 수 있습니다.")

    if values and values[0].replace(" ", "") in {"학교명", "기관명", "수신기관", "수신자"}:
        values = values[1:]
    return parse_recipients("\n".join(values))


def save_failures(results: list[RecipientResult]) -> Path | None:
    failures = [result for result in results if result.status == "실패"]
    if not failures:
        return None

    downloads = Path.home() / "Downloads"
    output_dir = downloads if downloads.exists() else Path.home()
    output = output_dir / f"문서24_실패목록_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with output.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관", "상태", "사유"])
        for item in failures:
            writer.writerow([item.recipient, item.status, item.reason])
    return output


class Doc24Automation:
    def __init__(self, log):
        self.log = log
        self.playwright: Playwright | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None
        self.caffeinate: subprocess.Popen | None = None

    def __enter__(self):
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        self.playwright = sync_playwright().start()
        chromium = self.playwright.chromium

        chrome_paths = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            str(Path.home() / "Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
        ]
        executable = next((path for path in chrome_paths if os.path.exists(path)), None)
        if not executable:
            raise RuntimeError("Google Chrome을 찾을 수 없습니다. 맥에 Google Chrome을 설치해주세요.")

        try:
            self.context = chromium.launch_persistent_context(
                user_data_dir=str(PROFILE_DIR),
                executable_path=executable,
                headless=False,
                viewport={"width": 1280, "height": 900},
                args=["--no-first-run", "--no-default-browser-check"],
            )
        except Exception as exc:
            raise RuntimeError(
                "문서24 전용 Chrome을 실행할 수 없습니다. 이미 열려 있다면 닫고 다시 실행해주세요."
            ) from exc

        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()

        # 자동화 중 맥이 잠자기 상태로 들어가지 않도록 유지한다.
        try:
            self.caffeinate = subprocess.Popen(
                ["caffeinate", "-dimsu"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except Exception:
            self.caffeinate = None

        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.context:
                self.context.close()
        finally:
            if self.playwright:
                self.playwright.stop()
            if self.caffeinate:
                self.caffeinate.terminate()
                self.caffeinate = None

    def ensure_login(self) -> None:
        assert self.page is not None
        page = self.page
        self.log("문서24 로그인 상태 확인 중...")

        page.goto(SENT_DOCS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1400)
        if "로그아웃" in page.content():
            self.log("저장된 로그인 세션으로 접속했습니다.")
            return

        self.log("로그인이 필요합니다. 열린 전용 Chrome에서 한 번만 로그인해주세요.")
        page.goto(DOC24_HOME, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(800)

        try:
            login_button = page.get_by_text("로그인", exact=True)
            if login_button.count() and login_button.first.is_visible():
                login_button.first.click()
                page.wait_for_timeout(900)
            corporate = page.locator("#entrprsHref")
            if corporate.count() and corporate.is_visible():
                corporate.click()
        except Exception:
            pass

        deadline = time.monotonic() + LOGIN_WAIT_SECONDS
        while time.monotonic() < deadline:
            try:
                if "로그아웃" in page.content():
                    self.log("로그인 확인 완료. 다음 실행부터 이 로그인 상태를 재사용합니다.")
                    return
            except Exception:
                pass
            page.wait_for_timeout(1000)

        raise RuntimeError("5분 안에 로그인이 확인되지 않았습니다. 다시 실행해주세요.")

    def get_last_document_title(self) -> str:
        assert self.page is not None
        page = self.page
        page.goto(SENT_DOCS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1800)
        row = page.locator("tbody tr").first
        if row.count() == 0:
            raise RuntimeError("보낸 문서함에서 마지막 문서를 찾지 못했습니다.")
        link = row.locator("a").first
        if link.count() == 0:
            raise RuntimeError("마지막 전송문서 링크를 찾지 못했습니다.")
        title = link.inner_text().strip()
        if not title:
            title = row.inner_text().strip().splitlines()[0]
        return title

    def _open_last_document_for_rewrite(self) -> None:
        assert self.page is not None
        page = self.page
        page.goto(SENT_DOCS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1400)
        page.locator("tbody tr").first.locator("a").first.click(force=True)
        page.wait_for_timeout(1600)
        page.locator("button:has-text('재작성')").click(force=True)
        page.wait_for_timeout(700)
        self._click_dialog_button("예")
        page.wait_for_timeout(1500)

        for index in range(1, 5):
            locator = page.locator(f"label[for='wteChk{index}']")
            if locator.count() and locator.is_visible():
                locator.click()
                page.wait_for_timeout(120)

        confirm = page.get_by_role("button", name="확인")
        if confirm.count() and confirm.last.is_visible():
            try:
                confirm.last.click()
                page.wait_for_timeout(700)
            except Exception:
                pass

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
                if locator.count() and locator.last.is_visible(timeout=1000):
                    locator.last.click(force=True)
                    return
            except Exception:
                continue
        raise RuntimeError(f"'{label}' 확인 버튼을 찾지 못했습니다.")

    def _select_recipient(self, recipient: str) -> None:
        assert self.page is not None
        page = self.page
        page.locator("#ldapSearch").click()
        page.wait_for_timeout(1000)
        search_box = page.locator("#searchOrgNm")
        search_box.fill(recipient)

        search_button = page.get_by_role("button", name="검색")
        if search_button.count() and search_button.last.is_visible():
            search_button.last.click()
        else:
            page.locator("button[type='submit']").last.click()
        page.wait_for_timeout(1300)

        rows = page.locator("tr:has(button:has-text('선택'))")
        target = None
        normalized_target = recipient.replace(" ", "")
        for index in range(rows.count()):
            row = rows.nth(index)
            try:
                row_text = row.inner_text().replace(" ", "")
            except Exception:
                continue
            if normalized_target in row_text:
                target = row
                break

        if target is None:
            raise RuntimeError("수신기관 검색 결과에서 일치하는 기관을 찾지 못했습니다.")

        target.get_by_role("button", name="선택").click()
        page.wait_for_timeout(600)

    def resend_to(self, recipient: str, actual_send: bool) -> RecipientResult:
        assert self.page is not None
        page = self.page
        try:
            self._open_last_document_for_rewrite()
            self._select_recipient(recipient)
            if not actual_send:
                return RecipientResult(recipient, "테스트", "수신기관 검색/선택까지만 확인")

            page.locator("#sendDoc").click()
            page.wait_for_timeout(700)
            self._click_dialog_button("보내기")
            page.wait_for_timeout(700)
            self._click_dialog_button("예")
            page.wait_for_timeout(1500)
            return RecipientResult(recipient, "완료")
        except Exception as exc:
            try:
                page.keyboard.press("Escape")
            except Exception:
                pass
            return RecipientResult(recipient, "실패", str(exc))


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("문서24 재발송기")
        self.root.geometry("780x650")
        self.root.minsize(720, 580)

        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.preview_title = ""
        self.running = False

        self.actual_send = tk.BooleanVar(value=False)
        self.preview_text = tk.StringVar(value="마지막 전송문서를 먼저 확인해주세요.")
        self.status_text = tk.StringVar(value="대기 중")

        self._build_ui()
        self.root.after(100, self._poll_events)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=20)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="문서24 재발송기", font=("Helvetica", 22, "bold")).pack(anchor="w")
        ttk.Label(
            outer,
            text="전용 Chrome 로그인 상태를 저장하고 마지막 전송문서를 여러 기관에 순차 재발송합니다.",
        ).pack(anchor="w", pady=(4, 8))
        ttk.Label(
            outer,
            text=f"로그인 프로필: {PROFILE_DIR}",
            foreground="#666666",
        ).pack(anchor="w", pady=(0, 14))

        preview_frame = ttk.LabelFrame(outer, text="마지막 전송문서", padding=12)
        preview_frame.pack(fill="x")
        ttk.Label(preview_frame, textvariable=self.preview_text, wraplength=610).pack(side="left", fill="x", expand=True)
        self.preview_button = ttk.Button(preview_frame, text="확인", command=self.preview_last_document)
        self.preview_button.pack(side="right", padx=(12, 0))

        recipients_frame = ttk.LabelFrame(outer, text="수신기관 (한 줄에 하나)", padding=12)
        recipients_frame.pack(fill="both", expand=True, pady=(14, 0))
        button_row = ttk.Frame(recipients_frame)
        button_row.pack(fill="x", pady=(0, 8))
        ttk.Button(button_row, text="CSV/XLSX 불러오기", command=self.import_recipients).pack(side="left")
        ttk.Button(button_row, text="목록 비우기", command=lambda: self.recipients.delete("1.0", "end")).pack(side="left", padx=6)
        self.recipients = tk.Text(recipients_frame, height=12, wrap="none")
        self.recipients.pack(fill="both", expand=True)

        options = ttk.Frame(outer)
        options.pack(fill="x", pady=(14, 0))
        ttk.Checkbutton(
            options,
            text="실제 발송 (체크하지 않으면 수신기관 검색/선택까지만 테스트)",
            variable=self.actual_send,
        ).pack(side="left")

        action_row = ttk.Frame(outer)
        action_row.pack(fill="x", pady=(12, 0))
        self.start_button = ttk.Button(action_row, text="마지막 문서 재발송 시작", command=self.start_sending)
        self.start_button.pack(side="left")
        ttk.Label(action_row, textvariable=self.status_text).pack(side="right")

        log_frame = ttk.LabelFrame(outer, text="진행 로그", padding=8)
        log_frame.pack(fill="both", expand=True, pady=(14, 0))
        self.log_box = tk.Text(log_frame, height=9, state="disabled", wrap="word")
        self.log_box.pack(fill="both", expand=True)

    def import_recipients(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("수신기관 목록", "*.csv *.xlsx")])
        if not path:
            return
        try:
            names = load_recipient_file(path)
        except Exception as exc:
            messagebox.showerror("불러오기 실패", str(exc))
            return
        self.recipients.delete("1.0", "end")
        self.recipients.insert("1.0", "\n".join(names))
        self._append_log(f"수신기관 {len(names)}개를 불러왔습니다.")

    def preview_last_document(self) -> None:
        if self.running:
            return
        self._set_running(True, "마지막 문서 확인 중")
        threading.Thread(target=self._preview_worker, daemon=True).start()

    def _preview_worker(self) -> None:
        try:
            with Doc24Automation(self._thread_log) as automation:
                automation.ensure_login()
                title = automation.get_last_document_title()
            self.events.put(("preview", title))
        except Exception as exc:
            self.events.put(("error", str(exc)))
        finally:
            self.events.put(("idle", "대기 중"))

    def start_sending(self) -> None:
        if self.running:
            return
        if not self.preview_title:
            messagebox.showwarning("마지막 문서 확인", "먼저 [확인]을 눌러 마지막 전송문서를 확인해주세요.")
            return

        recipients = parse_recipients(self.recipients.get("1.0", "end"))
        if not recipients:
            messagebox.showwarning("수신기관", "수신기관을 한 곳 이상 입력해주세요.")
            return

        mode = "실제 발송" if self.actual_send.get() else "테스트"
        if self.actual_send.get():
            answer = messagebox.askyesno(
                "실제 발송 확인",
                f"'{self.preview_title}' 문서를 {len(recipients)}개 기관에 실제 발송합니다.\n\n계속할까요?",
            )
            if not answer:
                return

        self._append_log(f"{mode} 시작: {len(recipients)}개 기관")
        self._set_running(True, f"{mode} 진행 중")
        threading.Thread(
            target=self._send_worker,
            args=(recipients, self.preview_title, self.actual_send.get()),
            daemon=True,
        ).start()

    def _send_worker(
        self,
        recipients: list[str],
        expected_title: str,
        actual_send: bool,
    ) -> None:
        results: list[RecipientResult] = []
        try:
            with Doc24Automation(self._thread_log) as automation:
                automation.ensure_login()
                current_title = automation.get_last_document_title()
                if current_title != expected_title:
                    raise RuntimeError(
                        "마지막 전송문서가 확인했을 때와 달라졌습니다. 다시 [확인]을 눌러주세요.\n"
                        f"확인 당시: {expected_title}\n현재: {current_title}"
                    )

                total = len(recipients)
                for index, recipient in enumerate(recipients, start=1):
                    self._thread_log(f"[{index}/{total}] {recipient} 처리 중")
                    result = automation.resend_to(recipient, actual_send)
                    results.append(result)
                    if result.status == "실패":
                        self._thread_log(f"실패: {recipient} - {result.reason}")
                    elif result.status == "테스트":
                        self._thread_log(f"테스트 성공: {recipient}")
                    else:
                        self._thread_log(f"발송 완료: {recipient}")

            failure_file = save_failures(results)
            failures = sum(1 for result in results if result.status == "실패")
            completed = len(results) - failures
            summary = f"처리 완료 {completed} / 실패 {failures}"
            if failure_file:
                summary += f"\n실패 목록: {failure_file}"
            self.events.put(("done", summary))
        except Exception as exc:
            failure_file = save_failures(results)
            message = str(exc)
            if failure_file:
                message += f"\n실패 목록: {failure_file}"
            self.events.put(("error", message))
        finally:
            self.events.put(("idle", "대기 중"))

    def _thread_log(self, text: str) -> None:
        self.events.put(("log", text))

    def _append_log(self, text: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{timestamp}] {text}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _set_running(self, running: bool, status: str) -> None:
        self.running = running
        self.status_text.set(status)
        state = "disabled" if running else "normal"
        self.preview_button.configure(state=state)
        self.start_button.configure(state=state)

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "log":
                    self._append_log(str(payload))
                elif event == "preview":
                    self.preview_title = str(payload)
                    self.preview_text.set(self.preview_title)
                    self._append_log(f"마지막 전송문서 확인: {self.preview_title}")
                elif event == "done":
                    self._append_log(str(payload).replace("\n", " | "))
                    messagebox.showinfo("완료", str(payload))
                elif event == "error":
                    self._append_log(f"오류: {payload}")
                    messagebox.showerror("오류", str(payload))
                elif event == "idle":
                    self._set_running(False, str(payload))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)


def smoke_test() -> int:
    assert parse_recipients("기관A\n기관B\n기관A") == ["기관A", "기관B"]
    assert parse_recipients("기관A, 기관B") == ["기관A", "기관B"]
    assert PROFILE_DIR.name == "chrome-profile"
    return 0


def main() -> int:
    if "--smoke-test" in sys.argv:
        return smoke_test()
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
