from __future__ import annotations

import csv
import re
import sys
from datetime import datetime
from pathlib import Path

from main import (
    Doc24Automation,
    PreventSleep,
    RESULT_DIR,
    SENT_DOCS_URL,
    ensure_dirs,
    load_recipients,
    log,
)


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text or "")


def row_is_today(text: str, now: datetime) -> bool:
    compact = normalize(text)
    year = now.year
    month = now.month
    day = now.day

    today_tokens = {
        f"{year}-{month:02d}-{day:02d}",
        f"{year}.{month:02d}.{day:02d}",
        f"{year}/{month:02d}/{day:02d}",
        f"{year}{month:02d}{day:02d}",
        f"{year}년{month}월{day}일",
        f"{month:02d}-{day:02d}",
        f"{month:02d}.{day:02d}",
        f"{month:02d}/{day:02d}",
        f"{month}월{day}일",
    }

    has_date = bool(
        re.search(r"20\d{2}[-./년]\d{1,2}[-./월]\d{1,2}", compact)
        or re.search(r"\d{1,2}[-./월]\d{1,2}", compact)
    )
    if not has_date:
        return True
    return any(token in compact for token in today_tokens)


def save_audit(
    recipients: list[str],
    sent: set[str],
    evidence: dict[str, str],
) -> tuple[Path, Path, Path]:
    ensure_dirs()
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    complete_path = RESULT_DIR / "발송완료.csv"
    missing_path = RESULT_DIR / "미발송.csv"
    report_path = RESULT_DIR / f"문서24_발송확인_{stamp}.csv"

    with complete_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관"])
        for name in recipients:
            if name in sent:
                writer.writerow([name])

    with missing_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관"])
        for name in recipients:
            if name not in sent:
                writer.writerow([name])

    with report_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["수신기관", "확인결과", "근거"])
        for name in recipients:
            writer.writerow([
                name,
                "발송확인" if name in sent else "미확인",
                evidence.get(name, ""),
            ])

    return complete_path, missing_path, report_path


def find_next_page(page) -> bool:
    return bool(
        page.evaluate(
            """
            () => {
                const visible = el => {
                    const style = window.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.display !== 'none'
                        && style.visibility !== 'hidden'
                        && rect.width > 0
                        && rect.height > 0;
                };

                const containers = Array.from(document.querySelectorAll(
                    '.pagination, .paging, .page, [class*="pagination"], [class*="paging"]'
                ));

                for (const container of containers) {
                    const buttons = Array.from(container.querySelectorAll('a, button'))
                        .filter(visible);
                    const next = buttons.find(el => {
                        const text = (el.innerText || el.textContent || '').trim();
                        const title = (el.getAttribute('title') || '').trim();
                        const cls = String(el.className || '');
                        const disabled = el.getAttribute('aria-disabled') === 'true'
                            || el.hasAttribute('disabled')
                            || cls.includes('disabled');
                        return !disabled && (
                            text === '다음'
                            || text === '>'
                            || text === '›'
                            || text === '»'
                            || title.includes('다음')
                        );
                    });
                    if (next) {
                        next.click();
                        return true;
                    }
                }

                const globalNext = Array.from(document.querySelectorAll('a, button'))
                    .filter(visible)
                    .find(el => {
                        const text = (el.innerText || el.textContent || '').trim();
                        const title = (el.getAttribute('title') || '').trim();
                        return text === '다음' || title.includes('다음');
                    });
                if (globalNext) {
                    globalNext.click();
                    return true;
                }
                return false;
            }
            """
        )
    )


def audit_sent_documents(max_pages: int = 100) -> int:
    recipients = load_recipients()
    normalized = {name: normalize(name) for name in recipients}
    sent: set[str] = set()
    evidence: dict[str, str] = {}
    now = datetime.now()

    with PreventSleep(), Doc24Automation() as automation:
        automation.ensure_login()
        assert automation.page is not None
        page = automation.page

        target_title = automation.get_last_document_title()
        target_override = " ".join(sys.argv[2:]).strip() if len(sys.argv) > 2 else ""
        if target_override:
            target_title = target_override
        normalized_title = normalize(target_title)

        print("\n" + "=" * 70)
        print("문서24 발송 확인 모드 - 읽기 전용")
        print(f"확인 날짜: {now:%Y-%m-%d}")
        print(f"확인 문서: {target_title}")
        print(f"대상 기관: {len(recipients)}개")
        print("=" * 70)

        page.goto(SENT_DOCS_URL)
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            pass
        page.wait_for_timeout(2000)

        seen_fingerprints: set[str] = set()

        for page_no in range(1, max_pages + 1):
            rows = page.locator("tbody tr")
            row_texts: list[str] = []

            for index in range(rows.count()):
                try:
                    text = rows.nth(index).inner_text().strip()
                except Exception:
                    continue
                if text:
                    row_texts.append(text)

            fingerprint = "\n".join(row_texts[:3])
            if fingerprint and fingerprint in seen_fingerprints:
                log("같은 페이지가 반복되어 확인을 종료합니다.")
                break
            if fingerprint:
                seen_fingerprints.add(fingerprint)

            matched_on_page = 0
            candidate_rows = 0

            for row_text in row_texts:
                compact_row = normalize(row_text)
                if normalized_title not in compact_row:
                    continue
                if not row_is_today(row_text, now):
                    continue

                candidate_rows += 1
                for name, compact_name in normalized.items():
                    if name in sent:
                        continue
                    if compact_name and compact_name in compact_row:
                        sent.add(name)
                        evidence[name] = re.sub(r"\s+", " ", row_text)[:500]
                        matched_on_page += 1

            log(
                f"보낸 문서함 {page_no}페이지 확인: "
                f"대상 문서 {candidate_rows}건 / 학교 {matched_on_page}개 추가 / 누적 {len(sent)}개"
            )

            save_audit(recipients, sent, evidence)

            if len(sent) == len(recipients):
                break

            before = fingerprint
            if not find_next_page(page):
                break

            page.wait_for_timeout(1500)
            try:
                page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass

            new_rows = page.locator("tbody tr")
            new_texts: list[str] = []
            for index in range(min(new_rows.count(), 3)):
                try:
                    new_texts.append(new_rows.nth(index).inner_text().strip())
                except Exception:
                    pass
            after = "\n".join(new_texts)
            if before and after == before:
                log("다음 페이지 이동이 확인되지 않아 종료합니다.")
                break

    complete_path, missing_path, report_path = save_audit(
        recipients, sent, evidence
    )
    missing = len(recipients) - len(sent)

    print("\n" + "=" * 70)
    print(f"발송 확인: {len(sent)}개")
    print(f"미확인: {missing}개")
    print(f"발송완료: {complete_path}")
    print(f"미발송: {missing_path}")
    print(f"전체 확인표: {report_path}")
    print("=" * 70)

    if not sent:
        print(
            "\n주의: 보낸 문서함 목록 행에 수신기관명이 표시되지 않는 구조라면 "
            "이 1차 확인에서는 0개로 나올 수 있습니다."
        )
        print("그 경우 보낸 문서 상세 화면을 읽는 2차 확인 방식으로 바꾸면 됩니다.")

    return 0


def main() -> int:
    if "--audit-sent" in sys.argv:
        return audit_sent_documents()
    print("사용법: python audit.py --audit-sent")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
