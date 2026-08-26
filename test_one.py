from __future__ import annotations

from main import Doc24Automation, PreventSleep, load_recipients


def main() -> int:
    recipients = load_recipients()
    recipient = recipients[0]

    print(f"테스트 대상: {recipient}")
    print("실제 발송은 하지 않습니다.\n")

    with PreventSleep(), Doc24Automation() as automation:
        automation.ensure_login()
        title = automation.get_last_document_title()
        print(f"마지막 전송문서: {title}")

        result = automation.resend_to(recipient, dry_run=True)
        print(f"결과: {result.status}")
        if result.reason:
            print(f"사유: {result.reason}")

        if result.status == "실패":
            print("\n브라우저 화면을 확인해주세요.")
            try:
                input("확인 후 Enter를 누르면 종료합니다: ")
            except Exception:
                pass
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
