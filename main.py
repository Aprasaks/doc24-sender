import pandas as pd
from playwright.sync_api import sync_playwright
import time
import os
import sys
import subprocess
from multiprocessing import freeze_support

# 🚀 브라우저 엔진 자동 설치 함수
def install_browser():
    print("🌐 시스템 환경을 점검하고 브라우저 엔진을 확인합니다...")
    try:
        # EXE 실행 시 자기 자신을 다시 부르지 않도록 subprocess로 분리해서 실행
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
    except Exception as e:
        print(f"⚠️ 브라우저 설치 시도 중 참고사항: {e}")

def run_macro():
    # 1. 엑셀 로드
    file_name = 'school_list.xlsx'
    if not os.path.exists(file_name):
        print(f"❌ '{file_name}' 파일이 없습니다. 엑셀 파일을 넣어줘 형!")
        time.sleep(5)
        return

    df = pd.read_excel(file_name)
    df.columns = df.columns.str.strip() 

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        # 2. 법인 로그인
        print("🏢 법인 계정 로그인 중...")
        page.goto("https://docu.gdoc.go.kr/index.do")
        page.get_by_text("로그인", exact=True).click()
        time.sleep(2)
        page.click("#entrprsHref") 
        page.fill("#id", "safejb")
        page.click("#password")
        page.keyboard.type("dkswjs26504#", delay=100)
        page.press("#password", "Enter")
        time.sleep(5)

        if "로그아웃" in page.content():
            for index, row in df.iterrows():
                school_name = str(row['학교명']).strip()
                print(f"🔄 [{index+1}/{len(df)}] '{school_name}' 작업 시작")

                # 3. 보낸 문서함으로 이동
                page.goto("https://docu.gdoc.go.kr/doc/snd/sendDocList.do")
                page.wait_for_load_state("networkidle")
                time.sleep(3) 

                # 4. 최신 문서 선택
                try:
                    page.get_by_text("2026년 교직원 응급처치 교육").first.click(force=True)
                except:
                    page.locator("tbody tr").first.locator("a").first.click(force=True)
                time.sleep(4)

                # 5. [재작성] -> '예' 팝업
                page.locator("button:has-text('재작성')").click(force=True)
                time.sleep(2)
                page.evaluate("""
                    const btn = Array.from(document.querySelectorAll('button.btnSkyBlue')).find(b => b.innerText.trim() === '예');
                    if (btn) btn.click();
                """)
                time.sleep(4)

                # 6. 사전 확인 4종 체크
                for i in range(1, 5):
                    page.click(f"label[for='wteChk{i}']")
                    time.sleep(0.3)
                try:
                    page.get_by_role("button", name="확인").click()
                    time.sleep(2)
                except: pass

                # 7. 수신처 검색 및 충청남도 선택
                page.click("#ldapSearch")
                time.sleep(2.5)
                search_box = page.locator("#searchOrgNm")
                search_box.click()
                page.keyboard.type(school_name, delay=150)
                page.click("button[type='submit']")
                time.sleep(3)

                try:
                    target_row = page.locator("tr:has-text('충청남도')").first
                    if not target_row.is_visible(timeout=2000):
                        target_row = page.locator("tr:has-text('충남')").first

                    if target_row.is_visible(timeout=1000):
                        target_row.get_by_role("button", name="선택").click()
                        print(f"✅ 수신처 지정 완료: {school_name} (충청남도)")
                    else: 
                        raise Exception("검색 결과에 충청남도/충남 없음")
                except Exception as e:
                    print(f"⚠️ {school_name} 충남 지역 검색 실패: {e}")
                    page.keyboard.press("Escape")
                    time.sleep(1)
                    continue

                # 8. 최종 발송 3단계 콤보
                print(f"🚀 '{school_name}' 최종 발송 중...")
                page.click("#sendDoc") 
                time.sleep(2)

                try:
                    page.wait_for_selector(".jconfirm-buttons button", state="visible", timeout=5000)
                    page.evaluate("""
                        const modalBtns = Array.from(document.querySelectorAll('.jconfirm-buttons button'));
                        const sendBtn = modalBtns.find(b => b.innerText.trim() === '보내기');
                        if (sendBtn) sendBtn.click();
                    """)
                except:
                    page.keyboard.press("Enter")
                time.sleep(2)

                try:
                    page.wait_for_selector("button.btnSkyBlue", state="visible", timeout=5000)
                    page.evaluate("""
                        const finalBtns = Array.from(document.querySelectorAll('button.btnSkyBlue'));
                        const yesBtn = finalBtns.find(b => b.innerText.trim() === '예');
                        if (yesBtn) yesBtn.click();
                    """)
                except:
                    page.keyboard.press("Enter")
                
                print(f"🏁 '{school_name}' 발송 완료!")
                time.sleep(3)

        browser.close()

# 💡 [핵심] 윈도우 EXE 실행 시 무한 루프를 방지하는 시작점
if __name__ == "__main__":
    freeze_support()    # 윈도우 멀티프로세싱 지원 (무한 실행 방지)
    install_browser()   # 브라우저 설치 시도
    run_macro()         # 매크로 실행