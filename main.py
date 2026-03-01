import pandas as pd
from playwright.sync_api import sync_playwright
import time
import os
import sys
from multiprocessing import freeze_support

def run_macro():
    # 1. 엑셀 로드
    file_name = 'school_list.xlsx'
    if not os.path.exists(file_name):
        print(f"❌ '{file_name}' 파일이 없습니다. 엑셀 파일을 프로그램 옆에 넣어주세요!")
        time.sleep(5)
        return

    try:
        df = pd.read_excel(file_name)
        df.columns = df.columns.str.strip() 
    except Exception as e:
        print(f"❌ 엑셀 파일을 읽는 중 오류 발생: {e}")
        time.sleep(5)
        return

    with sync_playwright() as p:
        # 💡 브라우저 실행 (이미 설치되어 있어야 함)
        try:
            browser = p.chromium.launch(headless=False)
        except Exception as e:
            print("❌ 브라우저 엔진을 찾을 수 없습니다. 'playwright install chromium'을 먼저 실행해주세요.")
            print(f"오류 내용: {e}")
            time.sleep(10)
            return
            
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

                # 4. 최신 문서 선택 (텍스트 방식)
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

                # 7. 수신처 검색 및 충청남도 선택 (정밀 타격)
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
                    print(f"⚠️ {school_name} 검색 실패 또는 타지역: {e}")
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

if __name__ == "__main__":
    freeze_support() # 윈도우 환경 안전 장치
    run_macro()