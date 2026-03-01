import pandas as pd
from playwright.sync_api import sync_playwright
import time
import os
import sys

# 🚀 [추가] 브라우저 엔진 자동 설치 로직
# 프로그램 실행 시 브라우저가 없으면 알아서 설치해주는 명령어임
print("🌐 시스템 환경을 점검하고 브라우저 엔진을 확인합니다...")
os.system(f"{sys.executable} -m playwright install chromium")

# 1. 엑셀 로드
file_name = 'school_list.xlsx'
if not os.path.exists(file_name):
    print(f"❌ '{file_name}' 파일이 없습니다. 엑셀 파일을 넣어줘 형!")
    time.sleep(5)
    exit()

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

            # 7. 수신처 검색 및 충청남도 선택 (정밀 타격 버전)
            page.click("#ldapSearch")
            time.sleep(2.5)
            search_box = page.locator("#searchOrgNm")
            search_box.click()
            page.keyboard.type(school_name, delay=150)
            page.click("button[type='submit']")
            time.sleep(3)

            try:
                # 💡 방법 1: '충청남도'라는 텍스트를 가진 행(tr)을 찾음
                target_row = page.locator("tr:has-text('충청남도')").first
                
                # 만약 위 방법으로 안 잡히면 '충남'으로 재시도
                if not target_row.is_visible(timeout=2000):
                    target_row = page.locator("tr:has-text('충남')").first

                if target_row.is_visible(timeout=1000):
                    # 해당 행 안에 있는 '선택' 버튼 클릭
                    target_row.get_by_role("button", name="선택").click()
                    print(f"✅ 수신처 지정 완료: {school_name} (충청남도)")
                else: 
                    # 검색 결과는 떴는데 충남이 아닐 경우
                    raise Exception("검색 결과에 충청남도/충남 없음")
            except Exception as e:
                # 💡 학교가 아예 안 나오거나 다른 지역만 뜰 때
                print(f"⚠️ {school_name} 충남 지역 검색 실패: {e}")
                page.keyboard.press("Escape")
                time.sleep(1)
                continue

            # 8. 최종 발송 3단계 콤보
            print(f"🚀 '{school_name}' 최종 발송 중...")
            page.click("#sendDoc") 
            time.sleep(2)

            # (2) 보내기 모달 버튼 강제 클릭
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

            # (3) 최종 확인 '예' 버튼 강제 클릭
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