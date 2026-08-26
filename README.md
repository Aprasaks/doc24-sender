# 문서24 재발송기 - macOS Python 버전

문서24에서 **마지막 전송문서를 재작성**하여 여러 수신기관에 순차 발송하는 macOS용 Playwright 자동화입니다.

## 실행 방식

GUI 앱이 아니라 원래처럼 Python으로 실행합니다.

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python main.py
```

Google Chrome이 `/Applications/Google Chrome.app`에 설치되어 있어야 합니다.

## 로그인

- 전용 Chrome 프로필은 `~/.doc24_sender/chrome-profile`에 저장됩니다.
- 로그인 세션이 살아 있으면 다음 실행부터 바로 재사용합니다.
- 현재 브랜치에서 처음 실행할 때 로컬 로그인 설정이 없으면, 같은 Git 저장소의 `main:main.py`에 있던 기존 로그인 코드를 읽어 계정정보를 `~/.doc24_sender/config.json`으로 이전할 수 있습니다.
- `config.json`은 macOS 로컬에만 저장되고 Git에는 올라가지 않습니다.
- 자동 로그인이 안 되면 전용 Chrome에서 한 번 직접 로그인한 뒤 Enter를 누르면 이후 세션을 재사용합니다.

## 수신기관 파일

프로젝트 폴더에 아래 중 하나를 넣으면 됩니다.

- `school_list.xlsx`
- `recipients.xlsx`
- `recipients.csv`
- `recipients.txt`

첫 번째 열 또는 TXT 각 줄에 기관명을 적습니다. 헤더는 `학교명`, `기관명`, `수신기관`, `수신자`를 지원합니다.

## 동작 순서

1. 전용 Chrome 실행
2. 저장된 로그인 세션 확인
3. 마지막 전송문서 제목 확인
4. 수신기관 목록 로드
5. 터미널에서 마지막 문서와 수신기관 수 확인
6. `YES` 입력 시 실제 재발송 시작
7. 각 기관마다 마지막 문서 재작성 → 수신기관 검색 → 선택 → 발송
8. 검색되지 않거나 오류 난 기관은 건너뛰고 다음 기관 계속 진행
9. 결과를 `~/.doc24_sender/results/문서24_발송결과_*.csv`에 저장

## 테스트 모드

실제 발송 버튼을 누르지 않고 수신기관 검색/선택까지만 확인하려면:

```bash
DOC24_DRY_RUN=1 python main.py
```

## 다른 작업을 해도 되는가

가능합니다. Playwright는 마우스를 직접 움직이는 매크로가 아니라 브라우저 DOM을 직접 제어하므로 자동화 Chrome이 뒤에 있어도 동작합니다.

다만 자동화 중에는 해당 Chrome 창을 직접 조작하거나 닫지 마세요. 스크립트 실행 중에는 `caffeinate`를 같이 실행하여 Mac이 잠자기 상태로 들어가지 않도록 합니다.

## 보안

공개 GitHub 코드에는 새 로그인 비밀번호를 다시 하드코딩하지 않습니다. 기존 공개 이력에 로그인정보가 남아 있었다면 비밀번호 변경을 권장합니다.
