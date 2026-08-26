#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SOURCE_APP="$SCRIPT_DIR/Doc24Sender.app"
TARGET_DIR="$HOME/Applications"
TARGET_APP="$TARGET_DIR/Doc24Sender.app"

if [ ! -d "$SOURCE_APP" ]; then
  echo "Doc24Sender.app을 찾을 수 없습니다. 설치 파일과 앱을 같은 폴더에 두고 다시 실행해주세요."
  echo
  read "?Enter 키를 누르면 종료합니다."
  exit 1
fi

echo "문서24 재발송기를 설치합니다..."
mkdir -p "$TARGET_DIR"

# 다운로드 격리 속성을 제거한 뒤 사용자 Applications 폴더로 복사합니다.
xattr -dr com.apple.quarantine "$SOURCE_APP" 2>/dev/null || true
rm -rf "$TARGET_APP"
ditto "$SOURCE_APP" "$TARGET_APP"
xattr -dr com.apple.quarantine "$TARGET_APP" 2>/dev/null || true

# 앱 번들 서명 상태를 확인합니다. 실패해도 설치 자체는 계속 진행합니다.
codesign --verify --deep --strict "$TARGET_APP" 2>/dev/null || true

echo "설치 완료: $TARGET_APP"
echo "앱을 실행합니다."
open "$TARGET_APP"
