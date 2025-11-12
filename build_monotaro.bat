@echo off
rem 使い方: ダブルクリック or cmdで build.bat

rem 文字化け回避（UTF-8）
chcp 65001 > nul

rem 1) クリーン
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
rmdir /s /q __pycache__ 2>nul

rem 2) 依存インストール（ビルド環境用）
python -m pip install -U pip
python -m pip install -r requirements_monotaro.txt

rem 3) exe作成（ワンファイル、コンソール非表示、出力名はMonotaroOrder）
python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name "MonotaroOrder" ^
  --clean ^
  monotaro.py

echo ビルド完了: dist\MonotaroOrder.exe を配布してください。
pause
