@echo off
chcp 65001 > nul

rem 1) クリーン
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
rmdir /s /q __pycache__ 2>nul

rem 2) 依存インストール
python -m pip install -U pip
python -m pip install -r requirements_amazon.txt

rem 3) exe作成（ワンファイル、GUI）
python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name "AmazonExcel" ^
  --clean ^
  amazon.py

echo ビルド完了: dist\AmazonExcel.exe を配布してください。
pause
