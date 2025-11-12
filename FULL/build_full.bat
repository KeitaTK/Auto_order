@echo off
rem 使い方: ダブルクリック or cmdで build_full.bat

rem 文字化け回避（UTF-8）
chcp 65001 > nul

rem 1) クリーン
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
rmdir /s /q __pycache__ 2>nul

rem 2) 依存インストール（ビルド環境用）
python -m pip install -U pip
python -m pip install -r requirements.txt

echo ビルド完了: dist\UnifiedOrder.exe を配布してください。
pause
rem 3) exe作成（ワンファイル、コンソール非表示、出力先はこのフォルダ）
python -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name "UnifiedOrder" ^
  --clean ^
  --distpath . ^
  main.py

echo ビルド完了: FULL\UnifiedOrder.exe を配布してください。
pause
