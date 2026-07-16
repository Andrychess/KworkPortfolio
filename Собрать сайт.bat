@echo off
cd /d "%~dp0"
python _sync.py
if errorlevel 1 (
  echo.
  echo Ошибка сборки. Нужны Python и зависимости: pip install pymupdf
  pause
  exit /b 1
)
echo.
echo Готово. Закоммитьте index.html и Works/ в git.
pause
