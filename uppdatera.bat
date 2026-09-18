@echo off
REM ---------------------------------------------------------------
REM  uppdatera.bat - hamtar senaste koden fran GitHub till den har
REM  mappen. Lagg filen i repo-mappen och dubbelklicka vid behov.
REM ---------------------------------------------------------------
cd /d "%~dp0"

git --version >nul 2>&1
if errorlevel 1 (
  echo Git hittades inte. Installera Git for Windows: https://git-scm.com/download/win
  pause
  exit /b 1
)

git rev-parse --git-dir >nul 2>&1
if errorlevel 1 (
  echo Den har mappen ar ingen git-klon.
  echo Klona repot istallet:
  echo    git clone https://github.com/ZilanCiftci/tv3analys.git
  pause
  exit /b 1
)

echo === Hamtar senaste andringar ===
git checkout main
if errorlevel 1 goto fel
git pull origin main
if errorlevel 1 goto fel

echo.
echo Uppdaterat till:
git log -1 --date=short --pretty=format:"  %%h  %%ad  %%s"
echo.
echo.
pause
exit /b 0

:fel
echo.
echo Hamtningen misslyckades (felkod %errorlevel%).
echo Har du andrat filer lokalt? Kor "git status" for att se vad som stoppar,
echo "git stash" for att lagga undan andringarna eller "git checkout -- ."
echo for att slanga dem.
pause
exit /b 1
