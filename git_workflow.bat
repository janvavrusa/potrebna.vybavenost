@echo off
setlocal
cd /d "%~dp0" || goto :error
set "REMOTE_URL=https://github.com/janvavrusa/potrebna.vybavenost.git"

echo [1/5] Kontrola Git a inicializace repozitare projektu...
where git >nul 2>&1 || (
  echo Git neni dostupny v PATH.
  goto :error
)
git init || goto :error

for /f "delims=" %%N in ('git config user.name 2^>nul') do set "GIT_USER_NAME=%%N"
for /f "delims=" %%E in ('git config user.email 2^>nul') do set "GIT_USER_EMAIL=%%E"
if not defined GIT_USER_NAME (
  echo Neni nastaveno git user.name. Nastavte ho prikazem: git config --global user.name "Vase jmeno"
  goto :error
)
if not defined GIT_USER_EMAIL (
  echo Neni nastaveno git user.email. Nastavte ho prikazem: git config --global user.email "vas@email.cz"
  goto :error
)

echo [2/5] Kontrola ciloveho GitHub repozitare...
git remote get-url origin >nul 2>&1
if errorlevel 1 goto :add_remote
for /f "delims=" %%U in ('git remote get-url origin 2^>nul') do set "ORIGIN_URL=%%U"
if /I not "%ORIGIN_URL%"=="%REMOTE_URL%" (
  echo Existujici origin ukazuje na jiny repozitar:
  echo %ORIGIN_URL%
  echo Ocekavany cil:
  echo %REMOTE_URL%
  echo Skript push z bezpecnostnich duvodu neprovede.
  goto :error
)
goto :remote_ready

:add_remote
git remote add origin "%REMOTE_URL%" || goto :error

:remote_ready
echo [3/5] Pridavani souboru z teto slozky do indexu...
git add -A -- . || goto :error

git diff --cached --quiet
if errorlevel 2 goto :error
if errorlevel 1 goto :commit_changes
goto :empty_commit

:commit_changes
echo [4/5] Vytvareni commitu se zmenami...
git commit -m "Update App Potrebna Vybavenost" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>" || goto :error
goto :push

:empty_commit
echo [4/5] Repo je ciste, vytvarim retry commit pro GitHub Pages...
git commit --allow-empty -m "Retry GitHub Pages deployment" -m "Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>" || goto :error

:push
echo [5/5] Odesilani na GitHub...
for /f "delims=" %%B in ('git branch --show-current') do set "BRANCH=%%B"
if not defined BRANCH (
  echo Nelze zjistit aktualni vetev.
  goto :error
)
git push --set-upstream origin "%BRANCH%" || goto :error

echo Hotovo. Zkontrolujte GitHub Actions, zda deployment uspesne dobehl.
pause
exit /b 0

:error
echo.
echo CHYBA: Deploy se nepodaril. Zkontrolujte vystup vyse.
pause
exit /b 1
