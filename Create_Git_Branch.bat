@echo off
setlocal

echo ========================================
echo       Git Initialize and Push
echo ========================================
echo.

:: Ask for repository URL
set /p REPO_URL=Enter Git repository URL: 

:: Ask for branch name
set /p BRANCH_NAME=Enter new branch name: 

:: Ask for commit message
set /p COMMIT_MSG=Enter commit message: 

echo.
echo [1/6] Checking Git repository...

if not exist ".git" (
    echo Git repository not found. Initializing...
    git init
) else (
    echo Git repository already initialized.
)

echo.
echo [2/6] Configuring remote...

git remote get-url origin >nul 2>&1

if errorlevel 1 (
    git remote add origin "%REPO_URL%"
    echo Remote 'origin' added.
) else (
    echo Remote 'origin' already exists.
)

echo.
echo [3/6] Creating branch "%BRANCH_NAME%"...

git checkout -b "%BRANCH_NAME%"

if errorlevel 1 (
    echo.
    echo Branch creation failed.
    pause
    exit /b 1
)

echo.
echo [4/6] Adding files...

git add .

echo.
echo [5/6] Committing changes...

git commit -m "%COMMIT_MSG%"

if errorlevel 1 (
    echo.
    echo Commit failed or there are no changes to commit.
)

echo.
echo [6/6] Pushing branch...

git push -u origin "%BRANCH_NAME%"

if errorlevel 1 (
    echo.
    echo ========================================
    echo PUSH FAILED
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo SUCCESS!
echo ========================================
echo Branch: %BRANCH_NAME%
echo Repository: %REPO_URL%
echo.
echo The branch has been pushed successfully.
echo.

pause