@echo off
echo ========================================
echo   Push Files to GitHub
echo ========================================
echo.

REM Check if git is installed
git --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Git is not installed!
    echo Please download and install Git from: https://git-scm.com/download/win
    echo.
    pause
    exit /b 1
)

echo Git is installed.
echo.

REM Check if .git exists
if not exist .git (
    echo Initializing git repository...
    git init
    echo.
)

REM Add Python and batch files
echo Adding Python and batch files...
git add *.py
git add *.bat
git add .gitignore
echo.

REM Check if there are changes to commit
git diff --cached --quiet
if errorlevel 1 (
    echo Committing files...
    git commit -m "Add experiment analyzer scripts and batch files"
    echo.
) else (
    echo No changes to commit.
    echo.
)

REM Ask for GitHub repository details
set /p GITHUB_USER="Enter your GitHub username: "
set /p REPO_NAME="Enter your repository name: "

REM Check if remote already exists
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    echo Adding remote repository...
    git remote add origin https://github.com/%GITHUB_USER%/%REPO_NAME%.git
) else (
    echo Remote already exists. Updating...
    git remote set-url origin https://github.com/%GITHUB_USER%/%REPO_NAME%.git
)
echo.

REM Set branch to main
git branch -M main

REM Push to GitHub
echo Pushing to GitHub...
echo You may be prompted for your GitHub username and password/token.
echo.
git push -u origin main

if errorlevel 1 (
    echo.
    echo ERROR: Push failed!
    echo.
    echo Common issues:
    echo 1. Repository doesn't exist on GitHub - create it first
    echo 2. Authentication failed - use a Personal Access Token instead of password
    echo    (Create one at: https://github.com/settings/tokens)
    echo 3. Wrong repository name or username
    echo.
) else (
    echo.
    echo ========================================
    echo   Successfully pushed to GitHub!
    echo ========================================
    echo.
)

pause

