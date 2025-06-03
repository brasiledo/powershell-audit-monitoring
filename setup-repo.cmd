@echo off
setlocal EnableDelayedExpansion

:: Step 1: Get current folder name
for %%F in (.) do set folderName=%%~nxF

:: Step 2: Ask user to confirm GitHub repo name (default to folder name)
set /p repoName=Enter GitHub repo name [%folderName%]: 
if "%repoName%"=="" set repoName=%folderName%

:: Step 3: Initialize Git
echo Initializing Git repo...
git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
    git init
) else (
    echo Git repo already initialized.
)

:: Step 4: Add files and commit
git add .
git commit -m "Initial commit of PowerShell scripts" 2>nul
if errorlevel 1 (
    echo Nothing new to commit or commit already exists.
) else (
    echo Files committed.
)

:: Step 5: Set branch
git branch -M main 2>nul

:: Step 6: Add or update remote origin
git remote -v | findstr "origin" >nul
if errorlevel 1 (
    git remote add origin https://github.com/brasiledo/%repoName%.git
) else (
    echo Remote 'origin' already exists.
)

:: Step 7: Push to GitHub
echo Pushing to GitHub...
git push -u origin main

:: Done
echo.
echo ================================
echo Repo pushed to GitHub (if no errors).
echo URL: https://github.com/brasiledo/%repoName%
echo ================================
pause
