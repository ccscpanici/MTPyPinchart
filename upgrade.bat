@echo off

rem *** these lines are needed in the initial upgrade file ***
rem echo "Changing to installation directory"
rem cd C:\CCS\MTPyPinchart

echo "pulling new code base from github"
git pull

echo "re-running dependency installation"
"C:\Program Files\Python39\Scripts\pip.exe" install -r requirements.txt

echo "new pinchart script installation completed"

echo on