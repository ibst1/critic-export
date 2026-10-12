@echo off
rem critic-export: the file to start.
rem   Double-click:              start the Obsidian hotkeys and set up autostart
rem                              and the Explorer right-click menu.
rem   critic-export note.md ...  export (critic-export --help for all options).
rem Everything else is in the app folder.
if "%~1"=="" goto start
python "%~dp0app\critic_export.py" %*
exit /b %errorlevel%
:start
python "%~dp0app\critic_export.py" --start
