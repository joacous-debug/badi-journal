@echo off
title Badi-Journal
cd /d "%~dp0"
where py >nul 2>nul && (py badi_journal.py) || (python badi_journal.py)
pause
