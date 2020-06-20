@echo off
D:
cd \xampp\htdocs\pyScripts
echo updating Tiles Remediation sheet...
call activate.bat
tiles_remediation.py
echo Done!
::pause
