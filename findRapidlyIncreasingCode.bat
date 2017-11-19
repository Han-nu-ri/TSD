reg add HKLM\Software\Classes\cmdfile\shell /ve /t REG_SZ /d "runas" /f

python findRapidlyIncreasingCode.py

reg delete HKLM\Software\Classes\cmdfile\shell /ve /f