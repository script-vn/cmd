REG ADD HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer /v DisallowRun /t REG_DWORD /reg:32 /d 1
REG ADD HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun /v 2 /t REG_SZ /d Zalo.exe
