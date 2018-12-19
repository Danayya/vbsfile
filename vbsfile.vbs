Set wshShell = CreateObject( "WScript.Shell" )
wshShell.RegWrite  "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\BrowserEmulation\IntranetCompatibilityMode","1","REG_DWORD"

