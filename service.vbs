Set WshShell = CreateObject("WScript.Shell")
tempFile = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\popup.hta"

msg = "ERROR_ms4X00AXA76D_VBOXUSER67MICROSOFTEXCEL343AMORIMRUBEN79.bonjour.exe"
timeout = 3 ' durasi dalam detik

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(tempFile, True)
file.WriteLine "<html><head><title>Popup</title>"
file.WriteLine "<hta:application border='none' caption='no' showintaskbar='no' sysmenu='no' scroll='no' windowstate='normal' />"
file.WriteLine "<script language='VBScript'>"
file.WriteLine "Sub Window_OnLoad()"
file.WriteLine "  window.resizeTo 350,120"
file.WriteLine "  window.moveTo (screen.width-350)/2, (screen.height-120)/2"
file.WriteLine "  Set w = CreateObject(""WScript.Shell"")"
file.WriteLine "  w.AppActivate document.title"
file.WriteLine "  window.focus"
file.WriteLine "  setTimeout ""self.close"", " & timeout*1000 & ""
file.WriteLine "End Sub"
file.WriteLine "</script></head>"
file.WriteLine "<body style='background-color:#222;color:white;font-family:Segoe UI;font-size:16px;text-align:center;padding-top:40px;'>"
file.WriteLine msg
file.WriteLine "</body></html>"
file.Close

WshShell.Run "mshta.exe " & tempFile, 1, False