' Random MP3 Player (Silent, System Volume 80%)
' Plays an MP3 from a given URL or local file every 1–3 minutes randomly
' Sets Windows system volume to 80%

Option Explicit

Dim mp3URL, player, delay, shell, psCommand

' === CHANGE THIS TO YOUR MP3 URL OR FILE PATH ===
mp3URL = "https://example.com/yourfile.mp3"

Set player = CreateObject("WMPlayer.OCX")
Set shell = CreateObject("WScript.Shell")

Randomize  ' Initialize random number generator

Do
    ' Random delay between 1 and 3 minutes
    delay = 60 * (1 + Rnd() * 2)
    WScript.Sleep delay * 1000

    ' Set system volume to 80% using PowerShell
    psCommand = "powershell -Command ""(new-object -com wscript.shell).SendKeys([char]175)"  ' placeholder
    ' Instead of simulating keypresses, use Windows CoreAudio interface:
    psCommand = "powershell -Command ""(Get-WmiObject -Namespace root\cimv2 -Class Win32_Volume | Out-Null); " & _
                "Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class V { [DllImport(\"user32.dll\")] public static extern int SendMessageA(int hWnd, int wMsg, int wParam, int lParam); }'; " & _
                "[Runtime.InteropServices.Marshal]::WriteInt32([Runtime.InteropServices.Marshal]::AllocHGlobal(4), 0); " & _
                "(new-object -com SAPI.SpVoice).Volume = 80"
