#NoTrayIcon
#RequireAdmin
#include "..\variables.au3"
Local $Title = $Name & " v" & $Version & " Uninstaller"
#include <MsgBoxConstants.au3>
If Not @Compiled Then Exit MsgBox($MB_ICONWARNING + $MB_TOPMOST, $Title, "The script must be a compiled exe to work correctly!")
#include "_SelfDelete.au3"
Local $SettingsDir = @AppDataDir & "\" & $Name
Local $Volume = 65536, $Percent = IniRead($SettingsDir & "\" & $Name & ".ini", "Settings", "MicVolume", "")

If MsgBox($MB_YESNO + $MB_ICONQUESTION, $Title, "Do you want to uninstall " & $Name & " v" & $Version & "?") <> $IDYES Then Exit

If Not RunWait('schtasks /query /tn "' & $Name & '"', "", @SW_HIDE) And RunWait('schtasks /delete /tn "' & $Name & '" /f', "", @SW_HIDE) Then Exit MsgBox($MB_ICONWARNING, $Title, "Failed to delete scheduled task!")

Func GetVolumeFromPercent()
    If $Percent = "" Then
        $Percent = $DefaultVolume
    Else
        $Percent = Floor(Number($Percent))
    EndIf
    If $Percent < 0 then
        $Percent = 0
    ElseIf $Percent > 100 Then
        $Percent = 100
    EndIf
    If $Percent <= 0 then
        $Volume = 0
    ElseIf $Percent < 100 Then
        $Volume = Round(65536 * $Percent / 100)
    Else
        $Volume = 65536
    EndIf
EndFunc

While ProcessExists($Name & ".exe")
    ProcessClose($Name & ".exe")
    Sleep(1000)
    GetVolumeFromPercent()
    RunWait("NirCmd.exe  mutesysvolume 0 default_record", @ScriptDir, @SW_HIDE)
    RunWait("NirCmd.exe setsysvolume " & $Volume & " default_record", @ScriptDir, @SW_HIDE)
WEnd

Local $delete = StringSplit($Name & ".exe,NirCmd.exe", ",")
FileChangeDir(@ScriptDir)
For $i = 1 To $delete[0]
    If FileExists($delete[$i]) And Not FileDelete($delete[$i]) Then Exit MsgBox($MB_ICONWARNING, $Title, "Failed to delete " & '"' & $delete[$i] & '"' & "!" & @CRLF & @CRLF & "Check to see if " & '"' & $delete[$i] & '"' & " is currently running and try again.")
Next

Local $RegLocation = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $Name

If ( RegRead($RegLocation, "DisplayName") <> "" Or RegRead($RegLocation, "DisplayVersion") <> "" Or RegRead($RegLocation, "Publisher") <> "" Or RegRead($RegLocation, "DisplayIcon") <> "" Or RegRead($RegLocation, "UninstallString") <> "" Or RegRead($RegLocation, "InstallLocation") <> "" ) And Not RegDelete($RegLocation) Then Exit MsgBox($MB_ICONWARNING, $Title, "Failed to delete uninstaller registry key!")

If FileExists(@AppDataDir & "\" & $Name) And MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_TOPMOST, $Title, "Keep settings files?") <> $IDYES Then DirRemove(@AppDataDir & "\" & $Name, 1)

DirRemove(@ProgramsCommonDir & "\" & $Name, 1)

_SelfDelete(5, 1, 1)
If @error Then Exit MsgBox($MB_ICONWARNING, "_SelfDelete()", "The script must be a compiled exe to work correctly.")
