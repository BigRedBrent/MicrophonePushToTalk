#NoTrayIcon
#include "variables.au3"
#include <Misc.au3>
#include <MsgBoxConstants.au3>
Global $Title = $Name & " v" & $Version & " Installer"
If _Singleton($Name & " Installer" & "346473046bWe46", 1) = 0 Then Exit MsgBox($MB_ICONWARNING, $Title, $Name & " Installer" & " is already running!")
#include <WinAPIFiles.au3>
#include <GUIConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <StringConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
Local $SettingsDir = @AppDataDir & "\" & $Name
Local $Volume = 65536, $Percent = IniRead($SettingsDir & "\" & $Name & ".ini", "Settings", "MicVolume", "")

Local $InstallDir = @ProgramFilesDir & "\" & $Name, $RegLocation = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & $Name, $InstallLocation = StringRegExpReplace(RegRead($RegLocation, "InstallLocation"), "\\+$", "")

Func GetInstallLocation($dir = $InstallDir)
    Local $GUI = GUICreate($Title, 434, 142, -1, -1, -1, $WS_EX_TOPMOST)
    Local $Input = GUICtrlCreateInput($dir, 16, 56, 329, 21)
    Local $ButtonChange = GUICtrlCreateButton("Change", 352, 54, 75, 25)
    Local $ButtonOK = GUICtrlCreateButton("OK", 168, 104, 75, 25, $BS_DEFPUSHBUTTON)
    Local $ButtonCancel = GUICtrlCreateButton("Cancel", 264, 104, 75, 25)
    Local $Label = GUICtrlCreateLabel("Select install location:", 16, 24, 332, 25)
    GUISetState()
    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE
                Exit
            Case $ButtonChange
                Local $sFileSelectFolder = FileSelectFolder("Select install location:", StringRegExpReplace(StringRegExpReplace(GUICtrlRead($Input), "\\+$", ""), "\\" & $Name & "$", ""), 0, "", $GUI)
                If @error = 0 Then GUICtrlSetData($Input, StringRegExpReplace($sFileSelectFolder, "\\+$", "") & "\" & $Name)
            Case $ButtonOK
                Local $sCurrInput = StringRegExpReplace(GUICtrlRead($Input), "\\+$", "")
                If StringRegExp($sCurrInput, "\\" & $Name & "$") And FileExists(StringRegExpReplace($sCurrInput, "\\" & $Name & "$", "")) Then
                    GUIDelete()
                    $InstallLocation = RegRead($RegLocation, "InstallLocation")
                    If $InstallLocation <> $sCurrInput And $InstallLocation <> "" Then
                        Local $UninstallString = StringReplace(RegRead($RegLocation, "UninstallString"), '"', "")
                        If StringInStr($UninstallString, $InstallLocation) And FileExists($UninstallString) Then
                            MsgBox($MB_ICONWARNING + $MB_TOPMOST, $Title, "Uninstall previous installation!")
                            RunWait($UninstallString, $InstallLocation)
                            $InstallLocation = RegRead($RegLocation, "InstallLocation")
                            If $InstallLocation <> "" Then Return GetInstallLocation($InstallLocation)
                            Return GetInstallLocation($sCurrInput)
                        EndIf
                    EndIf
                    Return $sCurrInput
                EndIf
            Case $ButtonCancel
                Exit
        EndSwitch
    WEnd
EndFunc

If $InstallLocation <> "" And StringRegExp($InstallLocation, "\\" & $Name & "$") And FileExists($InstallLocation) Then
    $InstallDir = $InstallLocation
    If MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_TOPMOST, $Title, "Install over existing installation?" & @CRLF & @CRLF & '"' & $InstallDir & '"') <> $IDYES Then Exit
Else
    $InstallDir = GetInstallLocation()
EndIf

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
    RunWait(@ScriptDir & "\" & $Name & "\NirCmd.exe  mutesysvolume 0 default_record", @ScriptDir & "\" & $Name, @SW_HIDE)
    RunWait(@ScriptDir & "\" & $Name & "\NirCmd.exe setsysvolume " & $Volume & " default_record", @ScriptDir & "\" & $Name, @SW_HIDE)
WEnd


If Not DirCopy($Name, $InstallDir, 1) Then Exit MsgBox($MB_ICONWARNING, $Title, "An error occurred while copying files to the programs folder!")

If Not RegWrite($RegLocation, "DisplayName", "REG_SZ", $Name & " v" & $Version) Or Not RegWrite($RegLocation, "DisplayVersion", "REG_SZ", $Version) Or Not RegWrite($RegLocation, "Publisher", "REG_SZ", "BigRedBrent") Or Not RegWrite($RegLocation, "DisplayIcon", "REG_SZ", $InstallDir & "\Uninstall.exe") Or Not RegWrite($RegLocation, "UninstallString", "REG_SZ", '"' & $InstallDir & '\Uninstall.exe"') Or Not RegWrite($RegLocation, "InstallLocation", "REG_SZ", $InstallDir) Then Exit MsgBox($MB_ICONWARNING, $Title, "An error occurred while creating the uninstaller registry keys!")

If Not RunWait('schtasks /query /tn "' & $Name & '"', "", @SW_HIDE) And RunWait('schtasks /delete /tn "' & $Name & '" /f', "", @SW_HIDE) Then Exit MsgBox($MB_ICONWARNING, $Title, "Failed to delete scheduled task!")

If MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_TOPMOST, $Title, "Run on startup?") = $IDYES Then
    Local $XMLText = _
    '<?xml version="1.0" encoding="UTF-16"?>' & @CRLF & _
    '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">' & @CRLF & _
    '<RegistrationInfo>' & @CRLF & _
    '  <Author>SYSTEM</Author>' & @CRLF & _
    '  <URI>\Neverwinter Invoke Bot Start Up</URI>' & @CRLF & _
    '</RegistrationInfo>' & @CRLF & _
    '<Triggers>' & @CRLF & _
    '  <LogonTrigger>' & @CRLF & _
    '    <Enabled>true</Enabled>' & @CRLF & _
    '  </LogonTrigger>' & @CRLF & _
    '</Triggers>' & @CRLF & _
    '<Principals>' & @CRLF & _
    '  <Principal id="Author">' & @CRLF & _
    '    <LogonType>InteractiveToken</LogonType>' & @CRLF & _
    '    <RunLevel>HighestAvailable</RunLevel>' & @CRLF & _
    '  </Principal>' & @CRLF & _
    '</Principals>' & @CRLF & _
    '<Settings>' & @CRLF & _
    '  <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>' & @CRLF & _
    '  <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>' & @CRLF & _
    '  <StopIfGoingOnBatteries>true</StopIfGoingOnBatteries>' & @CRLF & _
    '  <AllowHardTerminate>true</AllowHardTerminate>' & @CRLF & _
    '  <StartWhenAvailable>false</StartWhenAvailable>' & @CRLF & _
    '  <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>' & @CRLF & _
    '  <IdleSettings>' & @CRLF & _
    '    <StopOnIdleEnd>true</StopOnIdleEnd>' & @CRLF & _
    '    <RestartOnIdle>false</RestartOnIdle>' & @CRLF & _
    '  </IdleSettings>' & @CRLF & _
    '  <AllowStartOnDemand>true</AllowStartOnDemand>' & @CRLF & _
    '  <Enabled>true</Enabled>' & @CRLF & _
    '  <Hidden>false</Hidden>' & @CRLF & _
    '  <RunOnlyIfIdle>false</RunOnlyIfIdle>' & @CRLF & _
    '  <WakeToRun>false</WakeToRun>' & @CRLF & _
    '  <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>' & @CRLF & _
    '  <Priority>7</Priority>' & @CRLF & _
    '</Settings>' & @CRLF & _
    '<Actions Context="Author">' & @CRLF & _
    '  <Exec>' & @CRLF & _
    '      <Command>' & $InstallDir & '\' & $Name & '.exe' & '</Command>' & @CRLF & _
    '      <Arguments></Arguments>' & @CRLF & _
    '  </Exec>' & @CRLF & _
    '</Actions>' & @CRLF & _
    '</Task>'

    Local $XMLFile = _WinAPI_GetTempFileName(@TempDir)

    If Not FileWrite($XMLFile, $XMLText) Then Exit MsgBox($MB_ICONWARNING, $Title, "An error occurred while writing the temporary XML file!")

    If RunWait('schtasks /create /xml "' & $XMLFile & '" /tn "' & $Name & '"', "", @SW_HIDE) Then
        FileDelete($XMLFile)
        Exit MsgBox($MB_ICONWARNING, $Title, "Failed to create scheduled task!")
    Else
        FileDelete($XMLFile)
    EndIf
EndIf

MsgBox(0, $Title, "Install successful." & @CRLF & @CRLF & '"' & $InstallDir & '"')
If MsgBox($MB_YESNO + $MB_ICONQUESTION + $MB_TOPMOST, $Title, "Run now?") = $IDYES Then Exit ShellExecute($InstallDir & "\" & $Name & ".exe", "", $InstallDir)
