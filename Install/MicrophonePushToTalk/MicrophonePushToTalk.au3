#NoTrayIcon
#RequireAdmin
#include "..\variables.au3"
#include <Misc.au3>
#include <MsgBoxConstants.au3>
Local $Title = $Name & " v" & $Version
If _Singleton($Name & "B9mmEGD9YhIF", 1) = 0 Then Exit MsgBox($MB_ICONWARNING, $Title, $Name & " is already running!")
If @AutoItX64 Then Exit MsgBox($MB_ICONWARNING + $MB_TOPMOST, $Title, "Use 32bit")
Opt("OnExitFunc", "ExitFunction")
Local $hDLL = DllOpen("user32.dll")
Local $KeyPressed = 0
Local $SettingsDir = @AppDataDir & "\" & $Name
#include <GUIConstantsEx.au3>
#Include <Constants.au3>
#include <WinAPIFiles.au3>
#include "_Beep.au3"
#include "_OpenSoundControlPanel.au3"
FileChangeDir(@ScriptDir)
AutoItSetOption("TrayAutoPause", 0)
AutoItSetOption("TrayMenuMode", 3)
AutoItSetOption("TrayOnEventMode", 1)
TraySetIcon(@ScriptDir & "\mic_off.ico")
TrayItemSetOnEvent(TrayCreateItem("Open Sound Control Panel"), "OpenSoundControlPanel")
TrayItemSetOnEvent(TrayCreateItem("Set Hot Key"), "SelectHotKey")
TrayItemSetOnEvent(TrayCreateItem("Set Mic Volume"), "SetVolume")
$BeepSounds = 1
If IniRead($SettingsDir & "\" & $Name & ".ini", "Settings", "BeepSounds", "") = "Off" Then $BeepSounds = 0
$BeepSoundsTrayItem = TrayCreateItem("Beep Sounds")
TrayItemSetOnEvent($BeepSoundsTrayItem, "BeepSounds")
If $BeepSounds Then TrayItemSetState($BeepSoundsTrayItem, $TRAY_CHECKED)
$BeepVolumeTrayItem = TrayCreateItem("Set Beep Volume")
TrayItemSetOnEvent($BeepVolumeTrayItem, "SetBeepVolume")
If Not $BeepSounds Then TrayItemSetState($BeepVolumeTrayItem, $TRAY_DISABLE)
TrayCreateItem("")
TrayItemSetOnEvent(TrayCreateItem("Exit"), "ExitScript")
TraySetToolTip($Title)
Local $Volume = 65536, $Percent = IniRead($SettingsDir & "\" & $Name & ".ini", "Settings", "MicVolume", ""), $HotKey = IniRead($SettingsDir & "\" & $Name & ".ini", "Settings", "HotKey", ""), $BeepVolume = IniRead($SettingsDir & "\" & $Name & ".ini", "Settings", "BeepVolume", "")

Func ExitFunction()
    UnmuteMic()
    DllClose($hDLL)
EndFunc

Func ExitScript()
    Exit ExitFunction()
EndFunc

Local $AllKeys = "Left mouse button|Right mouse button|Control-break processing|Middle mouse button|X1 mouse button|X2 mouse button|BACKSPACE|TAB|CLEAR|ENTER|SHIFT|CTRL|ALT|PAUSE|CAPS LOCK|ESC|SPACEBAR|PAGE UP|PAGE DOWN|END|HOME|LEFT ARROW|UP ARROW|RIGHT ARROW|DOWN ARROW|SELECT|PRINT|EXECUTE|PRINT SCREEN|INS|DEL" & _
"|0|1|2|3|4|5|6|7|8|9|A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z" & _
"|Left Windows|Right Windows|PopUp Menu Key|Numeric keypad 0|Numeric keypad 1|Numeric keypad 2|Numeric keypad 3|Numeric keypad 4|Numeric keypad 5|Numeric keypad 6|Numeric keypad 7|Numeric keypad 8|Numeric keypad 9" & _
"|Multiply|Add|Separator|Subtract|Decimal|Divide|F1|F2|F3|F4|F5|F6|F7|F8|F9|F10|F11|F12|F13|F14|F15|F16|F17|F19|F19|F20|F21|F22|F23|F24" & _
"|NUM LOCK|SCROLL LOCK|Left SHIFT|Right SHIFT|Left CONTROL|Right CONTROL|Left MENU|Right MENU|;|=|,|-|.|/|`|[|\|]"

Local $AllKeyCodes = "01|02|03|04|05|06|08|09|0C|0D|10|11|12|13|14|1B|20|21|22|23|24|25|26|27|28|29|2A|2B|2C|2D|2E|30|31|32|33|34|35|36|37|38|39|41|42|43|44|45|46|47|48|49|4A|4B|4C|4D|4E|4F|50|51|52|53|54|55|56|57|58|59|5A|5B|5C|5D|60|61|62|63|64|65|66|67|68|69|6A|6B|6C|6D|6E|6F|70|71|72|73|74|75|76|77|78|79|7A|7B|7C|7D|7E|7F|80H|81H|82H|83H|84H|85H|86H|87H|90|91|A0|A1|A2|A3|A4|A5|BA|BB|BC|BD|BE|BF|C0|DB|DC|DD"

Local $AllKeysArray = StringSplit($AllKeys, "|")
Local $AllKeyCodesArray = StringSplit($AllKeyCodes, "|")

If $BeepVolume = "" Then
    $BeepVolume = $DefaultBeepVolume
Else
    $BeepVolume = Floor(Number($BeepVolume))
EndIf

Func OpenSoundControlPanel()
    _OpenSoundControlPanel(1)
EndFunc

Func _IsPressedCode($char)
    Local $c = $char
    For $_ = 1 To 2
        For $i = 1 To $AllKeysArray[0]
            If $c = $AllKeysArray[$i] Then Return $AllKeyCodesArray[$i]
        Next
        $c = $DefaultHotKey
    Next
    ExitScript()
EndFunc

If $HotKey = "" Then $HotKey = $DefaultHotKey
Local $HotKey_IsPressedCode = _IsPressedCode($HotKey)

Func SelectHotKey()
    Local $hGUI = GUICreate("Select Hot Key")
    Local $_Combo = GUICtrlCreateCombo("", 10, 12)
    GUICtrlSetData(-1, $AllKeys, $HotKey)
    Local $_Button = GUICtrlCreateButton ("&OK", 250, 10)
    GUISetState()
    While 1
        Local $msg = GUIGetMsg()
        Switch $msg
            Case $GUI_EVENT_CLOSE
                ExitLoop
            Case $_Button
                $HotKey = GUICtrlRead($_Combo)
                $HotKey_IsPressedCode = _IsPressedCode($HotKey)
                DirCreate($SettingsDir)
                IniWrite($SettingsDir & "\" & $Name & ".ini", "Settings", "HotKey", $HotKey)
                ExitLoop
        EndSwitch
    WEnd
    GUIDelete($hGUI)
EndFunc

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

Func SetVolume()
    While 1
        Local $strNumber = InputBox($Title, @CRLF & @CRLF & @CRLF & "Enter the volume percent that you want your default recording device to be set to:", $Percent)
        If @error <> 0 Then Return
        Local $number = Floor(Number($strNumber))
        If $number >= 0 And $number <= 100 Then
            $Percent = $number
            ExitLoop
        EndIf
        MsgBox($MB_ICONWARNING, $Title, "You did not enter a valid number!")
    WEnd
    DirCreate($SettingsDir)
    IniWrite($SettingsDir & "\" & $Name & ".ini", "Settings", "MicVolume", $Percent)
    GetVolumeFromPercent()
EndFunc

Func BeepSounds()
    If $BeepSounds Then
        $BeepSounds = 0
        TrayItemSetState($BeepSoundsTrayItem, $TRAY_UNCHECKED)
        TrayItemSetState($BeepVolumeTrayItem, $TRAY_DISABLE)
        DirCreate($SettingsDir)
        IniWrite($SettingsDir & "\" & $Name & ".ini", "Settings", "BeepSounds", "Off")
    Else
        $BeepSounds = 1
        TrayItemSetState($BeepSoundsTrayItem, $TRAY_CHECKED)
        TrayItemSetState($BeepVolumeTrayItem, $TRAY_ENABLE)
        DirCreate($SettingsDir)
        IniWrite($SettingsDir & "\" & $Name & ".ini", "Settings", "BeepSounds", "On")
    EndIf
EndFunc

Func SetBeepVolume()
    While 1
        Local $strNumber = InputBox($Title, @CRLF & @CRLF & @CRLF & "Enter the volume percent that you want the beep volume to be set to:", $BeepVolume)
        If @error <> 0 Then Return
        Local $number = Floor(Number($strNumber))
        If $number >= 0 And $number <= 100 Then
            $BeepVolume = $number
            ExitLoop
        EndIf
        MsgBox($MB_ICONWARNING, $Title, "You did not enter a valid number!")
    WEnd
    DirCreate($SettingsDir)
    IniWrite($SettingsDir & "\" & $Name & ".ini", "Settings", "BeepVolume", $BeepVolume)
EndFunc

Func MuteMic()
    Run("NirCmd.exe  mutesysvolume 1 default_record", @ScriptDir, @SW_HIDE)
    RunWait("NirCmd.exe setsysvolume 0 default_record", @ScriptDir, @SW_HIDE)
EndFunc

Func UnmuteMic()
    Run("NirCmd.exe  mutesysvolume 0 default_record", @ScriptDir, @SW_HIDE)
    RunWait("NirCmd.exe setsysvolume " & $Volume & " default_record", @ScriptDir, @SW_HIDE)
EndFunc

Func PushToTalk()
    If $KeyPressed Then Return
    $KeyPressed = 1
    UnmuteMic()
    TraySetIcon(@ScriptDir & "\mic_on.ico")
    If $BeepSounds Then _Beep(750, 100, $BeepVolume)
    While _IsPressed($HotKey_IsPressedCode, $hDLL)
        Sleep(10)
    WEnd
    MuteMic()
    TraySetIcon(@ScriptDir & "\mic_off.ico")
    $KeyPressed = 0
    If $BeepSounds Then _Beep(400, 100, $BeepVolume)
EndFunc

GetVolumeFromPercent()
MuteMic()

While 1
    if _IsPressed($HotKey_IsPressedCode) Then PushToTalk()
    Sleep(10)
WEnd

ExitScript()