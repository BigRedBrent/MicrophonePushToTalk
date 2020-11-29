#include-once
;~ Supported OS: Windows 7, Windows 8, Windows 8.1, Windows 10
Global Const $MicrosoftSound = "Microsoft.Sound"
Global Const $GUID_MicrosoftSound = "{F2DDFC82-8F12-4CDD-B7DC-D4FE1425AA4D}"

Global Const $sCLSID_OpenControlPanel = "{06622D85-6856-4460-8DE1-A81921B41C4B}"
Global Const $sIID_IOpenControlPanel =  "{D11AD862-66DE-4DF4-BF6C-1F5621996AF1}"
Global Const $sTagIOpenControlPanel = "Open hresult(wstr;wstr;ptr);GetPath hresult(wstr;wstr;uint);GetCurrentView hresult(int*)"

Local $oOpenControlPanel = ObjCreateInterface($sCLSID_OpenControlPanel, $sIID_IOpenControlPanel,  $sTagIOpenControlPanel)

Func _OpenSoundControlPanel($RecordingTab = 0)
    If $RecordingTab Then
        $oOpenControlPanel.Open($MicrosoftSound,"1",Null)
    Else
        $oOpenControlPanel.Open($MicrosoftSound,"",Null)
    EndIf
EndFunc
