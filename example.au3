#include "./src/toast.au3"
#include <GuiRichEdit.au3>
#include <GUIConstantsEx.au3>
#include <InetConstants.au3>

Opt("GuiOnEventMode", 1)

Global $hWnd = GUICreate("Toast example", 700, 320)
GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_CLOSE")

Global $hButton01 = GUICtrlCreateButton("From template", 10, 10, 200, 120)
GUICtrlSetOnEvent(-1, "ToastFromTemplateExample")

Global $hButton02 = GUICtrlCreateButton("From XML string", 10, 140, 200, 120)
GUICtrlSetOnEvent(-1, "ToastFromXmlString")

Global $hRich = _GUICtrlRichEdit_Create($hWnd, "", 220, 10, 470, 300)

GUISetState()

While 1
    Sleep(10)
Wend

Func ToastFromTemplateExample()
    Local $oXml = _Toast_CreateToastTemplateXmlDocument()

    Local $pToast = _Toast_CreateToastNotificationFromXmlObject($oXml)

    _Toast_Show($pToast)
EndFunc

Func DownloadImage()
    If FileExists(@TempDir & "\e21cd29c9fb51c3a5b82f009ec33fc997d2edd1ece931e8568f37e205c445778.jpeg") Then Return
    _GUICtrlRichEdit_AppendText($hRich, "Trying to download avatar image from gravatar..." & @CRLF)
    Local $iBytes = InetGet("https://gravatar.com/avatar/e21cd29c9fb51c3a5b82f009ec33fc997d2edd1ece931e8568f37e205c445778", @TempDir & "\e21cd29c9fb51c3a5b82f009ec33fc997d2edd1ece931e8568f37e205c445778.jpeg", $INET_FORCEBYPASS)
    Local $error = @error
    If  @error <> 0 Then
        _GUICtrlRichEdit_AppendText($hRich, "Failed to download image" & @CRLF)
        Return
    EndIf
    _GUICtrlRichEdit_AppendText($hRich, "Done! "& $iBytes & " bytes downloaded" & @CRLF)
EndFunc

Func ToastFromXmlString()
    DownloadImage()

    ; https://learn.microsoft.com/en-us/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml
    Local $sToast = _
        '<toast scenario="reminder" activationType="background" launch="action=mainContent" duration="short" useButtonStyle="true">' & _
        '  <visual>' & _
        '    <binding template="ToastGeneric">' & _
        '      <text>Sample toast</text>' & _
        '      <text>Sample content</text>' & _
        '      <image placement="appLogoOverride" src="file://' & @TempDir & '\e21cd29c9fb51c3a5b82f009ec33fc997d2edd1ece931e8568f37e205c445778.jpeg" hint-crop="circle"/>' & _
        '    </binding>' & _
        '  </visual>' & _
        '  <actions>' & _
        '    <action' & _
        '      content="Click me"' & _
        '      activationType="background"' & _
        '      hint-buttonStyle="Success"' & _
        '      arguments="action=click_me"/>' & _
        '    <action' & _
        '      content="Dismiss"' & _
        '      activationType="system"' & _
        '      hint-buttonStyle="Critical"' & _
        '      arguments="dismiss"/>' & _
        '  </actions>' & _
        "</toast>"

    Local $pToast = _Toast_CreateToastNotificationFromXmlString($sToast)

    If @error <> 0 Then
        _GUICtrlRichEdit_AppendText($hRich, _WinAPI_GetErrorMessage(@error))
        Return
    EndIf

    _Toast_Show($pToast)
EndFunc

Func GUI_CLOSE()
    Exit
EndFunc
