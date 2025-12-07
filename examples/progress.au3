#include "../src/toast.au3"

Global Const $sAppName = @ScriptName
Global $tCLSID = _Toast_CoCreateGuid()
Global $sGUID = _WinAPI_StringFromGUID($tCLSID)
ConsoleWrite("app CLSID: " & $sGUID & @CRLF)

_Toast_Initialize($sAppName, $tCLSID, OnToastActivation, "AutoIt Toast Example", @TempDir & "\e25dbe211ddfb027fcb8271d833159fc.png")

; https://learn.microsoft.com/en-us/windows/win32/api/notificationactivationcallback/nf-notificationactivationcallback-inotificationactivationcallback-activate
Func OnToastActivation($pSelf, $appUserModelId, $invokedArgs, $data, $count)
    _GUICtrlRichEdit_AppendText($hRich, _
        "Toast activated!" & @CRLF _
        & "    " & "appUserModelId: " & $appUserModelId & @CRLF _
        & "    " & "invokedArgs: " & $invokedArgs & @CRLF _
        )

    Return $_Toast_S_OK
EndFunc

Opt("GuiOnEventMode", 1)

Global $hWnd = GUICreate("Toast progress example", 700, 320)
GUISetOnEvent($GUI_EVENT_CLOSE, "GUI_CLOSE")

Global $hButton01 = GUICtrlCreateButton("Create toast", 10, 10, 200, 120)
GUICtrlSetOnEvent(-1, "CreateToast")

Global $hButton02 = GUICtrlCreateButton("Update Progress", 10, 140, 200, 120)
GUICtrlSetOnEvent(-1, "UpdateProgress")

Global $hRich = _GUICtrlRichEdit_Create($hWnd, "", 220, 10, 470, 300)

GUISetState()

While 1
    Sleep(10)
WEnd

Func CreateToast()
    ; https://learn.microsoft.com/en-us/windows/apps/design/shell/tiles-and-notifications/adaptive-interactive-toasts?tabs=xml
    Local $sToast = _
        '<toast scenario="reminder" activationType="background" launch="action=mainContent" duration="short" useButtonStyle="true">' & _
        '  <visual>' & _
        '    <binding template="ToastGeneric">' & _
        '      <text>Sample toast</text>' & _
        '      <text>Sample content</text>' & _
        '      <progress' & _
        '        title="Weekly playlist"' & _
        '        value="0.5"' & _
        '        _valueStringOverride="15/26 songs"' & _
        '        status="Downloading..."' & _
        '      />' & _
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

Func UpdateProgress()
    Local $oToastNotificationManager = __Toast_ToastNotificationManager()
    Local $pToastNotifier = 0
    Local $hr = $oToastNotificationManager.CreateToastNotifier($pToastNotifier)
    Local $oToastNotifier2 = ObjCreateInterface($pToastNotifier, "{354389C6-7C01-4BD5-9C20-604340CD2B74}", $sIInspectable & "UpdateWithTagAndGroup HRESULT(PTR;PTR;PTR;PTR);UpdateWithTag HRESULT(PTR;PTR;PTR)")
    
    Local $pNotificationData = CreateNotificationData()
    Local $oNotificationData = ObjCreateInterface($pNotificationData, "{9FFD2312-9D6A-4AAF-B6AC-FF17F0C1F280}", $sIInspectable & "Values HRESULT(PTR*);SequenceNumber HRESULT(UINT*);SetSequenceNumber HRESULT(UINT);")

    Local $pValues = 0
    Local $hr = $oNotificationData.Values($pValues)

    Local $pInsert = __Toast_VTable_get($pValues, 4)
    Local $pReplaced = 0
    Local $pKey = "value"
    $hr = __Toast_WindowsCreateString($pKey, $pKey)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf
    Local $pValue = "0.9"
    $hr = __Toast_WindowsCreateString($pValue, $pValue)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf
    $hr = DllCallAddress("HRESULT", $pInsert, "PTR", $pKey, "PTR", $pValue, "BOOLEAN*", $pReplaced)
    __Toast_WindowsDeleteString($pKey)
    __Toast_WindowsDeleteString($pValue)
    $oNotificationData.SetSequenceNumber(1)

    ; FIXME: apply new notificationData to the existing toast
EndFunc

Func CreateNotificationData()
    Local $classId = "Windows.UI.Notifications.NotificationData"
    Local $pInspectable = 0

    Local $hr = RoActivateInstance($classId, $pInspectable)

    If $hr <> 0 Then Return SetError($hr)

    Local $pINotificationData = 0
    $hr = __Toast_QueryInterface($pInspectable, $UIID_IXmlDocument, $pINotificationData)

    If @error <> 0 Then
        Local $error = @error, $extended = @extended
        __Toast_IUnknown_Release($pInspectable)
        Return SetError($error, $extended, 0)
    EndIf

    __Toast_IUnknown_Release($pInspectable)

    Return $pINotificationData
EndFunc

Func __Toast_VTable_get($pInterface, $iMethod)
    Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

    Local $methodOffset = (@AutoItX64 ? 8 : 4) * $iMethod
    
    Return DllStructGetData(DllStructCreate("ptr", $pVTable + $methodOffset), 1)
EndFunc
