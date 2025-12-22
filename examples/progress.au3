#include "../src/toast.au3"
#include <GuiRichEdit.au3>
#include <GUIConstantsEx.au3>

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
        '        value="{progress1}"' & _
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

    Local $pIToastNotification2 = 0
    $hr = __Toast_QueryInterface($pToast, "{9DFB9FD1-143A-490E-90BF-B9FBA7132DE7}", $pIToastNotification2)

    If @error <> 0 Or $hr <> 0 Then
        _GUICtrlRichEdit_AppendText(StringFormat("Failed to get IToastNotification2 COM interface!\n\t@error: %s\n\tHRESULT: %s\n", @error, $hr))
    Else
        Local $aRet = Null
        Local $pTag = "demo-tag"
        $hr = __Toast_WindowsCreateString($pTag, $pTag)
        $pSetTag = __Toast_VTable_get($pIToastNotification2, 6)
        $aRet = DllCallAddress("LONG", $pSetTag, "PTR", $pIToastNotification2, "PTR", $pTag)
        __Toast_WindowsDeleteString($pTag)

        Local $pGroup = "demo-group"
        $hr = __Toast_WindowsCreateString($pGroup, $pGroup)
        $pSetGroup = __Toast_VTable_get($pIToastNotification2, 8)
        DllCallAddress("LONG", $pSetGroup, "PTR", $pIToastNotification2, "PTR", $pGroup)
        __Toast_WindowsDeleteString($pGroup)

        __Toast_IUnknown_Release($pIToastNotification2)

        Local $pIToastNotification4 = 0
        $hr = __Toast_QueryInterface($pToast, "{15154935-28EA-4727-88E9-C58680E2D118}", $pIToastNotification4)
        If @error <> 0 Or $hr <> 0 Then
            _GUICtrlRichEdit_AppendText("Failed to get IToastNotification4 COM interface! default progressbar value will not be set!"&@CRLF)
        Else
            $pNotificationData = CreateNotificationData()
            Local $oNotificationData = ObjCreateInterface($pNotificationData, "{9FFD2312-9D6A-4AAF-B6AC-FF17F0C1F280}", $sIInspectable & "Values HRESULT(PTR*);SequenceNumber HRESULT(UINT*);SetSequenceNumber HRESULT(UINT);")
            Local $pValues = 0
            Local $hr = $oNotificationData.Values($pValues)

            Local $pInsert = __Toast_VTable_get($pValues, 10)
            Local $pReplaced = 0
            Local $pKey = "progress1"
            $hr = __Toast_WindowsCreateString($pKey, $pKey)
            Local $pValue = "0.5"
            ;$pValue = "indeterminate"
            $hr = __Toast_WindowsCreateString($pValue, $pValue)
            $hr = DllCallAddress("LONG", $pInsert, "PTR", $pValues, "PTR", $pKey, "PTR", $pValue, "BOOLEAN*", 0)
            __Toast_WindowsDeleteString($pKey)
            __Toast_WindowsDeleteString($pValue)

            $pSetData = __Toast_VTable_get($pIToastNotification4, 7)
            DllCallAddress("LONG", $pSetData, "PTR", $pIToastNotification4, "PTR", $pNotificationData)
        EndIf
    EndIf

    _Toast_Show($pToast)
EndFunc

Func UpdateProgress()
    Local $oToastNotificationManager = __Toast_ToastNotificationManager()
    Local $pToastNotifier = 0
    Local $sAppId = @ScriptName
    Local $hr = __Toast_WindowsCreateString($sAppId, $sAppId)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf
    Local $hr = $oToastNotificationManager.CreateToastNotifierWithId($sAppId, $pToastNotifier)
    If Not ($hr = 0) Then
        Return SetError($hr)
    EndIf

    Local $oToastNotifier2 = ObjCreateInterface($pToastNotifier, "{354389C6-7C01-4BD5-9C20-604340CD2B74}", $sIInspectable & "UpdateWithTagAndGroup HRESULT(PTR;PTR;PTR;PTR*);UpdateWithTag HRESULT(PTR;PTR;PTR*)")
    If @error Then
        Return SetError(@error)
    EndIf

    Local $pNotificationData = CreateNotificationData()
    If @error <> 0 Then
        Return SetError(@error)
    EndIf

    Local $oNotificationData = ObjCreateInterface($pNotificationData, "{9FFD2312-9D6A-4AAF-B6AC-FF17F0C1F280}", $sIInspectable & "Values HRESULT(PTR*);SequenceNumber HRESULT(UINT*);SetSequenceNumber HRESULT(UINT);")
    If @error Then
        Return SetError(@error)
    EndIf

    Local $pValues = 0
    Local $hr = $oNotificationData.Values($pValues)

    Local $pInsert = __Toast_VTable_get($pValues, 10)
    Local $pReplaced = 0
    Local $pKey = "progress1"
    $hr = __Toast_WindowsCreateString($pKey, $pKey)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf
    Local $pValue = "0.9"
    $hr = __Toast_WindowsCreateString($pValue, $pValue)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf
    $hr = DllCallAddress("LONG", $pInsert, "PTR", $pValues, "PTR", $pKey, "PTR", $pValue, "BOOLEAN*", 0)
    __Toast_WindowsDeleteString($pKey)
    __Toast_WindowsDeleteString($pValue)
    $oNotificationData.SetSequenceNumber(0) ; Need to be incremented (starting from 1) if the order of updates are important. 0 will apply without caring about the order.

    Local $pTag = "demo-tag"
    __Toast_WindowsCreateString($pTag, $pTag)
    Local $pGroup = "demo-group"
    __Toast_WindowsCreateString($pGroup, $pGroup)
    Local $iNotificationUpdateResult = 0

    $hr = $oToastNotifier2.UpdateWithTagAndGroup($pNotificationData, $pTag, $pGroup, $iNotificationUpdateResult)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf

    ; https://learn.microsoft.com/en-us/uwp/api/windows.ui.notifications.notificationupdateresult?view=winrt-26100
    Switch $iNotificationUpdateResult
        Case 0 ; Succeeded
            _GUICtrlRichEdit_AppendText($hRich, "The notification was updated."&@CRLF)
        Case 1 ; Failed
            _GUICtrlRichEdit_AppendText($hRich, "The notification update failed."&@CRLF)
        Case 2 ; NotificationNotFound
            _GUICtrlRichEdit_AppendText($hRich, "The specified notification couldn't be found."&@CRLF)
        Case Else
            _GUICtrlRichEdit_AppendText($hRich, "Unexpected notification update result: "&$iNotificationUpdateResult&@CRLF)
    EndSwitch

    __Toast_WindowsDeleteString($pTag)
    __Toast_WindowsDeleteString($pGroup)
EndFunc

Func CreateNotificationData()
    Local $classId = "Windows.UI.Notifications.NotificationData"
    Local $hr = __Toast_WindowsCreateString($classId, $classId)
    If @error <> 0 Then Return SetError(@error)

    Local $pInspectable = 0

    Local Static $IID_INotificationDataFactory = "{23c1e33a-1c10-46fb-8040-dec384621cf8}"

    $hr = RoGetActivationFactory($classId, $IID_INotificationDataFactory, $pInspectable)

    __Toast_WindowsDeleteString($classId)

    If $hr <> 0 Then Return SetError($hr)

    Local $pINotificationDataFactory = 0
    $hr = __Toast_QueryInterface($pInspectable, $IID_INotificationDataFactory, $pINotificationDataFactory)

    If @error <> 0 Then
        Local $error = @error, $extended = @extended
        __Toast_IUnknown_Release($pInspectable)
        Return SetError($error, $extended, 0)
    EndIf

    __Toast_IUnknown_Release($pInspectable)

    Local $pCreateNotificationDataWithValuesAndSequenceNumber = __Toast_VTable_get($pINotificationDataFactory, 6)

    Local $aRet = DllCallAddress("LONG", $pCreateNotificationDataWithValuesAndSequenceNumber, "PTR", $pINotificationDataFactory, "PTR", 0, "UINT", 0, "PTR*", 0)
    If $aRet[0] <> 0 Then
        __Toast_IUnknown_Release($pINotificationDataFactory)
        Return SetError($hr, 0, 0)
    EndIf

    __Toast_IUnknown_Release($pINotificationDataFactory)

    Return $aRet[4]
EndFunc

Func __Toast_VTable_get($pInterface, $iMethod)
    Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

    Local $methodOffset = (@AutoItX64 ? 8 : 4) * $iMethod
    
    Return DllStructGetData(DllStructCreate("ptr", $pVTable + $methodOffset), 1)
EndFunc

Func GUI_CLOSE()
    Exit
EndFunc
