#include <WinAPIConv.au3>

Global Const $sIInspectable = "GetIids HRESULT(ULONG*;PTR*);GetRuntimeClassName HRESULT(PTR);GetTrustLevel HRESULT(PTR);"

;Global $__Toast_

Global Enum _
    $_Toast_ToastTemplateType_ToastImageAndText01 = 0, _
    $_Toast_ToastTemplateType_ToastImageAndText02 = 1, _
    $_Toast_ToastTemplateType_ToastImageAndText03 = 2, _
    $_Toast_ToastTemplateType_ToastImageAndText04 = 3, _
    $_Toast_ToastTemplateType_ToastText01 = 4, _
    $_Toast_ToastTemplateType_ToastText02 = 5, _
    $_Toast_ToastTemplateType_ToastText03 = 6, _
    $_Toast_ToastTemplateType_ToastText04 = 7

Func RoGetActivationFactory($activatableClassId, $iid, ByRef $factory)
    Local $aRet = DllCall("Combase.dll", "LONG", "RoGetActivationFactory", "PTR", $activatableClassId, "PTR", $iid, "PTR*", 0)
    $factory = $aRet[3]
    Return $aRet[0]
EndFunc

Func __Toast_WindowsCreateString($sourceString, ByRef $pStr)
    Local $aRet = DllCall("combase.dll", "LONG", "WindowsCreateString", "WSTR", $sourceString, "long", StringLen($sourceString), "PTR*", 0)

    If @error<>0 Then Return SetError(@error, @extended, $aRet)

    If $aRet[0] <> 0 Then Return SetError($aRet[0], 0, $aRet)

    $pStr = $aRet[3]

    Return $aRet[0]
EndFunc

Func __Toast_WindowsDeleteString(ByRef $pStr)
    Local $aRet = DllCall("combase.dll", "LONG", "WindowsDeleteString", "PTR", $pStr)

    If @error<>0 Then Return SetError(@error, @extended, $aRet)

    If $aRet[0] <> 0 Then Return SetError($aRet[0], 0, $aRet)

    $pStr = 0

    Return $aRet[0]
EndFunc

Func __Toast_ToastNotificationManager()
    Local Static $oToastNotificationManager = Null

    If Not ($oToastNotificationManager = Null) Then
        Return $oToastNotificationManager
    EndIf

    Local Static $IID_IToastNotificationManagerStatics = "{50ac103f-d235-4598-bbef-98fe4d1a3ad4}"
    Local $tUUID_IToastNotificationManagerStatics = _WinAPI_GUIDFromString($IID_IToastNotificationManagerStatics)
    Local $hUUID_IToastNotificationManagerStatics = DllStructGetPtr($tUUID_IToastNotificationManagerStatics)
    Local Static $sIToastNotificationManagerStatics = $sIInspectable & "CreateToastNotifier HRESULT(ptr*);CreateToastNotifierWithId HRESULT(ptr;ptr*);GetTemplateContent HRESULT(ptr;ptr*)"

    Local $classId = "Windows.UI.Notifications.ToastNotificationManager"
    Local $hToastNotificationManager = 0
    Local $hr = __Toast_WindowsCreateString($classId, $classId)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf

    $hr = RoGetActivationFactory($classId, $hUUID_IToastNotificationManagerStatics, $hToastNotificationManager)
    if $hr <> 0 Then
        __Toast_WindowsDeleteString($classId)
        Return SetError($hr)
    EndIf

    __Toast_WindowsDeleteString($classId)

    $oToastNotificationManager = ObjCreateInterface($hToastNotificationManager, $IID_IToastNotificationManagerStatics, $sIToastNotificationManagerStatics)

    If @error <> 0 Then
        Local $error = @error, $extended = @extended
        __Toast_IUnknown_Release($hToastNotificationManager)
        Return SetError($error, $extended)
    EndIf

    ;__Toast_IUnknown_Release($hToastNotificationManager)

    Return $oToastNotificationManager
EndFunc

Func __Toast_QueryInterface($pInterface, $riid, ByRef $ppvObject)
    If IsString($riid) Then $riid = _WinAPI_GUIDFromString($riid)

    Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

    Local $pQueryInterface = DllStructGetData(DllStructCreate("ptr", $pVTable), 1)

    Local $aRet = DllCallAddress("long", $pQueryInterface, "ptr", $pInterface, "struct*", $riid, "ptr*", 0)

    If @error <> 0 Then Return SetError(@error, @extended, $aRet)

    $ppvObject = $aRet[3]

    Return $aRet[0]
EndFunc

Func __Toast_IUnknown_AddRef($pInterface)
    Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

    Local $releaseOffset = @AutoItX64 ? 8 : 4
    Local $pRelease = DllStructGetData(DllStructCreate("ptr", $pVTable + $releaseOffset), 1)

    Local $aRet = DllCallAddress("ulong", $pRelease, "ptr", $pInterface)

    If @error <> 0 Then Return SetError(@error, @extended, $aRet)

    Return $aRet[0]
EndFunc

Func __Toast_IUnknown_Release($pInterface)
    Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

    Local $releaseOffset = @AutoItX64 ? 16 : 8
    Local $pRelease = DllStructGetData(DllStructCreate("ptr", $pVTable + $releaseOffset), 1)

    Local $aRet = DllCallAddress("ulong", $pRelease, "ptr", $pInterface)

    If @error <> 0 Then Return SetError(@error, @extended, $aRet)

    Return $aRet[0]
EndFunc

Func _Toast_Show($oToastNotification, $sAppId = @ScriptName)
    Local $oToastNotificationManager = __Toast_ToastNotificationManager()
    If @error <> 0 Then Return SetError(@error, @extended, $oToastNotificationManager)

    Local $hr = __Toast_WindowsCreateString($sAppId, $sAppId)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf

    Local $pToastNotifier = 0
    $hr = $oToastNotificationManager.CreateToastNotifierWithId($sAppId, $pToastNotifier)
    If @error <> 0 Then Return SetError(@error, @extended, $hr)
    If $hr <> 0 Then
        __Toast_WindowsDeleteString($sAppId)
        Return SetError($hr)
    EndIf

    __Toast_WindowsDeleteString($sAppId)

    Local $oToastNotifier = ObjCreateInterface($pToastNotifier, "{75927B93-03F3-41EC-91D3-6E5BAC1B38E7}", $sIInspectable & "Show HRESULT(PTR)")

    If @error <> 0 Then
        __Toast_IUnknown_Release($pToastNotifier)
        Return SetError(@error, @extended, $oToastNotifier)
    EndIf

    $hr = $oToastNotifier.Show($oToastNotification)
    If @error <> 0 Then Return SetError(@error, @extended, $hr)
    If $hr <> 0 Then
        Return SetError($hr)
    EndIf

    __Toast_IUnknown_Release($pToastNotifier)

    Return $hr
EndFunc

Func _Toast_CreateToastNotificationFromXmlObject($oXml)
    Local Static $IID_IToastNotificationFactory = "{04124B20-82C6-4229-B109-FD9ED4662B53}"
    Local $tIID_IToastNotificationFactory = _WinAPI_GUIDFromString($IID_IToastNotificationFactory)
    Local $hIID_IToastNotificationFactory = DllStructGetPtr($tIID_IToastNotificationFactory)
    Local Static $sIToastNotificationFactory = $sIInspectable & "CreateToastNotification HRESULT(ptr;ptr*);"
    Local $classId = "Windows.UI.Notifications.ToastNotification"
    $hr = __Toast_WindowsCreateString($classId, $classId)
    if $hr <> 0 Then
        Return SetError($hr)
    EndIf

    Local $hIToastNotificationFactory = 0
    $hr = RoGetActivationFactory($classId, $hIID_IToastNotificationFactory, $hIToastNotificationFactory)
    if $hr <> 0 Then
        __Toast_WindowsDeleteString($classId)
        Return SetError($hr)
    EndIf

    __Toast_WindowsDeleteString($classId)
    
    Local $oIToastNotificationFactory = ObjCreateInterface($hIToastNotificationFactory, $IID_IToastNotificationFactory, $sIToastNotificationFactory)
    If @error <> 0 Then
        __Toast_IUnknown_Release($hIToastNotificationFactory)
        Return SetError(@error, @extended, $oIToastNotificationFactory)
    EndIf

    Local $pToastNotification = 0
    Local $hr = $oIToastNotificationFactory.CreateToastNotification($oXml, $pToastNotification)
    If $hr <> 0 Then Return SetError($hr)

    __Toast_IUnknown_Release($hIToastNotificationFactory)

    return $pToastNotification
EndFunc

Func _Toast_CreateToastNotificationFromXmlString($sXml)
    Local $pXmlDocument = __Toast_XmlDocument()

    If @error <> 0 Then
        Return SetError(@error)
    EndIf

    Local Static $UIID_IXmlDocumentIO = "{6cd0e74e-ee65-4489-9ebf-ca43e87ba637}"

    Local $pXmlDocumentIO = 0
    __Toast_QueryInterface($pXmlDocument, $UIID_IXmlDocumentIO, $pXmlDocumentIO)

    If @error <> 0 Then
        Return SetError(@error, @extended, $oXmlDocumentIO)
    EndIf


    Local $oXmlDocumentIO = ObjCreateInterface($pXmlDocumentIO, $UIID_IXmlDocumentIO, $sIInspectable & "LoadXml HRESULT(ptr);LoadXmlWithSettings HRESULT(ptr;ptr);SaveToFileAsync HRESULT(ptr;ptr*);")

    If @error <> 0 Then
        Return SetError(@error, @extended, $oXmlDocumentIO)
    EndIf

    __Toast_WindowsCreateString($sXml, $sXml)
    If @error <> 0 Then
        Return SetError(@error, @extended, $oXmlDocumentIO)
    EndIf

    $hr = $oXmlDocumentIO.LoadXml($sXml)
    If $hr <> 0 Then
        __Toast_WindowsDeleteString($sXml)
        Return SetError($hr)
    EndIf

    __Toast_WindowsDeleteString($sXml)

    $hr = _Toast_CreateToastNotificationFromXmlObject($pXmlDocument)

    Return SetError(@error, @extended, $hr)
EndFunc

Func _Toast_CreateToastTemplateXmlDocument($iToastTemplateType = $_Toast_ToastTemplateType_ToastImageAndText01)
    Local $oToastNotificationManager = __Toast_ToastNotificationManager()
    If @error <> 0 Then Return SetError(@error, @extended, $oToastNotificationManager)

    Local $pToastXml
    Local $hr = $oToastNotificationManager.GetTemplateContent($iToastTemplateType, $pToastXml)
    If $hr <> 0 Then Return SetError($hr)

    Local Static $UIID_IXmlDocument = "{f7f3a506-1e87-42d6-bcfb-b8c809fa5494}"
    Local Static $sIXmlDocument = $sIInspectable & "get_Doctype HRESULT(ptr*);get_Implementation HRESULT(ptr*);get_DocumentElement HRESULT(ptr*);CreateElement HRESULT(ptr;ptr*);CreateDocumentFragment HRESULT(ptr*);CreateTextNode HRESULT(ptr;ptr*);CreateComment HRESULT(ptr;ptr*);CreateProcessingInstruction HRESULT(ptr;ptr;ptr*);CreateAttribute HRESULT(ptr;ptr*);CreateEntityReference HRESULT(ptr;ptr*);GetElementsByTagName HRESULT(ptr;ptr*);CreateCDataSection HRESULT(ptr;ptr*);get_DocumentUri HRESULT(ptr);CreateAttributeNS HRESULT(ptr;ptr;ptr*);CreateElementNS HRESULT(ptr;ptr;ptr*);GetElementById HRESULT(ptr;ptr*);ImportNode HRESULT(ptr;boolean;ptr*);"

    $oToastXml = ObjCreateInterface($pToastXml, $UIID_IXmlDocument, $sIXmlDocument)
    If @error <> 0 Then
        __Toast_IUnknown_Release($pToastXml)
        Return SetError(@error, @extended)
    EndIf

    ;__Toast_IUnknown_Release($pToastXml)

    Return $oToastXml
EndFunc

Func __Toast_XmlDocument()
    Local $classId = "Windows.Data.Xml.Dom.XmlDocument"
    __Toast_WindowsCreateString($classId, $classId)

    Local $pInspectable = 0
    $hr = RoActivateInstance($classId, $pInspectable)

    If $hr <> 0 Then Return SetError($hr)

    Local Static $UIID_IXmlDocument = "{f7f3a506-1e87-42d6-bcfb-b8c809fa5494}"
    ;Local Static $sIXmlDocument = $sIInspectable & "get_Doctype HRESULT(ptr*);get_Implementation HRESULT(ptr*);get_DocumentElement HRESULT(ptr*);CreateElement HRESULT(ptr;ptr*);CreateDocumentFragment HRESULT(ptr*);CreateTextNode HRESULT(ptr;ptr*);CreateComment HRESULT(ptr;ptr*);CreateProcessingInstruction HRESULT(ptr;ptr;ptr*);CreateAttribute HRESULT(ptr;ptr*);CreateEntityReference HRESULT(ptr;ptr*);GetElementsByTagName HRESULT(ptr;ptr*);CreateCDataSection HRESULT(ptr;ptr*);get_DocumentUri HRESULT(ptr);CreateAttributeNS HRESULT(ptr;ptr;ptr*);CreateElementNS HRESULT(ptr;ptr;ptr*);GetElementById HRESULT(ptr;ptr*);ImportNode HRESULT(ptr;boolean;ptr*);"
    ;$oXmlDocument = ObjCreateInterface($pXmlDocument, $UIID_IXmlDocument, $sIXmlDocument)
    Local $pXmlDocument = 0
    $hr = __Toast_QueryInterface($pInspectable, $UIID_IXmlDocument, $pXmlDocument)

    If @error <> 0 Then
        Local $error = @error, $extended = @extended
        __Toast_IUnknown_Release($pXmlDocument)
        Return SetError($error, $extended, $oXmlDocument)
    EndIf

    ;__Toast_IUnknown_Release($pInspectable)

    Return $pXmlDocument
EndFunc

Func RoActivateInstance($activatableClassId, ByRef $instance)
    Local $aRet = DllCall("combase.dll", "LONG", "RoActivateInstance", "PTR", $activatableClassId, "PTR*", 0)
    $instance = $aRet[2]
    Return $aRet[0]
EndFunc

Func __Toast_ToastNotificationFactory()

    Local Static $IID_IToastNotificationFactory = "{04124B20-82C6-4229-B109-FD9ED4662B53}"
    Local $tIID_IToastNotificationFactory = _WinAPI_GUIDFromString($IID_IToastNotificationFactory)
    Local $hIID_IToastNotificationFactory = DllStructGetPtr($tIID_IToastNotificationFactory)
    Local Static $sIToastNotificationFactory = $sIInspectable & "CreateToastNotification HRESULT(ptr;ptr*);"
    Local $classId = "Windows.UI.Notifications.ToastNotification"
    $hr = __Toast_WindowsCreateString($classId, $classId)
    if $hr <> 0 Then
        Return SetError($hr)
    EndIf

    Local $hIToastNotificationFactory = 0
    $hr = RoGetActivationFactory($classId, $hIID_IToastNotificationFactory, $hIToastNotificationFactory)
    if $hr <> 0 Then
        __Toast_WindowsDeleteString($classId)
        Return SetError($hr)
    EndIf

    __Toast_WindowsDeleteString($classId)
    
    Local $oIToastNotificationFactory = ObjCreateInterface($hIToastNotificationFactory, $IID_IToastNotificationFactory, $sIToastNotificationFactory)
    If @error <> 0 Then
        __Toast_IUnknown_Release($hIToastNotificationFactory)
        Return SetError(@error, @extended, $oIToastNotificationFactory)
    EndIf

    Return $oIToastNotificationFactory
EndFunc
