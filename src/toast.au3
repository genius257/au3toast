#include <WinAPIConv.au3>
#include <Memory.au3>
#include <WinAPICom.au3>
#include <WinAPIShellEx.au3>
#include <WinAPIShPath.au3>
#include <WinApiReg.au3>
#include <WinAPIConv.au3>

Global Const $sIInspectable = "GetIids HRESULT(ULONG;PTR*);GetRuntimeClassName HRESULT(PTR);GetTrustLevel HRESULT(PTR);"

Global Const $_Toast_S_OK = 0x0
Global Const $_Toast_E_FAIL = 0x80004005
Global Const $_Toast_E_NOINTERFACE = 0x80004002
Global Const $_Toast_E_POINTER = 0x80004003

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
EndFunc   ;==>RoGetActivationFactory

Func __Toast_GlobalHandle($pMem)
	Local $aRet = DllCall("Kernel32.dll", "ptr", "GlobalHandle", "ptr", $pMem)
	If @error <> 0 Then Return SetError(@error, @extended, 0)
	If $aRet[0] = 0 Then Return SetError(-1, @extended, 0)
	Return $aRet[0]
EndFunc   ;==>__Toast_GlobalHandle

Func __Toast_WindowsCreateString($sourceString, ByRef $pStr)
	Local $aRet = DllCall("combase.dll", "LONG", "WindowsCreateString", "WSTR", $sourceString, "long", StringLen($sourceString), "PTR*", 0)

	If @error <> 0 Then Return SetError(@error, @extended, $aRet)

	If $aRet[0] <> 0 Then Return SetError($aRet[0], 0, $aRet)

	$pStr = $aRet[3]

	Return $aRet[0]
EndFunc   ;==>__Toast_WindowsCreateString

Func __Toast_WindowsDeleteString(ByRef $pStr)
	Local $aRet = DllCall("combase.dll", "LONG", "WindowsDeleteString", "PTR", $pStr)

	If @error <> 0 Then Return SetError(@error, @extended, $aRet)

	If $aRet[0] <> 0 Then Return SetError($aRet[0], 0, $aRet)

	$pStr = 0

	Return $aRet[0]
EndFunc   ;==>__Toast_WindowsDeleteString

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
	If $hr <> 0 Then
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
EndFunc   ;==>__Toast_ToastNotificationManager

Func __Toast_QueryInterface($pInterface, $riid, ByRef $ppvObject)
	If IsString($riid) Then $riid = _WinAPI_GUIDFromString($riid)

	Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

	Local $pQueryInterface = DllStructGetData(DllStructCreate("ptr", $pVTable), 1)

	Local $aRet = DllCallAddress("long", $pQueryInterface, "ptr", $pInterface, "struct*", $riid, "ptr*", 0)

	If @error <> 0 Then Return SetError(@error, @extended, $aRet)

	$ppvObject = $aRet[3]

	Return $aRet[0]
EndFunc   ;==>__Toast_QueryInterface

Func __Toast_IUnknown_AddRef($pInterface)
	Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

	Local $releaseOffset = @AutoItX64 ? 8 : 4
	Local $pRelease = DllStructGetData(DllStructCreate("ptr", $pVTable + $releaseOffset), 1)

	Local $aRet = DllCallAddress("ulong", $pRelease, "ptr", $pInterface)

	If @error <> 0 Then Return SetError(@error, @extended, $aRet)

	Return $aRet[0]
EndFunc   ;==>__Toast_IUnknown_AddRef

Func __Toast_IUnknown_Release($pInterface)
	Local $pVTable = DllStructGetData(DllStructCreate("ptr", $pInterface), 1)

	Local $releaseOffset = @AutoItX64 ? 16 : 8
	Local $pRelease = DllStructGetData(DllStructCreate("ptr", $pVTable + $releaseOffset), 1)

	Local $aRet = DllCallAddress("ulong", $pRelease, "ptr", $pInterface)

	If @error <> 0 Then Return SetError(@error, @extended, $aRet)

	Return $aRet[0]
EndFunc   ;==>__Toast_IUnknown_Release

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
EndFunc   ;==>_Toast_Show

Func _Toast_hide($oToastNotification, $sAppId = @ScriptName)
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

	Local $oToastNotifier = ObjCreateInterface($pToastNotifier, "{75927B93-03F3-41EC-91D3-6E5BAC1B38E7}", $sIInspectable & "Show HRESULT(PTR);Hide HRESULT(PTR);")

	If @error <> 0 Then
		__Toast_IUnknown_Release($pToastNotifier)
		Return SetError(@error, @extended, $oToastNotifier)
	EndIf

	$hr = $oToastNotifier.Hide($oToastNotification)
	If @error <> 0 Then Return SetError(@error, @extended, $hr)
	If $hr <> 0 Then
		Return SetError($hr)
	EndIf

	__Toast_IUnknown_Release($pToastNotifier)

	Return $hr
EndFunc   ;==>_Toast_hide

Func _Toast_CreateToastNotificationFromXmlObject($oXml)
	Local Static $IID_IToastNotificationFactory = "{04124B20-82C6-4229-B109-FD9ED4662B53}"
	Local $tIID_IToastNotificationFactory = _WinAPI_GUIDFromString($IID_IToastNotificationFactory)
	Local $hIID_IToastNotificationFactory = DllStructGetPtr($tIID_IToastNotificationFactory)
	Local Static $sIToastNotificationFactory = $sIInspectable & "CreateToastNotification HRESULT(ptr;ptr*);"
	Local $classId = "Windows.UI.Notifications.ToastNotification"
	$hr = __Toast_WindowsCreateString($classId, $classId)
	If $hr <> 0 Then
		Return SetError($hr)
	EndIf

	Local $hIToastNotificationFactory = 0
	$hr = RoGetActivationFactory($classId, $hIID_IToastNotificationFactory, $hIToastNotificationFactory)
	If $hr <> 0 Then
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

	Return $pToastNotification
EndFunc   ;==>_Toast_CreateToastNotificationFromXmlObject

Func _Toast_CreateToastNotificationFromXmlString($sXml)
	Local $pXmlDocument = __Toast_XmlDocument()

	If @error <> 0 Then
		Return SetError(@error)
	EndIf

	Local Static $UIID_IXmlDocumentIO = "{6cd0e74e-ee65-4489-9ebf-ca43e87ba637}"

	Local $pXmlDocumentIO = 0
	__Toast_QueryInterface($pXmlDocument, $UIID_IXmlDocumentIO, $pXmlDocumentIO)

	If @error <> 0 Then
		Return SetError(@error, @extended, 0)
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
EndFunc   ;==>_Toast_CreateToastNotificationFromXmlString

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
EndFunc   ;==>_Toast_CreateToastTemplateXmlDocument

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
		Return SetError($error, $extended, 0)
	EndIf

	;__Toast_IUnknown_Release($pInspectable)

	Return $pXmlDocument
EndFunc   ;==>__Toast_XmlDocument

Func RoActivateInstance($activatableClassId, ByRef $instance)
	Local $aRet = DllCall("combase.dll", "LONG", "RoActivateInstance", "PTR", $activatableClassId, "PTR*", 0)
	$instance = $aRet[2]
	Return $aRet[0]
EndFunc   ;==>RoActivateInstance

Func __Toast_ToastNotificationFactory()

	Local Static $IID_IToastNotificationFactory = "{04124B20-82C6-4229-B109-FD9ED4662B53}"
	Local $tIID_IToastNotificationFactory = _WinAPI_GUIDFromString($IID_IToastNotificationFactory)
	Local $hIID_IToastNotificationFactory = DllStructGetPtr($tIID_IToastNotificationFactory)
	Local Static $sIToastNotificationFactory = $sIInspectable & "CreateToastNotification HRESULT(ptr;ptr*);"
	Local $classId = "Windows.UI.Notifications.ToastNotification"
	$hr = __Toast_WindowsCreateString($classId, $classId)
	If $hr <> 0 Then
		Return SetError($hr)
	EndIf

	Local $hIToastNotificationFactory = 0
	$hr = RoGetActivationFactory($classId, $hIID_IToastNotificationFactory, $hIToastNotificationFactory)
	If $hr <> 0 Then
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
EndFunc   ;==>__Toast_ToastNotificationFactory

Func __Toast_ToastNotification($pToastNotification)
	Local Static $IID_IToastNotification = "{997E2675-059E-4E60-8B06-1760917C8B80}"
	Local Static $sToastNotification = $sIInspectable & "get_Content HRESULT(ptr*);put_ExpirationTime HRESULT(ptr);get_ExpirationTime HRESULT(ptr*);add_Dismissed HRESULT(ptr;ptr*);remove_Dismissed HRESULT(ptr);add_Activated HRESULT(ptr;ptr*);remove_Activated HRESULT(ptr);add_Failed HRESULT(ptr;ptr*);remove_Failed HRESULT(ptr);"

	Return ObjCreateInterface($pToastNotification, $IID_IToastNotification, $sToastNotification)
EndFunc   ;==>__Toast_ToastNotification

Func __Toast_ITypedEventHandler_Activated($fCallback)
	Local Static $hQueryInterface = DllCallbackRegister(__Toast_ITypedEventHandler_Activated_QueryInterface, "LONG", "ptr;ptr;ptr")
	Local Static $pQueryInterface = DllCallbackGetPtr($hQueryInterface)

	Return __Toast_ITypedEventHandler($fCallback, $pQueryInterface)
EndFunc   ;==>__Toast_ITypedEventHandler_Activated

Func __Toast_ITypedEventHandler_Dismissed($fCallback)
	Local Static $hQueryInterface = DllCallbackRegister(__Toast_ITypedEventHandler_Dismissed_QueryInterface, "LONG", "ptr;ptr;ptr")
	Local Static $pQueryInterface = DllCallbackGetPtr($hQueryInterface)

	Return __Toast_ITypedEventHandler($fCallback, $pQueryInterface)
EndFunc   ;==>__Toast_ITypedEventHandler_Dismissed

Func __Toast_ITypedEventHandler_Failed($fCallback)
	Local Static $hQueryInterface = DllCallbackRegister(__Toast_ITypedEventHandler_Failed_QueryInterface, "LONG", "ptr;ptr;ptr")
	Local Static $pQueryInterface = DllCallbackGetPtr($hQueryInterface)

	Return __Toast_ITypedEventHandler($fCallback, $pQueryInterface)
EndFunc   ;==>__Toast_ITypedEventHandler_Failed

Func __Toast_ITypedEventHandler($fCallback, $pQueryInterface)
	Local $hInvoke = DllCallbackRegister($fCallback, "dword", "ptr;ptr;ptr")
	Local $pInvoke = DllCallbackGetPtr($hInvoke)
	Local Static $hAddRef = DllCallbackRegister(__Toast_ITypedEventHandler_AddRef, "dword", "PTR")
	Local Static $pAddRef = DllCallbackGetPtr($hAddRef)
	Local Static $hRelease = DllCallbackRegister(__Toast_ITypedEventHandler_Release, "dword", "PTR")
	Local Static $pRelease = DllCallbackGetPtr($hRelease)

	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[4];handle dllcallback;")
	DllStructSetData($tObject, "refcount", 1)
	DllStructSetData($tObject, "vtable", $pQueryInterface, 1)
	DllStructSetData($tObject, "vtable", $pAddRef, 2)
	DllStructSetData($tObject, "vtable", $pRelease, 3)
	DllStructSetData($tObject, "vtable", $pInvoke, 4)
	DllStructSetData($tObject, "dllcallback", $hInvoke)

	Local $pObject = DllStructGetPtr($tObject)
	Local $iSize = DllStructGetSize($tObject)
	Local $hMemory = _MemGlobalAlloc($iSize, $GMEM_MOVEABLE)
	Local $pMemory = _MemGlobalLock($hMemory)
	_MemMoveMemory($pObject, $pMemory, $iSize)
	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[4];handle dllcallback;", $pMemory)
	DllStructSetData($tObject, "object", DllStructGetPtr($tObject, "vtable"))

	Return DllStructGetPtr($tObject, "object")
EndFunc   ;==>__Toast_ITypedEventHandler

Func __Toast_ITypedEventHandler_Activated_QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj = 0 Then Return $_Toast_E_POINTER

	Local $sGUID = _WinAPI_StringFromGUID($pRIID)

	Switch $sGUID
		Case '{00000000-0000-0000-C000-000000000046}' _ ; IID_IUnknown
				, '{AB54DE2D-97D9-5528-B6AD-105AFE156530}' ; ITypedEventHandler<ABI::Windows::UI::Notifications::ToastNotification*,IInspectable*>
			Local $tStruct = DllStructCreate("ptr", $pObj)
			DllStructSetData($tStruct, 1, $pSelf)
			__Toast_ITypedEventHandler_AddRef($pSelf)
			Return $_Toast_S_OK
		Case Else
			Return $_Toast_E_NOINTERFACE
	EndSwitch
EndFunc   ;==>__Toast_ITypedEventHandler_Activated_QueryInterface

Func __Toast_ITypedEventHandler_Dismissed_QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj = 0 Then Return $_Toast_E_POINTER

	Local $sGUID = _WinAPI_StringFromGUID($pRIID)

	Switch $sGUID
		Case '{00000000-0000-0000-C000-000000000046}' _ ; IID_IUnknown
				, '{61C2402F-0ED0-5A18-AB69-59F4AA99A368}' ; ITypedEventHandler<ABI::Windows::UI::Notifications::ToastNotification*,ABI::Windows::UI::Notifications::ToastDismissedEventArgs*>
			Local $tStruct = DllStructCreate("ptr", $pObj)
			DllStructSetData($tStruct, 1, $pSelf)
			__Toast_ITypedEventHandler_AddRef($pSelf)
			Return $_Toast_S_OK
		Case Else
			Return $_Toast_E_NOINTERFACE
	EndSwitch
EndFunc   ;==>__Toast_ITypedEventHandler_Dismissed_QueryInterface

Func __Toast_ITypedEventHandler_Failed_QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj = 0 Then Return $_Toast_E_POINTER

	Local $sGUID = _WinAPI_StringFromGUID($pRIID)

	Switch $sGUID
		Case '{00000000-0000-0000-C000-000000000046}' _ ; IID_IUnknown
				, '{95E3E803-C969-5E3A-9753-EA2AD22A9A33}' ; ITypedEventHandler<ABI::Windows::UI::Notifications::ToastNotification*,ABI::Windows::UI::Notifications::ToastFailedEventArgs*>
			Local $tStruct = DllStructCreate("ptr", $pObj)
			DllStructSetData($tStruct, 1, $pSelf)
			__Toast_ITypedEventHandler_AddRef($pSelf)
			Return $_Toast_S_OK
		Case Else
			Return $_Toast_E_NOINTERFACE
	EndSwitch
EndFunc   ;==>__Toast_ITypedEventHandler_Failed_QueryInterface

Func __Toast_ITypedEventHandler_AddRef($pSelf)
	Local $tStruct = DllStructCreate("int Ref", $pSelf - 4)
	$tStruct.Ref += 1
	Return $tStruct.Ref
EndFunc   ;==>__Toast_ITypedEventHandler_AddRef

Func __Toast_ITypedEventHandler_Release($pSelf)
	Local $tStruct = DllStructCreate("int Ref", $pSelf - 4)
	$tStruct.Ref -= 1

	If $tStruct.Ref > 0 Then Return $tStruct.Ref

	Local $pMemory = $pSelf - 4
	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[4];handle dllcallback;", $pMemory)
	DllCallbackFree($tObject.dllcallback)

	Local $hMemory = __Toast_GlobalHandle($pMemory)
	_MemGlobalFree($hMemory)

	Return 0
EndFunc   ;==>__Toast_ITypedEventHandler_Release

Func __Toast_ITypedEventHandler_Invoke($pSelf, $pSender, $pArgs)
	;
EndFunc   ;==>__Toast_ITypedEventHandler_Invoke

Global Const $__Toast_tNOTIFICATION_USER_INPUT_DATA = "STRUCT;PTR Key;PTR Value;ENDSTRUCT;"

Func __Toast_INotificationActivationCallback($hCallback = 0)
	Local Static $hQueryInterface = DllCallbackRegister(__Toast_INotificationActivationCallback_QueryInterface, "LONG", "ptr;ptr;ptr")
	Local Static $pQueryInterface = DllCallbackGetPtr($hQueryInterface)
	Local Static $hAddRef = DllCallbackRegister(__Toast_ITypedEventHandler_AddRef, "dword", "PTR")
	Local Static $pAddRef = DllCallbackGetPtr($hAddRef)
	Local Static $hRelease = DllCallbackRegister(__Toast_ITypedEventHandler_Release, "dword", "PTR")
	Local Static $pRelease = DllCallbackGetPtr($hRelease)
	Local Static $hActivate = DllCallbackRegister(__Toast_INotificationActivationCallback_Activate, "LONG", "ptr;WSTR;WSTR;ptr;ulong")
	Local Static $pActivate = DllCallbackGetPtr($hActivate)

	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[4];handle dllcallback;")
	DllStructSetData($tObject, "refcount", 1)
	DllStructSetData($tObject, "vtable", $pQueryInterface, 1)
	DllStructSetData($tObject, "vtable", $pAddRef, 2)
	DllStructSetData($tObject, "vtable", $pRelease, 3)
	Local $pCallback = $hCallback = 0 ? $pActivate : DllCallbackGetPtr($hCallback)
	DllStructSetData($tObject, "vtable", $pCallback, 4)
	;DllStructSetData($tObject, "dllcallback", $hCallback)

	Local $pObject = DllStructGetPtr($tObject)
	Local $iSize = DllStructGetSize($tObject)
	Local $hMemory = _MemGlobalAlloc($iSize, $GMEM_MOVEABLE)
	Local $pMemory = _MemGlobalLock($hMemory)
	_MemMoveMemory($pObject, $pMemory, $iSize)
	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[4];handle dllcallback;", $pMemory)
	DllStructSetData($tObject, "object", DllStructGetPtr($tObject, "vtable"))

	Return DllStructGetPtr($tObject, "object")
EndFunc   ;==>__Toast_INotificationActivationCallback

Func __Toast_INotificationActivationCallback_QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj = 0 Then Return $_Toast_E_POINTER

	Local $sGUID = _WinAPI_StringFromGUID($pRIID)

	Switch $sGUID
		Case '{00000000-0000-0000-C000-000000000046}' _ ; IID_IUnknown
				, '{53E31837-6600-4A81-9395-75CFFE746F94}' ; INotificationActivationCallback
			Local $tStruct = DllStructCreate("ptr", $pObj)
			DllStructSetData($tStruct, 1, $pSelf)
			__Toast_ITypedEventHandler_AddRef($pSelf)
			Return $_Toast_S_OK
		Case Else
			Return $_Toast_E_NOINTERFACE
	EndSwitch
EndFunc   ;==>__Toast_INotificationActivationCallback_QueryInterface

Func __Toast_INotificationActivationCallback_Activate($pSelf, $appUserModelId, $invokedArgs, $data, $count)
	ConsoleWrite("appUserModelId: " & $appUserModelId & @CRLF)
	ConsoleWrite("invokedArgs: " & $invokedArgs & @CRLF)
	ConsoleWrite("count: " & $count & @CRLF)

	Return $_Toast_S_OK
EndFunc   ;==>__Toast_INotificationActivationCallback_Activate

Func __Toast_CoRegisterClassObject($sAppId = @ScriptName, $tCLSID = _Toast_CoCreateGuid(), $fCallback = Null)
	_WinAPI_SetCurrentProcessExplicitAppUserModelID($sAppId)

	$pCLSID = DllStructGetPtr($tCLSID)

	Local Static $CLSCTX_LOCAL_SERVER = 0x4
	Local Static $REGCLS_MULTIPLEUSE = 1

	Local $pIClassFactor = __Toast_IClassFactory($fCallback)
	;Local $pINotificationActivationCallback = __Toast_INotificationActivationCallback()
	Local $aResult = DllCall("ole32.dll", "long", "CoRegisterClassObject", "ptr", $pCLSID, "ptr", $pIClassFactor, "uint", $CLSCTX_LOCAL_SERVER, "uint", $REGCLS_MULTIPLEUSE, "dword*", 0)
	If @error <> 0 Then Return SetError(@error, @extended, 0)

	If $aResult[0] <> 0 Then Return SetError($aResult[0], 0, 0)

	Return $aResult[5]
EndFunc   ;==>__Toast_CoRegisterClassObject

Func __Toast_CoRevokeClassObject($dwRegister)
	Local $aResult = DllCall("Ole32.dll", "LONG", "CoRevokeClassObject", "DWORD", $dwRegister)
	If @error <> 0 Then Return SetError(@error, @extended, 0)

	Return $aResult[0]
EndFunc   ;==>__Toast_CoRevokeClassObject

Func _Toast_CoCreateGuid()
	Local $tGUID = DllStructCreate($__tagWinAPICom_GUID)
	Local $aCall = DllCall('ole32.dll', 'long', 'CoCreateGuid', 'struct*', $tGUID)
	If @error Then Return SetError(@error, @extended, '')
	If $aCall[0] Then Return SetError(10, $aCall[0], '')

	Return $aCall[1]
EndFunc   ;==>_Toast_CoCreateGuid

Func __Toast_IClassFactory($fCallback = Null)
	Local Static $hQueryInterface = DllCallbackRegister(__Toast_IClassFactory_QueryInterface, "LONG", "ptr;ptr;ptr")
	Local Static $pQueryInterface = DllCallbackGetPtr($hQueryInterface)
	Local Static $hAddRef = DllCallbackRegister(__Toast_ITypedEventHandler_AddRef, "dword", "PTR")
	Local Static $pAddRef = DllCallbackGetPtr($hAddRef)
	Local Static $hRelease = DllCallbackRegister(__Toast_IClassFactory_Release, "dword", "PTR")
	Local Static $pRelease = DllCallbackGetPtr($hRelease)
	Local Static $hCreateInstance = DllCallbackRegister(__Toast_IClassFactory_CreateInstance, "LONG", "PTR;PTR;PTR;PTR")
	Local Static $pCreateInstance = DllCallbackGetPtr($hCreateInstance)
	Local Static $hLockServer = DllCallbackRegister(__Toast_IClassFactory_LockServer, "LONG", "BOOLEAN")
	Local Static $pLockServer = DllCallbackGetPtr($hLockServer)

	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[5];ptr callback;")
	DllStructSetData($tObject, "refcount", 1)
	DllStructSetData($tObject, "vtable", $pQueryInterface, 1)
	DllStructSetData($tObject, "vtable", $pAddRef, 2)
	DllStructSetData($tObject, "vtable", $pRelease, 3)
	DllStructSetData($tObject, "vtable", $pCreateInstance, 4)
	DllStructSetData($tObject, "vtable", $pLockServer, 5)
	If IsFunc($fCallback) Then
		Local $hCallback = DllCallbackRegister($fCallback, "LONG", "ptr;WSTR;WSTR;ptr;ulong")
		DllStructSetData($tObject, "callback", $hCallback)
	EndIf

	Local $pObject = DllStructGetPtr($tObject)
	Local $iSize = DllStructGetSize($tObject)
	Local $hMemory = _MemGlobalAlloc($iSize, $GMEM_MOVEABLE)
	Local $pMemory = _MemGlobalLock($hMemory)
	_MemMoveMemory($pObject, $pMemory, $iSize)
	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[5];ptr callback;", $pMemory)
	DllStructSetData($tObject, "object", DllStructGetPtr($tObject, "vtable"))

	Return DllStructGetPtr($tObject, "object")
EndFunc   ;==>__Toast_IClassFactory

Func __Toast_IClassFactory_Release($pSelf)
	Local $tStruct = DllStructCreate("int Ref", $pSelf - 4)
	$tStruct.Ref -= 1

	If $tStruct.Ref > 0 Then Return $tStruct.Ref

	Local $pMemory = $pSelf - 4
	Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[5];ptr callback;", $pMemory)

	Local $hMemory = __Toast_GlobalHandle($pMemory)
	_MemGlobalFree($hMemory)

	Return 0
EndFunc   ;==>__Toast_IClassFactory_Release

Func __Toast_IClassFactory_QueryInterface($pSelf, $pRIID, $pObj)
	If $pObj = 0 Then Return $_Toast_E_POINTER

	Local $sGUID = _WinAPI_StringFromGUID($pRIID)

	Switch $sGUID
		Case '{00000000-0000-0000-C000-000000000046}', _ ; IID_IUnknown
				'{00000001-0000-0000-C000-000000000046}' ; IClassFactory
			Local $tStruct = DllStructCreate("ptr", $pObj)
			DllStructSetData($tStruct, 1, $pSelf)
			__Toast_ITypedEventHandler_AddRef($pSelf)
			Return $_Toast_S_OK
		Case Else
			Return $_Toast_E_NOINTERFACE
	EndSwitch
EndFunc   ;==>__Toast_IClassFactory_QueryInterface

Func __Toast_IClassFactory_CreateInstance($pSelf, $pUnkOuter, $pRIID, $ppvObject)
	Local $sGUID = _WinAPI_StringFromGUID($pRIID)

	Switch $sGUID
		Case '{00000000-0000-0000-C000-000000000046}'
			Local $pMemory = $pSelf - 4
			Local $tObject = DllStructCreate("int refcount;ptr object;ptr vtable[5];ptr callback;", $pMemory)
			Local $pINotificationActivationCallback = __Toast_INotificationActivationCallback(DllStructGetData($tObject, 'callback'))

			DllStructSetData(DllStructCreate("PTR", $ppvObject), 1, $pINotificationActivationCallback)
			Return $_Toast_S_OK
	EndSwitch

	Return $_Toast_E_NOINTERFACE
EndFunc   ;==>__Toast_IClassFactory_CreateInstance

Func __Toast_IClassFactory_LockServer($pSelf, $fLock)
	Return $_Toast_E_FAIL
EndFunc   ;==>__Toast_IClassFactory_LockServer

Global $__Toast_Activator = 0

Func _Toast_Initialize( _
		$sAppName = @ScriptName, _
		$tCLSID = _Toast_CoCreateGuid(), _
		$fCallback = Null, _
		$sDisplayName = $sAppName, _
		$sIconUri = Null _
		)
	$__Toast_Activator = __Toast_CoRegisterClassObject($sAppName, $tCLSID, $fCallback)

	Local $hKey = _WinAPI_RegCreateKey($HKEY_CURRENT_USER, "Software\Classes\AppUserModelId\" & $sAppName, $KEY_ALL_ACCESS, $REG_OPTION_VOLATILE)

	Local $tValue

	$tValue = DllStructCreate("WCHAR[" & (StringLen($sDisplayName) + 2) & "]")
	DllStructSetData($tValue, 1, $sDisplayName)
	_WinAPI_RegSetValue($hKey, "DisplayName", $REG_SZ, $tValue, DllStructGetSize($tValue))

	If Not ($sIconUri = Null) Then
		$tValue = DllStructCreate("WCHAR[" & (StringLen($sIconUri) + 2) & "]")
		DllStructSetData($tValue, 1, $sIconUri)
		_WinAPI_RegSetValue($hKey, "IconUri", $REG_SZ, $tValue, DllStructGetSize($tValue))
	EndIf

	Local $sGUID = _WinAPI_StringFromGUID($tCLSID)
	$tValue = DllStructCreate("WCHAR[" & (StringLen($sGUID) + 2) & "]")
	DllStructSetData($tValue, 1, $sGUID)
	_WinAPI_RegSetValue($hKey, "CustomActivator", $REG_SZ, $tValue, DllStructGetSize($tValue))

	_WinAPI_RegCloseKey($hKey)

	OnAutoItExitRegister("__Toast_Terminate")
EndFunc   ;==>_Toast_Initialize

Func __Toast_Terminate()
	__Toast_CoRevokeClassObject($__Toast_Activator)
EndFunc   ;==>__Toast_Terminate
