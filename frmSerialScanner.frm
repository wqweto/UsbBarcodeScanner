VERSION 5.00
Begin VB.Form frmSerialScanner 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3624
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   3624
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmSerialScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' UsbBarcodeScanner (c) 2017 by wqweto@gmail.com
'
' A VB6 sample project for intercepting USB/HID devices input
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "frmSerialScanner"

'=========================================================================
' API
'=========================================================================

'--- for HID
Private Const RIDI_DEVICENAME               As Long = &H20000007
Private Const RID_INPUT                     As Long = &H10000003
Private Const RIM_TYPEKEYBOARD              As Long = 1
Private Const RIDEV_INPUTSINK               As Long = &H100
Private Const WM_ACTIVATEAPP                As Long = &H1C
Private Const WM_INPUT                      As Long = &HFF
Private Const WM_KEYDOWN                    As Long = &H100
'--- for setupapi
Private Const DIGCF_PRESENT                 As Long = &H2
Private Const DIGCF_ALLCLASSES              As Long = &H4
Private Const DIGCF_PROFILE                 As Long = &H8
Private Const DEVPROP_TYPE_STRING           As Long = &H12
Private Const INVALID_HANDLE_VALUE          As Long = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function GetRawInputDeviceInfo Lib "user32" Alias "GetRawInputDeviceInfoW" (ByVal hDevice As Long, ByVal uiCommand As Long, ByVal pData As Long, pcbSize As Long) As Long
Private Declare Function RegisterRawInputDevices Lib "user32" (pRawInputDevices As RAWINPUTDEVICE, ByVal uiNumDevices As Long, ByVal cbSize As Long) As Long
Private Declare Function GetRawInputData Lib "user32" (ByVal hRawInput As Long, ByVal uiCommand As Long, pData As Any, pcbSize As Long, ByVal cbSizeHeader As Long) As Long
Private Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, lpKeyState As Byte, ByVal pwszBuff As Long, ByVal cchBuff As Long, ByVal wFlags As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Private Declare Function GetRawInputDeviceList Lib "user32" (pRawInputDeviceList As Any, puiNumDevices As Long, ByVal cbSize As Long) As Long
'Private Declare Function GetRawInputDeviceInfo Lib "user32" Alias "GetRawInputDeviceInfoW" (ByVal hDevice As Long, ByVal uiCommand As Long, ByVal pData As Long, pcbSize As Long) As Long
Private Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByRef Class As Any, ByVal Enumerator As String, ByVal Parent As Long, ByVal Flag As Long) As Long
Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal List As Long) As Boolean
Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi.dll" (ByVal List As Long, ByVal Index As Long, ByRef Device As SP_DEVINFO) As Boolean
Private Declare Function SetupDiGetDeviceProperty Lib "setupapi.dll" Alias "SetupDiGetDevicePropertyW" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO, PropertyKey As SP_DEVPROPKEY, PropertyType As Long, ByVal PropertyBuffer As Long, ByVal PropertyBufferSize As Long, RequiredSize As Long, ByVal Flags As Long) As Long
Private Declare Function SetupDiGetDeviceRegistryProperty Lib "setupapi.dll" Alias "SetupDiGetDeviceRegistryPropertyA" (ByVal DeviceInfoSet As Long, DeviceInfoData As SP_DEVINFO, ByVal Property As DEVICEPROPERTYINDEX, PropertyRegDataType As REGPROPERTYTYPES, PropertyBuffer As Any, ByVal PropertyBufferSize As Long, RequiredSize As Long) As Long

Private Type RAWINPUTHEADER
    dwType          As Long
    dwSize          As Long
    hDevice         As Long
    wParam          As Long
End Type

Private Type RAWINPUTKEYBOARD
    hdr             As RAWINPUTHEADER
    MakeCode        As Integer
    Flags           As Integer
    Reserved        As Integer
    vKey            As Integer
    lMessage        As Long
    ExtraInformation As Long
End Type

Private Type RAWINPUTDEVICE
    usUsagePage     As Integer
    usUsage         As Integer
    dwFlags         As Long
    hWndTarget      As Long
End Type

Private Type POINTAPI
    X               As Long
    Y               As Long
End Type

Private Type MSG
    hWnd            As Long
    message         As Long
    wParam          As Long
    lParam          As Long
    time            As Long
    pt              As POINTAPI
End Type

Private Type RAWINPUTDEVICELIST
    hDevice             As Long
    dwType              As Long
End Type

Private Type SP_DEVINFO
    cbSize              As Long
    ClassGuid(0 To 3)   As Long
    DevInstance         As Long
    Reserved            As Long
End Type

Private Type SP_DEVPROPKEY
    fmtid(0 To 3)       As Long
    pid                 As Long
End Type

Public Enum DEVICEPROPERTYINDEX
    SPDRP_DEVICEDESC = &H0                       ' DeviceDesc (R/W)
    SPDRP_HARDWAREID = &H1                       ' HardwareID (R/W)
    SPDRP_COMPATIBLEIDS = &H2                    ' CompatibleIDs (R/W)
    SPDRP_UNUSED0 = &H3                          ' unused
    SPDRP_SERVICE = &H4                          ' Service (R/W)
    SPDRP_UNUSED1 = &H5                          ' unused
    SPDRP_UNUSED2 = &H6                          ' unused
    SPDRP_CLASS = &H7                            ' Class (R--tied to ClassGUID)
    SPDRP_CLASSGUID = &H8                        ' ClassGUID (R/W)
    SPDRP_DRIVER = &H9                           ' Driver (R/W)
    SPDRP_CONFIGFLAGS = &HA                      ' ConfigFlags (R/W)
    SPDRP_MFG = &HB                              ' Mfg (R/W)
    SPDRP_FRIENDLYNAME = &HC                     ' FriendlyName (R/W)
    SPDRP_LOCATION_INFORMATION = &HD             ' LocationInformation (R/W)
    SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = &HE      ' PhysicalDeviceObjectName (R)
    SPDRP_CAPABILITIES = &HF                     ' Capabilities (R)
    SPDRP_UI_NUMBER = &H10                       ' UiNumber (R)
    SPDRP_UPPERFILTERS = &H11                    ' UpperFilters (R/W)
    SPDRP_LOWERFILTERS = &H12                    ' LowerFilters (R/W)
    SPDRP_BUSTYPEGUID = &H13                     ' BusTypeGUID (R)
    SPDRP_LEGACYBUSTYPE = &H14                   ' LegacyBusType (R)
    SPDRP_BUSNUMBER = &H15                       ' BusNumber (R)
    SPDRP_ENUMERATOR_NAME = &H16                 ' Enumerator Name (R)
    SPDRP_SECURITY = &H17                        ' Security (R/W, binary form)
    SPDRP_SECURITY_SDS = &H18                    ' Security (W, SDS form)
    SPDRP_DEVTYPE = &H19                         ' Device Type (R/W)
    SPDRP_EXCLUSIVE = &H1A                       ' Device is exclusive-access (R/W)
    SPDRP_CHARACTERISTICS = &H1B                 ' Device Characteristics (R/W)
    SPDRP_ADDRESS = &H1C                         ' Device Address (R)
    SPDRP_UI_NUMBER_DESC_FORMAT = &H1E           ' UiNumberDescFormat (R/W)
    SPDRP_MAXIMUM_PROPERTY = &H1F                ' Upper bound on ordinals
End Enum

Private Enum REGPROPERTYTYPES
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
    REG_DWORD_LITTLE_ENDIAN = 4
    REG_DWORD_BIG_ENDIAN = 5
    REG_MULTI_SZ = 7
End Enum

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hWnd                  As Long
Private m_bHidCancelled         As Boolean
Private m_uHidSubclass          As SubClassData
Private m_uHidTimer             As FireOnceTimerData
Private m_bAppActive            As Boolean

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print STR_MODULE_NAME & "." & sFunc & ": " & Error
End Sub

'=========================================================================
' Methods
'=========================================================================

Friend Function frHidRegisterRawInput() As Boolean
    Dim uDevice         As RAWINPUTDEVICE
    
    m_hWnd = hWnd
    SubClass m_uHidSubclass, m_hWnd, ObjPtr(Me), AddressOf RedirectSerialScannerHidWndProc
    '--- listen for hid
    uDevice.usUsage = 6         ' Keyboard Usage ID
    uDevice.usUsagePage = 1     ' USB HID Generic Desktop Page
    uDevice.dwFlags = RIDEV_INPUTSINK
    uDevice.hWndTarget = m_hWnd
    Call RegisterRawInputDevices(uDevice, 1, Len(uDevice))
    '--- success
    frHidRegisterRawInput = True
End Function

Friend Function frWndProc( _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "frWndProc"
    Dim lNeeded         As Long
    Dim uRaw            As RAWINPUTKEYBOARD
    Dim baBuffer()      As Byte
    
    On Error GoTo EH
    Select Case wMsg
    Case WM_ACTIVATEAPP
        m_bAppActive = (wParam <> 0)
    Case WM_INPUT
        If m_bAppActive And Not g_oActiveForm Is Nothing And LenB(g_sScannerHidDevice) <> 0 Then
            Call GetRawInputData(lParam, RID_INPUT, ByVal 0&, lNeeded, Len(uRaw.hdr))
            If lNeeded >= Len(uRaw) Then
                ReDim baBuffer(0 To lNeeded) As Byte
                Call GetRawInputData(lParam, RID_INPUT, baBuffer(0), lNeeded, Len(uRaw.hdr))
                Call CopyMemory(uRaw, baBuffer(0), Len(uRaw))
            End If
            If uRaw.hdr.dwType = RIM_TYPEKEYBOARD And uRaw.lMessage = WM_KEYDOWN Then
                If pvGetHidDevice(uRaw.hdr.hDevice) = g_sScannerHidDevice Then ' "\\?\HID#VID_0000&PID_0001#7&8bf8dff&0&0000#{884b96c3-56ef-11d1-bc8c-00a0c91405dd}"
                    Screen.MousePointer = vbHourglass
                    g_oActiveForm.frFireScannerReceive pvHidScanBarcode(uRaw)
                    Screen.MousePointer = vbDefault
                    Exit Function
                End If
            End If
        End If
    End Select
    frWndProc = CallNextWndProc(m_uHidSubclass, hWnd, wMsg, wParam, lParam)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Sub frHidTimer()
    m_bHidCancelled = True
End Sub

Private Function pvHidScanBarcode(uRaw As RAWINPUTKEYBOARD) As String
    Const FUNC_NAME     As String = "pvHidScanBarcode"
    Dim uMsg            As MSG
    Dim lNeeded         As Long
    Dim sBuffer         As String
    Dim baBuffer()      As Byte
    Dim baState(0 To 255) As Byte
    Dim lEatMsg         As Long
    Dim hDevice         As Long
    Dim lTimeout        As Long
    
    On Error GoTo EH
    m_bHidCancelled = False
    hDevice = uRaw.hdr.hDevice
    lTimeout = g_lScannerTimeout
    GoTo InLoop
    Do While Not m_bHidCancelled
        Select Case GetMessage(uMsg, 0, 0, 0)
        Case 0, -1
            Exit Do
        End Select
        Select Case uMsg.message
        Case WM_INPUT
            Call GetRawInputData(uMsg.lParam, RID_INPUT, ByVal 0&, lNeeded, Len(uRaw.hdr))
            If lNeeded >= Len(uRaw) Then
                ReDim baBuffer(0 To lNeeded) As Byte
                Call GetRawInputData(uMsg.lParam, RID_INPUT, baBuffer(0), lNeeded, Len(uRaw.hdr))
                Call CopyMemory(uRaw, baBuffer(0), Len(uRaw))
InLoop:
                If uRaw.hdr.hDevice = hDevice And uRaw.hdr.dwType = RIM_TYPEKEYBOARD Then
                    lEatMsg = uRaw.lMessage
                    If uRaw.lMessage = WM_KEYDOWN Then
                        If uRaw.vKey = vbKeyReturn Or uRaw.vKey = &H6C Then ' VK_SEPARATOR
                            TerminateFireOnceTimer m_uHidTimer
                            InitFireOnceTimer m_uHidTimer, ObjPtr(Me), AddressOf RedirectSerialScannerHidTimerProc
                        Else
                            Call GetKeyState(0)
                            Call GetKeyboardState(baState(0))
                            sBuffer = String(64, 0)
                            If ToUnicode(uRaw.vKey, uRaw.MakeCode, baState(0), StrPtr(sBuffer), Len(sBuffer) - 1, 0) > 0 Then
                                sBuffer = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
                                pvHidScanBarcode = pvHidScanBarcode & sBuffer
                                TerminateFireOnceTimer m_uHidTimer
                                InitFireOnceTimer m_uHidTimer, ObjPtr(Me), AddressOf RedirectSerialScannerHidTimerProc, lTimeout
                            End If
                        End If
                    End If
                    GoTo Continue
                End If
            End If
        Case lEatMsg
            lEatMsg = 0
            GoTo Continue
        End Select
        Call TranslateMessage(uMsg)
        Call DispatchMessage(uMsg)
Continue:
    Loop
    '--- strip prefix if any
    If LenB(g_sScannerPrefix) <> 0 Then
        If g_sScannerPrefix = Left$(pvHidScanBarcode, Len(g_sScannerPrefix)) Then
            pvHidScanBarcode = Mid$(pvHidScanBarcode, Len(g_sScannerPrefix) + 1)
        End If
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
End Function

Public Function EnumHidDevices() As Variant
    Const FUNC_NAME     As String = "EnumHidDevices"
    Dim lNumDevices     As Long
    Dim uList()         As RAWINPUTDEVICELIST
    Dim lIdx            As Long
    Dim vRet            As Variant
    Dim lCount          As Long
    Dim oSetupDevs      As Collection
    Dim oFilterDevs     As Collection
    Dim sID             As String
    Dim vInfo           As Variant
    Dim oCol            As Collection
    
    On Error GoTo EH
    If OsVersion < ucsOsvXp Then
        GoTo QH
    End If
    If GetRawInputDeviceList(ByVal 0&, lNumDevices, Len(uList(0))) = -1 Then
        GoTo QH
    End If
    ReDim uList(0 To lNumDevices) As RAWINPUTDEVICELIST
    If GetRawInputDeviceList(uList(0), lNumDevices, Len(uList(0))) = -1 Then
        GoTo QH
    End If
    ReDim vRet(0 To lNumDevices) As Variant
    Set oFilterDevs = New Collection
    For lIdx = 0 To lNumDevices - 1
        If uList(lIdx).dwType = RIM_TYPEKEYBOARD Then
            sID = pvGetHidDevice(uList(lIdx).hDevice)
            If LenB(sID) <> 0 Then
                vRet(lCount) = sID
                lCount = lCount + 1
                sID = pvGetKeyFromID(sID)
                If LenB(sID) <> 0 And Not SearchCollection(oFilterDevs, sID) Then
                    oFilterDevs.Add sID, sID
                End If
            End If
        End If
    Next
    If lCount > 0 Then
        If oFilterDevs.Count > 0 Then
            pvEnumSetupDevices "USB", oFilterDevs, oSetupDevs
            pvEnumSetupDevices vbNullString, oFilterDevs, oSetupDevs
        End If
        For lIdx = 0 To lCount - 1
            sID = vRet(lIdx)
            If SearchCollection(oSetupDevs, pvGetKeyFromID(sID), RetVal:=vInfo) Then
                vRet(lIdx) = Array(sID, At(vInfo, 0) & IIf(Right$(At(vInfo, 0), 1) <> ")", " (" & At(vInfo, 1) & ")", vbNullString))  '--- name (class)
            Else
                vRet(lIdx) = Array(sID, sID)
            End If
        Next
        ReDim Preserve vRet(0 To lCount - 1) As Variant
        '--- uniqify names
        Set oCol = New Collection
        For lCount = 0 To UBound(vRet)
            If SearchCollection(oCol, vRet(lCount)(1)) Then
                For lIdx = 2 To 100
                    If Not SearchCollection(oCol, vRet(lCount)(1) & " [" & lIdx & "]") Then
                        vRet(lCount)(1) = vRet(lCount)(1) & " [" & lIdx & "]"
                        Exit For
                    End If
                Next
            End If
            oCol.Add vRet(lCount)(1), vRet(lCount)(1)
        Next
        EnumHidDevices = vRet
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetHidDevice(ByVal hDevice As Long) As String
    Dim lNeeded         As Long
        
    If GetRawInputDeviceInfo(hDevice, RIDI_DEVICENAME, 0, lNeeded) <> -1 Then
        pvGetHidDevice = String$(lNeeded + 1, 0)
        Call GetRawInputDeviceInfo(hDevice, RIDI_DEVICENAME, StrPtr(pvGetHidDevice), lNeeded)
        pvGetHidDevice = Left$(pvGetHidDevice, InStr(pvGetHidDevice, Chr$(0)) - 1)
        If InStrRev(pvGetHidDevice, "#") > InStrRev(pvGetHidDevice, "&") Then
            pvGetHidDevice = Left$(pvGetHidDevice, InStrRev(pvGetHidDevice, "#") - 1)
        End If
        If Left$(pvGetHidDevice, 4) = "\\?\" Then
            pvGetHidDevice = Replace(Mid$(pvGetHidDevice, 5), "#", "\")
        End If
    End If
End Function

Private Function pvEnumSetupDevices(Optional sEnumerator As String, Optional oFilter As Collection, Optional oRetVal As Collection) As Collection
    Const FUNC_NAME     As String = "pvEnumSetupDevices"
    Dim hDevInfo        As Long
    Dim uInfo           As SP_DEVINFO
    Dim lIdx            As Long
    Dim sID             As String
    Dim sClass          As String
    Dim sDevice         As String
    Dim uBusRptDeviceDesc As SP_DEVPROPKEY
    Dim sKey            As String

    On Error GoTo EH
    hDevInfo = SetupDiGetClassDevs(ByVal 0&, sEnumerator, 0, DIGCF_PRESENT Or DIGCF_PROFILE Or DIGCF_ALLCLASSES)
    If hDevInfo = INVALID_HANDLE_VALUE Then
        GoTo QH
    End If
    uBusRptDeviceDesc.fmtid(0) = &H540B947E
    uBusRptDeviceDesc.fmtid(1) = &H45BC8B40
    uBusRptDeviceDesc.fmtid(2) = &HB6AA2A8
    uBusRptDeviceDesc.fmtid(3) = &HA2BD4C89
    uBusRptDeviceDesc.pid = 4
    If oRetVal Is Nothing Then
        Set oRetVal = New Collection
    End If
    uInfo.cbSize = Len(uInfo)
    Do
        If SetupDiEnumDeviceInfo(hDevInfo, lIdx, uInfo) = 0 Then
            Exit Do
        End If
        sID = pvGetSetupRegSetting(hDevInfo, uInfo, SPDRP_HARDWAREID)
        If LenB(sID) <> 0 Then
            sKey = pvGetKeyFromID(sID)
            If SearchCollection(oFilter, sKey) Or oFilter Is Nothing Then
                sClass = pvGetSetupRegSetting(hDevInfo, uInfo, SPDRP_CLASS)
                sDevice = pvGetSetupRegSetting(hDevInfo, uInfo, SPDRP_FRIENDLYNAME)
                If LenB(sDevice) = 0 Then
                    sDevice = pvGetSetupSetting(hDevInfo, uInfo, uBusRptDeviceDesc)
                End If
                If LenB(sDevice) = 0 And OsVersion <= ucsOsvXp Then
                    sDevice = pvGetSetupRegSetting(hDevInfo, uInfo, SPDRP_LOCATION_INFORMATION)
                End If
                If LenB(sDevice) = 0 Then
                    sDevice = pvGetSetupRegSetting(hDevInfo, uInfo, SPDRP_DEVICEDESC)
                End If
            Else
                sDevice = vbNullString
                sClass = vbNullString
            End If
            If LenB(sKey) <> 0 And Not SearchCollection(oRetVal, sKey) Then
                oRetVal.Add Array(sDevice, sClass, sKey, UCase$(sID)), sKey
            Else
                oRetVal.Add Array(sDevice, sClass, sKey, UCase$(sID))
            End If
        End If
        lIdx = lIdx + 1
    Loop
    Call SetupDiDestroyDeviceInfoList(hDevInfo)
QH:
    Set pvEnumSetupDevices = oRetVal
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Function pvGetKeyFromID(sID As String, Optional sKeys As String = "VID PID MI|VEN DEV") As String
    Dim vSplit          As Variant
    Dim vKeys           As Variant
    Dim vKey            As Variant
    Dim vElem           As Variant
    
    vSplit = Split(Replace(Replace(Replace(UCase$(sID), "\", " "), "&", " "), "#", " "))
    For Each vKeys In Split(sKeys, "|")
        For Each vKey In Split(vKeys)
            For Each vElem In vSplit
                If Left$(vElem, Len(vKey) + 1) = vKey & "_" Then
                    If LenB(pvGetKeyFromID) <> 0 Then
                        pvGetKeyFromID = pvGetKeyFromID
                    End If
                    pvGetKeyFromID = pvGetKeyFromID & Mid$(vElem, Len(vKey) + 2)
                    GoTo NextLoop
                End If
            Next
            Exit For
NextLoop:
        Next
        If LenB(pvGetKeyFromID) <> 0 Then
            Exit Function
        End If
    Next
    For Each vElem In vSplit
        Select Case vElem
        Case "??", "0000", vbNullString
        Case Else
            If LenB(pvGetKeyFromID) <> 0 Then
                pvGetKeyFromID = vElem
                Exit For
            Else
                pvGetKeyFromID = vElem
            End If
        End Select
    Next
End Function

Private Function pvGetSetupRegSetting(ByVal hDevInfo As Long, uInfo As SP_DEVINFO, ByVal RegSetting As DEVICEPROPERTYINDEX) As String
    Dim lType           As Long
    Dim lSize           As Long
    Dim sBuffer         As String
    Dim lValue          As Long
    
    On Error GoTo QH
    Call SetupDiGetDeviceRegistryProperty(hDevInfo, uInfo, RegSetting, lType, ByVal sBuffer, Len(sBuffer), lSize)
    Select Case lType
    Case 0
        '--- do nothing
    Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ, REG_BINARY
        '--- note: double because of bug under win2000
        sBuffer = String$(2 * (lSize + 1), 0)
        If SetupDiGetDeviceRegistryProperty(hDevInfo, uInfo, RegSetting, lType, ByVal sBuffer, Len(sBuffer), lSize) <> 0 Then
            pvGetSetupRegSetting = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
        End If
    Case REG_DWORD, REG_DWORD_BIG_ENDIAN, REG_DWORD_LITTLE_ENDIAN
        If SetupDiGetDeviceRegistryProperty(hDevInfo, uInfo, RegSetting, lType, lValue, 4, lSize) <> 0 Then
            pvGetSetupRegSetting = lValue
        End If
    Case Else
        pvGetSetupRegSetting = "Unknown reg prop type (" & lType & ")"
    End Select
QH:
End Function

Private Function pvGetSetupSetting(ByVal hDevInfo As Long, uInfo As SP_DEVINFO, uPropKey As SP_DEVPROPKEY) As String
    Dim lType           As Long
    Dim lSize           As Long
    Dim sBuffer         As String
    
    On Error GoTo QH
    Call SetupDiGetDeviceProperty(hDevInfo, uInfo, uPropKey, lType, StrPtr(sBuffer), LenB(sBuffer), lSize, 0)
    Select Case lType
    Case 0
        '--- do nothing
    Case DEVPROP_TYPE_STRING
        sBuffer = String$(lSize + 1, 0)
        If SetupDiGetDeviceProperty(hDevInfo, uInfo, uPropKey, lType, StrPtr(sBuffer), LenB(sBuffer), lSize, 0) <> 0 Then
            pvGetSetupSetting = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
        End If
    Case Else
        pvGetSetupSetting = "Unknown prop type (" & lType & ")"
    End Select
QH:
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Unload(Cancel As Integer)
    TerminateFireOnceTimer m_uHidTimer
    UnSubClass m_uHidSubclass, m_hWnd
End Sub
