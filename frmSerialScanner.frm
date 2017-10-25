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

Private Type RAWINPUTHEADER
    dwType          As Long
    dwSize          As Long
    hDevice         As Long
    wParam          As Long
End Type

Private Type RAWINPUTKEYBOARD
    hdr             As RAWINPUTHEADER
    MakeCode        As Integer
    flags           As Integer
    Reserved        As Integer
    VKey            As Integer
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

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_hWnd                  As Long
Private m_bHidCancelled         As Boolean
Private m_uHidSubclass          As SubClassData
Private m_uHidTimer             As FireOnceTimerData
Private m_bAppActive            As Boolean

Private Sub PrintError(sFunc As String)
    Debug.Print STR_MODULE_NAME & "." & sFunc & ": " & Error
End Sub

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
                    ScreenMousePointer = vbHourglass
                    g_oActiveForm.frFireScannerReceive pvHidScanBarcode(uRaw)
                    ScreenMousePointer = vbDefault
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
                        If uRaw.VKey = vbKeyReturn Or uRaw.VKey = &H6C Then ' VK_SEPARATOR
                            TerminateFireOnceTimer m_uHidTimer
                            InitFireOnceTimer m_uHidTimer, ObjPtr(Me), AddressOf RedirectSerialScannerHidTimerProc
                        Else
                            Call GetKeyState(0)
                            Call GetKeyboardState(baState(0))
                            sBuffer = String(64, 0)
                            If ToUnicode(uRaw.VKey, uRaw.MakeCode, baState(0), StrPtr(sBuffer), Len(sBuffer) - 1, 0) > 0 Then
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

Private Function pvGetHidDevice(ByVal hDevice As Long) As String
    Dim lNeeded             As Long
        
    If GetRawInputDeviceInfo(hDevice, RIDI_DEVICENAME, 0, lNeeded) <> -1 Then
        pvGetHidDevice = String(lNeeded + 1, 0)
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

Private Sub Form_Unload(Cancel As Integer)
    TerminateFireOnceTimer m_uHidTimer
    UnSubClass m_uHidSubclass, m_hWnd
End Sub
