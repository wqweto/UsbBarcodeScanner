Attribute VB_Name = "mdGlobals"
'=========================================================================
'
' UsbBarcodeScanner (c) 2017 by wqweto@gmail.com
'
' A VB6 sample project for intercepting USB/HID devices input
'
'=========================================================================
Option Explicit
Private Const STR_MODULE_NAME As String = "mdGlobals"

Public g_sScannerHidDevice          As String
Public g_lScannerTimeout            As Long
Public g_sScannerPrefix             As String
Public g_oActiveForm                As Object

Private Sub PrintError(sFunc As String)
    Debug.Print STR_MODULE_NAME & "." & sFunc & ": " & Error
End Sub

Private Sub Main()
    g_lScannerTimeout = 100         '--- wait for 100 ms with no input before calling frFireScannerReceive
    g_sScannerPrefix = "~"          '--- prefix will be stripped if present
    g_sScannerHidDevice = ""        '--- usually read this user setting from registry
    frmSerialScanner.frHidRegisterRawInput
    With New Form1
        .Show
    End With
End Sub

Property Get ScreenMousePointer() As MousePointerConstants
    ScreenMousePointer = Screen.MousePointer
End Property

Property Let ScreenMousePointer(ByVal eValue As MousePointerConstants)
    Screen.MousePointer = eValue
End Property

Public Function RedirectSerialScannerHidWndProc( _
            ByVal This As frmSerialScanner, _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "RedirectSerialScannerHidWndProc"
    
    On Error GoTo EH
    RedirectSerialScannerHidWndProc = This.frWndProc(hWnd, wMsg, wParam, lParam)
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Sub RedirectSerialScannerHidTimerProc( _
            Data As FireOnceTimerData, _
            ByVal This As frmSerialScanner, _
            ByVal hWnd As Long, _
            ByVal wMsg As Long, _
            ByVal idEvent As Long, _
            ByVal dwTime As Long)
    #If hWnd And wMsg And dwTime Then '--- touch args
    #End If
    Data.TimerID = idEvent
    TerminateFireOnceTimer Data
    This.frHidTimer
End Sub



