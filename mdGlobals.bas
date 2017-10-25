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

'=========================================================================
' API
'=========================================================================

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Public g_sScannerHidDevice          As String
Public g_lScannerTimeout            As Long
Public g_sScannerPrefix             As String
Public g_oActiveForm                As Object

Public Enum UcsOsVersionEnum
    ucsOsvNt4 = 400
    ucsOsvWin98 = 410
    ucsOsvWin2000 = 500
    ucsOsvXp = 501
    ucsOsvVista = 600
    ucsOsvWin7 = 601
    ucsOsvWin8 = 602
    [ucsOsvWin8.1] = 603
    ucsOsvWin10 = 1000
End Enum

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print STR_MODULE_NAME & "." & sFunc & ": " & Error
End Sub

'=========================================================================
' Properties
'=========================================================================

Public Property Get OsVersion() As UcsOsVersionEnum
    Static lVersion     As Long
    Dim uVer            As OSVERSIONINFO
    
    If lVersion = 0 Then
        If lVersion = 0 Then
            uVer.dwOSVersionInfoSize = Len(uVer)
            If GetVersionEx(uVer) Then
                lVersion = uVer.dwMajorVersion * 100 + uVer.dwMinorVersion
            End If
        End If
    End If
    OsVersion = lVersion
End Property

'=========================================================================
' Functions
'=========================================================================

Private Sub Main()
    g_lScannerTimeout = 100         '--- wait for 100 ms with no input before calling frFireScannerReceive
    g_sScannerPrefix = "~"          '--- prefix will be stripped if present
    g_sScannerHidDevice = ""        '--- usually read this user setting from registry
    frmSerialScanner.frHidRegisterRawInput
    With New Form1
        .Show
    End With
End Sub

Public Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Public Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub

Public Function At(vData As Variant, ByVal lIdx As Long, Optional sDefault As String) As String
    On Error Resume Next
    At = sDefault
    At = vData(lIdx)
    On Error GoTo 0
End Function

'= redirectors ===========================================================

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



