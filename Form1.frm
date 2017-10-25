VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2316
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   1512
      TabIndex        =   3
      Top             =   924
      Width           =   4800
   End
   Begin VB.ComboBox cobHidDevices 
      Height          =   288
      Left            =   1512
      TabIndex        =   2
      Top             =   252
      Width           =   4800
   End
   Begin VB.Label Label2 
      Caption         =   "HID device"
      Height          =   432
      Left            =   336
      TabIndex        =   1
      Top             =   252
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   348
      Left            =   336
      TabIndex        =   0
      Top             =   1764
      Width           =   6060
   End
End
Attribute VB_Name = "Form1"
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
Private Const STR_MODULE_NAME As String = "Form1"

'=========================================================================
' API
'=========================================================================

'--- for GetRawInputDeviceInfo
Private Const RIDI_DEVICENAME                       As Long = &H20000007
Private Const RIM_TYPEKEYBOARD                      As Long = 1

Private Declare Function GetRawInputDeviceList Lib "user32" (pRawInputDeviceList As Any, puiNumDevices As Long, ByVal cbSize As Long) As Long
Private Declare Function GetRawInputDeviceInfo Lib "user32" Alias "GetRawInputDeviceInfoW" (ByVal hDevice As Long, ByVal uiCommand As Long, ByVal pData As Long, pcbSize As Long) As Long

Private Type RAWINPUTDEVICELIST
    hDevice         As Long
    dwType          As Long
End Type

Private Sub PrintError(sFunc As String)
    Debug.Print STR_MODULE_NAME & "." & sFunc & ": " & Error
End Sub

Private Sub cobHidDevices_Click()
    g_sScannerHidDevice = cobHidDevices.Text
End Sub

Private Sub Form_Activate()
    Set g_oActiveForm = Me
End Sub

Private Sub Form_Load()
    Dim vElem           As Variant

    For Each vElem In pvEnumHidDevices
        cobHidDevices.AddItem vElem
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSerialScanner
End Sub

'--- callback from frmSerialScanner
Public Sub frFireScannerReceive(sBarcode As String)
    Label1.Caption = sBarcode & "@" & Timer
End Sub

Private Function pvEnumHidDevices() As Variant
    Const FUNC_NAME     As String = "pvEnumHidDevices"
    Dim lNumDevices         As Long
    Dim uList()             As RAWINPUTDEVICELIST
    Dim lIdx                As Long
    Dim vRet                As Variant
    Dim lCount              As Long
    
    On Error GoTo EH
    If GetRawInputDeviceList(ByVal 0&, lNumDevices, Len(uList(0))) = -1 Then
        GoTo QH
    End If
    ReDim uList(0 To lNumDevices) As RAWINPUTDEVICELIST
    If GetRawInputDeviceList(uList(0), lNumDevices, Len(uList(0))) = -1 Then
        GoTo QH
    End If
    ReDim vRet(0 To lNumDevices) As String
    For lIdx = 0 To lNumDevices - 1
        If uList(lIdx).dwType = RIM_TYPEKEYBOARD Then
            vRet(lCount) = pvGetHidDevice(uList(lIdx).hDevice)
            lCount = lCount + 1
        End If
    Next
    If lCount > 0 Then
        ReDim Preserve vRet(0 To lCount - 1) As String
        pvEnumHidDevices = vRet
    End If
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
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

