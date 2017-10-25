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

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_vHidDevices           As Variant

'=========================================================================
' Methods
'=========================================================================

'--- callback from frmSerialScanner
Public Sub frFireScannerReceive(sBarcode As String)
    Label1.Caption = sBarcode & "@" & Timer
End Sub

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Dim vElem           As Variant
    
    m_vHidDevices = frmSerialScanner.EnumHidDevices
    For Each vElem In m_vHidDevices
        cobHidDevices.AddItem vElem(1)
    Next
End Sub

Private Sub cobHidDevices_Click()
    If cobHidDevices.ListIndex >= 0 Then
        g_sScannerHidDevice = m_vHidDevices(cobHidDevices.ListIndex)(0)
    Else
        g_sScannerHidDevice = vbNullString
    End If
End Sub

Private Sub Form_Activate()
    Set g_oActiveForm = Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSerialScanner
End Sub

