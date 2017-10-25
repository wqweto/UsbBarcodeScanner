Attribute VB_Name = "mdSubclass"
'=========================================================================
'
' UsbBarcodeScanner (c) 2017 by wqweto@gmail.com
'
' A VB6 sample project for intercepting USB/HID devices input
'
'=========================================================================
Option Explicit
DefObj A-Z

'==============================================================================
' API
'==============================================================================

'--- for Get/SetWindowLong
Private Const GWL_WNDPROC                       As Long = -4
'--- for VirtualQuery
Private Const PAGE_EXECUTE_READWRITE            As Long = &H40
'--- windows messages
Private Const WM_NCDESTROY                      As Long = &H82

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadID As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Private Type ThunkBytes
    Thunk(5)                As Long
End Type

Public Type PushParamThunk
    pfn                     As Long
    Code                    As ThunkBytes
End Type

Public Type SubClassData
    WndProcNext             As Long
    WndProcOrig             As Long
    WndProcThunkThis        As PushParamThunk
    #If DEBUGWINDOWPROC Then
        dbg_Hook            As WindowProcHook
    #End If
End Type

Public Type FireOnceTimerData
    TimerID                 As Long
    TimerProcThunkData      As PushParamThunk
    TimerProcThunkThis      As PushParamThunk
End Type

Public Type WindowsHookData
    hhkNext                 As Long
    hhkThunk                As PushParamThunk
    #If DEBUGHOOKPROC Then
        dbg_Hook            As HookProcHook
    #End If
End Type

Public g_cWndProcOrig       As New Collection

Public Sub InitPushParamThunk(Thunk As PushParamThunk, ByVal ParamValue As Long, ByVal pfnDest As Long)
'push [esp]
'mov eax, 16h // Dummy value for parameter value
'mov [esp + 4], eax
'nop // Adjustment so the next long is nicely aligned
'nop
'nop
'mov eax, 1234h // Dummy value for function
'jmp eax
'nop
'nop
    Dim dwDummy             As Long
    
    With Thunk.Code
        .Thunk(0) = &HB82434FF
        .Thunk(1) = ParamValue
        .Thunk(2) = &H4244489
        .Thunk(3) = &HB8909090
        .Thunk(4) = pfnDest
        .Thunk(5) = &H9090E0FF
        Call VirtualProtect(.Thunk(0), Len(Thunk), PAGE_EXECUTE_READWRITE, dwDummy)
    End With
    Thunk.pfn = VarPtr(Thunk.Code)
End Sub

Public Sub SubClass(Data As SubClassData, ByVal hWnd As Long, ByVal ThisPtr As Long, ByVal pfnRedirect As Long)
    If hWnd = 0 Then
        Exit Sub
    End If
    With Data
        If .WndProcOrig Then
            SetWindowLong hWnd, GWL_WNDPROC, .WndProcOrig
            .WndProcOrig = 0
            .WndProcNext = 0
        End If
        InitPushParamThunk .WndProcThunkThis, ThisPtr, pfnRedirect
#If DEBUGWINDOWPROC Then
        If Not CreateInstance("DbgWindowProc.WindowProcHookCreator") Is Nothing Then
            Set .dbg_Hook = CreateWindowProcHook
        End If
        If Not .dbg_Hook Is Nothing Then
            With .dbg_Hook
                .SetMainProc Data.WndProcThunkThis.pfn
                Data.WndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
                .SetDebugProc Data.WndProcNext
            End With
        Else
            .WndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, .WndProcThunkThis.pfn)
        End If
#Else
        .WndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, .WndProcThunkThis.pfn)
#End If
        .WndProcOrig = .WndProcNext
#If DebugMode Then
        If Not SearchCollection(g_cWndProcOrig, "#" & hWnd) Then
            g_cWndProcOrig.Add .WndProcNext, "#" & hWnd
        End If
#Else
        On Error Resume Next
        g_cWndProcOrig.Add .WndProcNext, "#" & hWnd
#End If
    End With
End Sub

Public Sub UnSubClass(Data As SubClassData, ByVal hWnd As Long)
    Const FUNC_NAME     As String = "UnSubClass"
    Dim pfn             As Long
    Dim lWndProc        As Long
    
    With Data
        If .WndProcOrig Then
            pfn = .WndProcThunkThis.pfn
#If DEBUGWINDOWPROC Then
            If Not .dbg_Hook Is Nothing Then
                pfn = .dbg_Hook.ProcAddress
            End If
#End If
            lWndProc = GetWindowLong(hWnd, GWL_WNDPROC)
            If lWndProc <> pfn And (lWndProc And &HF0000000) = 0 Then
                #If DebugMode Then
                    DebugPrint FUNC_NAME, "mdSubclass", "Skip unsubclass! pfn = &H" & Hex(pfn) & " <> &H" & Hex(lWndProc)
                #End If
            Else
                Call SetWindowLong(hWnd, GWL_WNDPROC, .WndProcOrig)
            End If
            .WndProcOrig = 0
            .WndProcNext = 0
        End If
#If DEBUGWINDOWPROC Then
        Set .dbg_Hook = Nothing
#End If
#If DebugMode Then
        RemoveCollection g_cWndProcOrig, "#" & hWnd
#Else
        On Error Resume Next
        g_cWndProcOrig.Remove "#" & hWnd
#End If
    End With
End Sub

Public Function CallNextWndProc(Data As SubClassData, ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Data.WndProcNext <> 0 Then
        CallNextWndProc = CallWindowProc(Data.WndProcNext, hWnd, wMsg, wParam, lParam)
        If wMsg = WM_NCDESTROY Then
            Data.WndProcNext = 0
            Data.WndProcOrig = 0
        End If
    End If
End Function

Public Sub InitFireOnceTimer(Data As FireOnceTimerData, ByVal ThisPtr As Long, ByVal pfnRedirect As Long, Optional ByVal Delay As Long)
    With Data
        InitPushParamThunk .TimerProcThunkData, VarPtr(Data), pfnRedirect
        InitPushParamThunk .TimerProcThunkThis, ThisPtr, .TimerProcThunkData.pfn
        .TimerID = SetTimer(0, 0, Delay, .TimerProcThunkThis.pfn)
    End With
End Sub

Public Sub TerminateFireOnceTimer(Data As FireOnceTimerData)
    With Data
        If .TimerID <> 0 Then
            Call KillTimer(0, .TimerID)
            .TimerID = 0
        End If
    End With
End Sub

'hMod and ThreadID are likely never to be used in VB because
'it isn't equipped to do global hooks (except for journal hooks
'which call back on the same thread). However, these are provided
'for completeness. If ThisPtr is 0, then pfnRedirect is passed the
'next hook procedure as the extra first parameter.
Public Sub StartWindowsHook(Data As WindowsHookData, ByVal HookType As Long, ByVal ThisPtr As Long, ByVal pfnRedirect As Long, Optional ByVal hMod As Long = -1, Optional ByVal ThreadID As Long = -1)
    With Data
        If .hhkNext Then
            UnhookWindowsHookEx .hhkNext
            .hhkNext = 0
        End If
        InitPushParamThunk .hhkThunk, ThisPtr, pfnRedirect
        If ThreadID = -1 Then ThreadID = App.ThreadID
        If hMod = -1 Then hMod = 0
#If DEBUGHOOKPROC Then
        Set .dbg_Hook = CreateInstance("DbgHookProc.HookProcHook")
        If Not .dbg_Hook Is Nothing Then
            With .dbg_Hook
                .SetMainProc Data.hhkThunk.pfn
                Data.hhkNext = SetWindowsHookEx(HookType, .ProcAddress, hMod, ThreadID)
                .SetDebugHandle Data.hhkNext
            End With
        Else
            .hhkNext = SetWindowsHookEx(HookType, .hhkThunk.pfn, hMod, ThreadID)
        End If
#Else
        .hhkNext = SetWindowsHookEx(HookType, .hhkThunk.pfn, hMod, ThreadID)
#End If
        'If a This pointer isn't provided, then pass the next
        'hook to the callback function. Reinitializing the thunk
        'will not change its pfn value.
        If ThisPtr = 0 Then InitPushParamThunk .hhkThunk, .hhkNext, pfnRedirect
    End With
End Sub

Public Sub StopWindowsHook(Data As WindowsHookData)
    With Data
        If .hhkNext Then
            UnhookWindowsHookEx .hhkNext
            .hhkNext = 0
        End If
#If DEBUGHOOKPROC Then
        Set .dbg_Hook = Nothing
#End If
    End With
End Sub

'==============================================================================
' Sample redirectors
'==============================================================================

'Public Function RedirectFormWndProc( _
'            ByVal This As cFormManager, _
'            ByVal hWnd As Long, _
'            ByVal wMsg As Long, _
'            ByVal wParam As Long, _
'            ByVal lParam As Long) As Long
'    RedirectFormWndProc = This.frWndProc(hWnd, wMsg, wParam, lParam)
'End Function

'Public Sub RedirectMenuExtTimerProc( _
'            Data As FireOnceTimerData, _
'            ByVal This As cMenuExtension, _
'            ByVal hWnd As Long, _
'            ByVal wMsg As Long, _
'            ByVal idEvent As Long, _
'            ByVal dwTime As Long)
'    #If hWnd And wMsg And dwTime Then '--- touch
'    #End If
'    Data.TimerID = idEvent
'    TerminateFireOnceTimer Data
'    This.frTimer
'End Sub

'Public Function RedirectApplicationGetMessageHookProc( _
'            ByVal This As cApplication, _
'            ByVal nCode As Long, _
'            ByVal wParam As Long, _
'            ByVal lParam As Long) As Long
'    RedirectApplicationGetMessageHookProc = This.frGetMessageHookProc(nCode, wParam, lParam)
'End Function

