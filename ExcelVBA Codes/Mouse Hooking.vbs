Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Type LLMOUSEHOOKSTRUCT
    pt As POINTAPI
    mousedata As Long
    flages As Long
    time As Long
    #If VBA7 Then
        dwExtraInfo As LongPtr
    #Else
        dwExtraInfo As Long
    #End If
End Type


#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal hData As LongPtr) As Long
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As LongPtr, ByVal lpString As String) As LongPtr
    Private Declare PtrSafe Function RegisterHotKey Lib "user32" (ByVal hWnd As LongPtr, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    Private Declare PtrSafe Function UnregisterHotKey Lib "user32" (ByVal hWnd As LongPtr, ByVal id As Long) As Long
    Private hHook As LongPtr
#Else
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As Long) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
    Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
    Private Declare PtrSafe Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    Private Declare PtrSafe Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
    Private hHook As Long
#End If
 
Private Const WH_MOUSE_LL As Long = 14
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const MOD_CONTROL As Long = &H2
Private Const WM_MOUSEMOVE As Long = &H200
Private lCtrlKey As Long

Public Sub HookTheMouse()
    #If VBA7 Then
        Dim lHinstance As LongPtr
        lHinstance = Application.HinstancePtr
    #Else
        Dim lHinstance As Long
        lHinstance = Application.Hinstance
    #End If
    hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LowLevelMouseProc, lHinstance, 0)
    SetProp Application.hWnd, "hHook", hHook
End Sub


Public Sub UnHookTheMouse()
    UnhookWindowsHookEx GetProp(Application.hWnd, "hHook")
    RemoveProp Application.hWnd, "hHook"
    Call UnregisterHotKey(0, lCtrlKey)
End Sub


Private Function LowLevelMouseProc _
(ByVal idHook As Long, ByVal wParam As LongPtr, lParam As LLMOUSEHOOKSTRUCT) As LongPtr
    Dim oObjectUnderMouse As Range
    
    On Error Resume Next
    Set oObjectUnderMouse = ActiveWindow.RangeFromPoint(lParam.pt.x, lParam.pt.y)
    If wParam = WM_MOUSEMOVE Then
        If TypeName(oObjectUnderMouse) = "Range" Then
            If oObjectUnderMouse.Hyperlinks.Count <> 0 Then
                Call RegisterHotKey(0, lCtrlKey, MOD_CONTROL, VBA.vbKeyControl)
            Else
                Call UnregisterHotKey(0, lCtrlKey)
            End If
        End If
    End If
    If wParam = WM_LBUTTONDOWN Then
        If TypeName(oObjectUnderMouse) = "Range" Then
            If oObjectUnderMouse.Hyperlinks.Count <> 0 Then
                If GetAsyncKeyState(vbKeyControl) = 0 Then
                    LowLevelMouseProc = -1
                    Exit Function
                End If
            End If
        End If
    End If
    LowLevelMouseProc = CallNextHookEx(GetProp(Application.hWnd, "hHook"), idHook, wParam, ByVal lParam)
End Function

