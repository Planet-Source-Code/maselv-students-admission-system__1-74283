Attribute VB_Name = "Mdl_SystemTray"
'****************************************************************************************************
'*  Author:             Masika .S. Elvas                                                            *
'*  Gender:             Male                                                                        *
'*  Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*  Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*  E-mail Address:     elvasmasika@lexeme-kenya.com\masika_elvas@live.com                          *
'*  Location            BUNGOMA, KENYA                                                              *
'****************************************************************************************************
'SystemTray Module (Belongs to the SystemTray Class)

Option Explicit

' Public Constants
Public Const GWL_USERDATA    As Long = -21
Public Const WM_USER_SYSTRAY As Long = &H405

' Private Variables
Private m_MessageInited      As Boolean
Private m_TaskbarRestart     As Long

' Public API
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

' Private API's
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function CreateRef(ByRef cObject As clsSystemTray) As Long
    Call CopyMemory(ByVal VarPtr(CreateRef), ByVal VarPtr(cObject), 4)
End Function

Public Function GetVersion(ByVal nValue As Long) As Long
    Call CopyMemory(GetVersion, ByVal nValue, 2)
End Function

Public Function Pass(ByVal nValue As Long) As Long
    Pass = nValue
End Function

Public Function SysTrayWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Const WM_TIMER As Long = &H113
    
    Dim cstObject  As clsSystemTray
    
    If Not m_MessageInited Then InitMessage
    
    If (uMsg = WM_USER_SYSTRAY) Or (uMsg = WM_TIMER) Then
        
        Set cstObject = DeRef(GetWindowLong(hWnd, GWL_USERDATA))
        Call cstObject.ProcessMessage(wParam, lParam)
        DestroyRef VarPtr(cstObject)
        
    ElseIf uMsg = m_TaskbarRestart Then
        
        Set cstObject = DeRef(GetWindowLong(hWnd, GWL_USERDATA))
        Call cstObject.RecreateIcon
        DestroyRef VarPtr(cstObject)
        
    End If
    
    SysTrayWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)

End Function

Public Sub DestroyRef(ByVal nObject As Long)
    Dim lngValue As Long
    Call CopyMemory(ByVal nObject, ByVal VarPtr(lngValue), 4)
End Sub

Private Function DeRef(ByVal nPointer As Long) As clsSystemTray
    Call CopyMemory(ByVal VarPtr(DeRef), ByVal VarPtr(nPointer), 4)
End Function

Private Function InitMessage()
    m_MessageInited = True
    m_TaskbarRestart = RegisterWindowMessage("TaskbarCreated")
End Function
