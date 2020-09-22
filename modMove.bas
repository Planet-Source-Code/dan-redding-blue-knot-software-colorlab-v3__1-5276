Attribute VB_Name = "modMove"
Option Explicit

Public Const GWL_WNDPROC = (-4)
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_MOVE = &H3


Public origWndProc As Long

' Original Subclass sample by Matt Hart - vbhelp@matthart.com
' http://matthart.com
' Modified to create 'Form_Move' event by Dan Redding

Public Sub SetHook(hwnd, bSet As Boolean)
    If bSet Then
        origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
    ElseIf origWndProc Then
        Dim lRet As Long
        lRet = SetWindowLong(hwnd, GWL_WNDPROC, origWndProc)
    End If
End Sub

Public Function AppWndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Static lOldTop As Long, lOldLeft As Long
    If Msg = WM_MOVE And frmColorRef.WindowState = 0 Then
        If frmColorRef.chkFav Then
            frmFav.left = frmFav.left + (frmColorRef.left - lOldLeft)
            frmFav.top = frmFav.top + (frmColorRef.top - lOldTop)
        End If
        If frmColorRef.chkBig Then
            frmBig.left = frmBig.left + (frmColorRef.left - lOldLeft)
            frmBig.top = frmBig.top + (frmColorRef.top - lOldTop)
        End If
        lOldTop = frmColorRef.top
        lOldLeft = frmColorRef.left
    End If
    AppWndProc = CallWindowProc(origWndProc, hwnd, Msg, wParam, lParam)
End Function
