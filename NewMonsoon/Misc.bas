Attribute VB_Name = "Misc"
Public Sub MoveForm(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long, ParentHWND As Long, DropMode As Long)
Dim X As Long
Dim Y As Long
Dim pos As POINTAPI
Dim window As RECT
Static OLDdropmode As Long
GetCursorPos pos
X = pos.X
Y = pos.Y
GetWindowRect ParentHWND, window
DropMode = 0
formWidth = 200
formHeight = 300
formLeft = X
formTop = Y
If Y < window.Bottom And Y > window.Top Then
  If X < window.Right And X > window.Left Then
    If X - window.Left < 100 Then
      formWidth = 200
      formHeight = (window.Bottom - window.Top)
      formLeft = X
      formTop = Y
      DropMode = 3
    End If
    If window.Right - X < 100 Then
      formWidth = 200
      formHeight = (window.Bottom - window.Top)
      formLeft = X
      formTop = Y
      DropMode = 4
    End If
    If Y - window.Top < 100 Then
      formWidth = (window.Right - window.Left)
      formHeight = 200
      formLeft = X
      formTop = Y
      DropMode = 1
    End If
    If window.Bottom - Y < 100 Then
      formWidth = (window.Right - window.Left)
      formHeight = 200
      formLeft = X
      formTop = Y
      DropMode = 2
    End If
  End If
End If
If DropMode <> OLDdropmode Then
  If DropMode = 0 Then
    Screen.MousePointer = 0
  End If
  If DropMode = 1 Then
    Screen.MouseIcon = LoadResPicture("top", vbResCursor)
    Screen.MousePointer = 99
  End If
  If DropMode = 2 Then
    Screen.MouseIcon = LoadResPicture("bottom", vbResCursor)
    Screen.MousePointer = 99
  End If
  If DropMode = 3 Then
    Screen.MouseIcon = LoadResPicture("left", vbResCursor)
    Screen.MousePointer = 99
  End If
  If DropMode = 4 Then
    Screen.MouseIcon = LoadResPicture("right", vbResCursor)
    Screen.MousePointer = 99
  End If
  OLDdropmode = DropMode
  End If
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
Dim lFlag As Long
If SetOnTop Then
  lFlag = HWND_TOPMOST
Else
  lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hWnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Sub DragForm(Frm As Form)
On Local Error Resume Next
Call ReleaseCapture
Call SendMessage(Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub

Public Sub DragControl(Ctl As Control)
On Local Error Resume Next
Call ReleaseCapture
Call SendMessage(Ctl.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
End Sub
