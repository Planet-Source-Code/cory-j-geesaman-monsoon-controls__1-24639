VERSION 5.00
Begin VB.UserControl mFlatButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   ScaleHeight     =   51
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   92
   ToolboxBitmap   =   "mFlatButton.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   240
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "mFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Created By
'Cory J. Geesaman
'www.naven.net/
'------------------------------------
'If This Or Any Of the Accomanying Files Are Disributed
'To Anyone Then Any Appearance Of My Name(Cory J. Geesaman),
'And The Surrounding Text Must Remain Intact.
'All The Files Accomanying This Were Created By Me(Cory J. Geesaman),
'Unless Otherwise Stated.
'------------------------------------
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()

Private Enum enumState
Up = 0
Down = 1
MouseOver = 2
Disabled = 3
End Enum

Private State As enumState, bDown As Boolean

Public Sub About()
Load frmAbout
frmAbout.Show vbModal
End Sub

Public Sub Init()
Image1.Enabled = True
Timer1.Enabled = True
End Sub

Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

Public Sub Refresh()
With UserControl
Select Case State
Case Disabled
Image1.Enabled = False
.Enabled = False
Case Down
UserControl_Resize
Image1.Left = Image1.Left + 1
Image1.Top = Image1.Top + 1
UserControl.Line (0, 0)-(0, .ScaleHeight), &H404040
UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), &HE0E0E0
UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), &HE0E0E0
UserControl.Line (0, 0)-(.ScaleWidth, 0), &H404040
Case MouseOver
UserControl_Resize
UserControl.Line (0, 0)-(0, .ScaleHeight), &HE0E0E0
UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(.ScaleWidth - 1, 0), &H404040
UserControl.Line (.ScaleWidth - 1, .ScaleHeight - 1)-(0, .ScaleHeight - 1), &H404040
UserControl.Line (0, 0)-(.ScaleWidth, 0), &HE0E0E0
Case Up
UserControl_Resize
UserControl.Cls
End Select
End With
End Sub

Private Sub Image1_Click()
UserControl_Click
End Sub

Private Sub Image1_DblClick()
UserControl_DblClick
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub Timer1_Timer()
If State <> Disabled Then
Dim tRect As RECT, mPos As POINTAPI
GetWindowRect UserControl.hWnd, tRect
GetCursorPos mPos
If mPos.X > tRect.Left And mPos.X < tRect.Right And mPos.Y > tRect.Top And mPos.Y < tRect.Bottom Then
If bDown = False Then
If State = MouseOver Then Exit Sub
State = MouseOver
Refresh
Else
If State <> Down Then
If State = Down Then Exit Sub
State = Down
Refresh
End If
End If
Else
If State = Up Then Exit Sub
State = Up
Refresh
End If
End If
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDown = True
State = Down
Refresh
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
bDown = False
State = Up
Refresh
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
  Set Picture = Image1.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
  Set Image1.Picture = New_Picture
  Image1.Left = ((UserControl.ScaleWidth / 2) - (Image1.Width / 2)) \ 1
  Image1.Top = ((UserControl.ScaleHeight / 2) - (Image1.Height / 2)) \ 1
  PropertyChanged "Picture"
End Property

Private Sub UserControl_Resize()
  Image1.Left = ((UserControl.ScaleWidth / 2) - (Image1.Width / 2)) \ 1
  Image1.Top = ((UserControl.ScaleHeight / 2) - (Image1.Height / 2)) \ 1
  RaiseEvent Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub
