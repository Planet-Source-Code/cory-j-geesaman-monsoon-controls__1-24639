VERSION 5.00
Begin VB.UserControl mSystemButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   825
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   55
   ToolboxBitmap   =   "mSystemButton.ctx":0000
   Begin VB.PictureBox C 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000012&
      Height          =   255
      Index           =   4
      Left            =   600
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox C 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      Height          =   255
      Index           =   3
      Left            =   480
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox C 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   255
      Index           =   2
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox C 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      Height          =   255
      Index           =   1
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox C 
      AutoRedraw      =   -1  'True
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "mSystemButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum enumButtons
Left = 0
Up = 1
Right = 2
Down = 3
End Enum

Private m_Enabled As Boolean
Private BUp As Boolean
Private m_Button As enumButtons
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(ButtonDown As Boolean, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonDown()
Public Event ButtonUp()

Public Sub About()
Attribute About.VB_UserMemId = -552
Load frmAbout
frmAbout.Show vbModal
End Sub

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
If m_Enabled <> New_Enabled Then
  m_Enabled = New_Enabled
  DrawButton
  PropertyChanged "Enabled"
End If
End Property

Public Property Get Button() As enumButtons
  Button = m_Button
End Property

Public Property Let Button(ByVal New_Button As enumButtons)
If m_Button <> New_Button Then
  m_Button = New_Button
  DrawButton
  PropertyChanged "Button"
End If
End Property

Private Sub DrawVLine(X As Long, Y1 As Long, Y2 As Long, Color As Long)
Dim i As Long
i = Y1
Do
DeleteObject SetPixel(UserControl.hdc, X, i, Color)
i = i + 1
Loop Until i > Y2
End Sub

Private Sub DrawHLine(Y As Long, X1 As Long, X2 As Long, Color As Long)
Dim i As Long
i = X1
Do
DeleteObject SetPixel(UserControl.hdc, i, Y, Color)
i = i + 1
Loop Until i > X2
End Sub

Private Sub DrawButton()
Dim mX As Long, mY As Long
UserControl.Cls
If BUp = True Then
DrawVLine 0, 0, UserControl.ScaleHeight - 1, GetPixel(C(0).hdc, 0, 0)
DrawHLine 0, 0, UserControl.ScaleWidth - 1, GetPixel(C(0).hdc, 0, 0)
DrawVLine 1, 1, UserControl.ScaleHeight - 2, GetPixel(C(1).hdc, 0, 0)
DrawHLine 1, 1, UserControl.ScaleWidth - 2, GetPixel(C(1).hdc, 0, 0)
DrawVLine UserControl.ScaleWidth - 2, 1, UserControl.ScaleHeight - 2, GetPixel(C(2).hdc, 0, 0)
DrawHLine UserControl.ScaleHeight - 2, 1, UserControl.ScaleWidth - 1, GetPixel(C(2).hdc, 0, 0)
DrawVLine UserControl.ScaleWidth - 1, 0, UserControl.ScaleHeight, GetPixel(C(3).hdc, 0, 0)
DrawHLine UserControl.ScaleHeight - 1, 0, UserControl.ScaleWidth, GetPixel(C(3).hdc, 0, 0)
mX = UserControl.ScaleWidth \ 2
mY = UserControl.ScaleHeight \ 2
Else
DrawVLine 0, 0, UserControl.ScaleHeight - 1, GetPixel(C(2).hdc, 0, 0)
DrawHLine 0, 0, UserControl.ScaleWidth - 1, GetPixel(C(2).hdc, 0, 0)
DrawVLine 1, 1, UserControl.ScaleHeight - 2, GetPixel(C(3).hdc, 0, 0)
DrawHLine 1, 1, UserControl.ScaleWidth - 2, GetPixel(C(3).hdc, 0, 0)
DrawVLine UserControl.ScaleWidth - 2, 1, UserControl.ScaleHeight - 2, GetPixel(C(0).hdc, 0, 0)
DrawHLine UserControl.ScaleHeight - 2, 1, UserControl.ScaleWidth - 1, GetPixel(C(0).hdc, 0, 0)
DrawVLine UserControl.ScaleWidth - 1, 0, UserControl.ScaleHeight, GetPixel(C(1).hdc, 0, 0)
DrawHLine UserControl.ScaleHeight - 1, 0, UserControl.ScaleWidth, GetPixel(C(1).hdc, 0, 0)
mX = (UserControl.ScaleWidth \ 2) + 1
mY = (UserControl.ScaleHeight \ 2) + 1
End If
Select Case m_Button
Case Left
DrawVLine mX, mY - 2, mY + 2, GetPixel(C(4).hdc, 0, 0)
DrawVLine mX - 1, mY - 1, mY + 1, GetPixel(C(4).hdc, 0, 0)
DrawHLine mY, mX - 2, mX - 2, GetPixel(C(4).hdc, 0, 0)
Case Up
DrawVLine mX, mY - 2, mY - 2, GetPixel(C(4).hdc, 0, 0)
DrawHLine mY - 1, mX - 1, mX + 1, GetPixel(C(4).hdc, 0, 0)
DrawHLine mY, mX - 2, mX + 2, GetPixel(C(4).hdc, 0, 0)
Case Right
DrawVLine mX, mY - 2, mY + 2, GetPixel(C(4).hdc, 0, 0)
DrawVLine mX + 1, mY - 1, mY + 1, GetPixel(C(4).hdc, 0, 0)
DrawHLine mY, mX + 2, mX + 2, GetPixel(C(4).hdc, 0, 0)
Case Down
DrawVLine mX, mY + 2, mY + 2, GetPixel(C(4).hdc, 0, 0)
DrawHLine mY + 1, mX - 1, mX + 1, GetPixel(C(4).hdc, 0, 0)
DrawHLine mY, mX - 2, mX + 2, GetPixel(C(4).hdc, 0, 0)
End Select
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
BUp = True
Enabled = True
m_Button = Down
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
BUp = False
DrawButton
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
RaiseEvent ButtonDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim bD As Boolean
bD = False
If Button = 1 Then
If (X < 0 Or Y < 0 Or X > UserControl.ScaleWidth Or Y > UserControl.ScaleHeight) = True Then
If BUp = False Then
BUp = True
bD = False
DrawButton
RaiseEvent ButtonUp
End If
Else
If BUp = True Then
BUp = False
bD = True
DrawButton
RaiseEvent ButtonDown
End If
End If
End If
RaiseEvent MouseMove(bD, Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
BUp = True
DrawButton
End If
RaiseEvent MouseUp(Button, Shift, X, Y)
RaiseEvent ButtonUp
End Sub

Private Sub UserControl_Paint()
DrawButton
End Sub

Private Sub UserControl_Resize()
DrawButton
End Sub

Private Sub UserControl_Show()
BUp = True
DrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_Button = PropBag.ReadProperty("Button", Left)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Button", m_Button, Left)
End Sub
