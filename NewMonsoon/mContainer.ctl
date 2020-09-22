VERSION 5.00
Begin VB.UserControl mContainer 
   Alignable       =   -1  'True
   BackColor       =   &H8000000C&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   2175
   ToolboxBitmap   =   "mContainer.ctx":0000
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   2040
      ScaleHeight     =   2745
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   255
      Width           =   135
      Begin VB.PictureBox picSelected 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   105
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Image UpN 
      Height          =   255
      Left            =   1200
      Picture         =   "mContainer.ctx":0312
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image DownN 
      Height          =   255
      Left            =   1200
      Picture         =   "mContainer.ctx":0530
      Top             =   2040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image DownU 
      Height          =   255
      Left            =   960
      Picture         =   "mContainer.ctx":074E
      Top             =   2040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image DownD 
      Height          =   255
      Left            =   1080
      Picture         =   "mContainer.ctx":096C
      Top             =   2040
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image sD 
      Height          =   255
      Left            =   2040
      Picture         =   "mContainer.ctx":0B8A
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image UpD 
      Height          =   255
      Left            =   1080
      Picture         =   "mContainer.ctx":0DA8
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image UpU 
      Height          =   255
      Left            =   960
      Picture         =   "mContainer.ctx":0FC6
      Top             =   1680
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image sU 
      Height          =   255
      Left            =   2040
      Picture         =   "mContainer.ctx":11E4
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "mContainer"
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
Private LastV, sUstop As Boolean, sDstop As Boolean, lY As Long, pST As Long, pSH As Long, mW As Boolean
Private mUm As Long, mLm As Long
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Sub About()
Attribute About.VB_UserMemId = -552
Load frmAbout
frmAbout.Show vbModal
End Sub

Private Sub picBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Y >= pST And Y <= pST + pSH Then
lY = CLng(Y)
mW = True
Else
mW = False
Exit Sub
End If
End If
End Sub

Private Sub picBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If mW = False Then Exit Sub
Dim o As Long, i As Long
o = Y - lY
lY = Y
If pST + o < pST Then If sU.Picture Is UpN.Picture Then Exit Sub
If pST + o > pST Then If sD.Picture Is DownN.Picture Then Exit Sub
With UserControl.ContainedControls
If .Count > 0 Then
i = 0
Do
.Item(i).Top = .Item(i).Top - o
i = i + 1
Loop Until i >= .Count
AdjustScrollBar
UserControl.Refresh
End If
End With
End If
End Sub

Private Sub sD_Click()
UserControl_Click
End Sub

Private Sub sD_DblClick()
UserControl_DblClick
End Sub

Private Sub sD_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i, ShowD
Set sD.Picture = DownD.Picture
RaiseEvent MouseDown(Button, Shift, X + sD.Left, Y + sD.Top)
sDstop = False
If UserControl.ContainedControls.Count > 0 Then
Do
i = 0
ShowD = False
Do
With UserControl.ContainedControls(i)
If .Top + .Height + mLm > UserControl.ScaleHeight And .Visible = True Then ShowD = True
.Top = .Top - 1
End With
i = i + 1
Loop Until i >= UserControl.ContainedControls.Count
AdjustScrollBar
DoEvents
Loop Until sDstop = True Or ShowD = False
If ShowD = False Then If Not sD.Picture Is DownN.Picture Then Set sD.Picture = DownN.Picture
End If
End Sub

Private Sub sD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X + sD.Left, Y + sD.Top)
End Sub

Private Sub sD_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set sD.Picture = DownU.Picture
RaiseEvent MouseUp(Button, Shift, X + sD.Left, Y + sD.Top)
sDstop = True
AdjustScrollBar
End Sub

Private Sub sU_Click()
UserControl_Click
End Sub

Private Sub sU_DblClick()
UserControl_DblClick
End Sub

Private Sub sU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i, ShowU
Set sU.Picture = UpD.Picture
RaiseEvent MouseDown(Button, Shift, X + sU.Left, Y + sU.Top)
sUstop = False
If UserControl.ContainedControls.Count > 0 Then
Do
i = 0
ShowU = False
Do
With UserControl.ContainedControls(i)
If .Top - mUm < 0 And .Visible = True Then ShowU = True
.Top = .Top + 1
End With
i = i + 1
Loop Until i >= UserControl.ContainedControls.Count
AdjustScrollBar
DoEvents
Loop Until sUstop = True Or ShowU = False
If ShowU = False Then If Not sU.Picture Is UpN.Picture Then Set sU.Picture = UpN.Picture
End If
End Sub

Private Sub sU_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X + sU.Left, Y + sU.Top)
End Sub

Private Sub sU_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set sU.Picture = UpU.Picture
RaiseEvent MouseUp(Button, Shift, X + sU.Left, Y + sU.Top)
sUstop = True
AdjustScrollBar
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
AdjustScrollBar
End Sub

Private Sub UserControl_InitProperties()
mLm = 60
mUm = 60
End Sub

Private Sub UserControl_Paint()
'UserControl_Resize
End Sub

Public Sub UserControl_Resize()
Dim Value, Max, ShowU, ShowD, i, OffS, tMax, tMin, mM
sU.Left = UserControl.ScaleWidth - sU.Width
sU.Top = 0
sD.Left = UserControl.ScaleWidth - sD.Width
sD.Top = UserControl.ScaleHeight - sD.Height
Value = 0
Max = 0
If UserControl.ContainedControls.Count > 0 Then
ShowU = False
ShowD = False
i = 0
Do
With UserControl.ContainedControls(i)
If .Top + .Height + mLm > UserControl.ScaleHeight And .Visible = True Then
If sD.Visible = False Then ShowD = True
End If
If .Top - mUm < 0 And .Visible = True Then
If sU.Visible = False Then ShowU = True
End If
If .Top + .Height > tMax Then tMax = .Top + .Height
If .Top < tMin Then tMin = .Top
End With
i = i + 1
Loop Until i >= UserControl.ContainedControls.Count
If ShowU = True Then
If sUstop = True Then
If Not sU.Picture Is UpU.Picture Then Set sU.Picture = UpU.Picture
Else
If Not sU.Picture Is UpD.Picture Then Set sU.Picture = UpD.Picture
End If
If sU.Enabled <> True Then sU.Enabled = True
If sU.Visible <> True Then sU.Visible = True
Else
If Not sU.Picture Is UpN.Picture Then Set sU.Picture = UpN.Picture
If sU.Enabled <> False Then sU.Enabled = False
If sU.Visible <> True Then sU.Visible = True
End If
If ShowD = True Then
If sDstop = True Then
If Not sD.Picture Is DownU.Picture Then Set sD.Picture = DownU.Picture
Else
If Not sD.Picture Is DownD.Picture Then Set sD.Picture = DownD.Picture
End If
If sD.Enabled <> True Then sD.Enabled = True
If sD.Visible <> True Then sD.Visible = True
Else
If Not sD.Picture Is DownN.Picture Then Set sD.Picture = DownN.Picture
If sD.Enabled <> False Then sD.Enabled = False
If sD.Visible <> True Then sD.Visible = True
End If
If picBar.Left <> UserControl.ScaleWidth - picBar.Width Then picBar.Left = UserControl.ScaleWidth - picBar.Width
If sU.Visible = True Then
If picBar.Top <> sU.Height Then picBar.Top = sU.Height
Else
If picBar.Top <> 0 Then picBar.Top = 0
End If
If picBar.Width <> sU.Width Then picBar.Width = sU.Width
OffS = 0
mM = mLm + mUm
If sU.Visible = False Then OffS = OffS + sU.Height
If sD.Visible = False Then OffS = OffS + sD.Height
If picBar.Height <> UserControl.ScaleHeight - sU.Height - sD.Height + OffS Then picBar.Height = UserControl.ScaleHeight - sU.Height - sD.Height + OffS
If picBar.ScaleHeight <> tMax - tMin + mM Then picBar.ScaleHeight = tMax - tMin + mM
If pST <> 0 - tMin Then pST = 0 - tMin
If pSH <> UserControl.ScaleHeight + mM Then pSH = UserControl.ScaleHeight + mM
If picSelected.Height <> pSH Then picSelected.Height = pSH
If picSelected.Top <> pST Then picSelected.Top = pST
picBar.Cls
DeleteObject SetPixel(picBar.hdc, 0, ((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight), GetPixel(picSelected.hdc, 0, 0))
StretchBlt picBar.hdc, 0, ((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight), (picBar.Width - 30) / 15, (((picBar.Height - 30) / 15) * ((picSelected.Top + picSelected.Height) / picBar.ScaleHeight)) - (((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight)), picBar.hdc, 0, ((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight), 1, 1, vbSrcCopy
picBar.Refresh
If picBar.Visible <> True Then picBar.Visible = True
Else
If sU.Visible <> False Then sU.Visible = False
If Not sU.Picture Is UpN.Picture Then Set sU.Picture = UpN.Picture
If sD.Visible <> False Then sD.Visible = False
If Not sD.Picture Is DownN.Picture Then Set sD.Picture = DownN.Picture
If picBar.Visible <> False Then picBar.Visible = False
End If
DoEvents
End Sub

Private Sub AdjustScrollBar()
Dim i, tMin, tMax, OffS, ShowD, ShowU, mM
tMin = 0
tMax = 0
If UserControl.ContainedControls.Count > 0 Then
ShowU = False
ShowD = False
i = 0
Do
With UserControl.ContainedControls(i)
If .Top + .Height + mLm > UserControl.ScaleHeight And .Visible = True Then ShowD = True
If .Top - mUm < 0 And .Visible = True Then ShowU = True
If .Top + .Height > tMax Then tMax = .Top + .Height
If .Top < tMin Then tMin = .Top
End With
i = i + 1
Loop Until i >= UserControl.ContainedControls.Count
If ShowU = True Then
If sUstop = True Then
If Not sU.Picture Is UpU.Picture Then Set sU.Picture = UpU.Picture
Else
If Not sU.Picture Is UpD.Picture Then Set sU.Picture = UpD.Picture
End If
If sU.Enabled <> True Then sU.Enabled = True
If sU.Visible <> True Then sU.Visible = True
Else
If Not sU.Picture Is UpN.Picture Then Set sU.Picture = UpN.Picture
If sU.Enabled <> False Then sU.Enabled = False
If sU.Visible <> True Then sU.Visible = True
End If
If ShowD = True Then
If sDstop = True Then
If Not sD.Picture Is DownU.Picture Then Set sD.Picture = DownU.Picture
Else
If Not sD.Picture Is DownD.Picture Then Set sD.Picture = DownD.Picture
End If
If sD.Enabled <> True Then sD.Enabled = True
If sD.Visible <> True Then sD.Visible = True
Else
If Not sD.Picture Is DownN.Picture Then Set sD.Picture = DownN.Picture
If sD.Enabled <> False Then sD.Enabled = False
If sD.Visible <> True Then sD.Visible = True
End If
If picBar.Left <> UserControl.ScaleWidth - picBar.Width Then picBar.Left = UserControl.ScaleWidth - picBar.Width
If sU.Visible = True Then
If picBar.Top <> sU.Height Then picBar.Top = sU.Height
Else
If picBar.Top <> 0 Then picBar.Top = 0
End If
If picBar.Width <> sU.Width Then picBar.Width = sU.Width
OffS = 0
mM = mLm + mUm
If sU.Visible = False Then OffS = OffS + sU.Height
If sD.Visible = False Then OffS = OffS + sD.Height
If picBar.Height <> UserControl.ScaleHeight - sU.Height - sD.Height + OffS Then picBar.Height = UserControl.ScaleHeight - sU.Height - sD.Height + OffS
If picBar.ScaleHeight <> tMax - tMin + mM Then picBar.ScaleHeight = tMax - tMin + mM
If pST <> 0 - tMin Then pST = 0 - tMin
If pSH <> UserControl.ScaleHeight + mM Then pSH = UserControl.ScaleHeight + mM
If picSelected.Height <> pSH Then picSelected.Height = pSH
If picSelected.Top <> pST Then picSelected.Top = pST
picBar.Cls
DeleteObject SetPixel(picBar.hdc, 0, ((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight), GetPixel(picSelected.hdc, 0, 0))
StretchBlt picBar.hdc, 0, ((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight), (picBar.Width - 30) / 15, (((picBar.Height - 30) / 15) * ((picSelected.Top + picSelected.Height) / picBar.ScaleHeight)) - (((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight)), picBar.hdc, 0, ((picBar.Height - 30) / 15) * (picSelected.Top / picBar.ScaleHeight), 1, 1, vbSrcCopy
picBar.Refresh
If picBar.Visible <> True Then picBar.Visible = True
Else
If sU.Visible <> False Then sU.Visible = False
If sD.Visible <> False Then sD.Visible = False
If Not sU.Picture Is UpN.Picture Then Set sU.Picture = UpN.Picture
If Not sD.Picture Is DownN.Picture Then Set sD.Picture = DownN.Picture
If picBar.Visible <> False Then picBar.Visible = False
End If
End Sub

Private Sub UserControl_Show()
sUstop = True
sDstop = True
UserControl_Resize
End Sub

Public Property Get UpperMargin() As Long
    UpperMargin = mUm
End Property

Public Property Let UpperMargin(ByVal New_UpperMargin As Long)
    mUm = New_UpperMargin
    PropertyChanged "UpperMargin"
    AdjustScrollBar
End Property

Public Property Get LowerMargin() As Long
    LowerMargin = mLm
End Property

Public Property Let LowerMargin(ByVal New_LowerMargin As Long)
    mLm = New_LowerMargin
    PropertyChanged "LowerMargin"
    AdjustScrollBar
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged BackColor
End Property

Public Property Get Appearance() As Integer
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get BackStyle() As Integer
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get ContainerHwnd() As Long
Attribute ContainerHwnd.VB_Description = "Returns a handle (from Microsoft Windows) to the window a UserControl is contained in."
    ContainerHwnd = UserControl.ContainerHwnd
End Property

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub PaintPicture(ByVal Picture As Picture, ByVal X1 As Single, ByVal Y1 As Single, Optional ByVal Width1 As Variant, Optional ByVal Height1 As Variant, Optional ByVal X2 As Variant, Optional ByVal Y2 As Variant, Optional ByVal Width2 As Variant, Optional ByVal Height2 As Variant, Optional ByVal Opcode As Variant)
Attribute PaintPicture.VB_Description = "Draws the contents of a graphics file on a Form, PictureBox, or Printer object."
    UserControl.PaintPicture Picture, X1, Y1, Width1, Height1, X2, Y2, Width2, Height2, Opcode
End Sub

Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal Flags As Variant, Optional ByVal X As Variant, Optional ByVal Y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, Flags, X, Y, DefaultMenu
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mUm = PropBag.ReadProperty("UpperMargin", 60)
    mLm = PropBag.ReadProperty("LowerMargin", 60)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", 0)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("UpperMargin", mUm, 60)
    Call PropBag.WriteProperty("LowerMargin", mLm, 60)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, 0)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
