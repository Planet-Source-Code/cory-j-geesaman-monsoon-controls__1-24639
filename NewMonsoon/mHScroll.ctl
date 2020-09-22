VERSION 5.00
Begin VB.UserControl mHScroll 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   ToolboxBitmap   =   "mHScroll.ctx":0000
   Begin VB.Timer rBT 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   960
   End
   Begin VB.Timer lBT 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1080
      Top             =   960
   End
   Begin prjMonsoonControls.mSystemButton lB 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   450
   End
   Begin prjMonsoonControls.mSystemButton rB 
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   0
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   450
      Button          =   2
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2220
      Begin VB.Label LabMax 
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LabDivider 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         ForeColor       =   &H80000014&
         Height          =   195
         Left            =   1050
         TabIndex        =   2
         Top             =   0
         Width           =   75
      End
      Begin VB.Label LabValue 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "mHScroll"
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
Public Event Change(Value As Long, Max As Long)
Private CtlEnabled As Boolean

Private bOff As Boolean

Public Enum enumRGB
ccRed = 0
ccGreen = 1
ccBlue = 2
End Enum

Public Sub About()
Load frmAbout
frmAbout.Show vbModal
End Sub

Public Function GetRGB(ColorValue As Long, cl As enumRGB)
Dim rCol, gCol, bCol
If cl = ccRed Then
rCol = ColorValue And &H10000FF
GetRGB = rCol
ElseIf cl = ccGreen Then
gCol = (ColorValue And &H100FF00) / (2 ^ 8)
GetRGB = gCol
ElseIf cl = ccBlue Then
bCol = (ColorValue And &H1FF0000) / (2 ^ 16)
GetRGB = bCol
End If
End Function

Public Property Get Enabled() As Boolean
Enabled = CtlEnabled
End Property

Public Property Let Enabled(Data As Boolean)
CtlEnabled = Data
lB.Enabled = Enabled
rB.Enabled = Enabled
LabValue.Enabled = Enabled
LabDivider.Enabled = Enabled
LabMax.Enabled = Enabled
Picture2.Enabled = Enabled
LabValue.Enabled = Enabled
LabDivider.Enabled = Enabled
LabMax.Enabled = Enabled
PropertyChanged Enabled
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = Picture2.BackColor
End Property

Public Property Let BackColor(Data As OLE_COLOR)
Picture2.BackColor = Data
PropertyChanged BackColor
End Property

Public Property Get BarCol1() As OLE_COLOR
BarCol1 = UserControl.ForeColor
End Property

Public Property Let BarCol1(Data As OLE_COLOR)
UserControl.ForeColor = Data
PropertyChanged BarCol1
DrawBar
End Property

Public Property Get BarCol2() As OLE_COLOR
BarCol2 = UserControl.FillColor
End Property

Public Property Let BarCol2(Data As OLE_COLOR)
UserControl.FillColor = Data
PropertyChanged BarCol2
DrawBar
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = LabValue.ForeColor
End Property

Public Property Let ForeColor(Data As OLE_COLOR)
LabValue.ForeColor = Data
LabDivider.ForeColor = Data
LabMax.ForeColor = Data
PropertyChanged ForeColor
End Property

Public Property Get Value() As Integer
Value = CInt(LabValue.Caption)
End Property

Public Property Let Value(Data As Integer)
LabValue.Caption = Data
PropertyChanged Value
DrawBar
RaiseEvent Change(CLng(Data), LabMax.Caption)
End Property

Public Property Get Max() As Integer
Max = CInt(LabMax.Caption)
End Property

Public Property Let Max(Data As Integer)
If Value > Data Then Value = Data
LabMax.Caption = Data
PropertyChanged Max
DrawBar
End Property

Private Sub LabDivider_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2_MouseDown Button, Shift, LabDivider.Left + (X / 15), LabDivider.Top + (Y / 15)
End Sub

Private Sub LabDivider_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabDivider_MouseDown Button, Shift, X, Y
End Sub

Private Sub LabMax_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2_MouseDown Button, Shift, LabMax.Left + (X / 15), LabMax.Top + (Y / 15)
End Sub

Private Sub LabMax_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabMax_MouseDown Button, Shift, X, Y
End Sub

Private Sub LabValue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2_MouseDown Button, Shift, LabValue.Left + (X / 15), LabValue.Top + (Y / 15)
End Sub

Private Sub LabValue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabValue_MouseDown Button, Shift, X, Y
End Sub

Private Sub lB_ButtonDown()
lBT.Enabled = True
End Sub

Private Sub lB_ButtonUp()
lBT.Enabled = False
End Sub

Private Sub lB_Click()
If Value > 1 Then Value = Value - 1
UserControl.SetFocus
End Sub

Private Sub lBT_Timer()
lB_Click
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i, v
If Button = 1 Then
i = (Picture2.ScaleWidth / Max)
v = ((X / i) + 0.5) \ 1
If v < 1 Then v = 1
If v > Max Then v = Max
If Value <> v Then Value = v
End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2_MouseDown Button, Shift, X, Y
End Sub

Private Sub rB_ButtonDown()
rBT.Enabled = True
End Sub

Private Sub rB_ButtonUp()
rBT.Enabled = False
End Sub

Private Sub rB_Click()
If Value < Max Then Value = Value + 1
UserControl.SetFocus
End Sub

Private Sub rBT_Timer()
rB_Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= Picture2.Left And X <= Picture2.Left + Picture2.Width Then
Picture2_MouseDown Button, Shift, X - Picture2.Left, Y
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= Picture2.Left And X <= Picture2.Left + Picture2.Width Then
Picture2_MouseDown Button, Shift, X - Picture2.Left, Y
End If
End Sub

Private Sub UserControl_Paint()
UserControl_Resize
DrawBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
LabValue.Caption = PropBag.ReadProperty("Value", 1)
LabMax.Caption = PropBag.ReadProperty("Max", 1)
LabValue.ForeColor = PropBag.ReadProperty("ForeColor", &H8000000F)
LabDivider.ForeColor = LabValue.ForeColor
LabMax.ForeColor = LabValue.ForeColor
Picture2.BackColor = PropBag.ReadProperty("BackColor", &H8000000C)
UserControl.ForeColor = PropBag.ReadProperty("BarCol1", &H8000000E)
UserControl.FillColor = PropBag.ReadProperty("BarCol2", &H8000000E)
Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
Picture2.Width = UserControl.ScaleWidth - lB.Width - rB.Width - 6
rB.Left = Picture2.Left + Picture2.Width + 3
lB.Height = UserControl.ScaleHeight
Picture2.Height = lB.Height
rB.Height = lB.Height
LabValue.Width = (Picture2.ScaleWidth - LabDivider.Width) / 2
LabMax.Width = LabValue.Width
LabDivider.Left = LabValue.Width
LabMax.Left = LabDivider.Left + LabDivider.Width
LabValue.Top = (Picture2.ScaleHeight / 2) - (LabValue.Height / 2)
LabDivider.Top = (Picture2.ScaleHeight / 2) - (LabDivider.Height / 2)
LabMax.Top = (Picture2.ScaleHeight / 2) - (LabMax.Height / 2)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Value", LabValue.Caption, 1
PropBag.WriteProperty "Max", LabMax.Caption, 1
PropBag.WriteProperty "ForeColor", LabValue.ForeColor, &H8000000F
PropBag.WriteProperty "BackColor", Picture2.BackColor, &H8000000C
PropBag.WriteProperty "BarCol1", UserControl.ForeColor, &H8000000E
PropBag.WriteProperty "BarCol2", UserControl.FillColor, &H8000000E
PropBag.WriteProperty "Enabled", CtlEnabled
End Sub

Public Sub DrawBar()
Dim Go2, i As Long
Go2 = (Picture2.ScaleWidth / Max) * Value
Picture2.Cls
For i = 0 To Go2 Step 1
If i <= Go2 - (Picture2.ScaleWidth / Max) - 1 Then GoTo SkipIt
DrawVerticalLine i, AverageColor(i, Picture2.ScaleWidth)
SkipIt:
Next i
Picture2.Refresh
End Sub

Private Sub DrawVerticalLine(X As Long, Color As Long)
Dim i As Long
i = 0
Do
DeleteObject SetPixel(Picture2.hdc, X, i, Color)
i = i + 1
Loop Until i > Picture2.ScaleHeight
End Sub

Private Function AverageColor(Percent, Ma)
Dim r, g, b
r = ((GetRGB(BarCol1, ccRed) * (Ma - Percent)) + (GetRGB(BarCol2, ccRed) * Percent)) \ Ma
g = ((GetRGB(BarCol1, ccGreen) * (Ma - Percent)) + (GetRGB(BarCol2, ccGreen) * Percent)) \ Ma
b = ((GetRGB(BarCol1, ccBlue) * (Ma - Percent)) + (GetRGB(BarCol2, ccBlue) * Percent)) \ Ma
AverageColor = RGB(r, g, b)
End Function
