VERSION 5.00
Begin VB.UserControl mProgress 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   275
   ToolboxBitmap   =   "mProgress.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   1815
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1815
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   1215
         Begin VB.Label lPercent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calisto MT"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   30
         End
      End
   End
End
Attribute VB_Name = "mProgress"
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
Private vPercent As Integer

Public Sub About()
Load frmAbout
frmAbout.Show vbModal
End Sub

Public Property Get BackStyle() As Boolean
Dim a
If UserControl.BackStyle = 0 Then a = False Else a = True
BackStyle = a
End Property

Public Property Let BackStyle(Data As Boolean)
Dim a
If Data = True Then a = 1 Else a = 0
UserControl.BackStyle = a
PropertyChanged "BackStyle"
ReDrawBar
End Property

Public Property Get Appearance3D() As Boolean
Dim a
If UserControl.Appearance = 0 Then a = False Else a = True
Appearance3D = a
End Property

Public Property Let Appearance3D(Data As Boolean)
Dim a
If Data = True Then a = 1 Else a = 0
UserControl.Appearance = a
PropertyChanged "Appearance3D"
End Property

Public Property Get BorderStyle() As Boolean
Dim a
If UserControl.BorderStyle = 0 Then a = False Else a = True
BorderStyle = a
End Property

Public Property Let BorderStyle(Data As Boolean)
Dim a
If Data = True Then a = 1 Else a = 0
UserControl.BorderStyle = a
PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(Data As OLE_COLOR)
UserControl.BackColor = Data
PropertyChanged "BackColor"
ReDrawBar
End Property

Public Property Get FColor1() As OLE_COLOR
FColor1 = UserControl.ForeColor
End Property

Public Property Let FColor1(Data As OLE_COLOR)
UserControl.ForeColor = Data
PropertyChanged "FColor1"
ReDrawBar
End Property

Public Property Get FColor2() As OLE_COLOR
FColor2 = UserControl.FillColor
End Property

Public Property Let FColor2(Data As OLE_COLOR)
UserControl.FillColor = Data
PropertyChanged "FColor2"
ReDrawBar
End Property

Public Property Get Percent() As Integer
Percent = vPercent
End Property

Public Property Let Percent(Data As Integer)
vPercent = Data
PropertyChanged "Percent"
UserControl.Refresh
ReDrawBar
End Property

Public Property Get LabelVisible() As Boolean
LabelVisible = lPercent.Visible
End Property

Public Property Let LabelVisible(Data As Boolean)
lPercent.Visible = Data
PropertyChanged "LabelVisible"
ReDrawBar
End Property

Public Property Get LabelColor() As OLE_COLOR
LabelColor = lPercent.ForeColor
End Property

Public Property Let LabelColor(Data As OLE_COLOR)
lPercent.ForeColor = Data
PropertyChanged "LabelColor"
ReDrawBar
End Property

Public Property Get LabelBColor() As OLE_COLOR
LabelBColor = lPercent.BackColor
End Property

Public Property Let LabelBColor(Data As OLE_COLOR)
lPercent.BackColor = Data
PropertyChanged "LabelBColor"
ReDrawBar
End Property

Public Property Get LabelBackStyle() As Boolean
Dim a
If lPercent.BackStyle = 0 Then a = False Else a = True
LabelBackStyle = a
End Property

Public Property Let LabelBackStyle(Data As Boolean)
Dim a
If Data = True Then a = 1 Else a = 0
lPercent.BackStyle = a
PropertyChanged "LabelBackStyle"
ReDrawBar
End Property

Public Sub ReDrawBar()
If vPercent < 0 Then vPercent = 100 - vPercent
Picture1.Height = UserControl.ScaleHeight
Picture1.Left = 0
Picture2.Width = UserControl.ScaleWidth
Picture2.Top = 0
Picture2.Height = Picture1.ScaleHeight
Picture2.ScaleHeight = 100
Picture1.Width = vPercent
If vPercent > 0 Then
Picture1.ScaleWidth = vPercent
Picture1.Visible = True
Else
Picture1.ScaleWidth = vPercent + 1
Picture1.Visible = False
End If
Picture2.Width = 100
Dim i As Integer
Picture2.ScaleMode = 3
Picture2.DrawWidth = (Picture2.ScaleWidth \ 100) + 1
Picture2.ScaleMode = 0
Picture2.ScaleHeight = 100
Picture2.ScaleWidth = 100
Picture2.Cls
For i = 0 To Picture2.ScaleWidth Step 1
Picture2.Line (i * (Picture2.ScaleWidth / 100), 0)-(i * (Picture2.ScaleWidth / 100), Picture2.ScaleHeight), Blend(FColor1, FColor2, i)
'DrawVerticalLine i * (Picture2.ScaleWidth / 100), Blend(FColor1, FColor2, i)
Next i
UserControl.ScaleWidth = 100
lPercent.Caption = vPercent & "%"
lPercent.Left = (vPercent / 2) - (lPercent.Width / 2)
lPercent.Top = (Picture2.ScaleHeight / 2) - (lPercent.Height / 2)
End Sub

Private Sub DrawVerticalLine(X As Long, Color As Long)
Dim i As Long
i = 0
Do
DeleteObject SetPixel(Picture2.hdc, X, i, Color)
i = i + 1
Loop Until i > Picture2.ScaleHeight
End Sub

Public Function RGBRed(RGBCol As Long) As Integer
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function

Private Sub Picture1_Paint()
ReDrawBar
End Sub

Private Sub Picture2_Paint()
ReDrawBar
End Sub

Private Sub UserControl_Paint()
ReDrawBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
UserControl.Appearance = PropBag.ReadProperty("Appearance3D", 1)
UserControl.BackColor = PropBag.ReadProperty("BGC", 0)
UserControl.ForeColor = PropBag.ReadProperty("C1", 0)
UserControl.FillColor = PropBag.ReadProperty("C2", 0)
vPercent = PropBag.ReadProperty("vP", 0)
lPercent.ForeColor = PropBag.ReadProperty("lfC", 0)
lPercent.BackColor = PropBag.ReadProperty("lbC", 0)
lPercent.Visible = PropBag.ReadProperty("lV", True)
lPercent.BackStyle = PropBag.ReadProperty("lS", 0)
UserControl.BackStyle = PropBag.ReadProperty("bS", 0)
End Sub

Private Sub UserControl_Resize()
ReDrawBar
End Sub

Public Function Blend(Color1 As OLE_COLOR, Color2 As OLE_COLOR, Number As Integer) As OLE_COLOR
Dim r, g, b
r = ((RGBRed(Color1) * (100 - Number)) + (RGBRed(Color2) * (Number))) / 100
g = ((RGBGreen(Color1) * (100 - Number)) + (RGBGreen(Color2) * (Number))) / 100
b = ((RGBBlue(Color1) * (100 - Number)) + (RGBBlue(Color2) * (Number))) / 100
Blend = RGB(r, g, b)
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "BorderStyle", UserControl.BorderStyle
PropBag.WriteProperty "Appearance3D", UserControl.Appearance
PropBag.WriteProperty "BGC", UserControl.BackColor
PropBag.WriteProperty "C1", UserControl.ForeColor
PropBag.WriteProperty "C2", UserControl.FillColor
PropBag.WriteProperty "vP", vPercent
PropBag.WriteProperty "lfC", lPercent.ForeColor
PropBag.WriteProperty "lbC", lPercent.BackColor
PropBag.WriteProperty "lV", lPercent.Visible
PropBag.WriteProperty "lS", lPercent.BackStyle
PropBag.WriteProperty "bS", UserControl.BackStyle
End Sub
