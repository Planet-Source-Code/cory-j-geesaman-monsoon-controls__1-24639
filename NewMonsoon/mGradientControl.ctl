VERSION 5.00
Begin VB.UserControl mGradient 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1845
   ControlContainer=   -1  'True
   ScaleHeight     =   17
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ToolboxBitmap   =   "mGradientControl.ctx":0000
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   480
   End
End
Attribute VB_Name = "mGradient"
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
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Event Click()
Event DblClick()
Event mGotFocus()
Event mLostFocus()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize()

Private vCol1 As Long, vCol2 As Long, fBold As Boolean, fItalic As Boolean, fName As String, fSize As Integer, fStrikethru As Boolean, fUnderline As Boolean, fCaption As String, fForeColor As OLE_COLOR

Public Sub About()
Load frmAbout
frmAbout.Show vbModal
End Sub

Public Property Get FontBold() As Boolean
  FontBold = Label1.FontBold
End Property

Public Property Let FontBold(ByVal New_Font As Boolean)
  Label1.FontBold = New_Font
  Refresh
End Property

Public Property Get FontItalic() As Boolean
  FontItalic = Label1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_Font As Boolean)
  Label1.FontItalic = New_Font
  Refresh
End Property

Public Property Get FontName() As String
  FontName = Label1.FontName
End Property

Public Property Let FontName(ByVal New_Font As String)
  Label1.FontName = New_Font
  Refresh
End Property

Public Property Get FontSize() As Integer
  FontSize = Label1.FontSize
End Property

Public Property Let FontSize(ByVal New_Font As Integer)
  Label1.FontSize = New_Font
  Refresh
End Property

Public Property Get FontStrikethru() As Boolean
  FontStrikethru = Label1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_Font As Boolean)
  Label1.FontStrikethru = New_Font
  Refresh
End Property

Public Property Get FontUnderline() As Boolean
  FontUnderline = Label1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_Font As Boolean)
  Label1.FontUnderline = New_Font
  Refresh
End Property

Public Property Get Caption() As String
  Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_String As String)
  Label1.Caption = New_String
  Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_Color As OLE_COLOR)
  Label1.ForeColor = New_Color
  Refresh
End Property

Public Property Get Color1() As OLE_COLOR
  Color1 = vCol1
End Property

Public Property Let Color1(ByVal New_Color As OLE_COLOR)
  vCol1 = New_Color
  PropertyChanged "Color1"
  Refresh
End Property

Public Property Get Color2() As OLE_COLOR
  Color2 = vCol2
End Property

Public Property Let Color2(ByVal New_Color As OLE_COLOR)
  vCol2 = New_Color
  PropertyChanged "Color2"
  Refresh
End Property

Private Sub Label1_Click()
UserControl_Click
End Sub

Private Sub Label1_DblClick()
UserControl_DblClick
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_GotFocus()
RaiseEvent mGotFocus
End Sub

Private Sub UserControl_LostFocus()
RaiseEvent mLostFocus
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  vCol1 = PropBag.ReadProperty("Color1", vbBlack)
  vCol2 = PropBag.ReadProperty("Color2", vbBlack)
  Label1.FontBold = PropBag.ReadProperty("Bold", True)
  Label1.FontItalic = PropBag.ReadProperty("Italic", False)
  Label1.FontName = PropBag.ReadProperty("Name", "MS Sans Serif")
  Label1.FontSize = PropBag.ReadProperty("Size", 8)
  Label1.FontStrikethru = PropBag.ReadProperty("Strikethru", False)
  Label1.FontUnderline = PropBag.ReadProperty("Underline", False)
  Label1.Caption = PropBag.ReadProperty("Caption", "")
  Label1.ForeColor = PropBag.ReadProperty("Forecolor", vbBlack)
End Sub

Private Sub UserControl_Resize()
Refresh
RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Color1", vCol1, vbBlack)
  Call PropBag.WriteProperty("Color2", vCol2, vbBlack)
  Call PropBag.WriteProperty("Bold", Label1.FontBold, True)
  Call PropBag.WriteProperty("Italic", Label1.FontItalic, False)
  Call PropBag.WriteProperty("Name", Label1.FontName, "MS Sans Serif")
  Call PropBag.WriteProperty("Size", Label1.FontSize, 8)
  Call PropBag.WriteProperty("Strikethru", Label1.FontStrikethru, False)
  Call PropBag.WriteProperty("Underline", Label1.FontUnderline, False)
  Call PropBag.WriteProperty("Caption", Label1.Caption, "")
  Call PropBag.WriteProperty("Forecolor", Label1.ForeColor, vbBlack)
End Sub

Public Sub Refresh()
Dim i As Long
i = 0
Do
DrawVerticalLine i, PercentColor(((i / UserControl.ScaleWidth) * 255) \ 1)
i = i + 1
Loop Until i > UserControl.ScaleWidth
UserControl.Refresh
End Sub

Private Sub DrawVerticalLine(X As Long, Color As Long)
Dim i As Long
i = 0
Do
DeleteObject SetPixel(UserControl.hdc, X, i, Color)
i = i + 1
Loop Until i > UserControl.ScaleHeight
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

Private Function PercentColor(Percent As Long) As Long
Dim r, g, b
r = ((RGBRed(vCol1) * (255 - Percent)) + (RGBRed(vCol2) * (Percent))) / 255
g = ((RGBGreen(vCol1) * (255 - Percent)) + (RGBGreen(vCol2) * (Percent))) / 255
b = ((RGBBlue(vCol1) * (255 - Percent)) + (RGBBlue(vCol2) * (Percent))) / 255
PercentColor = RGB(r, g, b)
End Function
