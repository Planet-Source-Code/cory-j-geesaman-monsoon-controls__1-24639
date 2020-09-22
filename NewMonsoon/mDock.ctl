VERSION 5.00
Begin VB.UserControl mDock 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3015
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   Begin prjMonsoonControls.mFormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   45
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
   End
   Begin VB.PictureBox SizeBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   3015
      TabIndex        =   4
      Top             =   3015
      Width           =   3015
   End
   Begin VB.PictureBox SizeRight 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   2925
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2715
      ScaleWidth      =   90
      TabIndex        =   3
      Top             =   300
      Width           =   90
   End
   Begin VB.PictureBox SizeLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2715
      ScaleWidth      =   90
      TabIndex        =   1
      Top             =   300
      Width           =   90
   End
   Begin VB.PictureBox SizeTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2805
      Left            =   30
      ScaleHeight     =   2805
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   270
      Width           =   750
   End
End
Attribute VB_Name = "mDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Dim mBoundHWND As Long
Dim WithEvents mBoundForm As Form
Attribute mBoundForm.VB_VarHelpID = -1
Dim mBoundFloat As New frmFloat
Public mBoundDrop As Long
Dim oldX As Long
Dim oldY As Long
Dim oldParent As Long
Dim oldStyle As Long
Dim drag As Boolean
Dim formLeft
Dim formTop
Dim formWidth
Dim formHeight
Public ParentHWND As Long

Public Sub About()
Attribute About.VB_UserMemId = -552
Load frmAbout
frmAbout.Show vbModal
End Sub

Public Sub Unload()
    Dim tmpForm As Form
    SetParent mBoundForm.hWnd, oldParent
    SetWindowLong mBoundHWND, GWL_STYLE, oldStyle
    Set tmpForm = mBoundForm
    Set mBoundFloat.mBoundControl = Nothing
    Set mBoundFloat.mBoundForm = Nothing
    Set mBoundFloat = Nothing
    Set mBoundForm = Nothing
    tmpForm.Hide
    
'    mBoundForm.Hide
'    mBoundFloat.Hide
'    Set mBoundFloat = Nothing
'    Set mBoundForm = Nothing
End Sub

Public Sub AttatchWindow(Where As Long)
    Dim xPos As Long
    Dim yPos As Long
    Dim WindowRect As RECT
    GetWindowRect mBoundFloat.hWnd, WindowRect
    xPos = WindowRect.Left * Screen.TwipsPerPixelX
    yPos = WindowRect.Top * Screen.TwipsPerPixelY
    mBoundFloat.Hide
    DoEvents
    SetParent mBoundHWND, UserControl.picClient.hWnd
    UserControl.Extender.Align = Where
    UserControl.Extender.Visible = True
    On Error Resume Next
    UserControl.Extender.Left = xPos
    UserControl.Extender.Top = yPos
    UserControl.Extender.Width = mBoundFloat.Width
    UserControl.Extender.Height = mBoundFloat.Height
    On Error GoTo 0
    FixMousePointer
    UserControl_Resize
    DoEvents
    mBoundForm.Visible = False
    DoEvents
    mBoundForm.Visible = True
    Screen.MousePointer = 0
End Sub

Public Sub Bind(Form As Object)
    Dim old
    Dim oldEX
    
    Form.Show
    
    ParentHWND = UserControl.Parent.hWnd
    
    formLeft = Form.Left
    formTop = Form.Top
    formWidth = Form.Width
    formHeight = Form.Height
    
    Set mBoundForm = Form
    mBoundHWND = Form.hWnd
    oldParent = SetParent(mBoundHWND, picClient.hWnd)
    old = GetWindowLong(mBoundHWND, GWL_STYLE)
    oldStyle = old
    oldEX = GetWindowLong(mBoundHWND, GWL_EXSTYLE)
    Set mBoundFloat.mBoundControl = Me
    SetWindowLong mBoundHWND, GWL_STYLE, old And Not WS_BORDER Or WS_CHILD Or WS_CHILDWINDOW And Not WS_CLIPCHILDREN And Not WS_CLIPSIBLINGS
    Form.Move -60, -60, UserControl.Width + 160, UserControl.Height
    Set mBoundFloat.mBoundForm = mBoundForm
    FixMousePointer
    mBoundForm.Visible = False
    DoEvents
    mBoundForm.Visible = True
End Sub

Private Sub DetatchWindow()
    Dim xPos As Long
    Dim yPos As Long
    Dim WindowRect As RECT
    GetWindowRect UserControl.hWnd, WindowRect
    xPos = WindowRect.Left * Screen.TwipsPerPixelX
    yPos = WindowRect.Top * Screen.TwipsPerPixelY
    UserControl.Extender.Visible = False
    mBoundFloat.Visible = False
    DoEvents
    mBoundFloat.Left = xPos
    mBoundFloat.Top = yPos
    mBoundFloat.Width = UserControl.Width
    mBoundFloat.Height = UserControl.Height
    SetParent mBoundHWND, mBoundFloat.picClient.hWnd
    AlwaysOnTop mBoundFloat, True
End Sub

Private Sub DrawSizerHorizontal(ByVal leftPos As Long, ByVal TopPos As Long, ByVal Height As Long)
    Dim hdc As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, 2, &H808080)
    hdc = GetDC(0)
    Call SelectObject(hdc, hPen)
    Call SetROP2(hdc, R2_NOTXORPEN)
    Call Rectangle(hdc, leftPos, TopPos - 10, leftPos + 2, TopPos + Height)
    Call SelectObject(hdc, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hdc, hPen)
    Call ReleaseDC(0, hdc)
End Sub

Private Sub DrawSizerVertical(ByVal leftPos As Long, ByVal TopPos As Long, ByVal Width As Long)
    Dim hdc As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, 2, &H808080)
    hdc = GetDC(0)
    Call SelectObject(hdc, hPen)
    Call SetROP2(hdc, R2_NOTXORPEN)
    Call Rectangle(hdc, leftPos, TopPos, leftPos + Width, TopPos + 2)
    Call SelectObject(hdc, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hdc, hPen)
    Call ReleaseDC(0, hdc)
End Sub


Private Sub FixMousePointer()
    SizeLeft.MousePointer = 0
    SizeTop.MousePointer = 0
    SizeRight.MousePointer = 0
    SizeBottom.MousePointer = 0
    Select Case UserControl.Extender.Align
        Case 1
            SizeBottom.MousePointer = 7
        Case 2
            SizeTop.MousePointer = 7
        Case 3
            SizeRight.MousePointer = 9
        Case 4
            SizeLeft.MousePointer = 9
    End Select
End Sub

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub SizeHorizontal(ByVal hWnd As Long)
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
  
    
    
    
    
    
    
    
   ' ReleaseCapture
    GetWindowRect hWnd, objRect
    GetCursorPos pt
    'Store the initial cursor position
    ptPrev.X = pt.X
    ptPrev.Y = pt.Y
    
    DrawSizerHorizontal pt.X, objRect.Top, objRect.Bottom - objRect.Top
    'Move the object
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos pt
        If pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y Then
            DrawSizerHorizontal ptPrev.X, objRect.Top, objRect.Bottom - objRect.Top
            ptPrev.X = pt.X
            ptPrev.Y = pt.Y
            DrawSizerHorizontal pt.X, objRect.Top, objRect.Bottom - objRect.Top
        End If
        DoEvents
    Loop
    'erase the previous drag rectangle if any
    DrawSizerHorizontal ptPrev.X, objRect.Top, objRect.Bottom - objRect.Top
    'move and repaint the window
    
    
End Sub



Private Sub SizeVertical(ByVal hWnd As Long)

    'Procedure which simulates windows dragging of an object.
    
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
    
   ' ReleaseCapture
    GetWindowRect hWnd, objRect
    GetCursorPos pt
    'Store the initial cursor position
    ptPrev.X = pt.X
    ptPrev.Y = pt.Y
    
    DrawSizerVertical objRect.Left, pt.Y, objRect.Right - objRect.Left
    'Move the object
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos pt
        If pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y Then
            DrawSizerVertical objRect.Left, ptPrev.Y, objRect.Right - objRect.Left
            ptPrev.X = pt.X
            ptPrev.Y = pt.Y
            DrawSizerVertical objRect.Left, pt.Y, objRect.Right - objRect.Left
        End If
        DoEvents
    Loop
    'erase the previous drag rectangle if any
    DrawSizerVertical objRect.Left, ptPrev.Y, objRect.Right - objRect.Left
    'move and repaint the window
End Sub






Private Sub cmdClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' picCaption.SetFocus
End Sub

Private Sub FormDragger1_CloseClick()
    UserControl.Extender.Visible = False
End Sub

Private Sub FormDragger1_DblClick()
    DetatchWindow
End Sub

Private Sub FormDragger1_FormDropped(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long)

    If mBoundDrop = 0 Then

        DetatchWindow
        mBoundFloat.Top = formTop * Screen.TwipsPerPixelY
        mBoundFloat.Left = formLeft * Screen.TwipsPerPixelX
        mBoundFloat.Width = formWidth * Screen.TwipsPerPixelX
        mBoundFloat.Height = formHeight * Screen.TwipsPerPixelY
        Screen.MousePointer = 0
    Else
        
        mBoundFloat.Top = formTop * Screen.TwipsPerPixelY
        mBoundFloat.Left = formLeft * Screen.TwipsPerPixelX
        mBoundFloat.Width = formWidth * Screen.TwipsPerPixelX
        mBoundFloat.Height = formHeight * Screen.TwipsPerPixelY
        AttatchWindow mBoundDrop
    End If
    
End Sub

Private Sub FormDragger1_FormMoved(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long)
    MoveForm formLeft, formTop, formWidth, formHeight, UserControl.Extender.Parent.hWnd, mBoundDrop
End Sub



Private Sub picCaption_DblClick()
    DetatchWindow
End Sub





Private Sub SizeBottom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim oPos As Long
    If UserControl.Extender.Align = 1 Then
        SizeVertical SizeBottom.hWnd
        GetCursorPos pt
        GetWindowRect UserControl.hWnd, rc
        
        'size of box
        tmp = rc.Bottom - rc.Top
      
      
        tmp = tmp + pt.Y - rc.Bottom
        tmp = tmp * Screen.TwipsPerPixelY
        oPos = UserControl.Extender.Left
        If tmp < 300 Then tmp = 300
        UserControl.Height = tmp
    End If
End Sub

Private Sub SizeLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim oPos As Long
    If UserControl.Extender.Align = 4 Then
        SizeHorizontal SizeLeft.hWnd
        GetCursorPos pt
        GetWindowRect UserControl.hWnd, rc
        tmp = rc.Right - rc.Left
        tmp = tmp + rc.Left - pt.X
        tmp = tmp * Screen.TwipsPerPixelX
        oPos = UserControl.Extender.Left
        If tmp < 300 Then tmp = 300
        UserControl.Width = tmp
    End If
End Sub


Private Sub SizeRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim oPos As Long
    If UserControl.Extender.Align = 3 Then
        SizeHorizontal SizeRight.hWnd
        GetCursorPos pt
        GetWindowRect UserControl.hWnd, rc
        tmp = rc.Right - rc.Left
        tmp = tmp + pt.X - rc.Right
        tmp = tmp * Screen.TwipsPerPixelX
        oPos = UserControl.Extender.Left
        If tmp < 300 Then tmp = 300
        UserControl.Width = tmp
    End If
End Sub



Private Sub SizeTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tmp As Long
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim oPos As Long
    If UserControl.Extender.Align = 2 Then
        SizeVertical SizeTop.hWnd
        GetCursorPos pt
        GetWindowRect UserControl.hWnd, rc
        tmp = rc.Bottom - rc.Top
        tmp = tmp + rc.Top - pt.Y
        tmp = tmp * Screen.TwipsPerPixelY
        oPos = UserControl.Extender.Left
        If tmp < 300 Then tmp = 300
        UserControl.Height = tmp
    End If
End Sub


Private Sub UserControl_Resize()
        picClient.Left = 2
        picClient.Top = 16
        picClient.Width = UserControl.Width
        picClient.Height = UserControl.Height
    If Not mBoundForm Is Nothing And UserControl.Extender.Visible = True Then
        mBoundForm.Height = UserControl.Height + 100
        mBoundForm.Width = UserControl.Width + 30
        mBoundForm.Left = -0
        mBoundForm.Top = -60
        mBoundForm.Show
    End If
    On Error Resume Next
    mBoundForm.Visible = False
    DoEvents
    mBoundForm.Visible = True
End Sub

Private Sub UserControl_Terminate()
   ' On Error Resume Next
   ' mBoundFloat.Hide
   ' mBoundFloat.Hide
   ' Set mBoundForm = Nothing
   ' Set mBoundFloat = Nothing
    
End Sub
