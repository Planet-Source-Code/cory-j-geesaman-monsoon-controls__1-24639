VERSION 5.00
Begin VB.UserControl mFormDragger 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   CanGetFocus     =   0   'False
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6375
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   Begin prjMonsoonControls.mGradient Gradient 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   450
      Color2          =   16711680
      Size            =   8.25
      Caption         =   "Testing..."
      Forecolor       =   14737632
   End
   Begin VB.Image btnClose 
      Height          =   150
      Left            =   4920
      Top             =   120
      Width           =   165
   End
   Begin VB.Image btnClosePressed 
      Height          =   150
      Left            =   5040
      Top             =   120
      Width           =   165
   End
End
Attribute VB_Name = "mFormDragger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'API Types


'API Constants


'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_MemberFlags = "200"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event FormDropped(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long)
Event FormMoved(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long)
Event CloseClick()
'Default Property Values:
Const m_def_RepositionForm = True
Const m_def_Caption = ""

'Property Variables:
Dim m_RepositionForm As Boolean
Dim m_Caption As String
Private Moving As Boolean

Public Sub About()
Attribute About.VB_UserMemId = -552
Load frmAbout
frmAbout.Show vbModal
End Sub

Private Sub btnClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnClose.Visible = False
UserControl_Paint
End Sub

Private Sub btnClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnClose.Visible = True
UserControl_Paint
If X >= 0 And X < 200 And Y >= 0 And Y < 200 Then
    RaiseEvent CloseClick
End If

End Sub

Private Sub Gradient_Click()
UserControl_Click
End Sub

Private Sub Gradient_DblClick()
UserControl_DblClick
End Sub

Private Sub Gradient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Gradient_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Gradient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub DragObject(ByVal hWnd As Long)

    'Procedure which simulates windows dragging of an object.
    
    Dim pt As POINTAPI
    Dim ptPrev As POINTAPI
    Dim objRect As RECT
    Dim DragRect As RECT
  '  Dim na As Long
    Dim lBorderWidth As Long
    Dim lObjWidth As Long
    Dim lObjHeight As Long
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim bMoved As Boolean
    
    ReleaseCapture
    GetWindowRect hWnd, objRect
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    'Store the initial cursor position
    ptPrev.X = pt.X
    ptPrev.Y = pt.Y
    
    'Set the initial rectangle, and draw it to show the user that
    'the object can be moved

    lXOffset = pt.X ' - objRect.left
    lYOffset = pt.Y ' - objRect.Top
    
    With DragRect
        .Left = pt.X - lObjWidth / 2
        .Top = pt.Y - 5
        .Right = .Left + lObjWidth
        .Bottom = .Top + lObjHeight
    End With
    'use form border width highlighting
    lBorderWidth = 3
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    'Move the object
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos pt
        If pt.X <> ptPrev.X Or pt.Y <> ptPrev.Y Then
            ptPrev.X = pt.X
            ptPrev.Y = pt.Y
            'erase the previous drag rectangle if any
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            'Tell the user we've moved
            RaiseEvent FormMoved(pt.X - lXOffset, pt.Y - lYOffset, lObjWidth, lObjHeight)
            'Adjust the height/width
            With DragRect
                .Left = pt.X - lObjWidth / 2 '- lXOffset
                .Top = pt.Y '- lYOffset
                .Right = .Left + lObjWidth
                .Bottom = .Top + lObjHeight
            End With
            DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
            bMoved = True
        End If
        DoEvents
    Loop
    'erase the previous drag rectangle if any
    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
    'move and repaint the window
    If bMoved Then
        If m_RepositionForm Then
            MoveWindow hWnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
        End If
        'tell the user we've dropped the form
        RaiseEvent FormDropped(DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top)
    End If
    
End Sub

Private Sub DrawDragRectangle(ByVal X As Long, ByVal Y As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal lWidth As Long)

    'Draw a rectangle using the Win32 API

    Dim hdc As Long
    Dim hPen As Long
    hPen = CreatePen(PS_SOLID, lWidth, &H808080)
    hdc = GetDC(0)
    Call SelectObject(hdc, hPen)
    Call SetROP2(hdc, R2_NOTXORPEN)
    Call Rectangle(hdc, X, Y, X1, Y1)
    Call SelectObject(hdc, GetStockObject(BLACK_PEN))
    Call DeleteObject(hPen)
    Call SelectObject(hdc, hPen)
    Call ReleaseDC(0, hdc)
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_Caption = m_def_Caption
    m_RepositionForm = m_def_RepositionForm
End Sub

Private Sub UserControl_Paint()
    
    Dim lBackColor As Long
    Dim sCaption As String
    
    'size, position, print caption etc.
    With UserControl
        .Cls
        .Extender.Align = vbAlignTop
        '.Extender.Top = 0
        '.Height = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY - 1
        .Height = Gradient.Height * Screen.TwipsPerPixelY
        Gradient.Width = UserControl.Width - Gradient.Left
        'Line1.X2 = UserControl.Width - 280
        'Line2.X2 = UserControl.Width - 280
        'Line3.X2 = UserControl.Width - 280
        'Line4.X2 = UserControl.Width - 280
        btnClose.Left = UserControl.Width - 165
        btnClosePressed.Left = btnClose.Left
        
        'draw the caption
'        If GetActiveWindow = UserControl.Extender.Parent.hwnd Then
'            .ForeColor = vbTitleBarText
            'lBackColor = vbActiveTitleBar
'            lBackColor = UserControl.BackColor
'        Else
'            .ForeColor = vbInactiveTitleBarText
            'lBackColor = vbInactiveTitleBar
'            lBackColor = UserControl.BackColor
'        End If
        
        'UserControl.Line (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - (2 * Screen.TwipsPerPixelX), UserControl.ScaleHeight - Screen.TwipsPerPixelY), lBackColor, BF
        '.CurrentX = 4 * Screen.TwipsPerPixelX
        '.CurrentY = 3 * Screen.TwipsPerPixelY
        '.Font.Name = "MS Sans Serif"
        '.Font.Bold = True
        'Check width
        'sCaption = m_Caption
        'If UserControl.TextWidth(sCaption) > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) Then
        '    Do While UserControl.TextWidth(sCaption & "...") > (UserControl.ScaleWidth - (4 * Screen.TwipsPerPixelX)) And Len(sCaption) > 0
        '        sCaption = Trim$(left$(sCaption, Len(sCaption) - 1))
        '    Loop
        '    sCaption = sCaption & "..."
        'End If
        
        'UserControl.Print sCaption;
    End With
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_RepositionForm = PropBag.ReadProperty("RepositionForm", m_def_RepositionForm)
    UserControl_Paint
End Sub

Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("RepositionForm", m_RepositionForm, m_def_RepositionForm)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets/Returns the caption of the control."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get RepositionForm() As Boolean
Attribute RepositionForm.VB_Description = "Specifies whether the control should move the form to it's new location."
    RepositionForm = m_RepositionForm
End Property

Public Property Let RepositionForm(ByVal New_RepositionForm As Boolean)
    m_RepositionForm = New_RepositionForm
    PropertyChanged "RepositionForm"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)  '  Dim na As Long
    If Button = 1 Then
    If Moving = True Then GoTo NoDice
    Moving = True
    Dim pt As POINTAPI
    Dim frmHwnd As Long
    
    UserControl_Paint
    frmHwnd = UserControl.Extender.Parent.hWnd
    
    'start 'dragging' the form
    If Button = vbLeftButton And X >= 0 And X <= UserControl.ScaleWidth And Y >= 0 And Y <= UserControl.ScaleHeight Then
       ' ReleaseCapture
        DragObject frmHwnd
    End If
    End If
NoDice:
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Moving = False
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub
