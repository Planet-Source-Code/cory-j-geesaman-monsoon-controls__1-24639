VERSION 5.00
Begin VB.Form frmFloat 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjMonsoonControls.mFormDragger FormDragger1 
      Align           =   1  'Align Top
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   450
   End
   Begin VB.PictureBox picClient 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   60
      ScaleHeight     =   1035
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   180
      Width           =   4680
   End
End
Attribute VB_Name = "frmFloat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim drag As Boolean
Dim oldX As Long
Dim oldY As Long
Public mBoundForm As Form
Public mBoundControl As mDock
Dim DropMode As Long
Public CurrentWidth As Long
Public CurrentHeight As Long

Private Sub Form_Resize()
        picClient.Height = Me.Height - 380
        picClient.Width = Me.Width - 240

        mBoundForm.Height = Me.Height + 80 '+ 160
        mBoundForm.Width = Me.Width - 20 ' + 0 '130
        mBoundForm.Left = -80
        mBoundForm.Top = -60
        mBoundForm.Show
      '  picClient.Height = Me.Height
      Dim rec As RECT
      GetWindowRect Me.hWnd, rec

End Sub

Private Sub FormDragger1_CloseClick()
Me.Hide

End Sub

Private Sub FormDragger1_DblClick()
    mBoundControl.AttatchWindow 4
    
End Sub

Private Sub FormDragger1_FormDropped(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long)
If DropMode = 0 Then
    Me.Top = formTop * Screen.TwipsPerPixelY
    Me.Left = formLeft * Screen.TwipsPerPixelX
    Me.Width = formWidth * Screen.TwipsPerPixelX
    Me.Height = formHeight * Screen.TwipsPerPixelY
    Screen.MousePointer = 0
Else
    mBoundControl.AttatchWindow DropMode
End If
    
End Sub

Private Sub FormDragger1_FormMoved(formLeft As Long, formTop As Long, formWidth As Long, formHeight As Long)

    MoveForm formLeft, formTop, formWidth, formHeight, mBoundControl.ParentHWND, DropMode
    'Me.Left = formLeft
    'Me.Top = formTop
End Sub

Private Sub tmrKill_Timer()
    
End Sub
