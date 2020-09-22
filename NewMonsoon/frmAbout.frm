VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5460
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjMonsoonControls.mContainer mContainer 
      Height          =   4695
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8281
      UpperMargin     =   0
      LowerMargin     =   0
      BackColor       =   -2147483636
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   960
         Index           =   0
         Left            =   0
         ScaleHeight     =   930
         ScaleWidth      =   4425
         TabIndex        =   16
         Top             =   4
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   7
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":0000
            ForeColor       =   &H00E0E0E0&
            Height          =   615
            Index           =   7
            Left            =   0
            TabIndex        =   17
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1140
         Index           =   12
         Left            =   0
         ScaleHeight     =   1110
         ScaleWidth      =   4425
         TabIndex        =   14
         Top             =   5254
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   6
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":00B7
            ForeColor       =   &H00E0E0E0&
            Height          =   780
            Index           =   6
            Left            =   0
            TabIndex        =   15
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   5
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   4425
         TabIndex        =   12
         Top             =   4264
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   5
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":0184
            ForeColor       =   &H00E0E0E0&
            Height          =   585
            Index           =   5
            Left            =   0
            TabIndex        =   13
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1140
         Index           =   4
         Left            =   0
         ScaleHeight     =   1110
         ScaleWidth      =   4425
         TabIndex        =   10
         Top             =   3109
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   4
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":0225
            ForeColor       =   &H00E0E0E0&
            Height          =   780
            Index           =   4
            Left            =   0
            TabIndex        =   11
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   975
         Index           =   3
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   4425
         TabIndex        =   8
         Top             =   2119
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   3
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":02ED
            ForeColor       =   &H00E0E0E0&
            Height          =   585
            Index           =   3
            Left            =   0
            TabIndex        =   9
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1125
         Index           =   2
         Left            =   0
         ScaleHeight     =   1095
         ScaleWidth      =   4425
         TabIndex        =   6
         Top             =   979
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":038F
            ForeColor       =   &H00E0E0E0&
            Height          =   825
            Index           =   2
            Left            =   0
            TabIndex        =   7
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1140
         Index           =   1
         Left            =   0
         ScaleHeight     =   1110
         ScaleWidth      =   4425
         TabIndex        =   4
         Top             =   6409
         Width           =   4455
         Begin prjMonsoonControls.mGradient mGradient1 
            Height          =   255
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   450
            Bold            =   0   'False
            Size            =   8.25
            Caption         =   "Label1"
            Forecolor       =   -2147483630
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":043E
            ForeColor       =   &H00E0E0E0&
            Height          =   780
            Index           =   1
            Left            =   0
            TabIndex        =   5
            Top             =   270
            Width           =   4425
            WordWrap        =   -1  'True
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wB 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://C:\WINDOWS\SYSTEM\SHDOCLC.DLL/dnserror.htm#http:///"
   End
   Begin prjMonsoonControls.mFlatButton btnOk 
      Height          =   540
      Left            =   3420
      TabIndex        =   1
      Top             =   4920
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   953
   End
   Begin prjMonsoonControls.mFlatButton btnWebSite 
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   953
   End
   Begin prjMonsoonControls.mGradient mGradient 
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      Bold            =   0   'False
      Size            =   8.25
      Caption         =   "Label1"
      Forecolor       =   -2147483630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
Unload Me
End Sub

Private Sub btnWebSite_Click()
'just to annoy ppl using vb5...lol
Dim a
a = "http://www.naaveen.net/"
wB.Navigate Replace(Replace(a, "aa", "a"), "ee", "e"), -1
End Sub

Private Sub Form_Load()
If App.Minor < 10 And App.Revision < 10 Then
mGradient.Caption = "About The Monsoon Controls v" & App.Major & "." & App.Minor & App.Revision
Else
mGradient.Caption = "About The Monsoon Controls v" & App.Major & "." & App.Minor & "." & App.Revision
End If
btnOk.Init
btnWebSite.Init
Me.Width = 4830
Me.Height = 5580
End Sub

Private Sub Form_Resize()
mGradient.Width = Me.ScaleWidth
mContainer.Width = Me.ScaleWidth
End Sub

Private Sub Label1_Click(Index As Integer)
wB.Navigate "mailto:martin.olivier@bigfoot.com", -1
End Sub

Private Sub mGradient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragForm Me
End Sub
