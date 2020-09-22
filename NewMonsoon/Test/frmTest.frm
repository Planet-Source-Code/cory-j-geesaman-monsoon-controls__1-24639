VERSION 5.00
Object = "*\A..\prjMonsoonControls.vbp"
Begin VB.Form frmTest 
   Caption         =   "Monsoon Controls"
   ClientHeight    =   2655
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   3  'Windows Default
   Begin prjMonsoonControls.mContainer mContainer 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
      UpperMargin     =   120
      LowerMargin     =   120
      BackColor       =   -2147483633
      BorderStyle     =   1
      Begin prjMonsoonControls.mHScroll mHScroll 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Max             =   "10"
         ForeColor       =   -2147483634
         BackColor       =   -2147483628
         BarCol1         =   0
         BarCol2         =   12582912
         Enabled         =   0   'False
      End
      Begin prjMonsoonControls.mSystemButton mSystemButton4 
         Height          =   255
         Left            =   2220
         TabIndex        =   7
         Top             =   4738
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Button          =   3
      End
      Begin prjMonsoonControls.mSystemButton mSystemButton3 
         Height          =   255
         Left            =   1940
         TabIndex        =   6
         Top             =   4458
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
      End
      Begin prjMonsoonControls.mSystemButton mSystemButton2 
         Height          =   255
         Left            =   2500
         TabIndex        =   5
         Top             =   4458
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Button          =   2
      End
      Begin prjMonsoonControls.mSystemButton mSystemButton1 
         Height          =   255
         Left            =   2220
         TabIndex        =   4
         Top             =   4178
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Button          =   1
      End
      Begin prjMonsoonControls.mFlatButton mFlatButton 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   3098
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         Picture         =   "frmTest.frx":0000
      End
      Begin prjMonsoonControls.mToggleButton mToggleButton 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2378
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         LC              =   "mToggleButton"
         BU              =   -1  'True
         bP              =   "frmTest.frx":0452
         bC              =   -2147483633
         fC              =   8388608
      End
      Begin prjMonsoonControls.mProgress mProgress 
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1658
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         BorderStyle     =   1
         Appearance3D    =   0
         BGC             =   -2147483633
         C1              =   8388608
         C2              =   0
         vP              =   10
         lfC             =   14737632
         lbC             =   -2147483633
         lV              =   -1  'True
         lS              =   0
         bS              =   1
      End
      Begin prjMonsoonControls.mGradient mGradient 
         Height          =   255
         Left            =   120
         Top             =   218
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color2          =   8388608
         Size            =   8.25
         Caption         =   "mGradient"
         Forecolor       =   12632256
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
mHScroll.DrawBar
mProgress.ReDrawBar
mFlatButton.Init
End Sub

Private Sub mHScroll_Change(Value As Long, Max As Long)
mProgress.Percent = (Value / Max) * 100
End Sub

Private Sub mnuHelpAbout_Click()
mContainer.About
End Sub
