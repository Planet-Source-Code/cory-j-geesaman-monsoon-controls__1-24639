VERSION 5.00
Begin VB.Form frmHeaderSafe 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1335
      Index           =   6
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   4425
      TabIndex        =   19
      Top             =   4590
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   9
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mFormDragger"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":0000
         ForeColor       =   &H00E0E0E0&
         Height          =   1020
         Index           =   9
         Left            =   0
         TabIndex        =   20
         Top             =   270
         Width           =   4425
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1140
      Index           =   0
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   4425
      TabIndex        =   17
      Top             =   2295
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   10
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mDock"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":011D
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   10
         Left            =   0
         TabIndex        =   18
         Top             =   270
         Width           =   4425
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1305
      Index           =   0
      Left            =   0
      ScaleHeight     =   1275
      ScaleWidth      =   4425
      TabIndex        =   14
      Top             =   975
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   11
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mCoolMenu"
         Forecolor       =   14737632
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Olivier Martin"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   2610
         MouseIcon       =   "frmHeaderSafe.frx":01F5
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":0347
         ForeColor       =   &H00E0E0E0&
         Height          =   975
         Index           =   11
         Left            =   0
         TabIndex        =   15
         Top             =   270
         Width           =   4425
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   4425
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   12
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mContainer"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":044C
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Index           =   12
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
      Index           =   7
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   4425
      TabIndex        =   10
      Top             =   9075
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   13
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mSystemButton"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":0503
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   13
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
      Index           =   8
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   4425
      TabIndex        =   8
      Top             =   8085
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   14
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mProgress"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":05D0
         ForeColor       =   &H00E0E0E0&
         Height          =   585
         Index           =   14
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
      Height          =   1140
      Index           =   9
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   4425
      TabIndex        =   6
      Top             =   6930
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   15
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mHScroll"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":0671
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   15
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
      Height          =   975
      Index           =   10
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   4425
      TabIndex        =   4
      Top             =   5940
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   16
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mGradient"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":0739
         ForeColor       =   &H00E0E0E0&
         Height          =   585
         Index           =   16
         Left            =   0
         TabIndex        =   5
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
      Index           =   11
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   3450
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   17
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mFlatButton"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":07DB
         ForeColor       =   &H00E0E0E0&
         Height          =   825
         Index           =   17
         Left            =   0
         TabIndex        =   3
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
      Index           =   0
      Left            =   0
      ScaleHeight     =   1110
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   10230
      Visible         =   0   'False
      Width           =   4455
      Begin prjMonsoonControls.mGradient mGradient1 
         Height          =   255
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         Color1          =   8388608
         Color2          =   8421504
         Size            =   8.25
         Caption         =   "mToggleButton"
         Forecolor       =   14737632
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmHeaderSafe.frx":088A
         ForeColor       =   &H00E0E0E0&
         Height          =   780
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   270
         Width           =   4425
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmHeaderSafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
