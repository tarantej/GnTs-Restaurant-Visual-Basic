VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Outputs."
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   6015
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16384003
         CurrentDate     =   39044
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16384003
         CurrentDate     =   39044
      End
      Begin lvButton.lvButtons_H btnShow 
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Show"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmDates.frx":0000
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnClose 
         Height          =   495
         Left            =   4200
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Close"
         CapAlign        =   2
         BackStyle       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777215
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmDates.frx":0C55
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   3
         Height          =   1815
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Desired Period."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   2730
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To &Date"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&From Date"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Period for Selected Report."
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5040
      Picture         =   "frmDates.frx":1C8E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodic Reports."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   345
      TabIndex        =   0
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    datesSelected = False
    Unload Me
End Sub

Private Sub btnShow_Click()
    datesSelected = True
    
    dtFrom = FromDate
    dtTo = ToDate
    
    Unload Me
End Sub

Private Sub Form_Load()
    FromDate = Date
    ToDate = Date
    datesSelected = False
End Sub
