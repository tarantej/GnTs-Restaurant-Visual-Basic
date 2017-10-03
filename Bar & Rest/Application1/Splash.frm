VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   240
      Top             =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rising Technologies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   600
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opp : Bharat Mata Temple. Shivaji Statue Road, JALNA. (Mh). Mob : 09423156065"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   5760
      Width           =   8370
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This Product is Licenced To."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4800
      TabIndex        =   4
      Top             =   720
      Width           =   2970
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Software By."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rupali"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1350
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   3315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beer Bar && Restaurant."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2280
      Width           =   3315
   End
   Begin VB.Label lblDisp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Checking Database ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   4800
      TabIndex        =   0
      Top             =   3840
      Width           =   1965
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
        Static count As Integer
        count = count + 1
        
        If count = 1 Then
            lblDisp = "Checking Classes ..."
            
        ElseIf count = 2 Then
            lblDisp = "Classes Initialized ..."
            
        ElseIf count = 3 Then
            lblDisp = "Applying System Database ..."
            
        ElseIf count = 4 Then
            lblDisp = "Applying Bar Application ..."
            
        ElseIf count = 5 Then
            lblDisp = "Loading Database ..."
            
        ElseIf count = 6 Then
            lblDisp = "Loading Bar Application ..."
            
        ElseIf count = 7 Then
            lblDisp = "Loading ..."
            
        ElseIf count = 8 Then
            Timer1.Enabled = False
            Unload Me
            frmMain.Show
            frmWelcome.Show
        End If
End Sub
