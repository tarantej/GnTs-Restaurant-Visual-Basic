VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      Picture         =   "Welcome.frx":0000
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   1440
         TabIndex        =   2
         Top             =   1440
         Width           =   105
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   1455
         TabIndex        =   1
         Top             =   960
         Width           =   105
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3000
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   120
      Top             =   1920
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    lblTime = "Login Time : " & Time
    lblDate = Format(Date, "dd-MMM-yyyy")
    
    Call popUp
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width - (Me.Width + 60)
    Me.Top = Screen.Height - 600 'assumed height for taskbar
End Sub

Private Sub popUp()
    Dim h As Integer
    
    h = Me.Height
    Me.Height = 0
    
    While Me.Height < h
        Me.Height = Me.Height + 1
        Me.Top = Me.Top - 1
        DoEvents
    Wend
End Sub

Private Sub popDown()
    On Error Resume Next
    
    While Me.Height > 0
        Me.Height = Me.Height - 1
        Me.Top = Me.Top + 1
        DoEvents
    Wend
    Unload Me
End Sub

Private Sub Timer1_Timer()
    popDown
End Sub
