VERSION 5.00
Begin VB.Form backup_fm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Back Up"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox textfile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   1920
         Width           =   4455
      End
      Begin VB.CommandButton cmdbackup 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CREATE BACK UP!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtpath 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ENTER PATH:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ENTER FILE NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "SELECTED PATH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
   End
End
Attribute VB_Name = "backup_fm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FileSystemObject As Object
Dim filename As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Drive1.Refresh
Me.Dir1.Refresh
Me.Drive1.Refresh
Me.textfile.Text = Format$(Now, "d-mm-YYYY")
End Sub
Private Sub Dir1_Change()
Me.txtpath.Text = "" & Dir1.Path
End Sub
Private Sub Drive1_Change()
Dim d, fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Set d = fs.getdrive(fs.getdrivename(Drive1.Drive))

If d.isready Then
    Dir1.Path = Drive1.Drive
    Dir1.SetFocus
Else
    MsgBox "DRIVE IS NOT READY!!"
End If
End Sub
Private Sub cmdbackup_Click()
filename = "" + Me.txtpath.Text + "\" + Me.textfile.Text + ".mdb"
Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
FileSystemObject.copyfile App.Path & "\db1.mdb", filename
MsgBox "DATA IS SAVED "
Me.Drive1.SetFocus
End Sub

