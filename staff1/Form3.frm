VERSION 5.00
Begin VB.Form Department_entry 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Dapartment Entry"
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   7275
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6975
      Begin VB.CommandButton exit2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E&xit"
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton saved 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton deleted 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Delete"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton newd 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&New"
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
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox dname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enter Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Department_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cd3 As ADODB.Connection
Dim rs3 As ADODB.Recordset
Dim flagcheck As Boolean

Private Sub exit2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set cd3 = New ADODB.Connection
 cd3.Provider = "Microsoft.Jet.OLEDB.4.0"
    cd3.Open App.Path & "\db1.mdb"
'cd3.Open "PROVIDER=Microsoft.jet.OLEDB.4.0;Data Source=c:\anita\vb\db1.mdb;Persist Security Info=False "
Set rs3 = New ADODB.Recordset
rs3.Open "select * from departmentname", cd3, adOpenStatic, adLockOptimistic

Frame1.Enabled = False
Me.saved.Enabled = False
Me.deleted.Enabled = False
End Sub

Private Sub newd_Click()
Frame1.Enabled = True
flagcheck = False
Me.dname.Enabled = True
Me.dname.SetFocus
Me.saved.Enabled = True
Me.deleted.Enabled = True
dname.Clear
If (rs3.RecordCount = 0) Then
 
Else
    rs3.MoveFirst
    While (rs3.EOF = False)
        Me.dname.AddItem (rs3.Fields(0))
        rs3.MoveNext
    Wend
End If

End Sub

Private Sub saved_Click()
If (rs3.RecordCount = 0) Then

Else

    rs3.MoveFirst
        While (rs3.EOF = False)
            If (rs3.Fields(0) = Me.dname.Text) Then
                flagcheck = True
            End If
            rs3.MoveNext
        Wend
End If
If flagcheck = True Then
    MsgBox "department is already in the database"
    Frame1.Enabled = False
    newd.SetFocus
Else
    rs3.AddNew
    rs3.Fields(0) = Me.dname.Text
    rs3.Update
    MsgBox "record has been saved"
    Me.dname.Enabled = False
End If
Me.saved.Enabled = False
End Sub
Private Sub deleted_Click()
Me.dname.Enabled = True

If rs3.RecordCount > 0 Then
rs3.MoveFirst
    While (rs3.EOF = False)
        If (rs3.Fields(0) = Me.dname.Text) Then
            rs3.delete adAffectCurrent
            MsgBox "record has been deleted"
        End If
        rs3.MoveNext
      Wend
Me.dname.Clear
Me.Frame1.Enabled = False
Me.deleted.Enabled = False
End If
End Sub

