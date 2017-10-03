VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Employee_detail_report 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Employee Detail"
   ClientHeight    =   4515
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   12640511
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame e_report_fr 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton exit1 
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton print1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Print"
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
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combodname 
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
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   3615
      End
      Begin VB.CommandButton search 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sea&rch"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1455
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Employee_detail_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ds4 As Connection
Dim rs4 As Recordset
Dim rs4a As Recordset



Private Sub Form_Load()

Set ds4 = New Connection
    ds4.Provider = "Microsoft.Jet.OLEDB.4.0"
    ds4.Open App.Path & "\db1.mdb"
   
Set rs4a = New Recordset
    rs4a.Open "select * from departmentname", ds4, adOpenStatic, adLockOptimistic
    rs4a.MoveFirst
    While (Not rs4a.EOF)
        Me.Combodname.AddItem (rs4a.Fields(0))
        rs4a.MoveNext
    Wend
Me.print1.Enabled = False
End Sub

Private Sub search_Click()
Set rs4 = New Recordset
rs4.Open "select dname,ename,address,city,pincode,phno,salary,format$(jdate,'d/mmm/yyyy')AS jdate,edu,remark from employeeentry where dname='" & Combodname.Text & "' ", ds4, adOpenStatic, adLockOptimistic

    If rs4.EOF Then
        MsgBox ("NO RECORDS FOUND FOR SELECTED DATE!")
    Else
   
        Set MSHFlexGrid1.DataSource = rs4
        Me.print1.Enabled = True
    End If
End Sub
Private Sub print1_Click()
If DataEnvironment1.rsCommand1.State = adStateOpen Then
    DataEnvironment1.rsCommand1.Close
End If
DataEnvironment1.rsCommand1.Open "Select dname,ename,address,city,pincode,phno,salary,format$(jdate,'d/mmm/yyyy') as jdate,edu,remark from employeeentry where dname= '" & Combodname.Text & "'", ds4, adOpenStatic, adLockOptimistic
employee_detail.Refresh
employee_detail.Show
DataEnvironment1.rsCommand1.Close
        Me.print1.Enabled = False
End Sub
Private Sub exit1_Click()
Unload Me
End Sub

 
