VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Attendence_report 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Attendence Report"
   ClientHeight    =   8700
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame att_dreport_fr 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   7575
      Begin VB.CommandButton search2 
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
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
         Width           =   1575
      End
      Begin VB.CommandButton print2 
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
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16744576
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Daily"
      TabPicture(0)   =   "Attendence_report.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DTPicker1"
      Tab(0).Control(1)=   "Label1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Monthly"
      TabPicture(1)   =   "Attendence_report.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Combomonth"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Comboyear"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "List2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "List1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Attendence_report.frx":0038
         Left            =   2160
         List            =   "Attendence_report.frx":003A
         TabIndex        =   3
         Top             =   1080
         Width           =   4695
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "Attendence_report.frx":003C
         Left            =   2160
         List            =   "Attendence_report.frx":003E
         TabIndex        =   4
         Top             =   1560
         Width           =   4695
      End
      Begin VB.ComboBox Comboyear 
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
         Left            =   5160
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox Combomonth 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   10
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "d/M/yyyy"
         Format          =   19267587
         CurrentDate     =   38600
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "Department Name:"
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
         TabIndex        =   15
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000016&
         Caption         =   "Employee Name:"
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
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
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
         Left            =   4080
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Month:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Date:"
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
         Left            =   -74280
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9340
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   12640511
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
End
Attribute VB_Name = "Attendence_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ds5 As Connection
Dim rs5 As Recordset
Dim rs As Recordset
Dim rse As Recordset
Private Sub Form_Load()
Set ds5 = New Connection
    ds5.Provider = "Microsoft.Jet.OLEDB.4.0"
    ds5.Open App.Path & "\db1.mdb"
Me.print2.Enabled = False
Me.Combomonth.AddItem ("JANUARY")
Me.Combomonth.AddItem ("FEBRUARY")
Me.Combomonth.AddItem ("MARCH")
Me.Combomonth.AddItem ("APRIL")
Me.Combomonth.AddItem ("MAY")
Me.Combomonth.AddItem ("JUNE")
Me.Combomonth.AddItem ("JULY")
Me.Combomonth.AddItem ("AUGUST")
Me.Combomonth.AddItem ("SUPTEMBER")
Me.Combomonth.AddItem ("OCTOBER")
Me.Combomonth.AddItem ("NOVEMBER")
Me.Combomonth.AddItem ("DECEMBER")

Me.Comboyear.AddItem ("2005")
Me.Comboyear.AddItem ("2006")
Me.Comboyear.AddItem ("2007")
Me.Comboyear.AddItem ("2008")
Me.Comboyear.AddItem ("2009")

Set rs = New ADODB.Recordset
rs.Open "select * from departmentname", ds5, adOpenStatic, adLockOptimistic

While (rs.EOF = False)
            Me.List1.AddItem (rs.Fields("dname"))
            rs.MoveNext
Wend


End Sub




Private Sub List1_Click()
Set rse = New ADODB.Recordset
rse.Open "select * from employeeentry", ds5, adOpenStatic, adLockOptimistic
Me.List2.Clear
        rse.MoveFirst
            While (rse.EOF = False)
                If (rse.Fields(0) = List1.Text) Then
                    Me.List2.AddItem (rse.Fields(2))
                End If
                rse.MoveNext
            Wend
End Sub

Private Sub search2_Click()
Set rs5 = New Recordset
If SSTab1.Caption = "Daily" Then
        rs5.Open "select dname,ename,date1,atype,avalue,dstart,dend,lstart,lend,totalhour from attendenceentry where date1 = '" & Format$(DTPicker1.Value, "d/M/yyyy") & "' order by date1", ds5, adOpenStatic, adLockOptimistic
        
        If rs5.EOF Then
            MsgBox ("NO RECORDS FOUND FOR SELECTED DATE!")
        Else
  
        Set MSHFlexGrid2.DataSource = rs5
        Me.print2.Enabled = True
        End If
ElseIf SSTab1.Caption = "Monthly" Then
        m = Me.Combomonth.ListIndex + 1
        rs5.Open "select dname,ename,date1,atype,avalue,dstart,dend,lstart,lend,totalhour from attendenceentry where month(date2) = '" & m & "' and year(date2) = '" & Comboyear.Text & "'And dname = '" & List1.Text & "' And ename = '" & List2.Text & "' order by date2", ds5, adOpenStatic, adLockOptimistic
        If rs5.EOF Then
            MsgBox ("NO RECORDS FOUND FOR SELECTED DATE!")
        Else
        Set MSHFlexGrid2.DataSource = rs5
        Me.print2.Enabled = True
        End If
End If
End Sub

Private Sub print2_Click()
If SSTab1.Caption = "Daily" Then
    DataEnvironment1.rsCommand2.Open "Select dname,ename,date1,atype,avalue,dstart,dend,lstart,lend,totalhour from attendenceentry WHERE date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' order by date1", ds5, adOpenStatic, adLockOptimistic
    While (Not DataEnvironment1.rsCommand2.EOF)
        If DataEnvironment1.rsCommand2.Fields("dstart") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("dstart") = Null
        End If
        If DataEnvironment1.rsCommand2.Fields("lstart") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("lstart") = Null
        End If
        If DataEnvironment1.rsCommand2.Fields("lend") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("lend") = Null
        End If
        If DataEnvironment1.rsCommand2.Fields("dend") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("dend") = Null
        End If
        DataEnvironment1.rsCommand2.MoveNext
    Wend
    attendence_detail.Refresh
    attendence_detail.Show
    attendence_detail.Refresh
    attendence_detail.Refresh
    
    DataEnvironment1.rsCommand2.Close
     
     Dim rs1 As ADODB.Recordset
     Set rs1 = New ADODB.Recordset
         rs1.Open "Select dname,ename,date1,atype,avalue,dstart,dend,lstart,lend,totalhour from attendenceentry WHERE date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' order by date2 ", ds5, adOpenStatic, adLockOptimistic
         While (Not rs1.EOF)
         If IsNull(rs1.Fields("dstart")) Then
            rs1.Fields("dstart") = "12:00:00 AM"
         End If
         If IsNull(rs1.Fields("lstart")) Then
            rs1.Fields("lstart") = "12:00:00 AM"
         End If
         If IsNull(rs1.Fields("lend")) Then
            rs1.Fields("lend") = "12:00:00 AM"
         End If
         If IsNull(rs1.Fields("dend")) Then
            rs1.Fields("dend") = "12:00:00 AM"
         End If
         rs1.MoveNext
    Wend
     
ElseIf SSTab1.Caption = "Monthly" Then ' for monthly report
    m = Me.Combomonth.ListIndex + 1
    DataEnvironment1.rsCommand2.Open "select dname,eno,ename,date1,date2,atype,avalue,dstart,dend,lstart,lend,totalhour from attendenceentry where month(date2) = '" & m & "' and year(date2) = '" & Comboyear.Text & "'And dname = '" & List1.Text & "' And ename = '" & List2.Text & "' order by date2", ds5, adOpenStatic, adLockOptimistic
    While (Not DataEnvironment1.rsCommand2.EOF)
        If DataEnvironment1.rsCommand2.Fields("dstart") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("dstart") = Null
        End If
        If DataEnvironment1.rsCommand2.Fields("lstart") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("lstart") = Null
        End If
        If DataEnvironment1.rsCommand2.Fields("lend") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("lend") = Null
        End If
        If DataEnvironment1.rsCommand2.Fields("dend") = "12:00:00 AM" Then
            DataEnvironment1.rsCommand2.Fields("dend") = Null
        End If
        DataEnvironment1.rsCommand2.MoveNext
    Wend
    attendence_detail.Refresh
    attendence_detail.Refresh
    attendence_detail.Show
    attendence_detail.Refresh
    attendence_detail.Refresh
   
DataEnvironment1.rsCommand2.Close
  
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
      rs.Open "select * from attendenceentry where month(date2) = '" & m & "' and year(date2) = '" & Comboyear.Text & "'And dname = '" & List1.Text & "' And ename = '" & List2.Text & "'", ds5, adOpenStatic, adLockOptimistic
    While (Not rs.EOF)
        If IsNull(rs.Fields("dstart")) Then
            rs.Fields("dstart") = "12:00:00 AM"
        End If
        If IsNull(rs.Fields("lstart")) Then
            rs.Fields("lstart") = "12:00:00 AM"
        End If
        If IsNull(rs.Fields("lend")) Then
            rs.Fields("lend") = "12:00:00 AM"
        End If
        If IsNull(rs.Fields("dend")) Then
            rs.Fields("dend") = "12:00:00 AM"
        End If
        rs.MoveNext
    Wend
     
End If
        Me.print2.Enabled = False
End Sub
Private Sub exit2_Click()
    Unload Me
End Sub



