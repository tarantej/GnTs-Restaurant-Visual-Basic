VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form pay_slip_report 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Monthly Pay Report"
   ClientHeight    =   8085
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   6135
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   10821
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "----------------------------------------------------------Pay Report--------------------------------------------------------"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.ComboBox L_year 
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
         Left            =   1920
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmd_print 
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
         Height          =   315
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmd_search 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox L_month 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   3015
      End
      Begin VB.ComboBox L_dname 
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
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Year:"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Month:"
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
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Depart.:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "pay_slip_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As ADODB.Connection
Dim r As ADODB.Recordset



Private Sub Form_Load()

Set d = New ADODB.Connection
    d.Provider = "Microsoft.Jet.OLEDB.4.0"
    d.Open App.Path & "\db1.mdb"

Set r = New ADODB.Recordset
    r.Open "select * from departmentname ", d, adOpenStatic, adLockOptimistic
        count1 = 0
        r.Requery
        Me.L_dname.Clear
        While (count1 < r.RecordCount)
            Me.L_dname.AddItem (r.Fields(0))
            count1 = count1 + 1
            r.MoveNext
        Wend
        

Me.L_month.AddItem ("JANUARY")
Me.L_month.AddItem ("FEBRUARY")
Me.L_month.AddItem ("MARCH")
Me.L_month.AddItem ("APRIL")
Me.L_month.AddItem ("MAY")
Me.L_month.AddItem ("JUN")
Me.L_month.AddItem ("JULY")
Me.L_month.AddItem ("AUGUST")
Me.L_month.AddItem ("SEPTEMBER")
Me.L_month.AddItem ("OCTOBER")
Me.L_month.AddItem ("NOVEMBER")
Me.L_month.AddItem ("DECEMBER")

Me.L_year.AddItem ("2005")
Me.L_year.AddItem ("2006")
Me.L_year.AddItem ("2007")
Me.L_year.AddItem ("2008")
Me.L_year.AddItem ("2009")
Me.L_year.AddItem ("2010")

Me.cmd_print.Enabled = False

End Sub

Private Sub cmd_search_Click()

Set r = New ADODB.Recordset
    r.Open "select Emp_name as 'Employee Name',Total_wday,Total_salary from pay_slip where Dept_name='" & L_dname.Text & "' and Month1='" & L_month.Text & "' AND Year1='" & Me.L_year.Text & "'", d, adOpenStatic, adLockOptimistic
    
    Me.MSHFlexGrid1.ColWidth(0) = 6000
    Me.MSHFlexGrid1.ColWidth(1) = 1500
    Me.MSHFlexGrid1.ColWidth(2) = 2000
    Me.MSHFlexGrid1.Font.Size = 10
    Me.MSHFlexGrid1.AllowUserResizing = flexResizeColumns
    Me.MSHFlexGrid1.Font.Bold = True
    
    If r.RecordCount = 0 Then
        MsgBox "No record!!!"
    Else
        Set Me.MSHFlexGrid1.DataSource = r
    End If
    
   Me.cmd_print.Enabled = True
    
End Sub
Private Sub cmd_print_Click()
DataEnvironment1.rsCommand3.Open "select Emp_name,Month1,Year1,Total_wday,Total_salary from pay_slip where Dept_name='" & L_dname.Text & "' and Month1='" & L_month.Text & "' AND Year1='" & Me.L_year.Text & "'", d, adOpenStatic, adLockOptimistic
    monthly_pay.Refresh
    monthly_pay.Show
    DataEnvironment1.rsCommand3.Close
End Sub
Private Sub cmd_exit_click()
Unload Me
End Sub

