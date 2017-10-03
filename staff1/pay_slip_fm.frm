VERSION 5.00
Begin VB.Form pay_slip_fm 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "PAY SLIP"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   6735
      Begin VB.CommandButton cmd_exit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E&XIT"
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_print 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&PRINT"
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
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&SAVE"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmd_add 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&NEW"
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "---------------------------Pay SLip---------------------------------------------"
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
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
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
         Left            =   5280
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox L_month 
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
         IntegralHeight  =   0   'False
         ItemData        =   "pay_slip_fm.frx":0000
         Left            =   2520
         List            =   "pay_slip_fm.frx":0002
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox T_salary 
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   3240
         TabIndex        =   9
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox T_wday 
         BorderStyle     =   0  'None
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
         Left            =   3240
         TabIndex        =   8
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox T_holiday 
         BorderStyle     =   0  'None
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
         Left            =   3240
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
      End
      Begin VB.ListBox L_ename 
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
         ItemData        =   "pay_slip_fm.frx":0004
         Left            =   2520
         List            =   "pay_slip_fm.frx":0006
         TabIndex        =   2
         Top             =   840
         Width           =   3855
      End
      Begin VB.ListBox L_dname 
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
         ItemData        =   "pay_slip_fm.frx":0008
         Left            =   2520
         List            =   "pay_slip_fm.frx":000A
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox T_prday 
         BorderStyle     =   0  'None
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
         Left            =   3240
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox T_abday 
         BorderStyle     =   0  'None
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
         Left            =   3240
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enter month and  year:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Salary for this month"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total Working day in month"
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
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   3135
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Holiday"
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
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Absent day:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Present day:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "pay_slip_fm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cd_pay As ADODB.Connection
Dim rs_pay As ADODB.Recordset
Dim rs_pay_save As ADODB.Recordset

Dim rsd As ADODB.Recordset
Dim rse As ADODB.Recordset
Dim count1 As Integer
Dim pr_day, ab_day, holy_day As Integer
Private Sub Form_Load()
Set cd_pay = New ADODB.Connection
    cd_pay.Provider = "Microsoft.Jet.OLEDB.4.0"
    cd_pay.Open App.Path & "\db1.mdb"

    Me.Frame1.Enabled = False
    Me.cmd_save.Enabled = False
End Sub
Private Sub cmd_add_Click()
Me.Frame1.Enabled = True
Me.cmd_save.Enabled = True

Set rsd = New ADODB.Recordset
    rsd.Open "select * from departmentname ", cd_pay, adOpenStatic, adLockOptimistic
        count1 = 0
        rsd.Requery
        Me.L_dname.Clear
    While (count1 < rsd.RecordCount)
        Me.L_dname.AddItem (rsd.Fields(0))
        count1 = count1 + 1
        rsd.MoveNext
    Wend
    Me.L_dname.SetFocus
End Sub
Private Sub L_dname_click()
Set rse = New ADODB.Recordset
    rse.Open "select * from employeeentry ", cd_pay, adOpenStatic, adLockOptimistic
        count1 = 0
        rse.Requery
        Me.L_ename.Clear
        While (count1 < rse.RecordCount)
            If rse.Fields("dname") = Me.L_dname.Text Then
                 Me.L_ename.AddItem (rse.Fields(2))
            End If
                count1 = count1 + 1
                rse.MoveNext
        Wend
    Me.L_ename.SetFocus
End Sub
Private Sub L_ename_click()

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

Me.L_month.ListIndex = 0

Me.L_year.AddItem ("2005")
Me.L_year.AddItem ("2006")
Me.L_year.AddItem ("2007")
Me.L_year.AddItem ("2008")
Me.L_year.AddItem ("2009")
Me.L_year.AddItem ("2010")

Me.L_year.ListIndex = 0
End Sub
Private Sub L_month_click()
Set rs_pay = New ADODB.Recordset
rs_pay.Open "select ename,date1,date2,atype from attendenceentry where  ename='" & L_ename.Text & " ' ", cd_pay, adOpenStatic, adLockOptimistic
        pr_day = 0
        ab_day = 0
        holy_day = 0
        count1 = 0
        rs_pay.Requery
        While (count1 < rs_pay.RecordCount)
            If Month(rs_pay.Fields("date2")) = (Me.L_month.ListIndex + 1) And Year(rs_pay.Fields("date2")) = (Me.L_year.Text) Then
                        
                        If (rs_pay.Fields("atype")) = "P" Then
                            pr_day = pr_day + 1
                        ElseIf (rs_pay.Fields("atype")) = "AB" Then
                            ab_day = ab_day + 1
                        ElseIf (rs_pay.Fields("atype")) = "HOLYDAY" Then
                            holy_day = holy_day + 1
                        End If
            End If
            count1 = count1 + 1
            rs_pay.MoveNext
        Wend
                
        Me.T_prday.Text = pr_day
        Me.T_abday.Text = ab_day
        Me.T_holiday.Text = holy_day
        Me.T_wday.Text = pr_day + holy_day
        
    Set rse = New ADODB.Recordset
    rse.Open "select * from employeeentry where ename='" & L_ename.Text & "'", cd_pay, adOpenStatic, adLockOptimistic
        Me.T_salary.Text = (rse.Fields("salary") / 30) * (Me.T_wday.Text)
        Me.T_salary.Text = Format(Me.T_salary.Text, "0.00")
End Sub
Private Sub L_year_click()
Set rs_pay = New ADODB.Recordset
rs_pay.Open "select ename,date1,date2,atype from attendenceentry where  ename='" & L_ename.Text & " ' ", cd_pay, adOpenStatic, adLockOptimistic
        pr_day = 0
        ab_day = 0
        holy_day = 0

        count1 = 0
        rs_pay.Requery
        While (count1 < rs_pay.RecordCount)
            If Month(rs_pay.Fields("date2")) = (Me.L_month.ListIndex + 1) And Year(rs_pay.Fields("date2")) = (Me.L_year.Text) Then
                        If (rs_pay.Fields("atype")) = "P" Then
                            pr_day = pr_day + 1
                        ElseIf (rs_pay.Fields("atype")) = "AB" Then
                            ab_day = ab_day + 1
                        ElseIf (rs_pay.Fields("atype")) = "HOLYDAY" Then
                            holy_day = holy_day + 1
                        End If
            End If
            count1 = count1 + 1
            rs_pay.MoveNext
        Wend
        
        Me.T_prday.Text = pr_day
        Me.T_abday.Text = ab_day
        Me.T_holiday.Text = holy_day
        Me.T_wday.Text = pr_day + holy_day
        
    Set rse = New ADODB.Recordset
        rse.Open "select * from employeeentry where ename='" & L_ename.Text & "'", cd_pay, adOpenStatic, adLockOptimistic
        Me.T_salary.Text = (rse.Fields("salary") / 30) * (Me.T_wday.Text)
End Sub
Private Sub cmd_save_Click()
    'If Me.L_dname.Text = "" Or Me.L_ename.Text = "" Or Me.L_month.Text = "" Then
    '    MsgBox "ENTER ENOUGH INFO"
    '    Me.L_dname.SetFocus
    '    Me.cmd_save.Enabled = False
    '    Exit Sub
    'End If

  Set rs_pay_save = New ADODB.Recordset
      rs_pay_save.Open "select * from pay_slip WHERE Dept_name='" & L_dname.Text & "' and Emp_name='" & L_ename.Text & "' AND Month1='" & L_month.Text & "' AND Year1='" & Me.L_year.Text & "'", cd_pay, adOpenStatic, adLockOptimistic
     If rs_pay_save.RecordCount = 0 Then
          
        rs_pay_save.AddNew
        
        rs_pay_save.Fields(0) = Me.L_dname.Text
        rs_pay_save.Fields(1) = Me.L_ename.Text
        rs_pay_save.Fields(2) = Me.L_month.Text
        rs_pay_save.Fields(3) = Me.T_prday.Text
        rs_pay_save.Fields(4) = Me.T_abday.Text
        rs_pay_save.Fields(5) = Me.T_holiday.Text
        rs_pay_save.Fields(6) = Me.T_wday.Text
        rs_pay_save.Fields(7) = Me.T_salary.Text
        rs_pay_save.Fields(8) = Me.L_year.Text
        
        rs_pay_save.Update
        MsgBox "RECORD SAVED!!"
     Else
        MsgBox "RECORD IS ALREADY SAVED!!"
     End If
End Sub
Private Sub cmd_exit_click()
Unload Me
End Sub
Private Sub cmd_print_Click()
    DataEnvironment1.rsCommand3.Open "Select * from pay_slip where Emp_name= '" & Me.L_ename.Text & "' AND Dept_name='" & Me.L_dname.Text & "' AND Month1='" & Me.L_month.Text & "' AND year1='" & Me.L_year.Text & "' ", cd_pay, adOpenStatic, adLockOptimistic
    pay_slip.Refresh
    pay_slip.Show
    DataEnvironment1.rsCommand3.Close
End Sub
Private Sub cleartext()
   Me.L_dname.Text = ""
   Me.L_ename.Text = ""
   Me.L_month.Text = ""
   Me.T_prday.Text = ""
   Me.T_abday.Text = ""
   Me.T_holiday.Text = ""
   Me.T_wday.Text = ""
   Me.T_salary.Text = ""
End Sub


