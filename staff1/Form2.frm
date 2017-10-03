VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Attendence_entry 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Attendence"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5340
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   6495
      Begin VB.CommandButton exit1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E&xit"
         Height          =   375
         Left            =   2400
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton search 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sea&rch"
         Height          =   375
         Left            =   4800
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton edit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3600
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
         Height          =   375
         Left            =   1200
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton add 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Add"
         Height          =   375
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame aentry 
      BackColor       =   &H00C0E0FF&
      Caption         =   "-------------------------------Attendence Entry-------------------------"
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.ComboBox ename 
         Height          =   360
         Left            =   2400
         TabIndex        =   12
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox totalhour 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   20
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox dstart 
         Height          =   345
         Left            =   1440
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox atype 
         Height          =   360
         Left            =   5040
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox avalue 
         Height          =   360
         Left            =   5040
         TabIndex        =   19
         Top             =   2400
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "d/M/yyyy"
         Format          =   19267587
         CurrentDate     =   38593
      End
      Begin VB.ComboBox dname 
         Height          =   360
         ItemData        =   "Form2.frx":0000
         Left            =   2400
         List            =   "Form2.frx":0002
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox lend 
         Height          =   360
         Left            =   1440
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox lstart 
         Height          =   360
         Left            =   1440
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox dend 
         Height          =   345
         Left            =   1440
         TabIndex        =   17
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------------------------------------------------------------------------"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3120
         Width           =   6015
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------------------------------------------------------------------------"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1440
         Width           =   6015
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total Hours of Work:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Attendence Type:"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2880
         TabIndex        =   30
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Department Name:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Attandence Value:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lunch End:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lunch Start:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Day End:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Date:"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Day Start:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Name:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame searchframe 
      BackColor       =   &H00C0E0FF&
      Caption         =   "--------------------------------Search-------------------------------------"
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   120
      TabIndex        =   34
      Top             =   240
      Width           =   6495
      Begin VB.ComboBox cmbdname 
         Height          =   360
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox cmbename 
         Height          =   360
         Left            =   2280
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.TextBox txttotalhour 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   330
         Left            =   3240
         TabIndex        =   10
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtavalue 
         Enabled         =   0   'False
         Height          =   360
         Left            =   4320
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtatype 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtlstart 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1560
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtdend 
         Enabled         =   0   'False
         Height          =   360
         Left            =   4320
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtdstart 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1560
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtlend 
         Enabled         =   0   'False
         Height          =   360
         Left            =   4320
         TabIndex        =   9
         Top             =   2640
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "d/M/yyyy"
         Format          =   19267587
         CurrentDate     =   38593
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "-------------------------------------------------------------------------------"
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   3120
         Width           =   6015
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "--------------------------------------------------------------------------------"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         Width           =   6015
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Total Hours of Work:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   45
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Day Start:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Day End:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   43
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lunch Start:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Lunch End:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   41
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Att. Value:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Att. Type:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Date:"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Employee Name:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Department Name:"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Attendence_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cd2 As ADODB.Connection
Dim rs2 As ADODB.Recordset

Dim rsd As ADODB.Recordset
Dim rse, rsf, rsg, rsh, rsi, rsj As ADODB.Recordset

Dim flagsave As Boolean
Dim flagedit As Boolean
Dim flagsearch As Boolean


Dim lunchtime, daytime, totaltime, thour, totalminute As Integer
Dim la, da As Date
Dim lb, db As Date
Dim strlunchtime As String

Dim count1 As Integer



Private Sub exit1_Click()
Unload Me
End Sub
Private Sub Form_Load()
Set cd2 = New ADODB.Connection
    cd2.Provider = "Microsoft.Jet.OLEDB.4.0"
    cd2.Open App.Path & "\db1.mdb"
Set rs2 = New ADODB.Recordset
rs2.Open "select * from attendenceentry", cd2, adOpenStatic, adLockOptimistic

Me.aentry.Enabled = False
Me.searchframe.Visible = False
Me.save.Enabled = False

flagsave = False
flagedit = False
flagsearch = False

Set rsd = New ADODB.Recordset
rsd.Open "select * from departmentname", cd2, adOpenStatic, adLockOptimistic

Set rse = New ADODB.Recordset
rse.Open "select * from employeeentry", cd2, adOpenStatic, adLockOptimistic


 '******* fetch department name**********
        
        While (rsd.EOF = False)
            Me.dname.AddItem (rsd.Fields("dname"))
            rsd.MoveNext
        Wend
Me.DTPicker1.Value = Now
Me.DTPicker2.Value = Now

Me.atype.AddItem ("AB")
Me.atype.AddItem ("P")
Me.atype.AddItem ("HOLYDAY")

End Sub
Private Sub add_Click()
 Me.searchframe.Visible = False
 
Me.aentry.Visible = True
Me.aentry.Enabled = True
Me.save.Enabled = True
Me.atype.Enabled = False
Me.dname.SetFocus

flagedit = False
   
End Sub
Private Sub dname_Click()
 Me.searchframe.Visible = False
        Me.ename.Clear
        rse.MoveFirst
            While (rse.EOF = False)
                If (rse.Fields(0) = dname.Text) Then
                    Me.ename.AddItem (rse.Fields(2))
                End If
                rse.MoveNext
            Wend
End Sub

Private Sub ename_Click()
 Me.searchframe.Visible = False
     
 Me.dstart.Text = "00:00:00 AM"
 Me.dend.Text = "00:00:00 AM"
 Me.lstart.Text = "00:00:00 AM"
 Me.lend.Text = "00:00:00 AM"
If flagedit = True Then

     Set rsf = New ADODB.Recordset
         rsf.Open "select * from attendenceentry where date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' and ename='" & ename.Text & "'", cd2, adOpenStatic, adLockOptimistic
        
         rsf.Requery
         If rsf.RecordCount = 0 Then
         Else
            rsf.MoveFirst
            count1 = 0
            While (count1 < rsf.RecordCount)
                  If ((Me.ename.Text) = (rsf.Fields(2))) Then
                         Me.atype.Text = rsf.Fields(4)
                         Me.avalue.Text = rsf.Fields(5)
                         
                         If rsf.Fields(6) = "12:00:00 AM" Then
                            Me.dstart.Text = "00:00:00 AM"
                         Else
                            Me.dstart.Text = rsf.Fields(6)
                         End If
                         
                         If rsf.Fields(7) = "12:00:00 AM" Then
                            Me.dend.Text = "00:00:00 AM"
                         Else
                            Me.dend.Text = rsf.Fields(7)
                         End If
                         
                         If rsf.Fields(8) = "12:00:00 AM" Then
                            Me.lstart.Text = "00:00:00 AM"
                         Else
                         Me.lstart.Text = rsf.Fields(8)
                         End If
                         
                         If rsf.Fields(9) = "12:00:00 AM" Then
                            Me.lend.Text = "00:00:00 AM"
                         Else
                            Me.lend.Text = rsf.Fields(9)
                         End If
                         
                         Me.totalhour.Text = rsf.Fields(10)
                   End If
                count1 = count1 + 1
                rsf.MoveNext
            Wend
        End If
    End If
Me.atype.Enabled = True
End Sub
Private Sub DTPicker1_closeup()
    Me.searchframe.Visible = False
    
   
 '********************************
 If flagedit = True Then

     Set rsf = New ADODB.Recordset
         rsf.Open "select * from attendenceentry where date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' and ename='" & ename.Text & "' ", cd2, adOpenStatic, adLockOptimistic
        
         rsf.Requery
         If rsf.RecordCount = 0 Then
            Me.dstart.Text = ""
            Me.dend.Text = ""
            Me.lstart.Text = ""
            Me.lend.Text = ""
         Else
            rsf.MoveFirst
            count1 = 0
            While (count1 < rsf.RecordCount)
                  If ((Me.ename.Text) = (rsf.Fields(2))) Then
                         Me.atype.Text = rsf.Fields(4)
                         Me.avalue.Text = rsf.Fields(5)
                         
                         If rsf.Fields(6) = "12:00:00 AM" Then
                            Me.dstart.Text = "00:00:00 AM"
                         Else
                            Me.dstart.Text = rsf.Fields(6)
                         End If
                         
                         If rsf.Fields(7) = "12:00:00 AM" Then
                            Me.dend.Text = "00:00:00 AM"
                         Else
                            Me.dend.Text = rsf.Fields(7)
                         End If
                         
                         If rsf.Fields(8) = "12:00:00 AM" Then
                            Me.lstart.Text = "00:00:00 AM"
                         Else
                         Me.lstart.Text = rsf.Fields(8)
                         End If
                         
                         If rsf.Fields(9) = "12:00:00 AM" Then
                            Me.lend.Text = "00:00:00 AM"
                         Else
                            Me.lend.Text = rsf.Fields(9)
                         End If
                         
                         Me.totalhour.Text = rsf.Fields(10)
                   End If
                count1 = count1 + 1
                rsf.MoveNext
            Wend
        End If
    End If
Me.atype.Enabled = True
End Sub
Private Sub DTPicker1_change()
        Me.searchframe.Visible = False
        
   If flagedit = True Then
 
        Set rsf = New ADODB.Recordset
        rsf.Open "select * from attendenceentry where date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' and ename='" & ename.Text & "' ", cd2, adOpenStatic, adLockOptimistic
        
        rsf.Requery
        
        If rsf.RecordCount = 0 Then
                Me.dstart.Text = ""
                Me.dend.Text = ""
                Me.lstart.Text = ""
                Me.lend.Text = ""
        Else
        rsf.MoveFirst
           count1 = 0
           While (count1 < rsf.RecordCount)
                  If ((Me.ename.Text) = (rsf.Fields(2))) Then
                         Me.atype.Text = rsf.Fields(4)
                         Me.avalue.Text = rsf.Fields(5)
                         If rsf.Fields(6) = "12:00:00 AM" Then
                            Me.dstart.Text = "00:00:00 AM"
                         Else
                            Me.dstart.Text = rsf.Fields(6)
                         End If
                         
                         If rsf.Fields(7) = "12:00:00 AM" Then
                            Me.dend.Text = "00:00:00 AM"
                         Else
                            Me.dend.Text = rsf.Fields(7)
                         End If
                         
                         If rsf.Fields(8) = "12:00:00 AM" Then
                            Me.lstart.Text = "00:00:00 AM"
                         Else
                         Me.lstart.Text = rsf.Fields(8)
                         End If
                         
                         If rsf.Fields(9) = "12:00:00 AM" Then
                            Me.lend.Text = "00:00:00 AM"
                         Else
                            Me.lend.Text = rsf.Fields(9)
                         End If
                         Me.totalhour.Text = rsf.Fields(10)
                   End If
                count1 = count1 + 1
                rsf.MoveNext
            Wend
      End If
    End If

Me.atype.Enabled = True
End Sub
Private Sub DTPicker1_lostfocus()
        Me.searchframe.Visible = False
   
   If flagedit = True Then
 
        Set rsf = New ADODB.Recordset
        rsf.Open "select * from attendenceentry where date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' and ename='" & ename.Text & "'", cd2, adOpenStatic, adLockOptimistic
        rsf.Requery
        
        If rsf.RecordCount = 0 Then
                Me.dstart.Text = ""
                Me.dend.Text = ""
                Me.lstart.Text = ""
                Me.lend.Text = ""
        Else
           rsf.MoveFirst
           count1 = 0
            While (count1 < rsf.RecordCount)
                  If ((Me.ename.Text) = (rsf.Fields(2))) Then
                         Me.atype.Text = rsf.Fields(4)
                         Me.avalue.Text = rsf.Fields(5)
                         If rsf.Fields(6) = "12:00:00 AM" Then
                            Me.dstart.Text = "00:00:00 AM"
                         Else
                            Me.dstart.Text = rsf.Fields(6)
                         End If
                         
                         If rsf.Fields(7) = "12:00:00 AM" Then
                            Me.dend.Text = "00:00:00 AM"
                         Else
                            Me.dend.Text = rsf.Fields(7)
                         End If
                         
                         If rsf.Fields(8) = "12:00:00 AM" Then
                            Me.lstart.Text = "00:00:00 AM"
                         Else
                         Me.lstart.Text = rsf.Fields(8)
                         End If
                         
                         If rsf.Fields(9) = "12:00:00 AM" Then
                            Me.lend.Text = "00:00:00 AM"
                         Else
                            Me.lend.Text = rsf.Fields(9)
                         End If
                         Me.totalhour.Text = rsf.Fields(10)
                   End If
                    count1 = count1 + 1
                    rsf.MoveNext
                Wend
        End If
    End If  'end of flag edit stat
Me.dstart.SetFocus
Me.atype.Enabled = True
End Sub
Private Sub atype_click()

 If (dstart.Text = "00:00:00 AM" Or dstart.Text = "00:00:00 PM" Or dend.Text = "00:00:00 AM" Or dend.Text = "00:00:00 PM") Then
    Me.totalhour.Text = ""
 Else
            If (lstart.Text = "00:00:00 AM" Or lstart.Text = "00:00:00 PM" Or lend.Text = "00:00:00 AM" Or lend.Text = "00:00:00 PM") Then
                                 lunchtime = 0
            Else
                                la = Format$(lstart.Text, "HH:mm")
                                lb = Format$(lend.Text, "HH:mm")
                                lunchtime = DateDiff("n", la, lb)
            End If
                                da = Format$(dstart, "HH:mm")
                                db = Format$(dend, "HH:mm")
                                daytime = DateDiff("n", da, db)
                                
                                totaltime = daytime - lunchtime
                                
                                thour = (totaltime / 60)
                                totalminute = (totaltime Mod 60)
                                strlunchtime = "" + Str(thour) + " h" + Str(totalminute) + " m"
                                Me.totalhour.Text = strlunchtime
End If


End Sub
Private Sub atype_gotfocus()

  If (dstart.Text = "00:00:00 AM" Or dstart.Text = "00:00:00 PM" Or dend.Text = "00:00:00 AM" Or dend.Text = "00:00:00 PM") Then
     Me.totalhour.Text = ""
 Else
            If (lstart.Text = "00:00:00 AM" Or lstart.Text = "00:00:00 PM" Or lend.Text = "00:00:00 AM" Or lend.Text = "00:00:00 PM") Then
                               lunchtime = 0
            Else
                                la = Format$(lstart.Text, "HH:mm")
                                lb = Format$(lend.Text, "HH:mm")
                                lunchtime = DateDiff("n", la, lb)
            End If
            
                                
                                da = Format$(dstart, "HH:mm")
                                db = Format$(dend, "HH:mm")
                                daytime = DateDiff("n", da, db)
                                
                                totaltime = daytime - lunchtime
                                
                                thour = (totaltime / 60)
                                totalminute = (totaltime Mod 60)
                                strlunchtime = "" + Str(thour) + " h" + Str(totalminute) + " m"
                                Me.totalhour.Text = strlunchtime
End If
End Sub
Private Sub avalue_click()
 Me.searchframe.Visible = False
        
End Sub
Private Sub avalue_gotfocus()
Me.searchframe.Visible = False

       Me.avalue.Clear
       Me.avalue.AddItem ("1")
      Me.avalue.AddItem ("0.5")
End Sub
Private Sub save_Click()
                               
                If Me.dname.Text = "" Then
                    MsgBox "ENTER ENOUGH DATA!!!"
                    Me.dname.SetFocus
                    Exit Sub
                End If
                If Me.ename.Text = "" Then
                    MsgBox "ENTER ENOUGH DATA!!!"
                    Me.ename.SetFocus
                    Exit Sub
                End If
                
                If Me.atype.Text = "" Then
                    Me.atype.Text = "-"
                End If
                
                If Me.avalue.Text = "" Then
                    Me.avalue.Text = "0"
                End If
                
                If Me.dstart.Text = "" Then
                    Me.dstart.Text = "00:00:00 AM"
                End If
                
                If Me.dend.Text = "" Then
                    Me.dend.Text = "00:00:00 AM"
                End If

                If Me.lstart.Text = "" Then
                    Me.lstart.Text = "00:00:00 AM"
                End If
                
                If Me.lend.Text = "" Then
                    Me.lend.Text = "00:00:00 AM"
                End If
                
                If Me.totalhour.Text = "" Then
                    Me.totalhour.Text = "0"
                End If
                
                If Me.atype.Text = "AB" Or Me.atype.Text = "HOLYDAY" Then
                    Me.dstart.Text = "00:00:00 AM"
                    Me.dend.Text = "00:00:00 AM"
                    Me.lstart.Text = "00:00:00 AM"
                    Me.lend.Text = "00:00:00 AM"
                    Me.avalue.Enabled = False
                Else
                    Me.avalue.Enabled = True
                End If
           
             
 If flagedit = False Then
 '*********** if record is already there
        Set rsg = New ADODB.Recordset
            rsg.Open "select * from attendenceentry where date1='" & Format$(DTPicker1.Value, "d/m/yyyy") & "' ", cd2, adOpenStatic, adLockOptimistic
            rsg.Requery
            
            If (rsg.RecordCount = 0) Then
            Else
                rsg.MoveFirst
                count1 = 0
            While (count1 < rsg.RecordCount)
                If ((Me.ename.Text) = (rsg.Fields(2))) Then
                      MsgBox "already in record"
                      Me.add.SetFocus
                      GoTo here
                End If
                count1 = count1 + 1
                rsg.MoveNext
            Wend
            '*********************************
            End If
                    rs2.AddNew
                    rs2.Fields(0) = Me.dname.Text
                    rs2.Fields(2) = Me.ename.Text
                    rs2.Fields(3) = Format$(Me.DTPicker1.Value, "dd/mmm/yyyy")
                    rs2.Fields(4) = Me.atype.Text
                    rs2.Fields(5) = Me.avalue.Text
                    rs2.Fields(6) = Me.dstart.Text
                    rs2.Fields(7) = Me.dend.Text
                    rs2.Fields(8) = Me.lstart.Text
                    rs2.Fields(9) = Me.lend.Text
                    rs2.Fields(10) = Me.totalhour.Text
                    rs2.Fields(11) = Format$(Me.DTPicker1.Value, "dd/M/yyyy")
                    
                    rs2.Update
                                                
                    MsgBox "record has been saved"
                    Me.save.Enabled = False
           
ElseIf flagedit = True Then
                '******* check for edit***
        Set rsh = New ADODB.Recordset
        
        rsh.Open "select * from attendenceentry where date1='" & Format$(DTPicker1.Value, "d/M/yyyy") & "' ", cd2, adOpenStatic, adLockOptimistic
        
        count1 = 0
        rsh.Requery
                          While (count1 < rsh.RecordCount)
                                 If ((Me.ename.Text) = (rsh.Fields(2))) Then
                                     rsh.delete
                                 End If
                                 count1 = count1 + 1
                                 rsh.MoveNext
                            Wend
        
                     
                '****************************
                            rs2.AddNew
                            rs2.Fields(0) = Me.dname.Text
                            rs2.Fields(2) = Me.ename.Text
                            rs2.Fields(3) = Format$(Me.DTPicker1.Value, "dd/mmm/yyyy")
                            rs2.Fields(4) = Me.atype.Text
                            rs2.Fields(5) = Me.avalue.Text
                            rs2.Fields(6) = Me.dstart.Text
                            rs2.Fields(7) = Me.dend.Text
                            rs2.Fields(8) = Me.lstart.Text
                            rs2.Fields(9) = Me.lend.Text
                            rs2.Fields(10) = Me.totalhour.Text
                            rs2.Fields(11) = Format$(Me.DTPicker1.Value, "dd/M/yyyy")
                            rs2.Update
                           'rs2.Sort
                    MsgBox "record has been saved"
       
End If

       
here:
                        Me.aentry.Enabled = False
                
                            Me.dname.Text = ""
                            Me.ename.Text = ""
                            Me.atype.Text = ""
                            Me.avalue.Text = ""
                            Me.dstart.Text = ""
                            Me.dend.Text = ""
                            Me.lstart.Text = ""
                            Me.lend.Text = ""
                            Me.totalhour.Text = ""
                            
                
End Sub
Private Sub edit_Click()
Me.searchframe.Visible = False
Me.aentry.Visible = True

aentry.Enabled = True
Me.save.Enabled = True
flagedit = True
Me.dname.SetFocus
End Sub

Private Sub search_Click()
 Me.searchframe.Visible = True
 Me.aentry.Visible = False
 Me.save.Enabled = False
 'Me.delete.Enabled = True
 
 Me.cmbename.Clear
 
 
 rsd.Requery
 If (rsd.RecordCount = 0) Then
 Else
 rsd.MoveFirst
    While (rsd.EOF = False)
        Me.cmbdname.AddItem (rsd.Fields(0))
        rsd.MoveNext
    Wend
 End If
Me.cmbdname.SetFocus

End Sub
Private Sub cmbdname_click()
        Me.cmbename.Clear
        If (rsd.RecordCount = 0) Then
        Else
        rse.MoveFirst
            While (rse.EOF = False)
                If (rse.Fields(0) = cmbdname.Text) Then
                Me.cmbename.AddItem (rse.Fields(2))
                End If
            rse.MoveNext
            Wend
            End If
  
End Sub
Private Sub DTPicker2_CLOSEUP()
flagsearch = False
    
Set rsi = New ADODB.Recordset
rsi.Open "select * from attendenceentry where date1='" & Format$(DTPicker2.Value, "d/M/yyyy") & "' and ename='" & cmbename.Text & "'", cd2, adOpenStatic, adLockOptimistic
        
rsi.Requery
If (rsi.RecordCount = 0) Then
   Me.txtdstart.Text = ""
   Me.txtdend.Text = ""
   Me.txtlstart.Text = ""
   Me.txtlend.Text = ""
   Me.txtatype.Text = ""
   Me.txtavalue.Text = ""
   'Me.DTPicker2.SetFocus
Else
    
    count1 = 0
        While (count1 < rsi.RecordCount)
         
                If (Me.cmbename.Text = rsi.Fields(2)) Then
                    Me.txtatype.Text = rsi.Fields(4)
                    Me.txtavalue.Text = rsi.Fields(5)
                    
                    
                        If rsi.Fields(6) = "12:00:00 AM" Then
                            Me.txtdstart.Text = "00:00:00 AM"
                        Else
                            Me.txtdstart.Text = rsi.Fields(6)
                        End If
                         
                        If rsi.Fields(7) = "12:00:00 AM" Then
                            Me.txtdend.Text = "00:00:00 AM"
                        Else
                            Me.txtdend.Text = rsi.Fields(7)
                        End If
                         
                        If rsi.Fields(8) = "12:00:00 AM" Then
                            Me.txtlstart.Text = "00:00:00 AM"
                        Else
                         Me.txtlstart.Text = rsi.Fields(8)
                        End If
                         
                        If rsi.Fields(9) = "12:00:00 AM" Then
                            Me.txtlend.Text = "00:00:00 AM"
                        Else
                            Me.txtlend.Text = rsi.Fields(9)
                        End If
                         
                        Me.txttotalhour.Text = rsi.Fields(10)
                 
                    flagsearch = True
                End If
            count1 = count1 + 1
            rsi.MoveNext
        Wend

        'If flagsearch = False Then
        'MsgBox "no record "
        'Me.DTPicker2.SetFocus
        'End If
End If
End Sub
Private Sub DTPicker2_change()
flagsearch = False
   
Set rsi = New ADODB.Recordset
rsi.Open "select * from attendenceentry where  date1='" & Format$(DTPicker2.Value, "d/M/yyyy") & "' and ename='" & cmbename.Text & "'", cd2, adOpenStatic, adLockOptimistic
        
rsi.Requery
If (rsi.RecordCount = 0) Then
 Me.txtdstart.Text = ""
   Me.txtdend.Text = ""
   Me.txtlstart.Text = ""
   Me.txtlend.Text = ""
   Me.txtatype.Text = ""
   Me.txtavalue.Text = ""
   'Me.DTPicker2.SetFocus
Else
    count1 = 0
        While (count1 < rsi.RecordCount)
         
                If (rsi.Fields(2) = Me.cmbename.Text) Then
                    Me.txtatype.Text = rsi.Fields(4)
                    Me.txtavalue.Text = rsi.Fields(5)
                    If rsi.Fields(6) = "12:00:00 AM" Then
                            Me.txtdstart.Text = "00:00:00 AM"
                        Else
                            Me.txtdstart.Text = rsi.Fields(6)
                        End If
                         
                        If rsi.Fields(7) = "12:00:00 AM" Then
                            Me.txtdend.Text = "00:00:00 AM"
                        Else
                            Me.txtdend.Text = rsi.Fields(7)
                        End If
                         
                        If rsi.Fields(8) = "12:00:00 AM" Then
                            Me.txtlstart.Text = "00:00:00 AM"
                        Else
                         Me.txtlstart.Text = rsi.Fields(8)
                        End If
                         
                        If rsi.Fields(9) = "12:00:00 AM" Then
                            Me.txtlend.Text = "00:00:00 AM"
                        Else
                            Me.txtlend.Text = rsi.Fields(9)
                        End If
                    Me.txttotalhour.Text = rsi.Fields(10)
                    flagsearch = True
               
                End If
            count1 = count1 + 1
            rsi.MoveNext
        Wend
 
        'If flagsearch = False Then
        ' Me.cmbename.SetFocus
        'End If
  End If
End Sub
Private Sub DTPicker2_lostfocus()
flagsearch = False
  
Set rsi = New ADODB.Recordset
rsi.Open "select * from attendenceentry where  date1='" & Format$(DTPicker2.Value, "d/M/yyyy") & "' and ename='" & cmbename.Text & "' ", cd2, adOpenStatic, adLockOptimistic
        
rsi.Requery
If (rsi.RecordCount = 0) Then
    Me.txtdstart.Text = ""
    Me.txtdend.Text = ""
    Me.txtlstart.Text = ""
    Me.txtlend.Text = ""
    Me.txtatype.Text = ""
    Me.txtavalue.Text = ""
    'Me.DTPicker2.SetFocus
Else
count1 = 0
        While (count1 < rsi.RecordCount)
         
                If (rsi.Fields(2) = Me.cmbename.Text) Then
                    Me.txtatype.Text = rsi.Fields(4)
                    Me.txtavalue.Text = rsi.Fields(5)
                    If rsi.Fields(6) = "12:00:00 AM" Then
                            Me.txtdstart.Text = "00:00:00 AM"
                        Else
                            Me.txtdstart.Text = rsi.Fields(6)
                        End If
                         
                        If rsi.Fields(7) = "12:00:00 AM" Then
                            Me.txtdend.Text = "00:00:00 AM"
                        Else
                            Me.txtdend.Text = rsi.Fields(7)
                        End If
                         
                        If rsi.Fields(8) = "12:00:00 AM" Then
                            Me.txtlstart.Text = "00:00:00 AM"
                        Else
                         Me.txtlstart.Text = rsi.Fields(8)
                        End If
                         
                        If rsi.Fields(9) = "12:00:00 AM" Then
                            Me.txtlend.Text = "00:00:00 AM"
                        Else
                            Me.txtlend.Text = rsi.Fields(9)
                        End If
                    Me.txttotalhour.Text = rsi.Fields(10)
                    flagsearch = True
               
                End If
            count1 = count1 + 1
            rsi.MoveNext
        Wend
 
        'If flagsearch = False Then
        '    MsgBox "No record!! "
        '    Me.cmbename.SetFocus
        'End If
   End If
End Sub



