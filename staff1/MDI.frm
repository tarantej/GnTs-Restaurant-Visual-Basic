VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C04080&
   Caption         =   "STAFF ATTENDENCE"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13950
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI.frx":29C12
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10695
      Left            =   0
      ScaleHeight     =   10635
      ScaleWidth      =   13890
      TabIndex        =   0
      Top             =   0
      Width           =   13950
      Begin VB.Image Image1 
         Height          =   10680
         Left            =   -240
         Picture         =   "MDI.frx":3FAFF
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15465
      End
   End
   Begin VB.Menu master 
      Caption         =   "MASTER"
      Index           =   1
      Begin VB.Menu eentry 
         Caption         =   "Employee Entry"
         Index           =   1
      End
      Begin VB.Menu dentry 
         Caption         =   "Department Entry"
         Index           =   2
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Index           =   3
      End
   End
   Begin VB.Menu attendence 
      Caption         =   "ATTANDENCE"
      Index           =   1
      Begin VB.Menu eattandence 
         Caption         =   "Employee Attendence"
         Index           =   1
      End
      Begin VB.Menu pay_slip 
         Caption         =   "PAY SLIP"
         Index           =   1
      End
   End
   Begin VB.Menu report 
      Caption         =   "REPORT"
      Index           =   1
      Begin VB.Menu areport 
         Caption         =   "ATTENDENCE REPORT"
         Index           =   1
      End
      Begin VB.Menu empdetail 
         Caption         =   "EMPLOYEE DETAIL"
         Index           =   1
      End
      Begin VB.Menu monthly_pay_report 
         Caption         =   "MONTHLY PAY REPORT "
         Index           =   3
      End
   End
   Begin VB.Menu backup 
      Caption         =   "BACKUP"
      Index           =   1
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub areport_Click(Index As Integer)
Attendence_report.Show

End Sub

Private Sub backup_Click(Index As Integer)
backup_fm.Show
End Sub

Private Sub dentry_Click(Index As Integer)
Department_entry.Show
End Sub

Private Sub eattandence_Click(Index As Integer)
Attendence_entry.Show
End Sub

Private Sub eentry_Click(Index As Integer)
Employee_entry.Show
Employee_entry.add.SetFocus

End Sub

Private Sub empdetail_Click(Index As Integer)
Employee_detail_report.Show
End Sub

Private Sub exit_Click(Index As Integer)
End
End Sub

Private Sub monthly_pay_report_Click(Index As Integer)
pay_slip_report.Show
End Sub

Private Sub pay_slip_Click(Index As Integer)
pay_slip_fm.Show
End Sub
