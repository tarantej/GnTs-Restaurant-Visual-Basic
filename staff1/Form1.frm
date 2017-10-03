VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Employee_entry 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Employee Entry"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsearch 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton exit 
      BackColor       =   &H00C0E0FF&
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      TabIndex        =   39
      Top             =   5400
      Width           =   3015
      Begin VB.CommandButton first 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&First"
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
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton next 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Next"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   960
         Width           =   1380
      End
      Begin VB.CommandButton privious 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Privious"
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
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   960
         Width           =   1380
      End
      Begin VB.CommandButton last 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Last"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   38
      Top             =   5400
      Width           =   2775
      Begin VB.CommandButton edit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Edit"
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
         TabIndex        =   43
         Top             =   960
         Width           =   1260
      End
      Begin VB.CommandButton delete 
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   960
         Width           =   1260
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1260
      End
      Begin VB.CommandButton add 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Add"
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
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   1260
      End
   End
   Begin VB.Frame searchframe 
      BackColor       =   &H00C0E0FF&
      Caption         =   "-------------------------Search Employee Record-------------------"
      Enabled         =   0   'False
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
      Height          =   5415
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   6975
      Begin VB.TextBox Text1 
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
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.TextBox lblphno_o 
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
         Left            =   2400
         TabIndex        =   7
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox lblphno_m 
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
         Left            =   2400
         TabIndex        =   8
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox lblremark 
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
         Left            =   2400
         TabIndex        =   12
         Top             =   4920
         Width           =   2175
      End
      Begin VB.TextBox lbledu 
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
         Left            =   2400
         TabIndex        =   11
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox lbljdate 
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
         Left            =   2400
         TabIndex        =   10
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox lblsalary 
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
         Left            =   2400
         TabIndex        =   9
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox lblphno_r 
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
         Left            =   2400
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox lblpincode 
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
         Left            =   2400
         TabIndex        =   5
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox lblcity 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox lbladdress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   3855
      End
      Begin VB.ListBox lstename 
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
         ItemData        =   "Form1.frx":0000
         Left            =   2400
         List            =   "Form1.frx":0002
         TabIndex        =   48
         Top             =   840
         Width           =   3855
      End
      Begin VB.ListBox lstdname 
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
         ItemData        =   "Form1.frx":0004
         Left            =   2400
         List            =   "Form1.frx":0006
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0E0FF&
         Caption         =   "(M)"
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
         Left            =   1920
         TabIndex        =   61
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0E0FF&
         Caption         =   "(O)"
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
         Left            =   1920
         TabIndex        =   60
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Phone No         (R)"
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
         Left            =   360
         TabIndex        =   57
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         Caption         =   "City:"
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
         Left            =   360
         TabIndex        =   56
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remark:"
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
         Left            =   360
         TabIndex        =   55
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Education:"
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
         Left            =   360
         TabIndex        =   54
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Joining Date:"
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
         Left            =   360
         TabIndex        =   53
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Salary:"
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
         Left            =   360
         TabIndex        =   52
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pin code:"
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
         Left            =   360
         TabIndex        =   51
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address:"
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
         Left            =   360
         TabIndex        =   50
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame eentry 
      BackColor       =   &H00C0E0FF&
      Caption         =   "----------------------  EMPLOYEE ENTRY  --------------------------"
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
      Height          =   5415
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   6975
      Begin VB.TextBox remark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2400
         TabIndex        =   24
         Top             =   4920
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker jdate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   22
         Top             =   4200
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19267587
         CurrentDate     =   38617
      End
      Begin VB.TextBox phno_o 
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
         Left            =   2400
         TabIndex        =   19
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox phno_m 
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
         Left            =   2400
         TabIndex        =   20
         Top             =   3480
         Width           =   1935
      End
      Begin VB.ListBox dname 
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
         ItemData        =   "Form1.frx":0008
         Left            =   2400
         List            =   "Form1.frx":000A
         TabIndex        =   13
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox ename 
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
         Left            =   2400
         TabIndex        =   14
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox salary 
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
         Left            =   2400
         TabIndex        =   21
         Top             =   3840
         Width           =   1935
      End
      Begin VB.TextBox edu 
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
         Left            =   2400
         TabIndex        =   23
         Top             =   4560
         Width           =   1935
      End
      Begin VB.TextBox phno_r 
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
         Left            =   2400
         TabIndex        =   18
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox address 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox city 
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
         Left            =   2400
         TabIndex        =   16
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox pincode 
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
         Left            =   2400
         TabIndex        =   17
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "(M)"
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
         Left            =   1920
         TabIndex        =   59
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "(O)"
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
         Left            =   1920
         TabIndex        =   58
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Joining Date:"
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
         Left            =   360
         TabIndex        =   37
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Education:"
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
         Left            =   360
         TabIndex        =   36
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remark:"
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
         Left            =   360
         TabIndex        =   35
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Salary:"
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
         Left            =   360
         TabIndex        =   34
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Phone No.       :(R)"
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
         Left            =   360
         TabIndex        =   33
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Pin Code:"
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
         Left            =   360
         TabIndex        =   32
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "City:"
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
         Left            =   360
         TabIndex        =   31
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Address:"
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
         Left            =   360
         TabIndex        =   30
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   2055
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Employee_entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cd As ADODB.Connection
Dim rsd As ADODB.Recordset
Dim rse  As ADODB.Recordset
Dim rsen As ADODB.Recordset
Dim rssearch As ADODB.Recordset
Dim rsnamesearch As ADODB.Recordset
Dim count1  As Integer
Dim flagedit, flagsave As Integer

Private Sub Form_Load()

Me.eentry.Enabled = False
Me.searchframe.Visible = False

Me.delete.Enabled = False
Me.save.Enabled = False
Me.delete.Enabled = False
Me.edit.Enabled = False

flagedit = False
flagsave = False

Set cd = New ADODB.Connection
    cd.Provider = "Microsoft.Jet.OLEDB.4.0;"
    cd.Open App.Path & "\db1.mdb"

Set rse = New ADODB.Recordset
    rse.Open "select * from employeeentry", cd, adOpenStatic, adLockOptimistic
    rse.Requery
Set rsen = New ADODB.Recordset
    rsen.Open "select * from employeeentry", cd, adOpenStatic, adLockOptimistic
    rsen.Requery

Me.dname.Text = ""
Me.ename.Text = ""
Me.address.Text = ""
Me.city.Text = ""
Me.pincode.Text = ""
Me.phno_r.Text = ""
Me.salary.Text = ""
Me.jdate.Value = Now
Me.edu.Text = ""
Me.remark.Text = ""

End Sub
' ************ for adding record*********
Private Sub add_Click()
Me.searchframe.Visible = False
Me.eentry.Visible = True
Me.eentry.Enabled = True

Me.edit.Enabled = False

flagedit = False
flagsave = False

Me.dname.Enabled = True
Me.ename.Enabled = True
Me.address.Enabled = True
Me.city.Enabled = True
Me.pincode.Enabled = True
Me.phno_r.Enabled = True
Me.salary.Enabled = True
Me.edu.Enabled = True
Me.remark.Enabled = True
Me.phno_o.Enabled = True
Me.phno_m.Enabled = True


Me.dname.SetFocus
Me.delete.Enabled = False
Me.save.Enabled = True
'**** for combo entry****
    Set rsd = New ADODB.Recordset
    rsd.Open "select * from departmentname", cd, adOpenStatic, adLockOptimistic
If rsd.RecordCount = 0 Then
MsgBox "NO record for department"
Exit Sub
End If
        count1 = 0
        rsd.Requery
        
    Me.dname.Clear
    While (count1 < rsd.RecordCount)
        Me.dname.AddItem (rsd.Fields(0))
        count1 = count1 + 1
    rsd.MoveNext
    Wend
'********************
    Me.dname.ListIndex = 0
    Me.ename.Text = ""
    Me.address.Text = ""
    Me.city.Text = ""
    Me.pincode.Text = ""
    Me.phno_r.Text = ""
    Me.salary.Text = ""
    Me.edu.Text = ""
    Me.remark.Text = ""
    Me.phno_o.Text = ""
    Me.phno_m.Text = ""

End Sub
Private Sub ename_keypress(KeyAscii As Integer)
Dim a As String
    a = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(a, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub ename_lostfocus()
    Set rse = New ADODB.Recordset
        rse.Open "select * from employeeentry where ename='" & ename.Text & "'", cd, adOpenStatic, adLockOptimistic
        If rse.RecordCount > 0 Then
            MsgBox "THIS NAME IS ALREADY STORED!!!"
            ename.SetFocus
        End If
End Sub
Private Sub city_keypress(KeyAscii As Integer)
Dim a As String
a = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
 KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(a, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub



Private Sub pincode_KeyPress(KeyAscii As Integer)
Dim strValid7 As String
    strValid7 = "0123456789"
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid7, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub phno_r_KeyPress(KeyAscii As Integer)
Dim strValid7 As String
    strValid7 = "0123456789()-"
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid7, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub phno_o_KeyPress(KeyAscii As Integer)
Dim strValid7 As String
    strValid7 = "0123456789()-"
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid7, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub phno_m_KeyPress(KeyAscii As Integer)
Dim strValid7 As String
    strValid7 = "0123456789()-"
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid7, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub salary_KeyPress(KeyAscii As Integer)
Dim strValid7 As String
    strValid7 = "0123456789"
    KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(strValid7, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub remark_keypress(KeyAscii As Integer)
Dim a As String
a = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789./-"
 KeyAscii = Asc(Chr(KeyAscii))
    If KeyAscii > 26 Then ' if it's not a control code
        If InStr(a, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub save_Click()
If flagsave = False Then
    If flagedit = False Then
        
        If Me.dname.Text = "" Then
            MsgBox "ENTER ENOUGH DATA!!"
            Exit Sub
            End If
        If Me.ename.Text = "" Then
            MsgBox "ENTER ENOUGH DATA!!"
            Exit Sub
        End If
        
        If Me.address.Text = "" Then
        Me.address.Text = "-"
        End If
        
        If Me.city.Text = "" Then
            Me.city.Text = "-"
        End If
        
        If Me.pincode.Text = "" Then
            Me.pincode.Text = "0"
        End If
        
        If Me.phno_r.Text = "" Then
            Me.phno_r.Text = "0"
        End If
        If Me.phno_o.Text = "" Then
            Me.phno_o.Text = "0"
        End If
        If Me.phno_m.Text = "" Then
            Me.phno_m.Text = "0"
        End If
        If Me.salary.Text = "" Then
            Me.salary.Text = "0"
        End If
        
        If Me.remark.Text = "" Then
            Me.remark.Text = "-"
        End If
        rse.AddNew

        rse.Fields(0) = Me.dname.Text
        rse.Fields(2) = Me.ename.Text
        rse.Fields(3) = Me.address.Text
        rse.Fields(4) = Me.city.Text
        rse.Fields(5) = Me.pincode.Text
        rse.Fields(6) = Me.phno_r.Text
        rse.Fields(7) = Me.salary.Text
        Me.jdate.Value = Format$(Me.jdate.Value, "d/M/yyyy")
        rse.Fields(8) = Format$(Me.jdate.Value, "dd/mmm/yyyy")
        rse.Fields(9) = Me.edu.Text
        rse.Fields(10) = Me.remark.Text
        rse.Fields(11) = Me.phno_o.Text
        rse.Fields(12) = Me.phno_m.Text
    Else
            rse.MoveFirst
             While (Not rse.EOF)
                 If (Me.lstename.Text = rse.Fields(2)) Then
                    rse.delete
                    rse.Update
                    flagedit = False
                 End If
                 rse.MoveNext
            Wend
        rse.AddNew
        rse.Fields(0) = Me.lstdname.List(Me.lstdname.ListIndex)
        rse.Fields(2) = Text1.Text
        rse.Fields(3) = Me.lbladdress.Text
        rse.Fields(4) = Me.lblcity.Text
        rse.Fields(5) = Me.lblpincode.Text
        rse.Fields(6) = Me.lblphno_r.Text
        rse.Fields(7) = Me.lblsalary.Text
        Me.lbljdate.Text = Format$(Me.lbljdate.Text, "d/M/yyyy")
        rse.Fields(8) = Format$(Me.lbljdate.Text, "dd/mmm/yyyy")
        rse.Fields(9) = Me.lbledu.Text
        rse.Fields(10) = Me.lblremark.Text
        rse.Fields(11) = Me.phno_o.Text
        rse.Fields(12) = Me.phno_m.Text
        Me.Text1.Visible = False
    End If
        flagsave = True
        rse.Update
        rse.Requery
        MsgBox "record has been saved"
        Me.save.Enabled = False
        Me.edit.Enabled = False
        Me.delete.Enabled = False
        Else
    MsgBox "record already saved"
End If

Me.dname.Text = ""
Me.ename.Text = ""
Me.address.Text = ""
Me.city.Text = ""
Me.pincode.Text = ""
Me.phno_r.Text = ""
Me.salary.Text = ""
Me.edu.Text = ""
Me.remark.Text = ""
Me.phno_o.Text = ""
Me.phno_m.Text = ""

Me.add.SetFocus
Me.eentry.Enabled = False

End Sub
Private Sub delete_Click()
Dim employee_no, i As Integer
 
    i = MsgBox("Do You Want To Delete This Employee Name:" + Me.lstename.List(Me.lstename.ListIndex) + "", vbQuestion + vbYesNo, "Deleting Record")
    If (i = vbYes) Then
           If rse.RecordCount <> 0 Then
             While (rse.EOF)
                 If (Me.lstename.Text = rse.Fields(2)) Then
                    rse.delete
                    rse.Update
                    MsgBox "record has been deleted"
                 End If
                 rse.MoveNext
            Wend
            
            Me.lbladdress.Text = ""
            Me.lblcity.Text = ""
            Me.lblpincode.Text = ""
            Me.lblphno_r.Text = ""
            Me.lblsalary.Text = ""
            Me.lbljdate.Text = ""
            Me.lbledu.Text = ""
            Me.lblremark.Text = ""
            Me.lblphno_o.Text = ""
            Me.lblphno_m.Text = ""
                           
            End If
    End If


Me.delete.Enabled = False
Me.edit.Enabled = False

End Sub

Private Sub first_Click()
Me.delete.Enabled = False
Me.edit.Enabled = False

Me.searchframe.Visible = False
Me.eentry.Enabled = True
Me.eentry.Visible = True

If (rsen.RecordCount = 0) Then
    MsgBox "No record"
    Exit Sub
End If

If (rsen.BOF = False) Then
        rsen.Requery
        rsen.MoveFirst

        Me.dname.List(0) = rsen.Fields(0)
        Me.ename.Text = rsen.Fields(2)
        Me.address.Text = rsen.Fields(3)
        Me.city.Text = rsen.Fields(4)
        Me.pincode.Text = rsen.Fields(5)
        Me.phno_r.Text = rsen.Fields(6)
        Me.salary.Text = rsen.Fields(7)
        Me.jdate.Value = Format$(rsen.Fields(8), "dd/mm/yyyy")
        Me.edu.Text = rsen.Fields(9)
        Me.remark.Text = rsen.Fields(10)
        Me.phno_o.Text = rsen.Fields(11)
        Me.phno_m.Text = rsen.Fields(12)
        
End If
Me.dname.Enabled = False
Me.ename.Enabled = False
Me.address.Enabled = False
Me.city.Enabled = False
Me.pincode.Enabled = False
Me.phno_r.Enabled = False
Me.salary.Enabled = False
Me.edu.Enabled = False
Me.remark.Enabled = False
Me.phno_o.Enabled = False
Me.phno_m.Enabled = False

End Sub
Private Sub last_Click()
Me.delete.Enabled = False
Me.edit.Enabled = False

Me.searchframe.Visible = False
Me.eentry.Enabled = True
Me.eentry.Visible = True

If (rsen.RecordCount = 0) Then
MsgBox "No record"
Exit Sub
End If

If (rsen.EOF = False) Then
rsen.MoveLast
        Me.dname.List(0) = rsen.Fields(0)
        Me.ename.Text = rsen.Fields(2)
        Me.address.Text = rsen.Fields(3)
        Me.city.Text = rsen.Fields(4)
        Me.pincode.Text = rsen.Fields(5)
        Me.phno_r.Text = rsen.Fields(6)
        Me.salary.Text = rsen.Fields(7)
        Me.jdate.Value = Format$(rsen.Fields(8), "dd/mm/yyyy")
        Me.edu.Text = rsen.Fields(9)
        Me.remark.Text = rsen.Fields(10)
        Me.phno_o.Text = rsen.Fields(11)
        Me.phno_m.Text = rsen.Fields(12)
End If
Me.dname.Enabled = False
Me.ename.Enabled = False
Me.address.Enabled = False
Me.city.Enabled = False
Me.pincode.Enabled = False
Me.phno_r.Enabled = False
Me.salary.Enabled = False
Me.edu.Enabled = False
Me.remark.Enabled = False
Me.phno_o.Enabled = False
Me.phno_m.Enabled = False

End Sub
Private Sub privious_Click()
Me.delete.Enabled = False
Me.edit.Enabled = False

Me.searchframe.Visible = False
Me.eentry.Enabled = True
Me.eentry.Visible = True

If (rsen.RecordCount = 0) Then
    MsgBox "No record"
    Exit Sub
End If
    rsen.MovePrevious
If (rsen.BOF = True) Then
    rsen.MoveFirst
End If
 Me.dname.List(0) = rsen.Fields(0)
Me.ename.Text = rsen.Fields(2)
Me.address.Text = rsen.Fields(3)
Me.city.Text = rsen.Fields(4)
Me.pincode.Text = rsen.Fields(5)
Me.phno_r.Text = rsen.Fields(6)
Me.salary.Text = rsen.Fields(7)
Me.jdate.Value = Format$(rsen.Fields(8), "dd/mm/yyyy")
Me.edu.Text = rsen.Fields(9)
Me.remark.Text = rsen.Fields(10)
Me.phno_o.Text = rsen.Fields(11)
Me.phno_m.Text = rsen.Fields(12)

Me.dname.Enabled = False
Me.ename.Enabled = False
Me.address.Enabled = False
Me.city.Enabled = False
Me.pincode.Enabled = False
Me.phno_r.Enabled = False
Me.salary.Enabled = False
Me.edu.Enabled = False
Me.remark.Enabled = False
Me.phno_o.Enabled = False
Me.phno_m.Enabled = False


End Sub
Private Sub next_Click()

Me.delete.Enabled = False
Me.edit.Enabled = False

Me.searchframe.Visible = False
Me.eentry.Enabled = True
Me.eentry.Visible = True

If (rsen.RecordCount = 0) Then
        MsgBox "No record"
    Exit Sub
End If
   rsen.MoveNext
    If (rsen.EOF = True) Then
        rsen.MoveLast
    End If
    Me.dname.List(0) = rsen.Fields(0)
    Me.ename.Text = rsen.Fields(2)
    Me.address.Text = rsen.Fields(3)
    Me.city.Text = rsen.Fields(4)
    Me.pincode.Text = rsen.Fields(5)
    Me.phno_r.Text = rsen.Fields(6)
    Me.salary.Text = rsen.Fields(7)
    Me.jdate.Value = Format$(rsen.Fields(8), "dd/mm/yyyy")
    Me.edu.Text = rsen.Fields(9)
    Me.remark.Text = rsen.Fields(10)
    Me.phno_o.Text = rsen.Fields(11)
    Me.phno_m.Text = rsen.Fields(12)

Me.dname.Enabled = False
Me.ename.Enabled = False
Me.address.Enabled = False
Me.city.Enabled = False
Me.pincode.Enabled = False
Me.phno_r.Enabled = False
Me.salary.Enabled = False
Me.edu.Enabled = False
Me.remark.Enabled = False
Me.phno_o.Enabled = False
Me.phno_m.Enabled = False

End Sub
Private Sub edit_Click()
flagedit = True
flagsave = False

Me.save.Enabled = True
Me.delete.Enabled = False

Me.eentry.Enabled = False
Me.eentry.Visible = False
Me.searchframe.Visible = True

Me.Text1.Visible = True
Me.Text1.Text = Me.lstename.List(Me.lstename.ListIndex)

Me.lbladdress.Enabled = True
Me.lblcity.Enabled = True
Me.lblpincode.Enabled = True
Me.lblphno_r.Enabled = True
Me.lblsalary.Enabled = True
Me.lbljdate.Enabled = True
Me.lbledu.Enabled = True
Me.lblremark.Enabled = True
Me.lblphno_o.Enabled = True
Me.lblphno_m.Enabled = True
Me.Text1.SetFocus
End Sub
Private Sub exit_Click()
Unload Me
End Sub
Private Sub cmdsearch_Click()
Me.searchframe.Visible = True
Me.searchframe.Enabled = True
Me.eentry.Visible = False

Me.delete.Enabled = True
Me.edit.Enabled = True
Me.save.Enabled = False

Me.lstdname.SetFocus
Me.lstdname.Clear

Set rssearch = New ADODB.Recordset
rssearch.Open "Select * from departmentname", cd, adOpenStatic, adLockOptimistic

    While (rssearch.EOF = False)
         Me.lstdname.AddItem (rssearch.Fields(0))
         rssearch.MoveNext
    Wend
End Sub
Private Sub lstdname_click()
Set rsnamesearch = New ADODB.Recordset
rsnamesearch.Open "Select * from employeeentry ", cd, adOpenStatic, adLockOptimistic
rsnamesearch.MoveFirst
Me.lstename.Clear
    While (rsnamesearch.EOF = False)
            If (rsnamesearch.Fields(0) = lstdname.Text) Then
                Me.lstename.AddItem (rsnamesearch.Fields(2))
            End If
            rsnamesearch.MoveNext
    Wend
            Me.lbladdress.Text = ""
            Me.lblcity.Text = ""
            Me.lblpincode.Text = ""
            Me.lblphno_r.Text = ""
            Me.lblsalary.Text = ""
            Me.lbljdate.Text = ""
            Me.lbledu.Text = ""
            Me.lblremark.Text = ""
            Me.lblphno_o.Text = ""
            Me.lblphno_m.Text = ""
End Sub
Private Sub lstename_click()
    rsnamesearch.MoveFirst
    rsnamesearch.Requery
    
    While (rsnamesearch.EOF = False)
            If (rsnamesearch.Fields(2) = Me.lstename.Text) Then
                  Me.lbladdress.Text = rsnamesearch.Fields(3)
                  Me.lblcity.Text = rsnamesearch.Fields(4)
                  Me.lblpincode.Text = rsnamesearch.Fields(5)
                  Me.lblphno_r.Text = rsnamesearch.Fields(6)
                  Me.lblsalary.Text = rsnamesearch.Fields(7)
                  Me.lbljdate.Text = Format$(rsnamesearch.Fields(8), "dd/mm/yyyy")
                  Me.lbledu.Text = rsnamesearch.Fields(9)
                  Me.lblremark.Text = rsnamesearch.Fields(10)
                  Me.lblphno_o.Text = rsnamesearch.Fields(11)
                  Me.lblphno_m.Text = rsnamesearch.Fields(12)
                  
            End If
            rsnamesearch.MoveNext
    Wend
        Me.lbladdress.Enabled = False
        Me.lblcity.Enabled = False
        Me.lblpincode.Enabled = False
        Me.lblphno_r.Enabled = False
        Me.lblsalary.Enabled = False
        Me.lbljdate.Enabled = False
        Me.lbledu.Enabled = False
        Me.lblremark.Enabled = False
        Me.lblphno_o.Enabled = False
        Me.lblphno_m.Enabled = False
End Sub

