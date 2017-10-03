VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmSales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bill Creation."
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   0
      Left            =   11520
      TabIndex        =   73
      Top             =   1080
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "0"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   1
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   2
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   3
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   4
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   5
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   6
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   7
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   8
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   9
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   10
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   11
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   12
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   13
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   14
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   15
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   16
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   17
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   18
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   19
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   20
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   21
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   22
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   23
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   24
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   25
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Index           =   26
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1695
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox chkLoose 
      BackColor       =   &H00FBE9E1&
      Caption         =   "Loose"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cmbCat 
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtNetAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   465
      Index           =   0
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtBill 
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   69861379
      CurrentDate     =   39043
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   750
      Index           =   0
      Left            =   11520
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1323
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   3
      Left            =   360
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   4
      Left            =   360
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   5
      Left            =   360
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   6
      Left            =   360
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   7
      Left            =   360
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   8
      Left            =   360
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   9
      Left            =   360
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   10
      Left            =   360
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   11
      Left            =   360
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   12
      Left            =   360
      TabIndex        =   28
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   13
      Left            =   360
      TabIndex        =   29
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   14
      Left            =   360
      TabIndex        =   30
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   15
      Left            =   360
      TabIndex        =   31
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   16
      Left            =   360
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   17
      Left            =   360
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   18
      Left            =   360
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   19
      Left            =   360
      TabIndex        =   35
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   20
      Left            =   360
      TabIndex        =   36
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   21
      Left            =   360
      TabIndex        =   37
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   22
      Left            =   360
      TabIndex        =   38
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   23
      Left            =   360
      TabIndex        =   39
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   24
      Left            =   360
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   25
      Left            =   360
      TabIndex        =   41
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   3015
      Index           =   26
      Left            =   360
      TabIndex        =   42
      Top             =   2040
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Loose"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Rate"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ProdID"
         Object.Width           =   0
      EndProperty
   End
   Begin lvButton.lvButtons_H btnSave 
      Height          =   495
      Left            =   360
      TabIndex        =   71
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Save"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      LockHover       =   2
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "Sales.frx":0000
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   1920
      TabIndex        =   72
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Dont Save"
      CapAlign        =   2
      BackStyle       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   16777215
      LockHover       =   1
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "Sales.frx":0588
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   1
      Left            =   8280
      TabIndex        =   74
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "1"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   2
      Left            =   8880
      TabIndex        =   75
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "2"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   3
      Left            =   9480
      TabIndex        =   76
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "3"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   4
      Left            =   8280
      TabIndex        =   77
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "4"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   5
      Left            =   8880
      TabIndex        =   78
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "5"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   6
      Left            =   9480
      TabIndex        =   79
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "6"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   7
      Left            =   8280
      TabIndex        =   80
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "7"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   8
      Left            =   8880
      TabIndex        =   81
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "8"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   9
      Left            =   9480
      TabIndex        =   82
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "9"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   10
      Left            =   8280
      TabIndex        =   83
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "10"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   11
      Left            =   8880
      TabIndex        =   84
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "11"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   12
      Left            =   9480
      TabIndex        =   85
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "12"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   13
      Left            =   8280
      TabIndex        =   86
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "13"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   14
      Left            =   8880
      TabIndex        =   87
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "14"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   15
      Left            =   9480
      TabIndex        =   88
      Top             =   3120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "15"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   16
      Left            =   8280
      TabIndex        =   89
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "16"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   17
      Left            =   8880
      TabIndex        =   90
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "17"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   18
      Left            =   9480
      TabIndex        =   91
      Top             =   3720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "18"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   19
      Left            =   8280
      TabIndex        =   92
      Top             =   4320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "19"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   20
      Left            =   8880
      TabIndex        =   93
      Top             =   4320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "20"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   21
      Left            =   9480
      TabIndex        =   94
      Top             =   4320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "21"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   22
      Left            =   8280
      TabIndex        =   95
      Top             =   4920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "22"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   23
      Left            =   8880
      TabIndex        =   96
      Top             =   4920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "23"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   24
      Left            =   9480
      TabIndex        =   97
      Top             =   4920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "24"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   25
      Left            =   8280
      TabIndex        =   98
      Top             =   5520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "25"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H btnCab 
      Height          =   495
      Index           =   26
      Left            =   8880
      TabIndex        =   99
      Top             =   5520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   "26"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click on Cabin Nos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8280
      TabIndex        =   70
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   300
      Left            =   5040
      TabIndex        =   69
      Top             =   5685
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Height          =   3255
      Left            =   240
      Top             =   1920
      Width           =   7695
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   195
      Left            =   7560
      TabIndex        =   16
      Top             =   1080
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   195
      Left            =   6435
      TabIndex        =   15
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Qty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   195
      Left            =   5220
      TabIndex        =   14
      Top             =   1080
      Width           =   300
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "0 th or Invisible Items"
      Height          =   195
      Left            =   11400
      TabIndex        =   3
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label lblCurrCab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   645
      TabIndex        =   2
      Top             =   300
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006A6A6A&
      Height          =   195
      Left            =   5760
      TabIndex        =   1
      Top             =   285
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Picture         =   "Sales.frx":15C1
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   10095
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As ADODB.Recordset
Dim vCurrCabin  As Integer

Dim cabButtons  As Collection
Dim LVs         As Collection
Dim NetAmts     As Collection

Private Sub btnCab_Click(Index As Integer)
    btnCab(Index).BackColor = vbCyan
    
    txtNetAmt(vCurrCabin).Visible = False
    lvBillDets(vCurrCabin).Visible = False
    
    vCurrCabin = Index
    
    txtNetAmt(vCurrCabin).Visible = True
    lvBillDets(vCurrCabin).Visible = True
    
    lblCurrCab = "Cabin/Table No : " & vCurrCabin
    
    cmbCat.SetFocus
End Sub

Private Sub btnCancel_Click()
    On Error GoTo handl
    
    If vCurrCabin = 0 Then Exit Sub
    
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Are You Sure...?", vbYesNo + vbQuestion + vbDefaultButton2)
    
    If resp = vbNo Then Exit Sub
    
    'change the button back color to origional
    btnCab(vCurrCabin).BackColor = &HD8E9EC
    
    'make all items clear
    lvBillDets(vCurrCabin).ListItems.Clear
    txtNetAmt(vCurrCabin).Text = ""
    
    'make lv and textbox of net amt invisible
    lvBillDets(vCurrCabin).Visible = False
    txtNetAmt(vCurrCabin).Visible = False
    
    vCurrCabin = 0
    lblCurrCab = ""
    
    Exit Sub
handl:
    'nothing
End Sub

Private Sub btnCancel_GotFocus()
    If Not vCurrCabin = 0 Then Merlin "Click 'Dont Save' Button to Clear This Bill..."
End Sub

Private Sub btnSave_Click()
    On Error GoTo handler
    
    If vCurrCabin = 0 Then
        MsgBox "Click on Cabin Nos button above..."
        Exit Sub
    ElseIf lvBillDets(vCurrCabin).ListItems.count = 0 Then
        MsgBox "There is nothing in this bill to save..."
        cmbCat.SetFocus
        Exit Sub
    End If
    
    Dim vBillNo As Long
    Dim X As ListItem
    
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        .Open "Select * from BillMaster order by Billno", myCn, 1, 2
        
        If .BOF Then
            vBillNo = 1
        Else
            .MoveLast
            vBillNo = !BillNo + 1
        End If
                
        .AddNew
            !BillNo = vBillNo
            !Date = dtBill.Value
            !Time = Time
            !acctNo = 0             ' 0 is acct no of 'Cash Account'
            !CustName = "Cabin : " & vCurrCabin
            !NetAmt = Val(txtNetAmt(vCurrCabin))
        .Update
        .Close
        
        .Open "Select * from BillDetail", myCn, 1, 2
        
        For Each X In lvBillDets(vCurrCabin).ListItems
            .AddNew
                !BillNo = vBillNo
                !ProdID = X.SubItems(6)
                !Qty = X.SubItems(3)
                !Loose = IIf(X.SubItems(2) = "Y", 1, 0)
                !Rate = X.SubItems(4)
                !amt = X.SubItems(5)
            .Update
            
            decreaseStock X.SubItems(6), X.SubItems(3), IIf(X.SubItems(2) = "Y", True, False)
        Next
        .Close
    End With
    
    Set rsTemp = Nothing
    'set the flag that new bills are created to be displayed on billing monitor
    newBills = True
    
    Merlin "Bill Has Been Created & Saved...", "Write"
    'MsgBox "Bill Saved." & vbCr & vbCr & "Bill No : " & vBillNo & ", Cabin : " & vCurrCabin & vbCr & "Bill Amount : " & Val(txtNetAmt(vCurrCabin))

    'change the button back color to origional
    btnCab(vCurrCabin).BackColor = &HD8E9EC
    
    'make all items clear
    lvBillDets(vCurrCabin).ListItems.Clear
    txtNetAmt(vCurrCabin).Text = ""
    
    'make lv and textbox of net amt invisible
    lvBillDets(vCurrCabin).Visible = False
    txtNetAmt(vCurrCabin).Visible = False
    
    vCurrCabin = 0
    lblCurrCab = ""
    
    Call initDtEnv
    DtEnv.cmdBill_Grouping vBillNo
    rptBill.Show 1
    
    Exit Sub
handler:
    MsgBox Err.Description
End Sub

Private Sub btnSave_GotFocus()
    If Not vCurrCabin = 0 Then Merlin "Click 'Save' Button to Save & Print This Bill..."
End Sub

Private Sub chkLoose_Click()
    On Error GoTo handler
    
    Set rsTemp = New ADODB.Recordset
    
    If chkLoose.Value Then
        rsTemp.Open "Select l" & Val(txtQty) & "MlRate from Products where ProdID = " & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
    Else
        rsTemp.Open "Select Rate from Products where ProdID = " & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
    End If
    
    txtRate = rsTemp.Fields(0)
    SendKeys "{Home}+{End}"
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    If chkLoose.Value Then
        txtAmt = txtRate
    Else
        txtAmt = Val(txtRate) * Val(txtQty)
    End If
    
    Exit Sub
handler:
        txtQty.SetFocus
End Sub

Private Sub chkLoose_GotFocus()
    Merlin "Check This 'Loose' Button if you sold the wine in loose"
End Sub

Private Sub chkLoose_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtQty.SetFocus
End Sub

Private Sub cmbCat_Click()
    On Error GoTo handl
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select Name, ProdID, SizeInMl, CatName from products, Sizes, Category where Products.SizeID=sizes.SizeId and products.catid=category.catid and CatName='" & cmbCat & "' order by products.name, sizes.sizeinml", myCn, 1, 2
    
    cmbProduct.Clear
    
    While Not rsTemp.EOF    'dont add '.' as size
        cmbProduct.AddItem IIf(rsTemp!SizeInMl = ".", rsTemp!Name, rsTemp!Name & " " & rsTemp!SizeInMl)
        cmbProduct.ItemData(cmbProduct.NewIndex) = rsTemp!ProdID
        
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    
    If cmbProduct.ListCount = 1 Then cmbProduct.ListIndex = 0
    Exit Sub
handl:
    MsgBox "Select Any Category Properly"
End Sub

Private Sub cmbCat_GotFocus()
    Merlin "Select Category And Press Enter"
End Sub

Private Sub cmbCat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbProduct.SetFocus
End Sub

Private Sub cmbCat_LostFocus()
    checkCombos cmbCat, False
End Sub

Private Sub cmbProduct_Click()
    On Error GoTo handl
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select * from Products where ProdID=" & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
    
    If rsTemp.BOF Then
        rsTemp.Close
        Exit Sub
    End If
    
    'if wine then
    If rsTemp!CatID = 3 Then
        chkLoose.Enabled = True
    Else
        chkLoose.Enabled = False
    End If
    
    txtQty = 1
    txtRate = rsTemp!Rate
    
    rsTemp.Close
    Exit Sub
handl:
    MsgBox "Select Any Item Properly..."
End Sub

Private Sub cmbProduct_GotFocus()
    Merlin "Now Select Product And Press Enter", "Pleased"
End Sub

Private Sub cmbProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkLoose.Enabled Then chkLoose.SetFocus Else txtQty.SetFocus
End Sub

Private Sub cmbProduct_LostFocus()
    checkCombos cmbProduct, False
End Sub

Private Sub Form_Load()
    dtBill.Value = Date
    
    Set cabButtons = New Collection
    Set LVs = New Collection
    Set NetAmts = New Collection

    Dim i As Integer
    Dim ctl As Control

    'add all controls to collections
    For i = 0 To 26
        cabButtons.Add btnCab(i)
        NetAmts.Add txtNetAmt(i)
        LVs.Add lvBillDets(i)
    Next

    'change back color of all the buttons
    For Each ctl In cabButtons
        ctl.BackColor = &HD8E9EC
    Next

    'make all netamt boxes bold
    For Each ctl In NetAmts
        ctl.Visible = False
    Next

    'make all Lvs inVisible
    For Each ctl In LVs
        ctl.Visible = False
        'ctl.Font.Bold = True
        'ctl.BackColor = &H80000018
    Next
    
    Call fillCombo(cmbCat, "Category", "CatName")
    Merlin "Click on Cabin/Table no Button to Start Billing...", "Pleased"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim X As lvButton.lvButtons_H
    
    For Each X In cabButtons
        If X.BackColor = vbCyan Then
            If MsgBox("Some Unsaved bills found... Do you relly want to exit ?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                Exit For
            Else
                Cancel = 1
                Exit Sub
            End If
        End If
    Next
    
    Set cabButtons = Nothing
    Set LVs = Nothing
    Set NetAmts = Nothing
    
    Set rsTemp = Nothing
    vCurrCabin = 0
    
    SaveSetting "Bar", "frmSales", "Top", Me.Top
    SaveSetting "Bar", "frmSales", "Left", Me.Left
End Sub

Private Sub lvBillDets_DblClick(Index As Integer)
    lvBillDets_KeyDown Index, 13, 1
End Sub

Private Sub lvBillDets_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        lvBillDets(Index).ListItems.Remove (lvBillDets(Index).SelectedItem.Index)
        calc Index
    ElseIf KeyCode = vbKeyReturn And Shift Then
        Dim X As ListItem
        Dim t As New ADODB.Recordset
        
        Set X = lvBillDets(Index).SelectedItem
        t.Open "Select * from Products where ProdID = " & X.SubItems(6), myCn
                
        cmbCat.ListIndex = 0
        While Not cmbCat.ItemData(cmbCat.ListIndex) = t!CatID
            cmbCat.ListIndex = cmbCat.ListIndex + 1
        Wend
        
        cmbProduct.ListIndex = 0
        While Not cmbProduct.ItemData(cmbProduct.ListIndex) = t!ProdID
            cmbProduct.ListIndex = cmbProduct.ListIndex + 1
        Wend
        
        t.Close
        
        chkLoose.Value = IIf(X.SubItems(2) = "Y", 1, 0)
        txtQty = X.SubItems(3)
        txtRate = X.SubItems(4)
        txtAmt = X.SubItems(5)
        
        lvBillDets(Index).ListItems.Remove (X.Index)
        calc Index
        
        cmbProduct.SetFocus
    End If
End Sub

Private Sub txtAmt_DblClick()
    txtAmt_KeyDown 13, 0
End Sub

Private Sub txtAmt_GotFocus()
    Merlin "Press Enter to Add This Item in Bill", "GetAttention"
End Sub

Private Sub txtAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo handl
    If vCurrCabin = 0 Then
        MsgBox "Click on Cabin Nos button above..."
        Exit Sub
    End If
    
    'if user want to add
    
    If KeyCode = 13 Then 'And Shift Then
        'check essential data
        If Trim(cmbCat) = "" Then
            MsgBox "Select Category from list..."
            cmbCat.SetFocus
            Exit Sub
        ElseIf Trim(cmbProduct) = "" Then
            MsgBox "Select Product from list..."
            cmbProduct.SetFocus
            Exit Sub
        ElseIf Val(txtAmt) = 0 Then
            MsgBox "Wrong Qty or Rate is Entered... Please Check Qty/Rate..."
            txtQty.SetFocus
            Exit Sub
        End If

        Dim X As ListItem
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select Products.Name, Sizes.SizeInMl from products, sizes where products.sizeid = sizes.sizeid and prodid=" & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
        
        With lvBillDets(vCurrCabin)
            Set X = lvBillDets(vCurrCabin).ListItems.Add(, , rsTemp!Name)
            
            X.SubItems(1) = IIf(rsTemp!SizeInMl = ".", " ", rsTemp!SizeInMl)
            X.SubItems(2) = IIf(chkLoose.Value = 1, "Y", " ")
            X.SubItems(3) = Val(txtQty)
            X.SubItems(4) = Val(txtRate)
            X.SubItems(5) = Val(txtAmt)
            X.SubItems(6) = Val(cmbProduct.ItemData(cmbProduct.ListIndex))  'store item code
            
            calc vCurrCabin
        End With
    End If
            
    If KeyCode = 13 Or KeyCode = 27 Then
        'make all text boxes and combos clear
        
        cmbCat = ""
        cmbProduct = ""
        chkLoose.Value = 0
        chkLoose.Enabled = False
        txtQty = ""
        txtRate = ""
        txtAmt = ""
        
        cmbCat.SetFocus
    End If
    Exit Sub
handl:
    MsgBox "Check Category/Item Name is properly Selected..."
End Sub

Private Sub txtQty_Change()
    If chkLoose.Value Then txtAmt = txtRate Else txtAmt = Val(txtRate) * Val(txtQty)
End Sub

Private Sub txtQty_GotFocus()
    SendKeys "{Home}+{End}"
    Merlin "Enter Sold Quantity Here", "Pleased"
End Sub

Private Sub txtRate_GotFocus()
    SendKeys "{Home}+{End}"
    Merlin "Change Rate of Selected Product if U Want", "Pleased"
    If chkLoose.Value = 0 Then Exit Sub
    
    On Error GoTo handler
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select l" & Val(txtQty) & "MlRate from Products where ProdID = " & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
    
    txtRate = rsTemp.Fields(0)
    rsTemp.Close
    SendKeys "{Home}+{End}"
    
    Exit Sub
handler:
    MsgBox "Check Qty..."
    txtQty.SetFocus
End Sub

Private Sub txtRate_Change()
    If chkLoose.Value Then txtAmt = txtRate Else txtAmt = Val(txtRate) * Val(txtQty)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRate.SetFocus Else setNumeric KeyAscii
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAmt.SetFocus Else setNumeric KeyAscii
End Sub

Private Sub calc(pCurrItem As Integer)
    Dim X As ListItem
    Dim amt As Single
    
    For Each X In lvBillDets(pCurrItem).ListItems
        amt = amt + Val(X.SubItems(5))
    Next
    
    txtNetAmt(pCurrItem) = Format(amt, "0.00")
End Sub

'("Acknowledge")
'("Announce")
'("Blink")
'("Congratulate")
'("DoMagic1")
'("DoMagic2")
'("Explain")
'("GestureDown")
'("GestureLeft")
'("GestureRight")
'("GetAttention")
'("LookUpBlink")
'("MoveDown")
'("MoveLeft")
'("MoveRight")
'("MoveUp")
'("Pleased")
'("Process")
'("Read")
'("Sad")
'("Search")
'("Show")
'("Think")
'("Wave")
'("Write")
