VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmBillingMonitor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Billing Monitor."
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   9855
      Begin MSComDlg.CommonDialog Cd 
         Left            =   7560
         Top             =   5400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraEnv 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   1560
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CheckBox chkGrid 
            Caption         =   "Show Grid Lines"
            Height          =   255
            Left            =   3360
            TabIndex        =   11
            Top             =   1680
            Width           =   1455
         End
         Begin lvButton.lvButtons_H btnExit 
            Height          =   375
            Left            =   4920
            TabIndex        =   31
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            CapAlign        =   2
            BackStyle       =   4
            Shape           =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBhover         =   12632256
            cGradient       =   12632256
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "Billing Monitor.frx":0000
            cBack           =   16777215
         End
         Begin lvButton.lvButtons_H btnFont 
            Height          =   495
            Left            =   240
            TabIndex        =   32
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Font"
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
            cGradient       =   16777215
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "Billing Monitor.frx":0619
            ImgSize         =   32
            cBack           =   12640511
         End
         Begin lvButton.lvButtons_H btnColor 
            Height          =   495
            Left            =   1680
            TabIndex        =   33
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            Caption         =   "&Color"
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
            cGradient       =   16777215
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "Billing Monitor.frx":3A5F
            ImgSize         =   24
            cBack           =   12640511
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set Your Display Environment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   4320
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   3
            Height          =   2175
            Left            =   0
            Top             =   0
            Width           =   5415
         End
      End
      Begin VB.CheckBox chkDblClick 
         Caption         =   "Use Mouse Double Click to Set the Bill as 'Paid'"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   5520
         Width           =   3855
      End
      Begin VB.Frame fraOptions 
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   7455
         Begin VB.ComboBox cmbCust 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1200
            TabIndex        =   7
            Top             =   1680
            Width           =   1935
         End
         Begin VB.CheckBox chkOnlyOf 
            Caption         =   "Only of"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   1680
            Width           =   855
         End
         Begin VB.OptionButton optUnpaid 
            Caption         =   "Unpaid Bills Only"
            Height          =   195
            Left            =   4080
            TabIndex        =   8
            Top             =   720
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optPaid 
            Caption         =   "Paid Bills Only"
            Height          =   195
            Left            =   4080
            TabIndex        =   9
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton optAll 
            Caption         =   "All Billing Records"
            Height          =   195
            Left            =   4080
            TabIndex        =   10
            Top             =   1440
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   16449539
            CurrentDate     =   39044
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   315
            Left            =   1200
            TabIndex        =   5
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   16449539
            CurrentDate     =   39044
         End
         Begin lvButton.lvButtons_H btnShow 
            Height          =   495
            Left            =   5760
            TabIndex        =   30
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            Caption         =   "&Show"
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
            cGradient       =   16777215
            Gradient        =   3
            Mode            =   0
            Value           =   0   'False
            Image           =   "Billing Monitor.frx":4565
            ImgSize         =   24
            cBack           =   12640511
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&From Date"
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "To &Date"
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set Your Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   120
            Width           =   2490
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000C0C0&
            BorderWidth     =   3
            Height          =   2175
            Left            =   0
            Top             =   0
            Width           =   7455
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   15000
         Left            =   240
         Top             =   5160
      End
      Begin MSComctlLib.ListView lvBills 
         Height          =   4815
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   4210752
         BackColor       =   -2147483628
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cabin No"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Bill No"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Net Amt"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Customer"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin lvButton.lvButtons_H btnOptions 
         Height          =   495
         Left            =   8160
         TabIndex        =   22
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Options"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":51BA
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnEnv 
         Height          =   495
         Left            =   8160
         TabIndex        =   23
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "E&nvironment"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":59E3
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnRefresh 
         Height          =   495
         Left            =   8160
         TabIndex        =   24
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Refresh/Default"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":629F
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnPaid 
         Height          =   495
         Left            =   8160
         TabIndex        =   25
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Set As &Paid"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":71A0
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnUnpaid 
         Height          =   495
         Left            =   8160
         TabIndex        =   26
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "Set As &Unpaid"
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
         Image           =   "Billing Monitor.frx":7DF5
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnReprint 
         Height          =   495
         Left            =   8160
         TabIndex        =   27
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Re-Print Bill"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":80FD
         ImgSize         =   32
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnEdit 
         Height          =   495
         Left            =   8160
         TabIndex        =   28
         Top             =   4200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Edit Bill"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":8696
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnCredit 
         Height          =   495
         Left            =   8160
         TabIndex        =   29
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&Make Credit"
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Billing Monitor.frx":8E9D
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   9735
      End
      Begin VB.Label lblNet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9540
         TabIndex        =   18
         Top             =   5520
         Width           =   75
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   8880
      Picture         =   "Billing Monitor.frx":97A3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Monitor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   345
      TabIndex        =   13
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View, Analyze && Manipulate All Your Billing Data."
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   600
      Width           =   3420
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "frmBillingMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsBills As ADODB.Recordset

Dim payStatus As String
Dim FromDate As String
Dim ToDate As String

Private Sub btnCredit_Click()
    On Error GoTo hand
    
    Dim vBillNo As Long
    vBillNo = lvBills.SelectedItem.SubItems(2)
    
    If getBillDets(vBillNo)!Paid Then   'if bill is already paid
        MsgBox "This Bill is already paid... You can't change it to a credit bill...", vbOKOnly + vbCritical
        Exit Sub
    End If

    Dim vAcctNo As Integer
    vAcctNo = frmFind.getKey("Accounts", "AcctName", " where TypeCd in(" & custAcct & "," & suppAcct & "," & cashAcct & ")")
    If vAcctNo = -1 Then Exit Sub        'user didn't selected any account name
    
    myCn.Execute "Update BillMaster set AcctNo =" & vAcctNo & " where billno = " & vBillNo
    
    Dim t As ADODB.Recordset
    Set t = getBillDets(vBillNo)
    
    'delete previous ledger entry. may the bill is already transferred to a party account
    DeleteLedgerEntry "Sales", t!BillNo
    
    If Not vAcctNo = BI_CashAcct Then
        'debit the party account
        Debit t!Date, t!acctNo, "Credit Bill : " & t!BillNo, "Sales", t!BillNo, t!BillNo, t!NetAmt
        
        'credit the sales account
        Credit t!Date, BI_SalesAcct, "Credit Bill : " & t!BillNo, "Sales", t!BillNo, t!BillNo, t!NetAmt
    End If
    
    Call RefreshLV
    Exit Sub
hand:
    MsgBox "Please Select Any Bill/Record..."
End Sub

Private Sub btnExit_Click()
    fraEnv.Visible = False
End Sub

Private Sub btnOptions_Click()
    fraOptions.Visible = Not fraOptions.Visible
End Sub

Private Sub btnPaid_Click()
    On Error GoTo handl
    
    Dim X As ListItem
    Dim t As ADODB.Recordset
    Dim rcpt As Long
    
    For Each X In lvBills.ListItems
        If X.Selected Then
            Set t = getBillDets(X.SubItems(2))
            
            If Not t!Paid Then
                If t!acctNo = BI_CashAcct Then  'if cash then
                    'debit the cash account
                    Debit t!Date, BI_CashAcct, "Cash Bill : " & t!BillNo, "Sales", t!BillNo, t!BillNo, t!NetAmt
                    
                    'credit the sales account
                    Credit t!Date, BI_SalesAcct, "Cash Bill : " & t!BillNo, "Sales", t!BillNo, t!BillNo, t!NetAmt
                    
                    myCn.Execute "Update BillMaster set Paid = 1 where billno = " & X.SubItems(2)
                Else                            'account then
                    If MsgBox("This bill Belongs to : " & getAcctDetailsByCode(t!acctNo)!AcctName & vbCr & "Do you received cash from above account...", vbYesNo + vbQuestion) = vbYes Then
                        rcpt = newReceipt(Date, t!acctNo, t!NetAmt, t!BillNo, "Received Against Bill No. " & t!BillNo)
                        myCn.Execute "Update BillMaster set Paid = 1 where billno = " & X.SubItems(2)
                        
                        MsgBox "A Receipt is generated to credit above account..." & vbCr & "The receipt no is : " & rcpt
                    End If
                End If
            End If
        End If
    Next
    
    t.Close
    Set t = Nothing
    Call RefreshLV
    
    Exit Sub
handl:
    'MsgBox Err.Description
End Sub

Private Sub btnRefresh_Click()
    Call Defaults
    Call RefreshLV
End Sub

Private Sub btnShow_Click()
    If chkOnlyOf.Value And Trim(cmbCust) = "" Then
        MsgBox "Select Any Customer... Bacause You Checked the Only of Button"
        cmbCust.SetFocus
        Exit Sub
    End If
    
    fraOptions.Visible = False
    Call RefreshLV
End Sub

Private Sub chkGrid_Click()
    If chkGrid.Value Then
        lvBills.GridLines = True
    Else
        lvBills.GridLines = False
    End If
End Sub

Private Sub chkOnlyOf_Click()
    cmbCust.Enabled = IIf(chkOnlyOf.Value = 0, False, True)
End Sub

Private Sub dtFrom_Change()
    FromDate = Format(dtFrom, "dd-MMM-yyyy")
End Sub

Private Sub dtTo_Change()
    ToDate = Format(dtTo, "dd-MMM-yyyy")
End Sub

Private Sub Form_Load()
    Call fillCombo(cmbCust, "Accounts", "AcctName", " where TypeCd in(" & custAcct & "," & suppAcct & "," & cashAcct & ")")
    Call Defaults
    
    'aquire saved settings.
    chkDblClick.Value = GetSetting("Bar", "frmBillingMonitor", "chkDblClick.Value", 0)
    chkGrid.Value = GetSetting("Bar", "frmBillingMonitor", "chkGrid.Value", 0)
    lvBills.Font.Name = GetSetting("Bar", "frmBillingMonitor", "lvBills.Font.Name", "MS Sans Serif")
    lvBills.Font.Size = GetSetting("Bar", "frmBillingMonitor", "lvBills.Font.Size", 8)
    lvBills.Font.Bold = GetSetting("Bar", "frmBillingMonitor", "lvBills.Font.Bold", False)
    lvBills.Font.Italic = GetSetting("Bar", "frmBillingMonitor", "lvBills.Font.Italic", False)
    
    lvBills.ForeColor = GetSetting("Bar", "frmBillingMonitor", "lvBills.ForeColor", vbBlack)
    
    Call RefreshLV
End Sub

Private Sub RefreshLV()
    On Error GoTo handl
    
    Dim vBillNo As ListItem
    Dim amt As Single
    Set rsBills = New ADODB.Recordset
    
    If chkOnlyOf.Value Then
        Select Case True
            Case optUnpaid
            rsBills.Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and BillMaster.AcctNo=" & cmbCust.ItemData(cmbCust.ListIndex) & " and (not paid) and `date` between #" & dtFrom & "# and #" & dtTo & "# order by BillNo", myCn, 1, 2
            
            Case optPaid
            rsBills.Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and BillMaster.AcctNo=" & cmbCust.ItemData(cmbCust.ListIndex) & " and paid and `date` between #" & dtFrom & "# and #" & dtTo & "# order by BillNo", myCn, 1, 2
            
            Case optAll
            rsBills.Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and BillMaster.AcctNo=" & cmbCust.ItemData(cmbCust.ListIndex) & " and `date` between #" & dtFrom & "# and #" & dtTo & "# order by BillNo", myCn, 1, 2
        End Select
    Else
        Select Case True
            Case optUnpaid
            rsBills.Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and (not paid) and `date` between #" & dtFrom & "# and #" & dtTo & "# order by BillNo", myCn, 1, 2
            
            Case optPaid
            rsBills.Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and paid and `date` between #" & dtFrom & "# and #" & dtTo & "# order by BillNo", myCn, 1, 2
            
            Case optAll
            rsBills.Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and `date` between #" & dtFrom & "# and #" & dtTo & "# order by BillNo", myCn, 1, 2
        End Select
    End If
    
    lvBills.ListItems.Clear
    
    While Not rsBills.EOF
        Set vBillNo = lvBills.ListItems.Add(, , Format(rsBills!Date, "dd-MMM-yyyy"))
        vBillNo.SubItems(1) = rsBills!CustName
        vBillNo.SubItems(2) = rsBills!BillNo
        vBillNo.SubItems(3) = Format(rsBills!NetAmt, "0.00")
        vBillNo.SubItems(4) = rsBills!AcctName
        vBillNo.SubItems(5) = IIf(rsBills!Paid, "Paid", "Not Paid")
        
        amt = amt + rsBills!NetAmt
        
        rsBills.MoveNext
    Wend
    rsBills.Close
    If amt > 0 Then lblNet = "Total: " & Format(amt, "0.00") Else lblNet = ""
    
    Call Msg
    Exit Sub
handl:
    MsgBox Err.Description
End Sub

Private Sub Msg()
    If chkOnlyOf.Value Then
        lblMsg = cmbCust & "'s " & payStatus & " From " & FromDate & " to " & ToDate
    Else
        lblMsg = payStatus & " From " & FromDate & " to " & ToDate
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsBills = Nothing
    
    SaveSetting "Bar", "frmBillingMonitor", "Top", Me.Top
    SaveSetting "Bar", "frmBillingMonitor", "Left", Me.Left
    
    SaveSetting "Bar", "frmBillingMonitor", "chkDblClick.Value", chkDblClick.Value
    SaveSetting "Bar", "frmBillingMonitor", "chkGrid.Value", chkGrid.Value
    
    SaveSetting "Bar", "frmBillingMonitor", "lvBills.Font.Name", lvBills.Font.Name
    SaveSetting "Bar", "frmBillingMonitor", "lvBills.Font.Size", lvBills.Font.Size
    SaveSetting "Bar", "frmBillingMonitor", "lvBills.Font.Italic", lvBills.Font.Italic
    SaveSetting "Bar", "frmBillingMonitor", "lvBills.Font.Bold", lvBills.Font.Bold
    
    SaveSetting "Bar", "frmBillingMonitor", "lvBills.ForeColor", lvBills.ForeColor
End Sub

Private Sub btnReprint_Click()
    On Error GoTo hand
    
    Dim vBillNo As Long
    vBillNo = lvBills.SelectedItem.SubItems(2)
        
    Call initDtEnv
    DtEnv.cmdBill_Grouping vBillNo
    rptBill.Show 1
    
    Exit Sub
hand:
    'MsgBox Err.Description
End Sub

Private Sub btnUnpaid_Click()
    On Error GoTo handl
    
    Dim X  As ListItem
    Dim t As ADODB.Recordset
    
    For Each X In lvBills.ListItems
        If X.Selected Then
            Set t = getBillDets(X.SubItems(2))
            
            If t!acctNo = BI_CashAcct Then
                myCn.Execute "Update BillMaster set Paid = 0 where billno = " & X.SubItems(2)
                
                'delete records from ledger because bill is set to un-paid
                DeleteLedgerEntry "Sales", t!BillNo
            End If
        End If
    Next
    
    t.Close
    Set t = Nothing
    
    Call RefreshLV
    Exit Sub
handl:
    'MsgBox Err.Description
End Sub

Private Sub btnEdit_Click()
    On Error GoTo hand
    
    Dim vBillNo As Long
    vBillNo = lvBills.SelectedItem.SubItems(2)
    
    If getBillDets(vBillNo)!Paid Then
        MsgBox "This Bill is already paid... You can't modify this bill...", vbOKOnly + vbCritical
    Else
        frmModifySales.configure vBillNo
        
        Dim t As New ADODB.Recordset
        Set t = getBillDets(vBillNo)
        
        If Not t!acctNo = BI_CashAcct Then  'if it is a credit bill
            'delete previous ledger entry. because this bill is already transferred to a party acct
            DeleteLedgerEntry "Sales", t!BillNo
        
            'make new ledger entries with modified bill net amt
        
            'debit the party account
            Debit t!Date, t!acctNo, "Credit Bill : " & t!BillNo, "Sales", t!BillNo, t!BillNo, t!NetAmt
            
            'credit the sales account
            Credit t!Date, BI_SalesAcct, "Credit Bill : " & t!BillNo, "Sales", t!BillNo, t!BillNo, t!NetAmt
        End If
        
        t.Close
        Set t = Nothing
        
        Call RefreshLV
    End If
    
    Exit Sub
hand:
    MsgBox "Please Select Any Bill/Record..."
End Sub

Private Sub btnEnv_Click()
    fraEnv.Visible = Not fraEnv.Visible
End Sub

Private Sub btnFont_Click()
    cd.Flags = cdlCFBoth
    
    cd.FontBold = lvBills.Font.Bold
    cd.FontName = lvBills.Font.Name
    cd.FontSize = lvBills.Font.Size
    cd.FontItalic = lvBills.Font.Italic
    
    cd.ShowFont
    
    lvBills.Font.Bold = cd.FontBold
    lvBills.Font.Name = cd.FontName
    lvBills.Font.Size = cd.FontSize
    lvBills.Font.Italic = cd.FontItalic
End Sub

Private Sub btnColor_Click()
    cd.Color = lvBills.ForeColor
    cd.ShowColor
    lvBills.ForeColor = cd.Color
End Sub

Private Sub lvBills_DblClick()
    If chkDblClick.Value Then Call btnPaid_Click
End Sub

Private Sub optAll_Click()
    payStatus = "All Bills"
End Sub

Private Sub optPaid_Click()
    payStatus = "Paid Bills"
End Sub

Private Sub optUnpaid_Click()
    payStatus = "Un-Paid Bills"
End Sub

Private Sub Timer1_Timer()
    If newBills Then
        Call RefreshLV
        newBills = False
    End If
End Sub

Private Sub Defaults()
    dtFrom.Value = Date
    dtTo.Value = Date
    optUnpaid.Value = True
    
    payStatus = "Un-Paid Bills"
    FromDate = Format(dtFrom, "dd-MMM-yyyy")
    ToDate = Format(dtTo, "dd-MMM-yyyy")
    
    chkOnlyOf.Value = 1
    cmbCust.ListIndex = 0
    
    While Not cmbCust.ItemData(cmbCust.ListIndex) = 0
        cmbCust.ListIndex = cmbCust.ListIndex + 1
    Wend
End Sub
