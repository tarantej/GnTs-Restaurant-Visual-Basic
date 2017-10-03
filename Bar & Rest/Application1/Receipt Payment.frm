VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmVoucher 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voucher."
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   8055
      Begin VB.TextBox txtRemark 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   2400
         Width           =   6015
      End
      Begin VB.TextBox txtBillBalance 
         BackColor       =   &H8000000A&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtAcctBal 
         BackColor       =   &H8000000A&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtAmt 
         Height          =   315
         Left            =   6240
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cmbAgainst 
         Height          =   315
         Left            =   6240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtCity 
         BackColor       =   &H8000000A&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H8000000A&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cmbAcct 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtEntry 
         Height          =   315
         Left            =   6240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16449539
         CurrentDate     =   39060
      End
      Begin VB.TextBox txtVchNo 
         BackColor       =   &H8000000A&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin lvButton.lvButtons_H btnNew 
         Height          =   495
         Left            =   6240
         TabIndex        =   22
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Caption         =   "&New Entry"
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
         Image           =   "Receipt Payment.frx":0000
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnCancel 
         Height          =   495
         Left            =   6240
         TabIndex        =   23
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
         cGradient       =   16777215
         Gradient        =   3
         Mode            =   0
         Value           =   0   'False
         Image           =   "Receipt Payment.frx":0829
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnSave 
         Height          =   495
         Left            =   4560
         TabIndex        =   24
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
         Image           =   "Receipt Payment.frx":1862
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Note / Narration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   1410
      End
      Begin VB.Label lblVchType 
         AutoSize        =   -1  'True
         Caption         =   "Bill Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   4800
         TabIndex        =   20
         Top             =   1200
         Width           =   1020
      End
      Begin VB.Label lblVchType 
         AutoSize        =   -1  'True
         Caption         =   "Acct's Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label lblAmtType 
         AutoSize        =   -1  'True
         Caption         =   "Recd Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label lblVchType 
         AutoSize        =   -1  'True
         Caption         =   "Against"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   4
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Party/Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   4800
         TabIndex        =   0
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblVchType 
         AutoSize        =   -1  'True
         Caption         =   "Receipt No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lblMain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   11
      Top             =   240
      Width           =   75
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   7320
      Picture         =   "Receipt Payment.frx":1DEA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblSub 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "frmVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As String
Dim rsTemp As ADODB.Recordset

Public Sub Init(pVchType As String)
    lblMain = pVchType
    lblVchType(0) = pVchType & " No."
    
    Select Case pVchType
        Case "Receipt"
        lblSub = "Make Entry of Received Cash..."
        lblAmtType = "Recd Amount"
        
        Case "Payment"
        lblSub = "Make Entry of Paid/Given Cash..."
        lblAmtType = "Paid Amount"
    End Select
    
    ShowInCentre Me
End Sub

Private Sub btnCancel_Click()
    Call clearAll(Me)
    Call btnSets(False, Me)
    btnNew.SetFocus
End Sub

Private Sub btnNew_Click()
    Call clearAll(Me)
    Call btnSets(True, Me)
    flag = "New"
    dtEntry = Date
    
    cmbAcct.SetFocus
    If lblMain = "Receipt" Then
        txtRemark = "Cash Recd"
    Else
        txtRemark = "Cash Paid"
    End If
End Sub

Private Sub btnSave_Click()
    'check essential data
    
    If cmbAcct = "" Then
        MsgBox "Select Account..."
        cmbAcct.SetFocus
        Exit Sub
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Enter Amount..."
        txtAmt.SetFocus
        Exit Sub
    ElseIf cmbAgainst <> "On Account" And Val(txtAmt) > Val(txtBillBalance) Then
        MsgBox "Bills Balance is only : " & txtBillBalance
        txtAmt.SetFocus
        Exit Sub
    End If
    
    If flag = "New" Then
        If lblMain = "Receipt" Then
            txtVchNo = newReceipt(dtEntry, cmbAcct.ItemData(cmbAcct.ListIndex), Val(txtAmt), IIf(cmbAgainst = "On Account", cmbAgainst, cmbAgainst.ItemData(cmbAgainst.ListIndex)), txtRemark)
        Else
            txtVchNo = newPayment(dtEntry, cmbAcct.ItemData(cmbAcct.ListIndex), Val(txtAmt), , txtRemark)
        End If
    ElseIf flag = "Edit" Then
        'edit code goes here
    
    End If
    
    Merlin "Entry Has Been Saved...", "Write"
    
    Call btnSets(False, Me)
End Sub

Private Sub cmbAcct_Click()
    If cmbAcct = "" Then Exit Sub
    
    Set rsTemp = getAcctDetailsByCode(cmbAcct.ItemData(cmbAcct.ListIndex))
    txtAddress = rsTemp!Address
    txtCity = rsTemp!City
    rsTemp.Close
    
    txtAcctBal = Format(getAcctBalance(cmbAcct.ItemData(cmbAcct.ListIndex)), "0.00")
    Call fillBills(cmbAcct.ItemData(cmbAcct.ListIndex))
End Sub

Private Sub cmbAcct_GotFocus()
    If lblMain = "Receipt" Then
        Merlin "Select Account/Party Name from which you have received cash"
    Else
        Merlin "Select Account/Party Name to which you have Paid cash"
    End If
End Sub

Private Sub cmbAcct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbAgainst.SetFocus
End Sub

Private Sub cmbAcct_LostFocus()
    checkCombos cmbAcct, False
End Sub

Private Sub cmbAgainst_Click()
    txtBillBalance = 0
    If Not cmbAgainst = "On Account" Then If lblMain = "Receipt" Then txtBillBalance = Format(getCrBillBalance(cmbAgainst.ItemData(cmbAgainst.ListIndex)), "0.00")
End Sub

Private Sub cmbAgainst_GotFocus()
    Merlin "Against which Bill...?", "Pleased"
End Sub

Private Sub cmbAgainst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAmt.SetFocus
End Sub

Private Sub cmbAgainst_LostFocus()
    checkCombos cmbAgainst, True
End Sub

Private Sub Form_Load()
    Call fillCombo(cmbAcct, "Accounts", "AcctName", " where AcctNo>1")
End Sub

Private Sub fillBills(ByVal pAcctNo As Long)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    cmbAgainst.Clear
    cmbAgainst.AddItem "On Account"
    
    If lblMain = "Receipt" Then
        rs.Open "Select Billno, Date from BillMaster where AcctNo=" & pAcctNo, myCn
        
        While Not rs.EOF
            If lblMain = "Receipt" Then
                'Send Uniq EntrySrNo to this function
                If getCrBillBalance(rs.Fields("BillNo")) > 0 Then
                    cmbAgainst.AddItem rs!BillNo & " - " & Format(rs!Date, "dd-MMM-yyyy")
                    cmbAgainst.ItemData(cmbAgainst.NewIndex) = rs!BillNo
                End If
            End If
            
            rs.MoveNext
        Wend
        rs.Close
    End If
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRemark.SetFocus Else setNumeric KeyAscii
End Sub

Private Sub txtAmt_GotFocus()
    If lblMain = "Receipt" Then
        Merlin "How Many Rupees you have Received from " & Trim(cmbAcct), "Pleased"
    Else
        Merlin "How Many Rupees you have Paid to " & Trim(cmbAcct), "Pleased"
    End If
End Sub

Private Sub txtRemark_GotFocus()
    Merlin "Any Remark you want to store with this transaction..."
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If btnSave.Visible Then btnSave.SetFocus Else btnNew.SetFocus Else tCase txtRemark, KeyAscii
End Sub
