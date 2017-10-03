VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLedger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ledger."
   ClientHeight    =   6915
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txtbalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   6240
         Width           =   1830
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   4920
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtDebit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   6015
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   6240
         Width           =   1395
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   7455
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   6240
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid grdDisplay 
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   8705
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbAcctName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   6480
         TabIndex        =   3
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16449539
         CurrentDate     =   38078
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   6480
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16449539
         CurrentDate     =   38280
      End
      Begin lvButton.lvButtons_H btnShow 
         Height          =   495
         Left            =   8955
         TabIndex        =   6
         Top             =   255
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "&Show/Refresh"
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
         Image           =   "Ledger Disp.frx":0000
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnPrint 
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Print"
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
         Image           =   "Ledger Disp.frx":0829
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnExit 
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   6240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "&Close"
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
         Image           =   "Ledger Disp.frx":0DC2
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnFont 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   6240
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
         Image           =   "Ledger Disp.frx":1DFB
         ImgSize         =   32
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnColor 
         Height          =   495
         Left            =   4440
         TabIndex        =   11
         Top             =   6240
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
         Image           =   "Ledger Disp.frx":5241
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ledger &Of."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " &To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5520
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " F&rom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCity 
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label lblArea 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsLedger As ADODB.Recordset
Dim acctNo As Integer

Private Sub ShowLedger()
    If Trim(cmbAcctName) = Empty Then Exit Sub

    acctNo = cmbAcctName.ItemData(cmbAcctName.ListIndex)
    prepareLedger acctNo
    Call configure
End Sub

Private Sub configure()
    Set rsLedger = New ADODB.Recordset
    rsLedger.Open "SELECT Ledger.srNo, Accounts.AcctName, Ledger.TranDate, Ledger.Particulars, Ledger.TranType, Ledger.DebitAmt, Ledger.CreditAmt, Ledger.ClosingBal, Ledger.PrintNo,Ledger.EntryNo, iif(Ledger.ClosingBal > 0, 'Dr', 'Cr') AS DrCr FROM Accounts, Ledger WHERE Accounts.AcctNo = Ledger.AcctNo AND Ledger.AcctNo =" & acctNo & " and Ledger.TranDate between #" & dtFrom & "# and #" & dtTo & "# ORDER BY Ledger.srNo", myCn
    Call display
End Sub

Private Sub display()
    Call setGridStatus
    Dim X As Integer

    txtDebit = 0
    txtCredit = 0
    
    With grdDisplay
        While Not rsLedger.EOF
            X = X + 1
            .Rows = .Rows + 1
            
            .TextMatrix(X, 0) = rsLedger.Fields("SrNo")
            .TextMatrix(X, 1) = Format(rsLedger.Fields("TranDate"), "dd-MMM-yyyy")
            .TextMatrix(X, 2) = rsLedger.Fields("Particulars")
            .TextMatrix(X, 3) = rsLedger.Fields("TranType")
            .TextMatrix(X, 4) = IIf(rsLedger.Fields("PrintNo") <> Empty, rsLedger.Fields("PrintNo"), 0)
            .TextMatrix(X, 5) = IIf(rsLedger.Fields("DebitAmt") <> 0, Format(rsLedger.Fields("DebitAmt"), "0.00"), "-")
            .TextMatrix(X, 6) = IIf(rsLedger.Fields("CreditAmt") <> 0, Format(rsLedger.Fields("CreditAmt"), "0.00"), "-")
            .TextMatrix(X, 7) = Format(Abs(rsLedger.Fields("ClosingBal")), "0.00")
            .TextMatrix(X, 8) = rsLedger.Fields("DrCr")
            .TextMatrix(X, 9) = rsLedger.Fields("EntryNo")
            
            txtDebit = Val(txtDebit) + rsLedger.Fields("DebitAmt")
            txtCredit = Val(txtCredit) + rsLedger.Fields("CreditAmt")
            
            rsLedger.MoveNext
        Wend
        rsLedger.Close
        
        txtDebit = Format(Val(txtDebit), "0.00")
        txtCredit = Format(Val(txtCredit), "0.00")
    End With
    
    txtbalance = Format(Val(txtDebit) - Val(txtCredit), "0.00")
    txtbalance = IIf(Val(txtbalance) > 0, txtbalance & " Dr", Format(Abs(txtbalance), "0.00") & " Cr")
End Sub

Private Sub setGridStatus()
    With grdDisplay
        .Clear
        .Rows = 1
        .Cols = 10
    
        .TextMatrix(0, 0) = "Sr"
        .ColWidth(0) = 500
        
        .TextMatrix(0, 1) = "Date"
        .ColWidth(1) = 1250
        
        .TextMatrix(0, 2) = "Particulars"
        .ColWidth(2) = 2650
        
        .TextMatrix(0, 3) = "Vch Type"
        .ColWidth(3) = 975
        .ColAlignment(3) = vbAlignLeft
        
        .TextMatrix(0, 4) = "Vch No"
        .ColWidth(4) = 900
        .ColAlignment(4) = vbAlignLeft
        
        .TextMatrix(0, 5) = "Debit Amt"
        .ColWidth(5) = 1100
        
        .TextMatrix(0, 6) = "Credit Amt"
        .ColWidth(6) = 1100
        
        .TextMatrix(0, 7) = "Closing"
        .ColWidth(7) = 1200
        
        .TextMatrix(0, 8) = "Type"
        .ColWidth(8) = 550
        
        .TextMatrix(0, 9) = "EntryNo"
        .ColWidth(9) = 0
    End With
End Sub

Private Sub btnShow_GotFocus()
    If Trim(cmbAcctName) = "" Then
        Merlin "You First Select Account/Party Name from List Beside", "Wave"
    Else
        Merlin "Press Enter Now to see ledger of " & Trim(cmbAcctName), "Pleased"
    End If
End Sub

Private Sub cmbAcctName_Click()
    If cmbAcctName = Empty Then Exit Sub
    
    Set rsLedger = getAcctDetailsByCode(cmbAcctName.ItemData(cmbAcctName.ListIndex))
    lblArea = rsLedger.Fields("Address")
    lblCity = rsLedger.Fields("City")
    
    Set rsLedger = Nothing
End Sub

Private Sub cmbAcctName_GotFocus()
    Merlin "Select Account/Party Name to Show It's Ledger"
End Sub

Private Sub cmbAcctName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtFrom.SetFocus
End Sub

Private Sub cmbAcctName_LostFocus()
    checkCombos cmbAcctName, False
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnFont_Click()
    Cd.Flags = cdlCFBoth
    
    Cd.FontName = grdDisplay.Font.Name
    Cd.FontSize = grdDisplay.Font.Size
    Cd.FontBold = grdDisplay.Font.Bold
    Cd.FontItalic = grdDisplay.Font.Italic
    
    Cd.ShowFont
    
    grdDisplay.Font.Name = Cd.FontName
    grdDisplay.Font.Size = Cd.FontSize
    grdDisplay.Font.Bold = Cd.FontBold
    grdDisplay.Font.Italic = Cd.FontItalic
End Sub

Private Sub dtFrom_GotFocus()
    If Trim(cmbAcctName) = "" Then
        Merlin "You First Select Account/Party Name from List Beside", "Wave"
    Else
        Merlin "Select Start Date here...", "Pleased"
    End If
End Sub

Private Sub dtTo_GotFocus()
    If Trim(cmbAcctName) = "" Then
        Merlin "You First Select Account/Party Name from List Beside", "Wave"
    Else
        Merlin "Select Stop/End Date here...", "Pleased"
    End If
End Sub

Private Sub Form_Load()
    fillCombo cmbAcctName, "Accounts", "AcctName"
    
    dtFrom = "01-Apr-2006"
    dtTo = Date
    
    grdDisplay.Font.Name = GetSetting("Bar", "frmLedger", "GrdDisplay.Font.Name", "Arial")
    grdDisplay.Font.Size = GetSetting("Bar", "frmLedger", "GrdDisplay.Font.size", 10)
    grdDisplay.Font.Bold = GetSetting("Bar", "frmLedger", "GrdDisplay.Font.Bold", False)
    grdDisplay.Font.Italic = GetSetting("Bar", "frmLedger", "GrdDisplay.Font.Italic", False)
    
    grdDisplay.BackColor = GetSetting("Bar", "frmLedger", "grdDisplay.BackColor", vbWhite)
End Sub

Private Sub btnShow_Click()
    If Trim(cmbAcctName) = "" Then
        Merlin "You First Select Account/Party Name from List Beside", "Wave"
    Else
        Call ShowLedger
        Merlin cmbAcctName & " Has a Balance of Rupees : " & txtbalance, "Read"
        grdDisplay.SetFocus
    End If
End Sub

Private Sub btnPrint_Click()
    If Trim(cmbAcctName) = Empty Then Exit Sub
    Dim acctNo As Integer
    
    Call initDtEnv
    acctNo = getAcctDetailsByCode(cmbAcctName.ItemData(cmbAcctName.ListIndex))!acctNo
    'DtEnv.cmdLedger_Grouping acctNo, dtFrom, dtTo
    
    'rptHeads rptLedger
    'rptLedger.Show
End Sub

Private Sub btnColor_Click()
    Cd.Color = grdDisplay.BackColor
    Cd.ShowColor
    grdDisplay.BackColor = Cd.Color
End Sub

Private Sub dtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtTo.SetFocus
End Sub

Private Sub dtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then btnShow.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Bar", "frmLedger", "GrdDisplay.Font.Name", grdDisplay.Font.Name
    SaveSetting "Bar", "frmLedger", "GrdDisplay.Font.Size", grdDisplay.Font.Size
    SaveSetting "Bar", "frmLedger", "GrdDisplay.Font.Bold", grdDisplay.Font.Bold
    SaveSetting "Bar", "frmLedger", "GrdDisplay.Font.Italic", grdDisplay.Font.Italic

    SaveSetting "Bar", "frmLedger", "grdDisplay.BackColor", grdDisplay.BackColor
End Sub

Private Sub grdDisplay_DblClick()
    Call editEntry
End Sub

Private Sub grdDisplay_KeyPress(KeyAscii As Integer)
    If Not KeyAscii = 13 Then Exit Sub
    Call editEntry
End Sub

Private Sub editEntry()
    Dim objectform As Form
    Dim entryNo As Long
    entryNo = grdDisplay.TextMatrix(grdDisplay.Row, 9)
    
'    Select Case grdDisplay.TextMatrix(grdDisplay.Row, 3)
'        Case "Sales"
'            frmModifySales.configure entryNo
'
'        Case "Pur"
'            Set objectform = New frmPurchase
'            objectform.Show
'            objectform.configureEdit entryNo
'
'
'        Case "Receipt", "Payment", "Contra", "Journal"
'            Load frmVoucher
'            frmVoucher.Show
'            frmVoucher.configureEdit entryNo
'    End Select
End Sub
