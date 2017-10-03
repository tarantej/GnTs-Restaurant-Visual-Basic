VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmAccounts 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H btnEdit 
      Height          =   495
      Left            =   4320
      TabIndex        =   25
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Edit Account"
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
      Image           =   "Accounts.frx":0000
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnNew 
      Height          =   495
      Left            =   2640
      TabIndex        =   24
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&New Account"
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
      Image           =   "Accounts.frx":0807
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   4320
      TabIndex        =   22
      Top             =   5760
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
      Image           =   "Accounts.frx":1030
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin VB.ComboBox cmbDrCr 
      Height          =   315
      ItemData        =   "Accounts.frx":2069
      Left            =   4440
      List            =   "Accounts.frx":2073
      TabIndex        =   18
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtOpnBal 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   17
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   2640
      TabIndex        =   15
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox txtPermit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
   End
   Begin VB.TextBox txtMobileNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtContNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   7
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   2520
      Width           =   4215
   End
   Begin VB.TextBox txtAcctNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1695
   End
   Begin lvButton.lvButtons_H btnSave 
      Height          =   495
      Left            =   2640
      TabIndex        =   23
      Top             =   5760
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
      Image           =   "Accounts.frx":207F
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnFind 
      Height          =   495
      Left            =   6000
      TabIndex        =   26
      Top             =   5760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Find Account"
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
      Image           =   "Accounts.frx":2607
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnExit 
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   240
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
      Image           =   "Accounts.frx":59FD
      cBack           =   16777215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Opening Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   16
      Top             =   5040
      Width           =   1470
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acct &Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   4680
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Permit No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Mobile No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   10
      Top             =   3960
      Width           =   870
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   3
      Left            =   1800
      Picture         =   "Accounts.frx":6016
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   2
      Left            =   1440
      Picture         =   "Accounts.frx":95B8
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   1
      Left            =   1080
      Picture         =   "Accounts.frx":CB02
      Top             =   5880
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   0
      Left            =   720
      Picture         =   "Accounts.frx":10000
      Top             =   5880
      Width           =   285
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Profile"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   480
      Left            =   720
      TabIndex        =   20
      Top             =   1200
      Width           =   3120
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Phone No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   3240
      Width           =   330
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Na&me"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rupali Bar && Restaurant."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   660
      TabIndex        =   19
      Top             =   255
      Width           =   3420
   End
   Begin VB.Image Image1 
      Height          =   7695
      Left            =   0
      Picture         =   "Accounts.frx":13593
      Stretch         =   -1  'True
      Top             =   -645
      Width           =   9855
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim rsAcct As ADODB.Recordset
Dim rsDisp As ADODB.Recordset
Dim rsTemp As ADODB.Recordset

Dim flag As String

Private Sub btnCancel_Click()
    flag = ""                           'remove any flag assigned
    Call btnSets(False, Me)             'change the button face
    Call Disp                           'display the contents on the screen
    btnNew.SetFocus                     'set the focus to new button
End Sub

Private Sub btnFind_Click()
    Dim vAcctNo As Integer
    vAcctNo = frmFind.getKey("Accounts", "AcctName")
    
    If Not vAcctNo = -1 Then            'user has been selected a account
        rsDisp.MoveFirst                'move to the first record in displays recordset
        rsDisp.Find "AcctNo=" & vAcctNo 'find the record in displays recordset
        Call Disp                       'display the contents on the screen
    End If
End Sub

Private Sub btnNav_Click(Index As Integer)
    On Error Resume Next
    
    With rsDisp
        Select Case Index
            Case 0                          'user clicked on FirstRecord button
            .MoveFirst
            
            Case 1
            If Not .BOF Then .MovePrevious  'user clicked on PrevRecord button
            If .BOF Then .MoveFirst
            
            Case 2
            If Not .EOF Then .MoveNext      'user clicked on NextRecord button
            If .EOF Then .MoveLast
            
            Case 3                          'user clicked on LastRecord button
            .MoveLast
        End Select
    End With
    
    Call Disp
End Sub

Private Sub Disp()
    With rsDisp
        If .BOF = False And .EOF = False Then   'If there is a current record
            txtAcctNo = !acctNo                 'display all the contents
            txtName = !AcctName
            txtAddress = !Address
            txtCity = !City
            
            txtContNo = !ContNo
            txtMobileNo = !MobileNo
            txtPermit = !PermitNo
            cmbType = !TypeName
            
            txtOpnBal = !Opening
            cmbDrCr = !DrCr
        End If
                                                'show current record position
        Label12 = .AbsolutePosition & " Of " & .RecordCount
    End With
End Sub

Private Sub btnNew_Click()
    Call clearAll(Me)                       'clear contents from all text boxes and combos
    flag = "New"                            'set flag to 'New Record'
    
    Call btnSets(True, Me)                  'change the button face
    
    cmbType_LostFocus                       'set account type as customer account
    cmbDrCr.ListIndex = 0                   'set debit by default
    txtName.SetFocus                        'set focus to the acct name text box
End Sub

Private Sub btnEdit_Click()
    flag = "Edit"                           'set flag to 'Edit Record'
    
    Call btnSets(True, Me)                  'change the button face
    txtName.SetFocus                        'set focus to the acct name text box
End Sub

Private Sub btnSave_Click()
    'check essential data
    
    If Trim(txtName) = "" Then
        MsgBox "Enter Accounts Name..."
        txtName.SetFocus
        Exit Sub
    ElseIf Trim(txtAddress) = "" Then
        If MsgBox("Address Not Found... Continue ? ", vbYesNo + vbQuestion) = vbNo Then
            txtAddress.SetFocus
            Exit Sub
        End If
    ElseIf Trim(cmbType) = "" Then
        MsgBox "Select Type..."
        cmbType.SetFocus
        Exit Sub
    End If
   
    '***********************************************************
    'get Account Type Code/ID
        Dim vAccTypeID As Byte
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select TypeCd from AccTypes where TypeName='" & cmbType & "'", myCn
            
        vAccTypeID = rsTemp!TypeCd
        rsTemp.Close
    '***********************************************************
    
    Set rsAcct = New ADODB.Recordset
    With rsAcct
        If flag = "New" Then
            .Open "Select * from Accounts order by AcctNo", myCn, 1, 2
            
            'Generate New Account ID
            If .BOF Then
                txtAcctNo = 1
            Else
                .MoveLast
                txtAcctNo = !acctNo + 1
            End If
            
            'Add a new blank record to the recordset
            .AddNew
                !acctNo = Val(txtAcctNo)
        
        ElseIf flag = "Edit" Then
            .Open "Select * from Accounts Where AcctNo = " & Val(txtAcctNo), myCn, 1, 2
        End If
            !AcctName = Trim(txtName)
            !Address = Trim(txtAddress)
            !City = Trim(txtCity)
            !ContNo = Trim(txtContNo)
            !MobileNo = Trim(txtMobileNo)
            !PermitNo = Trim(txtPermit)
            !TypeCd = vAccTypeID
            !Opening = Val(txtOpnBal)
            !DrCr = Trim(cmbDrCr)
        .Update                         'Save Record
        .Close
    End With
    
    If flag = "New" Then
        'Add a entry to the ledger table
        If cmbDrCr = "Dr" Then
            Debit Date, txtAcctNo, "Opening Balance", "Opening", 0, 0, Val(txtOpnBal)
        ElseIf cmbDrCr = "Cr" Then
            Credit Date, txtAcctNo, "Opening Balance", "Opening", 0, 0, Val(txtOpnBal)
        End If
    ElseIf flag = "Edit" Then
        'Modify entry to the ledger table
        If cmbDrCr = "Dr" Then
            myCn.Execute "Update Ledger set CreditAmt = 0, DebitAmt = " & Val(txtOpnBal) & " where TranType='Opening' and AcctNo=" & Val(txtAcctNo)
        ElseIf cmbDrCr = "Cr" Then
            myCn.Execute "Update Ledger set DebitAmt = 0, CreditAmt = " & Val(txtOpnBal) & " where TranType='Opening' and AcctNo=" & Val(txtAcctNo)
        End If
    End If
    
    'Refresh Displays Recordset
    Call refreshData
    rsDisp.Find "AcctNo=" & Val(txtAcctNo)
    
    Merlin "Account Entry Saved...", "Write"
    
    flag = ""
    Call btnSets(False, Me)     'Change Button Face
End Sub

Private Sub cmbDrCr_GotFocus()
    If Not flag = "" Then Merlin "Type of Opening Balance is it a Debit Balance or Credit Balance", "Read"
End Sub

Private Sub cmbDrCr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If btnSave.Visible Then btnSave.SetFocus Else btnNew.SetFocus
End Sub

Private Sub cmbDrCr_LostFocus()
    checkCombos cmbDrCr, True
End Sub

Private Sub cmbType_GotFocus()
    If Not flag = "" Then Merlin "Select Account Type from this list.", "Think"
End Sub

Private Sub cmbType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOpnBal.SetFocus
End Sub

Private Sub cmbType_LostFocus()
    checkCombos cmbType, False
    
    If cmbType = "" Then
        cmbType.ListIndex = 0
        While Not cmbType.ItemData(cmbType.ListIndex) = custAcct
            cmbType.ListIndex = cmbType.ListIndex + 1
        Wend
    End If
End Sub

Private Sub Form_Load()
    Call fillCombo(cmbType, "AccTypes", "TypeName", " where TypeCd>1")
    Call refreshData
    Call Disp
End Sub

Private Sub refreshData()
    Set rsDisp = New ADODB.Recordset
    rsDisp.Open "Select Accounts.*, AccTypes.TypeName from Accounts, AccTypes where Accounts.TypeCd = AccTypes.TypeCd and Accounts.AcctNo > 1 order by AcctNo", myCn, 1, 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsAcct = Nothing
    flag = ""
End Sub

Private Sub rsDisp_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Label12 = "Record : " & rsDisp.AbsolutePosition
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub txtAddress_GotFocus()
    If Not flag = "" Then Merlin "Enter " & Trim(txtName) & "'s Address Here..."
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCity.SetFocus
    Else
        tCase txtAddress, KeyAscii
    End If
End Sub

Private Sub txtCity_GotFocus()
    If Not flag = "" Then Merlin "Enter City of Account/Party Here", "Acknowledge"
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtContNo.SetFocus
    Else
        tCase txtCity, KeyAscii
    End If
End Sub

Private Sub txtContNo_GotFocus()
    If Not flag = "" Then Merlin "Phone Number of Party."
End Sub

Private Sub txtContNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMobileNo.SetFocus
    Else
        tCase txtContNo, KeyAscii
    End If
End Sub

Private Sub txtMobileNo_GotFocus()
    If Not flag = "" Then Merlin "Mobile Number of Party.", "Acknowledge"
End Sub

Private Sub txtMobileNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPermit.SetFocus
    Else
        tCase txtMobileNo, KeyAscii
    End If
End Sub

Private Sub txtName_GotFocus()
    If Not flag = "" Then Merlin "Enter Account/Party Name here...", "DoMagic1"
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAddress.SetFocus
    Else
        tCase txtName, KeyAscii
    End If
End Sub

Private Sub txtOpnBal_GotFocus()
    If Not flag = "" Then Merlin "Enter Opening Balance of Account/Party Here."
End Sub

Private Sub txtOpnBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbDrCr.SetFocus
    Else
        setNumeric KeyAscii
    End If
End Sub

Private Sub txtPermit_GotFocus()
    If Not flag = "" Then Merlin "Permit Number of Party in Case of A Customer."
End Sub

Private Sub txtPermit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbType.SetFocus
    Else
        tCase txtPermit, KeyAscii
    End If
End Sub
