VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmModifySales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bill Modification."
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H00FBE9E1&
      Caption         =   "Print Bill after Saved..."
      Height          =   195
      Left            =   6000
      TabIndex        =   18
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.CheckBox chkLoose 
      BackColor       =   &H00FBE9E1&
      Caption         =   "Loose"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.ComboBox cmbCat 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   2640
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
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtBill 
      Height          =   315
      Left            =   6600
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   69337091
      CurrentDate     =   39043
   End
   Begin MSComctlLib.ListView lvBillDets 
      Height          =   2415
      Left            =   360
      TabIndex        =   13
      Top             =   3240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4260
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
   Begin lvButton.lvButtons_H btnSave 
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   5865
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
      Image           =   "Sales Modify.frx":0000
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   1920
      TabIndex        =   24
      Top             =   5865
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
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      Image           =   "Sales Modify.frx":0588
      ImgSize         =   24
      cBack           =   14737632
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modify Bill."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   240
      TabIndex        =   21
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblBillNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6600
      TabIndex        =   20
      Top             =   1800
      Width           =   1485
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Bill No"
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
      Left            =   5880
      TabIndex        =   19
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblCust 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   1560
      Width           =   3045
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
      TabIndex        =   16
      Top             =   5925
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Height          =   2655
      Left            =   240
      Top             =   3120
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
      TabIndex        =   9
      Top             =   2400
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
      TabIndex        =   7
      Top             =   2400
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
      TabIndex        =   5
      Top             =   2400
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
      TabIndex        =   2
      Top             =   2400
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
      TabIndex        =   0
      Top             =   2400
      Width           =   975
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
      TabIndex        =   14
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
      Left            =   5880
      TabIndex        =   11
      Top             =   1485
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Picture         =   "Sales Modify.frx":15C1
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   10095
   End
End
Attribute VB_Name = "frmModifySales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As ADODB.Recordset

Dim vBillNo As Long
Dim vAcctNo As Long

Private Sub btnCancel_Click()
    Dim resp As VbMsgBoxResult
    resp = MsgBox("Are You Sure...?", vbYesNo + vbQuestion + vbDefaultButton2)
    
    If resp = vbYes Then Unload Me
End Sub

Private Sub btnSave_Click()
    On Error GoTo handler
    
    If lvBillDets.ListItems.count = 0 Then
        MsgBox "There is nothing in this bill to save..."
        cmbCat.SetFocus
        Exit Sub
    End If
    
    Dim X As ListItem
    
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        .Open "Select * from BillMaster where Billno = " & vBillNo, myCn, 1, 2
        
            !Date = dtBill.Value
            !Time = Time
            !NetAmt = Val(txtNetAmt)
        .Update
        .Close
        
        'delete previous records from detail table
        .Open "Select * from BillDetail where Billno = " & vBillNo, myCn, 1, 2
            While Not .EOF
                increaseStock !ProdID, !Qty, IIf(!Loose, True, False)
                .Delete
                .MoveNext
            Wend
        .Close
        
        'add new records to detail table
        .Open "Select * from BillDetail", myCn, 1, 2
        
        For Each X In lvBillDets.ListItems
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
    
    'set the flag that bills are altered to be displayed on billing monitor
    newBills = True
    
    If chkPrint.Value Then
        Call initDtEnv
        DtEnv.cmdBill_Grouping vBillNo
        rptBill.Show 1
    End If
    
    Unload Me
    Exit Sub

handler:
    MsgBox Err.Description
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
    MsgBox "Select Any Category Properly..."
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

Private Sub cmbProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If chkLoose.Enabled Then chkLoose.SetFocus Else txtQty.SetFocus
End Sub

Private Sub cmbProduct_LostFocus()
    checkCombos cmbProduct, False
End Sub

Private Sub Form_Load()
    Call fillCombo(cmbCat, "Category", "CatName")
    
    chkPrint.Value = GetSetting("Bar", "frmModifySales", "chkPrint.Value", 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsTemp = Nothing
    
    SaveSetting "Bar", "frmModifySales", "chkPrint.Value", chkPrint.Value
End Sub

Private Sub lvBillDets_DblClick()
    lvBillDets_KeyDown 13, 1
End Sub

Private Sub lvBillDets_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        lvBillDets.ListItems.Remove (lvBillDets.SelectedItem.Index)
        calc
    ElseIf KeyCode = vbKeyReturn And Shift Then
        Dim X As ListItem
        Dim t As New ADODB.Recordset
        
        Set X = lvBillDets.SelectedItem
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
        
        lvBillDets.ListItems.Remove (X.Index)
        calc
        
        cmbProduct.SetFocus
    End If
End Sub

Private Sub txtAmt_DblClick()
    txtAmt_KeyDown 13, 0
End Sub

Private Sub txtAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo handl
    'if user want to add
    
    If KeyCode = 13 Then 'And Shift Then
        'check essential data
        If Trim(cmbCat) = "" Then
            MsgBox "Select Category..."
            cmbCat.SetFocus
            Exit Sub
        ElseIf Trim(cmbProduct) = "" Then
            MsgBox "Select Product..."
            cmbProduct.SetFocus
            Exit Sub
        ElseIf Val(txtAmt) = 0 Then
            MsgBox "Check Qty/Rate..."
            txtQty.SetFocus
            Exit Sub
        End If

        Dim X As ListItem
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select Products.Name, Sizes.SizeInMl from products, sizes where products.sizeid = sizes.sizeid and prodid=" & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
        
        With lvBillDets
            Set X = lvBillDets.ListItems.Add(, , rsTemp!Name)
            
            X.SubItems(1) = IIf(rsTemp!SizeInMl = ".", " ", rsTemp!SizeInMl)
            X.SubItems(2) = IIf(chkLoose.Value = 1, "Y", " ")
            X.SubItems(3) = Val(txtQty)
            X.SubItems(4) = Val(txtRate)
            X.SubItems(5) = Val(txtAmt)
            X.SubItems(6) = Val(cmbProduct.ItemData(cmbProduct.ListIndex))  'store item code
            calc
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
    MsgBox "Check Category/Item ..."
End Sub

Private Sub txtQty_Change()
    If chkLoose.Value Then txtAmt = txtRate Else txtAmt = Val(txtRate) * Val(txtQty)
End Sub

Private Sub txtQty_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRate_GotFocus()
    SendKeys "{Home}+{End}"
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

Private Sub calc()
    Dim X As ListItem
    Dim amt As Single
    
    For Each X In lvBillDets.ListItems
        amt = amt + Val(X.SubItems(5))
    Next
    
    txtNetAmt = Format(amt, "0.00")
End Sub

Public Sub configure(ByVal pBillNo As Long)
    vBillNo = pBillNo
    lvBillDets.ListItems.Clear
    
    Dim X As ListItem
    
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        .Open "Select BillMaster.*, Accounts.AcctName from BillMaster, Accounts where Accounts.AcctNo = BillMaster.AcctNo and BillMaster.BillNo = " & vBillNo, myCn
        
            vAcctNo = !acctNo
            lblCust = !AcctName
            lblCurrCab = !CustName
            lblBillNo = !BillNo
            dtBill = !Date
            txtNetAmt = Format(!NetAmt, "0.00")
        .Close
        
        .Open "Select BillDetail.*, Products.Name, Sizes.SizeInMl from BillDetail, Products, Sizes where Products.SizeId = Sizes.SizeId and Products.ProdID = BillDetail.ProdID and BillDetail.Billno = " & vBillNo, myCn
            
        While Not .EOF
            Set X = lvBillDets.ListItems.Add(, , !Name)
            
            X.SubItems(1) = IIf(!SizeInMl = ".", " ", !SizeInMl)
            X.SubItems(2) = IIf(!Loose, "Y", " ")
            X.SubItems(3) = !Qty
            X.SubItems(4) = !Rate
            X.SubItems(5) = !amt
            X.SubItems(6) = !ProdID
            
            .MoveNext
        Wend
        
        .Close
    End With
    
    Me.Show 1
End Sub
