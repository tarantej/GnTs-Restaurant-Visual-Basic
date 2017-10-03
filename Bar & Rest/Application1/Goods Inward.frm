VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmInward 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Goods Inward."
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   7215
      Begin lvButton.lvButtons_H btnNew 
         Height          =   495
         Left            =   5520
         TabIndex        =   19
         Top             =   240
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
         Image           =   "Goods Inward.frx":0000
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.TextBox txtRemark 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   4320
         Width           =   5895
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4080
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbProduct 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   1080
         Width           =   2175
      End
      Begin VB.ComboBox cmbCat 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtEntrySrNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtInward 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16449539
         CurrentDate     =   39043
      End
      Begin MSComctlLib.ListView lvBillDets 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4895
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Particulars of Item"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty Inwarded"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ProdID"
            Object.Width           =   0
         EndProperty
      End
      Begin lvButton.lvButtons_H btnCancel 
         Height          =   495
         Left            =   5520
         TabIndex        =   17
         Top             =   840
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
         Image           =   "Goods Inward.frx":0829
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnSave 
         Height          =   495
         Left            =   5520
         TabIndex        =   18
         Top             =   240
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
         Image           =   "Goods Inward.frx":1862
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Remark/Note"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Qty"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4920
         TabIndex        =   6
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Product"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Category"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entry No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Material You Received on Counter."
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   6360
      Picture         =   "Goods Inward.frx":1DEA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Materials Inward Entry."
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
      TabIndex        =   12
      Top             =   240
      Width           =   1965
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmInward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTemp As ADODB.Recordset

Private Sub btnCancel_Click()
    Call btnSets(False, Me)
    Merlin "This Entry Has not Saved..."
End Sub

Private Sub btnNew_Click()
    txtEntrySrNo = ""
    cmbCat = ""
    cmbProduct = ""
    txtQty = ""
    txtRemark = ""
    
    lvBillDets.ListItems.Clear
    
    dtInward.Value = Date
    Call btnSets(True, Me)
    cmbCat.SetFocus
End Sub

Private Sub btnSave_Click()
    If lvBillDets.ListItems.count = 0 Then
        MsgBox "There is nothing to save in this entry...", vbOKOnly + vbExclamation
        cmbCat.SetFocus
        Exit Sub
    End If
    
    Dim vEntryNo As Long
    Dim X As ListItem
    
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        .Open "Select * from InwardMaster order by EntrySrNo", myCn, 1, 2
        
        If .BOF Then
            vEntryNo = 1
        Else
            .MoveLast
            vEntryNo = !EntrySrNo + 1
        End If
        
        txtEntrySrNo = vEntryNo
                
        .AddNew
            !EntrySrNo = vEntryNo
            !Date = dtInward.Value
            !Time = Time
            !Remark = Trim(txtRemark)
        .Update
        .Close
        
        .Open "Select * from InwardDetail", myCn, 1, 2
        
        For Each X In lvBillDets.ListItems
            .AddNew
                !EntrySrNo = vEntryNo
                !ProdID = X.SubItems(3)
                !Qty = X.SubItems(2)
            .Update
            
            increaseStock X.SubItems(3), X.SubItems(2), False
        Next
        .Close
    End With
    
    Merlin "This Entry Has Been Saved...", "Write"
    
    Set rsTemp = Nothing
    Call btnSets(False, Me)
End Sub

Private Sub cmbCat_GotFocus()
    Merlin "Select Desired Product Category..."
End Sub

Private Sub cmbProduct_GotFocus()
    Merlin "Select Desired Product...", "Pleased"
End Sub

Private Sub lvBillDets_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        lvBillDets.ListItems.Remove (lvBillDets.SelectedItem.Index)
    End If
End Sub

Private Sub cmbCat_Click()
    On Error GoTo handl
    If btnNew.Visible Then
        MsgBox "Please Click on 'New' Button First..."
        btnNew.SetFocus
        Exit Sub
    End If
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select Name, ProdID, SizeInMl, CatName from products, Sizes, Category where Products.SizeID=sizes.SizeId and products.catid=category.catid and CatName='" & cmbCat & "' order by products.name, sizes.sizeinml", myCn, 1, 2
    
    cmbProduct.Clear
    
    While Not rsTemp.EOF    'dont add '.' as size
        cmbProduct.AddItem IIf(rsTemp!SizeInMl = ".", rsTemp!Name, rsTemp!Name & " " & rsTemp!SizeInMl)
        cmbProduct.ItemData(cmbProduct.NewIndex) = rsTemp!ProdID
        
        rsTemp.MoveNext
    Wend
    rsTemp.Close
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

Private Sub cmbProduct_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtQty.SetFocus
End Sub

Private Sub cmbProduct_LostFocus()
    checkCombos cmbProduct, False
End Sub

Private Sub Form_Load()
    Call fillCombo(cmbCat, "Category", "CatName", " where CatName<>'Food'")
    Merlin "Here You Can Make Entries of Inwarded Goods on Your Sale Counter, Click on 'New Entry' Button to Start Entry", "Pleased"
End Sub

Private Sub txtQty_DblClick()
    Call txtQty_KeyDown(13, 0)
End Sub

Private Sub txtQty_GotFocus()
    Merlin "Enter Inwarded Qty Of " & Trim(cmbProduct) & " And Press Enter..."
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo handl
    
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
        ElseIf Val(txtQty) = 0 Then
            MsgBox "Enter Qty..."
            txtQty.SetFocus
            Exit Sub
        End If

        Dim X As ListItem
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select Products.Name, Sizes.SizeInMl from products, sizes where products.sizeid = sizes.sizeid and prodid=" & cmbProduct.ItemData(cmbProduct.ListIndex), myCn
        
        With lvBillDets
            Set X = lvBillDets.ListItems.Add(, , rsTemp!Name)
            
            X.SubItems(1) = IIf(rsTemp!SizeInMl = ".", " ", rsTemp!SizeInMl)
            X.SubItems(2) = Val(txtQty)
            X.SubItems(3) = Val(cmbProduct.ItemData(cmbProduct.ListIndex))  'store item code
        End With
    End If
            
    If KeyCode = 13 Or KeyCode = 27 Then
        'make all text boxes and combos clear
        
        cmbCat = ""
        cmbProduct = ""
        txtQty = ""
        
        cmbCat.SetFocus
    End If
    Exit Sub
handl:
    MsgBox "Check Category/Item ..."
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    setNumeric KeyAscii
End Sub

Private Sub txtRemark_GotFocus()
    Merlin "Add Any Remark/Narration to this entry..."
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If btnSave.Visible Then btnSave.SetFocus Else btnNew.SetFocus
End Sub
