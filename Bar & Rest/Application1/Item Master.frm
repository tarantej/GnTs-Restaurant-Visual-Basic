VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmProducts 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt30MlRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txt60MlRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   15
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox txt90MlRate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtOpnLoose 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtOpn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox cmbSize 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox cmbCategory 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtProductName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
   End
   Begin VB.TextBox txtProdID 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1575
   End
   Begin lvButton.lvButtons_H btnNew 
      Height          =   495
      Left            =   3240
      TabIndex        =   23
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&New Product"
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
      Image           =   "Item Master.frx":0000
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnEdit 
      Height          =   495
      Left            =   4800
      TabIndex        =   24
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Edit Product"
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
      Image           =   "Item Master.frx":0829
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
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
      Image           =   "Item Master.frx":1030
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnSave 
      Height          =   495
      Left            =   3240
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
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
      Image           =   "Item Master.frx":2069
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnFind 
      Height          =   495
      Left            =   6360
      TabIndex        =   27
      Top             =   5640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Caption         =   "&Find Product"
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
      Image           =   "Item Master.frx":25F1
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnExit 
      Height          =   375
      Left            =   7440
      TabIndex        =   28
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
      Image           =   "Item Master.frx":59E7
      cBack           =   16777215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rates of Loose Sale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   3360
      Width           =   1920
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30 ML"
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
      Left            =   5520
      TabIndex        =   12
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60 ML"
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
      Left            =   5520
      TabIndex        =   14
      Top             =   4200
      Width           =   540
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "90 Ml"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   3
      Left            =   1800
      Picture         =   "Item Master.frx":6000
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   2
      Left            =   1440
      Picture         =   "Item Master.frx":95A2
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   1
      Left            =   1080
      Picture         =   "Item Master.frx":CAEC
      Top             =   5160
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   0
      Left            =   720
      Picture         =   "Item Master.frx":FFEA
      Top             =   5160
      Width           =   285
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Frofile"
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
      Left            =   720
      TabIndex        =   21
      Top             =   1440
      Width           =   3000
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock &Loose (MLs)"
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
      Top             =   4560
      Width           =   2370
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Sto&ck"
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
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product &Rate"
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
      Top             =   3840
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Si&ze (if Wine/Beer/CD)"
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
      Top             =   3480
      Width           =   1980
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ca&tegory"
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
      Top             =   3120
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Na&me && Description"
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
      Top             =   2760
      Width           =   1680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
      TabIndex        =   18
      Top             =   2400
      Width           =   930
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
      TabIndex        =   20
      Top             =   255
      Width           =   3420
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   -30
      Picture         =   "Item Master.frx":1357D
      Stretch         =   -1  'True
      Top             =   -645
      Width           =   9855
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim rsDisp As ADODB.Recordset
Dim rsProd As ADODB.Recordset

Dim flag As String

Private Sub btnCancel_Click()
    flag = ""
    Call btnSets(False, Me)
    Call Disp
    btnNew.SetFocus
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnFind_Click()
    Dim vProdID As Integer
    vProdID = frmFind.getKey("Products", "Name")
    
    If Not vProdID = -1 Then            'user has been selected a account
        rsDisp.MoveFirst                'move to the first record in displays recordset
        rsDisp.Find "ProdID=" & vProdID 'find the record in displays recordset
        Call Disp                       'display the contents on the screen
    End If
End Sub

Private Sub btnNav_Click(Index As Integer)
    On Error Resume Next
    
    With rsDisp
        Select Case Index
            Case 0
            .MoveFirst
            
            Case 1
            If Not .BOF Then .MovePrevious
            
            Case 2
            If Not .EOF Then .MoveNext
            
            Case 3
            .MoveLast
        End Select
    End With
    
    Call Disp
End Sub

Private Sub Disp()
    With rsDisp
        If .BOF = False And .EOF = False Then
            txtProdID = !ProdID
            txtProductName = !Name
            cmbCategory = !CatName
            cmbSize = !SizeInMl
            txtRate = !Rate
            
            txt30MlRate = !l30MlRate
            txt60MlRate = !l60MlRate
            txt90MlRate = !l90MlRate
            
            txtOpn = !OpenStock
            txtOpnLoose = !OpenStockLoose
        End If
    End With
End Sub

Private Sub btnNew_Click()
    Call clearAll(Me)
    flag = "New"
    
    Call btnSets(True, Me)
    txtProductName.SetFocus
End Sub

Private Sub btnEdit_Click()
    flag = "Edit"
    
    Call btnSets(True, Me)
    Merlin "You Are Modifying " & Trim(txtProductName), "GetAttention"
    
    txtProductName.SetFocus
End Sub

Private Sub btnSave_Click()
    'check essential data
    
    If Trim(txtProductName) = "" Then
        MsgBox "Enter Item/Product Name..."
        txtProductName.SetFocus
        Exit Sub
    ElseIf Trim(cmbCategory) = "" Then
        MsgBox "Select Proper Category..."
        cmbCategory.SetFocus
        Exit Sub
    End If
    
    Dim vCatId, vSizeID As Integer
    vCatId = getCategoryDets(cmbCategory)!CatID
    
    If Trim(cmbSize) = "" Then vSizeID = 0 Else vSizeID = getSizeDets(cmbSize)!SizeID
    
    Set rsProd = New ADODB.Recordset
    With rsProd
        If flag = "New" Then
            .Open "Select * from Products order by ProdID", myCn, 1, 2
            
            If .BOF Then
                txtProdID = 1
            Else
                .MoveLast
                txtProdID = !ProdID + 1
            End If
            
            .AddNew
                !ProdID = Val(txtProdID)
                    'make current stock same as opening stock
                    !CurrStock = Val(txtOpn)
                    !CurrStockLoose = Val(txtOpnLoose)
        
        ElseIf flag = "Edit" Then
            .Open "Select * from Products Where ProdID = " & Val(txtProdID), myCn, 1, 2
        End If
                !Name = Trim(txtProductName)
                !SizeID = vSizeID
                !CatID = vCatId
                !Rate = Val(txtRate)
                
                !l30MlRate = Val(txt30MlRate)
                !l60MlRate = Val(txt60MlRate)
                !l90MlRate = Val(txt90MlRate)
                
                !OpenStock = Val(txtOpn)
                !OpenStockLoose = Val(txtOpnLoose)
            .Update
            .Close
    End With
    
    Set rsProd = Nothing
    
    Call btnSets(False, Me)
    Call refreshData
    
    flag = ""
    
    Merlin "Products Record Has Been Saved...", "Write"
    
    rsDisp.MoveFirst
    rsDisp.Find "ProdID=" & Val(txtProdID)
    
    Call Disp
End Sub

Private Sub cmbCategory_Click()
    Call fillCombo(cmbSize, "Sizes", "SizeInML", " Where Flag = '" & cmbCategory & "'")
End Sub

Private Sub refreshData()
    Set rsDisp = New ADODB.Recordset
    rsDisp.Open "Select Products.*, Sizes.SizeInMl, category.CatName from products, sizes, category where products.catid = category.catid and products.sizeid = sizes.sizeid order by prodid", myCn, 1, 2
End Sub

Private Sub cmbCategory_GotFocus()
    If Not flag = "" Then Merlin "Select Category of " & Trim(txtProductName) & " from this list", "Think"
End Sub

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbSize.SetFocus
End Sub

Private Sub cmbCategory_LostFocus()
    checkCombos cmbCategory, False
End Sub

Private Sub cmbSize_GotFocus()
    If Not flag = "" Then Merlin "Select Size... In Case of Wine, Beer or Cold Drinks", "Think"
End Sub

Private Sub cmbSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRate.SetFocus
End Sub

Private Sub cmbSize_LostFocus()
    checkCombos cmbSize, False
End Sub

Private Sub Form_Load()
    Call fillCombo(cmbCategory, "Category", "CatName")
    
    Call refreshData
    Call Disp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsProd = Nothing
    Set rsDisp = Nothing
    flag = ""
End Sub

Private Sub txt30MlRate_GotFocus()
    If Not flag = "" Then Merlin "Rate of 30 ML Loose Sale of " & Trim(txtProductName), "Pleased"
End Sub

Private Sub txt60MlRate_GotFocus()
    If Not flag = "" Then Merlin "Rate of 60 ML Loose Sale of " & Trim(txtProductName), "Pleased"
End Sub

Private Sub txt90MlRate_GotFocus()
    If Not flag = "" Then Merlin "Rate of 90 ML Loose Sale of " & Trim(txtProductName), "Pleased"
End Sub

Private Sub txt30MlRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt60MlRate.SetFocus Else setNumeric KeyAscii
End Sub

Private Sub txt60MlRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt90MlRate.SetFocus Else setNumeric KeyAscii
End Sub

Private Sub txt90MlRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If btnSave.Visible Then btnSave.SetFocus Else btnNew.SetFocus
End Sub

Private Sub txtOpn_GotFocus()
    If Not flag = "" Then Merlin "Opening Stock of " & Trim(txtProductName) & " which you currently have on your counter", "Pleased"
End Sub

Private Sub txtOpn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOpnLoose.SetFocus Else setNumeric KeyAscii
End Sub

Private Sub txtOpnLoose_GotFocus()
    If cmbCategory = "Wine" Then
        If Not flag = "" Then Merlin "Opening Stock in Loose of " & Trim(txtProductName) & " which you currently have on your counter", "Pleased"
    Else
        If Not flag = "" Then Merlin "Loose Opening Stock is Not applicable for " & Trim(txtProductName), "Pleased"
    End If
End Sub

Private Sub txtOpnLoose_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbCategory = "Wine" Then
            txt30MlRate.SetFocus
        Else
            If btnSave.Visible Then btnSave.SetFocus Else btnNew.SetFocus
        End If
    Else
        setNumeric KeyAscii
    End If
End Sub

Private Sub txtProductName_GotFocus()
    If flag = "New" Then Merlin "Enter Product Name And Description..."
End Sub

Private Sub txtProductName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbCategory.SetFocus Else tCase txtProductName, KeyAscii
End Sub

Private Sub txtRate_GotFocus()
    If Not flag = "" Then Merlin "Unit Rate of " & Trim(txtProductName)
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtOpn.SetFocus Else setNumeric KeyAscii
End Sub
