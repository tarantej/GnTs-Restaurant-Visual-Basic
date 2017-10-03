VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form frmLoose 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtLoose 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   16384003
      CurrentDate     =   39045
   End
   Begin VB.TextBox txtSrNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox cmbProduct 
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   3120
      Width           =   3855
   End
   Begin lvButton.lvButtons_H btnExit 
      Height          =   375
      Left            =   7440
      TabIndex        =   9
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
      Image           =   "Loose Wine Bottles.frx":0000
      cBack           =   16777215
   End
   Begin lvButton.lvButtons_H btnDelete 
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "&Delete"
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
      Image           =   "Loose Wine Bottles.frx":0619
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnNew 
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   5640
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
      Image           =   "Loose Wine Bottles.frx":08BF
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnCancel 
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   5640
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
      Image           =   "Loose Wine Bottles.frx":10E8
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin lvButton.lvButtons_H btnSave 
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   5640
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
      Image           =   "Loose Wine Bottles.frx":2121
      ImgSize         =   24
      cBack           =   12640511
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      BorderWidth     =   2
      Height          =   3015
      Left            =   720
      Top             =   2160
      Width           =   5895
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   0
      Left            =   720
      Picture         =   "Loose Wine Bottles.frx":26A9
      Top             =   5520
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   1
      Left            =   1080
      Picture         =   "Loose Wine Bottles.frx":5C3C
      Top             =   5520
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   2
      Left            =   1440
      Picture         =   "Loose Wine Bottles.frx":913A
      Top             =   5520
      Width           =   285
   End
   Begin VB.Image btnNav 
      Height          =   285
      Index           =   3
      Left            =   1800
      Picture         =   "Loose Wine Bottles.frx":C684
      Top             =   5520
      Width           =   285
   End
   Begin VB.Label Label4 
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3600
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry No"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   3120
      Width           =   675
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
      Left            =   690
      TabIndex        =   6
      Top             =   300
      Width           =   3420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entry of Loosed Bottles."
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
      TabIndex        =   5
      Top             =   1440
      Width           =   4875
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Picture         =   "Loose Wine Bottles.frx":FC26
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   9855
   End
End
Attribute VB_Name = "frmLoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsLoose As ADODB.Recordset

Private Sub btnCancel_Click()
    Call btnSets(False, Me)
    Call Disp
    Merlin "Entry Not Saved...", "Wave"
End Sub

Private Sub btnDelete_Click()
    Dim resp As VbMsgBoxResult
    Dim vProdID As Integer

    resp = MsgBox("Are You Sure to Delete this Record ? ", vbYesNo + vbQuestion + vbDefaultButton2)
    
    If resp = vbYes Then
        vProdID = cmbProduct.ItemData(cmbProduct.ListIndex)
        
        'increase a bottle
        increaseStock vProdID, 1, False
        
        'decrease in loose stock
        decreaseStock vProdID, getMlSizeOfProd(vProdID), True
    
        rsLoose.Delete
        
        Merlin "Entry Has Been Deleted...", "Wave"
    End If
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnNav_Click(Index As Integer)
    On Error Resume Next
    
    With rsLoose
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

Private Sub btnNew_Click()
    txtSrNo = ""
    cmbProduct = ""
    dtLoose = Date
    
    Call btnSets(True, Me)
    
    cmbProduct.SetFocus
End Sub

Private Sub btnSave_Click()
    If Trim(cmbProduct) = "" Then
        MsgBox "Select Product..."
        cmbProduct.SetFocus
        Exit Sub
    End If

    Dim vProdID As Integer
    vProdID = cmbProduct.ItemData(cmbProduct.ListIndex)
    
    Set rsLoose = New ADODB.Recordset
    With rsLoose
        .Open "Select * from LooseRecord order by SrNo", myCn, 1, 2
        
        If .BOF Then
            txtSrNo = 1
        Else
            .MoveLast
            txtSrNo = .Fields("SrNo") + 1
        End If
        
        .AddNew
            !SrNo = Val(txtSrNo)
            !Date = dtLoose
            !Time = Time
            !ProdID = vProdID
        .Update
    End With
        
    'decrease a bottle
    decreaseStock vProdID, 1, False
    
    'increase in loose stock
    increaseStock vProdID, getMlSizeOfProd(vProdID), True
    
    Merlin "This Entry Has Been Saved...", "Write"
    Call btnSets(False, Me)
End Sub

Private Sub cmbProduct_GotFocus()
    Merlin "Select Product which you have been loosed... And Click on Save Button Below", "Pleased"
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select Name, ProdID, SizeInMl, CatName from products, Sizes, Category where Products.SizeID=sizes.SizeId and products.catid=category.catid and CatName='Wine' order by products.name, sizes.sizeinml", myCn, 1, 2
    
    cmbProduct.Clear
    
    While Not rsTemp.EOF    'dont add '.' as size
        cmbProduct.AddItem IIf(rsTemp!SizeInMl = ".", rsTemp!Name, rsTemp!Name & " " & rsTemp!SizeInMl)
        cmbProduct.ItemData(cmbProduct.NewIndex) = rsTemp!ProdID
        
        rsTemp.MoveNext
    Wend
    rsTemp.Close
    
    Merlin "Here You Can Make Entries Loosed Wine Bottles, Click on 'New Entry' Button to Start Entry"
    
    Set rsLoose = New ADODB.Recordset
    rsLoose.Open "Select * from LooseRecord order by SrNo", myCn, 1, 2
    
    Call Disp
End Sub

Private Sub Disp()
    With rsLoose
        If .BOF = False And .EOF = False Then
            txtSrNo = !SrNo
            dtLoose = !Date
            
            cmbProduct.ListIndex = 0
            While Not cmbProduct.ItemData(cmbProduct.ListIndex) = !ProdID
                cmbProduct.ListIndex = cmbProduct.ListIndex + 1
            Wend
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsLoose = Nothing
End Sub
