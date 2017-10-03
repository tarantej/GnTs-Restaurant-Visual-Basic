VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProductsUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Product Profile."
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraProgress 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   6120
         Top             =   240
      End
      Begin MSComctlLib.ProgressBar rupaProgress 
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   3
         Height          =   1695
         Left            =   0
         Top             =   0
         Width           =   7455
      End
      Begin VB.Label lblWait 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait... Updating Closing Stock."
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
         TabIndex        =   9
         Top             =   240
         Width           =   5595
      End
   End
   Begin MSDataGridLib.DataGrid dtProducts 
      Bindings        =   "Update Product Rates & Openings.frx":0000
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "cmdProductProfile_form"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ProdID"
         Caption         =   "Sr No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Name"
         Caption         =   "Product Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "SizeInMl"
         Caption         =   "Size In Ml"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "OpenStock"
         Caption         =   "Opening Stock"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "OpenStockLoose"
         Caption         =   "Opening Loose"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Rate"
         Caption         =   "Rate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "L30MlRate"
         Caption         =   "30 ML"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "L60MlRate"
         Caption         =   "60 ML"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "L90MlRate"
         Caption         =   "90 ML"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   975.118
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Height          =   195
      Left            =   7800
      TabIndex        =   7
      Top             =   1155
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Rate"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   1155
      Width           =   420
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Opening Stock"
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
      Left            =   4215
      TabIndex        =   5
      Top             =   1155
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Particulars of Product"
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
      Left            =   435
      TabIndex        =   4
      Top             =   1155
      Width           =   1860
   End
   Begin VB.Line Line6 
      X1              =   7095
      X2              =   7095
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line5 
      X1              =   6075
      X2              =   6075
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line4 
      X1              =   3660
      X2              =   3660
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line3 
      X1              =   10440
      X2              =   10440
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10455
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You can change Opening Stock, Product Rate && Loose Sale Rates of Wine."
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   9480
      Picture         =   "Update Product Rates & Openings.frx":0014
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Modification in Registered Products."
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
      TabIndex        =   1
      Top             =   240
      Width           =   3105
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10335
   End
End
Attribute VB_Name = "frmProductsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOpnStockChanged As Boolean

Dim rsInward As ADODB.Recordset
Dim rsSold As ADODB.Recordset
Dim rsLoosed As ADODB.Recordset

Private Sub dtProducts_ColEdit(ByVal ColIndex As Integer)
    If ColIndex = 3 Or ColIndex = 4 Then vOpnStockChanged = True
End Sub

Private Sub Form_Load()
    vOpnStockChanged = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If vOpnStockChanged Then
        'change closing stock of every products
        fraProgressStatus True
        dtProducts.DataChanged = True
        
        Dim vOpenNos As Single
        Dim vInwardedNos As Single
        Dim vSoldNos As Single
        Dim vLoosedNos As Single
        Dim vOpenLoose As Single
        Dim vSoldLoose As Single
        Dim vLoosedMl As Single
        Dim i As Integer
                
        Dim rsProducts As New ADODB.Recordset
        With rsProducts
            .Open "Select * from Products order by ProdID", myCn, 1, 2
            rupaProgress.Max = .RecordCount
            
            While Not .EOF
                i = i + 1
                rupaProgress.Value = i
                DoEvents
                
                vOpenNos = !OpenStock
                vInwardedNos = Inwarded(!ProdID)
                vSoldNos = Sold(!ProdID, False)
                vLoosedNos = LoosedNos(!ProdID)
                
                vOpenLoose = !OpenStockLoose
                vSoldLoose = Sold(!ProdID, True)
                vLoosedMl = getMlSizeOfProd(!ProdID) * vLoosedNos
                                
                !CurrStock = (vOpenNos + vInwardedNos) - (vSoldNos + vLoosedNos)
                !CurrStockLoose = vOpenLoose + vLoosedMl - vSoldLoose
                
                .Update
                .MoveNext
            Wend
            .Close
        End With
        fraProgressStatus False
    End If
End Sub

Private Function Inwarded(ByVal pProdID As Integer) As Single
        Set rsInward = New ADODB.Recordset
        rsInward.Open "Select Sum(Qty) from InwardDetail group by ProdID having ProdID = " & pProdID, myCn, 1, 2
        
        If Not rsInward.BOF Then Inwarded = rsInward(0)
        Set rsInward = Nothing
End Function

Private Function Sold(ByVal pProdID As Integer, Optional pLoose As Boolean) As Single
        Set rsSold = New ADODB.Recordset
        
        If pLoose Then
            rsSold.Open "Select Sum(Qty) from BillDetail where Loose group by ProdID having ProdID = " & pProdID, myCn, 1, 2
        Else
            rsSold.Open "Select Sum(Qty) from BillDetail where not Loose group by ProdID having ProdID = " & pProdID, myCn, 1, 2
        End If
        
        If Not rsSold.BOF Then Sold = rsSold(0)
        Set rsSold = Nothing
End Function

Private Function LoosedNos(ByVal pProdID As Integer) As Single
        Set rsLoosed = New ADODB.Recordset
        rsLoosed.Open "Select count(ProdID) from LooseRecord where ProdID = " & pProdID, myCn, 1, 2
        
        If Not rsLoosed.BOF Then LoosedNos = rsLoosed(0)
        Set rsLoosed = Nothing
End Function

Private Sub Timer1_Timer()
    lblWait.Visible = Not lblWait.Visible
End Sub

Private Sub fraProgressStatus(ByVal pStatus As Boolean)
    fraProgress.Visible = pStatus
    Timer1.Enabled = pStatus
End Sub
