VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmSalesSummary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Summary"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid grdDisplay 
      Height          =   2055
      Left            =   4320
      TabIndex        =   4
      Top             =   6360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   6255
      Left            =   120
      OleObjectBlob   =   "Sales Summary Chart.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   600
      TabIndex        =   1
      Top             =   6240
      Width           =   3615
      Begin VB.ComboBox cmbType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin lvButton.lvButtons_H btnShow 
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "Sales Summary Chart.frx":24B8
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin lvButton.lvButtons_H btnClose 
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
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
         Image           =   "Sales Summary Chart.frx":310D
         ImgSize         =   24
         cBack           =   12640511
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chart Type"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   780
      End
   End
End
Attribute VB_Name = "frmSalesSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnShow_Click()
    Call initDtEnv
    DtEnv.cmdSummary dtFrom, dtTo
    
    rptSummary.Sections("rptHead").Controls("lblDuration").Caption = "During " & Format(dtFrom, "dd-MMM-yyyy") & " To " & Format(dtTo, "dd-MMM-yyyy")
    rptSummary.Show 1

End Sub

Private Sub cmbType_Click()
    MSChart1.chartType = cmbType.ItemData(cmbType.ListIndex)
End Sub

Private Sub Form_Load()
    Call fillChartTypes
    cmbType.ListIndex = 0

    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open "SELECT Category.CatName, SUM(BillDetail.Amt) AS SumOfSale FROM Category, Products, BillDetail, BillMaster WHERE Category.CatID = Products.CatID AND Products.ProdID = BillDetail.ProdID AND BillDetail.BillNo = BillMaster.BillNo AND BillMaster.`Date` BETWEEN #" & dtFrom & "# AND #" & dtTo & "# GROUP BY Category.CatName", myCn, 1, 2
    
    Set grdDisplay.DataSource = rs
    With MSChart1
        '.ShowLegend = True
        .Title = "Sale During " & Format(dtFrom, "dd-MMM-yyyy") & " To. " & Format(dtTo, "dd-MMM-yyyy")
        .EditCopy
        .EditPaste
        Set .DataSource = rs
    End With
End Sub

Private Sub fillChartTypes()
    cmbType.AddItem "VtChChartType2dBar"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dBar
    
    cmbType.AddItem "VtChChartType2dArea"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dArea
    
    'cmbType.AddItem "VtChChartType2dCombination"
    'cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dCombination
    
    cmbType.AddItem "VtChChartType2dLine"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dLine
    
    'cmbType.AddItem "VtChChartType2dPie"
    'cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dLine
    
    cmbType.AddItem "VtChChartType2dStep"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dStep
    
    'cmbType.AddItem "VtChChartType2dXY"
    'cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType2dXY
    
    cmbType.AddItem "VtChChartType3dArea"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType3dArea
    
    cmbType.AddItem "VtChChartType3dBar"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType3dBar
    
    'cmbType.AddItem "VtChChartType3dCombination"
    'cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType3dCombination
    
    cmbType.AddItem "VtChChartType3dLine"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType3dLine
    
    cmbType.AddItem "VtChChartType3dStep"
    cmbType.ItemData(cmbType.NewIndex) = VtChChartType.VtChChartType3dStep
End Sub
