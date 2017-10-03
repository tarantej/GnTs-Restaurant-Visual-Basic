VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Rupali Bar & Restaurant."
   ClientHeight    =   7560
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   9270
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar RupaToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Accounts"
            Object.ToolTipText     =   "Manage Accounts"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Products"
            Object.ToolTipText     =   "Manage Products"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Update_Products"
            Object.ToolTipText     =   "Update Product Profile"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep1"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sales"
            Object.ToolTipText     =   "Create Bills"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Inward"
            Object.ToolTipText     =   "Materials Inward"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Receipt"
            Object.ToolTipText     =   "Receipt Entry"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Payment"
            Object.ToolTipText     =   "Payment Entry"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Loose"
            Object.ToolTipText     =   "Loose a Wine Bottle"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep2"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ledger"
            Object.ToolTipText     =   "Show Ledger"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stock"
            Object.ToolTipText     =   "Stock Register"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sales_Summary"
            Object.ToolTipText     =   "Show Sales Summary"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Billing_Monitor"
            Object.ToolTipText     =   "Billing Monitor"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sep3"
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Quit"
            Object.ToolTipText     =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":268E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":3368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":4042
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":4D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":59F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":66D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":73AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":8084
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":8D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":9078
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":9D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":A9D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Window.frx":B626
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   1320
      Top             =   2640
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master"
      Begin VB.Menu mnuProduct 
         Caption         =   "&Products"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAcct 
         Caption         =   "&Accounts"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuUpdates 
         Caption         =   "Update Masters"
         Begin VB.Menu mnuProductUpdate 
            Caption         =   "Update Product Profile"
         End
      End
      Begin VB.Menu Dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuBilling 
         Caption         =   "&Create Bills"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuInward 
         Caption         =   "&Products Inward"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuLoose 
         Caption         =   "&Loose Wine Bottle"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReceipt 
         Caption         =   "&Receipt"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPayment 
         Caption         =   "&Payment"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuBillingMonitor 
         Caption         =   "&Billing Monitor"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuShowToolbar 
         Caption         =   "Toolbar"
      End
      Begin VB.Menu mnuAgent 
         Caption         =   "Agent"
      End
   End
   Begin VB.Menu mnuOutputs 
      Caption         =   "&Outputs"
      Begin VB.Menu mnuStock 
         Caption         =   "&Stock Register"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "&Ledger"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSalesSummary 
         Caption         =   "&Sales Summary"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuCurrBal 
         Caption         =   "&Account's Current Balance"
         Shortcut        =   ^B
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    With RupaToolbar
        .ImageList = ImageList1
        .Buttons.Item(1).Image = 4
        .Buttons.Item(2).Image = 1
        .Buttons.Item(3).Image = 15
        
        .Buttons.Item(5).Image = 3
        .Buttons.Item(6).Image = 11
        .Buttons.Item(7).Image = 14
        .Buttons.Item(8).Image = 13
        .Buttons.Item(9).Image = 10
        
        .Buttons.Item(11).Image = 16
        .Buttons.Item(12).Image = 7
        .Buttons.Item(13).Image = 12
        .Buttons.Item(14).Image = 2
        
        .Buttons.Item(16).Image = 8
    End With
    
    'set toolbar status
    RupaToolbar.Visible = GetSetting("Bar", "MDI", "RupaToolbar.Visible", True)
    mnuShowToolbar.Checked = GetSetting("Bar", "MDI", "RupaToolbar.Visible", True)
    mnuAgent.Checked = GetSetting("Bar", "MDI", "mnuAgent.Checked", True)

    Call Init
    
    'Initialize Agent
    MyAgent.Characters.Load "Merlin", "Merlin.Acs"
    Set myCharacter = MyAgent.Characters("Merlin")
    
    myCharacter.SoundEffectsOn = True
    
    showMerlin
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuAcct_Click()
    frmAccounts.Show
End Sub

Private Sub mnuAgent_Click()
    mnuAgent.Checked = Not mnuAgent.Checked
    
    SaveSetting "Bar", "MDI", "mnuAgent.Checked", mnuAgent.Checked
    showMerlin
End Sub

Private Sub mnuBilling_Click()
    frmSales.Show
    
    frmSales.Top = GetSetting("Bar", "frmSales", "Top", (frmMain.Height - frmSales.Height) / 3)
    frmSales.Left = GetSetting("Bar", "frmSales", "Left", (frmMain.Width - frmSales.Width) / 2)
End Sub

Private Sub mnuBillingMonitor_Click()
    frmBillingMonitor.Show
    
    frmBillingMonitor.Top = GetSetting("Bar", "frmBillingMonitor", "Top", (frmMain.Height - frmBillingMonitor.Height) / 3)
    frmBillingMonitor.Left = GetSetting("Bar", "frmBillingMonitor", "Left", (frmMain.Width - frmBillingMonitor.Width) / 2)
End Sub

Private Sub mnuCurrBal_Click()
    Dim vAcctName As String
    Dim vMsg As String
    Dim vAcctNo As Integer
    Dim vCurrBal As Single
    
    vAcctNo = frmFind.getKey("Accounts", "AcctName")
    
    If vAcctNo = -1 Then Exit Sub
    
    vAcctName = getAcctDetailsByCode(vAcctNo)!AcctName
    vCurrBal = getAcctBalance(vAcctNo)
    
    vMsg = vAcctName & " Has a Balance of Rs : " & IIf(vCurrBal > 0, Format(Abs(vCurrBal), "0.00") & " Dr", Format(Abs(vCurrBal), "0.00") & " Cr")
    Merlin vMsg, "Read"
End Sub

Private Sub mnuInward_Click()
    ShowInCentre frmInward
End Sub

Private Sub mnuLedger_Click()
    ShowInCentre frmLedger
End Sub

Private Sub mnuLoose_Click()
    frmLoose.Show
End Sub

Private Sub mnuPayment_Click()
    frmVoucher.Init ("Payment")
End Sub

Private Sub mnuProduct_Click()
    frmProducts.Show
End Sub

Private Sub mnuProductUpdate_Click()
    ShowInCentre frmProductsUpdate
End Sub

Private Sub mnuQuit_Click()
    End
End Sub

Private Sub mnuReceipt_Click()
    frmVoucher.Init ("Receipt")
End Sub

Private Sub mnuSalesSummary_Click()
    frmDates.Show
    If datesSelected Then ShowInCentre frmSalesSummary
End Sub

Private Sub mnuShowToolbar_Click()
    RupaToolbar.Visible = Not RupaToolbar.Visible
    mnuShowToolbar.Checked = Not mnuShowToolbar.Checked
    
    SaveSetting "Bar", "MDI", "RupaToolbar.Visible", RupaToolbar.Visible
End Sub

Private Sub mnuStock_Click()
    Call initDtEnv
    rptStock.Show
End Sub

Private Sub RupaToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Accounts"
        mnuAcct_Click
        
        Case "Products"
        mnuProduct_Click
        
        Case "Update_Products"
        mnuProductUpdate_Click
        
        Case "Sales"
        mnuBilling_Click
        
        Case "Inward"
        mnuInward_Click
        
        Case "Receipt"
        mnuReceipt_Click
        
        Case "Payment"
        mnuPayment_Click
        
        Case "Loose"
        mnuLoose_Click
        
        Case "Ledger"
        mnuLedger_Click
        
        Case "Stock"
        mnuStock_Click
        
        Case "Sales_Summary"
        mnuSalesSummary_Click
        
        Case "Billing_Monitor"
        mnuBillingMonitor_Click
        
        Case "Quit"
        mnuQuit_Click
    End Select
End Sub
