Attribute VB_Name = "modGeneral"
Option Explicit
Option Compare Text

Public myCn             As ADODB.Connection     'connection to software database
Public newBills         As Boolean              'flag tells billing monitor that new bills has been created
Public dtFrom, dtTo     As Date                 'period for periodic reports
Public datesSelected    As Boolean
Public myCharacter      As IAgentCtlCharacterEx 'ms agent

'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'System         : Restaurant & Bar Management System.
'Author         : Haresh Eshwarlal Jaiswal.
'Organization   : Rising Technologies, Jalna. (Maharashtra)
'E-Mail         : Haresh_Jaiswal@Rediffmail.com

'Contact Nos    : Ph: 02482 - 240212,  Mob: 9423156065
'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
'If you'have any suggestions please mail me at above said address

Sub main()
    Call Init
    frmSplash.Show
    'frmMain.Show
End Sub



'this sub should be called to establish a connection with database
Public Sub Init()
    Set myCn = New ADODB.Connection
    myCn.ConnectionString = "provider=microsoft.jet.oledb.4.0;data source=" & App.Path & "\Data\dbRupali.mdb"
    myCn.Open
    
    Call initDtEnv
End Sub

'this sub will reset data environment
Public Sub initDtEnv()
    Set DtEnv = New DtEnv
    DtEnv.Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Data\dbRupali.mdb"
End Sub

'this sub will make a debit entry to given accounts ledger
Public Sub Debit(ByVal pTranDate As Date, ByVal pAcctNo As Integer, ByVal pParticulars As String, pTranType As String, ByVal pEntryNo As Long, ByVal pPrintNo As Long, ByVal pDebitAmt As Single)
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * from Ledger", myCn, 1, 2
        
        .AddNew
            !TranDate = pTranDate
            !acctNo = pAcctNo
            !Particulars = pParticulars
            !TranType = pTranType
            !entryNo = pEntryNo
            !PrintNo = pPrintNo
            !DebitAmt = pDebitAmt
        .Update
        
        .Close
    End With
End Sub

'this sub will make a credit entry to given accounts ledger
Public Sub Credit(ByVal pTranDate As Date, ByVal pAcctNo As Integer, ByVal pParticulars As String, pTranType As String, ByVal pEntryNo As Long, ByVal pPrintNo As Long, ByVal pCreditAmt As Single)
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * from Ledger", myCn, 1, 2
        
        .AddNew
            !TranDate = pTranDate
            !acctNo = pAcctNo
            !Particulars = pParticulars
            !TranType = pTranType
            !entryNo = pEntryNo
            !PrintNo = pPrintNo
            !CreditAmt = pCreditAmt
        .Update
        
        .Close
    End With
End Sub

'this sub will delete a entry from given accounts ledger
Public Sub DeleteLedgerEntry(ByVal pTransType As String, ByVal pEntryNo As Long)
    myCn.Execute "Delete from Ledger where TranType='" & pTransType & "' and EntryNo=" & pEntryNo
End Sub

Public Function newReceipt(ByVal pDate As Date, ByVal pAcctNoToCredit As Integer, ByVal pAmt As Single, Optional ByVal pAgainst As String = "On Account", Optional ByVal pRemarks As String = "Cash Recd") As Long
    Dim rs As New ADODB.Recordset
    Dim tmp As New ADODB.Recordset
    
    '************************************************************
    'Generate Voucher Serial No
    Dim vSrNo As Long
    tmp.Open "Select VSrNo from Voucher order by vSrNo", myCn, 1, 2
        
    If tmp.BOF Then
        vSrNo = 1
    Else
        tmp.MoveLast
        vSrNo = tmp!vSrNo + 1
    End If
    Set tmp = Nothing
    '************************************************************
    
    With rs
        .Open "Select * from Voucher Where VchType='Receipt' order by vchNo", myCn, 1, 2
        
        If .BOF Then
            newReceipt = 1
        Else
            .MoveLast
            newReceipt = !vchno + 1
        End If
        
        .AddNew
            !vSrNo = vSrNo
            !vchtype = "Receipt"
            !vchno = newReceipt
            !Date = pDate
            !amt = pAmt
            !acctNo = pAcctNoToCredit
            !Against = pAgainst
            !Remark = pRemarks
        .Update
        .Close
    End With
    Set rs = Nothing
    
    'make entries to ledger
    
    'debit the cash account
    Debit pDate, BI_CashAcct, "Receipt : from. " & getAcctDetailsByCode(pAcctNoToCredit)!AcctName, "Rcpt", vSrNo, newReceipt, pAmt
    
    'credit the party account
    If pAgainst = "On Account" Then
        Credit pDate, pAcctNoToCredit, "Receipt", "Rcpt", vSrNo, newReceipt, pAmt
    Else
        Credit pDate, pAcctNoToCredit, "Receipt : Against Bill No. " & pAgainst, "Rcpt", vSrNo, newReceipt, pAmt
    End If
    
    'if against bill no is mentioned then make it paid
    If Not pAgainst = "On Account" Then
        If getCrBillBalance(pAgainst) = 0 Then myCn.Execute "Update BillMaster set Paid = 1 where billno = " & Val(pAgainst)
    End If
End Function

Public Function newPayment(ByVal pDate As Date, ByVal pAcctNoToDebit As Integer, ByVal pAmt As Single, Optional ByVal pAgainst As String = "On Account", Optional ByVal pRemarks As String = "Cash Recd") As Long
    Dim rs As New ADODB.Recordset
    Dim tmp As New ADODB.Recordset
    
    '************************************************************
    'Generate Voucher Serial No
    Dim vSrNo As Long
    tmp.Open "Select VSrNo from Voucher order by vSrNo", myCn, 1, 2
        
    If tmp.BOF Then
        vSrNo = 1
    Else
        tmp.MoveLast
        vSrNo = tmp!vSrNo + 1
    End If
    Set tmp = Nothing
    '************************************************************
    
    With rs
        .Open "Select * from Voucher Where VchType='Payment' order by vchNo", myCn, 1, 2
        
        If .BOF Then
            newPayment = 1
        Else
            .MoveLast
            newPayment = !vchno + 1
        End If
        
        .AddNew
            !vSrNo = vSrNo
            !vchtype = "Payment"
            !vchno = newPayment
            !Date = pDate
            !amt = pAmt
            !acctNo = pAcctNoToDebit
            !Against = pAgainst
            !Remark = pRemarks
        .Update
        .Close
    End With
    Set rs = Nothing
    
    'make entries to ledger
    
    'credit the cash account
    Credit pDate, BI_CashAcct, "Payment : To. " & getAcctDetailsByCode(pAcctNoToDebit)!AcctName, "Pymt", vSrNo, newPayment, pAmt
    
    'Debit the Party account
    Debit pDate, pAcctNoToDebit, "Payment", "Pymt", vSrNo, newPayment, pAmt
End Function

Public Sub clearAll(currentForm As Form)
    Dim ctl As Control
    
    For Each ctl In currentForm.Controls
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then ctl.Text = ""
    Next
End Sub

'this function will converts the text to title case format
Public Function tCase(str As String, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Function
    End If
    
    If str = Empty Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        If Right(str, 1) = " " Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            'KeyAscii = Asc(LCase(Chr(KeyAscii)))
        End If
    End If
    
    tCase = KeyAscii
End Function

Public Sub fillCombo(cmbName As ComboBox, tblName As String, fldName As String, Optional criteria As String)
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from " & tblName & criteria & " order by " & fldName, myCn, adOpenForwardOnly, adLockOptimistic
    
    cmbName.Clear
    
    Do Until rs.EOF
        cmbName.AddItem rs.Fields(fldName)
        cmbName.ItemData(cmbName.NewIndex) = rs.Fields(0)
        rs.MoveNext
    Loop
    rs.Close
End Sub

'this function will inforce only numeric data should be accepted by the text box
Public Function setNumeric(KeyAscii As Integer)
    If Not Chr(KeyAscii) Like "#" And KeyAscii <> 8 And KeyAscii <> 46 Then KeyAscii = 0
End Function

Public Sub btnSets(Status As Boolean, currentForm As Form)
    On Error Resume Next
    
    currentForm.btnNew.Visible = Not Status
    currentForm.btnEdit.Visible = Not Status
    
    currentForm.btnSave.Visible = Status
    currentForm.btnCancel.Visible = Status
    currentForm.btnDelete.Visible = Not Status
    
    currentForm.btnNav(0).Visible = Not Status
    currentForm.btnNav(1).Visible = Not Status
    currentForm.btnNav(2).Visible = Not Status
    currentForm.btnNav(3).Visible = Not Status
    
    currentForm.btnFind.Visible = Not Status
    currentForm.btnPrint.Visible = Not Status
End Sub

Public Sub checkCombos(cmbName As ComboBox, Optional SetFirst As Boolean = True)
    Dim cnt As Integer
    Dim found As Boolean
    
    For cnt = 0 To cmbName.ListCount - 1
        If cmbName = cmbName.List(cnt) Then
            found = True
            Exit For
        End If
    Next
    
    If Not found Then
        If cmbName.ListCount > 0 Then
            If SetFirst Then cmbName.ListIndex = 0 Else cmbName = ""
        Else
            cmbName = ""
        End If
    End If
End Sub

Public Sub increaseStock(ByVal vProductID As Integer, ByVal vQty As Single, vLoose As Boolean)
    If vLoose Then
        myCn.Execute "Update Products Set CurrStockLoose=CurrStockLoose + " & vQty & " Where ProdID=" & vProductID
    Else
        myCn.Execute "Update Products Set CurrStock=CurrStock + " & vQty & " Where ProdID=" & vProductID
    End If
End Sub

Public Sub decreaseStock(ByVal vProductID As Integer, ByVal vQty As Single, vLoose As Boolean)
    If vLoose Then
        myCn.Execute "Update Products Set CurrStockLoose=CurrStockLoose - " & vQty & " Where ProdID=" & vProductID
    Else
        myCn.Execute "Update Products Set CurrStock=CurrStock - " & vQty & " Where ProdID=" & vProductID
    End If
End Sub

Public Function getCategoryDets(strCatName As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from Category where catName='" & strCatName & "'", myCn
    
    Set getCategoryDets = rs
End Function

Public Function getCategoryDetsByCode(intCatID As Integer) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from Category where catID=" & intCatID, myCn
    
    Set getCategoryDetsByCode = rs
End Function

Public Function getSizeDets(strSizeInMl As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from Sizes where SizeInMl='" & strSizeInMl & "'", myCn
    
    Set getSizeDets = rs
End Function

Public Function getSizeDetsByCode(IntSizeID As Integer) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from Sizes where SizeId=" & IntSizeID, myCn
    
    Set getSizeDetsByCode = rs
End Function

Public Function getBillDets(ByVal pBillNo As Long) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "Select * from BillMaster where BillNo=" & pBillNo, myCn
    
    Set getBillDets = rs
End Function

Public Function getAcctBalance(ByVal acctNo As Integer) As Single
    Dim rs As New ADODB.Recordset
    rs.Open "select sum(debitAmt) as dbt,sum(creditAmt) as crd, dbt-crd as bal from Ledger group by AcctNo having AcctNo=" & acctNo, myCn
    
    If Not rs.BOF Then
        getAcctBalance = rs.Fields("bal")
    Else
        getAcctBalance = 0
    End If
    rs.Close
End Function

Public Function AddSpaces(sInput As String) As String
    Dim iPos As Integer
    Dim sTmpName As String
    
    sTmpName = Mid$(sInput, 1, 1)
    
    For iPos = 2 To Len(sInput)
        If Mid$(sInput, iPos, 1) <= "Z" Then sTmpName = sTmpName & " "
        
        sTmpName = sTmpName & Mid$(sInput, iPos, 1)
    Next
    
    AddSpaces = sTmpName
End Function

'this sub will prepare ledger to print
Public Sub prepareLedger(acctNo As Integer)
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "select * from Ledger where acctNo=" & acctNo & " order by tranDate,EntryNo", myCn, adOpenDynamic, adLockOptimistic
    
        If .BOF Then Exit Sub
        
        Dim bal As Single
        Dim i As Single
        
        While Not .EOF
            i = i + 1
            bal = bal + .Fields("DebitAmt") - .Fields("CreditAmt")
            .Fields("ClosingBal") = bal
            .Fields("srNo") = i
            
            .Update
            .MoveNext
        Wend
        .Close
    End With
End Sub

Public Sub ShowInCentre(formname As Form)
    formname.Show
    formname.Top = (frmMain.Height - formname.Height) / 4
    formname.Left = (frmMain.Width - formname.Width) / 2
End Sub

Public Function getMlSizeOfProd(ByVal pProdID As Integer) As Long
    Dim rs As New ADODB.Recordset
    rs.Open "Select SizeInMl from Sizes where SizeID=(Select SizeId from Products where ProdID=" & pProdID & ")", myCn
    If Val(rs!SizeInMl) > 0 Then
        getMlSizeOfProd = Val(Mid(rs!SizeInMl, 1, Len(rs!SizeInMl) - 3))
    Else
        getMlSizeOfProd = 0
    End If
    rs.Close
End Function

Public Function getAcctDetailsByCode(pAcctNo As Integer) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM Accounts where AcctNo=" & pAcctNo, myCn, adOpenDynamic, adLockOptimistic
    Set getAcctDetailsByCode = rs
End Function

'this function will return a balance amount of given bill nos srno
Public Function getCrBillBalance(ByVal vBillNo As Integer) As Single
    Dim rs As New ADODB.Recordset
    Dim BillAmt, PaidAmt As Single
       
    BillAmt = getBillDets(vBillNo).Fields("NetAmt")
    rs.Open "SELECT Voucher.Against, SUM(Voucher.Amt) AS sumOfPaidAmt FROM Voucher GROUP BY Voucher.Against HAVING Voucher.Against='" & vBillNo & "'", myCn
    If Not rs.BOF Then PaidAmt = rs.Fields(1)
    rs.Close
    
    getCrBillBalance = BillAmt - PaidAmt
End Function

Public Sub Merlin(Optional ByVal Msg As String, Optional ByVal Animation As String = "Explain")
    On Error Resume Next
    If myCharacter.Visible Then
        myCharacter.StopAll
        myCharacter.Play Animation
        
        If Not Msg = "" Then myCharacter.Speak Msg
    End If
End Sub

Public Sub showMerlin()
    If frmMain.mnuAgent.Checked Then
        myCharacter.Show
        
        'to botton right corner
        myCharacter.MoveTo 850, 550
    Else
        myCharacter.Play "Wave"
        myCharacter.Hide
    End If
End Sub
