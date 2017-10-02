<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Bill
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Calculator2 = New Calculator.Calculator
        Me.Label16 = New System.Windows.Forms.Label
        Me.TextBox16 = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.TextBox17 = New System.Windows.Forms.TextBox
        Me.TextBox18 = New System.Windows.Forms.TextBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.TextBox19 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ListView2 = New System.Windows.Forms.ListView
        Me.ItemCode = New System.Windows.Forms.ColumnHeader
        Me.ItemName = New System.Windows.Forms.ColumnHeader
        Me.ItemQuantity = New System.Windows.Forms.ColumnHeader
        Me.ItemPrice = New System.Windows.Forms.ColumnHeader
        Me.PrintDocument2 = New System.Drawing.Printing.PrintDocument
        Me.SuspendLayout()
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(12, 21)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(40, 13)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Bill No."
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(526, 15)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(30, 13)
        Me.Label11.TabIndex = 1
        Me.Label11.Text = "Date"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(599, 15)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(200, 20)
        Me.DateTimePicker1.TabIndex = 2
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(25, 65)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(55, 13)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Item Code"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(190, 65)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(58, 13)
        Me.Label13.TabIndex = 4
        Me.Label13.Text = "Item Name"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(539, 65)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(31, 13)
        Me.Label14.TabIndex = 5
        Me.Label14.Text = "Price"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(74, 14)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(100, 20)
        Me.TextBox2.TabIndex = 6
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(451, 65)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(46, 13)
        Me.Label15.TabIndex = 7
        Me.Label15.Text = "Quantity"
        '
        'Calculator2
        '
        Me.Calculator2.Location = New System.Drawing.Point(615, 84)
        Me.Calculator2.Name = "Calculator2"
        Me.Calculator2.Size = New System.Drawing.Size(264, 281)
        Me.Calculator2.TabIndex = 61
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(25, 400)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(43, 13)
        Me.Label16.TabIndex = 62
        Me.Label16.Text = "Amount"
        '
        'TextBox16
        '
        Me.TextBox16.Location = New System.Drawing.Point(542, 393)
        Me.TextBox16.Name = "TextBox16"
        Me.TextBox16.Size = New System.Drawing.Size(100, 20)
        Me.TextBox16.TabIndex = 63
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(25, 437)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(89, 13)
        Me.Label17.TabIndex = 64
        Me.Label17.Text = "Value Added Tax"
        '
        'TextBox17
        '
        Me.TextBox17.Location = New System.Drawing.Point(542, 437)
        Me.TextBox17.Name = "TextBox17"
        Me.TextBox17.Size = New System.Drawing.Size(100, 20)
        Me.TextBox17.TabIndex = 65
        '
        'TextBox18
        '
        Me.TextBox18.Location = New System.Drawing.Point(542, 477)
        Me.TextBox18.Name = "TextBox18"
        Me.TextBox18.Size = New System.Drawing.Size(100, 20)
        Me.TextBox18.TabIndex = 66
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(25, 477)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(47, 13)
        Me.Label18.TabIndex = 67
        Me.Label18.Text = "Total Bill"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(27, 507)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(46, 13)
        Me.Label19.TabIndex = 68
        Me.Label19.Text = "Balance"
        '
        'TextBox19
        '
        Me.TextBox19.Location = New System.Drawing.Point(542, 507)
        Me.TextBox19.Name = "TextBox19"
        Me.TextBox19.Size = New System.Drawing.Size(100, 20)
        Me.TextBox19.TabIndex = 69
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(227, 535)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 58)
        Me.Button1.TabIndex = 70
        Me.Button1.Text = "Print"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(358, 535)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(94, 58)
        Me.Button2.TabIndex = 71
        Me.Button2.Text = "Reset"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(99, 535)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(104, 58)
        Me.Button3.TabIndex = 72
        Me.Button3.Text = "Total Bill"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ListView2
        '
        Me.ListView2.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ItemCode, Me.ItemName, Me.ItemQuantity, Me.ItemPrice})
        Me.ListView2.Location = New System.Drawing.Point(29, 89)
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(554, 255)
        Me.ListView2.TabIndex = 73
        Me.ListView2.UseCompatibleStateImageBehavior = False
        '
        'Bill
        '
        Me.ClientSize = New System.Drawing.Size(889, 605)
        Me.Controls.Add(Me.ListView2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox19)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.TextBox18)
        Me.Controls.Add(Me.TextBox17)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.TextBox16)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Calculator2)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Name = "Bill"
        Me.Text = "Bill Receipt"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Bill_Date As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox20 As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox21 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBox22 As System.Windows.Forms.TextBox
    Friend WithEvents cmdBill_Print As System.Windows.Forms.Button
    Friend WithEvents cmdBill_Reset As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Item_Code As System.Windows.Forms.ColumnHeader
    Friend WithEvents Item_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents Item_Quantity As System.Windows.Forms.ColumnHeader
    Friend WithEvents Item_Price As System.Windows.Forms.ColumnHeader
    Friend WithEvents Calculator1 As Calculator.Calculator
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Calculator2 As Calculator.Calculator
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TextBox16 As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TextBox17 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox18 As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents TextBox19 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ListView2 As System.Windows.Forms.ListView
    Friend WithEvents ItemCode As System.Windows.Forms.ColumnHeader
    Friend WithEvents ItemName As System.Windows.Forms.ColumnHeader
    Friend WithEvents ItemQuantity As System.Windows.Forms.ColumnHeader
    Friend WithEvents ItemPrice As System.Windows.Forms.ColumnHeader
    Friend WithEvents PrintDocument2 As System.Drawing.Printing.PrintDocument
End Class
