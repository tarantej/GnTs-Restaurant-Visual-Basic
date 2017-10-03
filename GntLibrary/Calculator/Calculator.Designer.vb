<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Calculator
    Inherits System.Windows.Forms.UserControl

    'UserControl1 overrides dispose to clean up the component list.
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
        Me.cmdSign = New System.Windows.Forms.Button
        Me.cmdInverse = New System.Windows.Forms.Button
        Me.cmdSqrRoot = New System.Windows.Forms.Button
        Me.cmdPowerOf = New System.Windows.Forms.Button
        Me.cmdEqual = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdDecimal = New System.Windows.Forms.Button
        Me.txtInput = New System.Windows.Forms.TextBox
        Me.cmdSubtract = New System.Windows.Forms.Button
        Me.cmd8 = New System.Windows.Forms.Button
        Me.cmd9 = New System.Windows.Forms.Button
        Me.cmd4 = New System.Windows.Forms.Button
        Me.cmd5 = New System.Windows.Forms.Button
        Me.cmd6 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmd1 = New System.Windows.Forms.Button
        Me.cmd2 = New System.Windows.Forms.Button
        Me.cmd3 = New System.Windows.Forms.Button
        Me.cmd0 = New System.Windows.Forms.Button
        Me.cmdMultiply = New System.Windows.Forms.Button
        Me.cmdDivide = New System.Windows.Forms.Button
        Me.cmd7 = New System.Windows.Forms.Button
        Me.cmdClearAll = New System.Windows.Forms.Button
        Me.cmdClearEntry = New System.Windows.Forms.Button
        Me.cmdBackspace = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSign
        '
        Me.cmdSign.Location = New System.Drawing.Point(61, 147)
        Me.cmdSign.Name = "cmdSign"
        Me.cmdSign.Size = New System.Drawing.Size(35, 23)
        Me.cmdSign.TabIndex = 26
        Me.cmdSign.Text = "+/-"
        Me.cmdSign.UseVisualStyleBackColor = True
        '
        'cmdInverse
        '
        Me.cmdInverse.Location = New System.Drawing.Point(184, 118)
        Me.cmdInverse.Name = "cmdInverse"
        Me.cmdInverse.Size = New System.Drawing.Size(35, 23)
        Me.cmdInverse.TabIndex = 25
        Me.cmdInverse.Text = "1/x"
        Me.cmdInverse.UseVisualStyleBackColor = True
        '
        'cmdSqrRoot
        '
        Me.cmdSqrRoot.Location = New System.Drawing.Point(184, 89)
        Me.cmdSqrRoot.Name = "cmdSqrRoot"
        Me.cmdSqrRoot.Size = New System.Drawing.Size(35, 23)
        Me.cmdSqrRoot.TabIndex = 24
        Me.cmdSqrRoot.Text = "sqrt"
        Me.cmdSqrRoot.UseVisualStyleBackColor = True
        '
        'cmdPowerOf
        '
        Me.cmdPowerOf.Location = New System.Drawing.Point(184, 60)
        Me.cmdPowerOf.Name = "cmdPowerOf"
        Me.cmdPowerOf.Size = New System.Drawing.Size(35, 23)
        Me.cmdPowerOf.TabIndex = 23
        Me.cmdPowerOf.Text = "x^"
        Me.cmdPowerOf.UseVisualStyleBackColor = True
        '
        'cmdEqual
        '
        Me.cmdEqual.Location = New System.Drawing.Point(184, 147)
        Me.cmdEqual.Name = "cmdEqual"
        Me.cmdEqual.Size = New System.Drawing.Size(35, 23)
        Me.cmdEqual.TabIndex = 22
        Me.cmdEqual.Text = "="
        Me.cmdEqual.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(143, 147)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(35, 23)
        Me.cmdAdd.TabIndex = 21
        Me.cmdAdd.Text = "+"
        Me.cmdAdd.UseVisualStyleBackColor = True
        '
        'cmdDecimal
        '
        Me.cmdDecimal.Location = New System.Drawing.Point(102, 147)
        Me.cmdDecimal.Name = "cmdDecimal"
        Me.cmdDecimal.Size = New System.Drawing.Size(35, 23)
        Me.cmdDecimal.TabIndex = 20
        Me.cmdDecimal.Text = "."
        Me.cmdDecimal.UseVisualStyleBackColor = True
        '
        'txtInput
        '
        Me.txtInput.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.txtInput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInput.Location = New System.Drawing.Point(6, 16)
        Me.txtInput.Name = "txtInput"
        Me.txtInput.ReadOnly = True
        Me.txtInput.Size = New System.Drawing.Size(223, 20)
        Me.txtInput.TabIndex = 0
        Me.txtInput.TabStop = False
        Me.txtInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cmdSubtract
        '
        Me.cmdSubtract.Location = New System.Drawing.Point(143, 118)
        Me.cmdSubtract.Name = "cmdSubtract"
        Me.cmdSubtract.Size = New System.Drawing.Size(35, 23)
        Me.cmdSubtract.TabIndex = 19
        Me.cmdSubtract.Text = "-"
        Me.cmdSubtract.UseVisualStyleBackColor = True
        '
        'cmd8
        '
        Me.cmd8.Location = New System.Drawing.Point(61, 60)
        Me.cmd8.Name = "cmd8"
        Me.cmd8.Size = New System.Drawing.Size(35, 23)
        Me.cmd8.TabIndex = 18
        Me.cmd8.Text = "8"
        Me.cmd8.UseVisualStyleBackColor = True
        '
        'cmd9
        '
        Me.cmd9.Location = New System.Drawing.Point(102, 60)
        Me.cmd9.Name = "cmd9"
        Me.cmd9.Size = New System.Drawing.Size(35, 23)
        Me.cmd9.TabIndex = 17
        Me.cmd9.Text = "9"
        Me.cmd9.UseVisualStyleBackColor = True
        '
        'cmd4
        '
        Me.cmd4.Location = New System.Drawing.Point(20, 89)
        Me.cmd4.Name = "cmd4"
        Me.cmd4.Size = New System.Drawing.Size(35, 23)
        Me.cmd4.TabIndex = 16
        Me.cmd4.Text = "4"
        Me.cmd4.UseVisualStyleBackColor = True
        '
        'cmd5
        '
        Me.cmd5.Location = New System.Drawing.Point(61, 89)
        Me.cmd5.Name = "cmd5"
        Me.cmd5.Size = New System.Drawing.Size(35, 23)
        Me.cmd5.TabIndex = 15
        Me.cmd5.Text = "5"
        Me.cmd5.UseVisualStyleBackColor = True
        '
        'cmd6
        '
        Me.cmd6.Location = New System.Drawing.Point(102, 89)
        Me.cmd6.Name = "cmd6"
        Me.cmd6.Size = New System.Drawing.Size(35, 23)
        Me.cmd6.TabIndex = 14
        Me.cmd6.Text = "6"
        Me.cmd6.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdSign)
        Me.GroupBox2.Controls.Add(Me.cmdInverse)
        Me.GroupBox2.Controls.Add(Me.cmdSqrRoot)
        Me.GroupBox2.Controls.Add(Me.cmdPowerOf)
        Me.GroupBox2.Controls.Add(Me.cmdEqual)
        Me.GroupBox2.Controls.Add(Me.cmdAdd)
        Me.GroupBox2.Controls.Add(Me.cmdDecimal)
        Me.GroupBox2.Controls.Add(Me.cmdSubtract)
        Me.GroupBox2.Controls.Add(Me.cmd8)
        Me.GroupBox2.Controls.Add(Me.cmd9)
        Me.GroupBox2.Controls.Add(Me.cmd4)
        Me.GroupBox2.Controls.Add(Me.cmd5)
        Me.GroupBox2.Controls.Add(Me.cmd6)
        Me.GroupBox2.Controls.Add(Me.cmd1)
        Me.GroupBox2.Controls.Add(Me.cmd2)
        Me.GroupBox2.Controls.Add(Me.cmd3)
        Me.GroupBox2.Controls.Add(Me.cmd0)
        Me.GroupBox2.Controls.Add(Me.cmdMultiply)
        Me.GroupBox2.Controls.Add(Me.cmdDivide)
        Me.GroupBox2.Controls.Add(Me.cmd7)
        Me.GroupBox2.Controls.Add(Me.cmdClearAll)
        Me.GroupBox2.Controls.Add(Me.cmdClearEntry)
        Me.GroupBox2.Controls.Add(Me.cmdBackspace)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(235, 187)
        Me.GroupBox2.TabIndex = 8
        Me.GroupBox2.TabStop = False
        '
        'cmd1
        '
        Me.cmd1.Location = New System.Drawing.Point(20, 118)
        Me.cmd1.Name = "cmd1"
        Me.cmd1.Size = New System.Drawing.Size(35, 23)
        Me.cmd1.TabIndex = 13
        Me.cmd1.Text = "1"
        Me.cmd1.UseVisualStyleBackColor = True
        '
        'cmd2
        '
        Me.cmd2.Location = New System.Drawing.Point(61, 118)
        Me.cmd2.Name = "cmd2"
        Me.cmd2.Size = New System.Drawing.Size(35, 23)
        Me.cmd2.TabIndex = 12
        Me.cmd2.Text = "2"
        Me.cmd2.UseVisualStyleBackColor = True
        '
        'cmd3
        '
        Me.cmd3.Location = New System.Drawing.Point(102, 118)
        Me.cmd3.Name = "cmd3"
        Me.cmd3.Size = New System.Drawing.Size(35, 23)
        Me.cmd3.TabIndex = 11
        Me.cmd3.Text = "3"
        Me.cmd3.UseVisualStyleBackColor = True
        '
        'cmd0
        '
        Me.cmd0.Location = New System.Drawing.Point(20, 147)
        Me.cmd0.Name = "cmd0"
        Me.cmd0.Size = New System.Drawing.Size(35, 23)
        Me.cmd0.TabIndex = 10
        Me.cmd0.Text = "0"
        Me.cmd0.UseVisualStyleBackColor = True
        '
        'cmdMultiply
        '
        Me.cmdMultiply.Location = New System.Drawing.Point(143, 89)
        Me.cmdMultiply.Name = "cmdMultiply"
        Me.cmdMultiply.Size = New System.Drawing.Size(35, 23)
        Me.cmdMultiply.TabIndex = 9
        Me.cmdMultiply.Text = "*"
        Me.cmdMultiply.UseVisualStyleBackColor = True
        '
        'cmdDivide
        '
        Me.cmdDivide.Location = New System.Drawing.Point(143, 60)
        Me.cmdDivide.Name = "cmdDivide"
        Me.cmdDivide.Size = New System.Drawing.Size(35, 23)
        Me.cmdDivide.TabIndex = 8
        Me.cmdDivide.Text = "/"
        Me.cmdDivide.UseVisualStyleBackColor = True
        '
        'cmd7
        '
        Me.cmd7.Location = New System.Drawing.Point(20, 60)
        Me.cmd7.Name = "cmd7"
        Me.cmd7.Size = New System.Drawing.Size(35, 23)
        Me.cmd7.TabIndex = 7
        Me.cmd7.Text = "7"
        Me.cmd7.UseVisualStyleBackColor = True
        '
        'cmdClearAll
        '
        Me.cmdClearAll.Location = New System.Drawing.Point(165, 19)
        Me.cmdClearAll.Name = "cmdClearAll"
        Me.cmdClearAll.Size = New System.Drawing.Size(54, 23)
        Me.cmdClearAll.TabIndex = 6
        Me.cmdClearAll.Text = "C"
        Me.cmdClearAll.UseVisualStyleBackColor = True
        '
        'cmdClearEntry
        '
        Me.cmdClearEntry.Location = New System.Drawing.Point(103, 19)
        Me.cmdClearEntry.Name = "cmdClearEntry"
        Me.cmdClearEntry.Size = New System.Drawing.Size(56, 23)
        Me.cmdClearEntry.TabIndex = 5
        Me.cmdClearEntry.Text = "CE"
        Me.cmdClearEntry.UseVisualStyleBackColor = True
        '
        'cmdBackspace
        '
        Me.cmdBackspace.Location = New System.Drawing.Point(18, 19)
        Me.cmdBackspace.Name = "cmdBackspace"
        Me.cmdBackspace.Size = New System.Drawing.Size(76, 23)
        Me.cmdBackspace.TabIndex = 4
        Me.cmdBackspace.Text = "Backspace"
        Me.cmdBackspace.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtInput)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 29)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(235, 45)
        Me.GroupBox1.TabIndex = 7
        Me.GroupBox1.TabStop = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(249, 24)
        Me.MenuStrip1.TabIndex = 9
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'Calculator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Name = "Calculator"
        Me.Size = New System.Drawing.Size(249, 280)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSign As System.Windows.Forms.Button
    Friend WithEvents cmdInverse As System.Windows.Forms.Button
    Friend WithEvents cmdSqrRoot As System.Windows.Forms.Button
    Friend WithEvents cmdPowerOf As System.Windows.Forms.Button
    Friend WithEvents cmdEqual As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cmdDecimal As System.Windows.Forms.Button
    Friend WithEvents txtInput As System.Windows.Forms.TextBox
    Friend WithEvents cmdSubtract As System.Windows.Forms.Button
    Friend WithEvents cmd8 As System.Windows.Forms.Button
    Friend WithEvents cmd9 As System.Windows.Forms.Button
    Friend WithEvents cmd4 As System.Windows.Forms.Button
    Friend WithEvents cmd5 As System.Windows.Forms.Button
    Friend WithEvents cmd6 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmd1 As System.Windows.Forms.Button
    Friend WithEvents cmd2 As System.Windows.Forms.Button
    Friend WithEvents cmd3 As System.Windows.Forms.Button
    Friend WithEvents cmd0 As System.Windows.Forms.Button
    Friend WithEvents cmdMultiply As System.Windows.Forms.Button
    Friend WithEvents cmdDivide As System.Windows.Forms.Button
    Friend WithEvents cmd7 As System.Windows.Forms.Button
    Friend WithEvents cmdClearAll As System.Windows.Forms.Button
    Friend WithEvents cmdClearEntry As System.Windows.Forms.Button
    Friend WithEvents cmdBackspace As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip

End Class
