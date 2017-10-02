<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class New_Emp
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
        Dim ListViewGroup1 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Academic", System.Windows.Forms.HorizontalAlignment.Left)
        Dim ListViewGroup2 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Academic", System.Windows.Forms.HorizontalAlignment.Left)
        Dim ListViewGroup3 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Academic", System.Windows.Forms.HorizontalAlignment.Left)
        Dim ListViewGroup4 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Professional", System.Windows.Forms.HorizontalAlignment.Left)
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.Label11 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.Year = New System.Windows.Forms.ColumnHeader
        Me.Education = New System.Windows.Forms.ColumnHeader
        Me.EducationBoard = New System.Windows.Forms.ColumnHeader
        Me.University = New System.Windows.Forms.ColumnHeader
        Me.Percentage = New System.Windows.Forms.ColumnHeader
        Me.SuspendLayout()
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(343, 442)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(200, 20)
        Me.DateTimePicker1.TabIndex = 38
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(343, 401)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(139, 20)
        Me.TextBox6.TabIndex = 37
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(343, 369)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(139, 20)
        Me.TextBox5.TabIndex = 36
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(325, 269)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(218, 79)
        Me.RichTextBox1.TabIndex = 33
        Me.RichTextBox1.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(326, 21)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(181, 20)
        Me.TextBox2.TabIndex = 31
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(325, 83)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 30
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(239, 449)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(66, 13)
        Me.Label9.TabIndex = 28
        Me.Label9.Text = "Joining Date"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(239, 408)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(98, 13)
        Me.Label8.TabIndex = 27
        Me.Label8.Text = "Residence Number"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(239, 376)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(78, 13)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "Mobile Number"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(239, 272)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 13)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Address"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(236, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Employee Name"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(239, 90)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(26, 13)
        Me.Label1.TabIndex = 20
        Me.Label1.Text = "Age"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(239, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 22
        Me.Label3.Text = "Gender"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(328, 60)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(48, 17)
        Me.RadioButton1.TabIndex = 40
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Male"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(382, 60)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(59, 17)
        Me.RadioButton2.TabIndex = 41
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Female"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(236, 116)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(71, 13)
        Me.Label11.TabIndex = 42
        Me.Label11.Text = "Marital Status"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(325, 113)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 43
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(232, 152)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(70, 13)
        Me.Label5.TabIndex = 44
        Me.Label5.Text = "Qualifications"
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Year, Me.Education, Me.EducationBoard, Me.University, Me.Percentage})
        ListViewGroup1.Header = "Academic"
        ListViewGroup1.Name = "HighSchool"
        ListViewGroup2.Header = "Academic"
        ListViewGroup2.Name = "Graduation"
        ListViewGroup3.Header = "Academic"
        ListViewGroup3.Name = "Post Graduation"
        ListViewGroup4.Header = "Professional"
        ListViewGroup4.Name = "Company"
        Me.ListView1.Groups.AddRange(New System.Windows.Forms.ListViewGroup() {ListViewGroup1, ListViewGroup2, ListViewGroup3, ListViewGroup4})
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.LabelEdit = True
        Me.ListView1.Location = New System.Drawing.Point(325, 152)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(408, 97)
        Me.ListView1.TabIndex = 45
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'Year
        '
        Me.Year.Text = "Year"
        '
        'Education
        '
        Me.Education.Text = "Education"
        '
        'EducationBoard
        '
        Me.EducationBoard.DisplayIndex = 4
        Me.EducationBoard.Text = "Education Board"
        '
        'University
        '
        Me.University.DisplayIndex = 2
        Me.University.Text = "University"
        '
        'Percentage
        '
        Me.Percentage.DisplayIndex = 3
        Me.Percentage.Text = "Percentage"
        '
        'New_Emp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(843, 570)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.RadioButton2)
        Me.Controls.Add(Me.RadioButton1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "New_Emp"
        Me.Text = "New_Emp"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents Year As System.Windows.Forms.ColumnHeader
    Friend WithEvents Education As System.Windows.Forms.ColumnHeader
    Friend WithEvents University As System.Windows.Forms.ColumnHeader
    Friend WithEvents Percentage As System.Windows.Forms.ColumnHeader
    Friend WithEvents EducationBoard As System.Windows.Forms.ColumnHeader
End Class
