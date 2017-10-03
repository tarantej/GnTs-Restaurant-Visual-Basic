Public Class Calculator

    'variables to hold operands
    Private valHolder1 As Double
    Private valHolder2 As Double
    'Varible to hold temporary values
    Private tmpValue As Double
    'variable for Memory storage
    Private Memory As Double
    'True if "." is use else false
    Private hasDecimal As Boolean
    Private inputStatus As Boolean
    Private clearText As Boolean
    'variable to hold Operater
    Private calcFunc As String

#Region "Number Buttons "
    Private Sub cmd9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd9.Click
        If inputStatus = False Then
            txtInput.Text += cmd9.Text
        Else
            txtInput.Text = cmd9.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd8.Click
        If inputStatus = False Then
            txtInput.Text += cmd8.Text
        Else
            txtInput.Text = cmd8.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd7.Click
        If inputStatus = False Then
            txtInput.Text += cmd7.Text
        Else
            txtInput.Text = cmd7.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd6.Click
        If inputStatus = False Then
            txtInput.Text += cmd6.Text
        Else
            txtInput.Text = cmd6.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd5.Click
        If inputStatus = False Then
            txtInput.Text += cmd5.Text
        Else
            txtInput.Text = cmd5.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd4.Click
        If inputStatus = False Then
            txtInput.Text += cmd4.Text
        Else
            txtInput.Text = cmd4.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd3.Click
        If inputStatus = False Then
            txtInput.Text += cmd3.Text
        Else
            txtInput.Text = cmd3.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd2.Click
        If inputStatus = False Then
            txtInput.Text += cmd2.Text
        Else
            txtInput.Text = cmd2.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd1.Click
        If inputStatus = False Then
            txtInput.Text += cmd1.Text
        Else
            txtInput.Text = cmd1.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd0.Click
        If inputStatus = False Then
            If txtInput.Text.Length >= 1 Then
                txtInput.Text += cmd0.Text
            End If
        End If
    End Sub
#End Region


#Region " Calculation Buttons "
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Add"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdSubtract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubtract.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Subtract"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdDivide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDivide.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Divide"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdMultiply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMultiply.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Multiply"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdPowerOf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPowerOf.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "PowerOf"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdSqrRoot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSqrRoot.Click
        If txtInput.Text.Length <> 0 Then
            tmpValue = CDbl(txtInput.Text)
            tmpValue = System.Math.Sqrt(tmpValue)
            txtInput.Text = CStr(tmpValue)
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdEqual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEqual.Click
        If txtInput.Text.Length <> 0 AndAlso valHolder1 <> 0 Then
            CalculateTotals()
            calcFunc = ""
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdInverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInverse.Click
        If txtInput.Text.Length <> 0 Then
            tmpValue = CDbl(txtInput.Text)
            tmpValue = 1 / tmpValue
            txtInput.Text = CStr(tmpValue)
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdSign_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSign.Click
        If inputStatus = False Then
            If txtInput.Text.Length > 0 Then
                valHolder1 = -1 * CDbl(txtInput.Text)
                txtInput.Text = CStr(valHolder1)
            End If
        End If
    End Sub
#End Region

#Region " Other Buttons "
    Private Sub cmdClearEntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearEntry.Click
        txtInput.Text = String.Empty
        hasDecimal = False
    End Sub

    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click
        txtInput.Text = String.Empty
        valHolder1 = 0
        valHolder2 = 0
        calcFunc = String.Empty
        hasDecimal = False
    End Sub

    Private Sub cmdBackspace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBackspace.Click
        Dim str As String
        Dim loc As Integer
        If txtInput.Text.Length > 0 Then
            str = txtInput.Text.Chars(txtInput.Text.Length - 1)
            If str = "." Then
                hasDecimal = False
            End If
            loc = txtInput.Text.Length
            txtInput.Text = txtInput.Text.Remove(loc - 1, 1)
        End If
    End Sub

    Private Sub cmdDecimal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDecimal.Click
        'Check for input status (we want flase)
        If Not inputStatus Then
            'Check if it already has a decimal (if it does then do nothing)
            If Not hasDecimal Then
                'Check to make sure the length is > than 1
                'Dont want user to add decimal as first character
                If txtInput.Text.Length > 1 Then
                    'Make sure 0 isnt the first number
                    If Not txtInput.Text = "0" Then
                        'It met all our requirements so add the zero
                        txtInput.Text += cmdDecimal.Text
                        'Toggle the flag to true (only 1 decimal per calculation)
                        hasDecimal = True
                    End If
                Else
                    txtInput.Text = "0."
                End If
            End If
        End If
    End Sub
#End Region

#Region " Helpers "
    Private Sub CalculateTotals()
        valHolder2 = CDbl(txtInput.Text)
        Select Case calcFunc
            Case "Add"
                valHolder1 = valHolder1 + valHolder2
            Case "Subtract"
                valHolder1 = valHolder1 - valHolder2
            Case "Divide"
                valHolder1 = valHolder1 / valHolder2
            Case "Multiply"
                valHolder1 = valHolder1 * valHolder2
            Case "PowerOf"
                valHolder1 = System.Math.Pow(valHolder1, valHolder2)
        End Select
        txtInput.Text = CStr(valHolder1)
        inputStatus = True
    End Sub
#End Region


End Class
