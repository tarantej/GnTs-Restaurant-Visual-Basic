Option Strict Off
Option Explicit On
Imports System.Data.SqlClient
Public Class Login
    Inherits System.Windows.Forms.Form
    Private Sub cmdLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLogin.Click
        Dim ConnectionString As String
        Dim sql As String
        Dim conn As SqlConnection
        Dim cmd As SqlCommand
        ConnectionString = "Data Source=TARANTEJSINGH\SQLEXPRESS;Initial Catalog=gnt;Persist Security Info=True;User ID=sa; Password=broodwar"
        conn = New SqlConnection(ConnectionString)
        sql = "SELECT Username, Password from NewLogin where Username = '" & txtUsername.Text & "' and Password = '" & txtPassword.Text & "'"
        cmd = New SqlCommand(sql, conn)
        conn.Open()
        If txtuser.Text <> "tarantej" Or txtpass.Text <> "singh" Then
            Label2.ForeColor = Color.Red
            Label3.Text = "Invalid USERNAME OR  PASSWORD. Please Try Again"
            txtuser.Text = ""
            txtpass.Text = ""
        Else
            Label2.ForeColor = Color.Green
            Label3.Text = "Login successful..."
            Admin_Menu.Show()
            Me.Close()
        End If

        conn.Close()





    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Close()

    End Sub
End Class