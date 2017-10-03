
Option Strict Off
Option Explicit On
Imports System.Data.SqlClient




Public Class LoginForm1
    Inherits System.Windows.Forms.Form
    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim ConnectionString As String
        Dim sql As String
        Dim conn As SqlConnection
        Dim cmd As SqlCommand
        ConnectionString = "Data Source=TARANTEJSINGH\SQLEXPRESS;Initial Catalog=login;Persist Security Info=True;User ID=sa"
        conn = New SqlConnection(ConnectionString)
        sql = "SELECT Username, Password from NewLogin where Username = '" & txtUsername.Text & "' and Password = '" & txtPassword.Text & "'"
        cmd = New SqlCommand(sql, conn)
        conn.Open()
        If txtUsername.Text <> "tarantej" Or txtPassword.Text <> "singh" Then
            Label2.ForeColor = Color.Red
            Label3.Text = "Invalid USERNAME OR  PASSWORD. Please Try Again"
            txtUsername.Text = ""
            txtPassword.Text = ""
        Else
            Label2.ForeColor = Color.Green
            Label3.Text = "Login successful..."
            Form1.Show()
            Me.Close()
        End If

        conn.Close()



    End Sub
   

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
