Imports System.Data.SqlClient

Public Class frmDatabaseLogin

    Private Sub btnLogIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogIn.Click

        strConnection = "Data Source=" & txtServer.Text & "; " & _
                        "Initial Catalog=" & txtDatabase.Text & "; " & _
                        "User ID=" & txtUsername.Text & "; " & _
                        "Password=" & txtPassword.Text

        Dim cn As SqlConnection = New SqlConnection(strConnection)

        Try
            cn.Open()
            cn.Close()
        Catch ex As Exception
            strConnection = ""
        End Try
        Close()
    End Sub

    Private Sub frmDatabaseLogin_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

#If DEBUG Then
        Dim strConn As String = My.Settings.strTest
#Else
        Dim strConn As String = My.Settings.GBSConnectionString
#End If

        strConn = strConn.Substring(strConn.IndexOf("Source=") + Len("Source="))
        txtServer.Text = strConn.Substring(0, strConn.IndexOf(";"))

        strConn = strConn.Substring(strConn.IndexOf("Catalog=") + Len("Catalog="))
        txtDatabase.Text = strConn.Substring(0, strConn.IndexOf(";"))

        strConn = strConn.Substring(strConn.IndexOf("ID=") + Len("ID="))
        txtUsername.Text = strConn.Substring(0, strConn.IndexOf(";"))

        strConn = strConn.Substring(strConn.IndexOf("Password=") + Len("Password="))
        txtPassword.Text = strConn
    End Sub

End Class