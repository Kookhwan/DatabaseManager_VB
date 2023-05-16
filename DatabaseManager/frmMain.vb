Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Public Class frmMain

    Private Sub psubLoadDatabase()

        Dim conLocalSQL As New SqlConnection("Data Source=DBSERVER\AWTESTZONE17;Persist Security Info=True;User ID=sa;Password=&&AW1975&&")
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try

            conLocalSQL.Open()

            strQuery = ""
            strQuery = strQuery & "SELECT name "
            strQuery = strQuery & "FROM   sys.databases "
            strQuery = strQuery & "WHERE  name LIKE 'ArtwoodsSQL%'"
            strQuery = strQuery & "ORDER BY create_date ASC "

            dgvDatabase.DataSource = gfunDataTableSQL(strQuery, conLocalSQL)

            conLocalSQL.Close()

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try




    End Sub



End Class
