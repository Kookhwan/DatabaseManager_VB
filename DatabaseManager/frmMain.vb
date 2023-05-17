Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Public Class frmMain

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call psubLoadDatabase()
    End Sub

    Private Sub psubLoadDatabase()

        Dim conLocalSQL As New SqlConnection("Data Source=DBSERVER\AWTESTZONE17;Persist Security Info=True;User ID=sa;Password=&&AW1975&&")
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try

            conLocalSQL.Open()

            strQuery = ""
            strQuery = strQuery & "SELECT "
            strQuery = strQuery & "    database_id AS [ID], "
            strQuery = strQuery & "    name AS [DB Name], "
            strQuery = strQuery & "    create_date AS [Created DateTime] "
            strQuery = strQuery & "FROM "
            strQuery = strQuery & "    sys.databases "
            strQuery = strQuery & "WHERE "
            strQuery = strQuery & "    name LIKE 'ArtwoodsSQL%'"
            strQuery = strQuery & "ORDER BY "
            strQuery = strQuery & "    create_date ASC "

            dgvDatabase.DataSource = gfunDataTableSQL(strQuery, conLocalSQL)

            conLocalSQL.Close()

            dgvDatabase.Columns(0).Width = 60
            dgvDatabase.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvDatabase.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvDatabase.Columns(1).Width = 150
            dgvDatabase.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            dgvDatabase.Columns(2).Width = 150
            dgvDatabase.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvDatabase.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try

    End Sub

End Class
