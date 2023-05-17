Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Public Class frmMain

    Private m_dtProgress As New DataTable

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call psubDefineTable()
        Call psubGridDatabase()
        Call psubCmbDatabase1()
        Call psubCmbDatabase2()
    End Sub

    Private Sub psubDefineTable()

        Try

            dgvProgress.DataSource = Nothing

            m_dtProgress.Columns.Add("No", GetType(Integer))
            m_dtProgress.Columns.Add("Process Name", GetType(String))
            m_dtProgress.Columns.Add("Status", GetType(String))           '// Pending, Running, Finished

            m_dtProgress.Rows.Add(1, "Create a new backup for Local", "Pending")
            m_dtProgress.Rows.Add(2, "Restore Database for Local", "Pending")
            m_dtProgress.Rows.Add(3, "Log Clear for Local", "Pending")
            m_dtProgress.Rows.Add(4, "Copy Access DB", "Pending")
            m_dtProgress.Rows.Add(5, "Relink SQL Talbles", "Pending")
            m_dtProgress.Rows.Add(6, "Relink SQL Queries", "Pending")
            m_dtProgress.Rows.Add(7, "Create a new backup for Web", "Pending")
            m_dtProgress.Rows.Add(8, "Restore Database for Web", "Pending")
            m_dtProgress.Rows.Add(9, "Log Clear for Web", "Pending")

            dgvProgress.DataSource = m_dtProgress

            dgvProgress.Columns(0).Width = 40
            dgvProgress.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(1).Width = 180
            dgvProgress.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(2).Width = 100
            dgvProgress.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Catch ex As Exception
            MsgBox(ex.Message & " in psubDefineTable()")
        End Try

    End Sub

    Private Sub psubGridDatabase()

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
            MsgBox(ex.Message & " in psubGridDatabase()")

        End Try

    End Sub

    Private Sub psubCmbDatabase1()

        Dim conLocalSQL As New SqlConnection("Data Source=DBSERVER;Persist Security Info=True;User ID=sa;Password=&&AW1975&&")
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try
            cmbDatabase1.Items.Clear()

            conLocalSQL.Open()

            strQuery = ""
            strQuery = strQuery & "SELECT name "
            strQuery = strQuery & "FROM   sys.databases "
            strQuery = strQuery & "WHERE  name LIKE 'ArtwoodsSQL%'"
            strQuery = strQuery & "ORDER BY create_date ASC "

            sqlReader = gfunDataReaderSQL(strQuery, conLocalSQL)

            If sqlReader IsNot Nothing AndAlso sqlReader.HasRows Then

                While sqlReader.Read
                    cmbDatabase1.Items.Add(sqlReader.GetValue(sqlReader.GetOrdinal("name")))
                End While

                cmbDatabase1.Text = "Select database"

            End If

            conLocalSQL.Close()

        Catch ex As Exception
            MsgBox(ex.Message & " in psubCmbDatabase1()")

        End Try

    End Sub

    Private Sub psubCmbDatabase2()

        Dim conLocalSQL As New SqlConnection("Data Source=DBSERVER;Persist Security Info=True;User ID=sa;Password=&&AW1975&&")
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try
            cmbDatabase2.Items.Clear()

            conLocalSQL.Open()

            strQuery = "EXEC xp_dirtree 'C:\ArtwoodsSQL_Backups', 0, 1"

            sqlReader = gfunDataReaderSQL(strQuery, conLocalSQL)

            If sqlReader IsNot Nothing AndAlso sqlReader.HasRows Then

                While sqlReader.Read
                    cmbDatabase2.Items.Add(sqlReader.GetValue(sqlReader.GetOrdinal("subdirectory")))
                End While

                cmbDatabase2.Text = "Select database"

            End If

            conLocalSQL.Close()

        Catch ex As Exception
            MsgBox(ex.Message & " in psubCmbDatabase2()")

        End Try

    End Sub

    Private Sub tabDatabase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabDatabase.SelectedIndexChanged

        Try
            If tabDatabase.SelectedIndex = 1 Then
                m_dtProgress.Rows(0)("Status") = "N/A"
            Else
                m_dtProgress.Rows(0)("Status") = "Pending"
            End If

            dgvProgress.DataSource = m_dtProgress

        Catch ex As Exception
            MsgBox(ex.Message & " in tabDatabase_SelectedIndexChanged()")
        End Try

    End Sub

    Private Sub btnCreateBackup1_Click(sender As Object, e As EventArgs) Handles btnCreateBackup1.Click

        Try

        Catch ex As Exception
            MsgBox(ex.Message & " in btnCreateBackup1_Click()")
        End Try

    End Sub


End Class
