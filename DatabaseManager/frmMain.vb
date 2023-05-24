﻿Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Threading

Public Class frmMain

    Private Const MAX_PROCESS = 8

    Private m_dtProgress As New DataTable
    Private oThread As Thread
    Private nProcess As Integer

    Private m_strDate As String
    Private m_dtTarget As Date
    Private m_strSourceDB As String
    Private m_strTargetDB As String



    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call gsubLoadXML_Info()
        Call psubDefineTable()
        Call psubGridDatabase()
        Call psubCmbDatabase1()
        Call psubCmbDatabase2()
    End Sub

    Private Sub frmMain_Closed(sender As Object, e As EventArgs) Handles Me.Closed

        If gbStart = True Then
            Try
                gbStart = False
                oThread.Abort()
            Catch ex As Exception
                MsgBox(ex.Message & " in frmMain_Closed()")
            End Try
        End If

    End Sub

    Private Sub psubDefineTable()

        Try

            dgvProgress.DataSource = Nothing

            m_dtProgress.Columns.Add("No", GetType(Integer))
            m_dtProgress.Columns.Add("Process Name", GetType(String))
            m_dtProgress.Columns.Add("Status", GetType(String))           '// Pending, Running, Finished

            m_dtProgress.Rows.Add(1, "Create a new backup for Local", "Pending")
            m_dtProgress.Rows.Add(2, "Backup Access DB", "Pending")
            m_dtProgress.Rows.Add(3, "Restore Database for Local", "Pending")
            m_dtProgress.Rows.Add(4, "Log Clear for Local", "Pending")
            m_dtProgress.Rows.Add(5, "ReLink Access Objects", "Pending")
            m_dtProgress.Rows.Add(6, "Create a new backup for Web", "Pending")
            m_dtProgress.Rows.Add(7, "Restore Database for Web", "Pending")
            m_dtProgress.Rows.Add(8, "Log Clear for Web", "Pending")

            dgvProgress.DataSource = m_dtProgress

            dgvProgress.Columns(0).Width = 40
            dgvProgress.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(1).Width = 210
            dgvProgress.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(2).Width = 80
            dgvProgress.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvProgress.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Catch ex As Exception
            MsgBox(ex.Message & " in psubDefineTable()")
        End Try

    End Sub

    Private Sub psubGridDatabase()

        Dim conLocalSQL As New SqlConnection
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try

            Call gsubSqlConnectLocal(conLocalSQL, "master", True)

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

            Call gsubSqlDisconnet(conLocalSQL)

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

        Dim conLocalSQL As New SqlConnection
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try
            cmbDatabase1.Items.Clear()

            Call gsubSqlConnectLocal(conLocalSQL, "master")

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

            Call gsubSqlDisconnet(conLocalSQL)

        Catch ex As Exception
            MsgBox(ex.Message & " in psubCmbDatabase1()")

        End Try

    End Sub

    Private Sub psubCmbDatabase2()

        Dim conLocalSQL As New SqlConnection
        Dim sqlReader As SqlDataReader = Nothing
        Dim strQuery As String

        Try
            cmbDatabase2.Items.Clear()

            Call gsubSqlConnectLocal(conLocalSQL, "master")

            strQuery = "EXEC xp_dirtree 'C:\ArtwoodsSQL_Backups', 0, 1"

            sqlReader = gfunDataReaderSQL(strQuery, conLocalSQL)

            If sqlReader IsNot Nothing AndAlso sqlReader.HasRows Then

                While sqlReader.Read
                    cmbDatabase2.Items.Add(sqlReader.GetValue(sqlReader.GetOrdinal("subdirectory")))
                End While

                cmbDatabase2.Text = "Select a backup file"

            End If

            Call gsubSqlDisconnet(conLocalSQL)

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

        If cmbDatabase1.Text = "Select database" Then
            MessageBox.Show("Please Select a database", "Warning")
            Exit Sub
        End If

        m_strDate = Format(Now(), "_M_d")
        m_strSourceDB = cmbDatabase1.Text
        m_strTargetDB = m_strSourceDB & m_strDate

        If gfunIsExistDB(m_strTargetDB) Then
            If MessageBox.Show("There is already a database with the same name." & vbCrLf &
                               "Would you Like To overwrite On it?", "Duplicate DB", MessageBoxButtons.YesNo) = vbNo Then
                Exit Sub
            End If
        End If

        Call psubSetWaitCursor(True)

        nProcess = 0

        If gbStart = False Then

            Try

                oThread = New Thread(AddressOf psubWorkThread)
                gbStart = True
                oThread.Start()

            Catch ex As Exception
                MsgBox(ex.Message & " In btnCreateBackup1_Click()")
            End Try

        End If

    End Sub

    Private Sub btnCreateBackup2_Click(sender As Object, e As EventArgs) Handles btnCreateBackup2.Click

        If cmbDatabase2.Text = "Select a backup file" Then
            MessageBox.Show("Please Select a backup file", "Warning")
            Exit Sub
        End If

        m_dtTarget = Mid(cmbDatabase2.Text, 19, 10)

        m_strDate = Format(m_dtTarget, "_M_d")
        m_strSourceDB = Mid(cmbDatabase2.Text, 1, 11)
        m_strTargetDB = m_strSourceDB & m_strDate

        If gfunIsExistDB(m_strTargetDB) Then
            If MessageBox.Show("There is already a database with the same name." & vbCrLf &
                               "Would you Like To overwrite On it?", "Duplicate DB", MessageBoxButtons.YesNo) = vbNo Then
                Exit Sub
            End If
        End If

        Call psubSetWaitCursor(True)

        nProcess = 1

        If gbStart = False Then

            Try

                oThread = New Thread(AddressOf psubWorkThread)
                gbStart = True
                oThread.Start()

            Catch ex As Exception
                MsgBox(ex.Message & " In btnCreateBackup2_Click() Then")
            End Try

        End If

    End Sub

    Private Sub psubWorkThread()

        While (gbStart)

            nProcess += 1

            If (Me.InvokeRequired) Then
                Me.Invoke(Sub()

                              Call psubSetGridProgress(nProcess, "Running")

                              If pfunBackupProcess(nProcess) = True Then
                                  Call psubSetGridProgress(nProcess, "Finished")
                              Else
                                  Call psubSetGridProgress(nProcess, "failed")
                                  gbStart = False
                              End If

                          End Sub)

            End If

            If nProcess >= 8 Then

                If (Me.InvokeRequired) Then
                    Me.Invoke(Sub()
                                  MsgBox("Done")
                                  Call psubSetWaitCursor(False)
                              End Sub)
                End If

                gbStart = False
                oThread.Abort()

            End If

            Thread.Sleep(500)

        End While

    End Sub

    Private Sub psubResetGridProcess()

        Dim nStep As Integer

        For nStep = 0 To 7

            If nStep = 0 Then
                m_dtProgress.Rows(nStep)("Status") = IIf(tabDatabase.SelectedIndex = 1, "N/A", "Pending")
            Else
                m_dtProgress.Rows(nStep)("Status") = "Pending"
            End If

        Next

    End Sub
    Private Sub psubSetGridProgress(ByVal nStep As Integer, ByVal strStatus As String)

        If nStep <= 0 Then
            Exit Sub
        End If

        m_dtProgress.Rows(nStep - 1)("Status") = strStatus

        dgvProgress.DataSource = m_dtProgress

        dgvProgress.ClearSelection()
        dgvProgress.Rows(nStep - 1).Selected = True

        pbProgress.Value = nStep

        Application.DoEvents()

    End Sub

    Private Function pfunBackupProcess(ByVal nIndex As Integer) As Boolean

        Dim bProcess As Boolean

        Try

            Select Case nIndex
                Case 1
                    'bProcess = pfunCreateLocalBackup()
                    bProcess = True
                Case 2
                    'bProcess = pfunBackupAccessDB()
                    bProcess = True
                Case 3
                    'bProcess = pfunRestoreLocalBackup()
                    bProcess = True
                Case 4
                    'bProcess = pfunClearLogLocal()
                    bProcess = True
                Case 5
                    'bProcess = pfunRelinkObjects()
                    bProcess = True
                Case 6
                    bProcess = pfunCreateWebBackup()
                    bProcess = True
                Case 7
                    bProcess = pfunRestoreWebBackup()
                    bProcess = True
                Case 8
                    bProcess = pfunClearLogWeb()
                    bProcess = True
            End Select

        Catch ex As Exception
            bProcess = False
            MsgBox(ex.Message & " in pfunBackupProcess()")
        End Try

        Return bProcess

    End Function

    Private Function pfunCreateLocalBackup() As Boolean

        Dim conLocalSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean
        Dim strFilePath As String

        Try
            strFilePath = "C:\Database\AWTestZone\SQL_Backup\" & m_strTargetDB & ".bak"

            If System.IO.File.Exists(strFilePath) Then
                System.IO.File.Delete(strFilePath)
            End If

            Call gsubSqlConnectLocal(conLocalSQL, m_strSourceDB)

            strQuery = ""
            strQuery = strQuery & "BACKUP Database " & m_strSourceDB & " "
            strQuery = strQuery & "To disk = '" & strFilePath & "' "
            strQuery = strQuery & "WITH FORMAT, "
            strQuery = strQuery & "    MEDIANAME = '" & m_strSourceDB & "', "
            strQuery = strQuery & "    NAME = 'Full Backup of " & m_strSourceDB & "' "

            bResult = gfunDataWriterSQL(strQuery, conLocalSQL)

            Call gsubSqlDisconnet(conLocalSQL)

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunCreateLocalBackup()")
        End Try

        Return bResult

    End Function

    Private Function pfunRestoreLocalBackup() As Boolean

        Dim conLocalSQL As New SqlConnection
        Dim strQuery As String
        Dim strFilePath As String
        Dim strDataFile As String
        Dim strLogFile As String
        Dim bResult As Boolean

        Try
            strDataFile = "C:\Program Files\Microsoft SQL Server\MSSQL14.AWTESTZONE17\MSSQL\DATA\" & m_strTargetDB & ".mdf"
            strLogFile = "C:\Program Files\Microsoft SQL Server\MSSQL14.AWTESTZONE17\MSSQL\DATA\" & m_strTargetDB & "_log.ldf"

            If InStr(m_strSourceDB, "_") = 0 Then
                strFilePath = "C:\Database\AWTestZone\SQL_Backup\" & m_strTargetDB & ".bak"
            Else
                strFilePath = "C:\ArtwoodsSQL_Backups\" & cmbDatabase2.Text
            End If

            Call gsubSqlConnectLocal(conLocalSQL, "master", True)

            strQuery = ""
            strQuery = strQuery & "RESTORE Database " & m_strTargetDB & " "
            strQuery = strQuery & "FROM DISK = '" & strFilePath & "' "
            strQuery = strQuery & "WITH FILE = 1, "
            strQuery = strQuery & "    MOVE '" & m_strSourceDB & "' TO '" & strDataFile & "', "
            strQuery = strQuery & "    MOVE '" & m_strSourceDB & "_log' TO '" & strLogFile & "', "
            strQuery = strQuery & "    NOUNLOAD, "
            strQuery = strQuery & "    STATS = 5 "

            bResult = gfunDataWriterSQL(strQuery, conLocalSQL)

            Call gsubSqlDisconnet(conLocalSQL)

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunRestoreLocalBackup()")
        End Try

        Return bResult

    End Function

    Private Function pfunClearLogLocal() As Boolean

        Dim conLocalSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean

        Try
            Call gsubSqlConnectLocal(conLocalSQL, m_strTargetDB, True)

            strQuery = "ALTER Database " & m_strTargetDB & " SET RECOVERY SIMPLE"

            If gfunDataWriterSQL(strQuery, conLocalSQL) = False Then
                Call gsubSqlDisconnet(conLocalSQL)
                Return False
            End If

            strQuery = "DBCC Shrinkfile(ArtwoodsSQL_log, 1)"

            If gfunDataWriterSQL(strQuery, conLocalSQL) = False Then
                Call gsubSqlDisconnet(conLocalSQL)
                Return False
            End If

            strQuery = "ALTER Database " & m_strTargetDB & " SET RECOVERY FULL"

            If gfunDataWriterSQL(strQuery, conLocalSQL) = False Then
                Call gsubSqlDisconnet(conLocalSQL)
                Return False
            End If

            Call gsubSqlDisconnet(conLocalSQL)

            bResult = True

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunClearLogLocal()")

        End Try

        Return bResult

    End Function

    Private Function pfunBackupAccessDB() As Boolean

        Dim bResult As Boolean
        Dim strSourceDB As String
        Dim strTargetDB As String

        Try
            If InStr(m_strSourceDB, "_") = 0 Then
                strSourceDB = "N:\Master.mdb"
            Else
                strSourceDB = "D:\Backup." & Format(m_dtTarget, "d") & "\Master.mdb"
            End If

            strTargetDB = "N:\AWTestZone\Master\Master" & m_strDate & ".mdb"

            If System.IO.File.Exists(strSourceDB) = True Then

                If System.IO.File.Exists(strTargetDB) Then
                    MsgBox("Delete exist file")
                    System.IO.File.Delete(strTargetDB)
                End If

                System.IO.File.Copy(strSourceDB, strTargetDB)

                bResult = True

            Else
                MsgBox("Database restore could not be completed because MDB file does not exist.")

            End If

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunBackupAccessDB()")

        End Try

        Return bResult

    End Function

    Private Function pfunRelinkObjects() As Boolean

        Dim dbs As Access.Dao.Database
        Dim dbsE As New Access.Dao.DBEngine()
        Dim bResult As Boolean
        Dim strTargetAccess As String

        Try
            strTargetAccess = "N:\AWTestZone\Master\Master" & m_strDate & ".mdb"

            'Connection to the database 
            dbs = dbsE.OpenDatabase(strTargetAccess)

            For Each td As Access.Dao.TableDef In dbs.TableDefs

                If Len(td.Connect) > 0 Then

                    If InStr(td.Connect, "DATABASE=ArtwoodsSQL") > 0 Then
                        td.Connect = "ODBC;DSN=AWTestZone;" &
                                     "UID=" & gInfo.strAWG_UID & ";PWD=" & gInfo.strAWG_PWD & ";" &
                                     "Trusted_Connection=No;DATABASE=" & m_strTargetDB
                        td.RefreshLink()
                    ElseIf InStr(td.Connect, "Master-Web") Then
                        td.Connect = ";DATABASE=N:\AWTestZone\Master\Master-Web.mdb"
                        td.RefreshLink()
                    End If

                End If

                Application.DoEvents()

            Next

            For Each qd As Access.Dao.QueryDef In dbs.QueryDefs

                If Len(qd.Connect) > 0 Then

                    If InStr(qd.Connect, "DATABASE=ArtwoodsSQL") > 0 Then
                        qd.Connect = "ODBC;DSN=AWTestZone;" &
                                     "UID=" & gInfo.strAWG_UID & ";PWD=" & gInfo.strAWG_PWD & ";" &
                                     "Trusted_Connection=No;DATABASE=" & m_strTargetDB
                    End If

                End If

                Application.DoEvents()

            Next

            dbs.Close()

            bResult = True

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunRelinkObjects()")

        End Try

        Return bResult

    End Function

    Private Function pfunCreateWebBackup() As Boolean

        Dim conWebSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean

        Dim strFilePath As String
        Dim strSourceDB As String

        Try
            strSourceDB = "artwoods"
            strFilePath = "C:\Database\AWTestZone" & m_strDate & ".bak"

            If System.IO.File.Exists(strFilePath) Then
                System.IO.File.Delete(strFilePath)
            End If

            Call gsubSqlConnectWeb(conWebSQL, strSourceDB)

            strQuery = ""
            strQuery = strQuery & "BACKUP Database " & strSourceDB & " "
            strQuery = strQuery & "To disk = '" & strFilePath & "' "
            strQuery = strQuery & "WITH FORMAT, "
            strQuery = strQuery & "    MEDIANAME = '" & strSourceDB & "', "
            strQuery = strQuery & "    NAME = 'Full Backup of " & strSourceDB & "' "

            bResult = gfunDataWriterSQL(strQuery, conWebSQL)

            Call gsubSqlDisconnet(conWebSQL)

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunCreateWebBackup()")
        End Try

        Return bResult

    End Function

    Private Function pfunRestoreWebBackup() As Boolean

        Dim conWebSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean

        Dim strFilePath As String
        Dim strTargetDB As String
        Dim strDataFile As String
        Dim strLogFile As String

        Try
            strTargetDB = "AWTestZone" & m_strDate
            strDataFile = "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\" & strTargetDB & ".mdf"
            strLogFile = "C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\DATA\" & strTargetDB & "_log.ldf"
            strFilePath = "C:\Database\" & strTargetDB & ".bak"

            Call gsubSqlConnectWeb(conWebSQL, "master")

            strQuery = ""
            strQuery = strQuery & "RESTORE Database " & strTargetDB & " "
            strQuery = strQuery & "FROM DISK = '" & strFilePath & "' "
            strQuery = strQuery & "WITH FILE = 1, "
            strQuery = strQuery & "    MOVE 'artwoods_Data' TO '" & strDataFile & "', "
            strQuery = strQuery & "    MOVE 'artwoods_Log' TO '" & strLogFile & "', "
            strQuery = strQuery & "    NOUNLOAD, "
            strQuery = strQuery & "    STATS = 5 "

            bResult = gfunDataWriterSQL(strQuery, conWebSQL)

            Call gsubSqlDisconnet(conWebSQL)

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunRestoreWebBackup()")
        End Try

        Return bResult

    End Function

    Private Function pfunClearLogWeb() As Boolean

        Dim conWebSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean
        Dim strTargetDB As String

        strTargetDB = "AWTestZone" & m_strDate

        Try
            Call gsubSqlConnectWeb(conWebSQL, strTargetDB)

            strQuery = "ALTER Database " & strTargetDB & " SET RECOVERY SIMPLE"

            If gfunDataWriterSQL(strQuery, conWebSQL) = False Then
                Call gsubSqlDisconnet(conWebSQL)
                Return False
            End If

            strQuery = "DBCC Shrinkfile(artwoods_log, 1)"

            If gfunDataWriterSQL(strQuery, conWebSQL) = False Then
                Call gsubSqlDisconnet(conWebSQL)
                Return False
            End If

            strQuery = "ALTER Database " & strTargetDB & " SET RECOVERY FULL"

            If gfunDataWriterSQL(strQuery, conWebSQL) = False Then
                Call gsubSqlDisconnet(conWebSQL)
                Return False
            End If

            Call gsubSqlDisconnet(conWebSQL)

            bResult = True

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunClearLogWeb()")

        End Try

        Return bResult

    End Function

    Private Sub psubSetWaitCursor(ByVal bEnable As Boolean)

        pbProgress.Maximum = MAX_PROCESS

        If bEnable = True Then
            pbProgress.Value = 0
        End If

        btnCreateBackup1.Enabled = Not bEnable
        btnCreateBackup2.Enabled = Not bEnable

        Call psubResetGridProcess()

        Me.UseWaitCursor = bEnable

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'Dim strTargetAccess As String
        'Dim strTargetDB As String

        'strTargetDB = "ArtwoodsSQL_5_18"
        'strTargetAccess = "N:\AWTestZone\Master\Master_5_18.mdb"

        'If pfunRelinkObjects(strTargetDB, strTargetAccess) = True Then
        '    MsgBox("Success")
        'Else
        '    MsgBox("Fail")
        'End If

        'dgvProgress.SelectedRows = 5

        dgvProgress.ClearSelection()
        dgvProgress.Rows(5).Selected = True


        pbProgress.Maximum = 8
        pbProgress.Minimum = 0
        pbProgress.Value = 5



    End Sub

End Class

