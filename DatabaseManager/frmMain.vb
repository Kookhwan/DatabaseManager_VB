Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports System.Threading

Public Class frmMain

    Private m_dtProgress As New DataTable
    Private oThread As Thread

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

                cmbDatabase2.Text = "Select database"

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

        Call psubSetWaitCursor(True)

        btnCreateBackup1.Enabled = False
        btnCreateBackup2.Enabled = False

        'Call psubResetGridProcess()

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

        Try
            If pfunNewBackupDatabase(False) = True Then
                MsgBox("Test database created successfully.", "Result")
            Else
                MsgBox("Backup failed, please Try it again.", "Result")
            End If
        Catch ex As Exception
            MsgBox(ex.Message & " In btnCreateBackup1_Click()")
        End Try

    End Sub

    Private Sub psubWorkThread()

        Dim nIndex As Integer = 0

        While (gbStart)

            nIndex += 1

            If (Me.InvokeRequired) Then
                Me.Invoke(Sub()

                              Call psubSetGridProgress(nIndex, "Running")

                              If pfunBackupProcess(nIndex) = True Then
                                  Call psubSetGridProgress(nIndex, "Finished")
                              Else
                                  Call psubSetGridProgress(nIndex, "failed")
                                  gbStart = False
                              End If

                          End Sub)

            End If

            If nIndex >= 8 Then
                gbStart = False

                If (Me.InvokeRequired) Then
                    Me.Invoke(Sub()
                                  MsgBox("Done")
                                  Call psubSetWaitCursor(False)
                                  btnCreateBackup1.Enabled = True
                                  btnCreateBackup2.Enabled = True
                                  Call psubResetGridProcess()
                              End Sub)
                End If

                oThread.Abort()

            End If

            Thread.Sleep(1000)

        End While

    End Sub

    Private Sub psubResetGridProcess(Optional ByVal bRestore As Boolean = False)

        Dim nStep As Integer

        For nStep = 0 To 7

            If nStep = 0 Then
                m_dtProgress.Rows(nStep)("Status") = IIf(bRestore, "N/A", "Pending")
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

        Application.DoEvents()

    End Sub

    Private Function pfunBackupProcess(ByVal nIndex As Integer) As Boolean

        Dim strDate As String
        Dim strSourceDB As String
        Dim strTargetDB As String
        Dim strFilePath As String
        Dim strSourceAccess As String
        Dim strTargetAccess As String
        Dim bProcess As Boolean

        Try

            strDate = Format(Now(), "_M_d")
            strSourceDB = cmbDatabase1.SelectedItem
            strTargetDB = strSourceDB & strDate
            strFilePath = "C:\Database\AWTestZone\SQL_Backup\" & strTargetDB & ".bak"
            strSourceAccess = "N:\Master.mdb"
            strTargetAccess = "N:\AWTestZone\Master\Master" & strDate & ".mdb"

            Select Case nIndex
                Case 1
                    'bProcess = pfunCreateLocalBackup(strSourceDB, strFilePath)
                    bProcess = True
                Case 2
                    'bProcess = pfunBackupAccessDB(strSourceAccess, strTargetAccess)
                    bProcess = True
                Case 3
                    'bProcess = pfunRestoreBackupLocal(strSourceDB, strTargetDB, strFilePath)
                    bProcess = True
                Case 4
                    'bProcess = pfunClearLogLocal(strTargetDB)
                    bProcess = True
                Case 5
                    'bProcess = pfunRelinkObjects(strTargetDB, strTargetAccess)
                    bProcess = True
                Case 6
                    'bProcess = pfunCreateLocalBackup(strSourceDB, strFilePath)
                    bProcess = True
                Case 7
                    'bProcess = pfunRestoreBackupLocal(strSourceDB, strTargetDB, strFilePath)
                    bProcess = True
                Case 8
                    'bProcess = pfunClearLogLocal(strTargetDB)
                    bProcess = True
            End Select

        Catch ex As Exception
            bProcess = False
            MsgBox(ex.Message & " in pfunBackupProcess()")
        End Try

        Return bProcess

    End Function





    Private Function pfunNewBackupDatabase(ByVal bNewBackup As Boolean) As Boolean

        Dim strDate As String
        Dim strSourceDB As String
        Dim strTargetDB As String
        Dim strFilePath As String

        Dim strSourceAccess As String
        Dim strTargetAccess As String

        Dim bProcess As Boolean = True

        strDate = Format(Now(), "_M_d")
        strSourceDB = cmbDatabase1.SelectedItem
        strTargetDB = strSourceDB & strDate
        strFilePath = "C:\Database\AWTestZone\SQL_Backup\" & strTargetDB & ".bak"

        strSourceAccess = "N:\Master.mdb"
        strTargetAccess = "N:\AWTestZone\Master\Master" & strDate & ".mdb"

        '// Create a new backup file from Local Live Database
        If bNewBackup = True Then
            Call psubSetGridProgress(1, "Running")
            bProcess = pfunCreateLocalBackup(strSourceDB, strFilePath)
        End If

        '// Restore database from the backed up file
        If bProcess = True Then
            Call psubSetGridProgress(1, "Finished")
            Call psubSetGridProgress(2, "Running")
            bProcess = pfunRestoreBackupLocal(strSourceDB, strTargetDB, strFilePath)
        Else
            Call psubSetGridProgress(1, "Failed")
            Return False
        End If

        '// Clear log file
        If bProcess = True Then
            Call psubSetGridProgress(2, "Finished")
            Call psubSetGridProgress(3, "Running")
            bProcess = pfunClearLogLocal(strTargetDB)
        Else
            Call psubSetGridProgress(2, "Failed")
            Return False
        End If

        '// Backup Access DB
        If bProcess = True Then
            Call psubSetGridProgress(3, "Finished")
            Call psubSetGridProgress(4, "Running")
            bProcess = pfunBackupAccessDB(strSourceAccess, strTargetAccess)
        Else
            Call psubSetGridProgress(3, "Failed")
            Return False
        End If

        '// Must wait a moment to access it after backing up the Access DB.
        Threading.Thread.Sleep(5000)

        '// Re-Link objects into the MS Access
        If bProcess = True Then
            Call psubSetGridProgress(4, "Finished")
            Call psubSetGridProgress(5, "Running")
            bProcess = pfunRelinkObjects(strTargetDB, strTargetAccess)
        Else
            Call psubSetGridProgress(4, "Failed")
            Return False
        End If

        '// Create a new backup for Web
        If bProcess = True Then
            Call psubSetGridProgress(5, "Finished")
            Call psubSetGridProgress(6, "Running")
            'bProcess = pfunRelinkObjects(strTargetDB, strTargetAccess)
            bProcess = True
        Else
            Call psubSetGridProgress(5, "Failed")
            Return False
        End If

        '// Restore Database for Web
        If bProcess = True Then
            Call psubSetGridProgress(6, "Finished")
            Call psubSetGridProgress(7, "Running")
            'bProcess = pfunRelinkObjects(strTargetDB, strTargetAccess)
            bProcess = True
        Else
            Call psubSetGridProgress(6, "Failed")
            Return False
        End If

        '// Log Clear for Web
        If bProcess = True Then
            Call psubSetGridProgress(7, "Finished")
            Call psubSetGridProgress(8, "Running")
            'bProcess = pfunRelinkObjects(strTargetDB, strTargetAccess)
            bProcess = True
        Else
            Call psubSetGridProgress(7, "Failed")
            Return False
        End If

        If bProcess = True Then
            Call psubSetGridProgress(8, "Finished")
        Else
            Call psubSetGridProgress(8, "Failed")
        End If

        Return bProcess

    End Function

    Private Function pfunCreateLocalBackup(ByVal strSourceDB As String, ByVal strFilePath As String) As Boolean

        Dim conLocalSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean

        Try
            Call gsubSqlConnectLocal(conLocalSQL, strSourceDB)

            strQuery = ""
            strQuery = strQuery & "BACKUP Database " & strSourceDB & " "
            strQuery = strQuery & "To disk = '" & strFilePath & "' "
            strQuery = strQuery & "WITH FORMAT, "
            strQuery = strQuery & "    MEDIANAME = '" & strSourceDB & "', "
            strQuery = strQuery & "    NAME = 'Full Backup of " & strSourceDB & "' "

            bResult = gfunDataWriterSQL(strQuery, conLocalSQL)

            Call gsubSqlDisconnet(conLocalSQL)

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunCreateLocalBackup()")
        End Try

        Return bResult

    End Function

    Private Function pfunRestoreBackupLocal(ByVal strSourceDB As String, ByVal strTargetDB As String, ByVal strFilePath As String) As Boolean

        Dim conLocalSQL As New SqlConnection
        Dim strQuery As String
        Dim strDataFile As String
        Dim strLogFile As String
        Dim bResult As Boolean

        strDataFile = "C:\Program Files\Microsoft SQL Server\MSSQL14.AWTESTZONE17\MSSQL\DATA\" & strTargetDB & ".mdf"
        strLogFile = "C:\Program Files\Microsoft SQL Server\MSSQL14.AWTESTZONE17\MSSQL\DATA\" & strTargetDB & "_log.ldf"

        Try
            Call gsubSqlConnectLocal(conLocalSQL, "master", True)

            strQuery = ""
            strQuery = strQuery & "RESTORE Database " & strTargetDB & " "
            strQuery = strQuery & "FROM DISK = '" & strFilePath & "' "
            strQuery = strQuery & "WITH FILE = 1, "
            strQuery = strQuery & "    MOVE '" & strSourceDB & "' TO '" & strDataFile & "', "
            strQuery = strQuery & "    MOVE '" & strSourceDB & "_log' TO '" & strLogFile & "', "
            strQuery = strQuery & "    NOUNLOAD, "
            strQuery = strQuery & "    STATS = 5 "

            bResult = gfunDataWriterSQL(strQuery, conLocalSQL)

            Call gsubSqlDisconnet(conLocalSQL)

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunRestoreBackupLocal()")
        End Try

        Return bResult

    End Function

    Private Function pfunClearLogLocal(ByVal strTargetDB As String) As Boolean

        Dim conLocalSQL As New SqlConnection
        Dim strQuery As String
        Dim bResult As Boolean

        Try
            Call gsubSqlConnectLocal(conLocalSQL, strTargetDB, True)

            strQuery = "ALTER Database " & strTargetDB & " SET RECOVERY SIMPLE"

            If gfunDataWriterSQL(strQuery, conLocalSQL) = False Then
                Call gsubSqlDisconnet(conLocalSQL)
                Return False
            End If

            strQuery = "DBCC Shrinkfile(ArtwoodsSQL_log, 1)"

            If gfunDataWriterSQL(strQuery, conLocalSQL) = False Then
                Call gsubSqlDisconnet(conLocalSQL)
                Return False
            End If

            strQuery = "ALTER Database " & strTargetDB & " SET RECOVERY FULL"

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

    Private Function pfunBackupAccessDB(ByVal strSourceDB As String, ByVal strTargetDB As String) As Boolean

        Dim bResult As Boolean

        Try
            If System.IO.File.Exists(strSourceDB) Then

                If System.IO.File.Exists(strTargetDB) Then
                    MsgBox("Delete exist file")
                    System.IO.File.Delete(strTargetDB)
                End If

                System.IO.File.Copy(strSourceDB, strTargetDB)

                bResult = True

            End If

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunBackupAccessDB()")

        End Try

        Return bResult

    End Function

    Private Function pfunRelinkObjects(ByVal strTargetDB As String, ByVal strTargetAccess As String) As Boolean

        Dim dbs As Access.Dao.Database
        Dim dbsE As New Access.Dao.DBEngine()
        Dim bResult As Boolean

        'Connection to the database 
        dbs = dbsE.OpenDatabase(strTargetAccess)

        Try

            For Each td As Access.Dao.TableDef In dbs.TableDefs

                If Len(td.Connect) > 0 Then

                    If InStr(td.Connect, "DATABASE=ArtwoodsSQL") > 0 Then
                        td.Connect = "ODBC;DSN=AWTestZone;" &
                                     "UID=" & gInfo.strAWG_UID & ";PWD=" & gInfo.strAWG_PWD & ";" &
                                     "Trusted_Connection=No;DATABASE=" & strTargetDB
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
                                     "Trusted_Connection=No;DATABASE=" & strTargetDB
                    End If

                End If

                Application.DoEvents()

            Next

            bResult = True

        Catch ex As Exception
            bResult = False
            MsgBox(ex.Message & " in pfunRelinkObjects()")

        End Try

        dbs.Close()

        Return bResult

    End Function

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


    End Sub

    Private Sub psubSetWaitCursor(ByVal bEnable As Boolean)

        'Dim ctrl As Control = Me.GetNextControl(Me, True)

        'Do Until ctrl Is Nothing
        '    ctrl.UseWaitCursor = bEnable
        '    ctrl = Me.GetNextControl(ctrl, True)
        'Loop

        Me.UseWaitCursor = bEnable

    End Sub

End Class

