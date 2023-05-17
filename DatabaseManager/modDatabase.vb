Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient

Module modDatabase

    Public Function gfunDataTableSQL(ByVal strSource As String, ByVal oDatabase As SqlConnection) As DataTable

        Dim command As New SqlCommand(strSource, oDatabase)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()

        Try
            adapter.Fill(table)
        Catch ex As Exception
            '// Error occurred while trying to adapter fill
            '// Send error message to console (change below line to customize error handling)
            MsgBox(ex.Message)
            Return Nothing
        End Try

        Return table

    End Function

    Public Function gfunDataReaderSQL(ByVal strSource As String, ByVal oDatabase As SqlConnection) As SqlDataReader

        Dim command As New SqlCommand(strSource, oDatabase)
        Dim reader As SqlDataReader = Nothing

        Try
            '// Pass the connection to a command object
            reader = command.ExecuteReader()
        Catch ex As Exception
            '// Error occurred while trying to execute reader
            '// Send error message to console (change below line to customize error handling)
            MsgBox(ex.Message + "in gfunDataReaderSQL()")

            '// failed SQL execution, lets return null
            Return Nothing
        End Try

        Return reader

    End Function

End Module
