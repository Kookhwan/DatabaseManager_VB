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
            MsgBox(ex.Message)
        End Try

        gfunDataTableSQL = table

    End Function


End Module
