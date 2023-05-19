﻿Option Explicit On

Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Xml

Module modDatabase

    Private gInfo As ServerInfo

    Public Structure ServerInfo

        Public strAWG_IP As String
        Public strAWG_UID As String
        Public strAWG_PWD As String
        Public strWEB_IP As String
        Public strWEB_UID As String
        Public strWEB_PWD As String

    End Structure

    Public Sub gsubLoadXML_Info()

        Dim xmlDoc As New XmlDocument()

        xmlDoc.Load("C:\Files\ServerInfo.xml")

        gInfo.strAWG_IP = xmlDoc.SelectSingleNode("//System/Environment/AWG_IP").InnerText
        gInfo.strAWG_UID = xmlDoc.SelectSingleNode("//System/Environment/AWG_UID").InnerText
        gInfo.strAWG_PWD = xmlDoc.SelectSingleNode("//System/Environment/AWG_PWD").InnerText
        gInfo.strWEB_IP = xmlDoc.SelectSingleNode("//System/Environment/WEB_IP").InnerText
        gInfo.strWEB_UID = xmlDoc.SelectSingleNode("//System/Environment/WEB_UID").InnerText
        gInfo.strWEB_PWD = xmlDoc.SelectSingleNode("//System/Environment/WEB_PWD").InnerText

    End Sub

    Public Function gfunDataTableSQL(ByVal strSource As String, ByVal oDatabase As SqlConnection) As DataTable

        Dim command As New SqlCommand(strSource, oDatabase)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()

        Try
            adapter.Fill(table)
        Catch ex As Exception
            '// Error occurred while trying to adapter fill
            '// Send error message to console (change below line to customize error handling)
            MsgBox(ex.Message & " in gfunDataReaderSQL()")
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
            MsgBox(ex.Message & " in gfunDataReaderSQL()")

            '// failed SQL execution, lets return null
            Return Nothing
        End Try

        Return reader

    End Function

    Public Sub gsubSqlConnectLocal(ByRef dbConnect As SqlConnection, ByVal strDatabase As String, Optional ByVal bTestZone As Boolean = False)

        Dim strConSQL As String

        If bTestZone = True Then
            strConSQL = "Data Source=DBSERVER\AWTESTZONE17;Initial Catalog=" & strDatabase & ";Persist Security Info=True;User ID=sa;Password=&&AW1975&&"
        Else
            strConSQL = "Data Source=DBSERVER;Initial Catalog=" & strDatabase & ";Persist Security Info=True;User ID=sa;Password=&&AW1975&&"
        End If

        dbConnect = New SqlConnection

        dbConnect.ConnectionString = strConSQL

        If dbConnect.State <> ConnectionState.Open Then
            Try
                dbConnect.Open()
            Catch ex As Exception
                MsgBox(ex.Message & " in gsubSqlConnectLocal()")
            End Try
        End If

    End Sub

    Public Sub gsubSqlDisconnet(ByRef dbConnect As SqlConnection)

        dbConnect.Close()

    End Sub

    Public Function gfunDataWriterSQL(ByVal strSource As String, ByVal oDatabase As SqlConnection) As Boolean

        Dim command As New SqlCommand(strSource, oDatabase)

        Try
            command.CommandTimeout = 900
            command.ExecuteNonQuery()
            gfunDataWriterSQL = True
        Catch ex As Exception
            MsgBox(ex.Message + "in gfunDataWriterSQL()")
        End Try

        Return gfunDataWriterSQL

    End Function

End Module
