Imports System
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Imports System.Web.SessionState
Imports System.Web

Public Class DBAccess
#Region "DECLARE"
    Private db_trans As SqlTransaction
    Private Intvalue As Integer
    Private BolValue As Boolean
    Private CmdValue As SqlCommand
    Private IntI As Integer
    Private SqlCnn As SqlConnection

#End Region

#Region "CONSTRUCTORS"
    Sub New()
        db_trans = Nothing
        Intvalue = Nothing
        BolValue = Nothing
        CmdValue = Nothing
        IntI = Nothing
        SqlCnn = Nothing
    End Sub
#End Region




#Region "PUBLIC METHOD"
    'For NORMAL
    Public Function Sqlparameter(ByVal SqlCnn As SqlConnection, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal isparams As Boolean) As SqlCommand

        Dim Cmd As SqlCommand
        Cmd = New SqlCommand(CmdText, SqlCnn)
        Cmd.CommandType = CommandType.StoredProcedure
        If isparams = True Then
            If SqlParmas IsNot Nothing Then
                Dim LenParams As Integer
                LenParams = SqlParmas.Length
                IntI = 0
                While IntI < LenParams
                    Cmd.Parameters.Add(SqlParmas(IntI))
                    IntI = IntI + 1
                End While
            Else : End If
        ElseIf isparams = False Then
        End If
        Return Cmd
    End Function

    Public Function SQlExecute(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal isparams As Boolean) As Integer
        Dim Cmd As SqlCommand = Nothing
        Using SqlCnn As New SqlConnection(CnnStr)

            Cmd = Sqlparameter(SqlCnn, CmdText, SqlParmas, isparams)
            If Cmd.Connection.State <> ConnectionState.Open Then
                Cmd.Connection.Open()
            End If
            Intvalue = Cmd.ExecuteNonQuery()
        End Using
        Return Intvalue
    End Function
    ''' <summary>
    ''' Thuc Thi cau lenh Sql co tham so Va tra ve 1 gia tri
    ''' </summary>
    Public Function SQlExecuteReturnOneValue(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal isparams As Boolean) As SqlParameterCollection
        Dim LenParams As Integer

        LenParams = SqlParmas.Length
        Using SqlCnn As New SqlConnection(CnnStr)
            Dim Cmd As SqlCommand
            Cmd = Sqlparameter(SqlCnn, CmdText, SqlParmas, isparams)
            Cmd.ExecuteNonQuery()
            Cmd.Connection.Close()
            Return Cmd.Parameters
        End Using
    End Function


#Region "RETURN TABLE"
    Private Function FillDataTable(ByVal command As SqlCommand) As DataTable
        Using dAdapter As New SqlDataAdapter(command)
            Dim dTable As New DataTable()
            If command.Connection.State <> ConnectionState.Open Then
                command.Connection.Open()
            End If
            dAdapter.Fill(dTable)
            command.Connection.Close()
            Return dTable
        End Using
    End Function

    Public Function Filltable(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal isparams As Boolean) As DataTable

        Dim command As SqlCommand = Nothing
        Dim tbl As New DataTable

        Using connection As New SqlConnection(CnnStr)
            command = Sqlparameter(connection, CmdText, SqlParmas, isparams)
            tbl = FillDataTable(command)
        End Using

        Return tbl
    End Function

    Public Function Filltable_time_out(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal isparams As Boolean) As DataTable

        Dim command As SqlCommand = Nothing
        Dim tbl As New DataTable

        Using connection As New SqlConnection(CnnStr)
            command = Sqlparameter(connection, CmdText, SqlParmas, isparams)
            command.CommandTimeout = 1800
            tbl = FillDataTable(command)
        End Using

        Return tbl
    End Function
#End Region
#End Region

#Region "TRANSACTION METHOD"



    Public Function SQLExecuteTransaction(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal IsParams As Boolean) As Integer
        Dim LenParams As Integer
        LenParams = SqlParmas.Length
        Dim Cmd As SqlCommand
        Dim SqlCnn As SqlConnection
        SqlCnn = New SqlConnection(CnnStr)
        SqlCnn.Open()
        db_trans = SqlCnn.BeginTransaction()

        Try
            Cmd = New SqlCommand(CmdText, SqlCnn)
            Cmd.CommandType = CommandType.StoredProcedure
            Cmd.Transaction = db_trans
            If IsParams = True Then
                If SqlParmas IsNot Nothing Then
                    LenParams = SqlParmas.Length
                    IntI = 0
                    While IntI < LenParams
                        Cmd.Parameters.Add(SqlParmas(IntI))
                        IntI = IntI + 1
                    End While
                Else : End If
            ElseIf IsParams = False Then
            End If
            Intvalue = Cmd.ExecuteNonQuery
            db_trans.Commit()

        Catch ex As Exception
            db_trans.Rollback()

        Finally
            SqlCnn.Close()
        End Try

        Return Intvalue

    End Function
    'Public Shared Function SQLExecuteTransactionReturnOneValue(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal IsParams As Boolean) As SqlParameterCollection
    '    'Using SqlCnn As New SqlConnection(CnnStr)

    '    '    Cmd.Connection.Close()
    '    '    Return Cmd.Parameters
    '    'End Using
    '    ' Return Cmd.Parameters
    'End Function

    Private Function FillDataSet(ByVal command As SqlCommand) As DataSet
        Using dAdapter As New SqlDataAdapter(command)
            Dim d_ds As New DataSet()
            If command.Connection.State <> ConnectionState.Open Then
                command.Connection.Open()
            End If
            dAdapter.Fill(d_ds)
            command.Connection.Close()
            Return d_ds
        End Using
    End Function

    Public Function Filldataset(ByVal CnnStr As String, ByVal CmdText As String, ByVal SqlParmas() As SqlParameter, ByVal isparams As Boolean) As DataSet

        Dim command As SqlCommand = Nothing
        Dim ds As New DataSet

        Using connection As New SqlConnection(CnnStr)
            command = Sqlparameter(connection, CmdText, SqlParmas, isparams)
            ds = Filldataset(command)
        End Using

        Return ds
    End Function
#End Region
#Region "check permission"
    Private Sub check_permission()
        'Dim _security_common As ad

        'If HttpContext.Current.Session(_security_common.security_session_admin_uid) IsNot Nothing Then

        'End If
    End Sub

#End Region
End Class
