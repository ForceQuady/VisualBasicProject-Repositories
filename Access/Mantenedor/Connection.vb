Imports System.Data.OleDb

''' -----------------------------------------------------------------------------
''' Author   : Trx-Homie 
''' Project  : Visual Basic with Access
''' Class    : TestForm
''' Github   : HomieStart
''' License  : Creative Commons
''' -----------------------------------------------------------------------------
''' <summary>
''' This class only used for test Connections
''' </summary>
''' </version>0.1
''' </version>
''' <remarks>
''' If you need to use the Secuence Class, Using new Secuence(Object As OleDblConnection)
''' </remarks>
''' <history>
'''    [FQ-HomieStart] Created
''' </history>
''' 
''' -----------------------------------------------------------------------------

Public MustInherit Class Connection
    Private Const HOST As String = "127.0.0.1" 'IPHOST : is not necesary change'
    Private Const PORT As String = "3302" 'PORT : is not necesary change'
    Private Const DB_PATH As String = "C:\YOU_PATH\"
    Private Const DB_NAME As String = "YOU_DATABASE.mdb"
    Private Const PROVIDER As String = "Microsoft.Jet.OLEDB.4.0"
    Private Const DRIVER As String = "Provider=" & PROVIDER & ";Data Source=" & DB_PATH & DB_NAME
    Private flag As Boolean
    Private con As OleDbConnection
    Private cmd As OleDbCommand
    Private dapt As OleDbDataAdapter

    Public Function getConnection() As OleDbConnection
        'WARNING: Do erase After in future version ============='
        Return con
    End Function

    Sub New()
        flag = False
        Try
            con = New OleDbConnection(DRIVER)
            con.Open()
            System.Console.WriteLine("Connection Success")
        Catch ex As Exception
            System.Console.WriteLine("Connection Error:" & ex.ToString)
        End Try

    End Sub

    Private Sub Close()
        Try
            con.Close()
            System.Console.WriteLine("Connection Closed")
        Catch ex As Exception
            System.Console.WriteLine("Error Close Connection:" & ex.ToString)
        End Try
    End Sub

    Public MustOverride Function Execute_SendQuery() As Boolean
    Public MustOverride Function Execute_InsertQuery() As Boolean
    Public MustOverride Function Execute_DeleteQuery() As Boolean
End Class

Public NotInheritable Class Connect : Inherits Connection
    Sub New()
        MyBase.New()
    End Sub

    Public Overrides Function Execute_DeleteQuery() As Boolean
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Execute_InsertQuery() As Boolean
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Execute_SendQuery() As Boolean
        Throw New NotImplementedException()
    End Function
End Class
