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
''' <remarks>
''' If you need to use the Secuence Class, Using new Secuence(Object As OleDblConnection)
''' </remarks>
''' <history>
'''    [FQ-HomieStart] Created
'''    [GitHub] HomieStart
''' </history>
''' 
''' -----------------------------------------------------------------------------

Public Interface ISecuence(Of OleDbConnection)
    ReadOnly Property isConnected As Boolean
    Sub Send_QuerySelect(ByVal query As String)
    Sub Send_QueryDelete(ByVal query As String)
    Sub Send_QueryInsert(ByVal query As String)
    Sub Send_QueryUpdate(ByVal query As String)
End Interface

Public Class Secuence : Implements ISecuence(Of OleDbConnection)

    Private con As OleDbConnection
    Public ReadOnly Property isConnected As Boolean Implements ISecuence(Of OleDbConnection).isConnected
        Get
            Try
                If (con.State <> ConnectionState.Closed) Then
                    System.Console.WriteLine("Is Connected !")
                    Return True
                End If
            Catch ex As Exception
                System.Console.WriteLine("Is not Connected !!")
                System.Console.WriteLine(ex)
            End Try
            Return False
        End Get
    End Property

    Sub New(ByVal connection As OleDbConnection)
        con = connection
    End Sub


    Public Sub Send_QueryPersonalized(ByVal query As String)
        If (isConnected) Then
            Dim table As DataTable
            Dim dapt As OleDbDataAdapter = New OleDbDataAdapter(query, con)
            Dim autoID As Integer = 0
            dapt.Fill(table)
            If (table.Rows.Count > 0) Then
                autoID = table.Rows(0).Item(0) + 1
            Else
                autoID += 1
            End If
            con.Close()
            System.Console.WriteLine("Connection query Success !")
        End If
    End Sub


    Public Sub Send_QueryDelete(query As String) Implements ISecuence(Of OleDbConnection).Send_QueryDelete
        Throw New NotImplementedException()
    End Sub

    Public Sub Send_QueryInsert(query As String) Implements ISecuence(Of OleDbConnection).Send_QueryInsert
        If (isConnected) Then
            Dim cmd = New OleDbCommand(query, con)
            Try
                cmd.ExecuteNonQuery()
                System.Console.WriteLine("Connection query Success !")
            Catch ex As Exception
                System.Console.WriteLine("Error Connection query ! " & ex.ToString)
            End Try
            cmd.Dispose()
            con.Close()
        End If
    End Sub

    Public Sub Send_QuerySelect(query As String) Implements ISecuence(Of OleDbConnection).Send_QuerySelect
        Throw New NotImplementedException()
    End Sub

    Public Sub Send_QueryUpdate(query As String) Implements ISecuence(Of OleDbConnection).Send_QueryUpdate
        Throw New NotImplementedException()
    End Sub
End Class