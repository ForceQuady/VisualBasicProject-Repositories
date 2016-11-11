Imports TestConnection_VisualBasic



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
''' </history>
''' 
''' -----------------------------------------------------------------------------

Public Module mainModule
    Public Sub Main()
        WriteConsoleFor(10, "=", False)
        System.Console.Write(" -Begin- ")
        WriteConsoleFor(10, "=", False)
        'Dim a = New Connect()

        Dim Sentences = New Secuence((New Connect).getConnection())
        'QueryString = INSERT INTO Tabla1([nombre]) VALUES(?)
        'cmd.parameters.Add(New OleDbParameter('CType[NOMBRES_TRING, String])
        'try cmd.executeNonQuery() cmd.dispose() con.Close

        Sentences.Send_QueryInsert("INSERT INTO Tabla1(nombre) VALUES('" & CType("Hello World", String) & "')")


        WriteConsoleFor(10, "=", False)
        System.Console.Write(" -END- ")
        WriteConsoleFor(10, "=", False)
    End Sub


End Module

Public Class Test


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Main()'
    End Sub
End Class
