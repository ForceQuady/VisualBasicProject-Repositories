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
Public Module Auxiliary
    ''' <summary>
    ''' Aux Method for Write in Console
    ''' </summary>
    ''' <param name="HowMuch">How lenght of the Context to Write</param>
    ''' <param name="Context">Context to show in console</param>
    ''' <param name="Line">if Jump the line</param>
    ''' <example>
    '''     <code>
    '''         WriteConsoleFor(19, "=", false)  
    '''         WriteConsoleFor(3, "Hello", true)
    '''         WriteConsoleFor(10, "Hello", true)  
    '''     </code>
    '''     Out Put in Console
    '''     <code>
    '''         ==========
    '''         Hello
    '''         Hello
    '''         Hello
    '''         ==========
    '''     </code>
    ''' </example>
    Public Sub WriteConsoleFor(ByVal HowMuch As Integer, ByVal Context As String, Optional ByVal Line As Boolean = True)
        Dim iter As String = ""
        Dim i As Integer = 0
        Do
            iter += Context
            i += 1
        Loop While i < HowMuch
        If (Line) Then
            System.Console.WriteLine(iter)
        Else
            System.Console.Write(iter)
        End If
    End Sub
End Module
