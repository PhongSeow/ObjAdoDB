Imports ObjAdoDBLib

Public Class ConsoleDemo
    Public Conn As New Connection
    Public ConnStr As String
    Public SQL As String
    Public RS As Recordset
    Public DBSrv As String = "localhost"
    Public DBUser As String = "sa"
    Public DBPwd As String = ""
    Public CurrDB As String = "master"
    Public Provider As Connection.ProviderEnum
    Public CurrConsoleKey As ConsoleKey
    Public InpStr As String
    Public AccessFilePath As String

    Public Sub Main()
        With Me.Conn
            .ConnectionTimeout = 5
            .ConnectionString = Me.ConnStr
        End With
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to Set Connection String")
            Console.WriteLine("Press B to Open Connection")
            Console.WriteLine("Press C to Show Connection Information")
            Console.WriteLine("Press D to Create Recordset with Execute")
            Console.WriteLine("Press E to Show Recordset Information")
            Console.WriteLine("Press F to Recordset.MoveNext")
            Console.WriteLine("Press G to Recordset.NextRecordset")
            Console.WriteLine("Press H to Test Command")
            Console.WriteLine("Press I to Test JSon")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("Set Connection String")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Press Q to Up")
                    Console.WriteLine("Press A to SQL Server")
                    Console.WriteLine("Press B to Access")
                    Do While True
                        Me.CurrConsoleKey = Console.ReadKey().Key
                        Select Case Me.CurrConsoleKey
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.WriteLine("Is Use Microsoft SQL Server OLEDB ? (Y/n)")
                                Me.InpStr = Console.ReadLine()
                                Select Case Me.InpStr
                                    Case "Y", "y", ""
                                        Me.Provider = Connection.ProviderEnum.MicrosoftSQLServer
                                        Console.WriteLine("Provider=MicrosoftSQLServer")
                                    Case Else
                                        Me.Provider = Connection.ProviderEnum.MicrosoftSQLServer2012NativeClient
                                        Console.WriteLine("Provider=MicrosoftSQLServer2012NativeClient")
                                End Select
                                Console.WriteLine("Input SQL Server:" & Me.DBSrv)
                                Me.DBSrv = Console.ReadLine()
                                If Me.DBSrv = "" Then Me.DBSrv = "localhost"
                                Console.WriteLine("SQL Server=" & Me.DBSrv)
                                Console.WriteLine("Input Default DB:" & Me.CurrDB)
                                Me.CurrDB = Console.ReadLine()
                                If Me.CurrDB = "" Then Me.CurrDB = "master"
                                Console.WriteLine("Default DB=" & Me.CurrDB)
                                Console.WriteLine("Is Trusted Connection ? (Y/n)")
                                Me.InpStr = Console.ReadLine()
                                Select Case Me.InpStr
                                    Case "Y", "y", ""
                                        Me.Conn.SetConnSQLServer(Me.DBSrv, Me.CurrDB, Me.Provider)
                                    Case Else
                                        Console.WriteLine("Input DB User:" & Me.CurrDB)
                                        Me.DBUser = Console.ReadLine()
                                        If Me.DBUser = "" Then Me.DBUser = "sa"
                                        Console.WriteLine("DB User=" & Me.DBUser)
                                        Console.WriteLine("Input DB Password:")
                                        Me.DBPwd = Console.ReadLine()
                                        Console.WriteLine("DB Password=" & Me.DBPwd)
                                        Me.Conn.SetConnSQLServer(Me.DBSrv, Me.DBUser, Me.DBPwd, Me.CurrDB, Me.Provider)
                                End Select
                                Console.WriteLine("ConnectionString=" & Me.Conn.ConnectionString)
                                Exit Do
                            Case ConsoleKey.B
                                Console.WriteLine("Input Access File Path:" & Me.AccessFilePath)
                                Me.AccessFilePath = Console.ReadLine()
                                Console.WriteLine("Access File Path=" & Me.AccessFilePath)
                                Me.Conn.SetConnAccess(Me.AccessFilePath)
                                Console.WriteLine("ConnectionString=" & Me.Conn.ConnectionString)
                                Exit Do
                        End Select
                    Loop
                Case ConsoleKey.B
                    Console.WriteLine("#################")
                    Console.WriteLine("Open Connection")
                    Console.WriteLine("#################")
                    With Me.Conn
                        .Open()
                        If .LastErr <> "" Then
                            Console.WriteLine("Connect Error:" & .LastErr)
                        Else
                            Console.WriteLine("Connect OK")
                        End If
                    End With
                Case ConsoleKey.C
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Connection Information")
                    Console.WriteLine("#################")
                    Console.WriteLine("ConnectionString=" & Me.Conn.ConnectionString)
                    Console.WriteLine("State=" & Me.Conn.State)
                Case ConsoleKey.D
                    Console.WriteLine("#################")
                    Console.WriteLine("Create Recordset with Execute")
                    Console.WriteLine("#################")
                    Console.WriteLine("Input SQL:")
                    Me.SQL = Console.ReadLine()
                    With Me.Conn
                        Me.RS = .Execute(SQL)
                        If .LastErr <> "" Then
                            Console.WriteLine("Execute Error:" & .LastErr)
                        Else
                            Console.WriteLine("Execute OK")
                        End If
                    End With
                Case ConsoleKey.E
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Recordset Information")
                    Console.WriteLine("#################")
                    With Me.RS
                        Console.WriteLine("Fields.Count=" & .Fields.Count)
                        If .Fields.Count > 0 Then
                            Dim i As Integer
                            For i = 0 To .Fields.Count - 1
                                Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                            Next
                        End If
                        Console.WriteLine("PageCount=" & .PageCount)
                        Console.WriteLine("EOF=" & .EOF)
                    End With
                Case ConsoleKey.F
                    Console.WriteLine("#################")
                    Console.WriteLine("Recordset.MoveNext")
                    Console.WriteLine("#################")
                    With Me.RS
                        .MoveNext()
                        If .LastErr <> "" Then
                            Console.WriteLine("MoveNext Error:" & .LastErr)
                        Else
                            Console.WriteLine("MoveNext OK")
                        End If
                    End With
                Case ConsoleKey.G
                    Console.WriteLine("#################")
                    Console.WriteLine("Recordset.NextRecordset")
                    Console.WriteLine("#################")
                    With Me.RS
                        Dim oRs As Recordset = .NextRecordset
                        If .LastErr <> "" Then
                            Console.WriteLine("Error:" & .LastErr)
                        Else
                            Console.WriteLine("OK")
                            With oRs
                                Console.WriteLine("Fields.Count=" & .Fields.Count)
                                If .Fields.Count > 0 Then
                                    Dim i As Integer
                                    For i = 0 To .Fields.Count - 1
                                        Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                    Next
                                End If
                                Console.WriteLine("PageCount=" & .PageCount)
                                Console.WriteLine("EOF=" & .EOF)
                            End With
                        End If
                    End With
                Case ConsoleKey.H
                    Console.WriteLine("#################")
                    Console.WriteLine("Test Command")
                    Console.WriteLine("#################")
                    Dim oCommand As New Command
                    With oCommand
                        Console.WriteLine("Set ActiveConnection")
                        .ActiveConnection = Me.Conn
                        Console.WriteLine("CommandText=""sp_helpdb""")
                        .CommandText = "sp_helpdb"
                        Console.WriteLine("CreateParameter @dbname=""master""")
                        .Parameters.Append(.CreateParameter("@dbname", Field.DataTypeEnum.adVarChar, Parameter.ParameterDirectionEnum.adParamInput, 128, "master"))
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                        Console.WriteLine("Execute")
                        Dim rsAny = .Execute()
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                            With rsAny
                                Console.WriteLine("Fields.Count=" & .Fields.Count)
                                If .Fields.Count > 0 Then
                                    Dim i As Integer
                                    For i = 0 To .Fields.Count - 1
                                        Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                                    Next
                                End If
                                Console.WriteLine("PageCount=" & .PageCount)
                                Console.WriteLine("EOF=" & .EOF)
                            End With
                        End If
                        .Parameters.Delete("@dbname")
                    End With
                Case ConsoleKey.I
                    Console.WriteLine("*******************")
                    Console.WriteLine("Test JSon")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Press Q to Up")
                    Console.WriteLine("Press A to Convert current row to JSON")
                    Do While True
                        Me.CurrConsoleKey = Console.ReadKey().Key
                        Select Case Me.CurrConsoleKey
                            Case ConsoleKey.Q
                                Exit Do
                            Case ConsoleKey.A
                                Console.WriteLine(Me.RS.Row2JSon)
                                Exit Do
                        End Select
                    Loop

            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
