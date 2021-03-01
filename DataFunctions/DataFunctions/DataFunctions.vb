
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.ComponentModel
Imports System.Data.Odbc
Imports System.Reflection
Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
Imports GlobalFunction1.GlobalFunction1
Imports RegisterDllClass
Imports GlobalControl.Variables
Imports UserProgressBar
Imports System.IO
Public Class DataFunctions
#Region "DimVariables"
    Dim GF1 As New GlobalFunction1.GlobalFunction1
    Dim dllInteger As Integer = 280196 + 130858
    Public Sub New()
        Try
            If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then
                Dim rc As New RegisterDllClass.ValidateClass
                Dim AuFlag As Integer = rc.AdminVault
                If AuFlag <> dllInteger Then
                    QuitMessage(Me.ToString & " " & rc.ReverseString(rc.aumess, 1), "new")
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunctions.New")
        End Try
    End Sub
    ''' <summary>
    ''' To get the data of .csv file to a datatable using comma split
    ''' </summary>
    ''' <param name="csvpath">In this the full path with name of the .csv file is given.</param>
    ''' <returns></returns>
    Public Function GetDatafromCSV1(ByVal csvpath As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SR As StreamReader = New StreamReader(csvpath)

        Dim line As String = SR.ReadLine()
        Dim strArray As String() = line.Split(","c)
        Dim dt As New DataTable()
        Dim row As DataRow
        For Each s As String In strArray
            dt.Columns.Add(New DataColumn(s))
        Next
        Do
            line = SR.ReadLine
            If Not line = String.Empty Then
                row = dt.NewRow()
                row.ItemArray = line.Split(","c)
                dt.Rows.Add(row)
            Else
                Exit Do
            End If
        Loop
        Return dt
    End Function


    ''' <summary>
    ''' Error Message box before quitting the application
    ''' </summary>
    ''' <param name="ex">Error on exception </param>
    ''' <param name="err"> error object </param>
    ''' <remarks></remarks>
    Public Function QuitError(ByVal ex As Exception, ByVal err As ErrObject, ByVal ErrorString As String) As String
        If LCase("WebAzure,WebGodaddy,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            GlobalControl.Variables.ErrorString = "ERROR_MESSAGE = " & ex.Message & " " & vbCrLf & vbCrLf & "<BR/>STACK_TRACE  =" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "<BR />Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString
            MsgBox("ERROR_MESSAGE ( " & ex.Message & " )" & vbCrLf & vbCrLf & "STACK_TRACE  (" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString)
            'MsgBox(Application.ProductName)
            System.Windows.Forms.Application.ExitThread()
            System.Windows.Forms.Application.Exit()
            Process.GetCurrentProcess.Kill()
        Else
            GlobalControl.Variables.ErrorString = "ERROR_MESSAGE ( " & ex.Message & " )" & vbCrLf & vbCrLf & "STACK_TRACE  (" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString
            'GlobalControl.Variables.ErrorString = "ERROR_MESSAGE = " & ex.Message & " " & vbCrLf & vbCrLf & "<BR/>STACK_TRACE  =" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "<BR />Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString
        End If
        Return GlobalControl.Variables.ErrorString
    End Function
    ''' <summary>
    ''' Message box before quitting the application
    ''' </summary>
    ''' <param name="MessageString"> message as string</param>
    '''<param name="QuitProcedure" >Function or sunroutine name from exception thrown</param>
    ''' <remarks></remarks>
    Public Sub QuitMessage(ByVal MessageString As String, ByVal QuitProcedure As String)
        If LCase("WebAzure,WebGodaddy,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            If MessageString.Length > 0 Then
                MsgBox(MessageString & vbCrLf & vbCrLf & QuitProcedure)
            End If
            System.Windows.Forms.Application.ExitThread()
            System.Windows.Forms.Application.Exit()
            Process.GetCurrentProcess.Kill()
        End If
    End Sub

#End Region
    ''' <summary>
    ''' Add a control element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">SqlParameter element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new control </returns>
    ''' <remarks></remarks>


    Public Function ArrayAppendSqlParameter(ByRef ArrayName() As SqlParameter, ByVal LastValue As SqlParameter, Optional ByVal IgnoreIfExists As Boolean = False) As SqlParameter()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = GF1.ArrayFind(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ArrayAppendSqlParameter(ByRef ArrayName() As System.Data.SqlClient.SqlParameter, ByVal LastValue As Control, Optional ByVal IgnoreIfExists As Boolean = False) As System.Data.SqlClient.SqlParameter()")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    ''' Break String of serverdatabase string into a list of string where index item (0) is Sql Server Name and index item (1) is database name 
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of database with server</param>
    ''' <returns>Return a list of server name and data base name</returns>
    ''' <remarks></remarks>
    Public Function BreakServerDataBase(ByVal ServerDataBase As String) As List(Of String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Alist As New List(Of String)
        Try
            Dim AServer() As String = Split(ServerDataBase, ".")
            Dim MSqlServer As String = ""
            Dim MdataBase As String = ""
            If AServer.Length > 1 Then
                MdataBase = AServer(AServer.Length - 1)
                Array.Resize(AServer, AServer.Length - 1)
                MSqlServer = Join(AServer, ".")
            End If
            Alist.Add(MSqlServer)
            Alist.Add(MdataBase)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.BreakServerDataBase(ByVal ServerDataBase As String) As List(Of String)")
        End Try

        Return Alist
    End Function

    ''' <summary>
    ''' To open new SQL connection
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database eg. server1.database1</param>
    ''' <param name="MaxPoolSize" >Max pool size connections at a time default is 100</param>
    ''' <param name="ConnectionTimeOut" >Connection time out in seconds defau</param>
    ''' <returns>An new sql connection</returns>
    ''' <remarks></remarks>
    Public Function OpenSqlConnection(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcons As New SqlConnection
        Try
            Dim AList As List(Of String) = BreakServerDataBase(ConvertFromSrv0Mdf0(ServerDataBase))
            Dim MSqlServer As String = AList(0)
            Dim MdataBase As String = AList(1)

            Dim mUserId As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserName").ToString.Trim
            Dim mUserPwd As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserPassword").ToString.Trim
            If IsLocalServer(MSqlServer) = False Then
                Dim mSaralType As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SaralType").ToString.Trim
                Select Case LCase(mSaralType)
                    Case "lan"
                        mUserId = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserName").ToString.Trim
                        mUserPwd = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserPassword").ToString.Trim
                    Case "weblocal"
                        mUserId = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserName").ToString.Trim
                        mUserPwd = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserPassword").ToString.Trim
                    Case "webgodaddy", "webazure"
                        mUserId = GlobalControl.Variables.WebHostingUserName.ToString.Trim
                        mUserPwd = GlobalControl.Variables.WebHostingUserPassword.ToString.Trim
                    Case "cloud"
                        mUserId = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserName").ToString.Trim
                        mUserPwd = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserPassword").ToString.Trim
                End Select
            End If
            Dim Retryduration As TimeSpan = TimeSpan.FromSeconds(ConnectionTimeOut)
            Dim StartTime As DateTime = Now()
            Dim MSqlServer0 As String = RemoveSquareBrackets(MSqlServer)
            Dim MdataBase0 As String = RemoveSquareBrackets(MdataBase)
            Dim sconn As String = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";Integrated Security=True;Trusted_Connection=Yes" & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
GoToTry:

            Try
                If mUserId.Length > 0 Then
                    sconn = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";User Id=" & mUserId & ";Password=" & mUserPwd & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
                End If
                GlobalControl.Variables.ErrorString = sconn
                lcons = New SqlConnection(sconn)
                lcons.Open()
            Catch ex As Exception
                If lcons Is Nothing Then
                    Dim Mdur As TimeSpan = Now() - StartTime
                    If Mdur < Retryduration Then
                        GoTo GoToTry
                    End If
                End If
                QuitError(ex, Err, "Connection not established in 30 seconds (" & sconn & "  In DataFunction.OpenSqlConnection(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection   ")
            End Try
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.OpenSqlConnection(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection")
        End Try
        Return lcons
    End Function
    ''' <summary>
    ''' To open new Local SQL connection
    ''' </summary>
    ''' <param name="DataBaseName" >Database name to be connected</param>
    ''' <param name="MaxPoolSize" >Max pool size connections at a time default is 100</param>
    ''' <param name="ConnectionTimeOut" >Connection time out in seconds defau</param>
    ''' <returns>An new sql connection</returns>
    ''' <remarks></remarks>
    Public Function OpenSqlConnectionLocal(ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcons As New SqlConnection
        Try
            Dim MSqlServer As String = ConvertFromSrv0(GlobalControl.Variables.AllServers("0_srv_0"))
            If MSqlServer.Trim.Length = 0 Then
                GF1.QuitMessage(" Local server not defined in globalcontrol in AllServers (0_srv_0 key))", "In DataFuncion.OpenSqlConnectionLocal(ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection ")
            End If
            Dim MdataBase As String = ConvertFromMdf0(DataBaseName)
            Dim mUserId As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserName").ToString.Trim
            Dim mUserPwd As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserPassword").ToString.Trim
            Dim Retryduration As TimeSpan = TimeSpan.FromSeconds(30)
            Dim StartTime As DateTime = Now()
            Dim MSqlServer0 As String = RemoveSquareBrackets(MSqlServer)
            Dim MdataBase0 As String = RemoveSquareBrackets(MdataBase)

            Dim sconn As String = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";Integrated Security=True;Trusted_Connection=Yes" & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
GoToTry:
            Try
                If mUserId.Length > 0 Then
                    sconn = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";User Id=" & mUserId & ";Password=" & mUserPwd & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
                End If
                GlobalControl.Variables.ErrorString = sconn
                lcons = New SqlConnection(sconn)
                lcons.Open()
            Catch ex As Exception
                If lcons Is Nothing Then
                    Dim Mdur As TimeSpan = Now() - StartTime
                    If Mdur < Retryduration Then
                        GoTo GoToTry
                    End If
                End If
                QuitError(ex, Err, "Connection not established in 30 seconds (" & sconn & "  In DataFunction.OpenSqlConnection(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection   ")
            End Try
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.OpenSqlConnection(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection")
        End Try
        Return lcons
    End Function
    ''' <summary>
    ''' To open new Remote SQL connection
    ''' </summary>
    ''' <param name="DataBaseName" >Database name to be connected</param>
    ''' <param name="MaxPoolSize" >Max pool size connections at a time default is 100</param>
    ''' <param name="ConnectionTimeOut" >Connection time out in seconds defau</param>
    ''' <returns>An new sql connection</returns>
    ''' <remarks></remarks>
    Public Function OpenSqlConnectionRemote(ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MSqlServer As String = ConvertFromSrv0(GlobalControl.Variables.AllServers("1_srv_1").ToString.Trim)
        If MSqlServer.Trim.Length = 0 Then
            GF1.QuitMessage(" Remote server not defined in globalcontrol in AllServers (1_srv_1 key))", "In DataFunction.OpenSqlConnectionRemote(ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection ")
        End If
        Dim MdataBase As String = ConvertFromMdf0(DataBaseName)
        Dim mUserId As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName").ToString.Trim
        Dim mUserPwd As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerPassword").ToString.Trim
        Dim lcons As New SqlConnection
        Dim Retryduration As TimeSpan = TimeSpan.FromSeconds(30)
        Dim StartTime As DateTime = Now()
        Dim MSqlServer0 As String = RemoveSquareBrackets(MSqlServer)
        Dim MdataBase0 As String = RemoveSquareBrackets(MdataBase)
        Dim sconn As String = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";Integrated Security=True;Trusted_Connection=Yes" & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
GoToTry:
        Try
            If mUserId.Length > 0 Then
                sconn = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";User Id=" & mUserId & ";Password=" & mUserPwd & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
            End If
            GlobalControl.Variables.ErrorString = sconn
            lcons = New SqlConnection(sconn)
            lcons.Open()

        Catch ex As Exception
            If lcons Is Nothing Then
                Dim Mdur As TimeSpan = Now() - StartTime
                If Mdur < Retryduration Then
                    GoTo GoToTry
                End If
            End If
            QuitError(ex, Err, "Connection not established in 30 seconds (" & sconn & "  In DataFunction.OpenSqlConnection(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection   ")
        End Try
        Return lcons
    End Function
    ''' <summary>
    ''' Remove square brackets from serever or database name or both.
    ''' </summary>
    ''' <param name="ServerOrDatabaseName">Server Or DatabaseName</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RemoveSquareBrackets(ByVal ServerOrDatabaseName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Microsoft.VisualBasic.Left(ServerOrDatabaseName, 1) <> "." Then
            Dim aServerDatabase() As String = ServerOrDatabaseName.Split(".")
            Select Case True
                Case aServerDatabase.Length = 1 Or IsNumeric(aServerDatabase(aServerDatabase.Length - 1)) = True
                    If Left(ServerOrDatabaseName, 1) = "[" Then
                        ServerOrDatabaseName = Mid(ServerOrDatabaseName, 2, ServerOrDatabaseName.Length - 2)
                    End If
                Case Else
                    Dim astring() As String = ServerOrDatabaseName.Split(".")
                    Dim mdatabase As String = astring(astring.Length - 1)
                    Array.Resize(astring, astring.Length - 1)
                    Dim mserver As String = Join(astring, ".")
                    If Left(mserver, 1) = "[" Then
                        mserver = Mid(mserver, 2, mserver.Length - 2)
                    End If
                    If Left(mdatabase, 1) = "[" Then
                        mdatabase = Mid(mdatabase, 2, mdatabase.Length - 2)
                    End If
                    ServerOrDatabaseName = mserver & "." & mdatabase
            End Select
        End If
        Return ServerOrDatabaseName
    End Function
    ''' <summary>
    ''' Add square brackets to serever or database name or both.
    ''' </summary>
    ''' <param name="ServerOrDatabaseName">Server Or DatabaseName</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddSquareBrackets(ByVal ServerOrDatabaseName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Microsoft.VisualBasic.Left(ServerOrDatabaseName, 1) <> "." Then
            Dim aServerDatabase() As String = ServerOrDatabaseName.Split(".")
            Select Case True
                Case aServerDatabase.Length = 1 Or IsNumeric(aServerDatabase(aServerDatabase.Length - 1)) = True
                    If Left(ServerOrDatabaseName, 1) <> "[" Then
                        ServerOrDatabaseName = "[" & ServerOrDatabaseName & "]"
                    End If
                Case Else
                    Dim astring() As String = ServerOrDatabaseName.Split(".")
                    Dim mdatabase As String = astring(astring.Length - 1)
                    Array.Resize(astring, astring.Length - 1)
                    Dim mserver As String = Join(astring, ".")
                    If Left(mserver, 1) <> "[" Then
                        mserver = "[" & mserver & "]"
                    End If
                    If Left(mdatabase, 1) <> "[" Then
                        mdatabase = "[" & mdatabase & "]"
                    End If
                    ServerOrDatabaseName = mserver & "." & mdatabase
            End Select
        End If
        Return ServerOrDatabaseName
    End Function



    ''' <summary>
    ''' Check server name is local or not
    ''' </summary>
    ''' <param name="ServerName">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsLocalServer(ByVal ServerName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mflag As Boolean = False
        Select Case LCase(GlobalControl.Variables.SaralType)
            Case "webazure", "webgodaddy", "cloud", "webocal"
            Case Else
                Dim msrv0 As String = ConvertFromSrv0(GlobalControl.Variables.AllServers("0_srv_0"))
                ServerName = ConvertFromSrv0(ServerName)
                mflag = IIf(msrv0 = ServerName, True, False)
        End Select
        Return mflag
    End Function
    ''' <summary>
    ''' To open new SQL connection
    ''' </summary>
    ''' <param name="ServerName">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="DataBaseName">Database name</param>
    ''' <param name="MaxPoolSize" >Max pool size connections at a time default is 100</param>
    ''' <param name="ConnectionTimeOut" >Connection time out in seconds defau</param>
    ''' <returns>An Sql connection</returns>
    ''' <remarks></remarks>
    Public Function OpenSqlConnection(ByVal ServerName As String, ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MSqlServer As String = ConvertFromSrv0(ServerName)
        Dim MdataBase As String = ConvertFromMdf0(DataBaseName)
        Dim mUserId As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserName").ToString.Trim
        Dim mUserPwd As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserPassword").ToString.Trim
        If IsLocalServer(MSqlServer) = False Then
            Dim mSaralType As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SaralType").ToString.Trim
            Select Case LCase(mSaralType)
                Case "lan"
                    mUserId = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserName").ToString.Trim
                    mUserPwd = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserPassword").ToString.Trim
                Case "weblocal"
                    mUserId = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserName").ToString.Trim
                    mUserPwd = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserPassword").ToString.Trim
                Case "webgodaddy", "webazure"
                    mUserId = GlobalControl.Variables.WebHostingUserName.ToString.Trim
                    mUserPwd = GlobalControl.Variables.WebHostingUserPassword.ToString.Trim
                Case "cloud"
                    mUserId = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserName").ToString.Trim
                    mUserPwd = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserPassword").ToString.Trim
            End Select
        End If
        Dim lcons As New SqlConnection
        Dim Retryduration As TimeSpan = TimeSpan.FromSeconds(30)
        Dim StartTime As DateTime = Now()
        Dim MSqlServer0 As String = RemoveSquareBrackets(MSqlServer)
        Dim MdataBase0 As String = RemoveSquareBrackets(MdataBase)
        Dim sconn As String = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";Integrated Security=True;Trusted_Connection=Yes" & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
        If mUserId.Length > 0 Then
            sconn = "Data Source=" & MSqlServer0 & ";Initial Catalog=" & MdataBase0 & ";User Id=" & mUserId & ";Password=" & mUserPwd & ";Max Pool Size=" & MaxPoolSize & ";Connection TimeOut=" & ConnectionTimeOut
        End If
GoToTry:
        Try
            GlobalControl.Variables.ErrorString = sconn
            lcons = New SqlConnection(sconn)
            lcons.Open()
        Catch ex As Exception
            If lcons Is Nothing Then
                Dim Mdur As TimeSpan = Now() - StartTime
                If Mdur < Retryduration Then
                    GoTo GoToTry
                End If
            End If
            QuitError(ex, Err, " Connection not established in 30 seconds (" & sconn & ")  DataFunction.OpenSqlConnection(ByVal ServerName As String, ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection")
            QuitMessage(" Connection not established in 30 seconds (" & sconn & ")", "DataFunction.OpenSqlConnection(ByVal ServerName As String, ByVal DataBaseName As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlConnection")
        End Try
        Return lcons
    End Function

    ''' <summary>
    ''' Attach a database to server
    ''' </summary>
    ''' <param name="ServerName">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc. </param>
    ''' <param name="DataBaseFullName">Full path of database name</param>
    ''' <param name="DeleteMessageAlert">Delete message alert if database already attached to server</param>
    ''' <param name="NetworkedFolder" >True if file is attached from a Networked Computer</param>
    ''' <returns>Completion flag</returns>
    ''' <remarks></remarks>
    Public Function AttachDataBase(ByVal ServerName As String, ByVal DataBaseFullName As String, Optional ByVal DeleteMessageAlert As Boolean = True, Optional ByVal NetworkedFolder As Boolean = False) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim afiles As List(Of String) = GF1.FullFileNameToList(DataBaseFullName)
            ServerName = ConvertFromSrv0(ServerName)
            Dim MDFAttached As Boolean = DataBaseExists(ServerName, afiles(1))
            Dim Mattach As Boolean = True
            If MDFAttached And DeleteMessageAlert Then
                If MsgBox("Are you sure to remove database?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    DropDataBase(ServerName, afiles(1))
                    Mattach = True
                Else
                    Mattach = False
                End If
            End If
            If Not Mattach Then
                MsgBox("File was not attached")
                Exit Function
            End If
            Dim mdffile As String = afiles(0).Trim & "\" & afiles(1).Trim & ".mdf"
            IO.File.SetAttributes(mdffile, IO.FileAttributes.Normal)
            Dim ldffile As String = afiles(0).Trim & "\" & afiles(1).Trim & "_log.ldf"
            Dim mdbcc As String = IIf(NetworkedFolder = True, "DBCC TRACEON(1807, -1)" & vbCrLf, "")
            Dim qry As String = mdbcc & "EXEC sp_attach_db @dbname = '" & afiles(1) & "' , @filename1 = N'" & mdffile & "' ,@filename2 = N'" & ldffile & "'  "
            GlobalControl.Variables.ErrorString = qry
            SqlExecuteNonQuery(ServerName, "master", qry)
            Return True
            Exit Function
        Catch ex As Exception
            QuitError(ex, Err, " Unable to execute DATAFUNCTION.AttachDataBase(ByVal ServerName As String, ByVal DataBaseFullName As String, Optional ByVal DeleteMessageAlert As Boolean = True, Optional ByVal NetworkedFolder As Boolean = False) As Boolean")
        End Try
        Return False
    End Function
    ''' <summary>
    ''' Attach a database to server
    ''' </summary>
    ''' <param name="ServerName">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="DataBaseFolder">Folder name where database exists</param>
    ''' <param name="DataBaseName">Name of Database</param>
    ''' <param name="DeleteMessageAlert">Delete message alert if database already attached to server</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AttachDataBase(ByVal ServerName As String, ByVal DataBaseFolder As String, ByVal DataBaseName As String, Optional ByVal DeleteMessageAlert As Boolean = True) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim afiles As New List(Of String)
            ServerName = ConvertFromSrv0(ServerName)
            DataBaseName = ConvertFromMdf0(DataBaseName)
            afiles.Add(DataBaseFolder)
            afiles.Add(DataBaseName)
            afiles.Add("mdf")
            Dim MDFAttached As Boolean = DataBaseExists(ServerName, afiles(1))
            Dim Mattach As Boolean = True
            If MDFAttached And DeleteMessageAlert Then
                If MsgBox("Are you sure to remove database?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    DropDataBase(ServerName, afiles(1))
                    Mattach = True
                Else
                    Mattach = False
                End If
            End If
            If Not Mattach Then
                MsgBox("File was not attaced")
                Exit Function
            End If
            Dim mdffile As String = afiles(0).Trim & "\" & afiles(1).Trim & ".mdf"
            Dim ldffile As String = afiles(0).Trim & "\" & afiles(1) & "_log.ldf"
            Dim qry As String = "EXEC sp_attach_db @dbname = '" & afiles(1) & "' , @filename1 = '" & mdffile & "' ,@filename2 = '" & ldffile & "'  "
            GlobalControl.Variables.ErrorString = qry
            SqlExecuteNonQuery(ServerName, "master", qry)
            Return True
            Exit Function
        Catch ex As Exception
            QuitError(ex, Err, " Unable to execute DATAFUNCTION.AttachDataBase(ByVal ServerName As String, ByVal DataBaseFolder As String, ByVal DataBaseName As String, Optional ByVal DeleteMessageAlert As Boolean = True) As Boolean")
        End Try
        Return False
    End Function
    ''' <summary>
    ''' Detach connected database to the server
    ''' </summary>
    ''' <param name="ServerName">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="DataBaseName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DetachDataBase(ByVal ServerName As String, ByVal DataBaseName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            ServerName = ConvertFromSrv0(ServerName)
            DataBaseName = ConvertFromMdf0(DataBaseName)
            Dim MDFAttached As Boolean = DataBaseExists(ServerName, DataBaseName)
            If MDFAttached = True Then
                Dim str As String = "ALTER DATABASE " & DataBaseName.Trim & " SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
                GlobalControl.Variables.ErrorString = str
                SqlExecuteNonQuery(ServerName, "master", str)
                Dim qry As String = "EXEC sp_detach_db @dbname = '" & DataBaseName.Trim & "'  "
                GlobalControl.Variables.ErrorString = qry
                SqlExecuteNonQuery(ServerName, "master", qry)
                Return True
                Exit Function
            Else
                QuitMessage("Database not attached", " DetachDataBase(ByVal ServerName As String, ByVal DataBaseName As String) As Boolean ")
                Return False
            End If
        Catch ex As Exception
            QuitError(ex, Err, " Unable to execute DATAFUNCTION.DetachDataBase(ByVal ServerName As String, ByVal DataBaseName As String) As Boolean")
        End Try
        Return False
    End Function

    ''' <summary>
    ''' To create a new database 
    ''' </summary>
    ''' <param name="SqlServer">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="FullFileName">Database name with path</param>
    ''' <remarks></remarks>

    Public Sub CreateDataBase(ByVal SqlServer As String, ByVal FullFileName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            SqlServer = ConvertFromSrv0(SqlServer)
            Dim afiles As List(Of String) = GF1.FullFileNameToList(FullFileName)
            If Not DataBaseExists(SqlServer, afiles(1).Trim) Then
                Dim LSqlStr As String = "Create DataBase " & afiles(1) & " on primary"
                Select Case GlobalControl.Variables.SqlVersion
                    Case 2005
                        LSqlStr = LSqlStr & " " _
                                          & " (NAME = " & afiles(1) & ", " _
                                          & "FILENAME = '" & FullFileName.Trim & "', " _
                                          & "SIZE = 2MB, MAXSIZE = 2000MB, FILEGROWTH = 10%) " _
                                          & "LOG ON (NAME = " & afiles(1).Trim & "_Log, " _
                                          & "FILENAME = '" & afiles(0).Trim & "\" & afiles(1).Trim & "_log.ldf', " _
                                          & "SIZE = 1MB, " _
                                          & "MAXSIZE = 1000MB, " _
                                          & "FILEGROWTH = 10%)"
                    Case 2012
                        LSqlStr = LSqlStr & " " _
                                          & " (NAME = " & afiles(1) & ", " _
                                          & "FILENAME = '" & FullFileName.Trim & "', " _
                                          & "SIZE = 5120KB, MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB) " _
                                          & "LOG ON (NAME = " & afiles(1).Trim & "_Log, " _
                                          & "FILENAME = '" & afiles(0).Trim & "\" & afiles(1).Trim & "_log.ldf', " _
                                          & "SIZE = 2048KB, " _
                                          & "MAXSIZE = 2048GB, " _
                                          & "FILEGROWTH = 10%)"
                End Select

                GlobalControl.Variables.ErrorString = LSqlStr
                SqlExecuteNonQuery(SqlServer, "master", LSqlStr)

            End If
        Catch ex As Exception
            QuitError(ex, Err, " Unable to execute DATAFUNCTION.CreateDataBase(ByVal SqlServer As String, ByVal FullFileName As String)")
        End Try
    End Sub
    ''' <summary>
    ''' To create a new database 
    ''' </summary>
    ''' <param name="SqlServer">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="DataBaseFolder">DataBaseFolder name</param>
    ''' <param name="DataBaseName">DataBase Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function CreateDataBase(ByVal SqlServer As String, ByVal DataBaseFolder As String, ByVal DataBaseName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            SqlServer = ConvertFromSrv0(SqlServer)
            DataBaseName = ConvertFromMdf0(DataBaseName)
            Dim afiles As New List(Of String)
            afiles.Add(DataBaseFolder)
            afiles.Add(DataBaseName)
            afiles.Add("mdf")
            If Not DataBaseExists(SqlServer, afiles(1).Trim) Then
                Dim FullFileName As String = GF1.GetFullFileName(afiles)
                Dim LSqlStr As String = "Create DataBase " & afiles(1) & " on primary"
                Select Case GlobalControl.Variables.SqlVersion
                    Case 2005
                        LSqlStr = LSqlStr & " " _
                                          & " (NAME = " & afiles(1) & ", " _
                                          & "FILENAME = '" & FullFileName.Trim & "', " _
                                          & "SIZE = 2MB, MAXSIZE = 2000MB, FILEGROWTH = 10%) " _
                                          & "LOG ON (NAME = " & afiles(1).Trim & "_Log, " _
                                          & "FILENAME = '" & afiles(0).Trim & "\" & afiles(1).Trim & "_log.ldf', " _
                                          & "SIZE = 1MB, " _
                                          & "MAXSIZE = 1000MB, " _
                                          & "FILEGROWTH = 10%)"
                    Case 2012
                        LSqlStr = LSqlStr & " " _
                                          & " (NAME = " & afiles(1) & ", " _
                                          & "FILENAME = '" & FullFileName.Trim & "', " _
                                          & "SIZE = 5120KB, MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB) " _
                                          & "LOG ON (NAME = " & afiles(1).Trim & "_Log, " _
                                          & "FILENAME = '" & afiles(0).Trim & "\" & afiles(1).Trim & "_log.ldf', " _
                                          & "SIZE = 2048KB, " _
                                          & "MAXSIZE = 2048GB, " _
                                          & "FILEGROWTH = 10%)"
                End Select
                GlobalControl.Variables.ErrorString = LSqlStr
                SqlExecuteNonQuery(SqlServer, "master", LSqlStr)
                Return True
                Exit Function
            End If
        Catch ex As Exception
            QuitError(ex, Err, " Unable to execute DATAFUNCTION.CreateDataBase(ByVal SqlServer As String, ByVal DataBaseFolder As String, ByVal DataBaseName As String) As Boolean")
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Copy sql database to another location
    ''' </summary>
    ''' <param name="SourceFolder"></param>
    ''' <param name="SourceDatabaseName"></param>
    ''' <param name="TargetFolder"></param>
    ''' <param name="TargetDatabaseName"></param>
    ''' <remarks></remarks>


    Private Sub CopyDatabase(ByVal SourceFolder As String, ByVal SourceDatabaseName As String, ByVal TargetFolder As String, ByVal TargetDatabaseName As String)
        Try
            System.IO.File.Copy(SourceFolder & SourceDatabaseName & ".mdf", TargetFolder & TargetDatabaseName & ".mdf", True)
            System.IO.File.Copy(SourceFolder & SourceDatabaseName & "_log.ldf", TargetFolder & TargetDatabaseName & "_log.ldf", True)
        Catch ex As Exception
            QuitError(ex, Err, SourceDatabaseName & "   " & TargetDatabaseName)
        End Try
    End Sub
    ''' <summary>
    ''' To check wether a server name contained by local server
    ''' </summary>
    ''' <param name="LocalServer1">Local Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="FindServerName">Server Name as string to be searched in local server</param>
    ''' <returns>Existing flag</returns>
    ''' <remarks></remarks>
    Public Function ServerExists(ByVal LocalServer1 As String, ByVal FindServerName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim flag As Boolean = False
        Try
            LocalServer1 = ConvertFromSrv0(LocalServer1)
            FindServerName = ConvertFromSrv0(FindServerName)
            Dim con As SqlConnection = OpenSqlConnection(LocalServer1, "master")
            Dim qry As String = "select * from sysservers where srvname = '" & FindServerName.Trim & "'"
            Dim serverTable As New DataTable
            Dim ad As SqlDataAdapter
            GlobalControl.Variables.ErrorString = qry
            ad = New SqlDataAdapter(qry, con)
            ad.Fill(serverTable)
            If serverTable Is Nothing Then
                flag = False
            Else
                flag = IIf(serverTable.Rows.Count > 0, True, flag)
                serverTable.Dispose()
            End If
            ad.Dispose()
            con.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return flag
    End Function
    ''' <summary>
    ''' To check wether a database name contained by server
    ''' </summary>
    ''' <param name="ServerName">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="FindDataBaseName">Database name to be checked by existence</param>
    ''' <returns>Existing Flag</returns>
    ''' <remarks></remarks>

    Public Function DataBaseExists(ByVal ServerName As String, ByVal FindDataBaseName As String) As Boolean

        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim flag As Boolean = False
        Try
            ServerName = ConvertFromSrv0(ServerName)
            FindDataBaseName = ConvertFromMdf0(FindDataBaseName)
            Dim con As SqlConnection = OpenSqlConnection(ServerName, "master")
            Dim qry As String = "select name from sysdatabases where name ='" & FindDataBaseName.Trim & "'"
            Dim MDFTable As New DataTable
            Dim ad As SqlDataAdapter
            GlobalControl.Variables.ErrorString = qry
            ad = New SqlDataAdapter(qry, con)
            ad.Fill(MDFTable)
            If MDFTable Is Nothing Then
                flag = False
            Else
                flag = IIf(MDFTable.Rows.Count > 0, True, flag)
                MDFTable.Dispose()
            End If
            ad.Dispose()
            con.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return flag
    End Function
    ''' <summary>
    ''' To check wether a table name exists in a database 
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or 0_srv_0.0_mdf_0 format</param>
    ''' <param name="FindTableName">Table name to be searched</param>
    ''' <returns>Flag for existence</returns>
    ''' <remarks></remarks>
    Public Function TableExists(ByVal ServerDataBase As String, ByVal FindTableName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim flag As Boolean = False
        Try
            Dim LserverDataBase As String = GetServerDataBase(ServerDataBase)
            Dim con As SqlConnection = OpenSqlConnection(LserverDataBase)
            Dim qry As String = "SELECT name FROM sys.Tables where name ='" & FindTableName.Trim & "'"
            Dim MDFTable As New DataTable
            Dim ad As SqlDataAdapter
            GlobalControl.Variables.ErrorString = qry
            ad = New SqlDataAdapter(qry, con)
            ad.Fill(MDFTable)
            If MDFTable Is Nothing Then
                flag = False
            Else
                flag = IIf(MDFTable.Rows.Count > 0, True, flag)
                MDFTable.Dispose()
            End If
            ad.Dispose()
            con.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return flag
    End Function
    ''' <summary>
    ''' To check wether a table name exists in a sql connection
    ''' </summary>
    ''' <param name="LSqlConnection ">Opened sql connection</param>
    ''' <param name="FindTableName">Table name to be searched</param>
    ''' <returns>Flag for existence</returns>
    ''' <remarks></remarks>
    Public Function TableExists(ByVal LSqlConnection As SqlConnection, ByVal FindTableName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim flag As Boolean = False
        Try
            Dim qry As String = "SELECT name FROM sys.Tables where name ='" & FindTableName.Trim & "'"
            Dim MDFTable As New DataTable
            Dim ad As SqlDataAdapter
            GlobalControl.Variables.ErrorString = qry
            ad = New SqlDataAdapter(qry, LSqlConnection)
            ad.Fill(MDFTable)
            If MDFTable Is Nothing Then
                flag = False
            Else
                flag = IIf(MDFTable.Rows.Count > 0, True, flag)
                MDFTable.Dispose()
            End If
            ad.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return flag
    End Function



    ''' <summary>
    ''' To check wether a table name exists in a database
    ''' </summary>
    ''' <param name="SQServer">Server name as constant or as in the convention 0_srv_0 ,1_srv_1 etc.</param>
    ''' <param name="SqDataBase">Sql Database name</param>
    ''' <param name="FindTableName">Table name to be searched</param>
    ''' <returns>Flag for existence</returns>
    ''' <remarks></remarks>
    Public Function TableExists(ByVal SQServer As String, ByVal SqDataBase As String, ByVal FindTableName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim flag As Boolean = False
        Try
            Dim con As SqlConnection = OpenSqlConnection(SQServer, SqDataBase)
            Dim qry As String = "SELECT name FROM sys.Tables where name ='" & FindTableName.Trim & "'"
            Dim MDFTable As New DataTable
            Dim ad As SqlDataAdapter
            GlobalControl.Variables.ErrorString = qry
            ad = New SqlDataAdapter(qry, con)
            ad.Fill(MDFTable)
            If MDFTable Is Nothing Then
                flag = False
            Else
                flag = IIf(MDFTable.Rows.Count > 0, True, flag)
                MDFTable.Dispose()
            End If
            ad.Dispose()
            con.Close()
            Return flag
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return flag
    End Function

    ''' <summary>
    ''' To add a remote server to local server
    ''' </summary>
    ''' <param name="Detach">True if Remote server droped without prompting if already linked </param>
    ''' <remarks></remarks>

    Public Sub LinkServer(Optional ByVal detach As Boolean = True)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LocalServer As String = ConvertFromSrv0(GlobalControl.Variables.AllServers("0_srv_0").ToString.Trim)
            Dim RemoteServerName As String = ConvertFromSrv0(GlobalControl.Variables.AllServers("1_srv_1").ToString.Trim)
            If RemoteServerName.Length = 0 Then
                Exit Sub
            End If
            If LCase(LocalServer) = LCase(RemoteServerName) Then
                Exit Sub
            End If
            Dim mUserId As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserName").ToString.Trim
            Dim mUserPwd As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserPassword").ToString.Trim
            Dim RuserId As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName").ToString.Trim
            Dim RuserPwd As String = GF1.GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerPassword").ToString.Trim
            If detach = False Then
                If ServerExists(LocalServer, RemoteServerName) Then
                    If MsgBox("Server already linked ,Click OK to drop linked server", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
                        DropServer(LocalServer, RemoteServerName)
                    End If
                End If
            Else
                If ServerExists(LocalServer, RemoteServerName) Then
                    DropServer(LocalServer, RemoteServerName)
                End If
            End If
            If Not ServerExists(LocalServer, RemoteServerName) Then
                Dim con As SqlConnection = OpenSqlConnection(LocalServer, "master")
                '                Dim qryRegSrv As String = " Exec sp_addlinkedserver '" & RemoteServerName & "' "
                Dim qryRegSrv As String = "EXEC sp_addlinkedserver @server = '" & RemoteServerName & "' "
                qryRegSrv = qryRegSrv & ", @srvproduct = ''"
                qryRegSrv = qryRegSrv & ", @provider = 'MSDASQL'"
                qryRegSrv = qryRegSrv & ", @provstr = 'DRIVER={SQL Server};SERVER=" & LocalServer & ";UID=" & RuserId & ";PWD=" & RuserPwd & ";'"
                GlobalControl.Variables.ErrorString = qryRegSrv
                Dim cmd As New SqlCommand(qryRegSrv, con)
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                con.Close()
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
    End Sub
    ''' <summary>
    ''' To drop existing Linked server
    ''' </summary>
    ''' <param name="LocalServer1">Local server name</param>
    ''' <param name="RemoteServerName ">Remote server name to be linked to local server</param>
    ''' <remarks></remarks>
    Public Sub DropServer(ByVal LocalServer1 As String, ByVal RemoteServerName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            LocalServer1 = ConvertFromSrv0(LocalServer1)
            RemoteServerName = ConvertFromSrv0(RemoteServerName)
            If ServerExists(LocalServer1, RemoteServerName) Then
                Dim con As SqlConnection = OpenSqlConnection(LocalServer1, "master")
                Dim qryRegSrv As String = " Exec sp_dropserver '" & RemoteServerName & "' "
                GlobalControl.Variables.ErrorString = qryRegSrv
                Dim cmd As New SqlCommand(qryRegSrv, con)
                cmd.ExecuteNonQuery()
                cmd.Dispose()
                con.Close()
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
    End Sub


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or 0_srv_0.0_mdf_0 format</param>
    ''' <param name="Ltable">Sql Table names of  FROM CLAUSE . if SeverDataBase is blank then these names must be full qualifier table names eg. (Server1.SQLBASE1.DBO.TABLE1) or from the mapped table names such as "_srv_0._mdf_0.table1 where srv0 is a key of global collection of servers and mdf0 is a key of global collection of databases</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="LJoinStmt" >SQL JOIN CLAUSE of querry</param>
    ''' <param name="Lcondition">String condition after Where Clause </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="RecordPosition">Record posintion  F=FirstRecord,L=LastRecord, or "*"=All</param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If LfieldList.Trim.Length = 0 Then
            LfieldList = "*"
        End If
        Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Dim Ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            Select Case LCase(RecordPosition)
                Case "f"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case "l"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case Else
                    qrystr = "select " & LfieldList & " from " & Ltable & " m1 "
            End Select
            If LFilter.Length > 0 Then
                Lcondition = Lcondition & IIf(Lcondition.Length = 0, "", " and ") & LFilter
            End If
            qrystr = qrystr & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
            qrystr = qrystr & IIf(Lcondition.Length > 0, " where " & Lcondition, "")
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
                Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            If PrimaryCols.Trim.Length = 0 Then
                TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Else
                SetPrimaryColumns(Ldatatable, PrimaryCols)
            End If
            If TableExists(cons, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = LserverDatabase & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            TempDA.Dispose()
            cons.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="Sql_Connection" >Database sql connection</param>
    ''' <param name="Ltable">Sql Table names of  FROM CLAUSE . if SeverDataBase is blank then these names must be full qualifier table names eg. (Server1.SQLBASE1.DBO.TABLE1) or from the mapped table names such as "_srv_0._mdf_0.table1 where srv0 is a key of global collection of servers and mdf0 is a key of global collection of databases</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="LJoinStmt" >SQL JOIN CLAUSE of querry</param>
    ''' <param name="Lcondition">String condition after Where Clause </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="RecordPosition">Record posintion  F=FirstRecord,L=LastRecord, or "*"=All</param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSql(ByVal Sql_Connection As SqlConnection, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If LfieldList.Trim.Length = 0 Then
            LfieldList = "*"
        End If
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Dim Ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            Select Case LCase(RecordPosition)
                Case "f"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case "l"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case Else
                    qrystr = "select " & LfieldList & " from " & Ltable & " m1 "
            End Select
            If LFilter.Length > 0 Then
                Lcondition = Lcondition & IIf(Lcondition.Length = 0, "", " and ") & LFilter
            End If
            qrystr = qrystr & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
            qrystr = qrystr & IIf(Lcondition.Length > 0, " where " & Lcondition, "")
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, Sql_Connection)
            If PrimaryCols.Trim.Length = 0 Then
                TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Else
                SetPrimaryColumns(Ldatatable, PrimaryCols)
            End If
            If TableExists(Sql_Connection, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = Sql_Connection.DataSource & "." & Sql_Connection.Database & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            TempDA.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ClsObject" >Class object of Sql table</param>
    ''' <param name="Lcondition">String condition after Where Clause </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="aClsObject" >An array of all classobjects if there values are assigned or element of expression to any column of  datatable</param>
    ''' <param name="HashPublicValues" >A hashtable with keys of all public variables of the form</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSql(ByRef ClsObject As Object, ByVal Lcondition As String, Optional ByVal LFilter As String = "", Optional ByVal Lorder As String = "", Optional ByVal aClsObject As Object = Nothing, Optional ByVal HashPublicValues As Hashtable = Nothing) As DataTable
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Ldatatable As DataTable = ClsObject.PrevDt
        Ldatatable.Rows.Clear()
        Try
            Dim LserverDatabase As String = GetServerDataBase(ClsObject.ServerDataBase)
            Dim LTable As String = ClsObject.TableName
            Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
            Dim qrystr As String = GetDataQuery(ClsObject, Lcondition, LFilter, Lorder)
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            If TableExists(cons, LTable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & LTable & " not found in SQL Database"
            End If
            Dim ldatatable1 As New DataTable
            TempDA.Fill(ldatatable1)
            If LTable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = LserverDatabase & ".dbo." & LTable
            Else
                Ldatatable.TableName = LTable
            End If
            TempDA.Dispose()
            cons.Close()
            If ldatatable1.Rows.Count = 0 Then
                Return Ldatatable
                Exit Function
            End If
            Ldatatable = UpdateDataTables(Ldatatable, ldatatable1)
            Ldatatable = ReplaceGroupFieldsValueInDt(ClsObject, mGroupFieldsType, Ldatatable, aClsObject, HashPublicValues)
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function
    Private Function ReplaceGroupFieldsValueInDt(ByVal ClsObject As Object, ByVal mGroupFieldsType As Hashtable, ByVal LdataTable As DataTable, ByVal aclsObject As Object, ByVal HashPublicValues As Hashtable) As DataTable
        Dim mHash As New Hashtable
        For x = 0 To mGroupFieldsType.Count - 1
            Dim mkey As String = UCase(mGroupFieldsType.Keys(x))
            If InStr("GCHF", mkey) = 0 Then
                Continue For
            End If
            Dim mValue As String = GF1.GetValueFromHashTable(mGroupFieldsType, mkey)
            mHash = GF1.AddItemToHashTable(mHash, mkey, mValue)
        Next
        If mHash.Count = 0 Then
            Return LdataTable
            Exit Function
        End If

        For z = 0 To LdataTable.Rows.Count - 1
            For i = 0 To mHash.Count - 1
                Dim mkey As String = UCase(mHash.Keys(i))
                Dim mValue As String = GF1.GetValueFromHashTable(mHash, mkey)
                Select Case mkey
                    Case "C"
                        Dim afields() As String = mValue.Split("|")
                        For k = 0 To afields.Count - 1
                            Dim xfield() As String = afields(k).Split("#")
                            Dim mField As String = xfield(0)
                            Dim mexprField As String = xfield(1)
                            LdataTable.Rows(z).Item("Expr" & mField) = mexprField
                        Next
                    Case "G"
                        Dim afields() As String = mValue.Split("|")
                        For k = 0 To afields.Count - 1
                            Dim xfield() As String = afields(k).Split("#")
                            Dim mField As String = xfield(0)
                            Dim mexprField As String = xfield(1)
                            Dim evalue As Object = EvalCGroupTypeField(afields(k), aclsObject, HashPublicValues)
                            LdataTable.Rows(z).Item("Gxpr" & mField) = mexprField
                            If evalue IsNot Nothing Then
                                LdataTable.Rows(z).Item(mField) = evalue
                            End If
                        Next
                        'Case "A"
                        '    Dim afields() As String = mValue.Split("|")
                        '    For k = 0 To afields.Count - 1
                        '        Dim xfield() As String = afields(k).Split("#")
                        '        Dim mField As String = xfield(0)
                        '        Dim mexprField As String = xfield(1)
                        '        Dim evalue As Object = EvalAGroupTypeField(afields(k), aclsObject, HashPublicValues)
                        '        If evalue IsNot Nothing Then
                        '            LdataTable.Rows(z).Item(mField) = evalue
                        '        End If
                        '    Next
                    Case "F"
                        Dim afields() As String = mValue.Split("|")
                        For k = 0 To afields.Count - 1
                            Dim xfield() As String = afields(k).Split("#")
                            If xfield.Count = 2 Then
                                Dim mField As String = xfield(0)
                                Dim mFormat As String = xfield(1)
                                LdataTable.Rows(z).Item("Frmt" & mField) = GF1.FormattedValue(ClsObject.PrevDt.Rows(z).Item(mField), mFormat)
                            End If
 Next
                    Case "H"
                        Dim afields() As String = mValue.Split("|")
                        For k = 0 To afields.Count - 1
                            Dim xfield() As String = afields(k).Split("#")
                            Dim mField As String = xfield(0)
                            Dim mHashKey As String = xfield(6)
                            Dim mhash1 As Hashtable = GF1.CreateHashTable(mHashKey, ClsObject.PrevDt.Rows(z).Item(mField))
                            LdataTable.Rows(z).Item("Hash" & mField) = mhash1
                        Next
                End Select
            Next
        Next
        Return LdataTable
    End Function


        ''' <summary>
        ''' To get data table from Sql Table on specified order and conditions 
        ''' </summary>
        ''' <param name="ClsObject" >Class object of Sql table</param>
        ''' <param name="Hcondition" > condition as hashtable after Where Clause </param>
        ''' <param name="LFilter" >Filter criteria added in where clause</param>
        ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
        ''' <param name="aClsObject" >An array of all classobjects if there values are assigned or element of expression to any column of  datatable</param>
        ''' <param name="HashPublicValues" >A hashtable with keys of all public variables of the form</param>
        ''' <returns>Data Table Object</returns>
        ''' <remarks></remarks>
    Public Function GetDataFromSql(ByRef ClsObject As Object, ByVal Hcondition As Hashtable, Optional ByVal LFilter As String = "", Optional ByVal Lorder As String = "", Optional ByVal aClsObject As Object = Nothing, Optional ByVal HashPublicValues As Hashtable = Nothing) As DataTable
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcondition As String = GF1.GetStringConditionFromHashTable(Hcondition, True, True, " and ")
        Dim Ldatatable As DataTable = GetDataFromSql(ClsObject, Lcondition, LFilter, Lorder, aClsObject, HashPublicValues)
        Return Ldatatable
    End Function

    ''' <summary>
    ''' To check if duplicate value exists for a particular column in a table as per given condition
    ''' </summary>
    ''' <param name="clsTableclass">Table class of table in which value has to be searched </param>
    ''' <param name="colname">Name of column whose value is to be searched</param>
    ''' <param name="colValue">Value of column which needs to be check for duplicate</param>
    ''' <param name="WhereClause">additional condition which needs to be applied in query while searching for duplicate</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function checkIfDuplicateExists(ByVal clsTableclass As Object, ByVal colname As String, ByVal colValue As Object, Optional ByVal WhereClause As String = "") As Boolean

        Dim aClsObject() As Object = {clsTableclass}
        Dim mserverdb As String = GetServerMDFForTransanction(aClsObject)
        Dim mytrans As SqlTransaction = BeginTransaction(mserverdb)
        Dim dt As New DataTable
        dt = GetDataFromSql(clsTableclass, WhereClause)
        For i = 0 To dt.Rows.Count - 1
            If dt.Rows(i).Item(colname) = colValue Then
                MsgBox("Duplicate value for " & colname & " are not allowed, Please choose other value")
                Return False
                Exit For
            End If
        Next
        Return True
    End Function


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ClsObject" >Class object of Sql table</param>
    ''' <param name="Lcondition">String condition after Where Clause </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Private Function GetDataQuery(ByRef ClsObject As Object, ByVal Lcondition As String, Optional ByVal LFilter As String = "", Optional ByVal Lorder As String = "") As String
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        ' Dim LserverDatabase As String = GetServerDataBase(ClsObject.ServerDataBase)
        '    Dim LTable As String = ClsObject.TableName
        Dim TableFullPath As String = ClsObject.TableName
        If LCase("WebAzure,WebGodaddy,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            TableFullPath = ClsObject.TableWithSQLPath
        End If
        Dim QryStr As String = ""
        Try
            Dim JoinStr As String = ""
            Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
            Dim mFieldAlias As String = ""
            Dim JoinOnClause As String = ""
            If GF1.GetValueFromHashTable(mGroupFieldsType, "X") IsNot Nothing Then
                Dim Group_x As String = GF1.GetValueFromHashTable(mGroupFieldsType, "X")
                If Group_x.Length > 0 Then
                    Dim aGroup_x() As String = Group_x.Split("|")
                    For i = 0 To aGroup_x.Count - 1
                        '   mfieldnames  # LinkTable # InfoCode # ControlType # ValueType of ValueProperty of control type #  LinkTableTextField # LinkTableTextFieldType # LinkTableServerDatabase # LinkPrimarykey
                        Dim aaGroup_x() As String = aGroup_x(i).Split("#")
                        Dim mField As String = aaGroup_x(0)
                        Dim mLinkTable As String = aaGroup_x(1)
                        Dim mLinkTableTextField As String = aaGroup_x(6)
                        Dim mLinkTableServerDatabase As String = aaGroup_x(8)
                        Dim mLinkPrimarykey As String = aaGroup_x(9)
                        Dim LinkTableFullName As String = mLinkTable
                        If LCase("WebAzure,WebGodaddy,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
                            LinkTableFullName = GetServerDataBase(mLinkTableServerDatabase, True) & ".DBO." & mLinkTable
                        End If
                        Dim fieldString As String = "s" & i.ToString & "." & mLinkTableTextField & " As " & "Text" & mField
                        mFieldAlias = mFieldAlias & IIf(mFieldAlias.Length = 0, "", ",") & fieldString
                        Dim mjoinstring As String = " Left Outer Join " & LinkTableFullName & " s" & i.ToString & " on m1." & mField & " = " & "s" & i.ToString & "." & mLinkPrimarykey & " "
                        JoinOnClause = JoinOnClause & mjoinstring
                    Next
                End If
            End If
            Dim mAllFields As String = ClsObject.AllFields
            'Dim sfields As String = ""
            'For i = 0 To mAllFields.Count - 1
            '    Dim mfield As String = mAllFields(i).ToString.Trim
            '    If InStr(mfield, "m1") = 0 Then
            '        mfield = "m1." & mfield
            '    End If
            '    sfields = sfields & IIf(sfields.Length = 0, "", ",") & mfield
            'Next

            QryStr = "select " & mAllFields & IIf(mFieldAlias.Length > 0, "," & mFieldAlias, "") & " from " & TableFullPath & " m1 "

            If JoinOnClause.Length > 0 Then
                QryStr = QryStr & " " & JoinOnClause
            End If
            If ClsObject.WhereClauseDefault.Length > 0 Then
                Lcondition = ClsObject.WhereClauseDefault & IIf(Lcondition.Length = 0, "", " and ") & Lcondition
            End If

            If LFilter.Length > 0 Then
                Lcondition = Lcondition & IIf(Lcondition.Length = 0, "", " and ") & LFilter
            End If
            QryStr = QryStr & IIf(Lcondition.Length > 0, " where " & Lcondition, "")
            QryStr = QryStr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
        Catch ex As Exception
            QuitError(ex, Err, "")
        End Try
        Return QryStr
    End Function
    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="MyTrans" >Sql Transanction</param>
    ''' <param name="ClsObject" >Class object of Sql table</param>
    ''' <param name="Lcondition">String condition after Where Clause </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="aClsObject" >An array of all classobjects if there values are assigned or element of expression to any column of  datatable</param>
    ''' <param name="HashPublicValues" >A hashtable with keys of all public variables of the form</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSql(ByRef MyTrans As SqlTransaction, ByRef ClsObject As Object, ByVal Lcondition As String, Optional ByVal LFilter As String = "", Optional ByVal Lorder As String = "", Optional ByVal aClsObject As Object = Nothing, Optional ByVal HashPublicValues As Hashtable = Nothing) As DataTable
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Ldatatable As DataTable = ClsObject.PrevDt
        Try
            Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
            Dim qrystr As String = GetDataQuery(ClsObject, Lcondition, LFilter, Lorder)
            Dim Ldatatable1 As DataTable = SqlExecuteDataTable(MyTrans, qrystr)
            If Ldatatable1.Rows.Count = 0 Then
                ClsObject.Prevdt.Rows.clear()
                Return Ldatatable
                Exit Function
            End If
            Ldatatable = UpdateDataTables(Ldatatable, Ldatatable1)
            Ldatatable = ReplaceGroupFieldsValueInDt(ClsObject, mGroupFieldsType, Ldatatable, aClsObject, HashPublicValues)
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function

    ''' <summary>
    ''' To get data row from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ClsObject" >Class object of Sql table</param>
    ''' <param name="PrimaryKeyValue" >Primary key value to be searched</param>
    ''' <param name="aClsObject" >An array of all classobjects if there values are assigned or element of expression to any column of  datatable</param>
    ''' <param name="HashPublicValues" >A hashtable with keys of all public variables of the form</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function SeekRecordFromSql(ByRef ClsObject As Object, ByVal PrimaryKeyValue As Integer, Optional ByVal aClsObject As Object = Nothing, Optional ByVal HashPublicValues As Hashtable = Nothing) As DataRow
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBase(ClsObject.ServerDataBase)
        Dim LTable As String = ClsObject.TableName
        Dim Ldatatable As New DataTable
        Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
        Dim mrow As DataRow = ClsObject.NewRow
        Try
            Dim Lcondition As String = ClsObject.PrimaryKey & " = " & PrimaryKeyValue.ToString
            Dim qrystr As String = GetDataQuery(ClsObject, Lcondition)
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            If TableExists(cons, LTable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & LTable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If LTable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = LserverDatabase & ".dbo." & LTable
            Else
                Ldatatable.TableName = LTable
            End If
            TempDA.Dispose()
            cons.Close()
            If Ldatatable.Rows.Count = 0 Then
                ClsObject.PrevRow = ClsObject.NewRow
                Return ClsObject.PrevRow
                Exit Function
            End If
            mrow = UpdateDataRows(mrow, Ldatatable.Rows(0))
            mrow = ReplaceGroupFieldsValueInRow(mGroupFieldsType, mrow, aClsObject, HashPublicValues)
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        ClsObject.PrevRow.ItemArray = mrow.ItemArray
        Return ClsObject.PrevRow
    End Function
    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ClsObject" >Class object of Sql table</param>
    ''' <param name="WhereClause" >A hash table having the conditions of where clause , key is field name,value is fieldvalue,logical gate=And</param>
    ''' <param name="aClsObject" >An array of all classobjects if there values are assigned or element of expression to any column of  datatable</param>
    ''' <param name="HashPublicValues" >A hashtable with keys of all public variables of the form</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function SeekRecordFromSql(ByRef ClsObject As Object, ByVal WhereClause As Hashtable, Optional ByVal aClsObject As Object = Nothing, Optional ByVal HashPublicValues As Hashtable = Nothing) As DataRow
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcondition As String = GF1.GetStringConditionFromHashTable(WhereClause, True)
        Dim LserverDatabase As String = GetServerDataBase(ClsObject.ServerDataBase)
        Dim LTable As String = ClsObject.TableName
        Dim Ldatatable As New DataTable
        Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
        Dim mrow As DataRow = ClsObject.NewRow
        Try
            '          Dim Lcondition As String = ClsObject.PrimaryKey & " = " & PrimaryKeyValue.ToString
            '          Dim Lcondition As String = ClsObject.PrimaryKey & " = " & PrimaryKeyValue.ToString
            Dim qrystr As String = GetDataQuery(ClsObject, Lcondition)
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            If TableExists(cons, LTable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & LTable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If LTable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = LserverDatabase & ".dbo." & LTable
            Else
                Ldatatable.TableName = LTable
            End If
            TempDA.Dispose()
            cons.Close()
            If Ldatatable.Rows.Count = 0 Then
                ClsObject.PrevRow = ClsObject.NewRow
                Return ClsObject.PrevRow
                Exit Function
            End If
            mrow = UpdateDataRows(mrow, Ldatatable.Rows(0))
            mrow = ReplaceGroupFieldsValueInRow(mGroupFieldsType, mrow, aClsObject, HashPublicValues)
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        ClsObject.PrevRow.ItemArray = mrow.ItemArray
        Return ClsObject.PrevRow
    End Function
    Private Function ReplaceGroupFieldsValueInRow(ByVal mGroupFieldsType As Hashtable, ByVal mRow As DataRow, ByVal aclsObject As Object, ByVal HashPublicValues As Hashtable) As DataRow
        For i = 0 To mGroupFieldsType.Count - 1
            Dim mkey As String = UCase(mGroupFieldsType.Keys(i))
            If InStr("CGHF", mkey) = 0 Then
                Continue For
            End If
            Dim mValue As String = GF1.GetValueFromHashTable(mGroupFieldsType, mkey)
            Select Case mkey
                Case "C"
                    Dim afields() As String = mValue.Split("|")
                    For k = 0 To afields.Count - 1
                        Dim xfield() As String = afields(k).Split("#")
                        Dim mField As String = xfield(0)
                        Dim mexprField As String = xfield(1)
                        mRow.Item("Expr" & mField) = mexprField
                    Next
                Case "G"
                    Dim afields() As String = mValue.Split("|")
                    For k = 0 To afields.Count - 1
                        Dim xfield() As String = afields(k).Split("#")
                        Dim mField As String = xfield(0)
                        Dim mexprField As String = xfield(1)
                        mRow.Item("Gxpr" & mField) = mexprField
                        Dim evalue As Object = EvalCGroupTypeField(afields(k), aclsObject, HashPublicValues)
                        If evalue IsNot Nothing Then
                            mRow.Item(mField) = evalue
                        End If
                    Next

                    'Case "A"
                    '    Dim afields() As String = mValue.Split("|")
                    '    For k = 0 To afields.Count - 1
                    '        Dim xfield() As String = afields(k).Split("#")
                    '        Dim mField As String = xfield(0)
                    '        Dim mexprField As String = xfield(1)
                    '        Dim evalue As Object = EvalAGroupTypeField(afields(k), aclsObject, HashPublicValues)
                    '        If evalue IsNot Nothing Then
                    '            mRow.Item(mField) = evalue
                    '        End If
                    '    Next
                Case "F"
                    Dim afields() As String = mValue.Split("|")
                    For k = 0 To afields.Count - 1
                        Dim xfield() As String = afields(k).Split("#")
                        Dim mField As String = xfield(0)
                        Dim mFormat As String = xfield(1)
                        mRow.Item("Frmt" & mField) = GF1.FormattedValue(mRow.Item(mField), mFormat)
                    Next
                Case "H"
                    Dim afields() As String = mValue.Split("|")
                    For k = 0 To afields.Count - 1
                        Dim xfield() As String = afields(k).Split("#")
                        Dim mField As String = xfield(0)
                        Dim mHashKey As String = xfield(6)
                        Dim mhash1 As Hashtable = GF1.CreateHashTable(mHashKey, mRow.Item(mField))
                        mRow.Item("Hash" & mField) = mhash1
                    Next
            End Select
        Next
        Return mRow
    End Function
        ''' <summary>
        ''' To get data table from Sql Table on specified order and conditions 
        ''' </summary>
        ''' <param name="ClsObject" >Class object of Sql table</param>
        ''' <param name="PrimaryKeyValue" >Primary key value to be searched</param>
        ''' <param name="aClsObject" >An array of all classobjects if there values are assigned or element of expression to any column of  datatable</param>
        ''' <param name="HashPublicValues" >A hashtable with keys of all public variables of the form</param>
        ''' <returns>Data Table Object</returns>
        ''' <remarks></remarks>
    Public Function SeekRecordFromSql(ByRef MyTrans As SqlTransaction, ByRef ClsObject As Object, ByVal PrimaryKeyValue As Integer, Optional ByVal aClsObject As Object = Nothing, Optional ByVal HashPublicValues As Hashtable = Nothing) As DataRow
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LTable As String = ClsObject.TableName
        Dim Ldatatable As New DataTable
        Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
        Dim mrow As DataRow = ClsObject.NewRow
        Try
            Dim Lcondition As String = ClsObject.PrimaryKey & " = " & PrimaryKeyValue.ToString
            Dim qrystr As String = GetDataQuery(ClsObject, Lcondition)
            Dim mrow1 As DataRow = SeekRecord(MyTrans, LTable, PrimaryKeyValue)
            If Ldatatable.Rows.Count = 0 Then
                    ClsObject.PrevRow = ClsObject.NewRow
                Return ClsObject.PrevRow
                Exit Function
            End If
            mrow = UpdateDataRows(mrow, mrow1)
            mrow = ReplaceGroupFieldsValueInRow(mGroupFieldsType, mrow, aClsObject, HashPublicValues)
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        ClsObject.PrevRow.ItemArray = mrow.ItemArray
        Return mrow
    End Function


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or 0_srv_0.0_mdf_0 format</param>
    ''' <param name="Ltable">Sql Table names of  FROM CLAUSE . if SeverDataBase is blank then these names must be full qualifier table names eg. (Server1.SQLBASE1.DBO.TABLE1) or from the mapped table names such as "_srv_0._mdf_0.table1 where srv0 is a key of global collection of servers and mdf0 is a key of global collection of databases</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="LJoinStmt" >SQL JOIN CLAUSE of querry</param>
    ''' <param name="Hcondition">Condition as hashtable,where key is field and value is condition value with equality operator and logical gate AND </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="RecordPosition">Record posintion  F=FirstRecord,L=LastRecord, or "*"=All</param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Hcondition As Hashtable, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If LfieldList.Trim.Length = 0 Then
            LfieldList = "*"
        End If
        Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Dim Ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            Select Case LCase(RecordPosition)
                Case "f"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case "l"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case Else
                    qrystr = "select " & LfieldList & " from " & Ltable & " m1 "
            End Select
            Dim LCondition As String = GF1.GetStringConditionFromHashTable(Hcondition, True)
            If LFilter.Length > 0 Then
                LCondition = LCondition & IIf(LCondition.Length = 0, "", " and ") & LFilter
            End If
            qrystr = qrystr & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
            qrystr = qrystr & IIf(Lcondition.Length > 0, " where " & Lcondition, "")
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            If PrimaryCols.Trim.Length = 0 Then
                TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Else
                SetPrimaryColumns(Ldatatable, PrimaryCols)
            End If
            If TableExists(cons, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = LserverDatabase & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            TempDA.Dispose()
            cons.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="SqlTableIdentifier" >Full identifier of a sqltable with server name eg. server0.database0.dbo.table or 0_srv_0.0_mdf_0.dbo.table format</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="LJoinStmt" >SQL JOIN CLAUSE of querry</param>
    ''' <param name="Hcondition">Condition as hashtable,where key is field and value is condition value with equality operator and logical gate AND </param>
    ''' <param name="LFilter" >Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="RecordPosition">Record posintion  F=FirstRecord,L=LastRecord, or "*"=All</param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSqlIdentifier(ByVal SqlTableIdentifier As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Hcondition As Hashtable, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If LfieldList.Trim.Length = 0 Then
            LfieldList = "*"
        End If
        Dim LserverDatabase As String = GetServerDataBaseFromSqlIdentifier(SqlTableIdentifier)
        Dim Ltable As String = GetTableNameFromSqlIdentifier(SqlTableIdentifier)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Dim Ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            Select Case LCase(RecordPosition)
                Case "f"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case "l"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case Else
                    qrystr = "select " & LfieldList & " from " & Ltable & " m1 "
            End Select
            Dim LCondition As String = GF1.GetStringConditionFromHashTable(Hcondition, True)
            If LFilter.Length > 0 Then
                LCondition = LCondition & IIf(LCondition.Length = 0, "", " and ") & LFilter
            End If
            qrystr = qrystr & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
            qrystr = qrystr & IIf(LCondition.Length > 0, " where " & LCondition, "")
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            If PrimaryCols.Trim.Length = 0 Then
                TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Else
                SetPrimaryColumns(Ldatatable, PrimaryCols)
            End If
            If TableExists(cons, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                Ldatatable.TableName = LserverDatabase & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            TempDA.Dispose()
            cons.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function


    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or 0_srv_0.0_mdf_0 format</param>
    ''' <param name="Ltable">Sql Table names of  FROM CLAUSE . if SeverDataBase is blank then these names must be full qualifier table names eg. (Server1.SQLBASE1.DBO.TABLE1) or from the mapped table names such as "_srv_0._mdf_0.table1 where srv0 is a key of global collection of servers and mdf0 is a key of global collection of databases</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, Optional ByVal Lorder As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim Ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            qrystr = "select  *  from " & Ltable & " m1 "
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            If TableExists(cons, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                LserverDatabase = AddSquareBrackets(LserverDatabase)
                Ldatatable.TableName = LserverDatabase & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            TempDA.Dispose()
            cons.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function
    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="SqlTableIdentifier" >Full identifier of a sqltable with server name eg. server0.database0.dbo.table or 0_srv_0.0_mdf_0.dbo.table format</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSqlIdentifier(ByVal SqlTableIdentifier As String, Optional ByVal Lorder As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBaseFromSqlIdentifier(SqlTableIdentifier)
        Dim Ltable As String = GetTableNameFromSqlIdentifier(SqlTableIdentifier)
        Dim Ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            qrystr = "select  *  from " & Ltable & " m1 "
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            If TableExists(cons, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                LserverDatabase = AddSquareBrackets(LserverDatabase)
                Ldatatable.TableName = LserverDatabase & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            TempDA.Dispose()
            cons.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Ldatatable
    End Function
    ''' <summary>
    ''' To get data table from Sql Table on specified order and conditions 
    ''' </summary>
    ''' <param name="SqlTableIdentifier" >Full identifier of a sqltable with server name eg. server0.database0.dbo.table or 0_srv_0.0_mdf_0.dbo.table format</param>
    ''' <returns>DatRow  tempplate</returns>
    ''' <remarks></remarks>
    Public Function GetNewRowTemplateFromSqlIdentifier(ByVal SqlTableIdentifier As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBaseFromSqlIdentifier(SqlTableIdentifier)
        Dim Ltable As String = GetTableNameFromSqlIdentifier(SqlTableIdentifier)
        Dim Ldatatable As New DataTable
        Dim lrow As DataRow = Nothing
        Try
            Dim qrystr As String = ""
            qrystr = "select top (1)  *  from " & Ltable & " m1 "
            Dim cons As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim TempDA As New SqlDataAdapter
            GlobalControl.Variables.ErrorString = qrystr
            TempDA = New SqlDataAdapter(qrystr, cons)
            TempDA.MissingSchemaAction = MissingSchemaAction.AddWithKey
            If TableExists(cons, Ltable) = False Then
                GlobalControl.Variables.ErrorString = "Table Name " & Ltable & " not found in SQL Database"
            End If
            TempDA.Fill(Ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                LserverDatabase = AddSquareBrackets(LserverDatabase)
                Ldatatable.TableName = LserverDatabase & ".dbo." & Ltable
            Else
                Ldatatable.TableName = Ltable
            End If
            lrow = Ldatatable.NewRow
            Return lrow
            TempDA.Dispose()
            cons.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return lrow
    End Function

 



    ''' <summary>
    ''' To get a  string of  server.database format from  0_srv_0.0_mdf_0 or srv1.mdf1 fromat
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or 0_srv_0.0_mdf_0 format </param>
    ''' <param name="WithBrackets" >True, if servername will be enclosed in square brackets
    ''' </param>
    ''' <returns>Comma Separated string of SQL databases</returns>
    ''' <remarks></remarks>
    Public Function GetServerDataBase(ByVal ServerDataBase As String, Optional ByVal WithBrackets As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Sdatabase As String = ConvertFromSrv0Mdf0(ServerDataBase, WithBrackets)
        Return Sdatabase
    End Function
    ''' <summary>
    ''' To convert  a querry string in which tables are in the format (0_srv_0.0_mdf_0.dbo.table1)  are converted into format (server0.database0.dbo.table0)  
    ''' </summary>
    ''' <param name="StringToConvert">String having tables in the format (0_srv_0.0_mdf_0.dbo.table1) </param>
    ''' <param name="WithBrackets" >True, if servername will be enclosed in square brackets</param>
    ''' <returns>String of SQL databases with servers</returns>
    ''' <remarks></remarks>
    Public Function ConvertFromSrv0Mdf0(ByVal StringToConvert As String, Optional ByVal WithBrackets As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            GlobalControl.Variables.ErrorString = StringToConvert
            If StringToConvert.Contains("_srv_") = True Then
                For nn = 0 To GlobalControl.Variables.AllServers.Count - 1
                    Dim mkey As String = CStr(nn) & "_srv_" & CStr(nn)
                    Dim mvalue As String = GlobalControl.Variables.AllServers.Item(mkey)
                    If WithBrackets = True Then
                        If Left(mvalue, 1) <> "[" Then
                            mvalue = "[" & mvalue & "]"
                        End If
                    End If
                    StringToConvert = StringToConvert.Replace(mkey, mvalue)
                Next
            End If
            If StringToConvert.Contains("_mdf_") = True Then
                For nn = 0 To GlobalControl.Variables.MDFFiles.Count - 1
                    Dim mkey As String = CStr(nn) & "_mdf_" & CStr(nn)
                    Dim mvalue As String = GlobalControl.Variables.MDFFiles.Item(mkey)
                    If WithBrackets = True Then
                        If Left(mvalue, 1) <> "[" Then
                            mvalue = "[" & mvalue & "]"
                        End If
                    End If
                    StringToConvert = StringToConvert.Replace(mkey, mvalue)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ConvertFromSrv0Mdf0 (" & GlobalControl.Variables.ErrorString)
        End Try
        Return StringToConvert
    End Function


    ''' <summary>
    ''' To convert  a server string in which server name is in the format (_srv_0)  are converted into fixed name  
    ''' </summary>
    ''' <param name="SeverString">Server string in the format 0_srv_0 or fixed servername </param>
    ''' <param name="WithBrackets" >True, if servername will be enclosed in square brackets</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertFromSrv0(ByVal SeverString As String, Optional ByVal WithBrackets As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            GlobalControl.Variables.ErrorString = SeverString
            If SeverString.Contains("_srv_") = True Then
                For nn = 0 To GlobalControl.Variables.AllServers.Count - 1
                    Dim mkey As String = CStr(nn) & "_srv_" & CStr(nn)
                    Dim mvalue As String = GlobalControl.Variables.AllServers.Item(mkey)
                    If WithBrackets = True Then
                        If Left(mvalue, 1) <> "[" Then
                            mvalue = "[" & mvalue & "]"
                        End If
                    End If
                    SeverString = SeverString.Replace(mkey, mvalue)
                Next
            Else
                If WithBrackets = True Then
                    If Left(SeverString, 1) <> "[" Then
                        SeverString = "[" & SeverString & "]"
                    End If
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ConvertFromSrv0 (" & GlobalControl.Variables.ErrorString)
        End Try
        Return SeverString
    End Function
    ''' <summary>
    ''' To convert  a mdf name string in which database name is in the format (mdf0)  are converted into fixed name  
    ''' </summary>
    ''' <param name="MdfString">database string in the format 0_mdf_0 or fixed databasename</param>
    ''' <param name="WithBrackets" >True, if servername will be enclosed in square brackets</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertFromMdf0(ByVal MdfString As String, Optional ByVal WithBrackets As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            GlobalControl.Variables.ErrorString = MdfString
            If MdfString.Contains("_mdf_") = True Then
                For nn = 0 To GlobalControl.Variables.MDFFiles.Count - 1
                    Dim mkey As String = CStr(nn) & "_mdf_" & CStr(nn)
                    Dim mvalue As String = GlobalControl.Variables.MDFFiles.Item(mkey)
                    If WithBrackets = True Then
                        If Left(mvalue, 1) <> "[" Then
                            mvalue = "[" & mvalue & "]"
                        End If
                    End If
                    MdfString = MdfString.Replace(mkey, mvalue)
                Next
            Else
                If WithBrackets = True Then
                    If Left(MdfString, 1) <> "[" Then
                        MdfString = "[" & MdfString & "]"
                    End If
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ConvertFrommdf0 (" & GlobalControl.Variables.ErrorString)
        End Try
        Return MdfString
    End Function



    ''' <summary>
    ''' To get datatable of fixed rows from sql table on specified order and conditions 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format , if space or used table names are full identifier then this will extracted from full table names</param>
    ''' <param name="Ltable">Sql Table names (For more than one table ,comma should be used as delimited,If serverdatabase="" ,Table names must be full identifier ,such as SQLBase1.DBO.Table1 ,SQLBase2.DBO.Table3 etc. or  mdf0.table1,mdbf3.table2,mdf4.table2 ,where mdf0,mdf3,mdf4 are the keys of a collection having databasenames</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="LJoinStmt">SQL Join Clause of querry</param>
    ''' <param name="Lcondition">String condition after Where Clause</param>
    ''' <param name="Lfilter">Filter criteria added in where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="NoOfRows">No of rows to be populated</param>
    ''' <param name="StartPoint">A collection object having keys col0=Name of keyfield of main table,col1=Start Value of keyfield,Concanated start value of Order columns </param>
    ''' <param name="NavigstionType">"F"=Forward accessing of main table,"R"=Reverse accessing of main table</param>
    ''' <param name="StartRowPostion ">Rows count in the main table from top to the start point on this condition and filter criteria</param>
    ''' <param name="TotalRows  ">Total Rows in the main table on this condition and filter criteria </param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFixedRows(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal Lfilter As String, ByVal Lorder As String, ByVal NoOfRows As Integer, ByRef StartPoint As Collection, ByVal NavigstionType As String, ByRef StartRowPostion As Integer, ByRef TotalRows As Integer, Optional ByVal PrimaryCols As String = "") As DataTable
        'StartPoint collection
        'Name of Keyfield -col0
        'Value of FiKey field -col1
        'Start Value of orderkey  -col2
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LDataTable As DataTable = Nothing
        Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Try
            Dim Aorder As String() = Split(Lorder, ",")
            Dim OrderFields As String = Replace(Lorder, ",", "+")
            Dim OrderKey As String = ""
            If Aorder.Length > 1 Then
                LfieldList = LfieldList & "," & OrderFields & " as m1.OrderKey"
                OrderKey = "OrderKey"
            Else
                OrderKey = Aorder(0)  'changes by Neha
            End If
            Dim LKeyField As String = StartPoint.Item("col0"), LKeyStart As String = StartPoint.Item("col1"), OrderStart As String = StartPoint.Item("col2")
            Lorder = Lorder & "," & LKeyField

            Dim SQLStr1 As String = "", SqlStr2 As String = ""

            If Len(LTrim(RTrim(LKeyStart))) > 0 Then
                SQLStr1 = "select " & "TOP " & NoOfRows & " " & LfieldList & " from " & Ltable & " m1"
                SQLStr1 = SQLStr1 & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
                SqlStr2 = SQLStr1
                If UCase(NavigstionType) = "F" Then
                    SQLStr1 = SQLStr1 & " where " & "m1." & OrderKey & " = '" & OrderStart & "'" & " and " & "m1." & LKeyField & "  >= '" & LKeyStart & "'"
                    SqlStr2 = SqlStr2 & " where " & "m1." & OrderKey & " > '" & OrderStart & "'"
                    If Lcondition.Length > 0 Then
                        SQLStr1 = SQLStr1 & " and " & Lcondition
                        SqlStr2 = SqlStr2 & " and " & Lcondition
                    End If
                    If Len(LTrim(RTrim(Lfilter))) > 0 Then
                        SQLStr1 = SQLStr1 & " and " & Lfilter
                        SqlStr2 = SqlStr2 & " and " & Lfilter
                    End If
                    SQLStr1 = SQLStr1 & " order by " & "m1." & OrderKey & "," & LKeyField
                    SqlStr2 = SqlStr2 & " order by " & "m1." & OrderKey & "," & LKeyField
                End If
                If UCase(NavigstionType) = "R" Then
                    SQLStr1 = SQLStr1 & " where " & "m1." & OrderKey & " = '" & OrderStart & "'" & " and " & "m1." & LKeyField & "  <= '" & LKeyStart & "'"
                    SqlStr2 = SqlStr2 & " where " & "m1." & OrderKey & "<  '" & OrderStart & "'"
                    If Lcondition.Length > 0 Then
                        SQLStr1 = SQLStr1 & " and " & Lcondition
                        SqlStr2 = SqlStr2 & " and " & Lcondition
                    End If
                    If Len(LTrim(RTrim(Lfilter))) > 0 Then
                        SQLStr1 = SQLStr1 & " and " & Lfilter
                        SqlStr2 = SqlStr2 & " and " & Lfilter
                    End If
                    SQLStr1 = SQLStr1 & " order by " & "m1." & OrderKey & " desc," & LKeyField & " desc"
                    SqlStr2 = SqlStr2 & " order by " & "m1." & OrderKey & " desc," & LKeyField & " desc"
                End If
                GlobalControl.Variables.ErrorString = SQLStr1

                LDataTable = SqlExecuteDataTable(LserverDatabase, SQLStr1)

                If LDataTable.Rows.Count < NoOfRows Then
                    Dim dtable2 As DataTable = SqlExecuteDataTable(LserverDatabase, SqlStr2)
                    Dim n As Integer = NoOfRows - LDataTable.Rows.Count
                    Dim m As Integer = dtable2.Rows.Count
                    If m >= 1 Then
                        Dim Addrows As Integer = IIf(m < n, m, n)
                        For ii = 0 To Addrows - 1
                            LDataTable.ImportRow(dtable2.Rows(ii))
                        Next
                    End If
                End If


                If StartRowPostion = -1 Then
                    Dim SqlStr3 As String = "select count(*) from " & Ltable & " m1 where " & "m1." & OrderKey & "< " & "'" & LTrim(RTrim(OrderStart)) & "'" & ""
                    If Lcondition.Length > 0 Then
                        SqlStr3 = SqlStr3 & " and " & Lcondition
                    End If
                    If Len(LTrim(RTrim(Lfilter))) > 0 Then
                        SqlStr3 = SqlStr3 & " and " & Lfilter
                    End If
                    '  SqlStr3 = SqlStr3 & " order by " & "m1." & OrderKey & "," & "m1." & LKeyField
                    StartRowPostion = CInt(SqlExecuteScalarQuery(LserverDatabase, SqlStr3))
                End If

                If TotalRows = -1 Then
                    Dim sqlstr3 As String = "select count(*) from " & Ltable & " m1 "
                    Dim Lwhere As Boolean = False
                    If Lcondition.Length > 0 Then
                        sqlstr3 = sqlstr3 & " where " & IIf(Lwhere = False, "", " and ") & Lcondition
                        Lwhere = True
                    End If
                    If Len(LTrim(RTrim(Lfilter))) > 0 Then
                        sqlstr3 = sqlstr3 & IIf(Lwhere = False, " where ", " and ") & Lfilter
                    End If
                    ' sqlstr3 = sqlstr3 & " order by " & "m1." & OrderKey & "," & "m1." & LKeyField
                    TotalRows = CInt(SqlExecuteScalarQuery(LserverDatabase, sqlstr3))
                End If







                Return LDataTable
            Else
                If StartRowPostion = -1 Then
                    Dim SqlStr3 As String = "select count(*) from " & Ltable & " m1 where " & "m1." & OrderKey & "< " & "'" & LTrim(RTrim(OrderStart)) & "'" & ""
                    If Lcondition.Length > 0 Then
                        SqlStr3 = SqlStr3 & " and " & Lcondition
                    End If
                    If Len(LTrim(RTrim(Lfilter))) > 0 Then
                        SqlStr3 = SqlStr3 & " and " & Lfilter
                    End If
                    SqlStr3 = SqlStr3 & " order by " & "m1." & OrderKey & "," & "m1." & LKeyField
                    StartRowPostion = CInt(SqlExecuteScalarQuery(LserverDatabase, SqlStr3))
                End If

                If TotalRows = -1 Then
                    Dim sqlstr3 As String = "select count(*) from " & Ltable & " m1 "
                    Dim Lwhere As Boolean = False
                    If Lcondition.Length > 0 Then
                        sqlstr3 = sqlstr3 & IIf(Lwhere = False, "", " and ") & Lcondition
                        Lwhere = True
                    End If
                    If Len(LTrim(RTrim(Lfilter))) > 0 Then
                        sqlstr3 = sqlstr3 & IIf(Lwhere = False, "", " and ") & Lfilter
                    End If
                    sqlstr3 = sqlstr3 & " order by " & "m1." & OrderKey & "," & "m1." & LKeyField
                    TotalRows = CInt(SqlExecuteScalarQuery(LserverDatabase, sqlstr3))
                End If
                ' recpoint = dbCls.SqlExecuteScalarQuery(GlobalClasssLibrary.GlobalVarClass.SqlServerName, funcCls.processDbname(Trim(DatabaseName)), sql2)
                SQLStr1 = "select " & "TOP " & NoOfRows & "  " & LfieldList & "  from " & Ltable & " m1"
                SQLStr1 = SQLStr1 & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
                SQLStr1 = SQLStr1 & " where " & "m1." & OrderKey & ">= '" & LTrim(RTrim(OrderStart)) & "'"
                If Lcondition.Length > 0 Then
                    SQLStr1 = SQLStr1 & " and " & Lcondition
                End If
                If Len(LTrim(RTrim(Lfilter))) > 0 Then
                    SQLStr1 = SQLStr1 & " and " & Lfilter
                End If
                SQLStr1 = SQLStr1 & " order by " & "m1." & OrderKey & "," & LKeyField
                LDataTable = SqlExecuteDataTable(LserverDatabase, SQLStr1)
                Return LDataTable
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GetDataFixedRows(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal Lfilter As String, ByVal Lorder As String, ByVal NoOfRows As Integer, ByRef StartPoint As Collection, ByVal NavigstionType As String, ByRef StartRowPostion As Integer, ByRef TotalRows As Integer, Optional ByVal PrimaryCols As String = "") As DataTable")
        End Try
        Return LDataTable
    End Function

    Public Function getRowFromP_value(ByVal serverdatabase As String, ByVal tablename As String, ByVal columnstring As String, ByVal p_fieldname As String, ByVal p_fieldvalue As Integer) As DataRow
        Dim dtr As DataRow = Nothing
        Dim strsql As String = "select " & columnstring & " from " & tablename & " where " & p_fieldname & " = " & p_fieldvalue & " and rowstatus = 0"
        Dim dt As DataTable = SqlExecuteDataTable(serverdatabase, strsql)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)
        End If
        Return dtr
    End Function


   
    ''' <summary>
    ''' To get datatable of fixed rows from sql table on specified order and conditions 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format , if space or used table names are full identifier then this will extracted from full table names</param>
    ''' <param name="Ltable">Sql Table names (For more than one table ,comma should be used as delimited,If serverdatabase="" ,Table names must be full identifier ,such as SQLBase1.DBO.Table1 ,SQLBase2.DBO.Table3 etc. or  mdf0.table1,mdbf3.table2,mdf4.table2 ,where mdf0,mdf3,mdf4 are the keys of a collection having databasenames</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="LJoinStmt">SQL Join Clause of querry</param>
    ''' <param name="Lcondition">String condition after Where Clause</param>
    ''' <param name="Lfilter">Filter criteria to get page offset value within the above condition</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="StartRowPostion ">Rows count in the main table from top to the start point on this condition and filter criteria</param>
    ''' <param name="NoOfRows">No of rows to be populated</param>
    ''' <param name="TotalRows  ">Total Rows in the main table on this condition and filter criteria, calculated if -1 passed </param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns</param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromSqlFixedRows(ByRef ServerDataBase As String, ByRef Ltable As String, ByRef LfieldList As String, ByRef LJoinStmt As String, ByRef Lcondition As String, ByRef Lfilter As String, ByRef Lorder As String, ByRef StartRowPostion As Integer, ByVal NoOfRows As Integer, ByRef TotalRows As Integer, Optional ByVal PrimaryCols As String = "") As DataTable
        Dim LDataTable As DataTable = Nothing
        Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Dim Lcond As String = ""
        Try
            Dim OrderKeyVal As String = ""  'change by divya
            Dim Aorder As String() = Split(Lorder.ToLower, ",")
            If Aorder.Length > 1 Then
                For i = 0 To Aorder.Length - 1
                    If Aorder(i).Contains(" asc") Then
                        OrderKeyVal += Replace(Aorder(i), " asc", "") + "+"
                    ElseIf Aorder(i).Contains(" desc") Then
                        OrderKeyVal += Replace(Aorder(i), " desc", "") + "+"
                    End If
                Next
                If OrderKeyVal(OrderKeyVal.Length - 1) = "+" Then
                    OrderKeyVal = OrderKeyVal.Remove(OrderKeyVal.Length - 1)
                End If
                LfieldList = LfieldList & "," & OrderKeyVal & " as OrderKey"
            End If

            If StartRowPostion = -1 Then
                Dim lfilter1 As String = Lfilter.Replace(">=", "<")
                lfilter1 = lfilter1.Replace("=", "<")
                Lcond = Lcondition
                If lfilter1.Length > 0 Then
                    Lcond = Lcondition & IIf(Lcondition.Trim.Length = 0, "", " and ") & lfilter1
                    StartRowPostion = RowsCount(LserverDatabase, Ltable, Lcond, Lorder) 'change by divya
                    StartRowPostion = StartRowPostion - NoOfRows \ 2
                End If
                If StartRowPostion < 0 Then
                    StartRowPostion = 0
                End If
            End If

            Dim SQLStr As String = "Select "
            If LfieldList.Trim.Length = 0 Then
                SQLStr = SQLStr & " * "
            Else
                SQLStr = SQLStr & "  " & LfieldList
            End If
            SQLStr = SQLStr & " from " & Ltable & " as  m1"
            If LJoinStmt.Trim.Length > 0 Then
                SQLStr = SQLStr & " " & LJoinStmt
            End If
            Lcond = ""
            If Lcondition.Trim.Length > 0 Then
                Lcond = " where " & Lcondition
            End If
            If Lcond.Trim.Length > 0 Then
                SQLStr = SQLStr & " " & Lcond
            End If
            If Lorder.Trim.Length > 0 Then
                SQLStr = SQLStr & " order by " & Lorder 'change by divya
            End If
            '   MsgBox(SQLStr)
            SQLStr = SQLStr & " offset " & StartRowPostion.ToString & " Rows Fetch Next " & NoOfRows.ToString & " Rows Only"
            GlobalControl.Variables.ErrorString = SQLStr
            '  GlobalControl.Variables.LastProcessingData = SQLStr
            LDataTable = SqlExecuteDataTable(LserverDatabase, SQLStr)
            If TotalRows = -1 Then
                Lcond = Lcondition
                Dim countSTr As String = " select count (*) "
                countSTr = countSTr & " from " & Ltable & " as  m1"
                If LJoinStmt.Trim.Length > 0 Then
                    countSTr = countSTr & " " & LJoinStmt
                End If
                Lcond = ""
                If Lcondition.Trim.Length > 0 Then
                    countSTr = countSTr & " " & Lcond
                End If
                Dim dt As DataTable = SqlExecuteDataTable(ServerDataBase, countSTr).Rows(0).Item(0)
                If Not dt Is Nothing Then
                    TotalRows = dt.Rows(0).Item(0)
                End If
            End If
        Catch ex As Exception
            '  rdc.QuitError(ex, Err, New StackTrace(True))
        End Try
        '  rdc.EndMethod()
        Return LDataTable
    End Function










    ''' <summary>
    ''' To get data table from Sql Table on specified order and condition 
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql transaction name</param>
    ''' <param name="Ltable">Sql Table name</param>
    ''' <param name="LJoinStmt" >SQL Join Clause of querry</param>
    ''' <param name="LfieldList">Comma separated field list to get (*) for all</param>
    ''' <param name="Lcondition">String condition after Where Clause </param>
    ''' <param name="Lfilter" >Additional filter string for where clause</param>
    ''' <param name="Lorder">Comma separated field string after Order by Clause</param>
    ''' <param name="RecordPosition">Record posintion  F=FirstRecord,L=LastRecord, or "*"=All</param>
    ''' <param name="PrimaryCols">Comma separated string of data table primary columns </param>
    ''' <returns>Data Table Object</returns>
    ''' <remarks></remarks>

    Public Function GetDataFromSql(ByRef Sql_Transaction As SqlTransaction, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal Lfilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        LJoinStmt = ConvertFromSrv0Mdf0(LJoinStmt, True)
        Lcommand.Transaction = Sql_Transaction
        Dim ldatatable As New DataTable
        Try
            Dim qrystr As String = ""
            Select Case LCase(RecordPosition)
                Case "f"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case "l"
                    qrystr = "select top (1)  " & LfieldList & " from " & Ltable & " m1 "
                Case Else
                    qrystr = "select " & LfieldList & " from " & Ltable & " m1 "
            End Select
            If Lfilter.Length > 0 Then
                Lcondition = Lcondition & IIf(Lcondition.Length = 0, "", " and ") & Lfilter
            End If
            qrystr = qrystr & IIf(LJoinStmt.Length > 0, " " & LJoinStmt, "")
            qrystr = qrystr & IIf(Lcondition.Length > 0, " where " & Lcondition, "")
            qrystr = qrystr & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            Lcommand.CommandText = qrystr
            GlobalControl.Variables.ErrorString = qrystr
            Dim sqlda As New SqlDataAdapter(Lcommand)
            If PrimaryCols.Trim.Length = 0 Then
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            Else
                SetPrimaryColumns(ldatatable, PrimaryCols)
            End If
            sqlda.Fill(ldatatable)
            If Ltable.IndexOf(".dbo.") < 0 Then
                Dim cons As SqlConnection = Sql_Transaction.Connection
                Dim mdatabase As String = AddSquareBrackets(cons.Database)
                Dim msource As String = AddSquareBrackets(cons.DataSource)
                ldatatable.TableName = msource & "." & mdatabase & ".dbo." & Ltable
            Else
                ldatatable.TableName = Ltable
            End If
            sqlda.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return ldatatable
        Lcommand.Dispose()
    End Function

    ''' <summary>
    ''' To get increamental last key of a data table
    ''' </summary>
    ''' <param name="LDataTable">Data table from which lastkey found</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="Prefix">Prefix of lastkey</param>
    ''' <param name="LastKeyFieldSize">Size of the keyfield </param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function LastKeyPlus(ByVal LDataTable As DataTable, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mLastKey As String = Prefix & Right("000000000001", LastKeyFieldSize - Prefix.Length)
        Try
            If LDataTable.Rows.Count = 0 Then
                Return mLastKey
            End If
            Dim lfilter As String = ""
            Dim FdataTable As New DataTable
            ' LDataTable .
            If Prefix.Length > 0 Then
                lfilter = LastKeyField & "  like  '" & Prefix & "%'"
            End If
            Dim msort() As String = {LastKeyField}
            LDataTable = SortFilterDataTable(LDataTable, msort, "ASC", lfilter)
            Dim mcount As Integer = LDataTable.Rows.Count - 1
            Dim mlastno As Integer = 0
            If mcount > -1 Then
                mlastno = CInt(Right(LDataTable.Rows(mcount)(LastKeyField), LastKeyFieldSize - Prefix.Length))
            End If
            mLastKey = Prefix & Right(CStr(10000000001 + mlastno), LastKeyFieldSize - Prefix.Length)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByVal LDataTable As DataTable, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer) As String")
        End Try
        Return mLastKey
    End Function
    ''' <summary>
    ''' Set final values of  currdt or currrow  of TableClass object according to InterFieldValues.
    ''' </summary>
    ''' <param name="clsObject">TableClass object after defining FieldsFinalValues property</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetFinalFieldsValuesOneTable(ByRef clsObject As Object) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If clsObject.SqlUpdation = False Then
            Return clsObject
            Exit Function
        End If

        If clsObject.FieldsFinalValues.count = 0 Then
            Return clsObject
            Exit Function
        End If
        Dim i As Integer, j As Integer
        Dim mLastValuePlusOne As New Hashtable
        If clsObject.PlusOneColumns.Length > 0 Then
            Dim aFields() As String = Split(clsObject.PlusOneColumns, ",")
            For j = 0 To aFields.Length - 1
                Dim mfield As String = aFields(j)
                Dim LastValue As Integer = LastValuePlusOne(clsObject.CurrDt, mfield, False)
                mLastValuePlusOne = GF1.AddItemToHashTable(mLastValuePlusOne, mfield, LastValue)
            Next
        End If
        If clsObject.MultyRowsSqlHandling = True Then
            For i = 0 To clsObject.CurrDt.Rows.Count - 1
                If clsObject.PlusOneFields.Length > 0 Then
                    Dim aFields() As String = Split(clsObject.PlusOneFields, ",")
                    For j = 0 To aFields.Length - 1
                        If GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j)) IsNot Nothing Then
                            If IsDBNull(clsObject.CurrDt.Rows(i).Item(aFields(j))) = True Then
                                clsObject.CurrDt.Rows(i).Item(aFields(j)) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j)) + i
                            End If
                        End If
                    Next
                End If
                If clsObject.SameValueFields.Length > 0 Then
                    Dim aFields() As String = Split(clsObject.SameValueFields, ",")
                    For j = 0 To aFields.Length - 1
                        If GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j)) IsNot Nothing Then
                            clsObject.CurrDt.Rows(i).Item(aFields(j)) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j))
                        End If
                    Next
                End If
                If clsObject.PlusOneColumns.Length > 0 Then
                    Dim aFields() As String = Split(clsObject.PlusOneColumns, ",")
                    For j = 0 To aFields.Length - 1
                        Dim mlastvalue As Integer = GF1.GetValueFromHashTable(mLastValuePlusOne, aFields(j))
                        If IsDBNull(clsObject.CurrDt.Rows(i).Item(aFields(j))) = True Then
                            clsObject.CurrDt.Rows(i).Item(aFields(j)) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j))
                            mlastvalue = mlastvalue + i
                            mLastValuePlusOne = GF1.AddItemToHashTable(mLastValuePlusOne, aFields(j), mlastvalue)
                        End If
                    Next
                End If
            Next
        Else
            Dim rFields() As String = {}
            If clsObject.PlusOneFields.Length > 0 Then
                Dim aFields() As String = Split(clsObject.PlusOneFields, ",")
                For j = 0 To aFields.Length - 1
                    If GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j)) IsNot Nothing Then
                        If IsDBNull(clsObject.CurrRow.Item(aFields(j))) = True Then
                            clsObject.CurrRow.Item(aFields(j)) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j))
                            rFields = GF1.ArrayAppend(rFields, LCase(aFields(j)))
                        End If
                    End If
                Next
            End If
            If clsObject.SameValueFields.Length > 0 Then
                Dim cFields() As String = Split(clsObject.SameValueFields, ",")
                For j = 0 To cFields.Length - 1
                    If GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, cFields(j)) IsNot Nothing Then
                        clsObject.CurrRow.Item(cFields(j)) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, cFields(j))
                        rFields = GF1.ArrayAppend(rFields, LCase(cFields(j)))
                    End If
                Next
            End If
            If clsObject.PlusOneColumns.Length > 0 Then
                Dim aFields() As String = Split(clsObject.PlusOneColumns, ",")
                For j = 0 To aFields.Length - 1
                    Dim mlastvalue As Integer = GF1.GetValueFromHashTable(mLastValuePlusOne, aFields(j))
                    If IsDBNull(clsObject.CurrRow.Item(aFields(j))) = True Then
                        clsObject.CurrDt.CurrRow.Item(aFields(j)) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, aFields(j))
                        rFields = GF1.ArrayAppend(rFields, LCase(aFields(j)))
                    End If
                Next
            End If

            For j = 0 To clsObject.FieldsFinalValues.count - 1
                Dim mkey As String = LCase(clsObject.FieldsFinalValues.ToString.Trim)
                If GF1.ArrayFind(rFields, mkey) > -1 Then
                    Continue For
                End If
                If GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, mkey) IsNot Nothing Then
                    clsObject.CurrRow.Item(mkey) = GF1.GetValueFromHashTable(clsObject.FieldsFinalValues, mkey)
                End If
            Next
        End If
        Return clsObject
    End Function
    ''' <summary>
    ''' Set final values of  currdt or currrow  of an array of TableClass object according to InterFieldValues.
    ''' </summary>
    ''' <param name="clsObject">An array of TableClass object after defining FieldsFinalValues property for each table class.</param>
    ''' <param name="HashPublicValues" >A hashtable having keys as variablenames and values are variable's values used in expressions to assign fieldvalues</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetFinalFieldsValues(ByRef clsObject() As Object, ByVal HashPublicValues As Hashtable) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim AGroupType As New Hashtable
        Dim CGroupType As New Hashtable
        Dim mHeaderRowStatusFlag As Boolean = False
        Dim mHeaderRowStatusNo As Int16 = 0

        For k = 0 To clsObject.Count - 1
            '  Dim mClsObject As Object = clsObject(k)
            If clsObject(k).SqlUpdation = False Then
                Continue For
            End If
            If clsObject(k).TableType = "H" Then
                mHeaderRowStatusFlag = clsObject(k).RowStatusFlag
                Dim lhash As Hashtable = clsObject(k).FieldsFinalValues
                If lhash.Count > 0 Then
                    'Dim valw As Object = GF1.GetValueFromHashTable(clsObject(k).FieldsFinalValues, "Rowstatus")
                    mHeaderRowStatusNo = GF1.GetValueFromHashTable(clsObject(k).FieldsFinalValues, "RowStatus")
                End If
                'changes by Neha
            End If
            Try
                Dim mFieldsFinalValues As Hashtable = clsObject(k).FieldsFinalValues
                Dim rfields() As String = {}
                If mFieldsFinalValues.Count > 0 Then
                    Dim i As Integer, j As Integer
                    Dim mLastValuePlusOne As New Hashtable
                    If clsObject(k).PlusOneColumns.Length > 0 Then
                        Dim aFields() As String = Split(clsObject(k).PlusOneColumns, ",")
                        For j = 0 To aFields.Length - 1
                            Dim mfield As String = aFields(j)
                            rfields = GF1.ArrayAppend(rfields, mfield)
                            Dim LastValue As Integer = Val(LastValuePlusOne(clsObject(k).CurrDt, mfield, False))  'changed by Neha
                            mLastValuePlusOne = GF1.AddItemToHashTable(mLastValuePlusOne, mfield, LastValue)
                        Next
                    End If
                    Dim mPlusOneFields As New Hashtable
                    If clsObject(k).PlusOneFields.Length > 0 Then
                        Dim aFields() As String = Split(clsObject(k).PlusOneFields, ",")
                        For j = 0 To aFields.Length - 1
                            Dim mfield As String = aFields(j)
                            rfields = GF1.ArrayAppend(rfields, mfield)
                            Dim LastValue As Integer = Val(GF1.GetValueFromHashTable(clsObject(k).FieldsFinalValues, aFields(j)))   'Changed by Neha
                            mPlusOneFields = GF1.AddItemToHashTable(mPlusOneFields, mfield, LastValue)
                        Next
                    End If
                    Dim mSameValueFields As New Hashtable
                    If clsObject(k).SameValueFields.Length > 0 Then
                        Dim aFields() As String = Split(clsObject(k).SameValueFields, ",")
                        For j = 0 To aFields.Length - 1
                            Dim mfield As String = aFields(j)
                            rfields = GF1.ArrayAppend(rfields, mfield)
                            Dim LastValue As Integer = Val(GF1.GetValueFromHashTable(clsObject(k).FieldsFinalValues, aFields(j)))    'changed by Neha
                            mSameValueFields = GF1.AddItemToHashTable(mSameValueFields, mfield, LastValue)
                        Next
                    End If

                    For i = 0 To clsObject(k).CurrDt.Rows.Count - 1

                        If mPlusOneFields.Count > 0 Then
                            Dim mkeys(mPlusOneFields.Count - 1) As String
                            mPlusOneFields.Keys.CopyTo(mkeys, 0)
                            For z = 0 To mkeys.Count - 1
                                Dim mkey As String = mkeys(z)
                                Dim mvalue As Object = mPlusOneFields.Item(mkey)
                                If mvalue IsNot Nothing Then
                                    If IsDBNull(clsObject(k).CurrDt.Rows(i).Item(mkey)) = True Then
                                        clsObject(k).CurrDt.Rows(i).Item(mkey) = mvalue
                                        mPlusOneFields = GF1.AddItemToHashTable(mPlusOneFields, mkey, mvalue + 1)
                                    End If
                                    If CInt(clsObject(k).CurrDt.Rows(i).Item(mkey)) < 0 Then
                                        clsObject(k).CurrDt.Rows(i).Item(mkey) = mvalue
                                        mPlusOneFields = GF1.AddItemToHashTable(mPlusOneFields, mkey, mvalue + 1)
                                    End If
                                End If
                            Next
                        End If
                        If clsObject(k).SameValueFields.Length > 0 Then
                            Dim mkeys(mSameValueFields.Count - 1) As String
                            mSameValueFields.Keys.CopyTo(mkeys, 0)
                            For z = 0 To mkeys.Count - 1
                                Dim mkey As String = mkeys(z)
                                Dim mvalue As Object = mSameValueFields.Item(mkey)
                                If mvalue IsNot Nothing Then
                                    clsObject(k).CurrDt.Rows(i).Item(mkey) = mvalue
                                End If
                            Next
                        End If

                        If clsObject(k).PlusOneColumns.Length > 0 Then
                            Dim mkeys(mLastValuePlusOne.Count - 1) As String
                            mLastValuePlusOne.Keys.CopyTo(mkeys, 0)
                            For z = 0 To mkeys.Count - 1
                                Dim mkey As String = mkeys(z)
                                Dim mvalue As Object = mLastValuePlusOne.Item(mkey)
                                If mvalue IsNot Nothing Then
                                    If IsDBNull(clsObject(k).CurrDt.Rows(i).Item(mkey)) = True Then
                                        clsObject(k).CurrDt.Rows(i).Item(mkey) = mvalue
                                        mLastValuePlusOne = GF1.AddItemToHashTable(mLastValuePlusOne, mkey, mvalue + 1)
                                    End If
                                    If CInt(clsObject(k).CurrDt.Rows(i).Item(mkey)) < 0 Then
                                        clsObject(k).CurrDt.Rows(i).Item(mkey) = mvalue
                                        mLastValuePlusOne = GF1.AddItemToHashTable(mLastValuePlusOne, mkey, mvalue + 1)
                                    End If
                                End If
                            Next
                        End If
                        For j = 0 To mFieldsFinalValues.Count - 1
                            Dim mkey As String = mFieldsFinalValues.Keys(j)
                            If GF1.ArrayFind(rfields, mkey) > -1 Then
                                Continue For
                            End If
                            If GF1.GetValueFromHashTable(clsObject(k).FieldsFinalValues, mkey) IsNot Nothing Then
                                clsObject(k).CurrDt.Rows(i).Item(mkey) = GF1.GetValueFromHashTable(mFieldsFinalValues, mkey)
                            End If
                        Next
                    Next
                End If
                Dim kstr As String = ""
                kstr = GF1.GetValueFromHashTable(clsObject(k).GroupFieldsType, "A")

                If Not kstr = "" Then
                    AGroupType = GF1.AddItemToHashTable(AGroupType, k.ToString, GF1.GetValueFromHashTable(clsObject(k).GroupFieldsType, "A"))
                End If
                Dim lstr As String = ""
                lstr = GF1.GetValueFromHashTable(clsObject(k).GroupFieldsType, "C")
                If Not lstr = "" Then
                    CGroupType = GF1.AddItemToHashTable(CGroupType, k.ToString, GF1.GetValueFromHashTable(clsObject(k).GroupFieldsType, "C"))
                End If
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute datafunction.SetFinalFieldsValues(ByRef clsObject() As Object) As Object()" & clsObject(k).TableName)
            End Try
        Next
        For i = 0 To AGroupType.Count - 1
            Dim mkey As String = AGroupType.Keys(i)
            Dim aValue() As String = AGroupType.Item(mkey).ToString.Split("|")
            For j = 0 To aValue.Count - 1
                Dim mField As String = aValue(j).Split("#")(0)
                Dim mValue As Object = EvalAGroupTypeField(aValue(j), clsObject, HashPublicValues)
                If mValue Is Nothing Then Continue For
                If Not mValue Is Nothing Then '   If  mValue IsNot Nothing Then      'Changes by Neha
                    For x = 0 To clsObject(CInt(mkey)).CurrDt.Rows.Count - 1
                        clsObject(CInt(mkey)).CurrDt.Rows(x).Item(mField) = mValue
                        Dim mentryType As String = ""
                        mentryType = clsObject(CInt(mkey)).TableEntryType
                        If Not mentryType = "A" Then
                            Dim mUpcolStr As String = clsObject(CInt(mkey)).CurrDt.Rows(x).item("UpdatedColumnsStr").ToString.Trim
                            mUpcolStr = mUpcolStr & IIf(mUpcolStr = "", "", ",") & mField
                            clsObject(CInt(mkey)).CurrDt.Rows(x).item("UpdatedColumnsStr") = mUpcolStr
                        End If
                    Next
                End If
            Next
        Next
        For i = 0 To CGroupType.Count - 1
            Dim mkey As String = CGroupType.Keys(i)
            Dim aValue() As String = CGroupType.Item(mkey).ToString.Split("|")
            For j = 0 To aValue.Count - 1
                Dim mvalue As Object = EvalCGroupTypeField(aValue(j), clsObject, HashPublicValues)
                If mvalue IsNot Nothing Then
                    Dim mfield As String = aValue(j).Split("#")(0)
                    For x = 0 To clsObject(CInt(mkey)).CurrDt.Rows.Count - 1
                        clsObject(CInt(mkey)).CurrDt.Rows(x).Item(mfield) = mvalue
                    Next
                End If
            Next
        Next
        For k = 0 To clsObject.Length - 1
            If clsObject(k).TableType = "S" Then
                clsObject(k).HeaderRowStatusFlag = mHeaderRowStatusFlag
                clsObject(k).HeaderRowStatusNo = mHeaderRowStatusNo
            End If
        Next
        Return clsObject
    End Function

    ''' <summary>
    ''' This Function returns an auto increased integer value for a pertcular column in a table as per given condition
    ''' </summary>
    ''' <param name="aTableClass">Tableclass of table in which value has to be searched</param>
    ''' <param name="columnName"> Name of the column whose value is to be searched</param>
    ''' <param name="WhereClause">Additional condition which needs to be applied in query for searching.</param>
    ''' <returns>This function returns an Integer value</returns>
    ''' <remarks></remarks>
    Public Function AutoIncreaseInfSno(ByRef aTableClass As Object, ByVal columnName As String, Optional ByVal WhereClause As String = "") As Integer
        Dim SnoValue As Integer = 1
        Dim query As String = "Select top 1 *from " & aTableClass.TableName & " where " & WhereClause & " order by " & columnName & " desc"
        Dim dt As New DataTable
        dt = SqlExecuteDataTable(aTableClass.ServerDatabase, query)
        If dt.Rows.Count > 0 Then
            SnoValue = GetCellValue(dt, columnName, 0, "Integer") + 1
        End If
        Return SnoValue
    End Function


    Private Function EvalAGroupTypeField(ByVal FieldAssignExpression As String, ByVal aclsObject() As Object, ByVal HashPublicValues As Hashtable) As Object
        'For i = 0 To AGroupType.Count - 1
        '    Dim mkey As String = AGroupType.Keys(i)
        '    Dim aValue() As String = AGroupType.Item(mkey).ToString.Split("|")
        '    For j = 0 To aValue.Count - 1
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim jvalue() As String = FieldAssignExpression.Split("#")
        Dim mfield As String = jvalue(0)
        Dim mexpr As String = jvalue(1)
        Dim aObject As New Object
        aObject = Nothing
        Select Case True
            Case InStr(mexpr, "__") > 0
                Dim d As Int16 = InStr(mexpr, "__")
                Dim dtable As String = Microsoft.VisualBasic.Left(mexpr, d - 1)
                Dim dfield As String = Microsoft.VisualBasic.Mid(mexpr, d + 2, mexpr.Length - dtable.Length - 2)
                Dim kk As Int16 = -1
                If aclsObject IsNot Nothing Then
                    For k = 0 To aclsObject.Count - 1
                        If LCase(dtable) = LCase(aclsObject(k).TableName) Then
                            '  If aclsObject(k).SqlUpdation = True Then
                            kk = k
                            Exit For
                            'End If
                        End If
                    Next
                    If kk > -1 Then
                        Try
                            aObject = aclsObject(kk).CurrDt.Rows(0).Item(dfield)
                        Catch ex As Exception
                            QuitMessage("In valid value for field " & dfield & "  the table  " & dtable, "SetFinalFieldsValues")
                        End Try
                    End If
                End If
            Case LCase(mexpr) = "now()"
                aObject = Now()

            Case Else
                If HashPublicValues IsNot Nothing Then
                    If GF1.GetValueFromHashTable(HashPublicValues, mexpr) Is Nothing Then
                        QuitMessage("Variable " & mexpr & " not assigned in  HashPublicValues  ", "SetFinalFieldsValues")
                    Else
                        aObject = GF1.GetValueFromHashTable(HashPublicValues, mexpr)
                    End If
                End If
        End Select
        Return aObject
    End Function
    Private Function EvalCGroupTypeField(ByVal FieldAssignExpression As String, ByVal aclsObject() As Object, ByVal HashPublicValues As Hashtable) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim jvalue() As String = FieldAssignExpression.Split("#")
        Dim mfield As String = jvalue(0)
        Dim mexpr As String = jvalue(1)
        Dim aVar() As String = GF1.ExtractVariables(mexpr)
        Dim aValueHash As New Hashtable
        For z = 0 To aVar.Count - 1
            Dim mvar As String = aVar(z)
            If InStr(mvar, "__") > 0 Then
                Dim d As Int16 = InStr(mvar, "__")
                Dim dtable As String = Microsoft.VisualBasic.Left(mvar, d - 1)
                Dim dfield As String = Microsoft.VisualBasic.Mid(mvar, d + 2, mvar.Length - dtable.Length - 2)
                Dim kk As Int16 = -1
                If aclsObject IsNot Nothing Then
                    For k = 0 To aclsObject.Count - 1
                        If LCase(dtable) = LCase(aclsObject(k).TableName) Then
                            If aclsObject(k).SqlUpdation = True Then
                                kk = k
                                Exit For
                            End If
                        End If
                    Next
                    If kk > -1 Then
                        Try
                            aValueHash = GF1.AddItemToHashTable(aValueHash, mvar, aclsObject(kk).CurrDt.Rows(0).Item(dfield))
                            '    clsObject(CInt(mkey)).CurrDt(x).Item(mfield) = clsObject(kk).CurrDt(0).Item(dfield)
                        Catch ex As Exception
                            QuitMessage("In valid value of " & mvar & " for the table  " & dtable, "SetFinalFieldsValues")
                        End Try
                    Else
                        aValueHash = GF1.AddItemToHashTable(aValueHash, mvar, 0)
                    End If
                End If
            Else
                If HashPublicValues IsNot Nothing Then
                    If GF1.GetValueFromHashTable(HashPublicValues, mvar) Is Nothing Then
                        QuitMessage("Variable " & mvar & " not assigned in hashtable VariableValues  ", "SetFinalFieldsValues")
                    Else
                        aValueHash = GF1.AddItemToHashTable(aValueHash, mvar, GF1.GetValueFromHashTable(HashPublicValues, mvar))
                    End If
                End If
            End If
        Next
        Dim rexpr As String = GF1.ReplaceValuesInExpression(mexpr, aValueHash, "VB")
        Dim merror As Boolean = False
        Dim mvalue As Object = GF1.EvaluateExpression(rexpr, merror)
        Return mvalue
    End Function






        ''' <summary>
        ''' To get and assign increamental last value in a data table for a given field.
        ''' </summary>
        ''' <param name="LDataTable">Data table from which lastkey found and replaced</param>
        ''' <param name="LastKeyField">Field name of lastkey</param>
        ''' <param name="AssignValueToField" >True  if LastValuePlusOne to be assigned to column values of datatable. </param>
        ''' <param name="OnlyEmpty" >True ,if only empty ,zero or null cells to be replace</param>
        ''' <param name="StartFrom" >Field Value to be start from </param>
        ''' <returns></returns>
        ''' <remarks></remarks>

    Public Function LastValuePlusOne(ByRef LDataTable As DataTable, ByVal LastKeyField As String, Optional ByVal AssignValueToField As Boolean = True, Optional ByVal OnlyEmpty As Boolean = True, Optional ByVal StartFrom As Integer = 0) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mLastKey As Integer = 0
        Try
            If LDataTable.Rows.Count = 0 Then
                Return mLastKey
            End If
            Dim lfilter As String = ""
            Dim FdataTable As New DataTable
            ' LDataTable .

            For i = 0 To LDataTable.Rows.Count - 1
                If IsDBNull(LDataTable.Rows(i).Item(LastKeyField)) = False Then
                    If LDataTable.Rows(i).Item(LastKeyField) > mLastKey Then
                        mLastKey = LDataTable.Rows(i).Item(LastKeyField)
                    End If
                End If
            Next
            If OnlyEmpty = True Then
                mLastKey = mLastKey + 1 + StartFrom
            Else
                mLastKey = StartFrom + 1
            End If
            If AssignValueToField = True Then
                For i = 0 To LDataTable.Rows.Count - 1
                    If OnlyEmpty = True Then
                        If IsDBNull(LDataTable.Rows(i).Item(LastKeyField)) = True Then
                            LDataTable.Rows(i).Item(LastKeyField) = mLastKey
                        End If
                        If LDataTable.Rows(i).Item(LastKeyField) = 0 Then
                            LDataTable.Rows(i).Item(LastKeyField) = mLastKey
                        End If
                    Else
                        LDataTable.Rows(i).Item(LastKeyField) = mLastKey
                    End If
                    mLastKey = mLastKey + 1
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to Execute DataFunction.LastValuePlus(ByVal LDataTable As DataTable, ByVal LastKeyField As String, Optional ByVal AssignKeyValue As Boolean = True, Optional ByVal StartFrom As Integer = 0) As Integer")
        End Try
        Return mLastKey
    End Function



    ''' <summary>
    ''' To get increamental last key from an SQL Table
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format , if space or used table names are full identifier then this will extracted from full table name</param>
    ''' <param name="Ltable" >Sql Table name as string</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="Prefix">Prefix of lastkey</param>
    ''' <param name="LastKeyFieldSize">Size of the keyfield </param>
    ''' <param name="NewRowTemplate" >NewRowTemplate for RowSource used for new insert in sql table </param>
    ''' <param name="LastKeyValues" >A hashtable having LastKeyPlus  values.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LastKeyPlus(ByVal ServerDataBase As String, ByRef Ltable As String, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer, Optional ByRef NewRowTemplate As DataRow = Nothing, Optional ByRef LastKeyValues As Hashtable = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        ServerDataBase = ConvertFromSrv0Mdf0(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim mLastKey As String = Prefix & Right("000000000001", LastKeyFieldSize - Prefix.Length)
        If LastKeyValues IsNot Nothing Then
            Dim mLastno As String = GF1.GetValueFromHashTable(LastKeyValues, LastKeyField)
            If mLastno IsNot Nothing Then
                mLastno = CInt(Right(mLastno, LastKeyFieldSize - Prefix.Length)) + 1
                mLastKey = Prefix & Right(CStr(100000000000 + mLastno), LastKeyFieldSize - Prefix.Length)
                LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                Return mLastKey
                Exit Function
            End If
        End If
        Try
            Dim QuerryStr As String = "Select top(1) " & LastKeyField & " from " & Ltable
            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
            If Prefix.Trim.Length > 0 Then
                '  QuerryStr = QuerryStr & " Where Left(" & LastKeyField & "," & Prefix.Length.ToString & ") =  '" & Prefix & "'"
                QuerryStr = QuerryStr & " where " & LastKeyField & "  like  '" & Prefix & "%'"
                'QuerryStr = QuerryStr & "  " & LastKeyField & "  like  '" & Prefix & "%'"
            End If
            QuerryStr = QuerryStr & " order by " & LastKeyField & " desc"
            Dim dt As DataTable = SqlExecuteDataTable(ServerDataBase, QuerryStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim mlastno As Integer = 0
                    mlastno = CInt(Right(dt.Rows(0)(LastKeyField), LastKeyFieldSize - Prefix.Length)) + 1
                    mLastKey = Prefix & Right(CStr(100000000000 + mlastno), LastKeyFieldSize - Prefix.Length)
                    If LastKeyValues IsNot Nothing Then
                        LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                    End If
                End If
                If NewRowTemplate IsNot Nothing Then
                    NewRowTemplate = dt.NewRow
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByVal ServerDataBase As String, ByRef Ltable As String, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer) As String")

        End Try

        Return mLastKey
    End Function
    ''' <summary>
    ''' To get increamental last key from an SQL Table
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format , if space or used table names are full identifier then this will extracted from full table name</param>
    ''' <param name="Ltable" >Sql Table name as string</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="NewRowTemplate" >NewRowTemplate for RowSource used for new insert in sql table </param>
    ''' <param name="LastKeyValues" >An array of hashtable having LastKeyPlus  values involved in this transaction.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function LastKeyPlus(ByVal ServerDataBase As String, ByRef Ltable As String, ByVal LastKeyField As String, Optional ByRef NewRowTemplate As DataRow = Nothing, Optional ByRef LastKeyValues As Hashtable = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        ServerDataBase = ConvertFromSrv0Mdf0(ServerDataBase)
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim mLastKey As Integer = 1
        If LastKeyValues IsNot Nothing Then
            Dim mLastno As String = GF1.GetValueFromHashTable(LastKeyValues, LastKeyField)
            If mLastno IsNot Nothing Then
                mLastno = mLastno + 1
                mLastKey = mLastno
                LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                Return mLastKey
                Exit Function
            End If
        End If
        Try
            Dim QuerryStr As String = "Select top(1) " & LastKeyField & " from " & Ltable
            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
            QuerryStr = QuerryStr & " order by " & LastKeyField & " desc"
            Dim dt As DataTable = SqlExecuteDataTable(ServerDataBase, QuerryStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    mLastKey = CInt(dt.Rows(0)(LastKeyField)) + 1
                    If LastKeyValues IsNot Nothing Then
                        LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                    End If
                End If
                If NewRowTemplate IsNot Nothing Then
                    NewRowTemplate = dt.NewRow
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByVal ServerDataBase As String, ByRef Ltable As String, ByVal LastKeyField As String) As Integer")
        End Try
        Return mLastKey
    End Function
    ''' <summary>
    ''' To get increamental last key from an SQL Table
    ''' </summary>
    ''' <param name="Ltrans" >Sql Transaction</param>
    ''' <param name="Ltable" >Sql Table name as string</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="Prefix">Prefix of lastkey</param>
    ''' <param name="LastKeyFieldSize">Size of the keyfield </param>
    ''' <param name="NewRowTemplate" >NewRowTemplate for RowSource used for new insert in sql table </param>
    ''' <param name="LastKeyPlusInTransaction" >An array of hashtable having LastKeyPlus  values involved in this transaction.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function LastKeyPlus(ByRef Ltrans As SqlTransaction, ByVal Ltable As String, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer, Optional ByRef NewRowTemplate As DataRow = Nothing, Optional ByRef LastKeyPlusInTransaction As Hashtable = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim mLastKey As String = Prefix & Right("000000000001", LastKeyFieldSize - Prefix.Length)
        If LastKeyPlusInTransaction IsNot Nothing Then
            Dim LastKeyValues As Hashtable = GF1.GetValueFromHashTable(LastKeyPlusInTransaction, Ltable)
            Dim mLastno As String = GF1.GetValueFromHashTable(LastKeyValues, LastKeyField)
            If mLastno IsNot Nothing Then
                mLastno = CInt(Right(mLastno, LastKeyFieldSize - Prefix.Length)) + 1
                mLastKey = Prefix & Right(CStr(100000000000 + mLastno), LastKeyFieldSize - Prefix.Length)
                LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                Return mLastKey
                Exit Function
            End If
        End If
        Try
            Dim QuerryStr As String = "Select top(1) " & LastKeyField & " from " & Ltable
            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
            If Prefix.Trim.Length > 0 Then
                '  QuerryStr = QuerryStr & " Where Left(" & LastKeyField & "," & Prefix.Length.ToString & ") =  '" & Prefix & "'"
                QuerryStr = QuerryStr & " where " & LastKeyField & "  like  '" & Prefix & "%'"
                'QuerryStr = QuerryStr & "  " & LastKeyField & "  like  '" & Prefix & "%'"
            End If
            QuerryStr = QuerryStr & " order by " & LastKeyField & " desc"
            Dim dt As DataTable = SqlExecuteDataTable(Ltrans, QuerryStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim mlastno As Integer = 0
                    mlastno = CInt(Right(dt.Rows(0)(LastKeyField), LastKeyFieldSize - Prefix.Length)) + 1
                    mLastKey = Prefix & Right(CStr(100000000000 + mlastno), LastKeyFieldSize - Prefix.Length)
                    If LastKeyPlusInTransaction IsNot Nothing Then
                        Dim LastKeyValues As New Hashtable
                        ' LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, "Table", Ltable)
                        LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                        LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                    End If
                End If
                If NewRowTemplate IsNot Nothing Then
                    NewRowTemplate = dt.NewRow
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByRef Ltrans As SqlTransaction, ByVal Ltable As String, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer) As String")
        End Try
        Return mLastKey
    End Function
    ''' <summary>
    ''' To get increamental last key from an SQL Table
    ''' </summary>
    ''' <param name="LConnection" >Sql Transaction</param>
    ''' <param name="Ltable" >Sql Table name as string</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="Prefix">Prefix of lastkey</param>
    ''' <param name="LastKeyFieldSize">Size of the keyfield </param>
    ''' <param name="NewRowTemplate" >NewRowTemplate for RowSource used for new insert in sql table </param>
    ''' <param name="LastKeyPlusInTransaction" >An array of hashtable having LastKeyPlus  values involved in this transaction.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function LastKeyPlus(ByRef LConnection As SqlConnection, ByVal Ltable As String, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer, Optional ByRef NewRowTemplate As DataRow = Nothing, Optional ByRef LastKeyPlusInTransaction As Hashtable = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim mLastKey As String = Prefix & Right("000000000001", LastKeyFieldSize - Prefix.Length)
        If LastKeyPlusInTransaction IsNot Nothing Then
            Dim LastKeyValues As Hashtable = GF1.GetValueFromHashTable(LastKeyPlusInTransaction, Ltable)
            Dim mLastno As String = GF1.GetValueFromHashTable(LastKeyValues, LastKeyField)
            If mLastno IsNot Nothing Then
                mLastno = CInt(Right(mLastno, LastKeyFieldSize - Prefix.Length)) + 1
                mLastKey = Prefix & Right(CStr(100000000000 + mLastno), LastKeyFieldSize - Prefix.Length)
                LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                Return mLastKey
                Exit Function
            End If
        End If

        Try
            Dim QuerryStr As String = "Select top(1) " & LastKeyField & " from " & Ltable
            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
            If Prefix.Trim.Length > 0 Then
                '  QuerryStr = QuerryStr & " Where Left(" & LastKeyField & "," & Prefix.Length.ToString & ") =  '" & Prefix & "'"
                QuerryStr = QuerryStr & " where " & LastKeyField & "  like  '" & Prefix & "%'"
                'QuerryStr = QuerryStr & "  " & LastKeyField & "  like  '" & Prefix & "%'"
            End If
            QuerryStr = QuerryStr & " order by " & LastKeyField & " desc"
            Dim dt As DataTable = SqlExecuteDataTable(LConnection, QuerryStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim mlastno As Integer = 0
                    mlastno = CInt(Right(dt.Rows(0)(LastKeyField), LastKeyFieldSize - Prefix.Length)) + 1
                    mLastKey = Prefix & Right(CStr(100000000000 + mlastno), LastKeyFieldSize - Prefix.Length)
                    If LastKeyPlusInTransaction IsNot Nothing Then
                        Dim LastKeyValues As New Hashtable
                        ' LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, "Table", Ltable)
                        LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                        LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                    End If
                End If
                If NewRowTemplate IsNot Nothing Then
                    NewRowTemplate = dt.NewRow
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByRef LConnection As SqlConnection, ByVal Ltable As String, ByVal LastKeyField As String, ByVal Prefix As String, ByVal LastKeyFieldSize As Integer) As String")
        End Try
        Return mLastKey
    End Function
    ''' <summary>
    ''' To get increamental last key from an SQL Table
    ''' </summary>
    ''' <param name="Ltrans" >Sql Transaction</param>
    ''' <param name="Ltable" >Sql Table name as string</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="LCondition" >Optional condition for lastkeyplus query</param>
    ''' <param name="NewRowTemplate" >NewRowTemplate for RowSource used for new insert in sql table </param>
    ''' <param name="LastKeyPlusInTransaction" >An array of hashtable having LastKeyPlus  values involved in this transaction.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LastKeyPlus(ByRef Ltrans As SqlTransaction, ByVal Ltable As String, ByVal LastKeyField As String, Optional ByVal LCondition As String = "", Optional ByRef NewRowTemplate As DataRow = Nothing, Optional ByRef LastKeyPlusInTransaction As Hashtable = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim mLastKey As Integer = 1
        If LastKeyPlusInTransaction IsNot Nothing Then
            Dim LastKeyValues As Hashtable = GF1.GetValueFromHashTable(LastKeyPlusInTransaction, "Table", Ltable)
            Dim mLastno As String = GF1.GetValueFromHashTable(LastKeyValues, LastKeyField)
            If mLastno IsNot Nothing Then
                mLastno = mLastno + 1
                mLastKey = mLastno
                LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                Return mLastKey
                Exit Function
            End If
        End If
        Try
            Dim mFieldList As String = IIf(NewRowTemplate IsNot Nothing, " * ", LastKeyField)
            Dim QuerryStr As String = "Select top(1) " & mFieldList & " from " & Ltable
            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
            QuerryStr = QuerryStr & IIf(LCondition.Trim.Length > 0, " where " & LCondition, "")
            QuerryStr = QuerryStr & " order by " & LastKeyField & " desc"
            Dim dt As DataTable = SqlExecuteDataTable(Ltrans, QuerryStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    mLastKey = CInt(dt.Rows(0)(LastKeyField)) + 1
                    If LastKeyPlusInTransaction IsNot Nothing Then
                        Dim LastKeyValues As New Hashtable
                        LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                        LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                    End If
                End If
                If NewRowTemplate IsNot Nothing Then
                    NewRowTemplate = dt.NewRow
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByRef Ltrans As SqlTransaction, ByVal Ltable As String, ByVal LastKeyField As String, Optional ByVal LCondition As String = "") As Integer")
        End Try
        Return mLastKey
    End Function
    ''' <summary>
    ''' To get increamental last key from an SQL Table
    ''' </summary>
    ''' <param name="LConnection" >SqlConnection</param>
    ''' <param name="Ltable" >Sql Table name as string</param>
    ''' <param name="LastKeyField">Field name of lastkey</param>
    ''' <param name="LCondition" >Optional condition for lastkeyplus query</param>
    ''' <param name="NewRowTemplate" >NewRowTemplate for RowSource used for new insert in sql table </param>
    ''' <param name="LastKeyPlusInTransaction" >An array of hashtable having LastKeyPlus  values involved in this transaction.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LastKeyPlus(ByRef LConnection As SqlConnection, ByVal Ltable As String, ByVal LastKeyField As String, Optional ByVal LCondition As String = "", Optional ByRef NewRowTemplate As DataRow = Nothing, Optional ByRef LastKeyPlusInTransaction As Hashtable = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Ltable = ConvertFromSrv0Mdf0(Ltable, True)
        Dim mLastKey As Integer = 1
        If LastKeyPlusInTransaction IsNot Nothing Then
            Dim LastKeyValues As Hashtable = GF1.GetValueFromHashTable(LastKeyPlusInTransaction, "Table", Ltable)
            Dim mLastno As String = GF1.GetValueFromHashTable(LastKeyValues, LastKeyField)
            If mLastno IsNot Nothing Then
                mLastno = mLastno + 1
                mLastKey = mLastno
                LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                Return mLastKey
                Exit Function
            End If
        End If
        Try
            Dim QuerryStr As String = "Select top(1) " & LastKeyField & " from " & Ltable
            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
            QuerryStr = QuerryStr & IIf(LCondition.Trim.Length > 0, " where " & LCondition, "")
            QuerryStr = QuerryStr & " order by " & LastKeyField & " desc"
            Dim dt As DataTable = SqlExecuteDataTable(LCondition, QuerryStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    mLastKey = CInt(dt.Rows(0)(LastKeyField)) + 1
                    If LastKeyPlusInTransaction IsNot Nothing Then
                        Dim LastKeyValues As New Hashtable
                        LastKeyValues = GF1.AddItemToHashTable(LastKeyValues, LastKeyField, mLastKey)
                        LastKeyPlusInTransaction = GF1.AddItemToHashTable(LastKeyPlusInTransaction, Ltable, LastKeyValues)
                    End If
                End If
                If NewRowTemplate IsNot Nothing Then
                    NewRowTemplate = dt.NewRow
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.LastKeyPlus(ByRef LConnection As SqlConnection, ByVal Ltable As String, ByVal LastKeyField As String, Optional ByVal LCondition As String = "") As Integer")
        End Try
        Return mLastKey
    End Function
    ''' <summary>
    ''' To get increamental last key from  SQL Tables ,If a constraints are given in a collection
    ''' </summary>
    ''' <param name="Ltrans">Sql Transaction object</param>
    ''' <param name="ClsTable" >A table class object</param>
    ''' <param name="KeyPlusGroups" >Comma separated type of field group from grouptype Y,R,S,O,D default is Y</param>
    ''' <param name="LastKeysValues" >A Hash Table with keys are tablename and values are another hash table (keys= Increamenting fields(Including primary keys),Values= LastValue + 1 of sql table,</param>
    ''' <returns>A  ClsTable object with updated LastKeyValues</returns>
    ''' <remarks></remarks>
    Public Function LastKeysPlus(ByRef Ltrans As SqlTransaction, ByVal ClsTable As Object, Optional ByVal KeyPlusGroups As String = "Y", Optional ByRef LastKeysValues As Hashtable = Nothing) As Object

        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aCollection() As Collection = GF1.GetCollectionFromCollection(ClsTable.FieldsPlusCollection, "KeyPlusGroup", KeyPlusGroups)
        If aCollection.Count = 0 Then
            Return ClsTable
            Exit Function
        End If


        Dim mLastKeys As New Hashtable
        If LastKeysValues IsNot Nothing Then
            If GF1.GetValueFromHashTable(LastKeysValues, ClsTable.TableName) IsNot Nothing Then
                mLastKeys = GF1.GetValueFromHashTable(LastKeysValues, ClsTable.TableName)
            End If
        End If
        Select Case True
            Case mLastKeys.Count > 0
                For i = 0 To aCollection.Count - 1
                    Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                    Dim mLastno As Integer = 0
                    If GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield) IsNot Nothing Then
                        mLastno = CInt(GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield))
                    End If
                    mLastno = mLastno + 1
                    mLastKeys = GF1.AddItemToHashTable(mLastKeys, Lastkeyfield, mLastno)
                    If LCase(Lastkeyfield) = LCase(ClsTable.PrimaryKey) Then
                        ClsTable.PrimaryKeyValue = mLastno
                    End If
                Next
                ClsTable.FieldsFinalValues = mLastKeys
                If LastKeysValues IsNot Nothing Then
                    LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, ClsTable.TableName, mLastKeys)
                End If
            Case ClsTable.FieldsFinalValues.count > 0
                For i = 0 To aCollection.Count - 1
                    Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                    Dim mLastno As Integer = 0
                    If GF1.GetValueFromHashTable(ClsTable.FieldsFinalValues, Lastkeyfield) IsNot Nothing Then
                        mLastno = CInt(GF1.GetValueFromHashTable(ClsTable.FieldsFinalValues, Lastkeyfield))
                    End If
                    mLastno = mLastno + 1
                    ClsTable.FieldsFinalValues = GF1.AddItemToHashTable(ClsTable.FieldsFinalValues, Lastkeyfield, mLastno)
                    If LCase(Lastkeyfield) = LCase(ClsTable.PrimaryKey) Then
                        ClsTable.PrimaryKeyValue = mLastno
                    End If
                Next
                If LastKeysValues IsNot Nothing Then
                    LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, ClsTable.TableName, ClsTable.FieldsFinalValues)
                End If
            Case Else
                Dim QuerryStr As String = ""
                Try
                    For i = 0 To aCollection.Count - 1
                        Dim mCollection As Collection = aCollection(i)
                        Dim mTable As String = ConvertFromSrv0Mdf0(GF1.GetValueFromCollection(mCollection, "Table"), True)
                        Dim mLastkeyfield As String = GF1.GetValueFromCollection(mCollection, "LastKeyField")
                        Dim mCondition As String = GF1.GetValueFromCollection(mCollection, "Condition")
                        Dim mConditionVars As Hashtable = GF1.GetValueFromCollection(mCollection, "Variables")
                        mCondition = GF1.ReplaceValuesInExpression(mCondition, mConditionVars)
                        QuerryStr = QuerryStr & "Select top(1) " & mLastkeyfield & " from " & mTable
                        QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
                        QuerryStr = QuerryStr & IIf(mCondition.Trim.Length > 0, " where " & mCondition, "")
                        QuerryStr = QuerryStr & " order by " & mLastkeyfield & " desc" & vbCrLf
                    Next
                    Dim ds As DataSet = SqlExecuteDataSet(Ltrans, QuerryStr)
                    If ds IsNot Nothing Then
                        For i = 0 To aCollection.Count - 1
                            Dim mLastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                            Dim dt As DataTable = ds.Tables(i)
                            Dim mlastkey As Integer = 1
                            If dt.Rows.Count > 0 Then
                                mlastkey = CInt(dt.Rows(0)(mLastkeyfield)) + 1
                            End If
                            GF1.AddItemToHashTable(mLastKeys, mLastkeyfield, mlastkey)
                            If LCase(mLastkeyfield) = LCase(ClsTable.PrimaryKey) Then
                                ClsTable.PrimaryKeyValue = mlastkey
                            End If
                        Next
                        ClsTable.FieldsFinalValues = mLastKeys
                        If LastKeysValues IsNot Nothing Then
                            LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, ClsTable.TableName, mLastKeys)
                        End If
                    End If
                Catch ex As Exception
                    QuitError(ex, Err, "Unable to execute datafunction.LastKeysPlus(ByRef Ltrans As SqlTransaction, ByVal ClsTable As Object) As Hashtable")
                End Try
        End Select
        If ClsTable.TableEntryType = "A" And ClsTable.RowStatusFlag = True Then
            ClsTable.FieldsFinalValues = GF1.AddItemToHashTable(ClsTable.FieldsFinalValues, "RowStatus", 0)
        End If
        Return ClsTable

    End Function
    ''' <summary>
    ''' To get increamental last key from  SQL Tables ,If a constraints are given in a collection
    ''' </summary>
    ''' <param name="LConnection">SqlConnection object</param>
    ''' <param name="ClsTable" >A table class object</param>
    ''' <param name="KeyPlusGroups" >Comma separated type of field group from grouptype Y,R,S,O,D default is Y</param>
    ''' <param name="LastKeysValues" >A Hash Table with keys are tablename and values are another hash table (keys= Increamenting fields(Including primary keys),Values= LastValue + 1 of sql table,</param>
    ''' <returns>A  ClsTable object with updated LastKeyValues</returns>
    ''' <remarks></remarks>
    Public Function LastKeysPlus(ByRef LConnection As SqlConnection, ByVal ClsTable As Object, Optional ByVal KeyPlusGroups As String = "Y", Optional ByRef LastKeysValues As Hashtable = Nothing) As Object

        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aCollection() As Collection = GF1.GetCollectionFromCollection(ClsTable.FieldsPlusCollection, "KeyPlusGroup", KeyPlusGroups)
        If aCollection.Count = 0 Then
            Return ClsTable
            Exit Function
        End If


        Dim mLastKeys As New Hashtable
        If LastKeysValues IsNot Nothing Then
            If GF1.GetValueFromHashTable(LastKeysValues, ClsTable.TableName) IsNot Nothing Then
                mLastKeys = GF1.GetValueFromHashTable(LastKeysValues, ClsTable.TableName)
            End If
        End If
        Select Case True
            Case mLastKeys.Count > 0
                For i = 0 To aCollection.Count - 1
                    Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                    Dim mLastno As Integer = 0
                    If GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield) IsNot Nothing Then
                        mLastno = CInt(GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield))
                    End If
                    mLastno = mLastno + 1
                    mLastKeys = GF1.AddItemToHashTable(mLastKeys, Lastkeyfield, mLastno)
                    If LCase(Lastkeyfield) = LCase(ClsTable.PrimaryKey) Then
                        ClsTable.PrimaryKeyValue = mLastno
                    End If
                Next
                ClsTable.FieldsFinalValues = mLastKeys
                If LastKeysValues IsNot Nothing Then
                    LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, ClsTable.TableName, mLastKeys)
                End If
            Case ClsTable.FieldsFinalValues.count > 0
                For i = 0 To aCollection.Count - 1
                    Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                    Dim mLastno As Integer = 0
                    If GF1.GetValueFromHashTable(ClsTable.FieldsFinalValues, Lastkeyfield) IsNot Nothing Then
                        mLastno = CInt(GF1.GetValueFromHashTable(ClsTable.FieldsFinalValues, Lastkeyfield))
                    End If
                    mLastno = mLastno + 1
                    ClsTable.FieldsFinalValues = GF1.AddItemToHashTable(ClsTable.FieldsFinalValues, Lastkeyfield, mLastno)
                    If LCase(Lastkeyfield) = LCase(ClsTable.PrimaryKey) Then
                        ClsTable.PrimaryKeyValue = mLastno
                    End If
                Next
                If LastKeysValues IsNot Nothing Then
                    LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, ClsTable.TableName, ClsTable.FieldsFinalValues)
                End If
            Case Else
                Dim QuerryStr As String = ""
                Try
                    For i = 0 To aCollection.Count - 1
                        Dim mCollection As Collection = aCollection(i)
                        Dim mTable As String = ConvertFromSrv0Mdf0(GF1.GetValueFromCollection(mCollection, "Table"), True)
                        Dim mLastkeyfield As String = GF1.GetValueFromCollection(mCollection, "LastKeyField")
                        Dim mCondition As String = GF1.GetValueFromCollection(mCollection, "Condition")
                        Dim mConditionVars As Hashtable = GF1.GetValueFromCollection(mCollection, "Variables")
                        mCondition = GF1.ReplaceValuesInExpression(mCondition, mConditionVars)
                        QuerryStr = QuerryStr & "Select top(1) " & mLastkeyfield & " from " & mTable
                        QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
                        QuerryStr = QuerryStr & IIf(mCondition.Trim.Length > 0, " where " & mCondition, "")
                        QuerryStr = QuerryStr & " order by " & mLastkeyfield & " desc" & vbCrLf
                    Next
                    Dim ds As DataSet = SqlExecuteDataSet(LConnection, QuerryStr)
                    If ds IsNot Nothing Then
                        For i = 0 To aCollection.Count - 1
                            Dim mLastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                            Dim dt As DataTable = ds.Tables(i)
                            Dim mlastkey As Integer = 1
                            If dt.Rows.Count > 0 Then
                                mlastkey = CInt(dt.Rows(0)(mLastkeyfield)) + 1
                            End If
                            GF1.AddItemToHashTable(mLastKeys, mLastkeyfield, mlastkey)
                            If LCase(mLastkeyfield) = LCase(ClsTable.PrimaryKey) Then
                                ClsTable.PrimaryKeyValue = mlastkey
                            End If
                        Next
                        ClsTable.FieldsFinalValues = mLastKeys
                        If LastKeysValues IsNot Nothing Then
                            LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, ClsTable.TableName, mLastKeys)
                        End If
                    End If
                Catch ex As Exception
                    QuitError(ex, Err, "Unable to execute datafunction.LastKeysPlus(ByRef LConnection As SqlConnection, ByVal ClsTable As Object) As Hashtable")
                End Try
        End Select
        If ClsTable.TableEntryType = "A" And ClsTable.RowStatusFlag = True Then
            ClsTable.FieldsFinalValues = GF1.AddItemToHashTable(ClsTable.FieldsFinalValues, "RowStatus", 0)
        End If
        Return ClsTable

    End Function
    ''' <summary>
    ''' To get increamental last key from  SQL Tables ,If a constraints are given in a collection
    ''' </summary>
    ''' <param name="Ltrans">Sql Transaction object</param>
    ''' <param name="aClsTable" >An array of  table class object</param>
    ''' <param name="LastKeysValues" >A Hash Table with keys are tablename and values are another hash table (keys= Increamenting fields(Including primary keys),Values= LastValue + 1 of sql table,</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LastKeysPlus(ByRef Ltrans As SqlTransaction, ByVal aClsTable() As Object, Optional ByRef LastKeysValues As Hashtable = Nothing) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Dim cls As New ColorScheme.ColorScheme
        Dim aLastKeyFields() As String = {}
        Dim QuerryStr As String = ""
        For j = 0 To aClsTable.Count - 1
            Try
                If aClsTable(j).SqlUpdation = False Then
                    Continue For
                End If
                Dim aCollection() As Collection = GF1.GetCollectionFromCollection(aClsTable(j).FieldsPlusCollection, "KeyPlusGroup", aClsTable(j).KeyPlusGroups)
                If aCollection.Count = 0 Then
                    Continue For
                End If
                Dim mLastKeys As New Hashtable
                If LastKeysValues.count > 0 Then  'changed by Neha
                    If GF1.GetValueFromHashTable(LastKeysValues, aClsTable(j).TableName) IsNot Nothing Then
                        mLastKeys = GF1.GetValueFromHashTable(LastKeysValues, aClsTable(j).TableName)
                    End If
                End If
                Select Case True
                    Case mLastKeys.Count > 0
                        For i = 0 To aCollection.Count - 1
                            Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                            Dim mLastno As Integer = 0
                            If GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield) IsNot Nothing Then
                                mLastno = CInt(GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield))
                            End If
                            mLastno = mLastno + 1
                            mLastKeys = GF1.AddItemToHashTable(mLastKeys, Lastkeyfield, mLastno)
                            If LCase(Lastkeyfield) = LCase(aClsTable(j).PrimaryKey) Then
                                aClsTable(j).PrimaryKeyValue = mLastno
                            End If
                        Next
                        aClsTable(j).FieldsFinalValues = mLastKeys
                        If LastKeysValues IsNot Nothing Then
                            LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, aClsTable(j).TableName, mLastKeys)
                        End If
                    Case aClsTable(j).FieldsFinalValues.count > 0
                        For i = 0 To aCollection.Count - 1
                            Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                            Dim mLastno As Integer = 0
                            If GF1.GetValueFromHashTable(aClsTable(j).FieldsFinalValues, Lastkeyfield) IsNot Nothing Then
                                mLastno = CInt(GF1.GetValueFromHashTable(aClsTable(j).FieldsFinalValues, Lastkeyfield))
                            End If
                            mLastno = mLastno + 1
                            aClsTable(j).FieldsFinalValues = GF1.AddItemToHashTable(aClsTable(j).FieldsFinalValues, Lastkeyfield, mLastno)
                            If LCase(Lastkeyfield) = LCase(aClsTable(j).PrimaryKey) Then
                                aClsTable(j).PrimaryKeyValue = mLastno
                            End If
                        Next
                        If LastKeysValues IsNot Nothing Then
                            LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, aClsTable(j).TableName, aClsTable(j).FieldsFinalValues)
                        End If
                    Case Else
                        For i = 0 To aCollection.Count - 1
                            Dim mCollection As Collection = aCollection(i)
                            Dim mTable As String = ConvertFromSrv0Mdf0(GF1.GetValueFromCollection(mCollection, "Table"), True)
                            Dim mLastkeyfield As String = GF1.GetValueFromCollection(mCollection, "LastKeyField")
                            Dim mCondition As String = GF1.GetValueFromCollection(mCollection, "Condition")
                            Dim mConditionVars As Hashtable = GF1.GetValueFromCollection(mCollection, "Variables")
                            mCondition = GF1.ReplaceValuesInExpression(mCondition, mConditionVars)
                            QuerryStr = QuerryStr & "Select top(1) " & mLastkeyfield & " from " & mTable
                            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
                            QuerryStr = QuerryStr & IIf(mCondition.Trim.Length > 0, " where " & mCondition, "")
                            QuerryStr = QuerryStr & " order by " & mLastkeyfield & " desc" & vbCrLf
                            aLastKeyFields = GF1.ArrayAppend(aLastKeyFields, mTable & "," & mLastkeyfield)
                        Next
                End Select
                If aClsTable(j).TableEntryType = "A" And aClsTable(j).RowStatusFlag = True Then
                    aClsTable(j).FieldsFinalValues = GF1.AddItemToHashTable(aClsTable(j).FieldsFinalValues, "RowStatus", 0)
                End If
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute " & aClsTable(j).TableName & " datafunction.Function LastKeysPlus(ByRef Ltrans As SqlTransaction, ByVal LastKeyConstraints() As Collection) As Collection()")
            End Try
        Next
        If QuerryStr.Length > 0 Then
            Try
                Dim ds As DataSet = SqlExecuteDataSet(Ltrans, QuerryStr)
                If ds IsNot Nothing Then
                    For i = 0 To aLastKeyFields.Count - 1
                        Dim aLastFields() As String = aLastKeyFields(i).Trim.Split(",")
                        Dim mtable As String = aLastFields(0)
                        Dim mLastkeyfield As String = aLastFields(1)
                        Dim dt As DataTable = ds.Tables(i)
                        Dim mlastkey As Integer = 1
                        If dt.Rows.Count > 0 Then
                            mlastkey = CInt(dt.Rows(0)(mLastkeyfield)) + 1
                        End If

                        For j = 0 To aClsTable.Count - 1
                            Dim mclsTable As Object = aClsTable(j)
                            If LCase(mclsTable.TableName) = LCase(mtable) Or LCase(mclsTable.TableWithSQLPath) = LCase(mtable) Then
                                aClsTable(j).FieldsFinalValues = GF1.AddItemToHashTable(aClsTable(j).FieldsFinalValues, mLastkeyfield, mlastkey)
                                If LCase(mLastkeyfield) = LCase(aClsTable(j).PrimaryKey) Then
                                    aClsTable(j).PrimaryKeyValue = mlastkey
                                End If

                            End If
                        Next
                        If LastKeysValues IsNot Nothing Then
                            For j = 0 To aClsTable.Count - 1
                                LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, aClsTable(j).TableName, aClsTable(j).FieldsFinalValues)
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute " & QuerryStr & " datafunction.Function LastKeysPlus(ByRef Ltrans As SqlTransaction, ByVal LastKeyConstraints() As Collection) As Collection()")
            End Try
        End If
        Return aClsTable

    End Function
    ''' <summary>
    ''' To get increamental last key from  SQL Tables ,If a constraints are given in a collection
    ''' </summary>
    ''' <param name="LConnection">SqlConnection object</param>
    ''' <param name="aClsTable" >An array of  table class object</param>
    ''' <param name="LastKeysValues" >A Hash Table with keys are tablename and values are another hash table (keys= Increamenting fields(Including primary keys),Values= LastValue + 1 of sql table,</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LastKeysPlus(ByRef LConnection As SqlConnection, ByVal aClsTable() As Object, Optional ByRef LastKeysValues As Hashtable = Nothing) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Dim cls As New ColorScheme.ColorScheme
        Dim aLastKeyFields() As String = {}
        Dim QuerryStr As String = ""
        For j = 0 To aClsTable.Count - 1
            Try
                If aClsTable(j).SqlUpdation = False Then
                    Continue For
                End If
                Dim aCollection() As Collection = GF1.GetCollectionFromCollection(aClsTable(j).FieldsPlusCollection, "KeyPlusGroup", aClsTable(j).KeyPlusGroups)
                If aCollection.Count = 0 Then
                    Continue For
                End If
                Dim mLastKeys As New Hashtable
                If LastKeysValues IsNot Nothing Then
                    If GF1.GetValueFromHashTable(LastKeysValues, aClsTable(j).TableName) IsNot Nothing Then
                        mLastKeys = GF1.GetValueFromHashTable(LastKeysValues, aClsTable(j).TableName)
                    End If
                End If
                Select Case True
                    Case mLastKeys.Count > 0
                        For i = 0 To aCollection.Count - 1
                            Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                            Dim mLastno As Integer = 0
                            If GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield) IsNot Nothing Then
                                mLastno = CInt(GF1.GetValueFromHashTable(mLastKeys, Lastkeyfield))
                            End If
                            mLastno = mLastno + 1
                            mLastKeys = GF1.AddItemToHashTable(mLastKeys, Lastkeyfield, mLastno)
                            If LCase(Lastkeyfield) = LCase(aClsTable(j).PrimaryKey) Then
                                aClsTable(j).PrimaryKeyValue = mLastno
                            End If
                        Next
                        aClsTable(j).FieldsFinalValues = mLastKeys
                        If LastKeysValues IsNot Nothing Then
                            LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, aClsTable(j).TableName, mLastKeys)
                        End If
                    Case aClsTable(j).FieldsFinalValues.count > 0
                        For i = 0 To aCollection.Count - 1
                            Dim Lastkeyfield As String = GF1.GetValueFromCollection(aCollection(i), "LastKeyField")
                            Dim mLastno As Integer = 0
                            If GF1.GetValueFromHashTable(aClsTable(j).FieldsFinalValues, Lastkeyfield) IsNot Nothing Then
                                mLastno = CInt(GF1.GetValueFromHashTable(aClsTable(j).FieldsFinalValues, Lastkeyfield))
                            End If
                            mLastno = mLastno + 1
                            aClsTable(j).FieldsFinalValues = GF1.AddItemToHashTable(aClsTable(j).FieldsFinalValues, Lastkeyfield, mLastno)
                            If LCase(Lastkeyfield) = LCase(aClsTable(j).PrimaryKey) Then
                                aClsTable(j).PrimaryKeyValue = mLastno
                            End If
                        Next
                        If LastKeysValues IsNot Nothing Then
                            LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, aClsTable(j).TableName, aClsTable(j).FieldsFinalValues)
                        End If
                    Case Else
                        For i = 0 To aCollection.Count - 1
                            Dim mCollection As Collection = aCollection(i)
                            Dim mTable As String = ConvertFromSrv0Mdf0(GF1.GetValueFromCollection(mCollection, "Table"), True)
                            Dim mLastkeyfield As String = GF1.GetValueFromCollection(mCollection, "LastKeyField")
                            Dim mCondition As String = GF1.GetValueFromCollection(mCollection, "Condition")
                            Dim mConditionVars As Hashtable = GF1.GetValueFromCollection(mCollection, "Variables")
                            mCondition = GF1.ReplaceValuesInExpression(mCondition, mConditionVars)
                            QuerryStr = QuerryStr & "Select top(1) " & mLastkeyfield & " from " & mTable
                            QuerryStr = QuerryStr & " with (holdlock,tablock,updlock) "
                            QuerryStr = QuerryStr & IIf(mCondition.Trim.Length > 0, " where " & mCondition, "")
                            QuerryStr = QuerryStr & " order by " & mLastkeyfield & " desc" & vbCrLf
                            aLastKeyFields = GF1.ArrayAppend(aLastKeyFields, mTable & "," & mLastkeyfield)
                        Next
                End Select
                If aClsTable(j).TableEntryType = "A" And aClsTable(j).RowStatusFlag = True Then
                    aClsTable(j).FieldsFinalValues = GF1.AddItemToHashTable(aClsTable(j).FieldsFinalValues, "RowStatus", 0)
                End If
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute " & aClsTable(j).TableName & " datafunction.Function LastKeysPlus(ByRef Ltrans As SqlTransaction, ByVal LastKeyConstraints() As Collection) As Collection()")
            End Try
        Next
        If QuerryStr.Length > 0 Then
            Try
                Dim ds As DataSet = SqlExecuteDataSet(LConnection, QuerryStr)
                If ds IsNot Nothing Then
                    For i = 0 To aLastKeyFields.Count - 1
                        Dim aLastFields() As String = aLastKeyFields(i).Trim.Split(",")
                        Dim mtable As String = aLastFields(0)
                        Dim mLastkeyfield As String = aLastFields(1)
                        Dim dt As DataTable = ds.Tables(i)
                        Dim mlastkey As Integer = 1
                        If dt.Rows.Count > 0 Then
                            mlastkey = CInt(dt.Rows(0)(mLastkeyfield)) + 1
                        End If

                        For j = 0 To aClsTable.Count - 1
                            Dim mclsTable As Object = aClsTable(j)
                            If LCase(mclsTable.TableName) = LCase(mtable) Then
                                aClsTable(j).FieldsFinalValues = GF1.AddItemToHashTable(aClsTable(j).FieldsFinalValues, mLastkeyfield, mlastkey)
                                If LCase(mLastkeyfield) = LCase(aClsTable(j).PrimaryKey) Then
                                    aClsTable(j).PrimaryKeyValue = mlastkey
                                End If

                            End If
                        Next
                        If LastKeysValues IsNot Nothing Then
                            For j = 0 To aClsTable.Count - 1
                                LastKeysValues = GF1.AddItemToHashTable(LastKeysValues, aClsTable(j).TableName, aClsTable(j).FieldsFinalValues)
                            Next
                        End If
                    Next
                End If
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute " & QuerryStr & " datafunction.Function LastKeysPlus(ByRef LConnection As SqlConnection, ByVal LastKeyConstraints() As Collection) As Collection()")
            End Try
        End If
        Return aClsTable
    End Function



    ''' <summary>
    ''' To create SQL querry/non querry string for sql command execution 
    ''' </summary>
    ''' <param name="LTableName">Main Table name</param>
    ''' <param name="LastKeyCollection">A collection having fields informations of primary key and additional keyno.{col0=FieldName_PrimaryKey, col1=PrimaryFieldType, col2=PrimaryFieldSize, col3=PrimaryfieldPrefix, col4=FieldName_VoucherType, col5=FieldName_VoucherNo, col6=FieldName_Date, col7=IncreamentType(A=Always,F=Financial year wise,Y=Yearwise,M=Month wise,D=DateWise), col8=StringOfVoucherTypes Separated by ! for common voucher no increament}</param>
    ''' <param name="AddLast">TRUE, if row of lastkey is added onto the table ,otherwise take False if lastkey is only computed</param>
    ''' <param name="LPrimaryKey">True if primary key is calculated,False if additional key calculated</param>
    ''' <param name="CurrentDate">Current Date on which additional key generated</param>
    ''' <param name="LVoucherType">Current Voucher Type on which additional key generated</param>
    ''' <returns>String of SQL querry</returns>
    ''' <remarks></remarks>

    Private Function SqlStringLastKey(ByVal LTableName As String, ByVal LastKeyCollection As Collection, ByVal AddLast As Boolean, Optional ByVal LPrimaryKey As Boolean = True, Optional ByVal CurrentDate As String = "", Optional ByVal LVoucherType As String = "") As String
        If GlobalControl.Variables.AuthenticationChecked = False Then Return Nothing : Exit Function
        'collection  having columns of lastkeyplus
        'FieldName_PrimaryKey,col0
        'PrimaryFieldType, col1
        'PrimaryFieldSize, 'col2
        'PrimaryfieldPrefix 'col3

        'FieldName_VoucherType,col4
        ' FieldName_VoucherNo,col5
        ' FieldName_Date,col6
        'IncreamentType, col7
        'VoucherTypeValueString( ! separated),col8
        'VoucherTypeValueString( ! separated),col8 (for common increament)
        'VouchTypeValue for this key col9
        LTableName = ConvertFromSrv0Mdf0(LTableName)
        Dim QuerryStr As String = ""
        Try
            Dim lFinStartDate As String = GlobalControl.Variables.BusinessFirmRow.Item("YearStartDate")
            Dim lFinEndDate As String = GlobalControl.Variables.BusinessFirmRow.Item("YearEndDate")
            Dim ValCol() As String = {}
            For i = 0 To LastKeyCollection.Count - 1
                GF1.ArrayAppend(ValCol, LastKeyCollection("col" & i))
            Next

            If LPrimaryKey = True Then
                Dim FldListStr As String = ValCol(0).Trim
                QuerryStr = "Select top(1) " & FldListStr & " from " & LTableName
                QuerryStr = QuerryStr & IIf(AddLast, " with (holdlock,tablock,updlock) ", " ")
                If ValCol(3).Trim.Length > 0 Then
                    QuerryStr = QuerryStr & " where left(" & ValCol(0).Trim & "," & ValCol(3).Trim.Length & ")  =  '" & ValCol(3).Trim & "'"
                End If
                QuerryStr = QuerryStr & " order by " & FldListStr & " desc"
            Else
                Dim lwhere As Boolean = False
                Dim FldListStr As String = IIf(ValCol(4).Length > 0, "," & ValCol(4), "") & IIf(ValCol(5).Length > 0, "," & ValCol(5), "") & IIf(ValCol(6).Length > 0, "," & ValCol(6), "")
                QuerryStr = "Select top(1) " & FldListStr & " from " & LTableName
                If ValCol(8).Trim.Length > 0 Then
                    QuerryStr = QuerryStr & " where ( charindex(" & ValCol(4).Trim & ", '" & ValCol(8).Trim & "') > 0 )"
                    lwhere = True
                End If
                If LVoucherType.Trim.Length > 0 Then
                    QuerryStr = QuerryStr & IIf(lwhere, " and ( ", " where ( ") & ValCol(4).Trim & " = '" & LVoucherType.Trim & "' )"
                    lwhere = True
                End If

                Select Case UCase(ValCol(7).Trim)
                    Case "A"
                        QuerryStr = QuerryStr & " order by " & FldListStr & " desc"
                    Case "F"
                        QuerryStr = QuerryStr & IIf(lwhere, " and ( ", " where ( ") & ValCol(6).Trim & " >= '" & lFinStartDate.Trim & "' and  " & ValCol(6).Trim & " <= '" & lFinEndDate.Trim & "' )"
                    Case "Y"
                        If CurrentDate.Trim.Length = 8 Then
                            Dim CurrentYear As String = Left(CurrentDate, 4)
                            QuerryStr = QuerryStr & IIf(lwhere, " and ( ", " where ( ") & "left(" & ValCol(6).Trim & ",4) = '" & CurrentYear & "' )"
                        End If
                    Case "M"
                        If CurrentDate.Trim.Length = 8 Then
                            Dim CurrentMonth As String = Left(CurrentDate, 6)
                            QuerryStr = QuerryStr & IIf(lwhere, " and ( ", " where ( ") & "left(" & ValCol(6).Trim & ",6) = '" & CurrentMonth & "' )"
                        End If
                    Case "D"
                        If CurrentDate.Trim.Length = 8 Then
                            QuerryStr = QuerryStr & IIf(lwhere, " and ( ", " where ( ") & ValCol(6).Trim & "  = '" & CurrentDate.Trim & "' )"
                        End If
                End Select
                QuerryStr = QuerryStr & " order by " & FldListStr & " desc"
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute SqlStringLastKey(ByVal LTableName As String, ByVal LastKeyCollection As Collection, ByVal AddLast As Boolean, Optional ByVal LPrimaryKey As Boolean = True, Optional ByVal CurrentDate As String = "", Optional ByVal LVoucherType As String = "") As String")
        End Try
        Return QuerryStr
    End Function
    ''' <summary>
    ''' Check wether an index of an SQL table exists
    ''' </summary>
    ''' <param name="SqServer">Server name as string</param>
    ''' <param name="LDataBase">DataBase name as string</param>
    ''' <param name="IndexName">Index Name to be searched</param>
    ''' <returns>Existing flag</returns>
    ''' <remarks></remarks>
    Public Function IndexFileExists(ByVal SqServer As String, ByVal LDataBase As String, ByVal IndexName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        SqServer = ConvertFromSrv0(SqServer)
        LDataBase = ConvertFromMdf0(LDataBase)
        Dim success As Boolean = False
        Dim str1 As String = "select name from sys.indexes  where upper(name)  = '" & UCase(IndexName) & "'"
        Dim sqlconobj As SqlConnection = OpenSqlConnection(SqServer, LDataBase)
        Try
            Dim cmd As New SqlCommand(str1, sqlconobj)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            success = (IIf(rdr.Read, True, False))
            rdr.Close()
            cmd.Dispose()
            sqlconobj.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, str1)
        End Try
        Return success
    End Function
    ''' <summary>
    ''' Check wether an index of an SQL table exists
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="IndexName">Index name to be searched</param>
    ''' <returns>Existing flag</returns>
    ''' <remarks></remarks>
    Public Function IndexFileExists(ByVal ServerDataBase As String, ByVal IndexName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim success As Boolean = False
        Dim str1 As String = "select name from sys.indexes  where upper(name)  = '" & UCase(IndexName) & "'"
        Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
        Dim sqlconobj As SqlConnection = OpenSqlConnection(LserverDatabase)
        Try
            Dim cmd As New SqlCommand(str1, sqlconobj)
            Dim rdr As SqlDataReader = cmd.ExecuteReader
            success = (IIf(rdr.Read, True, False))
            rdr.Close()
            cmd.Dispose()
            sqlconobj.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, str1)
        End Try
        Return success
    End Function

    ''' <summary>
    ''' To open a coonection for a dbase III file
    ''' </summary>
    ''' <param name="DBFFolder">Folder which contains a DBF file</param>
    ''' <returns>An ODBC connection </returns>
    ''' <remarks></remarks>
    Public Function DbfConnection(ByVal DBFFolder As String) As OdbcConnection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim tempconn As New OdbcConnection
        Dim sconn As String = "Driver={Microsoft dBase Driver (*.dbf)}" _
                                   & ";collatingsequence=ASCII;dbq=" & DBFFolder & "" _
                                   & ";defaultdir=" & DBFFolder & ";deleted=0;driverid=21" _
                                   & ";fil=dBase III;" _
                                   & ";maxbuffersize=8192;maxscanrows=8;pagetimeout=5;safetransactions=0" _
                                   & ";statistics=0;threads=3;uid=admin;usercommitsync=Yes"
        Try
            tempconn.ConnectionString = sconn
            tempconn.Open()
        Catch ex As Exception
            QuitError(ex, Err, sconn)
        End Try
        Return tempconn
    End Function
    ''' <summary>
    ''' Get structure of a DBF file
    ''' </summary>
    ''' <param name="FullDBFName ">FullNameDBFFile</param>
    ''' <returns>Strucure of DBF as DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetStructureDBFFile(ByVal FullDBFName As String, Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim ltemp As New DataTable
        Dim afiles As List(Of String) = GF1.FullFileNameToList(FullDBFName)
        Try
            Dim dbfsource As OdbcConnection = DbfConnection(afiles(0))
            Dim cmdsource As New OdbcCommand
            cmdsource.Connection = dbfsource
            cmdsource.CommandText = "select top (1) * from " & afiles(1)
            GlobalControl.Variables.ErrorString = cmdsource.CommandText
            Dim rdrsource As OdbcDataReader = cmdsource.ExecuteReader
            ltemp = rdrsource.GetSchemaTable
            SetPrimaryColumns(ltemp, PrimaryCols)
            cmdsource.Dispose()
            rdrsource.Close()
            dbfsource.Close()
            dbfsource.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return ltemp
    End Function
    ''' <summary>
    ''' Set Primary columns of data table for fast searching
    ''' </summary>
    ''' <param name="LdataTable">Datatable as datatable</param>
    ''' <param name="PrimaryCols">Comma separated list of primary columns</param>
    ''' <remarks></remarks>
    Public Sub SetPrimaryColumns(ByRef LdataTable As DataTable, ByVal PrimaryCols As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        If PrimaryCols.Length > 0 Then
            Try
                Dim APcol() As DataColumn = {}
                Dim Apkey() As String = Split(PrimaryCols, ",")
                For i = 0 To Apkey.Count - 1
                    If CheckColumnInDataTable(Apkey(i), LdataTable) = -1 Then
                        LdataTable.Columns.Add(Apkey(i))
                    End If
                    GF1.ArrayAppend(APcol, LdataTable.Columns(Apkey(i)))
                Next
                LdataTable.PrimaryKey = APcol
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute datafunction.SetPrimaryColumns(ByRef LdataTable As DataTable, ByVal PrimaryCols As String)")
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Set Primary columns of data table for fast searching
    ''' </summary>
    ''' <param name="LdataTable">Datatable as datatable</param>
    ''' <param name="PrimaryCols">Primary columns with type as hashtable,where key is columnname,value is columntype</param>
    ''' <remarks></remarks>
    Public Sub SetPrimaryColumns(ByRef LdataTable As DataTable, ByVal PrimaryCols As Hashtable)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        If PrimaryCols.Count > 0 Then
            Try
                Dim APcol() As DataColumn = {}
                For i = 0 To PrimaryCols.Count - 1
                    Dim mkey As String = PrimaryCols.Keys(i).ToString
                    Dim mtype As Type = PrimaryCols.Item(mkey)
                    If CheckColumnInDataTable(mkey, LdataTable) = -1 Then
                        LdataTable.Columns.Add(mkey, mtype)
                    End If
                    GF1.ArrayAppend(APcol, LdataTable.Columns(mkey))
                Next
                LdataTable.PrimaryKey = APcol
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute datafunction.SetPrimaryColumns(ByRef LdataTable As DataTable, ByVal PrimaryCols As String)")
            End Try
        End If
    End Sub


    ''' <summary>
    ''' Set Primary columns of data table for fast searching
    ''' </summary>
    ''' <param name="LdataTable">Datatable as datatable</param>
    ''' <param name="PrimaryCols">Array of primary columns</param>
    ''' <remarks></remarks>
    Public Sub SetPrimaryColumns(ByRef LdataTable As DataTable, ByVal PrimaryCols() As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        If PrimaryCols.Count > 0 Then
            Try

                Dim APcol() As DataColumn = {}
                For i = 0 To PrimaryCols.Count - 1
                    If CheckColumnInDataTable(PrimaryCols(i), LdataTable) = -1 Then
                        LdataTable.Columns.Add(PrimaryCols(i))
                    End If
                    'Dim columns(1) As DataColumn
                    'columns(0) = workTable.Columns("CustID")
                    'workTable.PrimaryKey = columns
                    GF1.ArrayAppend(APcol, LdataTable.Columns(PrimaryCols(i)))
                Next
                LdataTable.PrimaryKey = APcol
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute datafunction.SetPrimaryColumns(ByRef LdataTable As DataTable, ByVal PrimaryCols() As String)")
            End Try

        End If

    End Sub

    ''' <summary>
    ''' Get schema/structure of a SQL table in a data table
    ''' </summary>
    ''' <param name="SqServer">Sql Server name as string</param>
    ''' <param name="lDatabase">Data base name as string</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Structure in a datatable object</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaTable(ByVal SqServer As String, ByVal lDatabase As String, ByVal TableName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'ColumnName
        'ColumnOrdinal
        'ColumnSize
        'DataTypeName
        'NumericPrecision
        'NumericScale
        SqServer = ConvertFromSrv0(SqServer)
        lDatabase = ConvertFromMdf0(lDatabase)
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim ltemp As New DataTable
        Try
            Dim sqlsource As SqlConnection = OpenSqlConnection(SqServer, lDatabase)
            Dim cmdsource As New SqlCommand
            Dim rdrsource As SqlDataReader = Nothing
            cmdsource.Connection = sqlsource
            TableName = GetTableNameFromSqlIdentifier(TableName)
            cmdsource.CommandText = "select top (1) * from " & TableName
            GlobalControl.Variables.ErrorString = cmdsource.CommandText
            rdrsource = cmdsource.ExecuteReader
            SetPrimaryColumns(ltemp, "ColumnName")
            ltemp = rdrsource.GetSchemaTable
            ltemp.TableName = TableName & "_schema"
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(SqServer & "." & lDatabase, primarykeystr)
            Dim PrimaryKey As String = ""
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
            ltemp = AddColumnsInDataTable(ltemp, "PrimaryKey,Column_Name")
            If PrimaryKey.Trim.Length > 0 Then
                For i = 0 To ltemp.Rows.Count - 1
                    If LCase(ltemp(i)("ColumnName")) = LCase(PrimaryKey) Then
                        ltemp(i)("PrimaryKey") = "Y"
                    Else
                        ltemp(i)("PrimaryKey") = "N"
                    End If
                    ltemp(i)("Column_Name") = ltemp(i)("ColumnName")
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return ltemp
    End Function
    ''' <summary>
    ''' Get schema/structure of a SQL table in a data table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Structure in a datatable object</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaTable(ByVal ServerDataBase As String, ByVal TableName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'ColumnName
        'ColumnOrdinal
        'ColumnSize
        'DataTypeName
        'NumericPrecision
        'NumericScale
        ServerDataBase = ConvertFromSrv0Mdf0(ServerDataBase)
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim ltemp As New DataTable
        Try
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            Dim sqlsource As SqlConnection = OpenSqlConnection(LServerDataBase)
            Dim cmdsource As New SqlCommand
            Dim rdrsource As SqlDataReader = Nothing
            cmdsource.Connection = sqlsource
            TableName = GetTableNameFromSqlIdentifier(TableName)
            cmdsource.CommandText = "select top (1) * from " & TableName
            GlobalControl.Variables.ErrorString = cmdsource.CommandText
            rdrsource = cmdsource.ExecuteReader
            SetPrimaryColumns(ltemp, "ColumnName")
            ltemp = rdrsource.GetSchemaTable
            ltemp.TableName = TableName & "_schema"
            rdrsource.Close()
            cmdsource.Dispose()
            sqlsource.Dispose()
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, primarykeystr)
            Dim PrimaryKey As String = ""
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
            ltemp = AddColumnsInDataTable(ltemp, "PrimaryKey,Column_Name")
            If PrimaryKey.Trim.Length > 0 Then
                For i = 0 To ltemp.Rows.Count - 1
                    If LCase(ltemp(i)("ColumnName")) = LCase(PrimaryKey) Then
                        ltemp(i)("PrimaryKey") = "Y"
                    Else
                        ltemp(i)("PrimaryKey") = "N"
                    End If
                    ltemp(i)("Column_Name") = ltemp(i)("ColumnName")
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return ltemp
    End Function
    ''' <summary>
    ''' Get base table name from full sql table identifier i.e. server.database.dbo.table or 0_srv_0.0_mdf_0.dbo.Table format
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetTableNameFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        SqlTableFullIdentifier = GetServerDataBase(SqlTableFullIdentifier)
        Dim Table_Name() As String = Split(SqlTableFullIdentifier, ".")
        'If Table_Name.Count < 4 Then
        'GF1.QuitMessage(SqlTableFullIdentifier & " is not full sql identifier", "GetTableNameFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String")
        'End If
        GetTableNameFromSqlIdentifier = Table_Name(Table_Name.Count - 1)
    End Function
    ''' <summary>
    ''' Get Server name from full sql table identifier i.e. server.database.dbo.table or 0_srv_0.0_mdf_0.dbo.Table format 
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetServerFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        SqlTableFullIdentifier = GetServerDataBase(SqlTableFullIdentifier)
        Dim Table_Name() As String = Split(SqlTableFullIdentifier, ".")
        If Table_Name.Count < 4 Then
            GF1.QuitMessage(SqlTableFullIdentifier & " is not full sql identifier", "GetServerFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String")
        End If
        GetServerFromSqlIdentifier = Table_Name(Table_Name.Count - 4)
    End Function
    ''' <summary>
    ''' Get ServerDatabase  from full sql table identifier i.e. server.database.dbo.table or 0_srv_0.0_mdf_0.dbo.Table format 
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetServerDataBaseFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        SqlTableFullIdentifier = GetServerDataBase(SqlTableFullIdentifier, True)
        Dim Table_Name() As String = Split(SqlTableFullIdentifier, ".")
        If Table_Name.Count < 4 Then
            GF1.QuitMessage(SqlTableFullIdentifier & " is not full sql identifier", "GetServerDataBaseFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String")
        End If
        GetServerDataBaseFromSqlIdentifier = Table_Name(Table_Name.Count - 4) & "." & Table_Name(Table_Name.Count - 3)
    End Function


    ''' <summary>
    ''' Get database name from full sql table identifier i.e. server.database.dbo.table format or 0_srv_0.0_mdf_0.dbo.Table format
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataBaseFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        SqlTableFullIdentifier = GetServerDataBase(SqlTableFullIdentifier)
        Dim Table_Name() As String = Split(SqlTableFullIdentifier, ".")
        Dim str As String = ""
        If Table_Name.Count < 4 Then
            GF1.QuitMessage(SqlTableFullIdentifier & " is not full sql identifier", "GetDataBaseFromSqlIdentifier(ByVal SqlTableFullIdentifier) As String")
        End If
        str = Table_Name(Table_Name.Count - 3)
        Return str
    End Function

    ''' <summary>
    ''' Get schema/structure of a SQL table in a data table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Structure in a datatable object</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaInformations(ByVal ServerDataBase As String, ByVal TableName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Table_Catalog   (database name)
        'Table_Name
        'Column_Name
        'Ordinal_Position
        'Character_maximum_length
        'Data_Type
        'numeric_Precision  size of numeric,decimal type
        'NUMERIC_SCALE    digits after decimal
        'Is_Nullable  = Nullable
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim ltemp As New DataTable
        Try
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            TableName = GetTableNameFromSqlIdentifier(TableName)
            Dim querrystr As String = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" & TableName & "'"
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            SetPrimaryColumns(ltemp, "Column_Name")
            GlobalControl.Variables.ErrorString = querrystr
            ltemp = SqlExecuteDataTable(LServerDataBase, querrystr)
            ltemp.TableName = TableName & "_schema"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, primarykeystr)
            Dim PrimaryKey As String = ""
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
            ltemp = AddColumnsInDataTable(ltemp, "PrimaryKey")
            If PrimaryKey.Trim.Length > 0 Then
                For i = 0 To ltemp.Rows.Count - 1
                    If LCase(ltemp(i)("Column_Name")) = LCase(PrimaryKey) Then
                        ltemp(i)("PrimaryKey") = "Y"
                    Else
                        ltemp(i)("PrimaryKey") = "N"
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return ltemp
    End Function

    ''' <summary>
    ''' Get  SQL table's primary key field name.
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Return Primary key column name of the table</returns>
    ''' <remarks></remarks>
    Public Function GetPrimaryKey(ByVal ServerDataBase As String, ByVal TableName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim PrimaryKey As String = ""
        Try
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            Dim BaseTableName As String = GetTableNameFromSqlIdentifier(TableName)
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & BaseTableName & "'"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, primarykeystr)
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return PrimaryKey
    End Function

    ''' <summary>
    ''' Get  SQL table's primary key field name.
    ''' </summary>
    ''' <param name="mSqlConnection">opened sql connection</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Return Primary key column name of the table</returns>
    ''' <remarks></remarks>
    Public Function GetPrimaryKey(ByVal mSqlConnection As SqlConnection, ByVal TableName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim PrimaryKey As String = ""
        Try
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(mSqlConnection, primarykeystr)
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return PrimaryKey
    End Function
    ''' <summary>
    ''' Get  SQL table's primary key field name from SchemaInformation.
    ''' </summary>
    ''' <param name="SchemaTable">SchemaTable of a SQL Table returned by GetSchemaTable or GetSchemaInformations </param>
    ''' <returns>Return Primary key column name of the table</returns>
    ''' <remarks></remarks>
    Public Function GetPrimaryKey(ByVal SchemaTable As DataTable) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim PrimaryKey As String = ""
        Try
            For i = 0 To SchemaTable.Rows.Count - 1
                If IsDBNull(SchemaTable(i)("PrimaryKey")) = False Then
                    If SchemaTable(i)("PrimaryKey") = "Y" Then
                        PrimaryKey = SchemaTable(i)("Column_Name")
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return PrimaryKey
    End Function
    ''' <summary>
    ''' Get schema/structure of a SQL table in a data table
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Structure in a datatable object</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaTable(ByRef Sql_Transaction As SqlTransaction, ByVal TableName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LConnection As SqlConnection = Sql_Transaction.Connection
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        'ColumnName
        'ColumnOrdinal
        'ColumnSize
        'DataTypeName
        'NumericPrecision  numeric size
        'NumericScale    digits after decimal
        Dim ltemp As New DataTable
        Try
            Dim rdrsource As SqlDataReader = Nothing
            TableName = GetTableNameFromSqlIdentifier(TableName)
            Lcommand.Connection = LConnection
            Lcommand.CommandText = "select top (1) * from " & TableName
            GlobalControl.Variables.ErrorString = Lcommand.CommandText
            rdrsource = Lcommand.ExecuteReader
            SetPrimaryColumns(ltemp, "ColumnName")
            ltemp = rdrsource.GetSchemaTable
            rdrsource.Close()
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(Sql_Transaction, primarykeystr)
            Dim PrimaryKey As String = ""
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
            ltemp = AddColumnsInDataTable(ltemp, "PrimaryKey,Column_Name")
            If PrimaryKey.Trim.Length > 0 Then
                For i = 0 To ltemp.Rows.Count - 1
                    If LCase(ltemp(i)("ColumnName")) = LCase(PrimaryKey) Then
                        ltemp(i)("PrimaryKey") = "Y"
                    Else
                        ltemp(i)("PrimaryKey") = "N"
                    End If
                    ltemp(i)("Column_Name") = ltemp(i)("ColumnName")
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Lcommand.Dispose()
        Return ltemp
    End Function
    ''' <summary>
    ''' Get schema/structure of a SQL table in a data table
    ''' </summary>
    ''' <param name="LConnection" >Sql Connection</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Structure in a datatable object</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaTable(ByRef LConnection As SqlConnection, ByVal TableName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = LConnection.CreateCommand
        'ColumnName
        'ColumnOrdinal
        'ColumnSize
        'DataTypeName
        'NumericPrecision  numeric size
        'NumericScale    digits after decimal
        Dim ltemp As New DataTable
        Try
            Dim rdrsource As SqlDataReader = Nothing
            TableName = GetTableNameFromSqlIdentifier(TableName)
            Lcommand.Connection = LConnection
            Lcommand.CommandText = "select top (1) * from " & TableName
            GlobalControl.Variables.ErrorString = Lcommand.CommandText
            rdrsource = Lcommand.ExecuteReader
            SetPrimaryColumns(ltemp, "ColumnName")
            ltemp = rdrsource.GetSchemaTable
            rdrsource.Close()
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(LConnection, primarykeystr)
            Dim PrimaryKey As String = ""
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
            ltemp = AddColumnsInDataTable(ltemp, "PrimaryKey,Column_Name")
            If PrimaryKey.Trim.Length > 0 Then
                For i = 0 To ltemp.Rows.Count - 1
                    If LCase(ltemp(i)("ColumnName")) = LCase(PrimaryKey) Then
                        ltemp(i)("PrimaryKey") = "Y"
                    Else
                        ltemp(i)("PrimaryKey") = "N"
                    End If
                    ltemp(i)("Column_Name") = ltemp(i)("ColumnName")
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Lcommand.Dispose()
        Return ltemp
    End Function

    ''' <summary>
    ''' Get schema/structure of a SQL table in a data table
    ''' </summary>
    ''' <param name="Sql_Transaction">Sql Transaction</param>
    ''' <param name="TableName">Name of Table of getting structue/schema</param>
    ''' <returns>Structure in a datatable object</returns>
    ''' <remarks></remarks>
    Public Function GetSchemaInformations(ByRef Sql_Transaction As SqlTransaction, ByVal TableName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Table_Catalog   (database name)
        'Table_Name
        'Column_Name
        'Ordinal_Position
        'Character_maximum_length
        'Data_Type
        'numeric_Precision  size of numeric,decimal type
        'NUMERIC_SCALE    digits after decimal
        'Is_Nullable  = Nullable
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim ltemp As New DataTable
        Try
            TableName = GetTableNameFromSqlIdentifier(TableName)
            Dim querrystr As String = "SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='" & TableName & "'"
            Dim primarykeystr As String = "Select column_name as Col_PrimaryKey FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE OBJECTPROPERTY(OBJECT_ID(constraint_name), 'IsPrimaryKey') = 1  and Table_name = '" & TableName & "'"
            SetPrimaryColumns(ltemp, "Column_Name")
            GlobalControl.Variables.ErrorString = querrystr
            ltemp = SqlExecuteDataTable(Sql_Transaction, querrystr)
            ltemp.TableName = TableName & "_schema"
            GlobalControl.Variables.ErrorString = primarykeystr
            Dim dt As DataTable = SqlExecuteDataTable(Sql_Transaction, primarykeystr)
            Dim PrimaryKey As String = ""
            If dt.Rows.Count > 0 Then
                PrimaryKey = dt(0)(0).ToString.Trim
            End If
            ltemp = AddColumnsInDataTable(ltemp, "PrimaryKey")
            If PrimaryKey.Trim.Length > 0 Then
                For i = 0 To ltemp.Rows.Count - 1
                    If LCase(ltemp(i)("Column_Name")) = LCase(PrimaryKey) Then
                        ltemp(i)("PrimaryKey") = "Y"
                    Else
                        ltemp(i)("PrimaryKey") = "N"
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return ltemp
    End Function

    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="SqServer">Sql server name</param>
    ''' <param name="lDatabase">Sql database name containing table</param>
    ''' <param name="TableName">Table from which rows deleted </param>
    ''' <param name="lCondition">Where Clause as string</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByVal SqServer As String, ByVal lDatabase As String, ByVal TableName As String, ByVal lCondition As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim nrows As Integer = 0
        SqServer = ConvertFromSrv0(SqServer)
        lDatabase = ConvertFromMdf0(lDatabase)
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Try
            Dim str As String = "delete  from " & TableName & IIf(lCondition.Length > 0, " where " & lCondition, "")

            nrows = SqlExecuteNonQuery(SqServer, lDatabase, str)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.DeleteRecords(ByVal SqServer As String, ByVal lDatabase As String, ByVal TableName As String, ByVal lCondition As String) As Integer")
        End Try
        Return nrows
    End Function

    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="SqServer">Sql server name</param>
    ''' <param name="lDatabase">Sql database name containing table</param>
    ''' <param name="TableName">Table from which rows deleted </param>
    ''' <param name="lCondition">Where Clause as HashTable</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByVal SqServer As String, ByVal lDatabase As String, ByVal TableName As String, ByVal lCondition As Hashtable) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim nrows As Integer = 0
        Try
            Dim lcond As String = GF1.GetStringConditionFromHashTable(lCondition, True)
            nrows = DeleteRecords(SqServer, lDatabase, TableName, lcond)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.DeleteRecords(ByVal SqServer As String, ByVal lDatabase As String, ByVal TableName As String, ByVal lCondition As String) As Integer")
        End Try
        Return nrows
    End Function

    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="ServerDatabase ">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table from which rows deleted </param>
    ''' <param name="lCondition">Where Clause as string</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByVal ServerDatabase As String, ByVal TableName As String, ByVal lCondition As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim nrows As Integer = 0
        Try
            Dim lserverdatabase As String = GetServerDataBase(ServerDatabase)
            Dim str As String = "delete  from " & TableName & IIf(lCondition.Length > 0, " where " & lCondition, "")
            nrows = SqlExecuteNonQuery(lserverdatabase, str)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.DeleteRecords(ByVal ServerDatabase As String, ByVal TableName As String, ByVal lCondition As String) As Integer")
        End Try
        Return nrows
    End Function

    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="ServerDatabase ">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table from which rows deleted </param>
    ''' <param name="HCondition">Where Clause as hashtable</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByVal ServerDatabase As String, ByVal TableName As String, ByVal HCondition As Hashtable) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcond As String = GF1.GetStringConditionFromHashTable(HCondition, True)
        Dim nrows As Integer = DeleteRecords(ServerDatabase, TableName, lcond)
        Return nrows
    End Function

    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction by reference</param>
    ''' <param name="TableName" >Table from which rows deleted</param>
    ''' <param name="lCondition" >Where Clause as string</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByRef Sql_Transaction As SqlTransaction, ByVal TableName As String, ByVal lCondition As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim sqlstr As String = "delete  from " & TableName & IIf(lCondition.Length > 0, " where " & lCondition, "")
        Lcommand.CommandText = sqlstr
        Dim k As Integer = 0
        Try
            k = Lcommand.ExecuteNonQuery()
        Catch ex As Exception
            QuitError(ex, Err, sqlstr)
        End Try
        Return k
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction by reference</param>
    ''' <param name="TableName" >Table from which rows deleted</param>
    ''' <param name="PrimaryKey" >Primary key field name</param>
    ''' <param name="PrimaryKeyValue" >Primary Key field value </param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByRef Sql_Transaction As SqlTransaction, ByVal TableName As String, ByVal PrimaryKey As String, ByVal PrimaryKeyValue As Integer) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim lcondition As String = PrimaryKey & " = " & PrimaryKeyValue
        Dim sqlstr As String = "delete  from " & TableName & IIf(lcondition.Length > 0, " where " & lcondition, "")
        Lcommand.CommandText = sqlstr
        Dim k As Integer = 0
        Try
            k = Lcommand.ExecuteNonQuery()
        Catch ex As Exception
            QuitError(ex, Err, sqlstr)
        End Try
        Return k
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction by reference</param>
    ''' <param name="TableName" >Table from which rows deleted</param>
    ''' <param name="PrimaryKey" >Primary key field name</param>
    ''' <param name="PrimaryKeyValue" >Primary Key field value </param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByRef Sql_Transaction As SqlTransaction, ByVal TableName As String, ByVal PrimaryKey As String, ByVal PrimaryKeyValue As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim lcondition As String = PrimaryKey & " = '" & PrimaryKeyValue & "'"
        Dim sqlstr As String = "delete  from " & TableName & IIf(lcondition.Length > 0, " where " & lcondition, "")
        Lcommand.CommandText = sqlstr
        Dim k As Integer = 0
        Try
            k = Lcommand.ExecuteNonQuery()
        Catch ex As Exception
            QuitError(ex, Err, sqlstr)
        End Try
        Return k
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction by reference</param>
    ''' <param name="TableName" >Table from which rows deleted</param>
    ''' <param name="FieldValues" >FieldValues as hashtable where key is fieldname and field value is hashtable value</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByRef Sql_Transaction As SqlTransaction, ByVal TableName As String, ByVal FieldValues As Hashtable) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcondition As String = GF1.GetStringConditionFromHashTable(FieldValues, True)
        Dim nrows As Integer = DeleteRecords(Sql_Transaction, TableName, lcondition)
        Return nrows
    End Function
    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="SQL_Connection">SQL connection already open</param>
    ''' <param name="SQL_Command">SQL command already created</param>
    ''' <param name="TableName">Table from which rows deleted</param>
    ''' <param name="lCondition">Where Clause as string</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByRef SQL_Connection As SqlConnection, ByRef SQL_Command As SqlCommand, ByVal TableName As String, ByVal lCondition As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim sqlstr As String = "delete  from " & TableName & IIf(lCondition.Length > 0, " where " & lCondition, "")
        SQL_Command.CommandText = sqlstr
        SQL_Command.Connection = SQL_Connection
        Dim k As Integer = 0
        Try
            k = SQL_Command.ExecuteNonQuery()
        Catch ex As Exception
            QuitError(ex, Err, sqlstr)
        End Try
        Return k
    End Function
    ''' <summary>
    ''' Delete records from an SQL Table
    ''' </summary>
    ''' <param name="SQL_Connection">SQL connection already open</param>
    ''' <param name="TableName">Table from which rows deleted</param>
    ''' <param name="lCondition">Where Clause as string</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeleteRecords(ByRef SQL_Connection As SqlConnection, ByVal TableName As String, ByVal lCondition As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        TableName = ConvertFromSrv0Mdf0(TableName, True)
        Dim sqlstr As String = "delete  from " & TableName & IIf(lCondition.Length > 0, " where " & lCondition, "")
        Dim SqlCmd As New SqlCommand
        SqlCmd.Connection = SQL_Connection
        SqlCmd.CommandText = sqlstr
        Dim k As Integer = 0
        Try
            k = SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, sqlstr)
        End Try
        Return k
    End Function

    ''' <summary>
    ''' Get Datatable from a DBF table
    ''' </summary>
    ''' <param name="FullDBFName">Full dbf file name with path and extension</param>
    ''' <param name="Lcondition">Where Clause as string</param>
    ''' <param name="Lorder ">Order By Clause as string </param> ''' 
    ''' <param name="PrimaryCols">Comma separated string of primary columns</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetDataFromDBF(ByVal FullDBFName As String, Optional ByVal Lcondition As String = "", Optional ByVal Lorder As String = "", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim ltemp As New DataTable
        Dim afiles As List(Of String) = GF1.FullFileNameToList(FullDBFName)
        Dim str As String = "select *   from " & afiles(1).Trim & IIf(Lcondition.Length > 0, " where " & Lcondition, "")
        str = str & IIf(Lorder.Length > 0, " order by " & Lorder, "")
        Try
            Dim dbfcon As OdbcConnection = DbfConnection(afiles(0))
            Dim dbfcmd As New OdbcCommand With {.commandtext = str, .Connection = dbfcon}
            Dim dbfda As New OdbcDataAdapter(dbfcmd)
            SetPrimaryColumns(ltemp, PrimaryCols)
            dbfda.Fill(ltemp)
            dbfcon.Close()
            dbfcmd.Dispose()
            dbfda.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, str)
        End Try
        Return ltemp
    End Function
    '''' <summary>
    '''' Get data table from an excel file
    '''' </summary>
    '''' <param name="FullExcelFile">Full exel file name with path and extension</param>
    '''' <param name="LCondition">Where Clause as string , it is case sensitive</param>
    '''' <param name="Lorder">Order By Clause as string</param>
    '''' <param name="PrimaryCols">Comma separated string of primary columns</param>
    '''' <returns>Data table of rows</returns>
    '''' <remarks></remarks>
    'Public Function GetDataFromExcel(ByVal FullExcelFile As String, Optional ByVal LCondition As String = "", Optional ByVal Lorder As String = "", Optional ByVal PrimaryCols As String = "") As DataTable
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    Dim Dt As New DataTable
    '    Try
    '        Dim msheet1 As String = FirstExcelSheetName(FullExcelFile)
    '        Dim ExcelVersion As String = IIf(LCase(FullExcelFile).Contains("xlsx"), "N", "O")
    '        Dim ConnectionString As String = IIf(ExcelVersion = "N", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullExcelFile & ";Extended Properties=Excel 12.0;", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FullExcelFile & ";Extended Properties=Excel 8.0; ")
    '        Dim ConnectionObject As New OleDb.OleDbConnection(ConnectionString)
    '        Dim QryString As String = "select  * from [" & msheet1 & "$]" & IIf(LCondition.Length = 0, "", " where " & LCondition)
    '        QryString = QryString & IIf(Lorder.Length > 0, " order by " & Lorder, "")
    '        GlobalControl.Variables.ErrorString = QryString
    '        SetPrimaryColumns(Dt, PrimaryCols)
    '        Dim ExcelAdapter As New OleDb.OleDbDataAdapter(QryString, ConnectionObject)
    '        Dim aa As New OleDb.OleDbCommand
    '        ExcelAdapter.Fill(Dt)
    '        ConnectionObject.Dispose()
    '        ExcelAdapter.Dispose()
    '    Catch ex As Exception
    '        QuitError(ex, Err, GlobalControl.Variables.ErrorString)
    '    End Try
    '    Return Dt
    'End Function

    ''' <summary>
    ''' Get data table from an excel file
    ''' </summary>
    ''' <param name="FullExcelFile">Full exel file name with path and extension</param>
    ''' <param name="LCondition">Where Clause as string , it is case sensitive</param>
    ''' <param name="Lorder">Order By Clause as string</param>
    ''' <param name="PrimaryCols">Comma separated string of primary columns</param>
    ''' <param name="SheetName" >SheetName of excel from which data fetched,default is first sheet</param>
    ''' <returns>Data table of rows</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromExcel(ByVal FullExcelFile As String, Optional ByVal LCondition As String = "", Optional ByVal Lorder As String = "", Optional ByVal PrimaryCols As String = "", Optional ByVal SheetName As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Dt As New DataTable
        Try
            Dim msheet1 As String = IIf(SheetName.Length = 0, FirstExcelSheetName(FullExcelFile), SheetName)
            Dim ExcelVersion As String = IIf(LCase(FullExcelFile).Contains("xlsx"), "N", "O")
            Dim ConnectionString As String = IIf(ExcelVersion = "N", "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FullExcelFile & ";Extended Properties=Excel 12.0;", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FullExcelFile & ";Extended Properties=Excel 8.0; ")
            Dim ConnectionObject As New OleDb.OleDbConnection(ConnectionString)
            Dim QryString As String = "select  * from [" & msheet1 & "$]" & IIf(LCondition.Length = 0, "", " where " & LCondition)
            QryString = QryString & IIf(Lorder.Length > 0, " order by " & Lorder, "")
            GlobalControl.Variables.ErrorString = QryString
            SetPrimaryColumns(Dt, PrimaryCols)
            Dim ExcelAdapter As New OleDb.OleDbDataAdapter(QryString, ConnectionObject)
            Dim aa As New OleDb.OleDbCommand
            ExcelAdapter.Fill(Dt)
 ExcelAdapter.Dispose()
            ConnectionObject.Close()
            ConnectionObject.Dispose()
           
        Catch ex As Exception
            MsgBox(ex.Message & "  " & Err.Description)
            Console.WriteLine(ex.Message & "  " & Err.Description)
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return Dt
    End Function



    ''' <summary>
    ''' To Get Duplicate rows having common column values
    ''' </summary>
    ''' <param name="LdataTable">DataTable whose rows to be checked</param>
    ''' <param name="ColumnsToCheck">Comma separated columns , whose values to be checked for duplicacy</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDuplicateRows(ByVal LdataTable As DataTable, ByVal ColumnsToCheck As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DupTable As DataTable = LdataTable.Clone
        Try
            Dim dt1 As DataTable = SortDataTable(LdataTable, ColumnsToCheck)
            Dim Icount As Integer = 0
            While Icount < dt1.Rows.Count
                Try
                    Dim aValues() As Object = GetColumnValuesFromRow(dt1.Rows(Icount), ColumnsToCheck)
                    Dim n As Integer = 0
                    While MatchColumnValuesOfRow(dt1.Rows(Icount), ColumnsToCheck, aValues) = True
                        Try
                            n = n + 1
                            If n > 1 Then
                                DupTable = AddRowInDataTable(DupTable, dt1.Rows(Icount))
                            End If
                            Icount = Icount + 1
                            If Icount > dt1.Rows.Count - 1 Then
                                Exit While
                            End If
                        Catch ex As Exception
                            GF1.QuitError(ex, Err, "error in GetDuplicateRows-1")
                        End Try
                    End While
                    If Icount > dt1.Rows.Count - 1 Then
                        Exit While
                    End If
                Catch ex As Exception
                    GF1.QuitError(ex, Err, "error in GetDuplicateRows")
                End Try
            End While
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.GetDuplicateRows")
        End Try
        Return DupTable
    End Function


    ''' <summary>
    ''' Create index files of sql tables , whoose names are stored in an excel file(column name -fname) and index fields and index filename are stored in (column name-ntxf) eg. value of ntfx may be (  IndexFld1*IndexFile1~IndexFld2*IndexFile2~IndexFld3*IndexFile3)
    ''' </summary>
    ''' <param name="Sqserver">Sql Server Name </param>
    ''' <param name="LDataBase">Sql DataBase Name</param>
    ''' <param name="ExcelSource">Excel file name with path and extension</param>
    ''' <param name="IncludeTables">Comma separated String of Table names, default is all</param>
    ''' <returns>Completion flag</returns>
    ''' <remarks></remarks>

    Public Function CreateIndexByExcelList(ByVal Sqserver As String, ByVal LDataBase As String, ByVal ExcelSource As String, Optional ByVal IncludeTables As String = "") As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DtTable As DataTable = GetDataFromExcel(ExcelSource)
        Dim mflag As Boolean = False
        Try
            If DtTable Is Nothing Then
                Return False
                Exit Function
            End If
            Sqserver = ConvertFromSrv0(Sqserver)
            LDataBase = ConvertFromMdf0(LDataBase)
            Dim Minclude() As String = Split(IncludeTables, ",")
            For i = 0 To DtTable.Rows.Count - 1
                Dim Mtable As String = DtTable.Rows(i).Item("fname")
                If IncludeTables.Length > 0 Then
                    If GF1.ArrayFind(Minclude, Mtable) < 0 Then
                        Continue For
                    End If
                End If
                Dim mntxf As String = DtTable.Rows(i).Item("ntxf")
                mflag = CreateIndex(Sqserver, LDataBase, Mtable, mntxf)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunction.CreateIndexByExcelList(ByVal Sqserver As String, ByVal LDataBase As String, ByVal ExcelSource As String, Optional ByVal IncludeTables As String = "") As Boolean")
        End Try
        Return mflag
    End Function
    ''' <summary>
    ''' To create index files on an SQL Table specifying its key fields and index file names
    ''' </summary>
    ''' <param name="SqServer">Sql Server Name</param>
    ''' <param name="LDataBase">Sql Data Base name</param>
    ''' <param name="LTable">Sql Table Name</param>
    ''' <param name="IndexKey ">Comma separated list of index fields</param>
    ''' <param name="IndexFile" >Name of index file</param> 
    ''' <returns>Completion flag</returns>
    ''' <remarks></remarks>
    Public Function CreateIndex(ByVal SqServer As String, ByVal LDataBase As String, ByVal LTable As String, ByVal IndexKey As String, ByVal IndexFile As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mflag As Boolean = False
        Dim str1 As String = ""
        SqServer = ConvertFromSrv0(SqServer)
        LDataBase = ConvertFromMdf0(LDataBase)
        LTable = ConvertFromSrv0Mdf0(LTable, True)
        Try
            If IndexFileExists(SqServer, LDataBase, IndexFile) Then
                str1 = "drop index " & LTable & "." & IndexFile
                Dim k1 As Integer = SqlExecuteNonQuery(SqServer, LDataBase, str1)
            End If
            str1 = "create index " & IndexFile & " on " & LTable & "(" & IndexKey.Trim & ")"
            Dim k2 As Integer = SqlExecuteNonQuery(SqServer, LDataBase, str1)
            mflag = True
            Return mflag
        Catch ex As Exception
            GF1.QuitError(ex, Err, str1)
        End Try
    End Function
    Public Function CreateIndex(ByVal ServerDatabase As String, ByVal LTable As String, ByVal IndexKey As String, ByVal IndexFile As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDataBase As String = GetServerDataBase(ServerDatabase)
        LTable = ConvertFromSrv0Mdf0(LTable, True)
        Dim mflag As Boolean = False
        Dim str1 As String = ""
        Try
            If IndexFileExists(LserverDataBase, IndexFile) Then
                str1 = "drop index " & LTable & "." & IndexFile
                Dim k1 As Integer = SqlExecuteNonQuery(LserverDataBase, str1)
            End If
            str1 = "create index " & IndexFile & " on " & LTable & "(" & IndexKey.Trim & ")"
            Dim k2 As Integer = SqlExecuteNonQuery(LserverDataBase, str1)
            mflag = True
            Return mflag
        Catch ex As Exception
            GF1.QuitError(ex, Err, str1)
        End Try
    End Function

    Public Function CreateIndex(ByVal TablesListRow As DataRow) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mfname As String = ConvertFromSrv0Mdf0(TablesListRow("tablename").ToString.Trim, True)
            Dim msqserver As String = ConvertFromSrv0(TablesListRow("sqlserver").ToString.Trim)
            Dim msqdatabase As String = ConvertFromMdf0(TablesListRow("sqldatabase").ToString.Trim)
            Dim mserverdatabase As String = msqserver.Trim & "." & msqdatabase
            Dim mflag As Boolean = False
            For i = 1 To 10
                Dim MindexKey As String = TablesListRow("indexkey" & i.ToString).ToString.Trim
                Dim MindexFile As String = TablesListRow("indexfile" & i.ToString).ToString.Trim
                If MindexKey.Length > 0 And MindexFile.Length > 0 Then
                    CreateIndex(mserverdatabase, mfname, MindexKey, MindexFile)
                End If
            Next

        Catch ex As Exception
            GF1.QuitError(ex, Err, "Unable to execute DataFunction.CreateIndex(ByVal TablesListRow As DataRow) As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' Execute T-SQL non query statements 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format </param>
    ''' <param name="SqlStr">Sql Non querry statements</param>
    ''' <param name="SqlParameterArr">Sql Parameters array</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>

    Public Function SqlExecuteNonQuery(ByVal ServerDataBase As String, ByVal SqlStr As String, Optional ByVal SqlParameterArr() As SqlParameter = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim k As Integer = 0
        Try
            Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
            Dim sqlcmd As New SqlCommand
            Dim sqlcon As SqlConnection = OpenSqlConnection(LserverDatabase)
            sqlcmd.CommandText = SqlStr
            sqlcmd.Connection = sqlcon
            If Not SqlParameterArr Is Nothing Then
                sqlcmd.Parameters.Clear()
                If SqlParameterArr.Count > 0 Then
                    sqlcmd.Parameters.AddRange(SqlParameterArr)
                End If
            End If
            GlobalControl.Variables.ErrorString = SqlStr
            k = sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
            sqlcon.Close()
           
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return k
    End Function
    ''' <summary>
    ''' Execute T-SQL non query statements 
    ''' </summary>
    ''' <param name="ServerName" >servername </param>
    ''' <param name="LDataBase" >Database name</param>
    ''' <param name="SqlStr">Sql Non querry statements</param>
    ''' <param name="SqlParameterArr">Sql Parameters array</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>

    Public Function SqlExecuteNonQuery(ByVal ServerName As String, ByVal LDataBase As String, ByVal SqlStr As String, Optional ByVal SqlParameterArr() As SqlParameter = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim k As Integer = 0
        Try
            Dim sqlcmd As New SqlCommand
            Dim sqlcon As SqlConnection = OpenSqlConnection(ServerName, LDataBase)
            GlobalControl.Variables.ErrorString = SqlStr
            sqlcmd.CommandText = SqlStr
            sqlcmd.Connection = sqlcon
            If Not SqlParameterArr Is Nothing Then
                sqlcmd.Parameters.Clear()
                If SqlParameterArr.Count > 0 Then
                    sqlcmd.Parameters.AddRange(SqlParameterArr)
                End If
            End If

            k = sqlcmd.ExecuteNonQuery()
            sqlcmd.Dispose()
            sqlcon.Close()
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return k
    End Function
    ''' <summary>
    ''' Execute T-SQL non query string command
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction by reference</param>
    ''' <param name="SqlStr">Non querry string to be executed</param>
    ''' <param name="SqlParameterArr" >Sql parameters as array</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteNonQuery(ByRef Sql_Transaction As SqlTransaction, ByVal SqlStr As String, Optional ByVal SqlParameterArr() As SqlParameter = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        Lcommand.CommandText = SqlStr
        Lcommand.Connection = Sql_Transaction.Connection
        GlobalControl.Variables.ErrorString = SqlStr
        Dim k As Integer = 0
        Try
            If Not SqlParameterArr Is Nothing Then
                Lcommand.Parameters.Clear()
                If SqlParameterArr.Count > 0 Then
                    Lcommand.Parameters.AddRange(SqlParameterArr)
                End If
            End If
            k = Lcommand.ExecuteNonQuery()
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Lcommand.Dispose()
        Return k
    End Function
    ''' <summary>
    ''' Execute T-SQL non query string command
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql Transaction by reference</param>
    ''' <param name="Lcommand" >sql command</param>
    ''' <param name="SqlStr">Non querry string to be executed</param>
    ''' <param name="SqlParameterArr" >Sql parameters as array</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteNonQuery(ByRef Sql_Transaction As SqlTransaction, ByRef Lcommand As SqlCommand, ByVal SqlStr As String, Optional ByVal SqlParameterArr() As SqlParameter = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Lcommand.CommandText = SqlStr
        Lcommand.Transaction = Sql_Transaction
        Lcommand.Connection = Sql_Transaction.Connection
        GlobalControl.Variables.ErrorString = SqlStr
        Dim k As Integer = 0
        Try
            If Not SqlParameterArr Is Nothing Then
                Lcommand.Parameters.Clear()
                If SqlParameterArr.Count > 0 Then
                    Lcommand.Parameters.AddRange(SqlParameterArr)
                End If
            End If
            k = Lcommand.ExecuteNonQuery()
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return k
        Lcommand.Dispose()
    End Function



    ''' <summary>
    ''' Execute T-SQL non query string command
    ''' </summary>
    ''' <param name="SQL_Connection">SQL connection already open </param>
    ''' <param name="SQL_Command">SQL command already created</param>
    ''' <param name="SqlStr">Non querry string to be executed</param>
    ''' <param name="SqlParameterArr" >Sql parameters as array</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteNonQuery(ByRef SQL_Connection As SqlConnection, ByRef SQL_Command As SqlCommand, ByVal SqlStr As String, Optional ByVal SqlParameterArr() As SqlParameter = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        SQL_Command.CommandText = SqlStr
        SQL_Command.Connection = SQL_Connection
        GlobalControl.Variables.ErrorString = SqlStr
        Dim k As Integer = 0
        Try
            If Not SqlParameterArr Is Nothing Then
                SQL_Command.Parameters.Clear()
                If SqlParameterArr.Count > 0 Then
                    SQL_Command.Parameters.AddRange(SqlParameterArr)
                End If
            End If
            k = SQL_Command.ExecuteNonQuery()
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return k
    End Function
    ''' <summary>
    ''' Execute T-SQL non query string command
    ''' </summary>
    ''' <param name="SQL_Connection">SQL connection already open </param>
    ''' <param name="SqlStr">Non querry string to be executed</param>
    ''' <param name="SqlParameterArr" >Sql parameters as array</param>
    ''' <returns>No. of rows affected</returns>
    ''' <remarks></remarks>

    Public Function SqlExecuteNonQuery(ByRef SQL_Connection As SqlConnection, ByVal SqlStr As String, Optional ByVal SqlParameterArr() As SqlParameter = Nothing) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlCmd As New SqlCommand
        SqlCmd.Connection = SQL_Connection
        SqlCmd.CommandText = SqlStr
        GlobalControl.Variables.ErrorString = SqlStr
        Dim k As Integer = 0
        Try
            If Not SqlParameterArr Is Nothing Then
                SqlCmd.Parameters.Clear()
                If SqlParameterArr.Count > 0 Then
                    SqlCmd.Parameters.AddRange(SqlParameterArr)
                End If
            End If
            k = SqlCmd.ExecuteNonQuery()
            SqlCmd.Dispose()
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return k
    End Function

    ''' <summary>
    ''' Execute T-SQL Scaler querry  statement 
    ''' </summary>
    ''' <param name="ServerDatabase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="Sqlstr">Scaler Querry Strig command to be executed</param>
    ''' <returns>An object type return</returns>
    ''' <remarks></remarks>

    Public Function SqlExecuteScalarQuery(ByVal ServerDatabase As String, ByVal Sqlstr As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim k As Integer = 0
        Dim LserverDatabase As String = GetServerDataBase(ServerDatabase)
        Try
            Dim sqlcmd As New SqlCommand
            Dim sqlcon As SqlConnection = OpenSqlConnection(LserverDatabase)
            GlobalControl.Variables.ErrorString = Sqlstr
            sqlcmd.CommandText = Sqlstr
            sqlcmd.Connection = sqlcon
            Dim aobj As Object = sqlcmd.ExecuteScalar
            If aobj IsNot Nothing Then
                k = CInt(aobj)
            End If
            sqlcmd.Dispose()
            sqlcon.Close()
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return k
    End Function
    ''' <summary>
    '''  Execute T-SQL Scaler querry  statement 
    ''' </summary>
    ''' <param name="Sql_Transaction" >Sql transaction by reference</param>
    ''' <param name="Sqlstr">Scaler Querry Strig command to ben executed</param>
    ''' <returns>An integer type return</returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteScalarQuery(ByRef Sql_Transaction As SqlTransaction, ByVal Sqlstr As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        Dim sqlval As Integer = 0
        Lcommand.CommandText = Sqlstr
        GlobalControl.Variables.ErrorString = Sqlstr
        Try
            Dim aobj As Object = Lcommand.ExecuteScalar
            If aobj IsNot Nothing Then
                sqlval = CInt(aobj)
            End If

        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return sqlval
        Lcommand.Dispose()
    End Function
    ''' <summary>
    '''  Execute T-SQL Scaler querry  statement 
    ''' </summary>
    ''' <param name="Sql_Connection" >Sql Connection by reference</param>
    ''' <param name="Sqlstr">Scaler Querry Strig command to ben executed</param>
    ''' <returns>An integer type return</returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteScalarQuery(ByRef Sql_Connection As SqlConnection, ByVal Sqlstr As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Connection.CreateCommand
        Lcommand.Connection = Sql_Connection
        Dim sqlval As Integer = 0
        Lcommand.CommandText = Sqlstr
        GlobalControl.Variables.ErrorString = Sqlstr
        Try
            Dim aobj As Object = Lcommand.ExecuteScalar
            If aobj IsNot Nothing Then
                sqlval = CInt(aobj)
            End If
            sqlval = CInt(Lcommand.ExecuteScalar())
        Catch ex As Exception
            GF1.QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return sqlval
        Lcommand.Dispose()
    End Function


    ''' <summary>
    ''' Execute T-SQL  querry  statement
    ''' </summary>
    ''' <param name="ServerDatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SqlStr"></param>
    ''' <param name="FromClauseTable" >TableName used in from clause of sqlstr query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteDataTable(ByVal ServerDatabase As String, ByVal SqlStr As String, Optional ByVal FromClauseTable As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBase(ServerDatabase)
        Dim dtble As New DataTable
        Dim sqlcmd As New SqlCommand
        Try
            Dim sqlcon As SqlConnection = OpenSqlConnection(LserverDatabase)
            GlobalControl.Variables.ErrorString = SqlStr
            sqlcmd.CommandText = SqlStr
            sqlcmd.Connection = sqlcon
            Dim sqlda As New SqlDataAdapter(sqlcmd)
            If FromClauseTable.Trim.Length > 0 Then
                Dim mserverdb As List(Of String) = BreakServerDataBase(LserverDatabase)
                sqlda.TableMappings.Add("Table", AddSquareBrackets(mserverdb(0)) & "." & AddSquareBrackets(mserverdb(1)) & ".dbo." & FromClauseTable)
            End If
            sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            sqlda.Fill(dtble)
            sqlcmd.Dispose()
            sqlcon.Close()
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtble
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="ServerDatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SqlStr">Querry Strig command to be executed</param>
    ''' <param name="FromClauseTables" >Comma separated string of tables used in from clause(s) of sqlstr query,Set as TableName property  of DataTable(server.database.dbo.table)</param>
    ''' <param name="PrimaryColumns" >~  separated string of columns(Pcol1,Pcol2 etc.) will be used as primarykeys of datatables ,These columns may be different sql tables primary keys,If not defined, Sql tables primary keys will be DataTables primary key</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteDataSet(ByVal ServerDatabase As String, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBase(ServerDatabase)
        Dim dtSet As New DataSet
        Dim sqlcmd As New SqlCommand
        Try
            Dim sqlcon As SqlConnection = OpenSqlConnection(LserverDatabase)
            GlobalControl.Variables.ErrorString = SqlStr
            sqlcmd.CommandText = SqlStr
            sqlcmd.Connection = sqlcon
            Dim sqlda As New SqlDataAdapter(sqlcmd)
            Dim n As Int16 = 0, m As Int16 = 0
            If FromClauseTables.Trim.Length > 0 Then
                Dim mserverdb As List(Of String) = BreakServerDataBase(LserverDatabase)
                Dim atables() As String = Split(FromClauseTables, ",")
                For i = 0 To atables.Count - 1
                    If i = 0 Then
                        sqlda.TableMappings.Add("Table", AddSquareBrackets(mserverdb(0)) & "." & AddSquareBrackets(mserverdb(1)) & ".dbo." & atables(i))
                    Else
                        sqlda.TableMappings.Add("Table" & CStr(i), AddSquareBrackets(mserverdb(0)) & "." & AddSquareBrackets(mserverdb(1)) & ".dbo." & atables(i))
                    End If
                    n = i
                Next
            End If
            If PrimaryColumns.Trim.Length > 0 Then
                Dim mpcols() As String = PrimaryColumns.Split("~")
                For i = 0 To mpcols.Count - 1
                    SetPrimaryColumns(dtSet.Tables(i), mpcols(i))
                    m = i
                Next
            Else
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            End If
            If n <> m Then
                QuitMessage("Each table does not have primary columns definition " & vbCrLf & SqlStr, "SqlExecuteDataSet(ByVal ServerDatabase As String, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet")
            End If
            sqlda.Fill(dtSet)
            sqlcmd.Dispose()
            sqlcon.Close()
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtSet
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="ServerDatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SqlStr">Querry Strig command to be executed</param>
    ''' <param name="StartRecord" >Start Record No</param>
    ''' <param name="MaxRecord" >Maximum Records to fetch</param>
    ''' <param name="PrimaryColumns" >~  separated string of columns(Pcol1,Pcol2 etc.) will be used as primarykeys of datatables ,These columns may be different sql tables primary keys,If not defined, Sql tables primary keys will be DataTables primary key</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SqlExecuteDataSet(ByVal ServerDatabase As String, ByVal SqlStr As String, ByVal StartRecord As Integer, ByVal MaxRecord As Integer, Optional ByVal PrimaryColumns As String = "") As DataSet
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBase(ServerDatabase)
        Dim dtSet As New DataSet
        Dim sqlcmd As New SqlCommand
        Try
            Dim sqlcon As SqlConnection = OpenSqlConnection(LserverDatabase)
            GlobalControl.Variables.ErrorString = SqlStr
            sqlcmd.CommandText = SqlStr
            sqlcmd.Connection = sqlcon
            Dim sqlda As New SqlDataAdapter(sqlcmd)
            If PrimaryColumns.Trim.Length > 0 Then
                Dim mpcols() As String = PrimaryColumns.Split("~")
                For i = 0 To mpcols.Count - 1
                    SetPrimaryColumns(dtSet.Tables(i), mpcols(i))
                Next
            Else
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            End If
            sqlda.Fill(dtSet, StartRecord, MaxRecord, sqlda.TableMappings(0).ToString)
            sqlcmd.Dispose()
            sqlcon.Close()
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtSet
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="Sql_Transaction">Sql Transaction </param>
    ''' <param name="SqlStr">Querry Strig command to be executed</param>
    ''' <param name="FromClauseTable" >Comma separated string of tables used in from clause(s) of sqlstr query,Set as TableName property  of DataTable(server.database.dbo.table)</param>
    ''' <param name="PrimaryColumns" >String of columns(Pcol1,Pcol2 etc.) will be used as primarykeys of datatable ,These columns may be different sql table's primary keys,If not defined, Sql tables primary keys will be DataTable primary key</param>
    ''' <returns>return a data table object </returns>
    ''' <remarks></remarks>
    ''' 
    Public Function SqlExecuteDataTable(ByRef Sql_Transaction As SqlTransaction, ByVal SqlStr As String, Optional ByVal FromClauseTable As String = "", Optional ByVal PrimaryColumns As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        Dim dtble As New DataTable
        Try
            Lcommand.CommandText = SqlStr
            GlobalControl.Variables.ErrorString = SqlStr
            Dim sqlda As New SqlDataAdapter(Lcommand)
            If FromClauseTable.Trim.Length > 0 Then
                Dim mserver As String = Sql_Transaction.Connection.DataSource
                Dim mdatabase As String = Sql_Transaction.Connection.Database
                sqlda.TableMappings.Add("Table", AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & FromClauseTable)
            End If

            If PrimaryColumns.Trim.Length > 0 Then
                SetPrimaryColumns(dtble, PrimaryColumns)
            Else
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            End If
            sqlda.Fill(dtble)
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtble
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="LConnection">Sql Connection </param>
    ''' <param name="SqlStr">Querry Strig command to be executed</param>
    ''' <param name="FromClauseTable" >Comma separated string of tables used in from clause(s) of sqlstr query,Set as TableName property  of DataTable(server.database.dbo.table)</param>
    ''' <param name="PrimaryColumns" >String of columns(Pcol1,Pcol2 etc.) will be used as primarykeys of datatable ,These columns may be different sql table's primary keys,If not defined, Sql tables primary keys will be DataTable primary key</param>
    ''' <returns>return a data table object </returns>
    ''' <remarks></remarks>
    ''' 
    Public Function SqlExecuteDataTable(ByRef LConnection As SqlConnection, ByVal SqlStr As String, Optional ByVal FromClauseTable As String = "", Optional ByVal PrimaryColumns As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = LConnection.CreateCommand
        Dim dtble As New DataTable
        Try
            Lcommand.CommandText = SqlStr
            GlobalControl.Variables.ErrorString = SqlStr
            Dim sqlda As New SqlDataAdapter(Lcommand)
            If FromClauseTable.Trim.Length > 0 Then
                Dim mserver As String = LConnection.DataSource
                Dim mdatabase As String = LConnection.Database
                sqlda.TableMappings.Add("Table", AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & FromClauseTable)
            End If

            If PrimaryColumns.Trim.Length > 0 Then
                SetPrimaryColumns(dtble, PrimaryColumns)
            Else
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            End If
            sqlda.Fill(dtble)
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtble
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="Sql_Transaction">Sql Transaction </param>
    ''' <param name="SqlStr">Querry Strig command to be executed</param>
    ''' <param name="FromClauseTables" >Comma separated string of tables used in from clause(s) of sqlstr query,Set as TableName property  of DataTable(server.database.dbo.table)</param>
    ''' <param name="PrimaryColumns" >~  separated string of columns(Pcol1,Pcol2 etc.) will be used as primarykeys of datatables ,These columns may be different sql tables primary keys,If not defined, Sql tables primary keys will be DataTables primary key</param>
    ''' <returns>return a data set object </returns>
    ''' <remarks></remarks>
    ''' 
    Public Function SqlExecuteDataSet(ByRef Sql_Transaction As SqlTransaction, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        Dim dtSet As New DataSet
        Try
            Lcommand.CommandText = SqlStr
            GlobalControl.Variables.ErrorString = SqlStr
            Dim sqlda As New SqlDataAdapter(Lcommand)
            Dim n As Int16 = 0, m As Int16 = 0
            If FromClauseTables.Trim.Length > 0 Then
                Dim mserver As String = Sql_Transaction.Connection.DataSource
                Dim mdatabase As String = Sql_Transaction.Connection.Database
                Dim atables() As String = Split(FromClauseTables, ",")
                For i = 0 To atables.Count - 1
                    If i = 0 Then
                        sqlda.TableMappings.Add("Table", AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & atables(i))
                    Else
                        sqlda.TableMappings.Add("Table" & CStr(i), AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & atables(i))
                    End If
                    n = i
                Next
            End If
            If PrimaryColumns.Trim.Length > 0 Then
                Dim mpcols() As String = PrimaryColumns.Split("~")
                For i = 0 To mpcols.Count - 1
                    SetPrimaryColumns(dtSet.Tables(i), mpcols(i))
                    m = i
                Next
            Else
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            End If
            If n <> m Then
                QuitMessage("Each table does not primary columns definition " & vbCrLf & SqlStr, "SqlExecuteDataSet(ByVal ServerDatabase As String, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet")
            End If
            sqlda.Fill(dtSet)
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtSet
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="Sql_Transaction">Sql Transaction </param>
    ''' <param name="aClsTables" >An array of tableclass objetcs</param>
    ''' <returns>return a data set  object </returns>
    ''' <remarks></remarks>
    ''' 
    Public Function SqlExecuteDataSet(ByRef Sql_Transaction As SqlTransaction, ByVal aClsTables() As Object) As DataSet
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = Sql_Transaction.Connection.CreateCommand
        Lcommand.Transaction = Sql_Transaction
        Dim dtSet As New DataSet
        Dim SqlStr As String = ""
        Try
            Dim sqlda As New SqlDataAdapter(Lcommand)
            Dim mserver As String = Sql_Transaction.Connection.DataSource
            Dim mdatabase As String = Sql_Transaction.Connection.Database
            For i = 0 To aClsTables.Count - 1
                SqlStr = SqlStr & " " & GetDataQuery(aClsTables(i), "")
                If i = 0 Then
                    sqlda.TableMappings.Add("Table", AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & aClsTables(i).TableName)
                Else
                    sqlda.TableMappings.Add("Table" & CStr(i), AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & aClsTables(i).TableName)
                End If
                '    Dim aa As New ColorScheme.ColorScheme
                Dim mPrimaryKey As String = aClsTables(i).PrimaryKey
                If mPrimaryKey.Trim.Length > 0 Then
                    SetPrimaryColumns(dtSet.Tables(i), mPrimaryKey)
                Else
                    sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
                End If
            Next
            Lcommand.CommandText = SqlStr
            GlobalControl.Variables.ErrorString = SqlStr
            sqlda.Fill(dtSet)
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtSet
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="aClsTables" >An array of tableclass objetcs</param>
    ''' <returns>return a data set  object </returns>
    ''' <remarks></remarks>
    ''' 
    Public Function SqlExecuteDataSet(ByVal aClsTables() As Object) As DataSet
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        '  Dim aa As New ColorFontScheme.ColorFontScheme
        Dim mserverdatabase As String = aClsTables(0).ServerDataBase
        Dim LserverDatabase As String = GetServerDataBase(mserverdatabase)
        Dim dtSet As New DataSet
        Dim sqlcmd As New SqlCommand
        Dim SqlStr As String = ""
        Try
            Dim sqlcon As SqlConnection = OpenSqlConnection(LserverDatabase)
            Dim sqlda As New SqlDataAdapter(sqlcmd)
            For i = 0 To aClsTables.Count - 1
                Dim mserver As String = aClsTables(0).Server
                Dim mdatabase As String = aClsTables(0).DataBase
                SqlStr = SqlStr & " " & GetDataQuery(aClsTables(i), "")
                If i = 0 Then
                    sqlda.TableMappings.Add("Table", AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & aClsTables(i).TableName)
                Else
                    sqlda.TableMappings.Add("Table" & CStr(i), AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & aClsTables(i).TableName)
                End If
                '    Dim aa As New ColorScheme.ColorScheme
                Dim mPrimaryKey As String = aClsTables(i).PrimaryKey
                If mPrimaryKey.Trim.Length > 0 Then
                    SetPrimaryColumns(dtSet.Tables(i), mPrimaryKey)
                Else
                    sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
                End If
            Next
            GlobalControl.Variables.ErrorString = SqlStr
            sqlcmd.CommandText = SqlStr
            sqlcmd.Connection = sqlcon
            GlobalControl.Variables.ErrorString = SqlStr
            sqlda.Fill(dtSet)
        Catch ex As Exception
            GF1.QuitError(ex, Err, SqlStr)
        End Try
        Return dtSet
        sqlcmd.Dispose()
    End Function



    ''' <summary>
    ''' Execute T-SQL  query  statement
    ''' </summary>
    ''' <param name="LConnection">SqlConnection </param>
    ''' <param name="SqlStr">Querry Strig command to be executed</param>
    ''' <param name="FromClauseTables" >Comma separated string of tables used in from clause(s) of sqlstr query,Set as TableName property  of DataTable(server.database.dbo.table)</param>
    ''' <param name="PrimaryColumns" >~  separated string of columns(Pcol1,Pcol2 etc.) will be used as primarykeys of datatables ,These columns may be different sql tables primary keys,If not defined, Sql tables primary keys will be DataTables primary key</param>
    ''' <returns>return a data table object </returns>
    ''' <remarks></remarks>
    ''' 
    Public Function SqlExecuteDataSet(ByRef LConnection As SqlConnection, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lcommand As SqlCommand = LConnection.CreateCommand
        Dim dtSet As New DataSet
        Try
            Lcommand.CommandText = SqlStr
            GlobalControl.Variables.ErrorString = SqlStr
            Dim sqlda As New SqlDataAdapter(Lcommand)
            Dim n As Int16 = 0, m As Int16 = 0
            If FromClauseTables.Trim.Length > 0 Then
                Dim mserver As String = LConnection.DataSource
                Dim mdatabase As String = LConnection.Database
                Dim atables() As String = Split(FromClauseTables, ",")
                For i = 0 To atables.Count - 1
                    If i = 0 Then
                        sqlda.TableMappings.Add("Table", AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & atables(i))
                    Else
                        sqlda.TableMappings.Add("Table" & CStr(i), AddSquareBrackets(mserver) & "." & AddSquareBrackets(mdatabase) & ".dbo." & atables(i))
                    End If
                    n = i
                Next
            End If
            If PrimaryColumns.Trim.Length > 0 Then
                Dim mpcols() As String = PrimaryColumns.Split("~")
                For i = 0 To mpcols.Count - 1
                    SetPrimaryColumns(dtSet.Tables(i), mpcols(i))
                    m = i
                Next
            Else
                sqlda.MissingSchemaAction = MissingSchemaAction.AddWithKey
            End If
            If n <> m Then
                QuitMessage("Each table does not primary columns definition " & vbCrLf & SqlStr, "SqlExecuteDataSet(ByVal ServerDatabase As String, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet")
            End If
            sqlda.Fill(dtSet)
        Catch ex As Exception
            GF1.QuitError(ex, Err, "SqlExecuteDataSet(ByRef LConnection As SqlConnection, ByVal SqlStr As String, Optional ByVal FromClauseTables As String = "", Optional ByVal PrimaryColumns As String = "") As DataSet" & SqlStr)
        End Try
        Return dtSet
        Lcommand.Dispose()
    End Function
    ''' <summary>
    ''' Create SQL Table 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table Name to be created</param>
    ''' <param name="mfields">String of fields eg "LastName varchar(255),FirstName varchar(255),Address varchar(255),City varchar(255)" </param>
    ''' <remarks></remarks>

    Public Sub CreateTableByFields(ByVal ServerDataBase As String, ByVal TableName As String, ByVal mfields As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase)
        Try
            SqlExecuteNonQuery(Lserverdatabase, "Create table " & TableName & " (" & mfields & ")")
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateTableByFields(ByVal ServerDataBase As String, ByVal TableName As String, ByVal mfields As String)")
        End Try

    End Sub
    ''' <summary>
    ''' Create SQL Table 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table Name to be created</param>
    ''' <param name="StructureTable" >A datatable containing the values of structure columns as rows</param>
    ''' <remarks></remarks>

    Public Sub CreateTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal StructureTable As DataTable)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        'StrictureColumnsString  "TableName,PrimaryKey,FieldName,FieldType,Nullable,Default"
        Try
            Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase)
            Dim mfields As String = ""
            Dim mstru() As String = Split(GlobalControl.Variables.StructureColumnsString, ",")
            Dim fldDefault As String = ""
            Dim fldNullAble As String = "NULL"
            Dim primarykeycols As String = ""
            For i = 0 To StructureTable.Rows.Count - 1
                Dim primarykeyflag As String = UCase(StructureTable.Rows(i).Item(mstru(1)).ToString.Trim)
                Dim fldnamevalue As String = StructureTable.Rows(i).Item(mstru(2)).ToString.Trim
                fldnamevalue = Replace(fldnamevalue, " ", "")
                Dim fldtypevalue As String = StructureTable.Rows(i).Item(mstru(3)).ToString.Trim
                If mstru.Length > 3 Then
                    fldNullAble = StructureTable.Rows(i).Item(mstru(4)).ToString.Trim
                    If UCase(Left(fldNullAble, 1)) = "N" Then
                        fldNullAble = "NOT NULL"
                    Else
                        fldNullAble = "NULL"
                    End If
                End If
                If mstru.Length > 4 Then
                    fldDefault = StructureTable.Rows(i).Item(mstru(5)).ToString.Trim
                    If fldDefault.Length > 0 Then
                        If InStr(LCase(fldtypevalue), "char") > 0 Then
                            If AscW(fldDefault) = 34 Then
                                fldDefault = Replace(fldDefault, ChrW(34), "")
                            End If
                            fldDefault = "'" & fldDefault & "'"
                        End If
                    End If
                End If
                mfields = mfields & IIf(mfields.Length = 0, "", ",") & fldnamevalue & "  " & fldtypevalue & " " & fldNullAble & " " & IIf(fldDefault.Length > 0, " DEFAULT " & fldDefault, "")
                primarykeycols = primarykeycols & IIf(InStr(primarykeyflag, "Y") > 0, IIf(primarykeycols.Length = 0, "", ",") & fldnamevalue & " ASC ", "")
            Next
            mfields = mfields & IIf(primarykeycols.Length > 0, " ,PRIMARY KEY CLUSTERED (" & primarykeycols & ")", "")
            SqlExecuteNonQuery(Lserverdatabase, "Create table " & TableName & " (" & mfields & ")")
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal StructureTable As DataTable)")
        End Try
    End Sub
    ''' <summary>
    ''' Create SQL Table by excel structure worksheet 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table Name to be created</param>
    ''' <param name="StructureExcelFile" >Full path of excel work sheet containing the whole schema as rows</param>
    ''' <remarks></remarks>

    Public Sub CreateTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal StructureExcelFile As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase)
            Dim mfields As String = ""
            Dim mstru() As String = Split(GlobalControl.Variables.StructureColumnsString, ",")
            Dim lwhere As String = mstru(0).ToString.Trim & " = '" & LCase(TableName).Trim & "'"
            Dim StructureTable As DataTable = GetDataFromExcel(StructureExcelFile, lwhere)
            CreateTable(ServerDataBase, TableName, StructureTable)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal StructureExcelFile As String)")
        End Try

    End Sub
    ''' <summary>
    ''' Create a new datatable whose structure is in a datatable
    ''' </summary>
    ''' <param name="StructureTable">Structure table as datatable</param>
    ''' <param name="StructureColumns">Comma separated names of columns which are corresponding to  Field_Name and Field_Type and PrimaryKeyFlag field </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateDataTable(ByVal StructureTable As DataTable, ByVal StructureColumns As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mdatatable As New DataTable
        Try
            Dim mstru() As String = Split(StructureColumns, ",")
            If mstru.Count < 2 Then
                GF1.QuitMessage("Invalid structure fields", "CreateDataTable(ByVal StructureTable As DataTable, ByVal StructureColumns As String) As DataTable  ")
            End If
            Dim fldPrimKey() As DataColumn = {}

            For i = 0 To StructureTable.Rows.Count - 1
                Dim fldnamevalue As String = StructureTable.Rows(i).Item(mstru(0)).ToString.Trim
                fldnamevalue = Replace(fldnamevalue, " ", "")
                Dim fldtypevalue As String = StructureTable.Rows(i).Item(mstru(1)).ToString.Trim
                Dim fldPrimKeyflg As String = "N"
                If mstru.Count > 2 Then
                    fldPrimKeyflg = StructureTable.Rows(i).Item(mstru(2)).ToString.Trim
                End If
                ' Dim fldsizevalue As String = CInt(StructureTable.Rows(i).Item(mstru(2)).ToString.Trim).ToString
                'Dim flddecvalue As String = CInt(StructureTable.Rows(i).Item(mstru(3)).ToString.Trim).ToString
                Dim mtype1 As String = ""
                Select Case LCase(fldtypevalue)
                    Case "nchar"
                        mtype1 = "System.String"
                    Case "nvarchar"
                        mtype1 = "System.String"
                    Case "int"
                        mtype1 = "Integer"
                    Case "tinyint"
                        mtype1 = "System.Int16"
                    Case "Smallint"
                        mtype1 = "System.Int32"
                    Case "decimal", "double", "single", "numeric"
                        mtype1 = "System.Decimal"
                    Case "datetime"
                        mtype1 = "System.DateTime"
                    Case "bit"
                        mtype1 = "System.Boolean"
                    Case Else
                        mtype1 = "System.Object"
                End Select
                mdatatable.Columns.Add(fldnamevalue, System.Type.GetType(mtype1))
                If fldPrimKeyflg = "Y" Then
                    GF1.ArrayAppend(fldPrimKey, mdatatable.Columns(fldnamevalue))
                End If
            Next
            If fldPrimKey.Count > 0 Then
                mdatatable.PrimaryKey = fldPrimKey
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.Function CreateDataTable(ByVal StructureTable As DataTable, ByVal StructureColumns As String) As DataTable")
        End Try
        Return mdatatable
    End Function



    ''' <summary>
    ''' Alter fields of  SQL Table 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table Name to be altered</param>
    ''' <param name="FieldsToAdd" >String of fields to add  eg "LastName varchar(255),FirstName varchar(255),Address varchar(255),City varchar(255)"</param>
    ''' <param name="FieldsToRemove" >String of fields eg "LastName ,FirstName,Address,City"</param>
    ''' <param name="FieldsToModify" >String of modifying fields eg "LastName varchar(255),FirstName varchar(255),Address varchar(255),City varchar(255)"</param>
    ''' <remarks></remarks>

    Public Sub AlterSQLTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal FieldsToAdd As String, Optional ByVal FieldsToRemove As String = "", Optional ByVal FieldsToModify As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase)
            Dim QryStr As String = "Alter Table " & TableName
            If FieldsToAdd.Trim.Length > 0 Then
                QryStr = QryStr & " Add " & FieldsToAdd.Trim
            End If
            If FieldsToRemove.Trim.Length > 0 Then
                QryStr = QryStr & " Drop Column " & FieldsToRemove.Trim
            End If
            If FieldsToRemove.Trim.Length > 0 Then
                QryStr = QryStr & " Alter Column " & FieldsToRemove.Trim
            End If

            SqlExecuteNonQuery(Lserverdatabase, QryStr)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AlterSQLTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal FieldsToAdd As String, Optional ByVal FieldsToRemove As String = "", Optional ByVal FieldsToModify As String = "")")
        End Try

    End Sub
    ''' <summary>
    ''' Alter (Add/Remove) columns in a datatable. 
    ''' </summary>
    ''' <param name="LDataTable" >DataTable name to be modified</param>
    ''' <param name="ColumnsToAdd" >Comma separated ColumnNames  as string to be added</param>
    ''' <param name="ColumnsToRemove" >Comma separated ColumnNames  as string to be removed</param>
    ''' <remarks></remarks>

    Public Function AlterDataTable(ByVal LDataTable As DataTable, ByVal ColumnsToAdd As String, Optional ByVal ColumnsToRemove As String = "", Optional ByVal OnlyStructure As Boolean = False) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim dtTemp As New DataTable
        If OnlyStructure = False Then
            dtTemp = LDataTable.Copy
        Else
            dtTemp = LDataTable.Clone
        End If
        Try
            If ColumnsToAdd.Trim.Length > 0 Then
                AddColumnsInDataTable(dtTemp, ColumnsToAdd)
            End If
            If ColumnsToRemove.Trim.Length > 0 Then
                dtTemp.PrimaryKey = Nothing
                Dim acol() As String = Split(ColumnsToRemove, ",")
                For i = 0 To acol.Count - 1
                    dtTemp.Columns.Remove(acol(i))
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AlterDataTable(ByVal LDataTable As DataTable, ByVal ColumnsToAdd As String, Optional ByVal ColumnsToRemove As String = "", Optional ByVal OnlyStructure As Boolean = False) As DataTable")
        End Try
        Return dtTemp
    End Function
    ''' <summary>
    ''' Shrink a datatable for given columns. 
    ''' </summary>
    ''' <param name="LDataTable" >DataTable name to be shrinked</param>
    ''' <param name="RemainingColumns" >Comma separated ColumnNames  as string to be remained</param>
    ''' <remarks></remarks>

    Public Function ShrinkDataTable(ByVal LDataTable As DataTable, ByVal RemainingColumns As String, Optional ByVal OnlyStructure As Boolean = False) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If RemainingColumns.Length = 0 Then
            Return LDataTable
            Exit Function
        End If
        Dim DtTemp As New DataTable
        If OnlyStructure = False Then
            DtTemp = LDataTable.Copy
        Else
            DtTemp = LDataTable.Clone
        End If
        Try
            Dim ColumnsToRemove As String = ""
            Dim aCols() As String = Split(LCase(RemainingColumns), ",")
            For i = 0 To DtTemp.Columns.Count - 1
                Dim mcol As String = LCase(DtTemp.Columns(i).ColumnName)
                If GF1.ArrayFind(aCols, mcol) < 0 Then
                    ColumnsToRemove = ColumnsToRemove & IIf(ColumnsToRemove.Length = 0, "", ",") & mcol
                End If
            Next
            If ColumnsToRemove.Trim.Length > 0 Then
                DtTemp.PrimaryKey = Nothing
                Dim acol() As String = Split(ColumnsToRemove, ",")
                For i = 0 To acol.Count - 1
                    DtTemp.Columns.Remove(acol(i))
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ShrinkDataTable(ByVal LDataTable As DataTable, ByVal RemainingColumns As String, Optional ByVal OnlyStructure As Boolean = False) As DataTable")
        End Try
        Return DtTemp
    End Function

    ''' <summary>
    ''' Shrink an SQL Table for remaining fields. 
    ''' </summary>
    ''' <param name="ServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table Name to be altered</param>
    ''' <remarks></remarks>

    Public Sub ShrinkSqlTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal RemainingFields As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        If RemainingFields.Trim.Length = 0 Then
            Exit Sub
        End If
        Try
            Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase)
            Dim dtSchema As DataTable = GetSchemaTable(Lserverdatabase, TableName)
            Dim aFields() As String = Split(RemainingFields, ",")
            Dim ToRemoveFields As String = ""
            For i = 0 To dtSchema.Rows.Count - 1
                Dim mfields As String = dtSchema(i).Item("columnname")
                If GF1.ArrayFind(aFields, mfields) < 0 Then
                    RemainingFields = RemainingFields & IIf(RemainingFields.Length = 0, "", ",") & mfields
                End If
            Next
            If RemainingFields.Length > 0 Then
                Dim QryStr As String = "Alter Table " & TableName
                QryStr = QryStr & " Drop Column " & RemainingFields.Trim
                SqlExecuteNonQuery(Lserverdatabase, QryStr)
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ShrinkSqlTable(ByVal ServerDataBase As String, ByVal TableName As String, ByVal RemainingFields As String)")
        End Try

    End Sub

    '''' <summary>
    '''' Create SQL Table 
    '''' </summary>
    '''' <param name="SqServer">Sql Server name</param>
    '''' <param name="SqDataBase" >Sql Database name</param>
    '''' <param name="TableName">Table Name to be created</param>
    '''' <param name="mfields">String of fields eg "LastName varchar(255),FirstName varchar(255),Address varchar(255),City varchar(255)" </param>
    '''' <remarks></remarks>
    'Public Sub CreateTable(ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String, ByVal mfields As String)
    '    If GlobalControl.Variables.AuthenticationChecked = False Then Exit Sub
    '    If GlobalControl.Variables.EventLogger = True Then GF1.WriteEventLogger(System.Reflection.MethodBase.GetCurrentMethod, True, GF1.GetParametersValueLine(SqServer, SqDataBase, TableName, mfields))
    '    'Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase, TableName, "")
    '    Try
    '        SqlExecuteNonQuery(SqServer, SqDataBase, "Create table " & TableName & " (" & mfields & ")")
    '    Catch ex As Exception
    '        GF1.QuitError(ex, Err)
    '    End Try
    'End Sub
    Private Function ColumnMaxLenght(ByVal DTable As DataTable, ByVal DCol As DataColumn, ByVal StartRowNo As Integer) As Integer
        Dim MaxLen As Integer = 0
        Try
            For j = StartRowNo + 1 To DTable.Rows.Count - 1
                MaxLen = IIf(MaxLen < DTable.Rows(j).Item(DCol.ColumnName).ToString.Length, DTable.Rows(j).Item(DCol.ColumnName).ToString.Length, MaxLen)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ColumnMaxLenght(ByVal DTable As DataTable, ByVal DCol As DataColumn, ByVal StartRowNo As Integer) As Integer")
        End Try

        Return MaxLen
    End Function
    ''' <summary>
    ''' Create SQL Table From Excel taking first row as header fields
    ''' </summary>
    ''' <param name="sqServer">SqL server name</param>
    ''' <param name="SqFolder">Folder where data base created if not exists in server</param>
    ''' <param name="SqDataBase">Sql DataBase Name</param>
    ''' <param name="TableName">Name of Table created</param>
    ''' <param name="ExcelFilePath">Excel file with full path</param>
    ''' <param name="RefreshTable">Remove Table </param>
    ''' <param name="WhereClause">Searching clause on excel file</param>
    ''' <param name="OrderClause">Orderby clause on excel file</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function CreateSQLTableFromExcel(ByVal sqServer As String, ByVal SqFolder As String, ByVal SqDataBase As String, ByVal TableName As String, ByVal ExcelFilePath As String, Optional ByVal RefreshTable As Boolean = True, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "") As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase, TableName, "")
        Try
            Dim Dtx As DataTable = GetDataFromExcel(ExcelFilePath, WhereClause, OrderClause)
            Dim mfld As String = ""
            'To hold valid column index converted into fields
            Dim FieldColumnIndex As Integer() = {}
            For i = 0 To Dtx.Columns.Count - 1
                Dim mname As String = Dtx.Columns(i).ColumnName.Trim
                If mname.Length = 0 Or mname = "F" & (i + 1).ToString Then
                    Continue For
                End If
                Dim mtype As String = ""
                Dim rr As Integer = 0
                For r = 0 To Dtx.Rows.Count - 1
                    mtype = LCase(Dtx.Rows(r)(i).GetType.Name.ToString)
                    If mtype = "byte" Or mtype = "string" Or mtype = "int32" Or mtype = "int64" Or mtype = "integer" Or mtype = "date" Or mtype = "datetime" Or mtype = "decimal" Or mtype = "double" Or mtype = "single" Then
                        rr = r
                        Exit For
                    End If
                Next
                If mtype = "dbnull" Then
                    Continue For
                End If
                Dim mtype1 As String = ""
                Select Case LCase(mtype)
                    Case "string"
                        Dim len As Integer = ColumnMaxLenght(Dtx, Dtx.Columns(i), rr)
                        mtype1 = "nchar(" & len & ") Null "
                    Case "integer", "int32", "int64", "int16", "byte"
                        mtype1 = "int Null "
                    Case "decimal", "double", "single"
                        mtype1 = "numeric(18,4) Null "
                    Case "datetime", "date"
                        mtype1 = "datetime Null"
                    Case Else
                        Continue For
                        'mtype1 = "nchar(" & len & ")"
                End Select
                GF1.ArrayAppend(FieldColumnIndex, i)
                mfld = mfld & IIf(mfld.Length = 0, "", ",") & mname & " " & mtype1 & " null "
            Next
            If Not DataBaseExists(sqServer, SqDataBase) Then
                CreateDataBase(sqServer, SqFolder, SqDataBase)
            End If

            If TableExists(sqServer, SqDataBase, TableName) Then
                If RefreshTable = True Then
                    DropTable(sqServer, SqDataBase, TableName)
                Else
                    GF1.QuitMessage("Table " & TableName & "Already exists nothing done", " CreateSQLTableFromExcel(ByVal sqServer As String, ByVal SqFolder As String, ByVal SqDataBase As String, ByVal TableName As String, ByVal ExcelFilePath As String, Optional ByVal RefreshTable As Boolean = True, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "") As Boolean ")
                    Return False
                    Exit Function
                End If
            End If
            Dim Mcons As SqlConnection = OpenSqlConnection(sqServer, SqDataBase)
            SqlExecuteNonQuery(Mcons, "Create table " & TableName & "(" & mfld & ")")
            Dim SqlCmd As New SqlCommand
            SqlCmd.Connection = Mcons
            Dim MyStr As String = ""
            For j = 0 To Dtx.Rows.Count - 1
                Dim xstring As String = ""
                For x = 0 To FieldColumnIndex.Count - 1
                    Dim mtype As String = LCase(Dtx.Rows(j)(x).GetType.Name.ToString)
                    Select Case mtype
                        Case "string", "datetime", "date"
                            xstring = xstring & Dtx.Rows(j).Item(x)
                        Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                            If Dtx.Rows(j).Item(x) <> 0 Then
                                xstring = xstring & Dtx.Rows(j).Item(x).ToString
                            End If
                    End Select
                Next
                If xstring.Trim.Length = 0 Then
                    Continue For
                End If

                MyStr = "Insert Into " & TableName & " Values ( "
                For i = 0 To FieldColumnIndex.Count - 1
                    Dim k As Integer = FieldColumnIndex(i)
                    'If Dtx.Rows(j).Item(k).ToString <> "" Then
                    Dtx.Rows(j).Item(k) = Dtx.Rows(j).Item(k).ToString.Replace("'", "")
                    Select Case LCase(Dtx.Rows(j).Item(k).GetType.Name.ToString)
                        Case "string", "datetime"
                            MyStr = MyStr & "'" & Dtx.Rows(j).Item(k).ToString.Trim & "',"
                        Case Else
                            MyStr = MyStr & Dtx.Rows(j).Item(k).ToString.Trim & ","
                    End Select
                    'End If
                Next
                MyStr = MyStr.Remove(MyStr.Length - 1, 1)
                MyStr = MyStr & ")"
                SqlExecuteNonQuery(Mcons, SqlCmd, MyStr)
            Next
            SqlCmd.Dispose()
            Mcons.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateSQLTableFromExcel(ByVal sqServer As String, ByVal SqFolder As String, ByVal SqDataBase As String, ByVal TableName As String, ByVal ExcelFilePath As String, Optional ByVal RefreshTable As Boolean = True, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "") As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' Create SQL Table From Excel taking first row as header fields
    ''' </summary>
    ''' <param name=" ServerDataBase ">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name=" TableName" >Name of Table created</param>
    ''' <param name="ExcelFilePath" >Excel file with full path</param>
    ''' <param name="RefreshTable">Remove Table </param>
    ''' <param name="WhereClause">Searching clause on excel file</param>
    ''' <param name="OrderClause">Orderby clause on excel file</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function CreateSQLTableFromExcel(ByVal ServerDataBase As String, ByVal TableName As String, ByVal ExcelFilePath As String, Optional ByVal RefreshTable As Boolean = True, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "") As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase, TableName, "")
        Try
            Dim Dtx As DataTable = GetDataFromExcel(ExcelFilePath, WhereClause, OrderClause)
            Dim mfld As String = ""
            'To hold valid column index converted into fields
            Dim FieldColumnIndex As Integer() = {}
            Dim FieldColumnIndexType As String() = {}
            For i = 0 To Dtx.Columns.Count - 1
                Dim mname As String = Dtx.Columns(i).ColumnName.Trim
                If mname.Length = 0 Or mname = "F" & (i + 1).ToString Then
                    Continue For
                End If
                Dim mtype As String = ""
                Dim rr As Integer = 0
                For r = 0 To Dtx.Rows.Count - 1
                    mtype = LCase(Dtx.Rows(r)(i).GetType.Name.ToString)
                    If mtype = "byte" Or mtype = "string" Or mtype = "int32" Or mtype = "int64" Or mtype = "integer" Or mtype = "date" Or mtype = "datetime" Or mtype = "decimal" Or mtype = "double" Or mtype = "single" Then
                        rr = r
                        Exit For
                    End If
                Next
                If mtype = "dbnull" Then
                    Continue For
                End If
                Dim mtype1 As String = ""
                Select Case LCase(mtype)
                    Case "string"
                        Dim len As Integer = ColumnMaxLenght(Dtx, Dtx.Columns(i), rr)
                        mtype1 = "nchar(" & len & ")"
                    Case "integer", "int32", "int64", "int16"
                        mtype1 = "int"
                    Case "byte"
                        mtype1 = "TinyInt"
                    Case "decimal", "double", "single"
                        mtype1 = "numeric(18,4)"
                    Case "datetime", "date"
                        mtype1 = "datetime"
                    Case Else
                        Continue For
                        'mtype1 = "nchar(" & len & ")"
                End Select
                GF1.ArrayAppend(FieldColumnIndex, i)
                GF1.ArrayAppend(FieldColumnIndexType, mtype)
                mfld = mfld & IIf(mfld.Length = 0, "", ",") & mname & " " & mtype1 & " null "
            Next
            Dim LserverDatabase As String = GetServerDataBase(ServerDataBase)
            If TableExists(LserverDatabase, TableName) Then
                If RefreshTable = True Then
                    DropTable(ServerDataBase, TableName)
                Else
                    GF1.QuitMessage("Table " & TableName & "Already exists nothing done", "CreateSQLTableFromExcel(ByVal ServerDataBase As String, ByVal TableName As String, ByVal ExcelFilePath As String, Optional ByVal RefreshTable As Boolean = True, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "") As Boolean  ")
                    Return False
                    Exit Function
                End If
            End If
            ' Dim aa As GlobalControl .Variables. .item("mdf0")

            Dim Mcons As SqlConnection = OpenSqlConnection(LserverDatabase)
            SqlExecuteNonQuery(Mcons, "Create table " & TableName & "(" & mfld & ")")
            Dim SqlCmd As New SqlCommand
            SqlCmd.Connection = Mcons
            Dim MyStr As String = ""
            For j = 0 To Dtx.Rows.Count - 1
                Dim xstring As String = ""
                For x = 0 To FieldColumnIndex.Count - 1
                    Dim mtype As String = LCase(Dtx.Rows(j)(x).GetType.Name.ToString)
                    Select Case mtype
                        Case "string", "datetime", "date"
                            xstring = xstring & Dtx.Rows(j).Item(x).ToString.Trim
                        Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                            If Dtx.Rows(j).Item(x) <> 0 Then
                                xstring = xstring & Dtx.Rows(j).Item(x).ToString.Trim
                            End If
                    End Select
                Next
                If xstring.Trim.Length = 0 Then
                    Continue For
                End If
                MyStr = "Insert Into " & TableName & " Values ( "
                For i = 0 To FieldColumnIndex.Count - 1
                    Dim k As Integer = FieldColumnIndex(i)
                    Dim CellValue As String = Dtx.Rows(j).Item(k).ToString.Trim
                    Dim ColFieldType As String = FieldColumnIndexType(i)
                    Select Case ColFieldType
                        Case "string", "datetime"
                            If IsDBNull(Dtx.Rows(j).Item(k)) = True Then
                                CellValue = ""
                            End If
                            CellValue = CellValue.ToString.Replace("'", "")
                            MyStr = MyStr & "'" & CellValue & "',"
                        Case Else
                            If IsDBNull(Dtx.Rows(j).Item(k)) = True Then
                                CellValue = "0"
                            End If
                            If CellValue.Length = 0 Then
                                CellValue = "0"
                            End If
                            MyStr = MyStr & CellValue & ","
                    End Select
                    'End If
                Next
                MyStr = MyStr.Remove(MyStr.Length - 1, 1)
                MyStr = MyStr & ")"
                SqlExecuteNonQuery(Mcons, SqlCmd, MyStr)
            Next
            SqlCmd.Dispose()
            Mcons.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateSQLTableFromExcel(ByVal ServerDataBase As String, ByVal TableName As String, ByVal ExcelFilePath As String, Optional ByVal RefreshTable As Boolean = True, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "") As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' Create SQL Table
    ''' </summary>
    ''' <param name="LTransaction">Transaction object passed by reference</param>
    ''' <param name="TableName">Table Name to be created</param>
    ''' <param name="mfields">String of fields eg "LastName varchar(255),FirstName varchar(255),Address varchar(255),City varchar(255)"</param>
    ''' <remarks></remarks>
    Public Sub CreateTable(ByRef LTransaction As SqlTransaction, ByVal TableName As String, ByVal mfields As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        'Dim Lserverdatabase As String = GetServerDataBase(ServerDataBase, TableName, "")
        Try
            SqlExecuteNonQuery(LTransaction, "Create table " & TableName & " (" & mfields & ")")
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateTable(ByRef LTransaction As SqlTransaction, ByVal TableName As String, ByVal mfields As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Remove table from sql database
    ''' </summary>
    ''' <param name="SqServer"  >Sql Server name</param>
    '''  <param name="SqDataBase" >Sql Database name</param>
    ''' <param name="TableName">Table name to be remove</param>
    ''' <remarks></remarks>
    Public Sub DropTable(ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            SqlExecuteNonQuery(SqServer, SqDataBase, "Drop Table " & TableName)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.DropTable(ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Remove SQL Table
    ''' </summary>
    ''' <param name="LTransaction">Transaction object passed by reference</param>
    ''' <param name="TableName">Table Name to be removed</param>
    ''' <remarks></remarks>
    Public Sub DropTable(ByRef LTransaction As SqlTransaction, ByVal TableName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            SqlExecuteNonQuery(LTransaction, "Drop Table " & TableName)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.DropTable(ByRef LTransaction As SqlTransaction, ByVal TableName As String)")
        End Try
    End Sub

    ''' <summary>
    ''' Remove table from sql database
    ''' </summary>
    ''' <param name="ServerDatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table name to be removed</param>
    ''' <remarks></remarks>
    Public Sub DropTable(ByVal ServerDatabase As String, ByVal TableName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim Lserverdatabase As String = GetServerDataBase(ServerDatabase)
        Try
            SqlExecuteNonQuery(Lserverdatabase, "Drop Table " & TableName)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.DropTable(ByVal ServerDatabase As String, ByVal TableName As String)")
        End Try
    End Sub

    ''' <summary>
    ''' To remove all rows from a sql table
    ''' </summary>
    ''' <param name="ServerDatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Rows of Table name to  be removed</param>
    ''' <param name="WhereClause">Filter condition on sql table</param>
    ''' <remarks></remarks>
    Public Sub TruncateTable(ByVal ServerDatabase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim Lserverdatabase As String = GetServerDataBase(ServerDatabase)
        Try
            If WhereClause.Length > 0 Then
                SqlExecuteNonQuery(Lserverdatabase, "delete from " & TableName & " where " & WhereClause)
            Else
                SqlExecuteNonQuery(Lserverdatabase, "Truncate Table " & TableName)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.TruncateTable(ByVal ServerDatabase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "")")
        End Try
    End Sub
    ''' <summary>
    ''' To remove all rows from a sql table
    ''' </summary>
    ''' <param name="LTransaction">Transaction object passed by reference</param>
    ''' <param name="TableName">Rows of Table name to  be removed</param>
    ''' <param name="WhereClause">Filter condition on sql table</param>
    ''' <remarks></remarks>
    Public Sub TruncateTable(ByRef LTransaction As SqlTransaction, ByVal TableName As String, Optional ByVal WhereClause As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            If WhereClause.Length > 0 Then
                SqlExecuteNonQuery(LTransaction, "delete from " & TableName & " where " & WhereClause)
            Else
                SqlExecuteNonQuery(LTransaction, "truncate table " & TableName)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.TruncateTable(ByRef LTransaction As SqlTransaction, ByVal TableName As String, Optional ByVal WhereClause As String = "")")
        End Try
    End Sub
    ''' <summary>
    ''' To Rename sql table of one database
    ''' </summary>
    ''' <param name="ServerDatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="OldTableName ">Old table name to  be renamed</param>
    ''' <param name="NewTableName ">New Table Name</param>
    ''' <remarks></remarks>
    Public Sub RenameTable(ByVal ServerDatabase As String, ByVal OldTableName As String, ByVal NewTableName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim Lserverdatabase As String = GetServerDataBase(ServerDatabase)
        OldTableName = ConvertFromSrv0Mdf0(OldTableName)
        NewTableName = ConvertFromSrv0Mdf0(NewTableName)
        Try
            SqlExecuteNonQuery(Lserverdatabase, "sp_rename '" & OldTableName & "', '" & NewTableName & "'")
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RenameTable(ByVal ServerDatabase As String, ByVal OldTableName As String, ByVal NewTableName As String)")
        End Try
    End Sub
    ''' <summary>
    ''' To Rename sql table of one database
    ''' </summary>
    ''' <param name="LTransaction">Transaction object passed by reference</param>
    ''' <param name="OldTableName">Old table name to  be renamed</param>
    ''' <param name="NewTableName">New Table Name</param>
    ''' <remarks></remarks>
    Public Sub RenameTable(ByRef LTransaction As SqlTransaction, ByVal OldTableName As String, ByVal NewTableName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            SqlExecuteNonQuery(LTransaction, "sp_rename '" & OldTableName & "', '" & NewTableName & "'")
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RenameTable(ByRef LTransaction As SqlTransaction, ByVal OldTableName As String, ByVal NewTableName As String)")
        End Try
    End Sub

    ''' <summary>
    ''' To remove SQL database 
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <remarks></remarks>
    Public Sub DropDataBase(ByVal ServerDataBase As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim LdataBase As String = BreakServerDataBase(ServerDataBase).Item(1)
        Dim Masterbase As String = BreakServerDataBase(ServerDataBase).Item(0) & ".[master]"
        Try
            SqlExecuteNonQuery(Masterbase, "Drop DataBase " & LdataBase)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.DropDataBase(ByVal ServerDataBase As String)")
        End Try
    End Sub
    ''' <summary>
    ''' TO Remove SQL Database
    ''' </summary>
    ''' <param name="ServerName">Server Name</param>
    ''' <param name="LDataBase">DataBase Name to be removed</param>
    ''' <remarks></remarks>
    Public Sub DropDataBase(ByVal ServerName As String, ByVal LDataBase As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim Masterbase As String = ServerName & ".master"
        Try
            SqlExecuteNonQuery(Masterbase, "Drop DataBase " & LDataBase)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.DropDataBase(ByVal ServerName As String, ByVal LDataBase As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Create a table by copying another table of same or distinct database/server
    ''' </summary>
    '''  <param name="LTransaction">Transaction object passed by reference</param>
    ''' <param name="To_Table">Target table name</param>
    ''' <param name="From_Table">Source table name</param>
    ''' <remarks></remarks>
    Public Sub CopySqlTable(ByRef LTransaction As SqlTransaction, ByVal To_Table As String, ByVal From_Table As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim str As String = "select * into " & To_Table & " from " & From_Table
            SqlExecuteNonQuery(LTransaction, str)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CopySqlTable(ByRef LTransaction As SqlTransaction, ByVal To_Table As String, ByVal From_Table As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Create a table by copying another table of same or distinct database/server
    ''' </summary>
    ''' <param name="ServerDataBase ">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="To_Table">Target table name</param>
    ''' <param name="From_Table">Source table name</param>
    ''' <remarks></remarks>
    Public Sub CopySqlTable(ByVal ServerDataBase As String, ByVal To_Table As String, ByVal From_Table As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim LServerDatabase As String = GetServerDataBase(ServerDataBase)
        To_Table = ConvertFromSrv0Mdf0(To_Table)
        From_Table = ConvertFromSrv0Mdf0(From_Table)
        Try
            Dim str As String = "select * into " & To_Table & " from " & From_Table
            SqlExecuteNonQuery(LServerDatabase, str)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CopySqlTable(ByVal ServerDataBase As String, ByVal To_Table As String, ByVal From_Table As String)")
        End Try
    End Sub

    ''' <summary>
    ''' To count total no. of rows  
    ''' </summary>
    ''' <param name="Serverdatabase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table name</param>
    ''' <param name="lCondition">Where clause of querry</param>
    ''' <param name="lorder">Order clause of querry</param>
    ''' <returns>Total nos. of rows</returns>
    ''' <remarks></remarks>
    Public Function RowsCount(ByVal Serverdatabase As String, ByVal TableName As String, ByVal lCondition As String, ByVal lorder As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim sqlval As Integer = 0
        Try
            Dim LserverDataBase As String = GetServerDataBase(Serverdatabase)
            TableName = ConvertFromSrv0Mdf0(TableName)
            Dim str1 As String = ""
            Select Case True
                Case lorder.Trim.Length = 0 And lCondition.Trim.Length = 0
                    str1 = "select sum(row_count) from sys.dm_db_partition_stats where object_id=object_id('" & TableName.Trim & "')  AND (index_id=0 or index_id=1) "
                Case lorder.Trim.Length = 0
                    str1 = "select count(*) from " & TableName & IIf(lCondition.Trim.Length = 0, "", " where " & lCondition)
                Case Else
                    str1 = "Select  ROW_NUMBER() OVER(ORDER BY " & lorder & ") As 'RowNumber'  from " & TableName & " As m1" & IIf(lCondition.Trim.Length = 0, "", " where " & lCondition) & " order by RowNumber Desc"
            End Select
            sqlval = SqlExecuteScalarQuery(LserverDataBase, str1)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RowsCount(ByVal Serverdatabase As String, ByVal TableName As String, ByVal lCondition As String, ByVal lorder As String) As Integer")
        End Try
        Return sqlval
    End Function
    ''' <summary>
    ''' To count total no. of rows
    ''' </summary>
    ''' <param name="LTransaction">Transaction object passed by reference</param>
    ''' <param name="TableName">Table name</param>
    ''' <param name="lCondition">Where clause of querry</param>
    ''' <param name="lorder">Order clause of querry</param>
    ''' <returns>Total nos. of rows</returns>
    ''' <remarks></remarks>
    Public Function RowsCount(ByRef LTransaction As SqlTransaction, ByVal TableName As String, ByVal lCondition As String, ByVal lorder As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim sqlval As Integer = 0
        Try
            TableName = ConvertFromSrv0Mdf0(TableName)
            Dim str1 As String = ""
            Select Case True
                Case lorder.Trim.Length = 0 And lCondition.Trim.Length = 0
                    str1 = "select sum(row_count) from sys.dm_db_partition_stats where object_id=object_id('" & TableName.Trim & "')  AND (index_id=0 or index_id=1) "
                Case lorder.Trim.Length = 0
                    str1 = "select count(*) from " & TableName & IIf(lCondition.Trim.Length = 0, "", " where " & lCondition)
                Case Else
                    str1 = "Select Top(1) ROW_NUMBER() OVER(ORDER BY " & lorder & ") As 'RowNumber'  from " & TableName & " As m1" & IIf(lCondition.Trim.Length = 0, "", " where " & lCondition) & " order by RowNumber Desc"
            End Select
            sqlval = SqlExecuteScalarQuery(LTransaction, str1)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RowsCount(ByRef LTransaction As SqlTransaction, ByVal TableName As String, ByVal lCondition As String, ByVal lorder As String) As Integer")
        End Try
        Return sqlval
    End Function
    ''' <summary>
    ''' To count total no. of rows
    ''' </summary>
    ''' <param name="LSqlConnection">SqlConnection object passed by reference</param>
    ''' <param name="TableName">Table name</param>
    ''' <param name="lCondition">Where clause of querry</param>
    ''' <param name="lorder">Order clause of querry</param>
    ''' <returns>Total nos. of rows</returns>
    ''' <remarks></remarks>
    Public Function RowsCount(ByRef LSqlConnection As SqlConnection, ByVal TableName As String, ByVal lCondition As String, ByVal lorder As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim sqlval As Integer = 0
        Try
            TableName = ConvertFromSrv0Mdf0(TableName)
            Dim str1 As String = ""
            Select Case True
                Case lorder.Trim.Length = 0 And lCondition.Trim.Length = 0
                    str1 = "select sum(row_count) from sys.dm_db_partition_stats where object_id=object_id('" & TableName.Trim & "')  AND (index_id=0 or index_id=1) "
                Case lorder.Trim.Length = 0
                    str1 = "select count(*) from " & TableName & IIf(lCondition.Trim.Length = 0, "", " where " & lCondition)
                Case Else
                    str1 = "Select Top(1) ROW_NUMBER() OVER(ORDER BY " & lorder & ") As 'RowNumber'  from " & TableName & " As m1" & IIf(lCondition.Trim.Length = 0, "", " where " & lCondition) & " order by RowNumber Desc"
            End Select
            sqlval = SqlExecuteScalarQuery(LSqlConnection, str1)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RowsCount(ByRef LTransaction As SqlTransaction, ByVal TableName As String, ByVal lCondition As String, ByVal lorder As String) As Integer")
        End Try
        Return sqlval
    End Function


    ''' <summary>
    ''' To get first excel work sheet name of excel file
    ''' </summary>
    ''' <param name="FullExcelFile">Excel file name with path and extension</param>
    ''' <returns>First work shhet name</returns>
    ''' <remarks></remarks>
    Public Function FirstExcelSheetName(ByVal FullExcelFile As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim objExcel As Excel.Application = CreateObject("Excel.Application")
            Dim objWorkBook As Excel.Workbook
            Dim objWorkSheets As Excel.Worksheet
            objWorkBook = objExcel.Workbooks.Open(FullExcelFile)
            For Each objWorkSheets In objWorkBook.Worksheets
                Dim str As String = objWorkSheets.Name
                objExcel.Workbooks.Close()
                Return str
                Exit Function
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FirstExcelSheetName(ByVal FullExcelFile As String) As String")
        End Try
        Return "sheet1"
    End Function
    ''' <summary>
    ''' To get physical path of a database in a sql server
    ''' </summary>
    ''' <param name="ServerName">Server name as string</param>
    ''' <param name="DatabaseName">Data base name to be searched</param>
    ''' <returns>Full path of database name</returns>
    ''' <remarks></remarks>
    Public Function GetDataBasePath(ByVal ServerName As String, ByVal DatabaseName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DBPath As String = ""
        Try
            ServerName = ConvertFromSrv0(ServerName)
            DatabaseName = ConvertFromMdf0(DatabaseName)
            Dim con As SqlConnection = OpenSqlConnection(ServerName, DatabaseName)
            Dim cmd As SqlCommand = New SqlCommand("SELECT physical_name AS current_file_location FROM sys.master_files WHERE name = '" & DatabaseName & "'", con)
            GlobalControl.Variables.ErrorString = cmd.CommandText
            Dim sdr As SqlDataReader = cmd.ExecuteReader()
            If (sdr.Read()) Then
                DBPath = sdr(0).ToString()
            End If
            sdr.Close()
            cmd.Dispose()
            con.Close()
            Return DBPath
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return DBPath
    End Function
    ''' <summary>
    ''' To get physical path of a database in a sql server
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <returns>Physical path of database</returns>
    ''' <remarks></remarks>
    Public Function GetSqlDataBasePath(ByVal ServerDataBase As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DBPath As String = ""
        Try
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            Dim AserverDataBase As List(Of String) = BreakServerDataBase(LServerDataBase)
            Dim con As SqlConnection = OpenSqlConnection(LServerDataBase)
            Dim cmd As SqlCommand = New SqlCommand("SELECT physical_name AS current_file_location FROM sys.master_files WHERE name = '" & AserverDataBase(1) & "'", con)
            GlobalControl.Variables.ErrorString = cmd.CommandText
            Dim sdr As SqlDataReader = cmd.ExecuteReader()
            If (sdr.Read()) Then
                DBPath = sdr(0).ToString()
            End If
            sdr.Close()
            cmd.Dispose()
            con.Close()
        Catch ex As Exception
            QuitError(ex, Err, GlobalControl.Variables.ErrorString)
        End Try
        Return DBPath
    End Function
    ''' <summary>
    ''' To add excel rows to an existing SQL table whoose columns are same as fields of sql table
    ''' </summary>
    ''' <param name="FullExcelFile">Full path of excel table</param>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TableName">Table name on which excel rows are added</param>
    ''' <param name="WhereClause" >Where clause to filter excel table</param>
    ''' <param name="ExcelColumns" >Identify excel columns by "name" or "index"  </param>
    ''' <param name="RemoveALLRows" >If True, remove all existing rows in target table</param>
    ''' <returns>Execution flag</returns>
    ''' <remarks></remarks>
    Public Function AppendExcelToSQL(ByVal FullExcelFile As String, ByVal ServerDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "Index", Optional ByVal RemoveAllRows As Boolean = True) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim LserverDataBase As String = GetServerDataBase(ServerDataBase)
            TableName = ConvertFromSrv0Mdf0(TableName)
            Dim ExcelRows As DataTable = GetDataFromExcel(FullExcelFile, WhereClause)
            Dim Lcons As SqlConnection = OpenSqlConnection(LserverDataBase)
            Dim MyTrans As SqlTransaction = Lcons.BeginTransaction(IsolationLevel.ReadCommitted, "MyTransaction")
            Dim Lcmd0 As SqlCommand = Lcons.CreateCommand()
            Lcmd0.Connection = Lcons
            Lcmd0.Transaction = MyTrans
            Dim mPrimaryKey As String = ""
            Dim LSchema As DataTable = GetSchemaTable(ServerDataBase, TableName)
            ExcelRows.CaseSensitive = False
            LSchema.CaseSensitive = False
            Try
                If RemoveAllRows = True Then
                    TruncateTable(MyTrans, TableName, WhereClause)
                End If
                InsertExcelRowsToSql(MyTrans, ExcelRows, LSchema, TableName, ExcelColumns)
                MyTrans.Commit()
            Catch
                MyTrans.Rollback()
            End Try
            LSchema.Dispose()
            ExcelRows.Dispose()
            Lcons.Dispose()
            Lcmd0.Dispose()
            MyTrans.Dispose()
            Return True
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AppendExcelToSQL(ByVal FullExcelFile As String, ByVal ServerDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "", Optional ByVal RemoveAllRows As Boolean = True) As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' To add excel rows to an existing SQL table whoose columns are same as fields of sql table
    ''' </summary>
    ''' <param name="FullExcelFile">Full path of excel table</param>
    ''' <param name="SqServer" >Sql server name</param>
    ''' <param name="SqDatabase">Database name</param>
    ''' <param name="TableName">Table name on which excel rows are added</param>
    ''' <param name="WhereClause" >Where clause to filter excel table</param>
    ''' <param name="ExcelColumns" >Mapping of  excel columns and sql table fields by "name" or "index"</param>
    ''' <param name="RemoveALLRows" >If True, remove all existing rows in target table</param>
    ''' <returns>Execution flag</returns>
    ''' <remarks></remarks>
    Public Function InserExcelToSql(ByVal FullExcelFile As String, ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "Index", Optional ByVal RemoveAllRows As Boolean = True) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ExcelRows As DataTable = GetDataFromExcel(FullExcelFile)
            Dim Lcons As SqlConnection = OpenSqlConnection(SqServer, SqDataBase)
            Dim LSchema As DataTable = GetSchemaTable(SqServer, SqDataBase, TableName)
            ExcelRows.CaseSensitive = False
            LSchema.CaseSensitive = False
            Dim Lcmd0 As SqlCommand = Lcons.CreateCommand()
            Dim LTrans As SqlTransaction = Lcons.BeginTransaction(IsolationLevel.ReadCommitted, "MyTrans")
            Lcmd0.Transaction = LTrans
            Try
                If RemoveAllRows = True Then
                    TruncateTable(LTrans, TableName, WhereClause)
                End If
                InsertExcelRowsToSql(LTrans, ExcelRows, LSchema, TableName, ExcelColumns)
                LTrans.Commit()
            Catch
                LTrans.Rollback()
            End Try
            LSchema.Dispose()
            ExcelRows.Dispose()
            Lcons.Dispose()
            Lcmd0.Dispose()
            LTrans.Dispose()
            Return True
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InserExcelToSql(ByVal FullExcelFile As String, ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "", Optional ByVal RemoveAllRows As Boolean = True) As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' To add a compatible string data template to an existing SQL table. </summary>
    ''' <param name="DataStringTemplate">String data template columns are separated by | and rows by chr(13)</param>
    ''' <param name="ServerDatabase " >Sql server name</param>
    ''' <param name="TableName">Table name on which excel rows are added</param>
    ''' <param name="RemoveALLRows" >If True, remove all existing rows in target table</param>
    ''' <returns>Execution flag</returns>
    ''' <remarks></remarks>
    Public Function InsertStringDataToSql(ByVal DataStringTemplate As String, ByVal ServerDatabase As String, ByVal TableName As String, Optional ByVal RemoveAllRows As Boolean = True, Optional ByVal ColSeparator As String = "|", Optional ByVal RowSeparator As String = Chr(13)) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ExcelRows As DataTable = StringToDataTable(DataStringTemplate, ColSeparator, RowSeparator)
            Dim Lcons As SqlConnection = OpenSqlConnection(ServerDatabase)
            Dim LSchema As DataTable = GetSchemaTable(ServerDatabase, TableName)
            ExcelRows.CaseSensitive = False
            LSchema.CaseSensitive = False
            Dim Lcmd0 As SqlCommand = Lcons.CreateCommand()
            Dim LTrans As SqlTransaction = Lcons.BeginTransaction(IsolationLevel.ReadCommitted, "MyTrans")
            Lcmd0.Transaction = LTrans
            Try
                If RemoveAllRows = True Then
                    TruncateTable(LTrans, TableName)
                End If
                InsertExcelRowsToSql(LTrans, ExcelRows, LSchema, TableName, "index")
                LTrans.Commit()
            Catch
                LTrans.Rollback()
            End Try
            LSchema.Dispose()
            ExcelRows.Dispose()
            Lcons.Dispose()
            Lcmd0.Dispose()
            LTrans.Dispose()
            Return True
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InserExcelToSql(ByVal FullExcelFile As String, ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "", Optional ByVal RemoveAllRows As Boolean = True) As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' To insert a CSV file into  SQL table. </summary>
    ''' <param name="CSVFileName">CSV File Name columns are separated by , and rows by chr(13)</param>
    ''' <param name="ServerDatabase " >Sql server name</param>
    ''' <param name="TableName">Table name on which excel rows are added</param>
    ''' <param name="StartRowNo" >Start row no of  CSV file from which data to be added i.e. rowno excluding header</param>
    ''' <param name="RemoveALLRows" >If True, remove all existing rows in target table</param>
    ''' <param name="ColSeparator" >Column separator of CSV/TXT file</param>
    ''' <param name="RowSeparator" >Row separator of CSV/TXT file</param>
    ''' <returns>Execution flag</returns>
    ''' <remarks></remarks>
    Public Function InsertCSVFileToSql(ByVal CSVFileName As String, ByVal ServerDatabase As String, ByVal TableName As String, Optional ByVal StartRowNo As Integer = 1, Optional ByVal RemoveAllRows As Boolean = True, Optional ByVal ColSeparator As String = ",", Optional ByVal RowSeparator As String = vbCrLf) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            '  Dim ExcelRows As DataTable = StringToDataTable(DataStringTemplate, ColSeparator, RowSeparator)
            Dim Lcons As SqlConnection = OpenSqlConnection(ServerDatabase)
            Dim LSchema As DataTable = GetSchemaTable(ServerDatabase, TableName)
            '      ExcelRows.CaseSensitive = False
            LSchema.CaseSensitive = False
            Dim Lcmd0 As SqlCommand = Lcons.CreateCommand()
            Dim LTrans As SqlTransaction = Lcons.BeginTransaction(IsolationLevel.ReadCommitted, "MyTrans")
            Lcmd0.Transaction = LTrans
            Try
                If RemoveAllRows = True Then
                    TruncateTable(LTrans, TableName)
                End If
                Dim fs As System.IO.FileStream = New System.IO.FileStream(CSVFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Dim sr As System.IO.StreamReader
                sr = New System.IO.StreamReader(fs, System.Text.Encoding.Default)
                sr.Peek()
                Dim ii As Integer = 0
                Do While sr.Peek() >= 0
                    Dim mstr As String = ""
                    If RowSeparator = vbCrLf Then
                        mstr = sr.ReadLine
                        ii = ii + 1
                        If ii < StartRowNo Then
                            Continue Do
                        End If
                    Else
                        Do While ChrW(sr.Read) <> RowSeparator And sr.Peek >= 0
                            mstr = mstr & ChrW(sr.Read)
                        Loop
                    End If
                    If mstr.Length > 0 Then
                        Dim mdt As DataTable = StringToDataTable(mstr, ColSeparator, RowSeparator)
                        InsertExcelRowsToSql(LTrans, mdt, LSchema, TableName, "index")
                    End If
                Loop
                sr.Close()
                fs.Close()
                LTrans.Commit()
            Catch
                LTrans.Rollback()
            End Try
            LSchema.Dispose()
            Lcons.Dispose()
            Lcmd0.Dispose()
            LTrans.Dispose()
            Return True
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertCSVFileToSql(ByVal CSVFileName As String, ByVal ServerDatabase As String, ByVal TableName As String, Optional ByVal StartRowNo As Integer = 1, Optional ByVal RemoveAllRows As Boolean = True, Optional ByVal ColSeparator As String = [, ]")
        End Try
    End Function
    ''' <summary>
    ''' To get data table from a CSV file . </summary>
    ''' <param name="CSVFileName">CSV File Name columns are separated by , and rows by chr(13)</param>
    ''' <param name="StartRowNo" >Start row no of  CSV file from which data to be added i.e. rowno excluding header</param>
    ''' <param name="ColSeparator" >Column separator of CSV/TXT file</param>
    ''' <param name="RowSeparator" >Row separator of CSV/TXT file</param>
    ''' <returns>Execution flag</returns>
    ''' <remarks></remarks>
    Public Function GetDataFromCSV(ByVal CSVFileName As String, Optional ByVal StartRowNo As Integer = 1, Optional ByVal ColSeparator As String = ",", Optional ByVal RowSeparator As String = vbCrLf) As DataTable
        If System.IO.File.Exists(CSVFileName) = False Then
            QuitMessage(CSVFileName & " not found", "")
        End If
        Dim mdt As New DataTable
        Try
            Dim fs As System.IO.FileStream = New System.IO.FileStream(CSVFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(fs, System.Text.Encoding.Default)
            Dim ii As Integer = 0
            Do While sr.Peek() >= 0
                Dim mstr As String = ""
                If RowSeparator = vbCrLf Then
                    mstr = sr.ReadLine
                    ii = ii + 1
                    If ii < StartRowNo Then
                        Continue Do
                    End If
                Else
                    Do While ChrW(sr.Read) <> RowSeparator And sr.Peek >= 0
                        mstr = mstr & ChrW(sr.Read)
                    Loop
                End If
                If mstr.Length > 0 Then
                    mdt = StringToDataTable(mstr, ColSeparator, RowSeparator, mdt)
                End If
            Loop
            sr.Close()
            fs.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Public Function GetDataFromCSV(ByVal CSVFileName As String, Optional ByVal StartRowNo As Integer = 1, Optional ByVal ColSeparator As String = Optional ByVal RowSeparator As String = vbCrLf) As DataTable")
        End Try
        Return mdt
    End Function
    ''' <summary>
    ''' Private function to insert a range of rows of datatable of excel work sheet into sql table
    ''' </summary>
    ''' <param name="LTransaction">Sql Transaction</param>
    ''' <param name="ExcelRows">DataTable object of excel worksheet</param>
    ''' <param name="Lschema">DataTable of Schema of SQL Table to be inserted </param>
    ''' <param name="TableName" >Table name into rows will be inserted</param>
    ''' <param name="ExcelCoumns">Mapping of  excel columns and sql table fields by "name" or "index"</param>
    ''' <remarks></remarks>
    Private Sub InsertExcelRowsToSql(ByRef LTransaction As SqlTransaction, ByVal ExcelRows As DataTable, ByVal Lschema As DataTable, ByVal TableName As String, ByVal ExcelCoumns As String)
        If GlobalControl.Variables.AuthenticationChecked = False Then Exit Sub
        Dim Lcommand As SqlCommand = LTransaction.Connection.CreateCommand
        Lcommand.Transaction = LTransaction
        Dim errorstr As String = ""

        Try
            For erow = 0 To ExcelRows.Rows.Count - 1
                Try
                    Dim FieldStr As String = ""
                    Dim ValueStr As String = ""
                    Dim aSqlParameters() As SqlParameter = {}
                    ReDim aSqlParameters(-1)
                    For Each lrow As DataRow In Lschema.Rows
                        Try
                            Dim mtype As String = LCase(lrow.Item("datatypename"))
                            Dim mfield As String = LCase(lrow.Item("columnname"))

                            'Dim Etype As String = LCase(ExcelRows.Rows(erow).Item(mfield).GetType.Name)
                            Dim Evalue As String = ""
                            Dim colfind As Boolean = True
                            Dim mfieldno As Integer = CInt(lrow.Item("columnordinal"))
                            If LCase(ExcelCoumns) = "name" Then
                                If CheckColumnInDataTable(mfield, ExcelRows) > -1 Then
                                    If IsDBNull(ExcelRows.Rows(erow).Item(mfield)) = False Then
                                        Evalue = ExcelRows.Rows(erow).Item(mfield).ToString
                                    End If
                                Else
                                    colfind = False
                                End If
                            Else
                                If mfieldno < ExcelRows.Columns.Count Then
                                    If IsDBNull(ExcelRows.Rows(erow).Item(mfieldno)) = False Then
                                        Evalue = ExcelRows.Rows(erow).Item(mfieldno).ToString
                                    End If
                                Else
                                    colfind = False
                                End If
                            End If
                            Select Case mtype
                                Case "int", "numeric", "decimal", "smallint", "tinyint"
                                    If Evalue.Length = 0 Then
                                        Evalue = "0"
                                    End If
                                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                                Case "bit"
                                    If Evalue.Length = 0 Then
                                        Evalue = "0"
                                    End If
                                    Evalue = LCase(Evalue)
                                    If Evalue = "true" Or Evalue = "yes" Or Evalue = "y" Then
                                        Evalue = "1"
                                    Else
                                        Evalue = "0"
                                    End If
                                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                                Case "datetime"
                                    Dim kk As New SqlParameter
                                    Dim dtvalue As DateTime = New Date(1900, 1, 1)
                                    kk.ParameterName = "@" & mfield
                                    If colfind = True Then
                                        If IsDBNull(ExcelRows.Rows(erow).Item(mfieldno)) = False Then
                                            'dtvalue = ExcelRows.Rows(erow).Item(mfieldno)
                                            If LCase(ExcelRows.Rows(erow).Item(mfieldno).GetType.Name) = "string" Then
                                                Evalue = ExcelRows.Rows(erow).Item(mfieldno).ToString
                                                Select Case Evalue.Trim.Length
                                                    Case 10
                                                        dtvalue = New Date(CInt(Right(Evalue, 4)), CInt(Left(Evalue, 2)), CInt(Mid(Evalue, 4, 2)))
                                                    Case 8
                                                        dtvalue = New Date(CInt(Left(Evalue, 4)), CInt(Mid(Evalue, 5, 2)), CInt(Right(Evalue, 2)))
                                                    Case Else
                                                        dtvalue = New Date(2000, 1, 1)
                                                End Select
                                            End If
                                        End If
                                    End If
                                    kk.SqlValue = dtvalue
                                    kk.SqlDbType = SqlDbType.DateTime
                                    ArrayAppendSqlParameter(aSqlParameters, kk)
                                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & kk.ParameterName
                                Case "nchar", "nvarchar"
                                    Evalue = Evalue.Replace("'", "")
                                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                                Case Else
                                    Evalue = Evalue.Replace("'", "")
                                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                            End Select
                        Catch ex As Exception
                            QuitError(ex, Err, "Error in " & FieldStr & " valuestr " & ValueStr & " in adding fields in InsertExcelRowsToSql(ByRef LTransaction As SqlTransaction, ByVal ExcelRows As DataTable, ByVal Lschema As DataTable, ByVal TableName As String, ByVal ExcelCoumns As String)")
                        End Try
                    Next
                    Dim QyrStr As String = "Insert  into  " & TableName & " (" & FieldStr & ") values (" & ValueStr & ")"
                    GlobalControl.Variables.ErrorString = QyrStr
                    Lcommand.Parameters.Clear()
                    Lcommand.Parameters.AddRange(aSqlParameters)
                    Lcommand.CommandText = QyrStr
                    Lcommand.ExecuteNonQuery()
                Catch ex As Exception
                    QuitError(ex, Err, GlobalControl.Variables.ErrorString)
                End Try
            Next

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InserExcelToSql(ByVal FullExcelFile As String, ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "", Optional ByVal RemoveAllRows As Boolean = True) As Boolean")
        End Try
        Lcommand.Dispose()

    End Sub
    ''' <summary>
    ''' To add/update rows from source sql table to target sql table for common field names,primary key names are same for both tables.  
    ''' </summary>
    ''' <param name="LocalServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTable">Source table name or full identifier table name</param>
    ''' <param name="TargetServerDataBase" >Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="TargetTable">Target table name or full identifier table name</param>
    ''' <param name="WhereClause" >Filter condition on source table</param>
    ''' <param name="OrderClause" >Order by Clause on source table</param>
    ''' <param name="ReplaceFlag">If True, update row if it exists in target table</param>
    ''' <param name="RemoveALLRows" >If True, remove all existing rows in target table</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AppendSQLToSQL(ByVal LocalServerDataBase As String, ByVal SourceTable As String, ByVal TargetServerDataBase As String, ByVal TargetTable As String, Optional ByVal WhereClause As String = "", Optional ByVal OrderClause As String = "", Optional ByVal ReplaceFlag As Boolean = True, Optional ByVal RemoveALLRows As Boolean = True) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim SserverDataBase As String = GetServerDataBase(LocalServerDataBase, True)
            SourceTable = ConvertFromSrv0Mdf0(SourceTable)
            TargetTable = ConvertFromSrv0Mdf0(TargetTable)
            Dim TserverDataBase As String = SserverDataBase
            If TargetServerDataBase.Trim.Length > 0 Then
                TserverDataBase = GetServerDataBase(TargetServerDataBase, True)
                TargetTable = TserverDataBase & ".DBO." & TargetTable
            End If
          
            Dim TotalRows As Integer = TotalRows = RowsCount(SserverDataBase, SourceTable, WhereClause, OrderClause)
            Dim TargetSchema As DataTable = GetSchemaTable(TserverDataBase, TargetTable)
            Dim mPrimaryKey As String = GetPrimaryKey(TargetSchema)

            Dim mOrderClause As String = IIf(OrderClause.Length = 0, mPrimaryKey, OrderClause)
            Dim SqlStr As String = "select *,row_number() over (order by " & mOrderClause & ") as RowSno from " & SourceTable
            Dim Lcons As SqlConnection = OpenSqlConnection(SserverDataBase)
            Dim LTrans As SqlTransaction = Lcons.BeginTransaction(IsolationLevel.ReadCommitted, "MyTtrans")
            Dim Lcmd0 As SqlCommand = Lcons.CreateCommand()
            Lcmd0.Transaction = LTrans
            If RemoveALLRows = True Then
                TruncateTable(LTrans, TargetTable, WhereClause)
            End If
            For i = 1 To TotalRows Step 100
                Dim mSqlStr As String = SqlStr & " where " & "RowSno between " & i.ToString & " and " & (i + 100).ToString
                If WhereClause.Length > 0 Then
                    mSqlStr = mSqlStr & " and " & WhereClause
                End If
                If OrderClause.Length > 0 Then
                    mSqlStr = mSqlStr & " order by " & OrderClause
                End If
                Dim DtSource As DataTable = SqlExecuteDataTable(SserverDataBase, SqlStr)
                InsertDataTableToSql(LTrans, Lcons, Lcmd0, DtSource, TargetSchema, TargetTable, ReplaceFlag, mPrimaryKey)
            Next
            LTrans.Commit()
            TargetSchema.Dispose()
            Lcons.Dispose()
            Lcmd0.Dispose()
            Return True
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InserExcelToSql(ByVal FullExcelFile As String, ByVal SqServer As String, ByVal SqDataBase As String, ByVal TableName As String, Optional ByVal WhereClause As String = "", Optional ByVal ExcelColumns As String = "", Optional ByVal RemoveAllRows As Boolean = True) As Boolean")
        End Try
    End Function
    ''' <summary>
    ''' To insert rows of a datatable into sql table
    ''' </summary>
    ''' <param name="LTransaction">Sql Transaction</param>
    ''' <param name="Lcons">Opened Sql Connection</param>
    ''' <param name="lcmd0">Command object</param>
    ''' <param name="LDtSource">DataTable object of source table</param>
    ''' <param name="LTargetSchema">DataTable of Schema of Target Table</param>
    ''' <param name="TableName">Target Table name </param>
    ''' <param name="ReplaceFlag">If True, update row if it exists in target table</param>
    ''' <param name="TargetPrimaryKey">Primary key of Target table</param>
    ''' <remarks></remarks>
    Public Sub InsertDataTableToSql(ByRef LTransaction As SqlTransaction, ByRef Lcons As SqlConnection, ByRef lcmd0 As SqlCommand, ByVal LDtSource As DataTable, ByVal LTargetSchema As DataTable, ByVal TableName As String, ByVal ReplaceFlag As Boolean, ByVal TargetPrimaryKey As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        'InsertRowsToTable(Lcons, lcmd0, DtSource, TargetSchema, TargetTable)
        Try

            For erow = 0 To LDtSource.Rows.Count - 1
                Dim FieldStr As String = ""
                Dim ValueStr As String = ""
                Dim mPrimaryValue As String = ""
                Dim UpdateStr As String = ""
                Dim QryStr As String = ""
                Try

                    For Each lrow As DataRow In LTargetSchema.Rows
                        lcmd0.Parameters.Clear()
                        Dim mtype As String = LCase(lrow.Item("datatypename"))
                        Dim mfield As String = LCase(lrow.Item("columnname"))
                        'Dim Etype As String = LCase(ExcelRows.Rows(erow).Item(mfield).GetType.Name)
                        Dim FieldFound As Boolean = False
                        For Each scolumn As DataColumn In LDtSource.Columns
                            If mfield = LCase(scolumn.ColumnName) Then
                                FieldFound = True
                                Exit For
                            End If
                        Next
                        Dim Evalue As String = ""
                        Dim ParaName As String = ""
                        If FieldFound = True Then
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mfield & " = "
                            'cmd = New SqlCeCommand("insert into Receiving (Id,PalletNo,Batch,Run,PCode,Qty,AddUser,AddDate) VALUES (@Id,@PalletNo,@Batch,@Run,@PCode,@Qty,@AddUser,@AddDate)", cn)
                            'cmd.Parameters.AddWithValue("@PalletNo", txtPallet.Text)
                            'cmd.Parameters.AddWithValue("@Batch", Trim(txtBatch.Text))
                            'cmd.Parameters.AddWithValue("@Run", Trim(txtRun.Text))
                            'cmd.Parameters.AddWithValue("@PCode", Trim(txtPCode.Text))
                            'cmd.Parameters.AddWithValue("@Qty", Val(txtQty.Text))
                            'cmd.Parameters.AddWithValue("@AddUser", txtEmpNo.Text)
                            'cmd.Parameters.AddWithValue("@Id", (ID + 1))
                            'cmd.Parameters.AddWithValue("@AddDate", Format(Now, "dd-mmm-yyyy"))
                            'cmd.ExecuteNonQuery()
                            If LCase(mfield) = LCase(TargetPrimaryKey) Then
                                mPrimaryValue = LDtSource.Rows(erow).Item(mfield)
                            End If
                            Select Case mtype
                                'int=10digits,smallint=32767,tinyint=255
                                Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                                    If IsDBNull(LDtSource.Rows(erow).Item(mfield)) = False Then
                                        Evalue = LDtSource.Rows(erow).Item(mfield).ToString
                                    Else
                                        Evalue = "0"
                                    End If
                                    If Evalue.Length = 0 Then
                                        Evalue = "0"
                                    End If
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                                    UpdateStr = UpdateStr & Evalue
                                Case "bit"
                                    Evalue = "0"
                                    If IsDBNull(LDtSource.Rows(erow).Item(mfield)) = False Then
                                        Evalue = LCase(LDtSource.Rows(erow).Item(mfield).ToString)
                                    End If
                                    If Evalue.Length = 0 Then
                                        Evalue = "0"
                                    End If
                                    If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                                        Evalue = "1"
                                    Else
                                        Evalue = "0"
                                    End If
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                                    UpdateStr = UpdateStr & Evalue
                                Case "image"
                                    Dim ImageBytes() As Byte = LDtSource.Rows(erow).Item(mfield)
                                    ParaName = "@" & mfield
                                    lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                                    UpdateStr = UpdateStr & ParaName
                                Case "datetime"
                                    Dim MdateTime As DateTime = LDtSource.Rows(erow).Item(mfield)
                                    ParaName = "@" & mfield
                                    lcmd0.Parameters.AddWithValue(ParaName, MdateTime)
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                                    UpdateStr = UpdateStr & ParaName
                                Case "nchar", "nvarchar"
                                    If IsDBNull(LDtSource.Rows(erow).Item(mfield)) = False Then
                                        Evalue = LDtSource.Rows(erow).Item(mfield).ToString
                                    Else
                                        Evalue = ""
                                    End If
                                    Evalue = Evalue.Replace("'", "")
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                                    UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                                Case Else
                                    If IsDBNull(LDtSource.Rows(erow).Item(mfield)) = False Then
                                        Evalue = LDtSource.Rows(erow).Item(mfield).ToString
                                    Else
                                        Evalue = ""
                                    End If
                                    Evalue = Evalue.Replace("'", " ")
                                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                                    UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                            End Select
                        End If
                    Next
                    Dim RowFound As DataRow = SeekRecord(LTransaction, TableName, TargetPrimaryKey, mPrimaryValue)
                    If RowFound IsNot Nothing Then
                        If ReplaceFlag = True Then
                            QryStr = "update  " & TableName & " set " & UpdateStr
                        Else
                            Continue For
                        End If
                    Else
                        QryStr = "Insert  into  " & TableName & " (" & FieldStr & ") values (" & ValueStr & ")"
                    End If
                    lcmd0.CommandText = QryStr
                    lcmd0.ExecuteNonQuery()
                Catch ex As Exception
                    QuitError(ex, Err, QryStr)
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertDataTableToSql(ByRef LTransaction As SqlTransaction, ByRef Lcons As SqlConnection, ByRef lcmd0 As SqlCommand, ByVal LDtSource As DataTable, ByVal LTargetSchema As DataTable, ByVal TableName As String, ByVal ReplaceFlag As Boolean, ByVal TargetPrimaryKey As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Insert Record to an SQL Table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="RowSource">DataRow having the column names as field names of table and values as  field's value  of record </param>
    ''' <param name="TargetTable">Name of table in which record inserted</param>
    ''' <param name="TargetPrimaryKey">Primary key field name of table , (if the key exists, record is updated otherwise inserted)</param>
    ''' <remarks></remarks>
    Public Sub InsertRowToSqlTable(ByVal ServerDataBase As String, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LTargetSchema As DataTable = GetSchemaTable(ServerDataBase, TargetTable)
            Dim SameColumns As New Hashtable
            For Each lrow As DataRow In LTargetSchema.Rows
                Dim mtype As String = LCase(lrow.Item("datatypename")).ToString
                Dim mfield As String = LCase(lrow.Item("columnname")).ToString
                Try
                    Dim aa As Object = RowSource(mfield)
                    SameColumns.Add(LCase(mfield), mtype)
                Catch
                    Continue For
                End Try
            Next
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            Dim mPrimaryValue As String = ""
            Dim UpdateStr As String = ""
            Dim lwhere As String = ""
            For i = 0 To SameColumns.Count - 1
                Dim mfield As String = LCase(SameColumns.Keys(i).ToString)
                Dim mtype As String = SameColumns.Item(mfield).ToString
                If IsDBNull(RowSource(mfield)) = True Then
                    Continue For
                End If
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mfield & " = "
                If LCase(mfield) = LCase(TargetPrimaryKey) Then
                    mPrimaryValue = RowSource(mfield)
                End If
                Select Case mtype
                    'int=10digits,smallint=32767,tinyint=255
                    Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                        If IsDBNull(RowSource(mfield)) = False Then
                            Evalue = RowSource(mfield).ToString.Trim
                        Else
                            Evalue = "0"
                        End If
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        UpdateStr = UpdateStr & Evalue
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If


                    Case "bit"
                        Evalue = "0"
                        If IsDBNull(RowSource(mfield)) = False Then
                            Evalue = LCase(RowSource(mfield).ToString.Trim)
                        End If
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                            Evalue = "1"
                        Else
                            Evalue = "0"
                        End If
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        UpdateStr = UpdateStr & Evalue

                    Case "image"
                        Dim ImageBytes() As Byte = RowSource(mfield)
                        ParaName = "@" & mfield
                        ' lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "datetime"
                        Dim MdateTime As DateTime = RowSource(mfield)
                        ParaName = "@" & mfield
                        'lcmd0.Parameters.AddWithValue(ParaName, MdateTime)
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "nchar", "nvarchar"
                        If IsDBNull(RowSource(mfield)) = False Then
                            Evalue = RowSource(mfield).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                    Case Else
                        If IsDBNull(RowSource(mfield)) = False Then
                            Evalue = RowSource(mfield).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                End Select
            Next
            Dim QryStr As String = ""
            Dim PrimaryKeyValue As New Hashtable
            GF1.AddItemToHashTable(PrimaryKeyValue, TargetPrimaryKey, mPrimaryValue)
            Dim RowFound As DataRow = SeekDataRow(ServerDataBase, TargetTable, PrimaryKeyValue)
            If RowFound IsNot Nothing Then
                QryStr = "update  " & TargetTable & " set " & UpdateStr & " where " & lwhere
            Else
                QryStr = "Insert into  " & TargetTable & " (" & FieldStr & ") values (" & ValueStr & ")"
            End If
            Dim aa1 As Integer = SqlExecuteNonQuery(ServerDataBase, QryStr)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertRowToSqlTable(ByVal ServerDataBase As String, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Insert Record to an SQL Table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="DtSource ">DataTable having the column names as field names of table and values as  field's value  of record </param>
    ''' <param name="TargetTable">Name of table in which record inserted</param>
    ''' <param name="TargetPrimaryKey">Primary key of table , if the key exists, record is updated otherwise inserted</param>
    ''' <remarks></remarks>
    Public Sub InsertDataTableToSql(ByVal ServerDataBase As String, ByVal DtSource As DataTable, ByVal TargetTable As String, ByVal TargetPrimaryKey As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim lcons As SqlConnection = OpenSqlConnection(GetServerDataBase(ServerDataBase))
            TargetTable = ConvertFromSrv0Mdf0(TargetTable)
            ' Dim MyTrans As SqlTransaction = lcons.BeginTransaction(IsolationLevel.ReadCommitted, "SampleTransaction")
            Dim MyTrans As SqlTransaction = lcons.BeginTransaction(IsolationLevel.Serializable, "SampleTransaction")
            Dim samecolumns As New Hashtable
            For i = 0 To DtSource.Rows.Count - 1
                InsertRowToSqlTable(MyTrans, DtSource.Rows(i), TargetTable, TargetPrimaryKey)
            Next
            MyTrans.Commit()
            lcons.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertRowToSqlTable(ByVal ServerDataBase As String, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Insert Record to an SQL Table
    ''' </summary>
    ''' <param name="Ltrans" >sql transaction </param>
    ''' <param name="RowSource">DataRow having the column names as field names of table and values as  field's value  of record </param>
    ''' <param name="TargetTable">Name of table in which record inserted</param>
    ''' <param name="TargetPrimaryKey">Primary key of table , if the key exists, record is updated otherwise inserted</param>
    ''' <remarks></remarks>
    Public Sub InsertRowToSqlTable(ByRef Ltrans As SqlTransaction, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LTargetSchema As DataTable = GetSchemaTable(Ltrans, TargetTable)
            Dim Lcmd0 As SqlCommand = Ltrans.Connection.CreateCommand
            Lcmd0.Transaction = Ltrans
            Dim SameColumns As New Hashtable
            For Each lrow As DataRow In LTargetSchema.Rows
                Dim mtype As String = LCase(lrow.Item("datatypename")).ToString
                Dim mfield As String = LCase(lrow.Item("columnname")).ToString
                Try
                    Dim aa As Object = RowSource(mfield)
                    SameColumns.Add(LCase(mfield), mtype)
                Catch
                    Continue For
                End Try
            Next
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            Dim mPrimaryValue As String = ""
            Dim UpdateStr As String = ""
            Dim lwhere As String = ""
            For i = 0 To SameColumns.Count - 1
                Dim mfield As String = LCase(SameColumns.Keys(i).ToString)
                Dim mtype As String = SameColumns.Item(mfield).ToString
                If IsDBNull(RowSource(mfield)) = True Then
                    Continue For
                End If
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mfield & " = "
                If LCase(mfield) = LCase(TargetPrimaryKey) Then
                    mPrimaryValue = RowSource(mfield)
                End If


                If mtype <> "image" Then
                    Evalue = RowSource(mfield).ToString.Trim
                    Evalue = Evalue.Replace("'", " ")
                End If
                Select Case mtype
                    'int=10digits,smallint=32767,tinyint=255
                    Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                        Evalue = RowSource(mfield).ToString.Trim
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        UpdateStr = UpdateStr & Evalue
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = " & mPrimaryValue
                        End If
                    Case "bit"
                        Evalue = LCase(RowSource(mfield).ToString.Trim)
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                            Evalue = "1"
                        Else
                            Evalue = "0"
                        End If
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        UpdateStr = UpdateStr & Evalue
                    Case "image"
                        Dim ImageBytes() As Byte = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "date"
                        Dim MdateTime As DateTime = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.Date).Value = MdateTime
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "datetime", "time"
                        Dim MdateTime As DateTime = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.DateTime).Value = MdateTime
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "timespan"
                        Dim MTimeSpan As TimeSpan = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.Time).Value = MTimeSpan
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "nchar", "nvarchar"
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                    Case Else
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                End Select
            Next
            Dim QryStr As String = ""
            Dim RowFound As DataRow = SeekRecord(Ltrans, TargetTable, TargetPrimaryKey, mPrimaryValue)
            If RowFound IsNot Nothing Then
                QryStr = "update  " & TargetTable & " set " & UpdateStr & " where " & lwhere
            Else
                QryStr = "Insert  into " & TargetTable & " (" & FieldStr & ") values (" & ValueStr & ")"
            End If
            Dim aa1 As Integer = SqlExecuteNonQuery(Ltrans, Lcmd0, QryStr)
            Lcmd0.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertRowToSqlTable(ByRef Ltrans As SqlTransaction, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String, Optional ByRef SameColumns As Hashtable = Nothing)")
        End Try
    End Sub
    ''' <summary>
    ''' Insert Record to an SQL Table
    ''' </summary>
    ''' <param name="LConnection" >sql transaction </param>
    ''' <param name="RowSource">DataRow having the column names as field names of table and values as  field's value  of record </param>
    ''' <param name="TargetTable">Name of table in which record inserted</param>
    ''' <param name="TargetPrimaryKey">Primary key of table , if the key exists, record is updated otherwise inserted</param>
    ''' <remarks></remarks>
    Public Sub InsertRowToSqlTable(ByRef LConnection As SqlConnection, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LTargetSchema As DataTable = GetSchemaTable(LConnection, TargetTable)
            Dim Lcmd0 As SqlCommand = LConnection.CreateCommand
            Dim SameColumns As New Hashtable
            For Each lrow As DataRow In LTargetSchema.Rows
                Dim mtype As String = LCase(lrow.Item("datatypename")).ToString
                Dim mfield As String = LCase(lrow.Item("columnname")).ToString
                Try
                    Dim aa As Object = RowSource(mfield)
                    SameColumns.Add(LCase(mfield), mtype)
                Catch
                    Continue For
                End Try
            Next
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            Dim mPrimaryValue As String = ""
            Dim UpdateStr As String = ""
            Dim lwhere As String = ""
            For i = 0 To SameColumns.Count - 1
                Dim mfield As String = LCase(SameColumns.Keys(i).ToString)
                Dim mtype As String = SameColumns.Item(mfield).ToString
                If IsDBNull(RowSource(mfield)) = True Then
                    Continue For
                End If
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mfield & " = "
                If LCase(mfield) = LCase(TargetPrimaryKey) Then
                    mPrimaryValue = RowSource(mfield)
                End If


                If mtype <> "image" Then
                    Evalue = RowSource(mfield).ToString.Trim
                    Evalue = Evalue.Replace("'", " ")
                End If
                Select Case mtype
                    'int=10digits,smallint=32767,tinyint=255
                    Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                        Evalue = RowSource(mfield).ToString.Trim
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        UpdateStr = UpdateStr & Evalue
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = " & mPrimaryValue
                        End If
                    Case "bit"
                        Evalue = LCase(RowSource(mfield).ToString.Trim)
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                            Evalue = "1"
                        Else
                            Evalue = "0"
                        End If
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        UpdateStr = UpdateStr & Evalue
                    Case "image"
                        Dim ImageBytes() As Byte = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "date"
                        Dim MdateTime As DateTime = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.Date).Value = MdateTime
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "datetime", "time"
                        Dim MdateTime As DateTime = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.DateTime).Value = MdateTime
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "timespan"
                        Dim MTimeSpan As TimeSpan = RowSource(mfield)
                        ParaName = "@" & mfield
                        Lcmd0.Parameters.Add(ParaName, SqlDbType.Time).Value = MTimeSpan
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                        UpdateStr = UpdateStr & ParaName
                    Case "nchar", "nvarchar"
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                    Case Else
                        ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & "'" & Evalue.Trim & "'"
                        If LCase(mfield) = LCase(TargetPrimaryKey) Then
                            mPrimaryValue = Evalue
                            lwhere = TargetPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                End Select
            Next
            Dim QryStr As String = ""
            Dim RowFound As DataRow = SeekRecord(LConnection, TargetTable, TargetPrimaryKey, mPrimaryValue)
            If RowFound IsNot Nothing Then
                QryStr = "update  " & TargetTable & " set " & UpdateStr & " where " & lwhere
            Else
                QryStr = "Insert  into " & TargetTable & " (" & FieldStr & ") values (" & ValueStr & ")"
            End If
            Dim aa1 As Integer = SqlExecuteNonQuery(LConnection, Lcmd0, QryStr)
            Lcmd0.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertRowToSqlTable(ByRef LConnection As SqlConnection, ByVal RowSource As DataRow, ByVal TargetTable As String, ByVal TargetPrimaryKey As String, Optional ByRef SameColumns As Hashtable = Nothing)")
        End Try
    End Sub

    ''' <summary>
    ''' Insert Records  to  multiple SQL Tables in a single batch execution.
    ''' </summary>
    ''' <param name="Ltrans" >sql transaction </param>
    ''' <param name="TableClassObject">An array of Class objects of tables</param>
    ''' <param name="SqlExecutionFlag" >Sql Execution Flag</param>
    Public Function InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject() As Object, Optional ByVal SqlExecutionFlag As Boolean = True) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return False : Exit Function
        If SqlExecutionFlag = False Then
            Return False
            Exit Function
        End If

        ' Try
        Dim FinalQuery As String = ""
        Dim lCommand As New SqlCommand
        lCommand.Transaction = Ltrans
        lCommand.Parameters.Clear()
        For i = 0 To TableClassObject.Count - 1
            If TableClassObject(i).SqlUpdation = False Then
                Continue For
            End If


            Dim mTable As String = IIf(SameDataSource(Ltrans, TableClassObject(i)) = True, TableClassObject(i).TableName, TableClassObject(i).TableWithSQLPath)
            Dim mRowStatusFlag As Boolean = TableClassObject(i).RowStatusFlag
            Dim mTableEntryType As String = TableClassObject(i).TableEntryType
            Dim mPrimaryKey As String = TableClassObject(i).PrimaryKey
            Dim Dtschema As DataTable = TableClassObject(i).SchemaTable
            Dim mFieldsFinalValues As Hashtable = TableClassObject(i).FieldsFinalValues
            Dim mPreviousExtraRows() As Integer = TableClassObject(i).PreviousExtraRows
            Dim mCurrentExtraRows() As Integer = TableClassObject(i).CurrentExtraRows
            Dim mTableType As String = TableClassObject(i).TableType
            Dim mHeaderRowStatusFlag As Boolean = TableClassObject(i).HeaderRowStatusFlag
            If mTableEntryType = "M" Or mTableEntryType = "D" Then
                For j = 0 To mPreviousExtraRows.Count - 1
                    Dim jj As Integer = mPreviousExtraRows(j)
                    Select Case True
                        Case mRowStatusFlag = True And InStr("H,M", mTableType) > 0
                            Dim mRowStatus As Integer = GF1.GetValueFromHashTable(mFieldsFinalValues, "RowStatus")
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mRowStatus.ToString & " where " & mPrimaryKey & " = " & TableClassObject(i).PrevDt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                        Case mTableType = "S" And mHeaderRowStatusFlag = True
                            Dim mHeaderRowStatusNo As Int16 = TableClassObject(i).HeaderRowStatusNo
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mHeaderRowStatusNo.ToString & " where " & mPrimaryKey & " = " & TableClassObject(i).PrevDt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                        Case Else
                            Dim pKeyValue As Integer = TableClassObject(i).PrevDt.Rows(jj).item(mPrimaryKey)
                            Dim Cind As Integer = FindRowIndexByPrimaryCols(TableClassObject(i).CurrDt, pKeyValue)
                            If Cind > -1 Then
                                Dim mRow As DataRow = TableClassObject(i).CurrDt.Rows(Cind)
                                Dim WhereClause As String = mPrimaryKey & " = " & pKeyValue
                                Dim QueryStr As String = CreateUpdateQuery(lCommand, mRow, mTable, Dtschema, Cind,, WhereClause)   'Changes by Neha
                                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
                            Else
                                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "delete from " & mTable & " where " & mPrimaryKey & " = " & TableClassObject(i).prevdt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                            End If
                    End Select
                Next

            End If
            If mCurrentExtraRows.Count > 0 Then
                Dim mRows() As DataRow = {}
                For j = 0 To mCurrentExtraRows.Count - 1
                    Dim jj As Integer = mCurrentExtraRows(j)
                    Dim mCurrDt As DataTable = TableClassObject(i).CurrDt
                    Dim nRow As DataRow = mCurrDt.Rows(jj)
                    If mRowStatusFlag = True Or (TableClassObject(i).TableType = "S" And TableClassObject(i).HeaderRowStatusFlag = True) Then
                        nRow("RowStatus") = 0
                        GF1.ArrayAppend(mRows, nRow)
                    Else
                        Dim mUpdatedColumnsStr As String = nRow("UpdatedColumnsStr").ToString.Trim
                        If mUpdatedColumnsStr.Length = 0 Then
                            GF1.ArrayAppend(mRows, nRow)
                        End If
                    End If
                Next
                Dim QueryStr As String = ""
                If mRows.Count = 1 Then
                    QueryStr = CreateInsertQuery(lCommand, mRows(0), mTable, Dtschema, i)
                Else
                    QueryStr = CreateInsertQuery(lCommand, mRows, mTable, Dtschema, i)
                End If
                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
            End If
        Next
        Dim a1 As Integer = SqlExecuteNonQuery(Ltrans, lCommand, FinalQuery)
        lCommand.Dispose()
        ' Catch ex As Exception
        'QuitError(ex, Err, "Unable to execute DataFunction.InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject() As Object)")
        'End Try
        Return True
    End Function

    ''' <summary>
    ''' Insert Records  to  multiple SQL Tables in a single batch execution.
    ''' </summary>
    ''' <param name="Ltrans" >sql transaction </param>
    ''' <param name="TableClassObject">An array of Class objects of tables</param>
    ''' <param name="Rowseffected" >Sql Execution Flag</param>
    Public Function InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject() As Object, ByRef Rowseffected As Integer, Optional ByVal SqlExecutionFlag As Boolean = True) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return False : Exit Function
        If SqlExecutionFlag = False Then
            Return False
            Exit Function
        End If
        Rowseffected = 0
        ' Try
        Dim FinalQuery As String = ""
        Dim lCommand As New SqlCommand
        lCommand.Transaction = Ltrans
        lCommand.Parameters.Clear()
        For i = 0 To TableClassObject.Count - 1
            If TableClassObject(i).SqlUpdation = False Then
                Continue For
            End If


            Dim mTable As String = IIf(SameDataSource(Ltrans, TableClassObject(i)) = True, TableClassObject(i).TableName, TableClassObject(i).TableWithSQLPath)
            Dim mRowStatusFlag As Boolean = TableClassObject(i).RowStatusFlag
            Dim mTableEntryType As String = TableClassObject(i).TableEntryType
            Dim mPrimaryKey As String = TableClassObject(i).PrimaryKey
            Dim Dtschema As DataTable = TableClassObject(i).SchemaTable
            Dim mFieldsFinalValues As Hashtable = TableClassObject(i).FieldsFinalValues
            Dim mPreviousExtraRows() As Integer = TableClassObject(i).PreviousExtraRows
            Dim mCurrentExtraRows() As Integer = TableClassObject(i).CurrentExtraRows
            Dim mTableType As String = TableClassObject(i).TableType
            Dim mHeaderRowStatusFlag As Boolean = TableClassObject(i).HeaderRowStatusFlag
            If mTableEntryType = "M" Or mTableEntryType = "D" Then
                For j = 0 To mPreviousExtraRows.Count - 1
                    Dim jj As Integer = mPreviousExtraRows(j)
                    Select Case True
                        Case mRowStatusFlag = True And InStr("H,M", mTableType) > 0
                            Dim mRowStatus As Integer = GF1.GetValueFromHashTable(mFieldsFinalValues, "RowStatus")
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mRowStatus.ToString & " where " & mPrimaryKey & " = " & TableClassObject(i).PrevDt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                        Case mTableType = "S" And mHeaderRowStatusFlag = True
                            Dim mHeaderRowStatusNo As Int16 = TableClassObject(i).HeaderRowStatusNo
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mHeaderRowStatusNo.ToString & " where " & mPrimaryKey & " = " & TableClassObject(i).PrevDt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                        Case Else
                            Dim pKeyValue As Integer = TableClassObject(i).PrevDt.Rows(jj).item(mPrimaryKey)
                            Dim Cind As Integer = FindRowIndexByPrimaryCols(TableClassObject(i).CurrDt, pKeyValue)
                            If Cind > -1 Then
                                Dim mRow As DataRow = TableClassObject(i).CurrDt.Rows(Cind)
                                Dim WhereClause As String = mPrimaryKey & " = " & pKeyValue
                                Dim QueryStr As String = CreateUpdateQuery(lCommand, mRow, mTable, Dtschema, Cind, , WhereClause)   'Changes by Neha
                                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
                            Else
                                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "delete from " & mTable & " where " & mPrimaryKey & " = " & TableClassObject(i).prevdt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                            End If
                    End Select
                Next

            End If
            If mCurrentExtraRows.Count > 0 Then
                Dim mRows() As DataRow = {}
                For j = 0 To mCurrentExtraRows.Count - 1
                    Dim jj As Integer = mCurrentExtraRows(j)
                    Dim mCurrDt As DataTable = TableClassObject(i).CurrDt
                    Dim nRow As DataRow = mCurrDt.Rows(jj)
                    If mRowStatusFlag = True Or (TableClassObject(i).TableType = "S" And TableClassObject(i).HeaderRowStatusFlag = True) Then
                        nRow("RowStatus") = 0
                        GF1.ArrayAppend(mRows, nRow)
                    Else
                        Dim mUpdatedColumnsStr As String = nRow("UpdatedColumnsStr").ToString.Trim
                        If mUpdatedColumnsStr.Length = 0 Then
                            GF1.ArrayAppend(mRows, nRow)
                        End If
                    End If
                Next
                Dim QueryStr As String = ""
                If mRows.Count = 1 Then
                    QueryStr = CreateInsertQuery(lCommand, mRows(0), mTable, Dtschema, i)
                Else
                    QueryStr = CreateInsertQuery(lCommand, mRows, mTable, Dtschema, i)
                End If
                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
            End If
        Next
        Dim a1 As Integer = SqlExecuteNonQuery(Ltrans, lCommand, FinalQuery)
        Rowseffected = a1
        Return Rowseffected
        lCommand.Dispose()
        ' Catch ex As Exception
        'QuitError(ex, Err, "Unable to execute DataFunction.InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject() As Object)")
        'End Try
        Return True
    End Function



    ''' <summary>
    ''' Insert Records  to  multiple SQL Tables in a single batch execution.
    ''' </summary>
    ''' <param name="TableClassObject">An array of Class objects of tables</param>

    Public Function InsertUpdateDeleteSqlTables(ByVal TableClassObject() As Object) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return False : Exit Function
       
        '  Rowseffected = 0
        ' Try
        Dim FinalQuery As String = ""

        ' lCommand.Transaction = Ltrans
        '  lCommand.Parameters.Clear()
        For i = 0 To TableClassObject.Count - 1
            If TableClassObject(i).SqlUpdation = False Then
                Continue For
            End If

            Dim mTable As String = TableClassObject(i).TableName
            ' Dim mTable As String = IIf(SameDataSource(Ltrans, TableClassObject(i)) = True, TableClassObject(i).TableName, TableClassObject(i).TableWithSQLPath)
            Dim mRowStatusFlag As Boolean = TableClassObject(i).RowStatusFlag
            Dim mTableEntryType As String = TableClassObject(i).TableEntryType
            Dim mPrimaryKey As String = TableClassObject(i).PrimaryKey
            Dim Dtschema As DataTable = TableClassObject(i).SchemaTable
            Dim mFieldsFinalValues As Hashtable = TableClassObject(i).FieldsFinalValues
            Dim mPreviousExtraRows() As Integer = TableClassObject(i).PreviousExtraRows
            Dim mCurrentExtraRows() As Integer = TableClassObject(i).CurrentExtraRows
            Dim mTableType As String = TableClassObject(i).TableType
            Dim mHeaderRowStatusFlag As Boolean = TableClassObject(i).HeaderRowStatusFlag
            If mTableEntryType = "M" Or mTableEntryType = "D" Then
                For j = 0 To mPreviousExtraRows.Count - 1
                    Dim jj As Integer = mPreviousExtraRows(j)
                    Select Case True
                        Case mRowStatusFlag = True And InStr("H,M", mTableType) > 0
                            Dim mRowStatus As Integer = GF1.GetValueFromHashTable(mFieldsFinalValues, "RowStatus")
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mRowStatus.ToString & " where " & mPrimaryKey & " = " & TableClassObject(i).PrevDt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                        Case mTableType = "S" And mHeaderRowStatusFlag = True
                            Dim mHeaderRowStatusNo As Int16 = TableClassObject(i).HeaderRowStatusNo
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mHeaderRowStatusNo.ToString & " where " & mPrimaryKey & " = " & TableClassObject(i).PrevDt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                        Case Else
                            Dim pKeyValue As Integer = TableClassObject(i).PrevDt.Rows(jj).item(mPrimaryKey)
                            Dim Cind As Integer = FindRowIndexByPrimaryCols(TableClassObject(i).CurrDt, pKeyValue)
                            If Cind > -1 Then
                                Dim mRow As DataRow = TableClassObject(i).CurrDt.Rows(Cind)
                                Dim WhereClause As String = mPrimaryKey & " = " & pKeyValue
                                Dim QueryStr As String = CreateUpdateQuery(mRow, mTable, Dtschema, Cind, , WhereClause)   'Changes by Neha
                                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
                            Else
                                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "delete from " & mTable & " where " & mPrimaryKey & " = " & TableClassObject(i).prevdt.Rows(jj).Item(mPrimaryKey).ToString & vbCrLf
                            End If
                    End Select
                Next

            End If
            If mCurrentExtraRows.Count > 0 Then
                Dim mRows() As DataRow = {}
                For j = 0 To mCurrentExtraRows.Count - 1
                    Dim jj As Integer = mCurrentExtraRows(j)
                    Dim mCurrDt As DataTable = TableClassObject(i).CurrDt
                    Dim nRow As DataRow = mCurrDt.Rows(jj)
                    If mRowStatusFlag = True Or (TableClassObject(i).TableType = "S" And TableClassObject(i).HeaderRowStatusFlag = True) Then
                        nRow("RowStatus") = 0
                        GF1.ArrayAppend(mRows, nRow)
                    Else
                        Dim mUpdatedColumnsStr As String = nRow("UpdatedColumnsStr").ToString.Trim
                        If mUpdatedColumnsStr.Length = 0 Then
                            GF1.ArrayAppend(mRows, nRow)
                        End If
                    End If
                Next
                Dim QueryStr As String = ""
                If mRows.Count = 1 Then
                    QueryStr = CreateInsertQuery(mRows(0), mTable, Dtschema, i)
                Else
                    QueryStr = CreateInsertQuery(mRows, mTable, Dtschema, i)
                End If
                FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
            End If
        Next

        Return FinalQuery
        '  Dim a1 As Integer = SqlExecuteNonQuery(Ltrans, lCommand, FinalQuery)
        'Rowseffected = a1
        'Return Rowseffected
        'lCommand.Dispose()
        ' Catch ex As Exception
        'QuitError(ex, Err, "Unable to execute DataFunction.InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject() As Object)")
        'End Try
        ' Return True
    End Function


   
    ''' <summary>
    ''' function to execute multiple queries in a transaction
    ''' </summary>
    ''' <param name="serverdatabase">serverdatabase name</param>
    ''' <param name="queries">space or semicolon separated queries</param>
    ''' <returns></returns>
    Public Function SqlExecuteQueriesInTransaction(ByVal serverdatabase As String, ByVal queries As String) As Boolean
        Dim command As New SqlCommand
        Dim result As Boolean
        Dim mytrans As SqlTransaction = BeginTransaction(serverdatabase)
        Try
            command = New SqlCommand(queries, mytrans.Connection, mytrans)
            result = command.ExecuteNonQuery()
            mytrans.Commit()
        Catch ex As Exception
            mytrans.Rollback()
        End Try
        command.Dispose()
        mytrans.Dispose()
        '  mytrans .
        Return result
    End Function

    ''' <summary>
    ''' Check  for data updations wether sql statements on TableClassOject() will be executed or not.
    ''' </summary>
    ''' <param name="TableClassObject">An array of Class objects of tables</param>
    Public Function CheckTableClassUpdations(ByRef TableClassObject() As Object) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return False : Exit Function
        Dim FinalFlag As Boolean = False
        Try
            For i = 0 To TableClassObject.Count - 1
                ' Dim TableClass As Object = TableClassObject(i)
                Dim mFlag As Boolean = CheckTableClassUpdation(TableClassObject(i))
                FinalFlag = IIf(mFlag = True, mFlag, FinalFlag)
            Next
            If FinalFlag = True Then
                For i = 0 To TableClassObject.Count - 1
                    If TableClassObject(i).TableType = "H" Then
                        TableClassObject(i).SqlUpdation = FinalFlag
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.Public Function CheckTableClassUpdations(ByRef TableClassObject As Object) As Boolean")
        End Try
        Return FinalFlag
    End Function


    ''' <summary>
    ''' Check for data updations wether sql statements on TableClassOject will be executed or not.
    ''' </summary>
    ''' <param name="TableClassObject">An array of Class objects of tables</param>
    Public Function CheckTableClassUpdation(ByRef TableClassObject As Object) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return False : Exit Function
        Dim mFlag As Boolean = False
        Try
            ' For i = 0 To TableClassObject.Count - 1
            If TableClassObject.MultyRowsSqlHandling = False Then
                If GF1.IsDataRowEmpty(TableClassObject.PrevRow) = False Then
                    TableClassObject.PrevDt.Rows.Clear()
                    Dim mrow As DataRow = TableClassObject.PrevDt.NewRow
                    mrow.ItemArray = TableClassObject.PrevRow.ItemArray
                    TableClassObject.PrevDt.Rows.Add(mrow)
                End If
                'TableClassObject(i).CurrRowsArray = DirectCast(TableClassObject(i).CurrRowsArray, System.Array)
                Dim RowArray() As DataRow = TableClassObject.CurrRowsArray
                If RowArray.Count > 0 Then
                    'TableClassObject.PrevDt.Rows.Clear()
                    'TableClassObject.PrevRow = TableClassObject.NewRow
                    TableClassObject.CurrDt.Rows.Clear()
                    For x = 0 To RowArray.Count - 1
                        If IsDBNull(RowArray(x).Item(TableClassObject.PrimaryKey)) = True Then
                            RowArray(x).Item(TableClassObject.PrimaryKey) = -100 - x
                        End If
                        TableClassObject.CurrDt.Rows.Add(RowArray(x))
                    Next
                Else
                    If GF1.IsDataRowEmpty(TableClassObject.CurrRow) = False Then
                        TableClassObject.CurrDt.Rows.Clear()
                        Dim mrow As DataRow = TableClassObject.CurrDt.NewRow
                        mrow.ItemArray = TableClassObject.CurrRow.ItemArray
                        If IsDBNull(mrow(TableClassObject.PrimaryKey)) = True Then
                            mrow(TableClassObject.PrimaryKey) = -1
                        End If
                        TableClassObject.CurrDt.Rows.Add(mrow)
                    End If
                End If
            Else
                TableClassObject.currdt = RemoveRowFromDataTable(TableClassObject.CurrDt, -9999)
            End If

            Dim mTableEntryType As String = TableClassObject.TableEntryType

            If mTableEntryType = "S" Then
                Return mFlag
                Exit Function
            End If
            Dim mPreviousSameRows() As Integer = {}
            Dim mPreviousExtraRows() As Integer = {}
            Dim mCurrentExtraRows() As Integer = {}
            Dim mCurrentSameRows() As Integer = {}
            If mTableEntryType = "M" Or mTableEntryType = "D" Then
                Dim mRowStatusFlag As Boolean = TableClassObject.RowStatusFlag
                Dim mTableType As String = TableClassObject.TableType
                Dim mExcludeColumns As String = TableClassObject.ExcludeFromCompare
                mFlag = CompareTwoDataTablesRows(TableClassObject.PrevDt, TableClassObject.CurrDt, mExcludeColumns, mPreviousExtraRows, mCurrentExtraRows, mPreviousSameRows, mCurrentSameRows)
                TableClassObject.SqlUpdation = mFlag
                TableClassObject.PreviousSameRows = mPreviousSameRows
                TableClassObject.PreviousExtraRows = mPreviousExtraRows
                TableClassObject.CurrentExtraRows = mCurrentExtraRows
                TableClassObject.CurrentSameRows = mCurrentSameRows
                If mRowStatusFlag = False Then
                    For u = 0 To mPreviousExtraRows.Length - 1
                        Dim uu As Integer = mPreviousExtraRows(u)
                        Dim pRow As DataRow = TableClassObject.PrevDt.Rows(uu)
                        Dim PkeyValue As Integer = pRow(TableClassObject.PrimaryKey)
                        Dim cInd As Integer = FindRowIndexByPrimaryCols(TableClassObject.CurrDt, PkeyValue)
                        If cInd > -1 Then
                            Dim cRow As DataRow = TableClassObject.CurrDt.Rows(cInd)
                            TableClassObject.CurrDt.Rows(cInd).Item("UpdatedColumnsStr") = CompareTwoDataRowsValues(pRow, cRow, mExcludeColumns)
                        End If
                    Next
                End If
                If mRowStatusFlag = True Then
                    If TableClassObject.TableType = "H" And mCurrentExtraRows.Length = 0 And mTableEntryType = "M" And mPreviousExtraRows.Length > 0 Then
                        mCurrentExtraRows = mPreviousExtraRows
                        TableClassObject.CurrentExtraRows = mPreviousExtraRows
                    End If
                    If TableClassObject.TableType = "H" And mCurrentExtraRows.Length > 0 And mTableEntryType = "M" And mPreviousExtraRows.Length > 0 Then
                        mCurrentExtraRows = mPreviousExtraRows
                        TableClassObject.CurrentExtraRows = mPreviousExtraRows
                    End If
                    For u = 0 To TableClassObject.CurrDt.Rows.Count - 1
                        TableClassObject.CurrDt.Rows(u).item(TableClassObject.PrimaryKey) = -1 * (u + 1)
                    Next
                End If
            End If

            If mTableEntryType = "A" Then
                TableClassObject.PreviousSameRows = mPreviousSameRows
                TableClassObject.PreviousExtraRows = mPreviousExtraRows
                TableClassObject.CurrentSameRows = mCurrentSameRows
                For u = 0 To TableClassObject.CurrDt.Rows.Count - 1
                    GF1.ArrayAppend(TableClassObject.CurrentExtraRows, u)
                Next
                TableClassObject.SqlUpdation = True
                mFlag = True
            End If
            TableClassObject.KeyPlusGroups = GetKeyPlusGroups(TableClassObject)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.Public Function CheckTableClassUpdations(ByRef TableClassObject As Object) As Boolean")
        End Try
        Return mFlag
    End Function
    Private Function GetKeyPlusGroups(ByRef ClsObject As Object) As String
        Dim mKeyPlusGroups As String = ""
        Dim yflag As Boolean = False, rflag As Boolean = False, oflag As Boolean = False, sflag As Boolean = False, dflag As Boolean = False
        If ClsObject.RowStatusFlag = True Then
            If ClsObject.PreviousExtraRows.Length > 0 And ClsObject.CurrentExtraRows.Length = 0 Then
                sflag = True
            End If
            If ClsObject.CurrentExtraRows.Length > 0 And ClsObject.PreviousExtraRows.Length = 0 Then
                yflag = True : oflag = True : rflag = True : dflag = True
            End If
            If ClsObject.CurrentExtraRows.Length > 0 And ClsObject.PreviousExtraRows.Length > 0 Then
                yflag = True : sflag = True
            End If
        Else
            If ClsObject.CurrentExtraRows.Length > 0 And ClsObject.PreviousExtraRows.Length = 0 Then
                yflag = True : oflag = True : dflag = True : rflag = True
            End If
            If ClsObject.CurrentExtraRows.Length > 0 And ClsObject.PreviousExtraRows.Length > 0 Then
                For x = 0 To ClsObject.CurrentExtraRows.Length - 1
                    If IsDBNull(ClsObject.CurrDt.Rows(x).Item("UpdatedColumnsStr")) = False Then
                        If ClsObject.CurrDt.Rows(x).Item("UpdatedColumnsStr").ToString.Length = 0 Then
                            yflag = True : rflag = True
                        End If
                    End If
                Next
            End If
        End If

        If yflag = True Then
            mKeyPlusGroups = mKeyPlusGroups & IIf(mKeyPlusGroups.Length = 0, "", ",") & "Y"
        End If
        If rflag = True Then
            mKeyPlusGroups = mKeyPlusGroups & IIf(mKeyPlusGroups.Length = 0, "", ",") & "R"
        End If
        If sflag = True Then
            mKeyPlusGroups = mKeyPlusGroups & IIf(mKeyPlusGroups.Length = 0, "", ",") & "S"
        End If
        If oflag = True Then
            mKeyPlusGroups = mKeyPlusGroups & IIf(mKeyPlusGroups.Length = 0, "", ",") & "O"
        End If
        If dflag = True Then
            mKeyPlusGroups = mKeyPlusGroups & IIf(mKeyPlusGroups.Length = 0, "", ",") & "D"
        End If
        Return mKeyPlusGroups
    End Function


    ''' <summary>
    ''' Check the connection server-database with  table class object server-database
    ''' </summary>
    ''' <param name="SqlTrans"></param>
    ''' <param name="TableClas"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SameDataSource(ByVal SqlTrans As SqlTransaction, ByVal TableClas As Object) As Boolean
        Dim mflag As Boolean = False
        Dim mserver As String = SqlTrans.Connection.DataSource
        Dim mdatabase As String = SqlTrans.Connection.Database
        If LCase(TableClas.Server) = LCase(mserver) And LCase(TableClas.Database) = LCase(mdatabase) Then
            mflag = True
        End If
        Return mflag
    End Function
    ''' <summary>
    ''' Check the connection server-database with  table class object server-database
    ''' </summary>
    ''' <param name="SqlConn"></param>
    ''' <param name="TableClas"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SameDataSource(ByVal SqlConn As SqlConnection, ByVal TableClas As Object) As Boolean
        Dim mflag As Boolean = False
        Dim mserver As String = SqlConn.DataSource
        Dim mdatabase As String = SqlConn.Database
        If LCase(TableClas.Server) = LCase(mserver) And LCase(TableClas.Database) = LCase(mdatabase) Then
            mflag = True
        End If
        Return mflag
    End Function
    ''' <summary>
    ''' Get most used server database connection available in tables of  TableClass()
    ''' </summary>
    ''' <param name="TableClass"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetServerMDFForTransanction(ByVal TableClass() As Object) As String
        Dim mserverdatabase As New Hashtable
        For i = 0 To TableClass.Count - 1
            Dim mstr As String = TableClass(i).serverdatabase
            Dim mno As Int16 = 0
            If GF1.GetValueFromHashTable(mserverdatabase, mstr) IsNot Nothing Then
                mno = val(GF1.GetValueFromHashTable(mserverdatabase, mstr)) 'changed by Neha
            End If
            GF1.AddItemToHashTable(mserverdatabase, mstr, mno + 1)
        Next
        If mserverdatabase.Count = 1 Then
            Return TableClass(0).serverdatabase
            Exit Function
        End If
        Dim high As String = ""
        Dim nhigh As Int16 = 0
        For i = 0 To mserverdatabase.Count - 1
            Dim mkey As String = mserverdatabase.Keys(i).ToString
            Dim mitem As Int16 = mserverdatabase.Item(mkey(i))
            If mitem >= nhigh Then
                nhigh = mitem
                high = mkey
            End If
        Next
        Return high
    End Function


    ''' <summary>
    ''' Insert Records  to  SQL Tables in  rowwise batch execution.
    ''' </summary>
    ''' <param name="Ltrans" >sql transaction </param>
    ''' <param name="TableClassObject">A Class object of sql table</param>
    ''' <param name="FinalFlag" >Final datatable updated</param>
    Public Function InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject As Object, Optional ByVal FinalFlag As Boolean = True) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Function
        If FinalFlag = False Then
            Return False
            Exit Function
        End If
        If TableClassObject.SqlUpdation = False Then
            Return False
            Exit Function
        End If
        Dim mTable As String = IIf(SameDataSource(Ltrans, TableClassObject) = True, TableClassObject.TableName, TableClassObject.TableWithSQLPath)
        Try
            Dim FinalQuery As String = ""
            Dim mTableType As String = TableClassObject.TableType
            Dim mRowStatusFlag As Boolean = IIf(mTableType = "S", TableClassObject.HeaderRowStatusFlag, TableClassObject.RowStatusFlag)
            Dim mTableEntryType As String = TableClassObject.TableEntryType
            Dim mPrimaryKey As String = TableClassObject.PrimaryKey
            Dim Dtschema As DataTable = TableClassObject.SchemaTable
            Dim lCommand As New SqlCommand
            lCommand.Connection = Ltrans.Connection
            lCommand.Transaction = Ltrans
            lCommand.Parameters.Clear()
            If mTableEntryType = "M" Or mTableEntryType = "D" Then
                For j = 0 To TableClassObject.PreviousExtraRows.Count - 1
                    Dim jj As Integer = TableClassObject.PreviousExtraRows(j)
                    If mRowStatusFlag = True Then
                        Dim mRowStatus As Integer = GF1.GetValueFromHashTable(TableClassObject.FieldsFinalValues, "RowStatus")
                        If mTableType = "S" Then
                            mRowStatus = TableClassObject.HeaderRowStatusNo
                        End If


                        FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mRowStatus.ToString & " where " & mPrimaryKey & " = " & TableClassObject.PrevDt(jj)(mPrimaryKey).ToString & vbCrLf
                    Else
                        Dim pKeyValue As Integer = TableClassObject.PrevDt(jj)(mPrimaryKey)
                        Dim Cind As Integer = FindRowIndexByPrimaryCols(TableClassObject.CurrDt, pKeyValue)
                        If Cind > -1 Then
                            Dim mRow As DataRow = TableClassObject.CurrDt(Cind)
                            Dim QueryStr As String = CreateUpdateQuery(lCommand, mRow, mTable, Dtschema, Cind)
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
                        Else
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "delete from " & mTable & " where " & mPrimaryKey & " = " & TableClassObject.prevdt(jj)(mPrimaryKey).ToString & vbCrLf
                        End If
                    End If
                Next
            End If
            Dim mRows() As DataRow = {}
            For j = 0 To TableClassObject.CurrentExtraRows.Count - 1
                Dim jj As Integer = TableClassObject.CurrentExtraRows(j)
                GF1.ArrayAppend(mRows, TableClassObject.CurrDt.rows(jj))
            Next
            If mRows.Count = 1 Then
                FinalQuery = FinalQuery & CreateInsertQuery(lCommand, mRows(0), mTable, Dtschema)
            Else
                FinalQuery = FinalQuery & CreateInsertQuery(lCommand, mRows, mTable, Dtschema)
            End If
            Dim a1 As Integer = SqlExecuteNonQuery(Ltrans, lCommand, FinalQuery)
            lCommand.Dispose()
        Catch ex As Exception

            QuitError(ex, Err, "Unable to execute DataFunction.InsertDataTablesToSqlTables(ByRef Ltrans As SqlTransaction, ByVal CurrentDataTables() As DataTable, ByVal PreviousDataTables() As DataTable, ByVal ExcludeColumns() As String)")
        End Try
    End Function
    ''' <summary>
    ''' Insert Records  to  SQL Tables in  rowwise batch execution.
    ''' </summary>
    ''' <param name="Ltrans" >sql transaction </param>
    ''' <param name="TableClassObject">A Class object of sql table</param>
    ''' <param name="RowsEffected" >Number of Rows effected</param>
    Public Function InsertUpdateDeleteSqlTables(ByRef Ltrans As SqlTransaction, ByVal TableClassObject As Object, ByRef RowsEffected As Integer) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Function
        If TableClassObject.SqlUpdation = False Then
            Return False
            Exit Function
        End If
        Dim mTable As String = IIf(SameDataSource(Ltrans, TableClassObject) = True, TableClassObject.TableName, TableClassObject.TableWithSQLPath)
        Try
            Dim FinalQuery As String = ""
            Dim mTableType As String = TableClassObject.TableType
            Dim mRowStatusFlag As Boolean = IIf(mTableType = "S", TableClassObject.HeaderRowStatusFlag, TableClassObject.RowStatusFlag)
            Dim mTableEntryType As String = TableClassObject.TableEntryType
            Dim mPrimaryKey As String = TableClassObject.PrimaryKey
            Dim Dtschema As DataTable = TableClassObject.SchemaTable
            Dim lCommand As New SqlCommand
            lCommand.Connection = Ltrans.Connection
            lCommand.Transaction = Ltrans
            lCommand.Parameters.Clear()
            If mTableEntryType = "M" Or mTableEntryType = "D" Then
                For j = 0 To TableClassObject.PreviousExtraRows.Count - 1
                    Dim jj As Integer = TableClassObject.PreviousExtraRows(j)
                    If mRowStatusFlag = True Then
                        Dim mRowStatus As Integer = GF1.GetValueFromHashTable(TableClassObject.FieldsFinalValues, "RowStatus")
                        If mTableType = "S" Then
                            mRowStatus = TableClassObject.HeaderRowStatusNo
                        End If
                        FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "update " & mTable & " set RowStatus = " & mRowStatus.ToString & " where " & mPrimaryKey & " = " & TableClassObject.PrevDt(jj)(mPrimaryKey).ToString & vbCrLf
                    Else
                        Dim pKeyValue As Integer = TableClassObject.PrevDt(jj)(mPrimaryKey)
                        Dim Cind As Integer = FindRowIndexByPrimaryCols(TableClassObject.CurrDt, pKeyValue)
                        If Cind > -1 Then
                            Dim mRow As DataRow = TableClassObject.CurrDt(Cind)
                            Dim QueryStr As String = CreateUpdateQuery(lCommand, mRow, mTable, Dtschema, Cind)
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & QueryStr
                        Else
                            FinalQuery = FinalQuery & IIf(FinalQuery.Length = 0, "", vbCrLf) & "delete from " & mTable & " where " & mPrimaryKey & " = " & TableClassObject.prevdt(jj)(mPrimaryKey).ToString & vbCrLf
                        End If
                    End If
                Next
            End If
            Dim mRows() As DataRow = {}
            For j = 0 To TableClassObject.CurrentExtraRows.Count - 1
                Dim jj As Integer = TableClassObject.CurrentExtraRows(j)
                GF1.ArrayAppend(mRows, TableClassObject.CurrDt.rows(jj))
            Next
            If mRows.Count = 1 Then
                FinalQuery = FinalQuery & CreateInsertQuery(lCommand, mRows(0), mTable, Dtschema)
            Else
                FinalQuery = FinalQuery & CreateInsertQuery(lCommand, mRows, mTable, Dtschema)
            End If
            Dim a1 As Integer = SqlExecuteNonQuery(Ltrans, lCommand, FinalQuery)
            RowsEffected = a1
            Return RowsEffected
            lCommand.Dispose()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertDataTablesToSqlTables(ByRef Ltrans As SqlTransaction, ByVal CurrentDataTables() As DataTable, ByVal PreviousDataTables() As DataTable, ByVal ExcludeColumns() As String)")
        End Try
    End Function
    ''' <summary>
    ''' Create a string query with sql parameters from a datarow to be inserted.
    ''' </summary>
    ''' <param name="Lcommand">Command as sql command</param>
    ''' <param name="mRow">row of values inserted</param>
    ''' <param name="mTable">Table name</param>
    ''' <param name="DtSchema">Schema of table</param>
    ''' <param name="RowSNo">If sql command has more than one row then row no</param>
    ''' <param name="TableNo">If sql command has more than one table than table no.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateInsertQuery(ByRef Lcommand As SqlCommand, ByVal mRow As DataRow, ByVal mTable As String, ByVal DtSchema As DataTable, Optional ByVal RowSNo As Integer = 0, Optional ByVal TableNo As Int16 = 0) As String
        '   Dim mRow As DataRow = TableClassObject.CurrDt.rows(CurrentExtraRows(j))
        Dim FieldStr As String = ""
        Dim ValueStr As String = ""
        For k = 0 To DtSchema.Rows.Count - 1
            Dim mfield As String = DtSchema(k)("Column_Name").ToString
            Dim mtype As String = LCase(DtSchema(k)("Data_Type").ToString)
            Dim Evalue As String = ""
            Dim ParaName As String = ""
            If IsDBNull(mRow(mfield)) = True Then
                Continue For
            End If
            Select Case mtype
                Case "int", "numeric", "decimal", "smallint", "tinyint"
                    Evalue = mRow(mfield)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                Case "bit"
                    Evalue = LCase(mRow(mfield).ToString.Trim)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                        Evalue = "1"
                    Else
                        Evalue = "0"
                    End If
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                Case "datetime"
                    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSNo.ToString
                    Lcommand.Parameters.Add(ParaName, SqlDbType.DateTime).Value = mRow(mfield)
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Case "image"
                    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSNo.ToString
                    Dim mimagebyte() As Byte = mRow(mfield)
                    Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = mimagebyte
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Case "nchar", "nvarchar"
                    Evalue = mRow(mfield).ToString
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                Case "varbinary"
                    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSNo.ToString
                    Dim mimagebyte() As Byte = mRow(mfield)
                    Lcommand.Parameters.Add(ParaName, SqlDbType.VarBinary).Value = mimagebyte
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Case Else
                    Evalue = mRow(mfield).ToString
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
            End Select
        Next
        Dim FinalQuery As String = "Insert  into " & mTable & " (" & FieldStr & ") values (" & ValueStr & ")"
        Return FinalQuery
    End Function



    ''' <summary>
    ''' Create a string query with sql parameters from a datarow to be inserted.
    ''' </summary>
    ''' <param name="mRow">row of values inserted</param>
    ''' <param name="mTable">Table name</param>
    ''' <param name="DtSchema">Schema of table</param>
    ''' <param name="RowSNo">If sql command has more than one row then row no</param>
    ''' <param name="TableNo">If sql command has more than one table than table no.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateInsertQuery(ByVal mRow As DataRow, ByVal mTable As String, ByVal DtSchema As DataTable, Optional ByVal RowSNo As Integer = 0, Optional ByVal TableNo As Int16 = 0) As String
        '   Dim mRow As DataRow = TableClassObject.CurrDt.rows(CurrentExtraRows(j))
        Dim FieldStr As String = ""
        Dim ValueStr As String = ""
        For k = 0 To DtSchema.Rows.Count - 1
            Dim mfield As String = DtSchema(k)("Column_Name").ToString
            Dim mtype As String = LCase(DtSchema(k)("Data_Type").ToString)
            Dim Evalue As String = ""
            Dim ParaName As String = ""
            If IsDBNull(mRow(mfield)) = True Then
                Continue For
            End If
            Select Case mtype
                Case "int", "numeric", "decimal", "smallint", "tinyint"
                    Evalue = mRow(mfield)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                Case "bit"
                    Evalue = LCase(mRow(mfield).ToString.Trim)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                        Evalue = "1"
                    Else
                        Evalue = "0"
                    End If
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                Case "datetime"
                    '  ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSNo.ToString
                    Evalue = mRow(mfield).ToString("yyyy-MM-dd hh:mm:ss")
                    '  Lcommand.Parameters.Add(ParaName, SqlDbType.DateTime).Value = mRow(mfield)
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    'Case "image"
                    '    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSNo.ToString
                    '    Dim mimagebyte() As Byte = mRow(mfield)
                    '    Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = mimagebyte
                    '    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    '    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Case "nchar", "nvarchar"
                    Evalue = mRow(mfield).ToString
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "N'" & Evalue.Trim & "'"
                    'Case "varbinary"
                    '    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSNo.ToString
                    '    Dim mimagebyte() As Byte = mRow(mfield)
                    '    Lcommand.Parameters.Add(ParaName, SqlDbType.VarBinary).Value = mimagebyte
                    '    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    '    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Case Else
                    Evalue = mRow(mfield).ToString
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
            End Select
        Next
        Dim FinalQuery As String = "Insert  into " & mTable & " (" & FieldStr & ") values (" & ValueStr & ")"
        Return FinalQuery
    End Function






    ''' <summary>
    ''' Create a string query with sql parameters from a datarow to be inserted.
    ''' </summary>
    ''' <param name="Lcommand">Command as sql command</param>
    ''' <param name="mRow">row of values inserted</param>
    ''' <param name="mTable">Table name</param>
    ''' <param name="DtSchema">Schema of table</param>
    ''' <param name="RowSNo">If sql command has more than one row then row no</param>
    ''' <param name="TableNo">If sql command has more than one table than table no.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateUpdateQuery(ByRef Lcommand As SqlCommand, ByVal mRow As DataRow, ByVal mTable As String, ByVal DtSchema As DataTable, Optional ByVal RowSNo As Integer = 0, Optional ByVal TableNo As Int16 = 0, Optional ByVal WhereClause As String = "") As String     'Changes by Neha
        '   Dim mRow As DataRow = TableClassObject.CurrDt.rows(CurrentExtraRows(j))
        Dim ValueStr As String = ""
        If mRow("UpdatedColumnsStr").ToString.Trim.Length = 0 Then
            Return ""
        End If
        Dim uColumns() As String = mRow("UpdatedColumnsStr").ToString.Trim.Split(",")

        For k = 0 To uColumns.Count - 1
            Dim mfield As String = uColumns(k)
            Dim sRow As DataRow = FindRowByPrimaryCols(DtSchema, mfield)
            Dim mtype As String = LCase(sRow("Data_Type").ToString)
            Dim Evalue As String = ""
            Dim ParaName As String = ""
            If IsDBNull(mRow(mfield)) = True Then
                Continue For
            End If
            Select Case mtype
                Case "int", "numeric", "decimal", "smallint", "tinyint"
                    Evalue = mRow(mfield)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & Evalue
                Case "bit"
                    Evalue = LCase(mRow(mfield).ToString.Trim)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                        Evalue = "1"
                    Else
                        Evalue = "0"
                    End If
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                Case "datetime"
                    ParaName = "@" & mfield & "_upd" & TableNo.ToString & "_" & RowSNo.ToString
                    Lcommand.Parameters.Add(ParaName, SqlDbType.DateTime).Value = mRow(mfield)
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & ParaName
                Case "image"
                    ParaName = "@" & mfield & "_upd" & TableNo.ToString & "_" & RowSNo.ToString
                    Dim mimagebyte() As Byte = mRow(mfield)
                    Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = mimagebyte
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & ParaName
                Case "nchar", "nvarchar"
                    Evalue = mRow(mfield).ToString
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & "'" & Evalue.Trim & "'"

                Case "varbinary"

                    ParaName = "@" & mfield & "_upd" & TableNo.ToString & "_" & RowSNo.ToString
                    Dim mimagebyte() As Byte = mRow(mfield)
                    Lcommand.Parameters.Add(ParaName, SqlDbType.VarBinary).Value = mimagebyte
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & ParaName

                Case Else
                    Evalue = mRow(mfield).ToString
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & "'" & Evalue.Trim & "'"
            End Select
        Next
        Dim FinalQuery As String = "update  " & mTable & " set " & ValueStr & " where " & WhereClause   'changes by Neha
        Return FinalQuery
    End Function



    ''' <summary>
    ''' Create a string query with sql parameters from a datarow to be inserted.
    ''' </summary>
    ''' <param name="mRow">row of values inserted</param>
    ''' <param name="mTable">Table name</param>
    ''' <param name="DtSchema">Schema of table</param>
    ''' <param name="RowSNo">If sql command has more than one row then row no</param>
    ''' <param name="TableNo">If sql command has more than one table than table no.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateUpdateQuery(ByVal mRow As DataRow, ByVal mTable As String, ByVal DtSchema As DataTable, Optional ByVal RowSNo As Integer = 0, Optional ByVal TableNo As Int16 = 0, Optional ByVal WhereClause As String = "") As String     'Changes by Neha
        '   Dim mRow As DataRow = TableClassObject.CurrDt.rows(CurrentExtraRows(j))
        Dim ValueStr As String = ""
        If mRow("UpdatedColumnsStr").ToString.Trim.Length = 0 Then
            Return ""
        End If
        Dim uColumns() As String = mRow("UpdatedColumnsStr").ToString.Trim.Split(",")

        For k = 0 To uColumns.Count - 1
            Dim mfield As String = uColumns(k)
            Dim sRow As DataRow = FindRowByPrimaryCols(DtSchema, mfield)
            Dim mtype As String = LCase(sRow("Data_Type").ToString)
            Dim Evalue As String = ""
            Dim ParaName As String = ""
            If IsDBNull(mRow(mfield)) = True Then
                Continue For
            End If
            Select Case mtype
                Case "int", "numeric", "decimal", "smallint", "tinyint"
                    Evalue = mRow(mfield)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & Evalue
                Case "bit"
                    Evalue = LCase(mRow(mfield).ToString.Trim)
                    If Evalue.Length = 0 Then
                        Evalue = "0"
                    End If
                    If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                        Evalue = "1"
                    Else
                        Evalue = "0"
                    End If
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                Case "datetime"
                    ' ParaName = "@" & mfield & "_upd" & TableNo.ToString & "_" & RowSNo.ToString
                    'Lcommand.Parameters.Add(ParaName, SqlDbType.DateTime).Value = mRow(mfield)
                    Evalue = mRow(mfield).ToString("yyyy-MM-dd hh:mm:ss")

                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & "'" & Evalue & "'"
                    'Case "image"
                    '    ParaName = "@" & mfield & "_upd" & TableNo.ToString & "_" & RowSNo.ToString
                    '    Dim mimagebyte() As Byte = mRow(mfield)
                    '    Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = mimagebyte
                    '    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & ParaName
                Case "nchar", "nvarchar"
                    Evalue = mRow(mfield).ToString
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & "'" & Evalue.Trim & "'"

                    'Case "varbinary"

                    '    ParaName = "@" & mfield & "_upd" & TableNo.ToString & "_" & RowSNo.ToString
                    '    Dim mimagebyte() As Byte = mRow(mfield)
                    '    Lcommand.Parameters.Add(ParaName, SqlDbType.VarBinary).Value = mimagebyte
                    '    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & ParaName

                Case Else
                    Evalue = mRow(mfield).ToString
                    Evalue = Evalue.Replace("'", " ")
                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & mfield & " = " & "'" & Evalue.Trim & "'"
            End Select
        Next
        Dim FinalQuery As String = "update  " & mTable & " set " & ValueStr & " where " & WhereClause   'changes by Neha
        Return FinalQuery
    End Function






    ''' <summary>
    ''' Create a string query with sql parameters from a datarow to be inserted.
    ''' </summary>
    ''' <param name="Lcommand">Command as sql command</param>
    ''' <param name="mRows">An array of datarows  to be  inserted</param>
    ''' <param name="mTable">Table name</param>
    ''' <param name="DtSchema">Schema of table</param>
    ''' <param name="TableNo">If sql command has more than one table than table no.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateInsertQuery(ByRef Lcommand As SqlCommand, ByVal mRows() As DataRow, ByVal mTable As String, ByVal DtSchema As DataTable, Optional ByVal TableNo As Int16 = 0) As String
        '   Dim mRow As DataRow = TableClassObject.CurrDt.rows(CurrentExtraRows(j))
        If mRows.Count = 0 Then
            Return ""
            Exit Function
        End If
        Dim FinalValues As String = ""
        Dim FinalFields As String = ""
        Dim sql2005str As String = ""
        For i = 0 To mRows.Count - 1
            Dim mrow As DataRow = mRows(i)
            Dim RowSno As Integer = i
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            For k = 0 To DtSchema.Rows.Count - 1
                Dim mfield As String = DtSchema(k)("Column_Name").ToString
                Dim mtype As String = LCase(DtSchema(k)("Data_Type").ToString)
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                Dim NullValue As Boolean = False
                If IsDBNull(mrow(mfield)) = True Then
                    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                    '   Dim evalue1 As System.DBNull = mrow(mfield)
                    If mtype = "image" Then
                        Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = DBNull.Value
                    Else
                        Lcommand.Parameters.AddWithValue(ParaName, DBNull.Value)
                    End If

                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Else
                    Select Case mtype
                        Case "int", "numeric", "decimal", "smallint", "tinyint"
                            Evalue = mrow(mfield)
                            If Evalue.Length = 0 Then
                                Evalue = "0"
                            End If
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        Case "bit"
                            Evalue = LCase(mrow(mfield).ToString.Trim)
                            If Evalue.Length = 0 Then
                                Evalue = "0"
                            End If
                            If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                                Evalue = "1"
                            Else
                                Evalue = "0"
                            End If
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        Case "datetime"
                            ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                            Lcommand.Parameters.Add(ParaName, SqlDbType.DateTime).Value = mrow(mfield)
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                        Case "image"
                            ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                            Dim mimagebyte() As Byte = mrow(mfield)
                            Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = mimagebyte
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                        Case "nchar", "nvarchar"
                            Evalue = mrow(mfield).ToString
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            Evalue = Evalue.Replace("'", " ")
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & IIf(Evalue.Trim.Length > 0, "N'", "'") & Evalue.Trim & "'"
                        Case "varbinary"
                            ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                            Dim mimagebyte() As Byte = mrow(mfield)
                            Lcommand.Parameters.Add(ParaName, SqlDbType.VarBinary).Value = mimagebyte
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield

                        Case Else
                            Evalue = mrow(mfield).ToString
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            Evalue = Evalue.Replace("'", " ")
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                    End Select
                End If
            Next
            FinalValues = FinalValues & IIf(FinalValues.Length = 0, "", "),(") & ValueStr
            FinalFields = FieldStr
            Dim str As String = "Insert  into " & mTable & " (" & FinalFields & ") values (" & ValueStr & ")"
            sql2005str = sql2005str & IIf(sql2005str.Length = 0, "", vbCrLf) & str
        Next
        Dim FinalQuery As String = "Insert  into " & mTable & " (" & FinalFields & ") values (" & FinalValues & ")"
        Return sql2005str
        'Return FinalQuery
    End Function


    ''' <summary>
    ''' Create a string query with sql parameters from a datarow to be inserted.
    ''' </summary>
    ''' <param name="mRows">An array of datarows  to be  inserted</param>
    ''' <param name="mTable">Table name</param>
    ''' <param name="DtSchema">Schema of table</param>
    ''' <param name="TableNo">If sql command has more than one table than table no.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateInsertQuery(ByVal mRows() As DataRow, ByVal mTable As String, ByVal DtSchema As DataTable, Optional ByVal TableNo As Int16 = 0) As String
        '   Dim mRow As DataRow = TableClassObject.CurrDt.rows(CurrentExtraRows(j))
        If mRows.Count = 0 Then
            Return ""
            Exit Function
        End If
        Dim FinalValues As String = ""
        Dim FinalFields As String = ""
        Dim sql2005str As String = ""
        For i = 0 To mRows.Count - 1
            Dim mrow As DataRow = mRows(i)
            Dim RowSno As Integer = i
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            For k = 0 To DtSchema.Rows.Count - 1
                Dim mfield As String = DtSchema(k)("Column_Name").ToString
                Dim mtype As String = LCase(DtSchema(k)("Data_Type").ToString)
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                Dim NullValue As Boolean = False
                If IsDBNull(mrow(mfield)) = True Then
                    '   ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                    '   Dim evalue1 As System.DBNull = mrow(mfield)
                    'If mtype = "image" Then
                    '    Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = DBNull.Value
                    'Else
                    '    Lcommand.Parameters.AddWithValue(ParaName, DBNull.Value)
                    'End If

                    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                Else
                    Select Case mtype
                        Case "int", "numeric", "decimal", "smallint", "tinyint"
                            Evalue = mrow(mfield)
                            If Evalue.Length = 0 Then
                                Evalue = "0"
                            End If
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        Case "bit"
                            Evalue = LCase(mrow(mfield).ToString.Trim)
                            If Evalue.Length = 0 Then
                                Evalue = "0"
                            End If
                            If Evalue = "true" Or Evalue = "y" Or Evalue = "yes" Then
                                Evalue = "1"
                            Else
                                Evalue = "0"
                            End If
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & Evalue
                        Case "datetime"
                            '   ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                            ' Lcommand.Parameters.Add(ParaName, SqlDbType.DateTime).Value = mrow(mfield)
                            Evalue = mrow(mfield).ToString("yyyy-MM-dd hh:mm:ss")
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue & "'"
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            'Case "image"
                            '    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                            '    Dim mimagebyte() As Byte = mrow(mfield)
                            '    Lcommand.Parameters.Add(ParaName, SqlDbType.Image).Value = mimagebyte
                            '    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                            '    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                        Case "nchar", "nvarchar"
                            Evalue = mrow(mfield).ToString
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            Evalue = Evalue.Replace("'", " ")
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & IIf(Evalue.Trim.Length > 0, "N'", "'") & Evalue.Trim & "'"
                            'Case "varbinary"
                            '    ParaName = "@" & mfield & "_" & TableNo.ToString & "_" & RowSno.ToString
                            '    Dim mimagebyte() As Byte = mrow(mfield)
                            '    Lcommand.Parameters.Add(ParaName, SqlDbType.VarBinary).Value = mimagebyte
                            '    ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & ParaName
                            '    FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield

                        Case Else
                            Evalue = mrow(mfield).ToString
                            FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mfield
                            Evalue = Evalue.Replace("'", " ")
                            ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                    End Select
                End If
            Next
            FinalValues = FinalValues & IIf(FinalValues.Length = 0, "", "),(") & ValueStr
            FinalFields = FieldStr
            Dim str As String = "Insert  into " & mTable & " (" & FinalFields & ") values (" & ValueStr & ")"
            sql2005str = sql2005str & IIf(sql2005str.Length = 0, "", vbCrLf) & str
        Next
        Dim FinalQuery As String = "Insert  into " & mTable & " (" & FinalFields & ") values (" & FinalValues & ")"
        Return sql2005str
        'Return FinalQuery
    End Function





    ''' <summary>
    ''' Insert Records  to  multiple SQL Tables in a single batch execution.
    ''' </summary>
    ''' <param name="TableClassObject">An Class objects of tables</param>
    Public Sub InsertUpdateDeleteSqlTables(ByVal TableClassObject As Object)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim ltrans As SqlTransaction = BeginTransaction(TableClassObject.ServerDataBase)
        Try
            InsertUpdateDeleteSqlTables(ltrans, TableClassObject)
            Try
                ltrans.Commit()
            Catch ex As Exception
                ltrans.Rollback()
                QuitError(ex, Err, "Unable to commit transaction at  DataFunction.InsertUpdateDeleteSqlTables(ByVal TableClassObject As Object)")
            End Try
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.InsertUpdateDeleteSqlTables(ByVal TableClassObject As Object)")
        End Try
    End Sub


    ''' <summary>
    ''' Update Record field's values of  an SQL Table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="ValueSource">Updating values as hashtable, where columnname as key and fieldvalue as hashtable value </param>
    ''' <param name="TargetTable">Name of table in which record updated</param>
    ''' <param name="WhereClause">Condition for records updated</param>
    ''' <remarks></remarks>
    Public Sub UpdateValuesToSqlRecords(ByVal ServerDataBase As String, ByVal ValueSource As Hashtable, ByVal TargetTable As String, ByVal WhereClause As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LTargetSchema As DataTable = GetSchemaTable(ServerDataBase, TargetTable)
            Dim mPrimaryKey As String = GetPrimaryKey(LTargetSchema)
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            Dim mPrimaryValue As String = ""
            Dim UpdateStr As String = ""
            Dim lwhere As String = ""
            Dim asqlparam() As SqlParameter = {}
            For i = 0 To ValueSource.Count - 1
                Dim mColumnName As String = LCase(ValueSource.Keys(i).ToString)
                Dim mindx() As Integer = SearchDataTableRowIndexSingleColumn(LTargetSchema, "columnname", mColumnName, True)
                If mindx.Count = 0 Then
                    Continue For
                End If
                Dim mtype As String = LCase(LTargetSchema.Rows(mindx(0)).Item("datatypename").ToString)
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                If LCase(mColumnName) = LCase(mPrimaryKey) Then
                    mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                    lwhere = mPrimaryKey & " = " & mPrimaryValue
                End If
                Select Case mtype
                    'int=10digits,smallint=32767,tinyint=255
                    Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = "0"
                        End If
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & Evalue
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = " & mPrimaryValue
                        End If
                    Case "image"
                        Dim ImageBytes() As Byte = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        ' lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Image
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "datetime"
                        Dim MdateTime As DateTime = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.DateTime
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "bit"
                        Dim mFlag As Boolean = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Bit
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "time"
                        Dim MdateTime As DateTime = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Time
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "nchar", "nvarchar"
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        '  ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & "N'" & Evalue.Trim & "'"
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If


                    Case Else
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        '  ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & "'" & Evalue.Trim & "'"
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                End Select
            Next
            Dim QryStr As String = "update  " & TargetTable & " set " & UpdateStr & IIf(WhereClause.Trim.Length > 0, " where " & WhereClause, "")
            Dim aa1 As Integer = SqlExecuteNonQuery(ServerDataBase, QryStr)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.UpdateValuesToSqlRecords(ByVal ServerDataBase As String, ByVal ValueSource As Hashtable, ByVal TargetTable As String, ByVal WhereClause As String)")
        End Try
    End Sub

    ''' <summary>
    ''' Update Record field's values of  an SQL Table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="ValueSource">Updating values in datarow , where columnnames are fieldNames </param>
    ''' <param name="TargetTable">Name of table in which record updated</param>
    ''' <param name="WhereClause">Condition for records updated</param>
    ''' <remarks></remarks>
    Public Sub UpdateValuesToSqlRecords(ByVal ServerDataBase As String, ByVal ValueSource As DataRow, ByVal TargetTable As String, ByVal WhereClause As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LTargetSchema As DataTable = GetSchemaTable(ServerDataBase, TargetTable)
            Dim mPrimaryKey As String = GetPrimaryKey(LTargetSchema)
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            Dim mPrimaryValue As String = ""
            Dim UpdateStr As String = ""
            Dim lwhere As String = ""
            Dim asqlparam() As SqlParameter = {}
            Dim mDataTable As DataTable = ValueSource.Table


            For i = 0 To mDataTable.Columns.Count - 1
                Dim mColumnName As String = mDataTable.Columns(i).ColumnName
                Dim mindx() As Integer = SearchDataTableRowIndexSingleColumn(LTargetSchema, "columnname", mColumnName, True)
                If mindx.Count = 0 Then
                    Continue For
                End If
                Dim mtype As String = LCase(LTargetSchema.Rows(mindx(0)).Item("datatypename").ToString)
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                If LCase(mColumnName) = LCase(mPrimaryKey) Then
                    mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                End If
                Select Case mtype
                    'int=10digits,smallint=32767,tinyint=255
                    Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = "0"
                        End If
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & Evalue
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = " & mPrimaryValue
                        End If
                    Case "image"
                        Dim ImageBytes() As Byte = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        ' lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Image
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "datetime"
                        Dim MdateTime As DateTime = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Image
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "bit"
                        Dim Mflag As Boolean = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Bit
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "time"
                        Dim MdateTime As DateTime = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Time
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "nchar", "nvarchar"
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        '  ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & "N'" & Evalue.Trim & "'"
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                    Case Else
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        '  ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & "'" & Evalue.Trim & "'"
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                End Select
            Next
            Dim QryStr As String = "update  " & TargetTable & " set " & UpdateStr & IIf(WhereClause.Trim.Length > 0, " where " & WhereClause, "")
            Dim aa1 As Integer = SqlExecuteNonQuery(ServerDataBase, QryStr, asqlparam)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.UpdateValuesToSqlRecords(ByVal ServerDataBase As String, ByVal ValueSource As DataRow, ByVal TargetTable As String, ByVal WhereClause As String)")
        End Try
    End Sub
    ''' <summary>
    ''' Update Record field's values of  an SQL Table
    ''' </summary>
    ''' <param name="mSqlTransanction">sql transanction already created</param>
    ''' <param name="ValueSource">Updating values in datarow , where columnnames are fieldNames </param>
    ''' <param name="TargetTable">Name of table in which record updated</param>
    ''' <param name="WhereClause">Condition for records updated</param>
    ''' <remarks></remarks>
    Public Sub UpdateValuesToSqlRecords(ByVal mSqlTransanction As SqlTransaction, ByVal ValueSource As DataRow, ByVal TargetTable As String, Optional ByVal WhereClause As String = "", Optional ByVal WhereClause1 As Hashtable = Nothing)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim mPrimaryValue As String = ""
            Dim LTargetSchema As DataTable = GetSchemaInformations(mSqlTransanction, TargetTable)
            Dim mPrimaryKey As String = GetPrimaryKey(LTargetSchema)
            Dim FieldStr As String = ""
            Dim ValueStr As String = ""
            Dim UpdateStr As String = ""
            Dim lwhere As String = ""
            Dim lwhere1 As String = ""
            Dim asqlparam() As SqlParameter = {}
            Dim mDataTable As DataTable = ValueSource.Table
            For i = 0 To mDataTable.Columns.Count - 1
                Dim mColumnName As String = mDataTable.Columns(i).ColumnName
                Dim mindx() As Integer = SearchDataTableRowIndexSingleColumn(LTargetSchema, "column_name", mColumnName, True)
                If mindx.Count = 0 Then
                    Continue For
                End If
                Dim mtype As String = LCase(LTargetSchema.Rows(mindx(0)).Item("data_type").ToString)
                Dim Evalue As String = ""
                Dim ParaName As String = ""
                'If mtype <> "image" Then
                '    Evalue = ValueSource(mColumnName)
                'End If
                If LCase(mColumnName) = LCase(mPrimaryKey) Then
                    mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                End If
                Select Case mtype
                    'int=10digits,smallint=32767,tinyint=255
                    Case "int", "numeric", "decimal", "smallint", "tinyint", "bigint"
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = "0"
                        End If
                        If Evalue.Length = 0 Then
                            Evalue = "0"
                        End If
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & Evalue
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = " & mPrimaryValue
                        End If
                    Case "image"
                        Dim ImageBytes() As Byte = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        ' lcmd0.Parameters.Add(ParaName, SqlDbType.Image, ImageBytes.Count).Value = ImageBytes
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Image
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "datetime"
                        Dim MdateTime As DateTime = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.DateTime
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "bit"
                        Dim Mflag As Boolean = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Bit
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "time"
                        Dim MdateTime As DateTime = ValueSource(mColumnName)
                        ParaName = "@" & mColumnName
                        Dim kk As New SqlParameter
                        kk.ParameterName = ParaName
                        kk.SqlValue = ValueSource(mColumnName)
                        kk.SqlDbType = SqlDbType.Time
                        ArrayAppendSqlParameter(asqlparam, kk)
                        FieldStr = FieldStr & IIf(FieldStr.Length = 0, "", ",") & mColumnName
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & ParaName
                    Case "nchar", "nvarchar"
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        '  ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & "N'" & Evalue.Trim & "'"
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                    Case Else
                        If IsDBNull(ValueSource(mColumnName)) = False Then
                            Evalue = ValueSource(mColumnName).ToString.Trim
                        Else
                            Evalue = ""
                        End If
                        Evalue = Evalue.Replace("'", " ")
                        '  ValueStr = ValueStr & IIf(ValueStr.Length = 0, "", ",") & "'" & Evalue.Trim & "'"
                        UpdateStr = UpdateStr & IIf(UpdateStr.Length = 0, "", ",") & mColumnName & " = " & "'" & Evalue.Trim & "'"
                        If LCase(mColumnName) = LCase(mPrimaryKey) Then
                            mPrimaryKey = ValueSource(mColumnName).ToString.Trim
                            WhereClause = mPrimaryKey & " = '" & mPrimaryValue & "'"
                        End If
                End Select
            Next
            Dim QryStr As String = "update  " & TargetTable & " set " & UpdateStr & IIf(WhereClause.Trim.Length > 0, " where " & WhereClause, "")
            If WhereClause1 IsNot Nothing Then
                For i = 0 To WhereClause1.Keys.Count - 1
                    Dim mkey As String = WhereClause1.Keys(i)
                    Dim mvalue As Object = GF1.GetValueFromHashTable(WhereClause1, mkey)
                    Dim mtype As String = LCase(mvalue.GetType.Name)
                    Dim mcond As String = ""

                    Select Case mtype
                        Case "string", "datetime"
                            mvalue = mvalue.Replace("'", " ")
                            mcond = mkey & " = '" & mvalue & "'"
                        Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                            mcond = mkey & " = " & mvalue
                    End Select
                    lwhere1 = lwhere1 & IIf(lwhere1.Length = 0, "", " and ") & mcond
                Next
                QryStr = QryStr & IIf(lwhere1.Length > 0, IIf(WhereClause.Length > 0, " and ", " where ") & lwhere1, "")
            End If
            Dim aa1 As Integer = SqlExecuteNonQuery(mSqlTransanction, QryStr, asqlparam)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.UpdateValuesToSqlRecords(ByVal mSqlTransanction As SqlTransaction, ByVal ValueSource As DataRow, ByVal TargetTable As String, ByVal WhereClause As String)")
        End Try
    End Sub


    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="Ltransaction">Sql Transaction already specified</param>
    ''' <param name="LTableName">Sql Table Nale</param>
    ''' <param name="PrimaryKeyName">Primary Key field name on which searching to be done</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & LTableName & " where " & PrimaryKeyName & " = '" & SearchKeyValue & "'"
        Try
            Dim dt As DataTable = SqlExecuteDataTable(Ltransaction, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As String) As DataRow")
        End Try
        Return Nothing
    End Function


    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="LConnection">Sql Connection already specified</param>
    ''' <param name="LTableName">Sql Table Nale</param>
    ''' <param name="PrimaryKeyName">Primary Key field name on which searching to be done</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef LConnection As SqlConnection, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & LTableName & " where " & PrimaryKeyName & " = '" & SearchKeyValue & "'"
        Try
            Dim dt As DataTable = SqlExecuteDataTable(LConnection, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef Lconnection As SqlConnection, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As String) As DataRow")
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="Ltransaction">Sql Transaction already specified</param>
    ''' <param name="LTableName">Sql Table Nale</param>
    ''' <param name="PrimaryKeyName">Primary Key field name on which searching to be done</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & LTableName & " where " & PrimaryKeyName & " = " & SearchKeyValue.ToString
        Try
            Dim dt As DataTable = SqlExecuteDataTable(Ltransaction, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="LConnection">SqlConnection already specified</param>
    ''' <param name="LTableName">Sql Table Nale</param>
    ''' <param name="PrimaryKeyName">Primary Key field name on which searching to be done</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef LConnection As SqlConnection, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & LTableName & " where " & PrimaryKeyName & " = " & SearchKeyValue.ToString
        Try
            Dim dt As DataTable = SqlExecuteDataTable(LConnection, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef LConnection As SqlConnection, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="Ltransaction">Sql Transaction already specified</param>
    ''' <param name="ClsObject">Sql Table Nale</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef Ltransaction As SqlTransaction, ByRef ClsObject As Object, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & ClsObject.TableName & " where " & ClsObject.PrimaryKey & " = " & SearchKeyValue.ToString
        Try
            Dim dt As DataTable = SqlExecuteDataTable(Ltransaction, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="LConnection">SqlConnection already specified</param>
    ''' <param name="ClsObject">Sql Table class object</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef LConnection As SqlConnection, ByRef ClsObject As Object, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & ClsObject.TableName & " where " & ClsObject.PrimaryKey & " = " & SearchKeyValue.ToString
        Try
            Dim dt As DataTable = SqlExecuteDataTable(LConnection, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef LConnection As SqlConnection, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="Ltransaction">Sql Transaction already specified</param>
    ''' <param name="ClsObject">Sql Table Nale</param>
    ''' <param name="WhereClause" >A hashtable to define whereclause ,where key is fieldname and value is field value</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef Ltransaction As SqlTransaction, ByRef ClsObject As Object, ByVal WhereClause As Hashtable) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lwhere As String = GF1.GetStringConditionFromHashTable(WhereClause, True)

        Dim SqlStr As String = "select * from " & ClsObject.TableName & IIf(lwhere.Trim.Length = 0, "", " where " & lwhere)
        Try
            Dim dt As DataTable = SqlExecuteDataTable(Ltransaction, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim mrow As DataRow = ClsObject.NewRow
                    mrow = UpdateDataRows(mrow, dt.Rows(0))
                    Dim mgroupfieldstype As Hashtable = ClsObject.groupfieldstype
                    mrow = ReplaceGroupFieldsValueInRow(mgroupfieldstype, mrow, Nothing, Nothing)
                    ClsObject.PrevRow.ItemArray = mrow.ItemArray
                    Return mrow
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="LConnection">SqlConnection already specified</param>
    ''' <param name="ClsObject">Sql Table Class object</param>
    ''' <param name="WhereClause" >A hashtable to define whereclause ,where key is fieldname and value is field value</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef LConnection As SqlConnection, ByRef ClsObject As Object, ByVal WhereClause As Hashtable) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lwhere As String = GF1.GetStringConditionFromHashTable(WhereClause, True)

        Dim SqlStr As String = "select * from " & ClsObject.TableName & IIf(lwhere.Trim.Length = 0, "", " where " & lwhere)
        Try
            Dim dt As DataTable = SqlExecuteDataTable(LConnection, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Dim mrow As DataRow = ClsObject.NewRow
                    mrow = UpdateDataRows(mrow, dt.Rows(0))
                    Dim mgroupfieldstype As Hashtable = ClsObject.groupfieldstype
                    mrow = ReplaceGroupFieldsValueInRow(mgroupfieldstype, mrow, Nothing, Nothing)
                    ClsObject.PrevRow.ItemArray = mrow.ItemArray
                    Return mrow
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef LConnection As SqlConnection, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ClsObject">Sql Table Nale</param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByRef ClsObject As Object, ByVal SearchKeyValue As Integer) As DataRow

        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SqlStr As String = "select * from " & ClsObject.TableName & " where " & ClsObject.PrimaryKey & " = " & SearchKeyValue.ToString
        Try
            Dim dt As DataTable = SqlExecuteDataTable(ClsObject.ServerDataBase, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByRef Ltransaction As SqlTransaction, ByVal LTableName As String, ByVal PrimaryKeyName As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function


    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <param name="PrimaryKey" >Source Table Primary key field name </param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>DataRow of keyValue</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal PrimaryKey As String, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
        SourceTableName = ConvertFromSrv0Mdf0(SourceTableName)
        Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
        Dim SqlStr As String = "select * from " & SourceTableName & " where " & PrimaryKey & " = " & SearchKeyValue
        Try
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal KeyField As String, ByVal SearchKeyValue As integer) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier">full sql table identifier i.e. server.database.dbo.table or 0_srv_0.0_mdf_0.dbo.Table format </param>
    ''' <param name="PrimaryKey" >Source Table Primary key field name </param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>DataRow of keyValue</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal SqlTableFullIdentifier As String, ByVal PrimaryKey As String, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LServerDataBase As String = GetServerDataBase(GetServerDataBaseFromSqlIdentifier(SqlTableFullIdentifier))
        Dim SourceTableName As String = GetTableNameFromSqlIdentifier(SqlTableFullIdentifier)
        Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
        Dim SqlStr As String = "select * from " & SourceTableName & " where " & PrimaryKey & " = " & SearchKeyValue
        Try
            Dim dt As DataTable = SqlExecuteDataTable(sqlcon, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function

                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal SqlTableFullIdentifier As String, ByVal PrimaryKey As String, ByVal SearchKeyValue As Integer) As DataRow")
        End Try
        sqlcon.Dispose()
        Return Nothing
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier">full sql table identifier i.e. server.database.dbo.table or 0_srv_0.0_mdf_0.dbo.Table format </param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>DataRow of keyValue</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal SqlTableFullIdentifier As String, ByVal SearchKeyValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LServerDataBase As String = GetServerDataBase(GetServerDataBaseFromSqlIdentifier(SqlTableFullIdentifier))
        Dim SourceTableName As String = GetTableNameFromSqlIdentifier(SqlTableFullIdentifier)
        Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
        Dim PrimaryKey As String = GetPrimaryKey(sqlcon, SourceTableName)
        Dim SqlStr As String = "select * from " & SourceTableName & " where " & PrimaryKey & " = " & SearchKeyValue
        Try
            Dim dt As DataTable = SqlExecuteDataTable(sqlcon, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal SqlTableFullIdentifier As String, ByVal PrimaryKey As String, ByVal SearchKeyValue As Integer) As DataRow")
        End Try
        sqlcon.Dispose()
        Return Nothing
    End Function
    ''' <summary>
    ''' delete  row   from a Sql table
    ''' </summary>
    ''' <param name="SqlTableFullIdentifier">full sql table identifier i.e. server.database.dbo.table or 0_srv_0.0_mdf_0.dbo.Table format </param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>no. of record deleted successfully</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecord(ByVal SqlTableFullIdentifier As String, ByVal SearchKeyValue As Integer) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LServerDataBase As String = GetServerDataBase(GetServerDataBaseFromSqlIdentifier(SqlTableFullIdentifier))
        Dim SourceTableName As String = GetTableNameFromSqlIdentifier(SqlTableFullIdentifier)
        Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
        Dim PrimaryKey As String = GetPrimaryKey(sqlcon, SourceTableName)
        Dim SqlStr As String = "delete  from " & SourceTableName & " where " & PrimaryKey & " = " & SearchKeyValue
        Dim i As Integer = 0
        Try
            i = SqlExecuteNonQuery(sqlcon, SqlStr)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal SqlTableFullIdentifier As String, ByVal PrimaryKey As String, ByVal SearchKeyValue As Integer) As DataRow")
        End Try
        sqlcon.Dispose()
        Return i
    End Function
    ''' <summary>
    ''' To delete  row from Sql Table on specified primary key. 
    ''' </summary>
    ''' <param name="ClsObject" >Class object of Sql table</param>
    ''' <param name="PrimaryKeyValue" >Primary key value to be searched</param>
    ''' <returns>Rows affected</returns>
    ''' <remarks></remarks>
    Public Function DeleteRecord(ByRef ClsObject As Object, ByVal PrimaryKeyValue As Integer) As Integer
        ' Public Function GetDataFromSql(ByVal ServerDataBase As String, ByVal Ltable As String, ByVal LfieldList As String, ByVal LJoinStmt As String, ByVal Lcondition As String, ByVal LFilter As String, ByVal Lorder As String, Optional ByVal RecordPosition As String = "*", Optional ByVal PrimaryCols As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LserverDatabase As String = GetServerDataBase(ClsObject.ServerDataBase)
        Dim LTable As String = ClsObject.TableName
        Dim ii As Integer = DeleteRecord(LserverDatabase & ".dbo." & LTable, PrimaryKeyValue)
        Return ii
    End Function



    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <param name="PrimaryKey" >Source Table Primary key field name </param>
    ''' <param name="SearchKeyValue">Primary key field value of row to be fetched</param>
    ''' <returns>DataRow of keyValue</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal PrimaryKey As String, ByVal SearchKeyValue As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
        SourceTableName = ConvertFromSrv0Mdf0(SourceTableName)
        Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
        Dim SqlStr As String = "select * from " & SourceTableName & " where " & PrimaryKey & " = '" & SearchKeyValue & "'"
        Try
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, SqlStr)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    Return dt.Rows(0)
                    Exit Function
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal KeyField As String, ByVal SearchKeyValue As string) As DataRow")
        End Try
        Return Nothing
    End Function


    ''' <summary>
    ''' Creates the backup of database (.mdf and .ldf files) at the target folder
    ''' </summary>
    ''' <param name="servername">Servername</param>
    ''' <param name="databasename">Databasename without mdf for ex. saralweb</param>
    ''' <param name="sourcefolder">Source folder where the database is located</param>
    ''' <param name="targetfolder">Target folder location where you want to copy database files</param>
    ''' <remarks>returns exception message if error occurs else return ""</remarks>
    Public Function BackupDatabaseFiles(ByVal servername As String, ByVal databasename As String, ByVal sourcefolder As String, ByVal targetfolder As String) As String
        Dim databasefullname As String = sourcefolder & "\" & databasename & ".mdf"
        Dim result As String = ""
        Try
            ' firstly for detach
            Dim res As Boolean = DetachDataBase(servername, databasename)
            'to copy ldf file
            Dim sourcefile, destinationfile As String
            Dim dateAndTime As String
            dateAndTime = getDateTimeISTNow.ToString("_yyyy_MM_dd_hh_mm_ss")
            Dim copylogfile As String = databasename & dateAndTime & "_log.ldf"
            Dim srcfile As String = databasename & "_log.ldf"
            sourcefile = Path.Combine(sourcefolder, srcfile)
            destinationfile = Path.Combine(targetfolder, copylogfile)
            System.IO.File.Copy(sourcefile, destinationfile, True)
            'to copy mdf file
            Dim srcmdf As String = databasename & ".mdf"
            Dim copyMDFfile As String = databasename & dateAndTime & ".mdf"
            sourcefile = Path.Combine(sourcefolder, srcmdf)
            destinationfile = Path.Combine(targetfolder, copyMDFfile)
            System.IO.File.Copy(sourcefile, destinationfile, True)
            'to attach database
            AttachDataBase(servername, databasefullname, True)
        Catch ex As Exception
            AttachDataBase(servername, databasefullname, True)
            result = ex.Message
        End Try
        Return result
    End Function
    ''' <summary>
    ''' Fetch  all  values of a fieldname in a string array,return an sql  condition expression for where clause
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <param name="SourceFieldName" >Name of field of source table which values to be considered</param>
    ''' <param name="TargetFieldName" >Name of field of condition expression</param>
    '''<param name="mOperator" >Equaily operator i.e  "="  or  "#" for not equal  </param>
    ''' <param name="mLogicGate" >AND or  OR  used in condition expression</param>
    ''' <param name="FieldValues" >A string arrays of field values</param>
    ''' <returns>Condition clause as string</returns>
    ''' <remarks></remarks>
    Public Function GetConditionExpressionOnAllRows(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SourceFieldName As String, ByVal TargetFieldName As String, ByVal mOperator As String, ByVal mLogicGate As String, ByRef FieldValues() As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim valuestring As String = ""
        Try
            Dim dt As DataTable = GetDataFromSql(ServerDataBase, SourceTableName, SourceFieldName, "", "", "", SourceFieldName)
            valuestring = GetConditionExpressionOnAllRows(dt, SourceFieldName, TargetFieldName, mOperator, mLogicGate, FieldValues)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetConditionExpressionOnAllRows(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SourceFieldName As String, ByVal TargetFieldName As String, ByVal mOperator As String, ByVal mLogicGate As String, ByRef FieldValues() As String) As String")
        End Try
        Return valuestring
    End Function
    ''' <summary>
    ''' Fetch  all  values of a fieldname in a string array,return an sql  condition expression.
    ''' </summary>
    ''' <param name="LdataTable" >DataTable as datatable </param>
    ''' <param name="SourceFieldName" >Name of field of source table which values to be considered</param>
    ''' <param name="TargetFieldName" >Name of field of condition expression</param>
    '''<param name="mOperator" >Equaily operator i.e  "="  or  "#" for not equal  </param>
    ''' <param name="mLogicGate" >AND or  OR  used in condition expression</param>
    ''' <param name="FieldValues" >A string arrays of field values</param>
    ''' <returns>Condition clause as string</returns>
    ''' <remarks></remarks>
    Public Function GetConditionExpressionOnAllRows(ByVal LdataTable As DataTable, ByVal SourceFieldName As String, ByVal TargetFieldName As String, ByVal mOperator As String, ByVal mLogicGate As String, ByRef FieldValues() As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim valuestring As String = ""
        Try
            mOperator = IIf(mOperator = "#", "<>", mOperator)
            If LdataTable IsNot Nothing Then
                For i = 0 To LdataTable.Rows.Count - 1
                    Dim eValue As String = LdataTable.Rows(i).Item(SourceFieldName).ToString
                    GF1.ArrayAppend(FieldValues, eValue)
                    valuestring = valuestring & IIf(valuestring.Length = 0, "", " " & mLogicGate & " ") & TargetFieldName & " " & mOperator & " '" & eValue & "'"
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetConditionExpressionOnAllRows(ByVal LdataTable As DataTable, ByVal SourceFieldName As String, ByVal TargetFieldName As String, ByVal mOperator As String, ByVal mLogicGate As String, ByRef FieldValues() As String) As String")
        End Try
        Return valuestring
    End Function

    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <param name="SearchKeyValue">A hash table containg Primary key field as key and value as item</param>
    ''' <returns>data row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekDataRow(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SearchKeyValue As Hashtable) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SearchKeyRow As DataRow = Nothing
        Try
            Dim lwhere As String = GF1.GetStringConditionFromHashTable(SearchKeyValue, True)
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            SourceTableName = ConvertFromSrv0Mdf0(SourceTableName)
            Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
            Dim SqlStr As String = "select * from " & SourceTableName & " where " & lwhere
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, SqlStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    SearchKeyRow = dt.Rows(0)
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SearchKeyValue As Hashtable, Optional ByRef SearchKeyRow As DataRow = Nothing) As Boolean")

        End Try
        Return SearchKeyRow
    End Function

    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <param name="SearchKeyValue">A hash table containg Primary key field as key and value as item</param>
    ''' <param name="SearchKeyRow">DataRow to hold searhed result</param>
    ''' <returns>True if row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SearchKeyValue As Hashtable, Optional ByRef SearchKeyRow As DataRow = Nothing) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim seekrow As Boolean = False
        Try
            Dim lwhere As String = GF1.GetStringConditionFromHashTable(SearchKeyValue, True)

            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            SourceTableName = ConvertFromSrv0Mdf0(SourceTableName)
            Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)

            ' Dim mPrimaryKey As String = LCase(SearchKeyValue.Keys(0))
            'Dim mValue As String = SearchKeyValue.Item(mPrimaryKey)
            'Dim Pktype As String = LCase(mValue.GetType.Name)
            '  Dim SqlStr As String = "select * from " & SourceTableName & " where " & mPrimaryKey & " = " & IIf(Pktype = "string", "'" & mValue & "'", mValue)
            Dim SqlStr As String = "select * from " & SourceTableName & " where " & lwhere
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, SqlStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    SearchKeyRow = dt.Rows(0)
                    seekrow = True
                End If
            End If
            Return seekrow
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SearchKeyValue As Hashtable, Optional ByRef SearchKeyRow As DataRow = Nothing) As Boolean")

        End Try
    End Function

    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <param name="Lcondition">A string type sql condition</param>
    ''' <returns>Data row is found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal Lcondition As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SearchKeyRow As DataRow = Nothing
        Try
            Dim lwhere As String = Lcondition
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            SourceTableName = ConvertFromSrv0Mdf0(SourceTableName)
            Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
            Dim SqlStr As String = "select * from " & SourceTableName & " where " & lwhere
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, SqlStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    SearchKeyRow = dt.Rows(0)
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal Lcondition As String) As DataRow")

        End Try
        Return SearchKeyRow
    End Function
    ''' <summary>
    ''' Fetch  row as datarow  from a Sql table
    ''' </summary>
    ''' <param name="ServerDataBase">Full identifier of a database with server name eg. server0.database0 or _srv_0._mdf_0 format</param>
    ''' <param name="SearchKeyValue">A hash table containg Primary key field as key and value as item</param>
    ''' <param name="SourceTableName">Source table name or full identifier table name</param>
    ''' <returns> datarow found</returns>
    ''' <remarks></remarks>
    Public Function SeekRecord(ByVal ServerDataBase As String, ByVal SearchKeyValue As Hashtable, ByVal SourceTableName As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SearchKeyRow As DataRow = Nothing
        Try
            Dim lwhere As String = GF1.GetStringConditionFromHashTable(SearchKeyValue, True)
            Dim LServerDataBase As String = GetServerDataBase(ServerDataBase)
            SourceTableName = ConvertFromSrv0Mdf0(SourceTableName)
            Dim sqlcon As SqlConnection = OpenSqlConnection(LServerDataBase)
            'Dim mPrimaryKey As String = LCase(SearchKeyValue.Keys(0))
            'Dim mValue As String = SearchKeyValue.Item(mPrimaryKey)
            'Dim Pktype As String = LCase(mValue.GetType.Name)
            'Dim SqlStr As String = "select * from " & SourceTableName & " where " & mPrimaryKey & " = " & IIf(Pktype = "string", "'" & mValue & "'", mValue)
            Dim SqlStr As String = "select * from " & SourceTableName & " where " & lwhere
            Dim dt As DataTable = SqlExecuteDataTable(LServerDataBase, SqlStr)
            If dt IsNot Nothing Then
                If dt.Rows.Count > 0 Then
                    SearchKeyRow = dt.Rows(0)
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SeekRecord(ByVal ServerDataBase As String, ByVal SourceTableName As String, ByVal SearchKeyValue As Hashtable, Optional ByRef SearchKeyRow As DataRow = Nothing) As Boolean")

        End Try
        Return SearchKeyRow

    End Function


    ''' <summary>
    ''' To sort a datatable on specified columns
    ''' </summary>
    ''' <param name="DtTable">DataTable to be sorted</param>
    ''' <param name="SortedOnColumns">Comma separated string of column names on which sorting executed eg "Col1,Col2,col3" ,if SortOrder not defined it is ascending by default </param>
    ''' <param name="SortingDirection">Default value is "ASC" ,Permissible values are ASC,DESC and * =Set sort direction with column name eg. "column1 ASC ,column2 DESC" etc.</param>
    ''' <returns>Return sorted datatable</returns>
    ''' <remarks></remarks>
    Public Function SortDataTable(ByVal DtTable As DataTable, ByVal SortedOnColumns As String, Optional ByVal SortingDirection As String = "ASC") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If SortingDirection.Trim.Length > 0 Then
            If UCase(SortingDirection) <> "*" And UCase(SortingDirection) <> "ASC" And UCase(SortingDirection) <> "DESC" Then
                QuitMessage("Invalid sorting direction " & SortingDirection, "SortDataTable(ByVal DtTable As DataTable, ByVal SortedOnColumns As String, Optional ByVal SortingDirection As String = ASC) As DataTable  ")
                Return DtTable
                Exit Function
            End If
            SortingDirection = IIf(SortingDirection = "*", "", SortingDirection)
        End If
        Try
            If GlobalControl.Variables.AuthenticationChecked = False Then Return Nothing : Exit Function
            If DtTable.Rows.Count > 0 Then
                Dim LSort() As String = Split(SortedOnColumns, ",")
                SortedOnColumns = Join(LSort, " " & SortingDirection.Trim & ",") & " " & SortingDirection.Trim
                Dim dtView As New DataView(DtTable)
                dtView.Sort = SortedOnColumns
                Dim SortedTable As Data.DataTable = dtView.ToTable("SortedDataTable")
                Return SortedTable
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SortDataTable(ByVal DtTable As DataTable, ByVal SortedOnColumns As String, Optional ByVal SortingDirection As String = "") As DataTable")
        End Try
        Return DtTable
    End Function
    ''' <summary>
    ''' This function sorts and filter a datatable on  columns and expression  specified
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be sorted</param>
    ''' <param name="SortColumns">An array of column names on which datatable sorted</param>
    ''' <param name="SortingDirection">Default value is "ASC" ,Permissible values are ASC,DESC and * =Set sort direction with column name eg. "column1 ASC ,column2 DESC" etc.</param>
    ''' <param name="FilterString">Filter expression string eg "Column1 = 'Value1" </param>
    ''' <returns>Sorted DataTable</returns>
    ''' <remarks></remarks>
    Public Function SortFilterDataTable(ByVal LDataTable As DataTable, ByVal SortColumns() As String, Optional ByVal SortingDirection As String = "ASC", Optional ByVal FilterString As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable
        If SortingDirection.Trim.Length > 0 Then
            If UCase(SortingDirection) <> "*" And UCase(SortingDirection) <> "ASC" And UCase(SortingDirection) <> "DESC" Then
                QuitMessage("Invalid sorting direction " & SortingDirection, "SortFilterDataTable(ByVal LDataTable As DataTable, ByVal SortColumns() As String, Optional ByVal SortingDirection As String = ASC, Optional ByVal FilterString As String = "") As DataTable  ")
                Return mDataTable
                Exit Function
            End If
            SortingDirection = IIf(SortingDirection = "*", "", SortingDirection)
        End If
        Try
            If LDataTable.Rows.Count > 0 Then
                Dim Masc As String = " " & SortingDirection & " ,"
                Dim SortString As String = ""
                If SortColumns IsNot Nothing Then
                    SortString = IIf(SortColumns.Count > 0, Join(SortColumns, Masc) & " " & SortingDirection.Trim, "")
                End If
                Dim mRows() As DataRow = mDataTable.Select(FilterString, SortString)
                If mRows.Count > 0 Then
                    mDataTable = mRows.CopyToDataTable
                Else
                    mDataTable.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SortFilterDataTable(ByVal LDataTable As DataTable, ByVal SortColumns() As String, Optional ByVal SortingDirection As String = "", Optional ByVal FilterString As String = "") As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' This function sorts and filter a datatable on  columns and expression  specified
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be sorted</param>
    ''' <param name="SortColumns">Comma separated column names on which datatable sorted</param>
    ''' <param name="SortingDirection">Default value is "ASC" ,Permissible values are ASC,DESC and * =Set sort direction with column name eg. "column1 ASC ,column2 DESC" etc.</param>
    ''' <param name="FilterString">Filter expression string eg "Column1 = 'Value1" </param>
    ''' <returns>Sorted DataTable</returns>
    ''' <remarks></remarks>
    Public Function SortFilterDataTable(ByVal LDataTable As DataTable, ByVal SortColumns As String, Optional ByVal SortingDirection As String = "ASC", Optional ByVal FilterString As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Copy
        If SortingDirection.Trim.Length > 0 Then
            If UCase(SortingDirection) <> "*" And UCase(SortingDirection) <> "ASC" And UCase(SortingDirection) <> "DESC" Then
                QuitMessage("Invalid sorting direction " & SortingDirection, "SortFilterDataTable(ByVal LDataTable As DataTable, ByVal SortColumns As String, Optional ByVal SortingDirection As String = ASC, Optional ByVal FilterString As String = "") As DataTable  ")
                Return mDataTable
                Exit Function
            End If
            SortingDirection = IIf(SortingDirection = "*", "", SortingDirection)
        End If
        Try
            If LDataTable.Rows.Count > 0 Then

                Dim Masc As String = " " & SortingDirection & " ,"
                Dim SortString As String = ""
                If SortColumns IsNot Nothing Then
                    Dim asortcolumns() As String = Split(SortColumns, ",")
                    SortString = IIf(SortColumns.Count > 0, Join(asortcolumns, Masc) & " " & SortingDirection, "")
                End If
                Dim mRows() As DataRow = mDataTable.Select(FilterString, SortString)
                If mRows.Count > 0 Then
                    mDataTable = mRows.CopyToDataTable
                Else
                    mDataTable.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SortFilterDataTable(ByVal LDataTable As DataTable, ByVal SortColumns As String, Optional ByVal SortingDirection As String = "", Optional ByVal FilterString As String = "") As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' Search datatable on a specified condition.
    ''' </summary>
    ''' <param name="LDataTable">Table to be saerched </param>
    ''' <param name="FilterString">Filter condition as string eg. column1=Value1 and column2> Value2 etc.</param>
    ''' <param name="OnlyFirst">False if return datatable contains only first row of specified criteria</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SearchDataTable(ByVal LDataTable As DataTable, ByVal FilterString As String, Optional ByVal OnlyFirst As Boolean = False) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Copy
        Try
            If LDataTable.Rows.Count > 0 Then
                Dim mRows() As DataRow = mDataTable.Select(FilterString)
                If mRows.Count > 0 Then
                    If OnlyFirst = True Then
                        Dim mrows1() As DataRow = {}
                        GF1.ArrayAppend(mrows1, mRows(0))
                        mDataTable = mrows1.CopyToDataTable
                    Else
                        mDataTable = mRows.CopyToDataTable
                    End If
                Else
                    mDataTable.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTable(ByVal LDataTable As DataTable, ByVal FilterString As String, Optional ByVal OnlyFirst As Boolean = False) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' Search datatable on a specified condition and return first row.
    ''' </summary>
    ''' <param name="LDataTable">Table to be saerched </param>
    ''' <param name="FilterString">Filter condition as string eg. column1=Value1 and column2> Value2 etc.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal FilterString As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Copy
        Try
            If LDataTable.Rows.Count > 0 Then
                Dim mRows() As DataRow = mDataTable.Select(FilterString)
                If mRows.Count > 0 Then
                    Dim mrows1() As DataRow = {}
                    GF1.ArrayAppend(mrows1, mRows(0))
                    mDataTable = mrows1.CopyToDataTable
                    Return mDataTable(0)
                    Exit Function
                Else
                    mDataTable.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal FilterString As String) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Get Cell value from a data table.
    ''' </summary>
    ''' <param name="LDataTable">Data table </param>
    ''' <param name="LColumnName" >Column Name of cell</param>
    ''' <param name="RowNo" >Row no of cell</param>
    ''' <param name="ValueType">Permissible values are String,Integer,Date and default is Object</param> 
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCellValue(ByVal LDataTable As DataTable, ByVal LColumnName As String, Optional ByVal RowNo As Integer = 0, Optional ByVal ValueType As String = "object") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mvalue As Object = Nothing
        Select Case LCase(ValueType)
            Case "string"
                mvalue = ""
            Case "integer"
                mvalue = -1
            Case "date"
                mvalue = New Date(1900, 1, 1)
        End Select

        If LDataTable Is Nothing Then
            Return mvalue
            Exit Function
        End If
        If LDataTable.Rows.Count = 0 Then
            Return mvalue
            Exit Function
        End If
        If RowNo > LDataTable.Rows.Count - 1 Then
            Return mvalue
            Exit Function
        End If
        If CheckColumnInDataTable(LColumnName, LDataTable) = -1 Then
            Return mvalue
            Exit Function
        End If
        Try
            If IsDBNull(LDataTable.Rows(RowNo).Item(LColumnName)) = False Then
                mvalue = LDataTable.Rows(RowNo).Item(LColumnName)
                Select Case LCase(ValueType)
                    Case "string"
                        mvalue = mvalue.ToString
                    Case "integer"
                        mvalue = CInt(mvalue)
                    Case "date"
                        mvalue = CDate(mvalue)
                End Select
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetCellValue(ByVal LDataTable As DataTable, ByVal LColumnName As String, Optional ByVal RowNo As Integer = 0, Optional ByVal ValueType As String = object) As Object")
        End Try
        Return mvalue
    End Function
    ''' <summary>
    ''' Get Cell value from a data table.
    ''' </summary>
    ''' <param name="LDataRow">DataRow of cell </param>
    ''' <param name="LColumnName" >Column Name of cell</param>
    ''' <param name="ValueType">Permissible values are String,Integer,Date and default is Object</param> 
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCellValue(ByVal LDataRow As DataRow, ByVal LColumnName As String, Optional ByVal ValueType As String = "object") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mvalue As Object = Nothing
        Select Case LCase(ValueType)
            Case "string"
                mvalue = ""
            Case "integer"
                mvalue = -1
            Case "date"
                mvalue = New Date(1900, 1, 1)
        End Select
        If LDataRow Is Nothing Then
            Return mvalue
            Exit Function
        End If
        If CheckColumnInDataTable(LColumnName, LDataRow.Table) = -1 Then
            Return mvalue
            Exit Function
        End If
        Try
            If IsDBNull(LDataRow.Item(LColumnName)) = False Then
                mvalue = LDataRow.Item(LColumnName)
                Select Case LCase(ValueType)
                    Case "string"
                        mvalue = mvalue.ToString
                    Case "integer"
                        mvalue = CInt(mvalue)
                    Case "date"
                        mvalue = CDate(mvalue)
                End Select
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetCellValue(ByVal LDataRow As DataRow, ByVal LColumnName As String, Optional ByVal ValueType As String = object) As Object")
        End Try
        Return mvalue
    End Function


    ''' <summary>
    ''' Search DataTable on a specified condition
    ''' </summary>
    ''' <param name="LDataTable"></param>
    ''' <param name="FilterString"></param>
    ''' <param name="OnlyFirst"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal FilterString As String, Optional ByVal OnlyFirst As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mindx() As Integer = {}
        Try
            If LDataTable.Rows.Count > 0 Then
                Dim mRows() As DataRow = LDataTable.Select(FilterString)
                If mRows.Count > 0 Then
                    If OnlyFirst = True Then
                        Dim ind As Integer = CheckRowInDataTable(mRows(0), LDataTable)
                        If ind > -1 Then
                            GF1.ArrayAppend(mindx, ind)
                            Return mindx
                            Exit Function
                        End If
                    Else
                        For i = 0 To mRows.Count - 1
                            Dim ind As Integer = CheckRowInDataTable(mRows(i), LDataTable)
                            If ind > -1 Then
                                GF1.ArrayAppend(mindx, ind)
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal FilterString As String, Optional ByVal OnlyFirst As Boolean = False) As Integer() " & FilterString)
        End Try
        Return mindx
    End Function


    ''' <summary>
    ''' Fetch the array of column values from the given datarow and columnames array
    ''' </summary>
    ''' <param name="LDataRow ">DataTable to be sorted</param>
    ''' <param name="RowColumns ">An array of column namesd</param>
    ''' <returns>Array of object values</returns>
    ''' <remarks></remarks>
    Public Function GetDataRowValuesByColumns(ByVal LDataRow As DataRow, ByVal RowColumns() As String) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mValues() As Object = {}
        Try
            For i = 0 To RowColumns.Count - 1
                GF1.ArrayAppend(mValues, LDataRow(RowColumns(i)))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetDataRowValuesByColumns(ByVal LDataRow As DataRow, ByVal RowColumns() As String) As Object()")
        End Try
        Return mValues
    End Function
    ''' <summary>
    ''' This function creates a unique/distinct rows datatable on  specified columns from a datatable to remove duplicacy of specified columns.
    ''' </summary>
    ''' <param name="LDataTable">DataTable with duplicate rows  on specified columns</param>
    ''' <param name="UniqueColumns ">Comma separated column names   specified columns on which distinct clause executed eg. "Column1,Column2, ..  } etc.</param>
    '''<param name="SameColumns" >false if unique file has only unique columns specified,true if returming datatable has same schema</param>
    ''' <returns>Unique rows datatable</returns>
    ''' <remarks></remarks>
    Public Function GetDistinctRowsFromDataTable(ByVal LDataTable As DataTable, ByVal UniqueColumns As String, Optional ByVal SameColumns As Boolean = False) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As New DataTable
        Try
            Dim mView As DataView = LDataTable.DefaultView
            Dim aUniqueColumns() As String = Split(UniqueColumns, ",")
            Dim SortString As String = Join(aUniqueColumns, " ASC,") & " ASC"
            mView.Sort = SortString
            mDataTable = mView.ToTable(True, aUniqueColumns)
            If SameColumns = True Then
                Dim MissingColumns() As String = {}
                For i = 0 To LDataTable.Columns.Count - 1
                    Dim mcolumn As String = LDataTable.Columns(i).ColumnName
                    Dim ColumnExists As Boolean = False
                    For k = 0 To mDataTable.Columns.Count - 1
                        If mcolumn = mDataTable.Columns(k).ColumnName Then
                            ColumnExists = True
                            Exit For
                        End If
                    Next
                    If ColumnExists = False Then
                        Dim newColumn As String = LDataTable.Columns(i).ColumnName
                        mDataTable.Columns.Add(newColumn)
                        GF1.ArrayAppend(MissingColumns, LDataTable.Columns(i).ColumnName)
                    End If
                Next
                For i = 0 To mDataTable.Rows.Count - 1
                    Dim UniqueValues() As Object = GetDataRowValuesByColumns(mDataTable.Rows(i), aUniqueColumns)
                    Dim mrowindex() As Integer = SearchDataTableRowIndex(LDataTable, aUniqueColumns, UniqueValues, True)
                    For k = 0 To MissingColumns.Count - 1
                        mDataTable.Rows(i).Item(MissingColumns(k)) = LDataTable.Rows(mrowindex(0)).Item(MissingColumns(k))
                    Next
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetDistinctRowsFromDataTable(ByVal LDataTable As DataTable, ByVal UniqueColumns As String, Optional ByVal SameColumns As Boolean = False) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' This function adds two datatables of same structure creates a new data table.
    ''' </summary>
    ''' <param name="FirstDataTable">First data table to be added</param>
    ''' <param name="SecondDataTable">Second data table to be added</param>
    ''' <param name="IgnoreDuplicateRows">Ignore duplicate rows</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddTwoDataTables(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable, Optional ByVal IgnoreDuplicateRows As Boolean = False) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If FirstDataTable.Rows.Count = 0 Then
            Return SecondDataTable
            Exit Function
        End If
        If FirstDataTable.Columns.Count <> SecondDataTable.Columns.Count Then
            MsgBox("Columns count not matched" & " AddDataTable")
        End If
        Try
            For i = 0 To SecondDataTable.Rows.Count - 1
                If IgnoreDuplicateRows = True Then
                    Dim j As Integer = CheckRowInDataTable(SecondDataTable(i), FirstDataTable)
                    If j > -1 Then
                        Continue For
                    End If
                End If
                Dim mRow As DataRow = FirstDataTable.NewRow
                For k = 0 To SecondDataTable.Columns.Count - 1
                    mRow(k) = SecondDataTable.Rows(i).Item(k)
                Next
                FirstDataTable.Rows.Add(mRow)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddTwoDataTables(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable, Optional ByVal IgnoreDuplicateRows As Boolean = False) As DataTable")
        End Try
        Return FirstDataTable
    End Function
    ''' <summary>
    ''' This function updates all rows of firstdatatable according to second datatable.
    ''' </summary>
    ''' <param name="FirstDataTable">First data table to be updated.</param>
    ''' <param name="SecondDataTable">Second data table from column's value updated.</param>
    ''' <param name="RemoveExcessRows" >Remove excess rows of first datatable w.r.to second datatable.</param>
    ''' <param name="UpdatingColumns" >Comma separated columnnames of firstdatatable to be updated,if exist in second datatable,* for all</param>
    ''' <param name="IgnoreColumns">Comma separated columnnames of firstdatatable to be ignored for updation</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateDataTables(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable, Optional ByVal RemoveExcessRows As Boolean = True, Optional ByVal UpdatingColumns As String = "*", Optional ByVal IgnoreColumns As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ts As Integer = SecondDataTable.Rows.Count - 1
            Dim tf As Integer = FirstDataTable.Rows.Count - 1
            If RemoveExcessRows = True Then
                If tf > ts Then
                    For i = ts + 1 To tf
                        FirstDataTable.Rows(i).Delete()
                    Next
                    tf = ts
                End If
            End If
            For i = 0 To SecondDataTable.Rows.Count - 1
                If tf <= i And tf > -1 Then
                    For j = 0 To SecondDataTable.Columns.Count - 1
                        Dim mcolumnname As String = SecondDataTable.Columns(j).ColumnName
                        If CheckColumnInDataTable(mcolumnname, FirstDataTable) > -1 Then
                            If InStr(UpdatingColumns, mcolumnname) > 0 Or UpdatingColumns = "*" Then
                                If InStr(IgnoreColumns, mcolumnname) = 0 Then
                                    FirstDataTable.Rows(i).Item(mcolumnname) = SecondDataTable.Rows(i).Item(mcolumnname)
                                End If
                            End If
                        End If
                    Next
                Else
                    Dim mrow As DataRow = FirstDataTable.NewRow
                    For j = 0 To SecondDataTable.Columns.Count - 1
                        Dim mcolumnname As String = SecondDataTable.Columns(j).ColumnName
                        If CheckColumnInDataTable(mcolumnname, FirstDataTable) > -1 Then
                            If InStr(UpdatingColumns, mcolumnname) > 0 Or UpdatingColumns = "*" Then
                                If InStr(IgnoreColumns, mcolumnname) = 0 Then
                                    mrow(mcolumnname) = SecondDataTable.Rows(i).Item(mcolumnname)
                                End If
                            End If
                        End If
                    Next
                    FirstDataTable.Rows.Add(mrow)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.UpdateDataTables(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable, Optional ByVal UpdatingColumns As String = " * ", Optional ByVal IgnoreColumns As String = "") As DataTable")
        End Try
        Return FirstDataTable
    End Function
    ''' <summary>
    ''' This function updates FirstDataRow  according to SecondDataRow.
    ''' </summary>
    ''' <param name="FirstDataRow">First data table to be updated.</param>
    ''' <param name="SecondDataRow">Second data table from column's value updated.</param>
    ''' <param name="UpdatingColumns" >Comma separated columnnames of firstdatarow to be updated,if exist in second datarow,* for all</param>
    ''' <param name="IgnoreColumns">Comma separated columnnames of firstdatarow to be ignored for updation</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateDataRows(ByVal FirstDataRow As DataRow, ByVal SecondDataRow As DataRow, Optional ByVal UpdatingColumns As String = "*", Optional ByVal IgnoreColumns As String = "") As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim firstdt As DataTable = FirstDataRow.Table
            Dim seconddt As DataTable = SecondDataRow.Table

            For j = 0 To seconddt.Columns.Count - 1
                Dim mcolumnname As String = seconddt.Columns(j).ColumnName
                If CheckColumnInDataTable(mcolumnname, firstdt) > -1 Then
                    If InStr(UpdatingColumns, mcolumnname) > 0 Or UpdatingColumns = "*" Then
                        If InStr(IgnoreColumns, mcolumnname) = 0 Then
                            If Not IsDBNull(SecondDataRow(mcolumnname)) Then
                                FirstDataRow(mcolumnname) = SecondDataRow(mcolumnname)
                            End If
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            '  MsgBox(ex.Message)
            QuitError(ex, Err, "Unable to execute DataFunction.UpdateDataRows(ByVal FirstDataRow As DataRow, ByVal SecondDataRow As DataRow, Optional ByVal UpdatingColumns As String = " * ", Optional ByVal IgnoreColumns As String = "") As DataRow")
        End Try
        Return FirstDataRow
    End Function



    ''' <summary>
    ''' This function rename oldcolumnname to newcolumnname of a datatable.
    ''' </summary>
    ''' <param name="LDataTable" >DataTable in which columns renamed</param>
    ''' <param name="OldColumnName" >Column name to be renamed</param>
    ''' <param name="NewColumnName" >New Column Name </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RenameDataTableColumn(ByVal LDataTable As DataTable, ByVal OldColumnName As String, ByVal NewColumnName As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As New DataTable
        Try
            If CheckColumnInDataTable(OldColumnName, LDataTable) < 0 Then
                MsgBox("Column name not exists in datatable " & OldColumnName)
                Return LDataTable
                Exit Function
            End If
            Dim mtype As String = LDataTable.Columns(OldColumnName).DataType.ToString
            LDataTable = AddColumnsInDataTable(LDataTable, NewColumnName, mtype)
            For i = 0 To LDataTable.Rows.Count - 1
                LDataTable.Rows(i).Item(NewColumnName) = LDataTable.Rows(i).Item(OldColumnName)
            Next
            LDataTable.Columns.Remove(OldColumnName)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RenameDataTableColumn(ByVal LDataTable As DataTable, ByVal OldColumnName As String, ByVal NewColumnName As String) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' Compare two data rows 
    ''' </summary>
    ''' <param name="FirstDataRow">First data row to be compared</param>
    ''' <param name="SecondDataRow">Second data row to be compared</param>
    ''' <param name="ExcludeColumns" >Comma separated list of columns to be excluded for comparing datarows</param>
    ''' <returns>If matched return true</returns>
    ''' <remarks></remarks>
    Public Function CompareTwoDataRows(ByVal FirstDataRow As DataRow, ByVal SecondDataRow As DataRow, Optional ByVal ExcludeColumns As String = "") As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim CompareFlag As Boolean = True
        Try
            Dim aColumns() As String = Split(LCase(ExcludeColumns), ",")
            For k = 0 To FirstDataRow.ItemArray.Count - 1
                Dim columnname As String = LCase(FirstDataRow.Table.Columns(k).ColumnName)
                If GF1.ArrayFind(aColumns, columnname) > -1 Then
                    Continue For
                End If
                Dim mType As String = LCase(FirstDataRow(columnname).GetType.Name)
                Select Case mType
                    Case "string"
                        If Not FirstDataRow.Item(columnname).ToString.Trim = SecondDataRow.Item(columnname).ToString.Trim Then
                            CompareFlag = False
                            Exit For
                        End If
                    Case "datarow"
                        CompareFlag = CompareTwoDataRows(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname), ExcludeColumns)
                        If CompareFlag = False Then
                            Exit For
                        End If
                    Case "datatable"
                        CompareFlag = CompareTwoDataTables(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname), ExcludeColumns)
                        If CompareFlag = False Then
                            Exit For
                        End If
                    Case "hashtable"
                        CompareFlag = CompareTwoHashTables(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname), ExcludeColumns)
                        If CompareFlag = False Then
                            Exit For
                        End If
                    Case Else
                        If Equals(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname)) = False Then
                            CompareFlag = False
                            Exit For
                        End If
                End Select
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareTwoDataRows(ByVal FirstDataRow As DataRow, ByVal SecondDataRow As DataRow, Optional ByVal ExcludeColumns As String = "") As Boolean")
        End Try
        Return CompareFlag
    End Function
    ''' <summary>
    ''' Compare two hashtables 
    ''' </summary>
    ''' <param name="FirstHashTable">First hash table to be compared</param>
    ''' <param name="SecondHashTable">Second hash table to be compared</param>
    ''' <param name="ExcludeKeys" >Comma separated list of keys to be excluded for comparing datarows</param>
    ''' <returns>If matched return true</returns>
    ''' <remarks></remarks>
    Public Function CompareTwoHashTables(ByVal FirstHashTable As Hashtable, ByVal SecondHashTable As Hashtable, Optional ByVal ExcludeKeys As String = "") As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim CompareFlag As Boolean = True
        Try
            Dim aColumns() As String = Split(LCase(ExcludeKeys), ",")
            For k = 0 To FirstHashTable.Keys.Count - 1
                Dim mkey As String = LCase(FirstHashTable.Keys(k))
                If GF1.ArrayFind(aColumns, mkey) > -1 Then
                    Continue For
                End If
                Dim mType As String = LCase(FirstHashTable.Item(mkey).GetType.Name)
                Select Case mType
                    Case "string"
                        If Not FirstHashTable.Item(mkey).ToString.Trim = SecondHashTable.Item(mkey).ToString.Trim Then
                            CompareFlag = False
                            Exit For
                        End If
                    Case "datarow"
                        CompareFlag = CompareTwoDataRows(FirstHashTable.Item(mkey), SecondHashTable.Item(mkey), ExcludeKeys)
                        If CompareFlag = False Then
                            Exit For
                        End If
                    Case "datatable"
                        CompareFlag = CompareTwoDataTables(FirstHashTable.Item(mkey), SecondHashTable.Item(mkey), ExcludeKeys)
                        If CompareFlag = False Then
                            Exit For
                        End If
                    Case "hashtable"
                        CompareFlag = CompareTwoHashTables(FirstHashTable.Item(mkey), SecondHashTable.Item(mkey), ExcludeKeys)
                        If CompareFlag = False Then
                            Exit For
                        End If
                    Case Else
                        If Equals(FirstHashTable.Item(mkey), SecondHashTable.Item(mkey)) = False Then
                            CompareFlag = False
                            Exit For
                        End If
                End Select
            Next
           Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareTwoHashTables(ByVal FirstHashTable As Hashtable, ByVal SecondHashTable As Hashtable, Optional ByVal ExcludeKeys As String = "") As Boolean")
        End Try
        Return CompareFlag
    End Function
    ''' <summary>
    ''' Compare two data rows 
    ''' </summary>
    ''' <param name="FirstDataRow">First data row to be compared</param>
    ''' <param name="SecondDataRow">Second data row to be compared</param>
    ''' <param name="ExcludeColumns" >Comma separated list of columns to be excluded for comparing datarows</param>
    ''' <returns>Comma separated column names of first table which values not equal to second table column's value</returns>
    ''' <remarks></remarks>
    Public Function CompareTwoDataRowsValues(ByVal FirstDataRow As DataRow, ByVal SecondDataRow As DataRow, Optional ByVal ExcludeColumns As String = "") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MismatchColumns As String = ""
        Dim CompareFlag As Boolean = False
        Try
            Dim aColumns() As String = Split(LCase(ExcludeColumns), ",")
            For k = 0 To FirstDataRow.ItemArray.Count - 1
                Dim columnname As String = LCase(FirstDataRow.Table.Columns(k).ColumnName)
                If GF1.ArrayFind(aColumns, columnname) > -1 Then
                    Continue For
                End If
                Dim mType As String = LCase(FirstDataRow(k).GetType.Name)
                Select Case mType
                    Case "string"
                        If Not FirstDataRow.Item(columnname).ToString.Trim = SecondDataRow.Item(columnname).ToString.Trim Then
                            MismatchColumns = MismatchColumns & IIf(MismatchColumns.Trim.Length = 0, "", ",") & columnname
                        End If
                    Case "datarow"
                        CompareFlag = CompareTwoDataRows(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname), ExcludeColumns)
                        If CompareFlag = False Then
                            MismatchColumns = MismatchColumns & IIf(MismatchColumns.Trim.Length = 0, "", ",") & columnname
                        End If
                    Case "datatable"
                        CompareFlag = CompareTwoDataTables(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname), ExcludeColumns)
                        If CompareFlag = False Then
                            MismatchColumns = MismatchColumns & IIf(MismatchColumns.Trim.Length = 0, "", ",") & columnname
                        End If
                    Case "hashtable"
                        CompareFlag = CompareTwoHashTables(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname), ExcludeColumns)
                        If CompareFlag = False Then
                            MismatchColumns = MismatchColumns & IIf(MismatchColumns.Trim.Length = 0, "", ",") & columnname
                        End If
                    Case Else
                        If Equals(FirstDataRow.Item(columnname), SecondDataRow.Item(columnname)) = False Then
                            MismatchColumns = MismatchColumns & IIf(MismatchColumns.Trim.Length = 0, "", ",") & columnname
                        End If
                End Select
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareTwoDataRowsFields(ByVal FirstDataRow As DataRow, ByVal SecondDataRow As DataRow, Optional ByVal ExcludeColumns As String = "") As String()")
        End Try
        Return MismatchColumns
    End Function

    ''' <summary>
    ''' Check  row if it exists in a datatable and return its row index ,-1 if not found 
    ''' </summary>
    ''' <param name="SearchingRow">DataRow is being Checked for existence</param>
    ''' <param name="TableToBeSearched">Table object where above row to be searched</param>
    ''' <param name="ExcludeColumns" >Comma separated list of columns to be excluded for comparing datarows</param>
    ''' <returns>Return first row index of table if found, otherwise -1</returns>
    ''' <remarks></remarks>
    Public Function CheckRowInDataTable(ByVal SearchingRow As DataRow, ByVal TableToBeSearched As DataTable, Optional ByVal ExcludeColumns As String = "") As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mIndex As Integer = -1
        Try
            Dim PrimaryCol As String = ""
            If TableToBeSearched.PrimaryKey.Count > 0 Then
                PrimaryCol = TableToBeSearched.PrimaryKey(0).ColumnName
            End If
            If PrimaryCol.Trim.Length = 0 Then
                For k = 0 To TableToBeSearched.Rows.Count - 1
                    If CompareTwoDataRows(SearchingRow, TableToBeSearched.Rows(k), ExcludeColumns) = True Then
                        mIndex = k
                        Exit For
                    End If
                Next
            Else
                Dim mIndex0 As Integer = FindRowIndexByPrimaryCols(TableToBeSearched, SearchingRow(PrimaryCol))
                If mIndex0 > -1 Then
                    If CompareTwoDataRows(SearchingRow, TableToBeSearched.Rows(mIndex0), ExcludeColumns) = True Then
                        mIndex = mIndex0
                    End If
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CheckRowInDataTable(ByVal SearchingRow As DataRow, ByVal TableToBeSearched As DataTable, Optional ByVal ExcludeColumns As String = "") As Integer")
        End Try
        Return mIndex
    End Function
    ''' <summary>
    ''' Check  ColumnName if it exists in a DataTable and return its column index ,-1 if not found 
    ''' </summary>
    ''' <param name="ColumnName" >ColumnName to be searched</param>
    '''<param name="TableToBeSearched" >TableName being searched</param>
    ''' <returns >Return column index in table if found, otherwise -1</returns>
    ''' <remarks ></remarks>
    Public Function CheckColumnInDataTable(ByVal ColumnName As String, ByVal TableToBeSearched As DataTable) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mIndex As Integer = -1
        Try
            For k = 0 To TableToBeSearched.Columns.Count - 1
                If LCase(ColumnName.Trim) = LCase(TableToBeSearched.Columns(k).ColumnName) Then
                    mIndex = k
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CheckColumnInDataTable(ByVal ColumnName As String, ByVal TableToBeSearched As DataTable) As Integer")
        End Try
        Return mIndex
    End Function
    ''' <summary>
    ''' Compare FirstDataTableRows to SecondDataTable rows and return index() of SecondDataTable rows not found in first table
    ''' </summary>
    ''' <param name="FirstDataTable">Table to be searched</param>
    ''' <param name="SecondDataTable">Table being searched</param>
    ''' <returns>Mismatched array of row indexes of Second Data Table as integer</returns>
    ''' <remarks></remarks>
    Public Function CompareTwoDataTableRows(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MismatchedRow() As Integer = {}
        Try
            For k = 0 To SecondDataTable.Rows.Count - 1
                Dim mIndex As Integer = CheckRowInDataTable(SecondDataTable.Rows(k), FirstDataTable)
                If mIndex > -1 Then
                    GF1.ArrayAppend(MismatchedRow, mIndex)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareTwoDataTablesRows(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable) As Integer()")
        End Try
        Return MismatchedRow
    End Function
    ''' <summary>
    ''' Compare FirstDataTableRows to SecondDataTable rows and return index() of SecondDataTable rows not found in first table
    ''' </summary>
    ''' <param name="CurrentDataTable">Table to be searched</param>
    ''' <param name="PreviousDataTable">Table being searched</param>
    ''' <param name="ExcludeColumns" >Comma separated ColumnNames which are excluded from comparing</param>
    ''' <returns>CompareFlag as True or False</returns>
    ''' <remarks></remarks>
    Public Function CompareTwoDataTables(ByVal CurrentDataTable As DataTable, ByVal PreviousDataTable As DataTable, Optional ByVal ExcludeColumns As String = "") As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MismatchedFlag As Boolean = True
        If PreviousDataTable.Rows.Count = 0 Or CurrentDataTable.Rows.Count = 0 Then
            MismatchedFlag = False
            Return MismatchedFlag
            Exit Function
        End If
        If PreviousDataTable.Rows.Count <> CurrentDataTable.Rows.Count Then
            MismatchedFlag = False
            Return MismatchedFlag
            Exit Function
        End If
        Try
            For k = 0 To CurrentDataTable.Rows.Count - 1
                Dim mIndex As Integer = CheckRowInDataTable(CurrentDataTable.Rows(k), PreviousDataTable, ExcludeColumns)
                If mIndex > -1 Then
                    MismatchedFlag = False
                    Return MismatchedFlag
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareTwoDataTables(ByVal CurrentDataTable As DataTable, ByVal PreviousDataTable As DataTable, Optional ByVal ExcludeColumns As String = "") As Boolean")
        End Try
        Return MismatchedFlag
    End Function


    ''' <summary>
    ''' Compare Previous DataTable Rows to Current DataTable Rows and return Row index array of on SecondDataTable rows not found in first table
    ''' </summary>
    ''' <param name="PreviousDataTable">Table to be searched</param>
    ''' <param name="CurrentDataTable">Table being searched</param>
    ''' <param name="ExcludeColumns" >Comma separated list of columns to be excluded for comparing datarows</param>
    ''' <param name="PreviousExtraRows" >An array of row indexes of previous datatable which are missing in current datatable </param>
    ''' <param name="CurrentExtraRows" >An array of row indexes of current datatable which are missing in previous datatable</param>
    ''' <param name="PreviousSameRows" >An array of row indexes of previous datatable which are same in current datatable</param>
    ''' <param name="CurrentSameRows" >An array of row indexes of previous datatable which are same in current datatable</param>
    ''' <returns>An array of row indexes of current datatable which are same in previous datatable</returns>
    ''' <remarks></remarks>
    Public Function CompareTwoDataTablesRows(ByVal PreviousDataTable As DataTable, ByVal CurrentDataTable As DataTable, Optional ByVal ExcludeColumns As String = "", Optional ByRef PreviousExtraRows() As Integer = Nothing, Optional ByRef CurrentExtraRows() As Integer = Nothing, Optional ByRef PreviousSameRows() As Integer = Nothing, Optional ByRef CurrentSameRows() As Integer = Nothing) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mFlag As Boolean = False
        Try
            ReDim PreviousSameRows(-1)
            ReDim PreviousExtraRows(-1)
            ReDim CurrentExtraRows(-1)
            ReDim CurrentSameRows(-1)
            Dim PrimaryCol As String = ""
            If PreviousDataTable.PrimaryKey.Count > 0 Then
                PrimaryCol = PreviousDataTable.PrimaryKey(0).ColumnName
            End If
            Select Case True
                Case PreviousDataTable.Rows.Count > 0 And CurrentDataTable.Rows.Count > 0
                    Try
                        For k = 0 To PreviousDataTable.Rows.Count - 1
                            Dim mIndex As Integer = CheckRowInDataTable(PreviousDataTable.Rows(k), CurrentDataTable, ExcludeColumns)
                            If mIndex < 0 Then
                                GF1.ArrayAppend(PreviousExtraRows, k)
                            Else
                                GF1.ArrayAppend(PreviousSameRows, k)
                            End If
                        Next
                        For k = 0 To CurrentDataTable.Rows.Count - 1
                            Dim mIndex As Integer = CheckRowInDataTable(CurrentDataTable.Rows(k), PreviousDataTable, ExcludeColumns)
                            If mIndex < 0 Then
                                GF1.ArrayAppend(CurrentExtraRows, k)
                            Else
                                GF1.ArrayAppend(CurrentSameRows, k)
                            End If
                        Next
                    Catch ex As Exception
                        GF1.QuitError(ex, Err, "Row comparing failed in CompareTwoDataTablesRows(ByVal PreviousDataTable As DataTable, ByVal CurrentDataTable As DataTable, ByRef PreviousExtraRows() As Integer, ByRef CurrentExtraRows() As Integer, ByRef PreviousSameRows() As Integer, Optional ByVal ExcludeColumns As String = "") As Integer()")
                    End Try
                Case PreviousDataTable.Rows.Count > 0 And CurrentDataTable.Rows.Count = 0
                    For k = 0 To PreviousDataTable.Rows.Count - 1
                        GF1.ArrayAppend(PreviousExtraRows, k)
                    Next
                Case PreviousDataTable.Rows.Count = 0 And CurrentDataTable.Rows.Count > 0
                    For k = 0 To CurrentDataTable.Rows.Count - 1
                        GF1.ArrayAppend(CurrentExtraRows, k)
                    Next
            End Select
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareTwoDataTablesRows(ByVal PreviousDataTable As DataTable, ByVal CurrentDataTable As DataTable, Optional ByVal ExcludeColumns As String = "", Optional ByRef PreviousExtraRows() As Integer = Nothing, Optional ByRef CurrentExtraRows() As Integer = Nothing, Optional ByRef PreviousSameRows() As Integer = Nothing, Optional ByRef CurrentSameRows() As Integer = Nothing) As Boolean")
        End Try
        If PreviousExtraRows.Count > 0 Or CurrentExtraRows.Count > 0 Then
            mFlag = True
        End If

        Return mFlag
    End Function





    ''' <summary>
    ''' Compare columns of two data tables and get missing columns
    ''' </summary>
    ''' <param name="FirstDataTable">First Data Table whoose columns compared</param>
    ''' <param name="SecondDataTable">Second Data Table whoose columns compared</param>
    ''' <param name="MissingColumns">Missing columns in second table as hashtable where key is columnname and value is type</param>
    ''' <param name="MismatchColumnsTYpe" >Mismattched columns types  in second table as hashtable where key is columnname and value is type</param>
    ''' <param name="ExtraColumns">Extra columns in second table as hashtable where key is columnname and value is type</param>
    ''' <param name="SameColumns">Same columns in second table as hashtable where key is columnname and value is type</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CompareDataColumns(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable, Optional ByRef MissingColumns As Hashtable = Nothing, Optional ByRef MismatchColumnsTYpe As Hashtable = Nothing, Optional ByRef ExtraColumns As Hashtable = Nothing, Optional ByRef SameColumns As Hashtable = Nothing) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim CompareFlag As Boolean = True
        Dim mMissingColumns As New Hashtable
        Dim mExtraColumns As New Hashtable
        Dim mSameColumns As New Hashtable
        Dim mMismatchColumnTypes As New Hashtable
        Try
            For i = 0 To FirstDataTable.Columns.Count - 1
                Dim mfirstColumnName As String = FirstDataTable.Columns(i).ColumnName
                Dim mfirstColumnType As Type = FirstDataTable.Columns(i).DataType
                Dim mMissingColumnsFlag As Boolean = True
                Dim mExtraColumnsFlag As Boolean = False
                Dim mSameColumnsFlag As Boolean = False
                Dim mMismatchColumnTypesFlag As Boolean = False
                Dim mSecondColumnName As String = ""
                Dim mSecondColumnType As Type = Nothing
                For k = 0 To SecondDataTable.Columns.Count - 1
                    mSecondColumnName = SecondDataTable.Columns(k).ColumnName
                    mSecondColumnType = SecondDataTable.Columns(k).DataType
                    If mfirstColumnName = mSecondColumnName And Equals(mfirstColumnType, mSecondColumnType) = True Then
                        mMissingColumnsFlag = False
                        mSameColumnsFlag = True
                        Exit For
                    End If
                    If mfirstColumnName = mSecondColumnName And Equals(mfirstColumnType, mSecondColumnType) = False Then
                        mMismatchColumnTypesFlag = True
                        Exit For
                    End If
                Next
                If mSameColumnsFlag = True Then
                    mSameColumns.Add(mfirstColumnName, mfirstColumnType)
                End If
                If mMismatchColumnTypesFlag = True Then
                    mMismatchColumnTypes.Add(mfirstColumnName, mSecondColumnType)
                End If
                If mMissingColumnsFlag = True Then
                    mMissingColumns.Add(mfirstColumnName, mfirstColumnType)
                End If
            Next
            If Not ExtraColumns Is Nothing Then
                For i = 0 To SecondDataTable.Columns.Count - 1
                    Dim msecondColumnName As String = SecondDataTable.Columns(i).ColumnName
                    Dim mSecondType As Type = SecondDataTable.Columns(i).DataType
                    Dim mExtraColumnsFlag As Boolean = True
                    Dim mFirstColumnName As String = ""
                    For k = 0 To FirstDataTable.Columns.Count - 1
                        mFirstColumnName = SecondDataTable.Columns(k).ColumnName
                        If mFirstColumnName = msecondColumnName Then
                            mExtraColumnsFlag = False
                            Exit For
                        End If
                    Next
                    If mExtraColumnsFlag = True Then
                        mExtraColumns.Add(mFirstColumnName, mSecondType)
                    End If
                Next
            End If
            If mMissingColumns.Count > 0 Or mMissingColumns.Count > 0 Or ExtraColumns.Count > 0 Then
                CompareFlag = False
            End If
            MissingColumns = mMissingColumns
            SameColumns = mSameColumns
            MismatchColumnsTYpe = mMismatchColumnTypes
            ExtraColumns = mExtraColumns
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CompareDataColumns(ByVal FirstDataTable As DataTable, ByVal SecondDataTable As DataTable, Optional ByRef MissingColumns As Hashtable = Nothing, Optional ByRef MismatchColumnsTYpe As Hashtable = Nothing, Optional ByRef ExtraColumns As Hashtable = Nothing, Optional ByRef SameColumns As Hashtable = Nothing) As Boolean")
        End Try
        Return CompareFlag
    End Function
    ''' <summary>
    ''' Searching a DataTable row on  columns() with values() and return RowIndex()
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="SearchColumns  ">An array of  column names eg. {"Column1","Column2", ..  } etc.</param>
    '''<param name="SearchValues" >An array of values eg. {value1,value2,value3...  </param>
    '''<param name="OnlyFirstIndex" >True=Only first row index returned to match the criteria </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns() As String, ByVal SearchValues() As Object, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Not SearchColumns.Count = SearchValues.Count Then
            QuitMessage("Size of SearchColumns differs SearchValues" & "   " & "SearchDataTableRowIndex", "SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns() As String, ByVal SearchValues() As Object, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()  ")
        End If
        Dim mRowIndex() As Integer = {}
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                For j = 0 To SearchColumns.Count - 1
                    If Not LDataTable.Rows(i).Item(SearchColumns(j)) = SearchValues(j) Then
                        Matched = False
                    End If
                    If Matched = False Then
                        Exit For
                    End If
                Next
                If Matched = True Then
                    GF1.ArrayAppend(mRowIndex, i)
                    If OnlyFirstIndex = True Then
                        Return mRowIndex
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns() As String, ByVal SearchValues() As Object, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()")
        End Try
        Return mRowIndex
    End Function
    ''' <summary>
    ''' Searching a DataTable row on  a column with value and return RowIndex()
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="SearchColumn  ">column name as string</param>
    '''<param name="SearchValue" >Column value  as object..  </param>
    '''<param name="OnlyFirstIndex" >True=Only first row index returned to match the criteria </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableRowIndexSingleColumn(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Object, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mRowIndex() As Integer = {}
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                If Not UCase(LDataTable.Rows(i).Item(SearchColumn).ToString.Trim) = UCase(SearchValue).ToString.Trim Then
                    Matched = False
                End If
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If OnlyFirstIndex = True Then
                    Return mRowIndex
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableRowIndexSingleColumn(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Object, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()")
        End Try
        Return mRowIndex
    End Function
    ''' <summary>
    ''' Searching a DataTable row on  a column with value and return a datarow
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="SearchColumn  ">column name as string</param>
    '''<param name="SearchValue" >Column value  as object..  </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableRow(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Object) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mRowIndex() As Integer = {}
        Dim OnlyFirstIndex As Boolean = True
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                If Not UCase(LDataTable.Rows(i).Item(SearchColumn).ToString.Trim) = UCase(SearchValue.ToString.Trim) Then
                    Matched = False
                End If
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If OnlyFirstIndex = True Then
                    Return LDataTable(mRowIndex(0))
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableRow(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Object) As DataRow")
        End Try
        Return Nothing

    End Function



    ''' <summary>
    ''' Searching a DataTable row on  columns with values and return RowIndex()
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="SearchColumns  ">Comma separated column name as string</param>
    '''<param name="SearchValues" >Comma separated Column values  as string..  </param>
    '''<param name="OnlyFirstIndex" >True=Only first row index returned to match the criteria </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Split(SearchColumns, ",").Count <> Split(SearchValues, ",").Count Then
            QuitMessage("Invalid count of searchcolumns and searchValues " & SearchColumns & " " & SearchValues, "SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()  ")
        End If
        Dim mRowIndex() As Integer = {}
        Dim aColumns() As String = Split(SearchColumns, ",")
        Dim aValues() As String = Split(SearchValues, ",")
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                For j = 0 To aColumns.Count - 1
                    If Not UCase(LDataTable.Rows(i).Item(aColumns(j)).ToString.Trim) = UCase(aValues(j).ToString.Trim) Then
                        Matched = False
                        Exit For
                    End If
                Next
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If OnlyFirstIndex = True Then
                    Return mRowIndex
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()")
        End Try
        Return mRowIndex
    End Function

    ''' <summary>
    ''' Searching a DataTable row on  columns with values and return RowIndex()
    ''' </summary>
    ''' <param name="LGridView">DataGridViewe is being searched</param>
    ''' <param name="SearchColumns  ">Comma separated column name as string</param>
    '''<param name="SearchValues" >Comma separated Column values  as string..  </param>
    '''<param name="OnlyFirstIndex" >True=Only first row index returned to match the criteria </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataGridViewRowIndex(ByVal LGridView As System.Windows.Forms.DataGridView, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Split(SearchColumns, ",").Count <> Split(SearchValues, ",").Count Then
            QuitMessage("Invalid count of searchcolumns and searchValues " & SearchColumns & " " & SearchValues, "SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()  ")
        End If
        Dim mRowIndex() As Integer = {}
        Dim aColumns() As String = Split(SearchColumns, ",")
        Dim aValues() As String = Split(SearchValues, ",")
        Try
            For i = 0 To LGridView.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                For j = 0 To aColumns.Count - 1
                    If Not UCase(LGridView.Rows(i).Cells(aColumns(j)).Value.ToString.Trim) = UCase(aValues(j).ToString.Trim) Then
                        Matched = False
                        Exit For
                    End If
                Next
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If OnlyFirstIndex = True Then
                    Return mRowIndex
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Public Function SearchDataGridViewRowIndex(ByVal LGridView As System.Windows.Forms.DataGridView, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()")
        End Try
        Return mRowIndex
    End Function

    ''' <summary>
    ''' Searching a DataTable row on  columns with values and return RowIndex()
    ''' </summary>
    ''' <param name="LGridView">DataGridViewe is being searched</param>
    ''' <param name="SearchHashValues  ">A hash table having keys as field names and values as fieldvalues</param>
    '''<param name="OnlyFirstIndex" >True=Only first row index returned to match the criteria </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataGridViewRowIndex(ByVal LGridView As System.Windows.Forms.DataGridView, ByVal SearchHashValues As Hashtable, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mRowIndex() As Integer = {}
        Try
            For i = 0 To LGridView.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                For j = 0 To SearchHashValues.Count - 1
                    Dim mcolumn As String = SearchHashValues.Keys(j)
                    Dim mvalue As Object = GF1.GetValueFromHashTable(SearchHashValues, mcolumn)
                    If Not LGridView.Rows(i).Cells(mcolumn).Value = mvalue Then
                        Matched = False
                        Exit For
                    End If
                Next
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If OnlyFirstIndex = True Then
                    Return mRowIndex
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Public Function Public Function SearchDataGridViewRowIndex(ByVal LGridView As System.Windows.Forms.DataGridView, ByVal SearchHashValues As Hashtable, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()")
        End Try
        Return mRowIndex
    End Function



    ''' <summary>
    ''' Searching a DataTable row on  columns with values and return RowIndex()
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="SearchColumns  ">Comma separated column name as string</param>
    '''<param name="SearchValues" >Comma separated Column values  as string..  </param>
    ''' <returns>A datarow with specified condition</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Split(SearchColumns, ",").Count <> Split(SearchValues, ",").Count Then
            QuitMessage("Invalid count of searchcolumns and searchValues " & SearchColumns & " " & SearchValues, "SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataRow  ")
        End If
        Dim onlyfirstindex As Boolean = True
        Dim mRowIndex() As Integer = {}
        Dim aColumns() As String = Split(SearchColumns, ",")
        Dim aValues() As String = Split(SearchValues, ",")
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                For j = 0 To aColumns.Count - 1
                    If Not UCase(LDataTable.Rows(i).Item(aColumns(j)).ToString.Trim) = UCase(aValues(j)).ToString.Trim Then
                        Matched = False
                        Exit For
                    End If
                Next
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If onlyfirstindex = True Then
                    Return LDataTable(mRowIndex(0))
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataRow")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Searching a DataTable row on  columns with values and return RowIndex()
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="SearchFieldValues  ">HashTable Containg the Key is FieldName and value is selectedkeyfieldvalue.</param>
    ''' <returns>A datarow with specified condition</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal SearchFieldValues As Hashtable) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim onlyfirstindex As Boolean = True
        Dim mRowIndex() As Integer = {}
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                '  For j = 0 To SearchColumns.Count - 1
                For j = 0 To SearchFieldValues.Count - 1
                    Dim mcolumn As String = SearchFieldValues.Keys(j)
                    Dim mvalue As Object = GF1.GetValueFromHashTable(SearchFieldValues, mcolumn)
                    'If mvalue IsNot Nothing Then
                    If Not LDataTable.Rows(i).Item(mcolumn) = mvalue Then
                        Matched = False
                        Exit For
                    End If
                    'End If
                Next
                If Matched = False Then
                    Continue For
                End If
                GF1.ArrayAppend(mRowIndex, i)
                If onlyfirstindex = True Then
                    Return LDataTable(mRowIndex(0))
                    Exit Function
                End If
                'Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Public Function SearchDataTableFirstRow(ByVal LDataTable As DataTable, ByVal SearchFieldValues As Hashtable) As DataRow")
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Find a DataTable row on  primary columns() with values() and return Row
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValues" >An array of values of primary columns eg. {value1,value2,value3...  </param>
    ''' <returns>A DataRow having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues() As Object) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mrow As DataRow = Nothing
        Try
            mrow = LDataTable.Rows.Find(SearchValues)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues() As Object) As DataRow")
        End Try
        Return mrow
    End Function
    ''' <summary>
    ''' Find a DataTable row on  primary columns() with values() and return Row
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchKeyValues" >A Dictionary of having values in integerto be searched,hashset key is column names</param>
    ''' <returns>A DataRow having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchKeyValues As Dictionary(Of String, Integer)) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mrow As DataRow = Nothing
        Try
            Dim SearchValues() As Integer = {}
            For i = 0 To SearchKeyValues.Count - 1
                Dim mkey As String = SearchKeyValues.Keys(i)
                GF1.ArrayAppend(SearchValues, SearchKeyValues.Item(mkey))
            Next
            mrow = LDataTable.Rows.Find(SearchValues)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues() As Object) As DataRow")
        End Try
        Return mrow
    End Function



    ''' <summary>
    ''' Find a DataTable row on   primary column with value and return Row
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValue" >value of primary column...  </param>
    ''' <returns>A DataRow having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValue As Object) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mrow As DataRow = Nothing
        Try
            ' Dim searchvalues() As Object = {SearchValue}
            mrow = LDataTable.Rows.Find(SearchValue)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValue As Object) As DataRow")
        End Try
        Return mrow
    End Function



    ''' <summary>
    ''' Find a DataTable row on   primary column with value and return Row
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValue" >value of primary column...  </param>
    ''' <returns>A DataRow having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValue As Integer) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mrow As DataRow = Nothing
        Try
            'Dim searchvalues() As Integer = {SearchValue}
            mrow = LDataTable.Rows.Find(SearchValue)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValue As Integer) As DataRow")
        End Try
        Return mrow
    End Function
    ''' <summary>
    ''' Find a DataTable row on string type   primary column with value and return Row
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValues" >Comma separated values of primary column...  </param>
    ''' <returns>A DataRow having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues As String) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mrow As DataRow = Nothing
        Try
            Dim asearchvalues() As String = Split(SearchValues, ",")
            mrow = LDataTable.Rows.Find(asearchvalues)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues As String) As DataRow")
        End Try
        Return mrow
    End Function
  

    ''' <summary>
    ''' Find a DataTable row on  primary columns() with values() and return RowIndex
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValues" >An array of values of primary columns eg. {value1,value2,value3...  </param>
    ''' <returns>A DataRow index as integer  having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues() As Object) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mindex As Integer = -1

        Try
            Dim mrow As DataRow = LDataTable.Rows.Find(SearchValues)
            If Not mrow Is Nothing Then
                If Not mrow Is Nothing Then
                    mindex = LDataTable.Rows.IndexOf(mrow)
                End If
                'For i = 0 To LDataTable.Rows.Count - 1
                '    Dim comparer As IEqualityComparer(Of DataRow) = DataRowComparer.Default
                '    Dim CompareFlag As Boolean = comparer.Equals(LDataTable.Rows(i), mrow)
                '    If CompareFlag = True Then
                '        mindex = i
                '        Exit For
                '    End If
                'Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues() As Object) As Integer")
        End Try
        Return mindex
    End Function

    ''' <summary>
    ''' Find a DataTable row on  primary columns() with values() and return RowIndex
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="PrimaryKeyHashTable" >A hashtable having keys as primary columns and values are primary key values..  </param>
    ''' <returns>A DataRow index as integer  having primary keys of values</returns>
    ''' <remarks></remarks>
    Public Function FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal PrimaryKeyHashTable As Hashtable) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mindex As Integer = -1

        Try
            Dim SearchValues() As Object = {}
            For i = 0 To PrimaryKeyHashTable.Count - 1
                Dim mkey As String = PrimaryKeyHashTable.Keys(i)
                GF1.ArrayAppend(searchvalues, PrimaryKeyHashTable.Item(mkey))
            Next
            Dim mrow As DataRow = LDataTable.Rows.Find(SearchValues)
            If Not mrow Is Nothing Then
                mindex = LDataTable.Rows.IndexOf(mrow)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute Public Function FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal PrimaryKeyHashTable As Hashtable) As Integer")
        End Try
        Return mindex
    End Function
    ''' <summary>
    ''' Find a DataTable row on  primary columns with value and return RowIndex
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValue" >value of primary column to be found...  </param>
    ''' <returns>A DataRow index as integer  having primary key of value</returns>
    ''' <remarks></remarks>
    Public Function FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValue As Object) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mindex As Integer = -1
        Try
            ' Dim asearchvalues() As Object = {SearchValue}
            Dim mrow As DataRow = LDataTable.Rows.Find(SearchValue)
            If Not mrow Is Nothing Then
                If Not mrow Is Nothing Then
                    mindex = LDataTable.Rows.IndexOf(mrow)
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValue As Object) As Integer")
        End Try
        Return mindex
    End Function

    ''' <summary>
    ''' Find a DataTable row on  primary columns with value and return RowIndex
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    '''<param name="SearchValues" >Comma separated values of string primary column to be found...  </param>
    ''' <returns>A DataRow index as integer  having primary key of value</returns>
    ''' <remarks></remarks>
    Public Function FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mindex As Integer = -1

        Try
            Dim asearchvalues() As String = Split(SearchValues, ",")
            Dim mrow As DataRow = LDataTable.Rows.Find(asearchvalues)
            If Not mrow Is Nothing Then
                If Not mrow Is Nothing Then
                    mindex = LDataTable.Rows.IndexOf(mrow)
                End If
                'For i = 0 To LDataTable.Rows.Count - 1
                '    Dim comparer As IEqualityComparer(Of DataRow) = DataRowComparer.Default
                '    Dim CompareFlag As Boolean = comparer.Equals(LDataTable.Rows(i), mrow)
                '    If CompareFlag = True Then
                '        mindex = i
                '        Exit For
                '    End If
                'Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FindRowIndexByPrimaryCols(ByVal LDataTable As DataTable, ByVal SearchValues As String) As Integer")
        End Try
        Return mindex
    End Function



    ''' <summary>
    ''' Searching a DataTable row on  columns() with values() and return RowIndex()
    ''' </summary>
    ''' <param name="LDataTable">DataTable is being searched</param>
    ''' <param name="ColumnValuePair" >A hashtable oject of columns where key is columnname and value is column value</param>
    '''<param name="OnlyFirstIndex" >True=Only first row index returned to match the criteria </param>
    ''' <returns>An integer array of rowindex</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal ColumnValuePair As Hashtable, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mRowIndex() As Integer = {}
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                Dim Matched As Boolean = True
                For j = 0 To ColumnValuePair.Count - 1
                    Dim mkey As String = LCase(ColumnValuePair.Keys(j))
                    Dim aValues() As String = Split(GF1.GetValueFromHashTable(ColumnValuePair, mkey).ToString, ",")
                    If aValues.Count > 1 Then
                        Dim matched1 As Boolean = False
                        For k = 0 To aValues.Count - 1
                            If LCase(LDataTable.Rows(i).Item(mkey).ToString) = LCase(aValues(k)) Then
                                matched1 = True
                                Exit For
                            End If
                        Next
                        Matched = matched1
                    Else
                        If Not LCase(LDataTable.Rows(i).Item(mkey).ToString) = LCase(GF1.GetValueFromHashTable(ColumnValuePair, mkey).ToString) Then
                            Matched = False
                        End If
                    End If
                    If Matched = False Then
                        Exit For
                    End If
                Next
                If Matched = True Then
                    GF1.ArrayAppend(mRowIndex, i)
                    If OnlyFirstIndex = True Then
                        Return mRowIndex
                        Exit Function
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTableRowIndex(ByVal LDataTable As DataTable, ByVal ColumnValuePair As Hashtable, Optional ByVal OnlyFirstIndex As Boolean = False) As Integer()")
        End Try
        Return mRowIndex
    End Function
    ''' <summary>
    '''Adding Name column and its value to a datatable corresponding to its code value searching from a code master data table
    ''' </summary>
    ''' <param name="LDataTable">DataTable object on which name columns added</param>
    ''' <param name="TableCodeFields">Comma separated string of column names of table,their values to be searched in code master</param>
    ''' <param name="AddingName" >ColumnName to be inserted in ldatatable for name value of masternamefield</param>
    ''' <param name="CodeMaster">A datatable having columns corressponding to TableKeyColumn and TableNameColumn</param>
    ''' <param name="MasterCodeFields">Comma separated string of column names of CodeMastertable to be linked</param>
    ''' <param name="MasterNameField">ColumnName of CodeMaster table which contains the name</param>
    ''' <param name="AddCodeInName" >"P" if table column code is added as prefix,"S" if table column code is added as suffix,"N" for do not add</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddingNameForCodes(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingName As String, ByVal CodeMaster As DataTable, ByVal MasterCodeFields As String, ByVal MasterNameField As String, Optional ByVal AddCodeInName As String = "N") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim aTablecodeFields() As String = Split(TableCodeFields, ",")
            Dim TableNameField As String = AddingName
            If CheckColumnInDataTable(TableNameField, LDataTable) = -1 Then
                LDataTable.Columns.Add(TableNameField)
            End If
            Dim aMasterCodeFields() As String = Split(MasterCodeFields, ",")
            For i = 0 To LDataTable.Rows.Count - 1
                Dim aTableFieldsValue() As Object = {}
                For k = 0 To aTablecodeFields.Count - 1
                    GF1.ArrayAppend(aTableFieldsValue, LDataTable.Rows(i).Item(aTablecodeFields(k)))
                Next
                Dim row1() As Integer = SearchDataTableRowIndex(CodeMaster, aMasterCodeFields, aTableFieldsValue, True)
                If row1.Count > 0 Then
                    Dim MasterNameValue As String = CodeMaster.Rows(row1(0)).Item(MasterNameField).ToString.Trim
                    Select Case UCase(AddCodeInName)
                        Case "P"
                            MasterNameValue = " (" & Join(aTableFieldsValue, ".") & ")" & MasterNameValue
                        Case "S"
                            MasterNameValue = MasterNameValue & " (" & Join(aTableFieldsValue, ".") & ")"
                    End Select
                    LDataTable.Rows(i).Item(TableNameField) = MasterNameValue
                Else
                    QuitMessage("Values " & Join(aTableFieldsValue, ",") & " not found in " & CodeMaster.TableName, "AddingNameForCodes(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingName As String, ByVal CodeMaster As DataTable, ByVal MasterCodeFields As String, ByVal MasterNameField As String, Optional ByVal AddCodeInName As String = N) As DataTable  ")
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddingNameForCodes(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingName As String, ByVal CodeMaster As DataTable, ByVal MasterCodeFields As String, ByVal MasterNameField As String, Optional ByVal AddCodeInName As String = "") As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    '''Adding Name columns and its value to a datatable corresponding to its code value searching from a code master data table
    ''' </summary>
    ''' <param name="LDataTable">DataTable object on which name columns added</param>
    ''' <param name="TableCodeFields">Comma separated string of column names of table,their values to be searched in code master</param>
    ''' <param name="AddingNames" >Corresponding comma separated ColumnNames to be inserted in ldatatable for name value of masternamefield</param>
    ''' <param name="CodeMaster">A datatable having a primary key with values corressponding to columns of TableCodeFields</param>
    ''' <param name="MasterCodeFields">Comma separated string of column names of CodeMasterTable to be linked</param>
    ''' <param name="MasterNameField">ColumnName of CodeMaster table which contains the name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddingNameColumnsForCodes(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingNames As String, ByVal CodeMaster As DataTable, ByVal MasterCodeFields As String, ByVal MasterNameField As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim aTablecodeFields() As String = Split(TableCodeFields, ",")
            Dim TableNameField() As String = AddingNames.Split(",")
            If aTablecodeFields.Count <> TableNameField.Count Then
                GF1.QuitMessage("Columns in TableCodeFields  must be same in columns in  AddingNames", "AddingNameColumnsForCodes")
            End If
            For i = 0 To TableNameField.Count - 1
                If CheckColumnInDataTable(TableNameField(i), LDataTable) = -1 Then
                    LDataTable.Columns.Add(TableNameField(i))
                End If
            Next
            ' Dim MasterCodePrimary As String = GetPrimaryKey(CodeMaster)
            For i = 0 To LDataTable.Rows.Count - 1
                For k = 0 To aTablecodeFields.Count - 1
                    Dim row1 As DataRow = FindRowByPrimaryCols(CodeMaster, LDataTable.Rows(i).Item(aTablecodeFields(k)))
                    If row1 IsNot Nothing Then
                        Dim MasterNameValue As String = row1.Item(MasterNameField).ToString.Trim
                        LDataTable.Rows(i).Item(TableNameField(k)) = MasterNameValue
                    End If
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddingNameColumnsForCodes(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingNames As String, ByVal CodeMaster As DataTable, ByVal MasterCodeFields As String, ByVal MasterNameField As String) As DataTable")
        End Try
        Return LDataTable
    End Function



    ''' <summary>
    ''' Replace Column Values into a DataTable of selective rows by a condition 
    ''' </summary>
    ''' <param name="LDataTable">DataTable whoose columns replaced</param>
    ''' <param name="ColumnNamesValuePair">A hash table containing column names as key and value as column value</param>
    ''' <param name="Lcondition">Condition string eg.  "Column1 = 'Value1' etc.</param>
    ''' <returns>Final data table </returns>
    ''' <remarks></remarks>
    Public Function ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNamesValuePair As Hashtable, Optional ByVal Lcondition As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If Lcondition.Trim.Length > 0 Then
                Dim mDataTable As DataTable = SortFilterDataTable(LDataTable, "", "", Lcondition)
                For i = 0 To mDataTable.Rows.Count - 1
                    Dim mindex As Integer = CheckRowInDataTable(mDataTable.Rows(i), LDataTable)
                    If mindex > 0 Then
                        For j = 0 To ColumnNamesValuePair.Count - 1
                            Dim mkey As String = ColumnNamesValuePair.Keys(j)
                            Dim mvalue As Object = GF1.GetValueFromHashTable(ColumnNamesValuePair, mkey)
                            LDataTable.Rows(mindex).Item(mkey) = mvalue
                        Next
                    End If
                Next
            Else
                For i = 0 To LDataTable.Rows.Count - 1
                    For j = 0 To ColumnNamesValuePair.Count - 1
                        Dim mkey As String = ColumnNamesValuePair.Keys(j)
                        Dim mvalue As Object = GF1.GetValueFromHashTable(ColumnNamesValuePair, mkey)
                        LDataTable.Rows(i).Item(mkey) = mvalue
                    Next
                Next
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNamesValuePair As Hashtable, Optional ByVal Lcondition As String = "") As DataTable")
        End Try
        Return LDataTable
    End Function

    ''' <summary>
    ''' Replace Column Values into a DataTable of selective rows by a condition 
    ''' </summary>
    ''' <param name="LDataTable">DataTable whoose columns replaced</param>
    ''' <param name="ColumnNamesValuePair">A hash table containing column names as key and value as column value</param>
    ''' <param name="Hcondition">Condition HashTable where key is column names and values are column values,nothing for none.</param>
    ''' <returns>Final data table </returns>
    ''' <remarks></remarks>
    Public Function ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNamesValuePair As Hashtable, ByVal Hcondition As Hashtable) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim LCondition As String = GF1.GetStringConditionFromHashTable(Hcondition, True)
            If LCondition.Trim.Length > 0 Then
                Dim mDataTable As DataTable = SortFilterDataTable(LDataTable, "", "", LCondition)
                For i = 0 To mDataTable.Rows.Count - 1
                    Dim mindex As Integer = CheckRowInDataTable(mDataTable.Rows(i), LDataTable)
                    If mindex > 0 Then
                        For j = 0 To ColumnNamesValuePair.Count - 1
                            Dim mkey As String = ColumnNamesValuePair.Keys(j)
                            Dim mvalue As Object = GF1.GetValueFromHashTable(ColumnNamesValuePair, mkey)
                            LDataTable.Rows(mindex).Item(mkey) = mvalue
                        Next
                    End If
                Next
            Else
                For i = 0 To LDataTable.Rows.Count - 1
                    For j = 0 To ColumnNamesValuePair.Count - 1
                        Dim mkey As String = ColumnNamesValuePair.Keys(j)
                        Dim mvalue As Object = GF1.GetValueFromHashTable(ColumnNamesValuePair, mkey)
                        LDataTable.Rows(i).Item(mkey) = mvalue
                    Next
                Next
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNamesValuePair As Hashtable, Optional ByVal Lcondition As String = "") As DataTable")
        End Try
        Return LDataTable
    End Function


    ''' <summary>
    ''' Replace Column Values into a DataTable of selective rows by a condition 
    ''' </summary>
    ''' <param name="LDataTable">DataTable whoose columns replaced</param>
    ''' <param name="ColumnNames">A comma separated string of  column names</param>
    ''' <param name="ColumnValues" >A comma  separated string of  column values</param>
    ''' <param name="Lcondition">Condition string eg.  "Column1 = 'Value1' etc.</param>
    ''' <returns>Final data table </returns>
    ''' <remarks></remarks>
    Public Function ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, ByVal ColumnValues As String, Optional ByVal Lcondition As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If Lcondition.Trim.Length > 0 Then
                Dim mDataTable As DataTable = SortFilterDataTable(LDataTable, "", "", Lcondition)
                For i = 0 To mDataTable.Rows.Count - 1
                    Dim mindex As Integer = CheckRowInDataTable(mDataTable.Rows(i), LDataTable)
                    If mindex > 0 Then
                        Dim aColumns() As String = Split(ColumnNames, ",")
                        Dim aValues() As String = Split(ColumnValues, ",")
                        For j = 0 To aColumns.Count - 1
                            LDataTable.Rows(mindex).Item(aColumns(j)) = aValues(j)
                        Next
                    End If
                Next
            Else
                For i = 0 To LDataTable.Rows.Count - 1
                    Dim aColumns() As String = Split(ColumnNames, ",")
                    Dim aValues() As String = Split(ColumnValues, ",")
                    For j = 0 To aColumns.Count - 1
                        LDataTable.Rows(i).Item(aColumns(j)) = aValues(j)
                    Next
                Next
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, ByVal ColumnValues As String, Optional ByVal Lcondition As String = "") As DataTable")
        End Try
        Return LDataTable
    End Function


    ''' <summary>
    ''' Replace Column Values into a DataTable of selective rows by a condition 
    ''' </summary>
    ''' <param name="LDataTable">DataTable whoose columns replaced</param>
    ''' <param name="ColumnNames">A comma separated string of  column names</param>
    ''' <param name="ColumnValues" >A comma  separated string of  column values</param>
    ''' <param name="Hcondition">Condition hash table where key is column name and value is column value.</param>
    ''' <param name="LastRowNo" >Last row no which was replaced</param>
    ''' <returns>Final data table </returns>
    ''' <remarks></remarks>
    Public Function ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, ByVal ColumnValues As String, ByVal Hcondition As Hashtable, Optional ByRef LastRowNo As Integer = -1) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim LCondition As String = GF1.GetStringConditionFromHashTable(Hcondition, True)
            If LCondition.Trim.Length > 0 Then
                Dim mDataTable As DataTable = SortFilterDataTable(LDataTable, "", "", LCondition)
                For i = 0 To mDataTable.Rows.Count - 1
                    Dim mindex As Integer = CheckRowInDataTable(mDataTable.Rows(i), LDataTable)
                    If mindex > 0 Then
                        Dim aColumns() As String = Split(ColumnNames, ",")
                        Dim aValues() As String = Split(ColumnValues, ",")
                        For j = 0 To aColumns.Count - 1
                            LDataTable.Rows(mindex).Item(aColumns(j)) = aValues(j)
                        Next
                        LastRowNo = mindex
                    End If
                Next
            Else
                For i = 0 To LDataTable.Rows.Count - 1
                    Dim aColumns() As String = Split(ColumnNames, ",")
                    Dim aValues() As String = Split(ColumnValues, ",")
                    For j = 0 To aColumns.Count - 1
                        LDataTable.Rows(i).Item(aColumns(j)) = aValues(j)
                    Next
                    LastRowNo = i
                Next
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ReplaceValuesInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, ByVal ColumnValues As String, Optional ByVal Lcondition As String = "") As DataTable")
        End Try
        Return LDataTable
    End Function


    ''' <summary>
    ''' Replace Column Values into a DataTable of selective row by primary key value
    ''' </summary>
    ''' <param name="LDataTable">DataTable whoose columns replaced</param>
    ''' <param name="ColumnNames">A comma separated string of  column names</param>
    ''' <param name="ColumnValues" >A comma  separated string of  column values</param>
    ''' <param name="PrimaryKeyValue" >Primary key value </param>
    ''' <returns>Final data table </returns>
    ''' <remarks></remarks>
    Public Function ReplaceValuesInDTbyPrimaryKey(ByRef LDataTable As DataTable, ByVal ColumnNames As String, ByVal ColumnValues As String, ByVal PrimaryKeyValue As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mflag As Boolean = False
        Try
            Dim mindex As Integer = FindRowIndexByPrimaryCols(LDataTable, PrimaryKeyValue)
            If mindex > 0 Then
                Dim aColumns() As String = Split(ColumnNames, ",")
                Dim aValues() As String = Split(ColumnValues, ",")
                For j = 0 To aColumns.Count - 1
                    LDataTable.Rows(mindex).Item(aColumns(j)) = aValues(j)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.ReplaceValuesInDTbyPrimaryKey(ByRef LDataTable As DataTable, ByVal ColumnNames As String, ByVal ColumnValues As String, ByVal PrimaryKeyValue As String) As Boolean")
        End Try
        Return mflag
    End Function

    ''' <summary>
    '''Adding Names column and its value to a datatable corresponding to its code value searching from a code master data table
    ''' </summary>
    ''' <param name="LDataTable">DataTable object on which name columns added</param>
    ''' <param name="TableCodeFields">Comma separated  string whoose items are column names of table,their values to be searched in code master</param>
    ''' <param name="AddingName" >Column separated ColumnName to be inserted in ldatatable for name value of masternamefield</param>
    ''' <param name="CodeMasterWithPrimaryColumns">A datatable having Primary columns corressponding to TableCodeFields and TableNameColumn</param>
    ''' <param name="MasterNameField">ColumnName of CodeMaster table which contains the name</param>
    ''' <param name="AddCodeInName" >"P" if table column code is added as prefix,"S" if table column code is added as suffix,"N" for do not add</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddingNameForCodesPrimamryCols(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingName As String, ByVal CodeMasterWithPrimaryColumns As DataTable, ByVal MasterNameField As String, Optional ByVal AddCodeInName As String = "N") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            ' Dim TableNameField As String = AddingName
            Dim aTableNameField() As String = Split(AddingName, ",")
            Dim aTableCodeFields() As String = Split(TableCodeFields, ",")
            If aTableCodeFields.Count <> aTableNameField.Count Then
                QuitMessage("No. of name fields differs no. of code fields", "AddingNameForCodesPrimamryCols")
            End If

            For i = 0 To aTableNameField.Count - 1
                If CheckColumnInDataTable(aTableNameField(i), LDataTable) < 0 Then
                    LDataTable.Columns.Add(aTableNameField(i))
                End If
            Next
            For i = 0 To LDataTable.Rows.Count - 1
                For k = 0 To aTableCodeFields.Count - 1
                    Dim TableFieldsValue As Object = LDataTable.Rows(i).Item(aTableCodeFields(k))
                    Dim row1 As DataRow = FindRowByPrimaryCols(CodeMasterWithPrimaryColumns, TableFieldsValue)
                    If Not row1 Is Nothing Then
                        Dim MasterNameValue As String = row1.Item(MasterNameField)
                        Select Case UCase(AddCodeInName)
                            Case "P"
                                MasterNameValue = " (" & Join(TableFieldsValue, ".") & ")" & MasterNameValue
                            Case "S"
                                MasterNameValue = MasterNameValue & " (" & Join(TableFieldsValue, ".") & ")"
                        End Select
                        LDataTable.Rows(i).Item(aTableNameField(k)) = MasterNameValue
                    Else
                        If IsDBNull(TableFieldsValue) = False Then
                            If TableFieldsValue > -1 Then
                                LDataTable.Rows(i).Item(aTableNameField(k)) = "missing"
                                '   GF1.QuitMessage("Values " & Join(TableFieldsValue, ",") & " not found in " & CodeMasterWithPrimaryColumns.TableName)
                            End If
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddingNameForCodesPrimamryCols(ByVal LDataTable As DataTable, ByVal TableCodeFields As String, ByVal AddingName As String, ByVal CodeMasterWithPrimaryColumns As DataTable, ByVal MasterNameField As String, Optional ByVal AddCodeInName As String = N) As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    '''Fill Values for blank fields in a datatable 
    ''' </summary>
    ''' <param name="LDataTable">Data table to be filled</param>
    ''' <param name="TableCodeFields">Array of field's names to be checked for blank</param>
    ''' <param name="TableValues">Array of field values to be filled</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FillValuesForBlanks(ByVal LDataTable As DataTable, ByVal TableCodeFields() As String, ByVal TableValues() As Object) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To LDataTable.Rows.Count - 1
                For j = 0 To TableCodeFields.Count - 1
                    Dim mType As String = LCase(LDataTable.Rows(i).Item(TableCodeFields(j)).GetType.Name.ToString)
                    Dim mValue As Object = LDataTable.Rows(i).Item(TableCodeFields(j)).ToString.Trim
                    If IsDBNull(LDataTable.Rows(i).Item(TableCodeFields(j))) = True Then
                        Continue For
                    End If
                    Select Case LCase(mType)
                        Case "string", "datetime", "date"
                            If LDataTable.Rows(i).Item(TableCodeFields(j)) Is Nothing Then
                                LDataTable.Rows(i).Item(TableCodeFields(j)) = TableValues(j)
                            End If
                            If LDataTable.Rows(i).Item(TableCodeFields(j)).ToString.Trim.Length = 0 Then
                                LDataTable.Rows(i).Item(TableCodeFields(j)) = TableValues(j)
                            End If
                        Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                            If LDataTable.Rows(i).Item(TableCodeFields(j)) = 0 Then
                                LDataTable.Rows(i).Item(TableCodeFields(j)) = TableValues(j)
                            End If
                    End Select
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.FillValuesForBlanks(ByVal LDataTable As DataTable, ByVal TableCodeFields() As String, ByVal TableValues() As Object) As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    ''' This function searches a DataTable on specified columns with values and return datatable satisfying the criteria.
    ''' </summary>
    ''' <param name="LDataTable">DataTable searched on specified columns</param>
    ''' <param name="SearchColumns">An array of  column names eg. {"Column1","Column2", ..  } etc.</param>
    '''<param name="SearchValues" >An array of values eg. {value1,value2,value3...  </param>
    ''' <returns>Datatable of specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns() As String, ByVal SearchValues() As Object) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Not SearchColumns.Count = SearchValues.Count Then
            QuitMessage("Size of SearchColumns differs SearchValues" & " SearchDataTable", "SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns() As String, ByVal SearchValues() As Object) As DataTable  ")
        End If
        Dim mDataTable As New DataTable
        Dim SearchExpr As String = ""
        For i = 0 To SearchColumns.Count - 1
            Dim mtype As String = LCase(LDataTable.Columns(SearchColumns(i)).DataType.Name.ToString)
            Select Case LCase(mtype)
                Case "string", "datetime", "date"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " and ") & SearchColumns(i) & " = '" & SearchValues(i).ToString & "'"
                Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " and ") & SearchColumns(i) & " = " & SearchValues(i).ToString
                Case Else
                    Continue For
            End Select
        Next
        Try
            Dim mRows() As DataRow = LDataTable.Select(SearchExpr)
            If mRows.Count > 0 Then
                mDataTable = mRows.CopyToDataTable
            Else
                mDataTable.Rows.Clear()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns() As String, ByVal SearchValues() As Object) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' This function searches a DataTable on specified columns with values and return datatable satisfying the criteria.
    ''' </summary>
    ''' <param name="LDataTable">DataTable searched on specified columns</param>
    ''' <param name="SearchColumns">Comma separated column names to be searched eg column1,column2,column3} etc.</param>
    '''<param name="SearchValues" >Comma separated string values to be searched eg. {value1,value2,value3...  </param>
    ''' <returns>Datatable of specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Not Split(SearchColumns, ",").Count = Split(SearchValues, ",").Count Then
            QuitMessage("Size of SearchColumns differs SearchValues" & " SearchDataTable", "SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable  ")
        End If
        Dim mDataTable As New DataTable
        Dim SearchExpr As String = ""
        Dim aSearchColumns() As String = Split(SearchColumns, ",")
        Dim aSearchvalues() As String = Split(SearchValues, ",")
        For i = 0 To aSearchColumns.Count - 1
            Dim mtype As String = LCase(LDataTable.Columns(aSearchColumns(i)).DataType.Name.ToString)
            Select Case LCase(mtype)
                Case "string", "datetime", "date"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " and ") & aSearchColumns(i) & " = '" & aSearchvalues(i).ToString & "'"
                Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " and ") & aSearchColumns(i) & " = " & aSearchvalues(i).ToString
                Case Else
                    Continue For
            End Select
        Next
        Try
            Dim mRows() As DataRow = LDataTable.Select(SearchExpr)
            If mRows.Count > 0 Then
                mDataTable = mRows.CopyToDataTable
            Else
                mDataTable.Rows.Clear()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' This function searches a DataTable on specified column with its integer value and return datatable satisfying the criteria.
    ''' </summary>
    ''' <param name="LDataTable">DataTable searched on specified columns</param>
    ''' <param name="SearchColumn">column name to be searched</param>
    '''<param name="SearchValue" >integer value to be searched  </param>
    ''' <returns>Datatable of specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As New DataTable
        Dim SearchExpr As String = SearchColumn & " = " & SearchValue.ToString
        Try
            Dim mRows() As DataRow = LDataTable.Select(SearchExpr)
            If mRows.Count > 0 Then
                mDataTable = mRows.CopyToDataTable
                mDataTable.Rows.Clear()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' This function searches a string value into a column of DataTable and return datatable .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched.</param>
    ''' <param name="SearchColumn">Column name to be searched.</param>
    '''<param name="SearchString" >String value to be Sought.  </param>
    ''' <param name="IgnoreCase" >if Searching is not case sensitive.</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchStringInDTColumn(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchString As String, Optional ByVal IgnoreCase As Boolean = True) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        If LDataTable.Rows.Count = 0 Then
            Return mDataTable
            Exit Function
        End If
        Try
            SearchString = IIf(IgnoreCase = True, LCase(SearchString), SearchString)
            For i = 0 To LDataTable.Rows.Count - 1
                Dim mcolvalue As String = LDataTable.Rows(i).Item(SearchColumn).ToString
                mcolvalue = IIf(IgnoreCase = True, LCase(mcolvalue), mcolvalue)
                If InStr(mcolvalue, SearchString) > 0 Then
                    mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(i), True)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchStringInDTColumn(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchString As String, Optional ByVal IgnoreCase As Boolean = True) As DataTable")
        End Try
        Return mDataTable
    End Function
    '''' <summary>
    '''' This function searches  value of  specified column of Sorted DataTable on specified column .
    '''' </summary>
    '''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumns.</param>
    '''' <param name="SearchColumns">Comma separated column name(s) to be searched(2 columns only).</param>
    ''''<param name="SearchValues" >Comma separated column value(s) to be Sought.  </param>
    '''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    '''' <returns>Datatable on specified criteria</returns>
    '''' <remarks></remarks>
    'Public Function SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OutputOnSorting As String = "") As DataTable
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    Dim mDataTable As DataTable = LDataTable.Clone
    '    If LDataTable.Rows.Count = 0 Then
    '        Return mDataTable
    '        Exit Function
    '    End If
    '    Try
    '        Dim aSearchColumn() As String = SearchColumns.Split(",")
    '        Dim aSearchValue() As String = SearchValues.Split(",")
    '        If aSearchColumn.Count > 1 Then
    '            If aSearchColumn.Count <> aSearchValue.Count Then
    '                QuitMessage("Comma separated Columns and values count must be same", "SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OutputOnSorting As String = "") As DataTable")
    '            End If
    '            If aSearchColumn.Count > 2 Then
    '                QuitMessage("Only two Columns permissible", "SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OutputOnSorting As String = "") As DataTable")
    '            End If
    '        End If

    '        If aSearchColumn.Count = 1 Then
    '            Dim SearchColumn As String = aSearchColumn(0)
    '            Dim SearchValue As String = aSearchValue(0)
    '            Dim first As Integer = 0
    '            Dim last As Integer = LDataTable.Rows.Count - 1
    '            If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
    '                Return mDataTable
    '                Exit Function
    '            End If
    '            Dim middle As Integer = last \ 2
    '            Dim fnd As Boolean = False
    '            While first <= last
    '                Select Case True
    '                    Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn)
    '                        first = middle + 1
    '                        middle = (first + last) \ 2
    '                    Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn)
    '                        last = middle - 1
    '                        middle = (first + last) \ 2
    '                    Case LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
    '                        fnd = True
    '                        Dim k As Integer = middle
    '                        While LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
    '                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
    '                            middle = middle - 1
    '                            If middle < 0 Then
    '                                Exit While
    '                            End If
    '                        End While
    '                        While LDataTable.Rows(k).Item(SearchColumn) = SearchValue
    '                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
    '                            k = k + 1
    '                            If k > last Then
    '                                Exit While
    '                            End If
    '                        End While
    '                End Select
    '                If fnd = True Then
    '                    Exit While
    '                End If
    '            End While
    '            If OutputOnSorting.Length > 0 Then
    '                mDataTable = SortDataTable(mDataTable, OutputOnSorting)
    '            End If
    '        Else
    '            Dim first As Integer = 0
    '            Dim last As Integer = LDataTable.Rows.Count - 1
    '            Dim col1 As String = aSearchColumn(0)
    '            Dim col2 As String = aSearchColumn(1)
    '            Dim val1 As String = aSearchValue(0)
    '            Dim val2 As String = aSearchValue(1)
    '            Dim LastValue1 As String = LDataTable(last).Item(col1)
    '            Dim FirstValue1 As String = LDataTable(first).Item(col1)
    '            Dim LastValue2 As String = LDataTable(last).Item(col2)
    '            Dim FirstValue2 As String = LDataTable(first).Item(col2)

    '            If (val1 > LastValue1 Or val1 < FirstValue1) Or (val2 > LastValue2 Or val2 < FirstValue2) Then
    '                Return mDataTable
    '                Exit Function
    '            End If
    '            Dim middle As Integer = last \ 2
    '            Dim fnd As Boolean = False
    '            While first <= last
    '                Select Case True
    '                    Case LDataTable(middle).Item(col1) < val1
    '                        first = middle + 1
    '                        middle = (first + last) \ 2
    '                    Case LDataTable(middle).Item(col1) > val1
    '                        last = middle - 1
    '                        middle = (first + last) \ 2
    '                    Case val2 > LDataTable(middle).Item(col2)
    '                        first = middle + 1
    '                        middle = (first + last) \ 2
    '                    Case val2 < LDataTable(middle).Item(col2)
    '                        last = middle - 1
    '                        middle = (first + last) \ 2
    '                    Case LDataTable(middle).Item(col2) = val2
    '                        fnd = True
    '                        Dim k As Integer = middle
    '                        While LDataTable(middle).Item(col2) = val2 And LDataTable(middle).Item(col1) = val1
    '                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
    '                            middle = middle - 1
    '                            If middle < 0 Then
    '                                Exit While
    '                            End If
    '                        End While
    '                        While LDataTable(middle).Item(col2) = val2 And LDataTable(middle).Item(col1) = val1
    '                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
    '                            k = k + 1
    '                            If k > last Then
    '                                Exit While
    '                            End If
    '                        End While
    '                End Select
    '                If fnd = True Then
    '                    Exit While
    '                End If
    '            End While
    '            If OutputOnSorting.Length > 0 Then
    '                mDataTable = SortDataTable(mDataTable, OutputOnSorting)
    '            End If
    '        End If
    '    Catch ex As Exception
    '        QuitError(ex, Err, "Unable to execute DataFunction.SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String) As DataTable")
    '    End Try
    '    Return mDataTable
    'End Function


    ''' <summary>
    ''' This function searches  value of  specified column of Sorted DataTable on specified column .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumns.</param>
    ''' <param name="SearchColumns">Comma separated column name(s) to be searched(2 columns only).</param>
    '''<param name="SearchValues" >Comma separated column value(s) to be Sought.  </param>
    ''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OutputOnSorting As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        If LDataTable.Rows.Count = 0 Then
            Return mDataTable
            Exit Function
        End If
        Try
            Dim aSearchColumn() As String = SearchColumns.Split(",")
            Dim aSearchValue() As String = SearchValues.Split(",")
            If aSearchColumn.Count > 1 Then
                If aSearchColumn.Count <> aSearchValue.Count Then
                    QuitMessage("Comma separated Columns and values count must be same", "SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OutputOnSorting As String = "") As DataTable")
                End If
                If aSearchColumn.Count > 2 Then
                    QuitMessage("Only two Columns permissible", "SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String, Optional ByVal OutputOnSorting As String = "") As DataTable")
                End If
            End If

            If aSearchColumn.Count = 1 Then
                Dim SearchColumn As String = aSearchColumn(0)
                Dim SearchValue As String = aSearchValue(0)
                Dim first As Integer = 0
                Dim last As Integer = LDataTable.Rows.Count - 1
                If SearchValue > LDataTable.Rows(last).Item(SearchColumn).ToString Or SearchValue < LDataTable.Rows(first).Item(SearchColumn).ToString Then
                    Return mDataTable
                    Exit Function
                End If
                Dim middle As Integer = last \ 2
                Dim fnd As Boolean = False
                While first <= last
                    Select Case True
                        Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn).ToString
                            first = middle + 1
                            middle = (first + last) \ 2
                        Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn).ToString
                            last = middle - 1
                            middle = (first + last) \ 2
                        Case LDataTable.Rows(middle).Item(SearchColumn).ToString = SearchValue
                            fnd = True
                            Dim k As Integer = middle
                            While LDataTable.Rows(middle).Item(SearchColumn).ToString = SearchValue
                                mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                                middle = middle - 1
                                If middle < 0 Then
                                    Exit While
                                End If
                            End While
                            While LDataTable.Rows(k).Item(SearchColumn).ToString = SearchValue
                                mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
                                k = k + 1
                                If k > last Then
                                    Exit While
                                End If
                            End While
                    End Select
                    If fnd = True Then
                        Exit While
                    End If
                End While
                If OutputOnSorting.Length > 0 Then
                    mDataTable = SortDataTable(mDataTable, OutputOnSorting)
                End If
            Else
                Dim first As Integer = 0
                Dim last As Integer = LDataTable.Rows.Count - 1
                Dim col1 As String = aSearchColumn(0)
                Dim col2 As String = aSearchColumn(1)
                Dim val1 As String = aSearchValue(0)
                Dim val2 As String = aSearchValue(1)
                Dim LastValue1 As String = LDataTable(last).Item(col1).ToString
                Dim FirstValue1 As String = LDataTable(first).Item(col1).ToString
                Dim LastValue2 As String = LDataTable(last).Item(col2).ToString
                Dim FirstValue2 As String = LDataTable(first).Item(col2).ToString

                If (val1 > LastValue1 Or val1 < FirstValue1) Or (val2 > LastValue2 Or val2 < FirstValue2) Then
                    Return mDataTable
                    Exit Function
                End If
                Dim middle As Integer = last \ 2
                Dim fnd As Boolean = False
                While first <= last
                    Select Case True
                        Case LDataTable(middle).Item(col1) < val1
                            first = middle + 1
                            middle = (first + last) \ 2
                        Case LDataTable(middle).Item(col1) > val1
                            last = middle - 1
                            middle = (first + last) \ 2
                        Case val2 > LDataTable(middle).Item(col2).ToString
                            first = middle + 1
                            middle = (first + last) \ 2
                        Case val2 < LDataTable(middle).Item(col2).ToString
                            last = middle - 1
                            middle = (first + last) \ 2
                        Case LDataTable(middle).Item(col2).ToString = val2
                            fnd = True
                            Dim k As Integer = middle
                            While LDataTable(middle).Item(col2).ToString = val2 And LDataTable(middle).Item(col1).ToString = val1
                                mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                                middle = middle - 1
                                If middle < 0 Then
                                    Exit While
                                End If
                            End While
                            If middle > -1 And last > -1 Then
                                While LDataTable(middle).Item(col2).ToString = val2 And LDataTable(middle).Item(col1).ToString = val1
                                    mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
                                    k = k + 1
                                    If k > last Then
                                        Exit While
                                    End If
                                End While
                            End If
                    End Select
                    If fnd = True Then
                        Exit While
                    End If
                End While
                If OutputOnSorting.Length > 0 Then
                    mDataTable = SortDataTable(mDataTable, OutputOnSorting)
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String) As DataTable")
        End Try
        Return mDataTable
    End Function





    ''' <summary>
    ''' This function searches  value of  specified column of Sorted DataTable in a ~ separated string expression .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    ''' <param name="SearchColumn">Column name to be sought in StringExpression.</param>
    '''<param name="StringExpression" >~  separated string expression of  column value(s) to be searched  </param>
    ''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchColumnInStringExpression(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal StringExpression As String, Optional ByVal OutputOnSorting As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        If LDataTable.Rows.Count = 0 Then
            Return mDataTable
            Exit Function
        End If
        Try
            Dim aSearchValue() As String = StringExpression.Split("~")
            For i = 0 To aSearchValue.Count - 1
                Dim SearchValue As String = aSearchValue(i)
                Dim first As Integer = 0
                Dim last As Integer = LDataTable.Rows.Count - 1
                If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
                    Return mDataTable
                    Exit Function
                End If
                Dim middle As Integer = last \ 2
                Dim fnd As Boolean = False
                While first <= last
                    Select Case True
                        Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn)
                            first = middle + 1
                            middle = (first + last) \ 2
                        Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn)
                            last = middle - 1
                            middle = (first + last) \ 2
                        Case LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
                            fnd = True
                            Dim k As Integer = middle
                            While LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
                                mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                                middle = middle - 1
                                If middle < 0 Then
                                    Exit While
                                End If
                            End While
                            While LDataTable.Rows(k).Item(SearchColumn) = SearchValue
                                mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
                                k = k + 1
                                If k > last Then
                                    Exit While
                                End If
                            End While
                    End Select
                    If fnd = True Then
                        Exit While
                    End If
                End While
            Next
            If OutputOnSorting.Length > 0 Then
                mDataTable = SortDataTable(mDataTable, OutputOnSorting)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String) As DataTable")
        End Try
        Return mDataTable
    End Function



    '''' <summary>
    '''' This function searches  value of  specified column of Sorted DataTable on specified column .
    '''' </summary>
    '''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    '''' <param name="SearchColumn">Column name to be searched.</param>
    ''''<param name="SearchValue" >Column value to be Sought.  </param>
    '''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    '''' <returns>Datatable on specified criteria</returns>
    '''' <remarks></remarks>
    'Public Function SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer, Optional ByVal OutputOnSorting As String = "") As DataTable
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    Dim mDataTable As DataTable = LDataTable.Clone
    '    If LDataTable.Rows.Count = 0 Then
    '        Return mDataTable
    '        Exit Function
    '    End If
    '    Try
    '        Dim first As Integer = 0
    '        Dim last As Integer = LDataTable.Rows.Count - 1
    '        If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
    '            Return mDataTable
    '            Exit Function
    '        End If
    '        Dim middle As Integer = last \ 2
    '        Dim fnd As Boolean = False
    '        While first <= last
    '            Select Case True
    '                Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn)
    '                    first = middle + 1
    '                    middle = (first + last) \ 2
    '                Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn)
    '                    last = middle - 1
    '                    middle = (first + last) \ 2
    '                Case LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
    '                    fnd = True
    '                    Dim k As Integer = middle
    '                    While LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
    '                        mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
    '                        middle = middle - 1
    '                        If middle < 0 Then
    '                            Exit While
    '                        End If
    '                    End While
    '                    While LDataTable.Rows(k).Item(SearchColumn) = SearchValue
    '                        mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
    '                        k = k + 1
    '                        If k > last Then
    '                            Exit While
    '                        End If
    '                    End While
    '            End Select
    '            If fnd = True Then
    '                Exit While
    '            End If
    '        End While
    '        If OutputOnSorting.Length > 0 Then
    '            mDataTable = SortDataTable(mDataTable, OutputOnSorting)
    '        End If
    '    Catch ex As Exception
    '        QuitError(ex, Err, "Unable to execute DataFunction.SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer) As DataTable")
    '    End Try
    '    Return mDataTable
    'End Function


    ''' <summary>
    ''' This function searches  value of  specified column of Sorted DataTable on specified column .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    ''' <param name="SearchColumn">Column name to be searched.</param>
    '''<param name="SearchValue" >Column value to be Sought.  </param>
    ''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer, Optional ByVal OutputOnSorting As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        If LDataTable.Rows.Count = 0 Then
            Return mDataTable
            Exit Function
        End If
        Try
            Dim first As Integer = 0
            Dim last As Integer = LDataTable.Rows.Count - 1
            If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
                Return mDataTable
                Exit Function
            End If
            Dim middle As Integer = last \ 2
            Dim fnd As Boolean = False
            While first <= last
                Dim mvalue As Integer = IIf(IsDBNull(LDataTable.Rows(middle).Item(SearchColumn)) = True, 0, CInt(LDataTable.Rows(middle).Item(SearchColumn)))
                Select Case True
                    Case SearchValue > mvalue
                        first = middle + 1
                        middle = (first + last) \ 2
                    Case SearchValue < mvalue
                        last = middle - 1
                        middle = (first + last) \ 2
                    Case mvalue = SearchValue
                        fnd = True
                        Dim k As Integer = middle
                        While mvalue = SearchValue
                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                            middle = middle - 1
                            If middle < 0 Then
                                Exit While
                            End If
                            mvalue = IIf(IsDBNull(LDataTable.Rows(middle).Item(SearchColumn)) = True, 0, CInt(LDataTable.Rows(middle).Item(SearchColumn)))
                        End While
                        mvalue = IIf(IsDBNull(LDataTable.Rows(k).Item(SearchColumn)) = True, 0, CInt(LDataTable.Rows(k).Item(SearchColumn)))
                        While LDataTable.Rows(k).Item(SearchColumn) = SearchValue
                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
                            k = k + 1
                            If k > last Then
                                Exit While
                            End If
                            mvalue = IIf(IsDBNull(LDataTable.Rows(k).Item(SearchColumn)) = True, 0, CInt(LDataTable.Rows(k).Item(SearchColumn)))
                        End While
                End Select
                If fnd = True Then
                    Exit While
                End If
            End While
            If OutputOnSorting.Length > 0 Then
                mDataTable = SortDataTable(mDataTable, OutputOnSorting)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer) As DataTable")
        End Try
        Return mDataTable
    End Function







    '''' <summary>
    '''' This function searches  value of  specified column of Sorted DataTable on specified column .
    '''' </summary>
    '''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    '''' <param name="SearchColumn">Column name to be searched.</param>
    ''''<param name="SearchValue" >Column value to be Sought.  </param>
    '''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    '''' <returns>Datatable on specified criteria</returns>
    '''' <remarks></remarks>
    'Public Function SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer, Optional ByVal OutputOnSorting As String = "") As DataRow
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    Dim mDataTable As DataTable = LDataTable.Clone
    '    Dim mrow As DataRow = Nothing
    '    If LDataTable.Rows.Count = 0 Then
    '        Return mrow
    '        Exit Function
    '    End If


    '    Try
    '        Dim first As Integer = 0
    '        Dim last As Integer = LDataTable.Rows.Count - 1
    '        If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
    '            Return mrow
    '            Exit Function
    '        End If
    '        Dim middle As Integer = last \ 2
    '        Dim fnd As Boolean = False
    '        While first <= last
    '            Select Case True
    '                Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn)
    '                    first = middle + 1
    '                    middle = (first + last) \ 2
    '                Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn)
    '                    last = middle - 1
    '                    middle = (first + last) \ 2
    '                Case LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
    '                    fnd = True
    '                    mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
    '            End Select
    '            If fnd = True Then
    '                Exit While
    '            End If
    '        End While
    '        If OutputOnSorting.Length > 0 Then
    '            mDataTable = SortDataTable(mDataTable, OutputOnSorting)
    '        End If
    '    Catch ex As Exception
    '        QuitError(ex, Err, "Unable to execute DataFunction.SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer) As DataRow")
    '    End Try
    '    If mDataTable.Rows.Count > 0 Then
    '        mrow = mDataTable.Rows(0)
    '    End If
    '    Return mrow
    'End Function


    ''' <summary>
    ''' This function searches  value of  specified column of Sorted DataTable on specified column .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    ''' <param name="SearchColumn">Column name to be searched.</param>
    '''<param name="SearchValue" >Column value to be Sought.  </param>
    ''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer, Optional ByVal OutputOnSorting As String = "") As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        Dim mrow As DataRow = Nothing
        If LDataTable.Rows.Count = 0 Then
            Return mrow
            Exit Function
        End If


        Try
            Dim first As Integer = 0
            Dim last As Integer = LDataTable.Rows.Count - 1
            If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
                Return mrow
                Exit Function
            End If
            Dim middle As Integer = last \ 2
            Dim fnd As Boolean = False
            While first <= last
                Dim mvalue As Integer = IIf(IsDBNull(LDataTable.Rows(middle).Item(SearchColumn)) = True, 0, CInt(LDataTable.Rows(middle).Item(SearchColumn)))
                Select Case True
                    Case SearchValue > mvalue
                        first = middle + 1
                        middle = (first + last) \ 2
                    Case SearchValue < mvalue
                        last = middle - 1
                        middle = (first + last) \ 2
                    Case mvalue = SearchValue
                        fnd = True
                        mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                End Select
                If fnd = True Then
                    Exit While
                End If
            End While
            If OutputOnSorting.Length > 0 Then
                mDataTable = SortDataTable(mDataTable, OutputOnSorting)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As Integer) As DataRow")
        End Try
        If mDataTable.Rows.Count > 0 Then
            mrow = mDataTable.Rows(0)
        End If
        Return mrow
    End Function


    '''' <summary>
    '''' This function searches  value of  specified column of Sorted DataTable on specified column .
    '''' </summary>
    '''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    '''' <param name="SearchColumn">Column name to be searched.</param>
    ''''<param name="SearchValue" >Column value to be Sought.  </param>
    '''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    '''' <returns>Datatable on specified criteria</returns>
    '''' <remarks></remarks>
    'Public Function SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String, Optional ByVal OutputOnSorting As String = "") As DataRow
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    Dim mDataTable As DataTable = LDataTable.Clone
    '    Dim mrow As DataRow = Nothing
    '    If LDataTable.Rows.Count = 0 Then
    '        Return mrow
    '        Exit Function
    '    End If
    '    Try
    '        Dim first As Integer = 0
    '        Dim last As Integer = LDataTable.Rows.Count - 1
    '        If SearchValue > LDataTable.Rows(last).Item(SearchColumn) Or SearchValue < LDataTable.Rows(first).Item(SearchColumn) Then
    '            Return mrow
    '            Exit Function
    '        End If
    '        Dim middle As Integer = last \ 2
    '        Dim fnd As Boolean = False
    '        While first <= last
    '            Select Case True
    '                Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn)
    '                    first = middle + 1
    '                    middle = (first + last) \ 2
    '                Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn)
    '                    last = middle - 1
    '                    middle = (first + last) \ 2
    '                Case LDataTable.Rows(middle).Item(SearchColumn) = SearchValue
    '                    fnd = True
    '                    mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
    '            End Select
    '            If fnd = True Then
    '                Exit While
    '            End If
    '        End While
    '        If OutputOnSorting.Length > 0 Then
    '            mDataTable = SortDataTable(mDataTable, OutputOnSorting)
    '        End If
    '    Catch ex As Exception
    '        QuitError(ex, Err, "Unable to execute DataFunction.SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String) As DataRow")
    '    End Try
    '    If mDataTable.Rows.Count > 0 Then
    '        mrow = mDataTable.Rows(0)
    '    End If
    '    Return mrow
    'End Function

    ''' <summary>
    ''' This function searches  value of  specified column of Sorted DataTable on specified column .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    ''' <param name="SearchColumn">Column name to be searched.</param>
    '''<param name="SearchValue" >Column value to be Sought.  </param>
    ''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String, Optional ByVal OutputOnSorting As String = "") As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        Dim mrow As DataRow = Nothing
        If LDataTable.Rows.Count = 0 Then
            Return mrow
            Exit Function
        End If
        Try
            Dim first As Integer = 0
            Dim last As Integer = LDataTable.Rows.Count - 1
            If SearchValue > LDataTable.Rows(last).Item(SearchColumn).ToString Or SearchValue < LDataTable.Rows(first).Item(SearchColumn).ToString Then
                Return mrow
                Exit Function
            End If
            Dim middle As Integer = last \ 2
            Dim fnd As Boolean = False
            While first <= last
                Select Case True
                    Case SearchValue > LDataTable.Rows(middle).Item(SearchColumn).ToString
                        first = middle + 1
                        middle = (first + last) \ 2
                    Case SearchValue < LDataTable.Rows(middle).Item(SearchColumn).ToString
                        last = middle - 1
                        middle = (first + last) \ 2
                    Case LDataTable.Rows(middle).Item(SearchColumn).ToString = SearchValue
                        fnd = True
                        mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                End Select
                If fnd = True Then
                    Exit While
                End If
            End While
            If OutputOnSorting.Length > 0 Then
                mDataTable = SortDataTable(mDataTable, OutputOnSorting)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchRowInSortedDataTable(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String) As DataRow")
        End Try
        If mDataTable.Rows.Count > 0 Then
            mrow = mDataTable.Rows(0)
        End If
        Return mrow
    End Function


    ''' <summary>
    ''' This function searches string in a specified column of Sorted DataTable on search column .
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched originally sorted on SearchColumn.</param>
    ''' <param name="SearchColumn">Column name to be searched.</param>
    '''<param name="SearchString" >String value to be Sought from left.  </param>
    ''' <param name="OutputOnSorting" >Comma separated columnnames on which output datatable must be sorted</param>
    ''' <returns>Datatable on specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchStringInSortedDTColumn(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchString As String, Optional ByVal OutputOnSorting As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mDataTable As DataTable = LDataTable.Clone
        Try
            Dim first As Integer = 0
            Dim last As Integer = LDataTable.Rows.Count - 1
            SearchString = LCase(Trim(SearchString))
            Dim mlen As Int16 = SearchString.Length

            If SearchString > LCase(Left(LDataTable.Rows(last).Item(SearchColumn).ToString, mlen)) Or SearchString < LCase(Left(LDataTable.Rows(first).Item(SearchColumn).ToString, mlen)) Then
                Return mDataTable
                Exit Function
            End If
            Dim middle As Integer = last \ 2
            Dim fnd As Boolean = False
            While first <= last
                Select Case True
                    Case SearchString > LCase(Left(LDataTable.Rows(middle).Item(SearchColumn).ToString, mlen))
                        first = middle + 1
                        middle = (first + last) \ 2
                    Case SearchString < LCase(Left(LDataTable.Rows(middle).Item(SearchColumn).ToString, mlen))
                        last = middle - 1
                        middle = (first + last) \ 2
                    Case SearchString = LCase(Left(LDataTable.Rows(middle).Item(SearchColumn).ToString, mlen))
                        fnd = True
                        Dim k As Integer = middle
                        While SearchString = LCase(Left(LDataTable.Rows(middle).Item(SearchColumn).ToString, mlen))
                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(middle), True)
                            middle = middle - 1
                            If middle < 0 Then
                                Exit While
                            End If
                        End While
                        While SearchString = LCase(Left(LDataTable.Rows(k).Item(SearchColumn).ToString, mlen))
                            mDataTable = AddRowInDataTable(mDataTable, LDataTable.Rows(k), True)
                            k = k + 1
                            If k > last Then
                                Exit While
                            End If
                        End While
                End Select
                If fnd = True Then
                    If OutputOnSorting.Length > 0 Then
                        mDataTable = SortDataTable(mDataTable, OutputOnSorting)
                    End If
                    Exit While
                End If
            End While
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchStringInSortedDTColumn(ByVal LDataTable As DataTable, ByVal SearchColumn As String, ByVal SearchValue As String) As DataTable")
        End Try
        Return mDataTable
    End Function


    ''' <summary>
    ''' This function searches a DataTable on specified columns with values and return datatable after removing the rows that satisfy the criteria.
    ''' </summary>
    ''' <param name="LDataTable">DataTable with duplicate rows  on specified columns</param>
    ''' <param name="SearchColumns  ">Comma separated column names to be searched eg column1,column2,column3} etc.</param>
    '''<param name="SearchValues" >Comma separated string values to be searched eg. {value1,value2,value3...  </param>
    ''' <returns>Datatable after removing rows</returns>
    ''' <remarks></remarks>
    Public Function RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Not Split(SearchColumns, ",").Count = Split(SearchValues, ",").Count Then
            QuitMessage("Size of SearchColumns differs SearchValues" & " SearchDataTable", "RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable  ")
        End If
        Dim mDataTable As DataTable = LDataTable
        Dim SearchExpr As String = ""
        Dim aSearchColumns() As String = Split(SearchColumns, ",")
        Dim aSearchvalues() As String = Split(SearchValues, ",")
        For i = 0 To aSearchColumns.Count - 1
            Dim mtype As String = LCase(LDataTable.Columns(aSearchColumns(i)).DataType.Name.ToString)
            Select Case LCase(mtype)
                Case "string", "datetime", "date"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " or ") & aSearchColumns(i) & " <> '" & aSearchvalues(i).ToString & "'"
                Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " or ") & aSearchColumns(i) & " <> " & aSearchvalues(i).ToString
                Case Else
                    Continue For
            End Select
        Next
        Try
            Dim mRows() As DataRow = LDataTable.Select(SearchExpr)
            If mRows.Count > 0 Then
                mDataTable = mRows.CopyToDataTable
            Else
                mDataTable.Rows.Clear()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal SearchValues As String) As DataTable")
        End Try
        Return mDataTable
    End Function

    ''' <summary>
    ''' Removed specified rows from a table.
    ''' </summary>
    ''' <param name="LDataTable">DataTable</param>
    ''' <param name="HashColumnValues  ">A Hashtable having the keys are column names and values are column values.</param>
    ''' <returns>Datatable after removing rows</returns>
    ''' <remarks></remarks>
    Public Function RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal HashColumnValues As Hashtable) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim SearchExpr As String = GF1.GetStringConditionFromHashTable(HashColumnValues, True, True)
        Dim mDataTable As DataTable = LDataTable

        Try
            Dim mRows() As DataRow = LDataTable.Select(SearchExpr)
            If mRows.Count > 0 Then
                mDataTable = mRows.CopyToDataTable
            Else
                mDataTable.Rows.Clear()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunctionPublic Function RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal HashColumnValues As Hashtable) As DataTable")
        End Try
        Return mDataTable
    End Function
    ''' <summary>
    ''' Remove data row from a datatable on primary key.
    ''' </summary>
    ''' <param name="LDataTable">DataTable with duplicate rows  on specified columns</param>
    '''<param name="PrimaryKeyValue" >Value of primary key to be removed  </param>
    ''' <returns>Datatable after removing rows</returns>
    ''' <remarks></remarks>
    Public Function RemoveRowFromDataTable(ByVal LDataTable As DataTable, ByVal PrimaryKeyValue As Integer) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mind As Integer = FindRowIndexByPrimaryCols(LDataTable, PrimaryKeyValue)
            If mind > -1 Then
                LDataTable.Rows.RemoveAt(mind)
            End If
            Return LDataTable
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RemoveRowFromDataTable(ByVal LDataTable As DataTable, ByVal PrimaryKeyValue As Integer) As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    ''' This function searches a DataTable on specified columns with values and return datatable after removing the rows that satisfy the criteria.
    ''' </summary>
    ''' <param name="LDataTable">DataTable with duplicate rows  on specified columns</param>
    ''' <param name="SearchColumns  ">Comma separated column names to be searched eg column1,column2,column3} etc.</param>
    '''<param name="aSearchValues" >An array of value object to be searched eg. {value1,value2,value3...  </param>
    ''' <returns>Datatable after removing rows</returns>
    ''' <remarks></remarks>
    Public Function RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal aSearchValues() As Object) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If Not Split(SearchColumns, ",").Count = aSearchValues.Count Then
            QuitMessage("Size of SearchColumns differs SearchValues" & " SearchDataTable", "RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal aSearchValues() As Object) As DataTable   " & "RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal aSearchValues() As Object) As DataTable")
        End If
        Dim mDataTable As DataTable = LDataTable
        Dim SearchExpr As String = ""
        Dim aSearchColumns() As String = Split(SearchColumns, ",")
        'Dim aSearchvalues() As String = Split(SearchValues, ",")
        For i = 0 To aSearchColumns.Count - 1
            Dim mtype As String = LCase(LDataTable.Columns(aSearchColumns(i)).DataType.Name.ToString)
            Select Case LCase(mtype)
                Case "string", "datetime", "date"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " or ") & aSearchColumns(i) & " <> '" & aSearchValues(i).ToString & "'"
                Case "integer", "int32", "int64", "int16", "decimal", "double", "single", "byte"
                    SearchExpr = SearchExpr & IIf(SearchExpr.Length = 0, "", " or ") & aSearchColumns(i) & " <> " & aSearchValues(i).ToString
                Case Else
                    Continue For
            End Select
        Next
        Try
            Dim mRows() As DataRow = LDataTable.Select(SearchExpr)
            If mRows.Count > 0 Then
                mDataTable = mRows.CopyToDataTable
            Else
                mDataTable.Rows.Clear()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.RemoveRowsFromDataTable(ByVal LDataTable As DataTable, ByVal SearchColumns As String, ByVal aSearchValues() As Object) As DataTable")
        End Try
        Return mDataTable
    End Function



    ''' <summary>
    ''' This function searches a DataTable on specified columns with values and return datatable satisfying the criteria.
    ''' </summary>
    ''' <param name="LDataTable">DataTable to be searched on specified columns</param>
    ''' <param name="SearchColumnsValues  ">A hashtable containg columnnames as keys and column values as valu of hashtable .</param>
    ''' <returns>Datatable of specified criteria</returns>
    ''' <remarks></remarks>
    Public Function SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumnsValues As Hashtable) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mtable As New DataTable

        Try
            Dim searchColumns() As String = {}
            Dim searchvalues() As Object = {}
            GF1.ConvertHashTableToArrays(SearchColumnsValues, searchColumns, searchvalues)
            mtable = SearchDataTable(LDataTable, searchColumns, searchvalues)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.SearchDataTable(ByVal LDataTable As DataTable, ByVal SearchColumnsValues As Hashtable) As DataTable")
        End Try

        Return mtable
    End Function

    ''' <summary>
    ''' This function creates a unique/distinct rows datatable on  specified columns from a datatable and make totals of numeric columns.
    ''' </summary>
    ''' <param name="LDataTable">DataTable with duplicate rows  on specified columns</param>
    ''' <param name="TotalOnColumns">comma separated  column names on which total of numeric columns will be calculated eg. "Column1,Column2, ..  } etc.</param>
    ''' <param name="NumericColumns" >comma separated  numeric columns names which are added to get total </param>
    ''' <param name="AllColumns" >if false the returning total table will has TotalOnColumns+Numeric Columns only,if true it has the same columns as original datatable</param>
    '''<param name="ColumnsOfSameValues" >A hashtable which has keys as columnnames,values as columnvalues, which are same for all rows of TotalTable</param>
    ''' <returns>Unique rows datatable</returns>
    ''' <remarks></remarks>
    Public Function TotalOnDataTable(ByVal LDataTable As DataTable, ByVal TotalOnColumns As String, ByVal NumericColumns As String, Optional ByVal AllColumns As Boolean = False, Optional ByVal ColumnsOfSameValues As Hashtable = Nothing) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim UniqueTable As New DataTable
        Try
            UniqueTable = GetDistinctRowsFromDataTable(LDataTable, TotalOnColumns, AllColumns)
            Dim aNumericColumns() As String = Split(NumericColumns, ",")
            Dim aTotalOnColumns() As String = Split(TotalOnColumns, ",")
            If AllColumns = False Then
                For k = 0 To aNumericColumns.Count - 1
                    UniqueTable.Columns.Add(aNumericColumns(k))
                Next
            End If
            For k = 0 To UniqueTable.Rows.Count - 1
                For j = 0 To aNumericColumns.Count - 1
                    UniqueTable.Rows(k).Item(aNumericColumns(j)) = 0
                Next
                If ColumnsOfSameValues IsNot Nothing Then
                    For j = 0 To ColumnsOfSameValues.Count - 1
                        Dim mkey As String = LCase(ColumnsOfSameValues(j))
                        UniqueTable.Rows(k).Item(mkey) = ColumnsOfSameValues.Item(mkey)
                    Next
                End If
            Next
            For i = 0 To LDataTable.Rows.Count - 1
                Dim ColumnValues() As Object = {}
                For k = 0 To aTotalOnColumns.Count - 1
                    GF1.ArrayAppend(ColumnValues, LDataTable.Rows(i).Item(aTotalOnColumns(k)))
                Next
                Dim mRowIndex() As Integer = SearchDataTableRowIndex(UniqueTable, aTotalOnColumns, ColumnValues, True)
                For k = 0 To aNumericColumns.Count - 1
                    UniqueTable.Rows(mRowIndex(0)).Item(aNumericColumns(k)) = UniqueTable.Rows(mRowIndex(0)).Item(aNumericColumns(k)) + LDataTable.Rows(i).Item(aNumericColumns(k))
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.TotalOnDataTable(ByVal LDataTable As DataTable, ByVal TotalOnColumns As String, ByVal NumericColumns As String, Optional ByVal AllColumns As Boolean = False, Optional ByVal ColumnsOfSameValues As Hashtable = Nothing) As DataTable")
        End Try
        Return UniqueTable
    End Function
    ''' <summary>
    ''' To Get Parent nodes/groups of a groupcode from a GroupMasterDataTable as an array of ParentGroupCodes
    ''' </summary>
    ''' <param name="DtGroupMaster">GroupMaster datatable where groupcode to be searched </param>
    ''' <param name="GroupCodeField">Name of Groupcode Field in GroupMaster </param>
    ''' <param name="GroupCodeValue">Value of GroupCode whoose parents to be found</param>
    ''' <param name="ParentCodeField">Name of ParentCode Field</param>
    ''' <param name="LastParentValue">Value of Last/Top parent code </param>
    ''' <param name="ParentCodeCeiling">Final size of parent array,If No. of parents are less than this then add lastparentvalue to compele the size,(-1)  for actual size</param>
    ''' <returns>An array of parent codes</returns>
    ''' <remarks></remarks>
    Function GetParentNodes(ByVal DtGroupMaster As DataTable, ByVal GroupCodeField As String, ByVal GroupCodeValue As String, ByVal ParentCodeField As String, ByVal LastParentValue As String, Optional ByVal ParentCodeCeiling As Integer = -1) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mParent() As String = {}
        Try
StartAgain:
            If Not GroupCodeValue = LastParentValue Then
                Dim expr As String = GroupCodeField & " = " & "'" & GroupCodeValue & "'"
                Dim mGroupCodeRow As DataRow() = DtGroupMaster.Select(expr)
                Dim mparentcode As String = mGroupCodeRow(0).Item(ParentCodeField)
                If mparentcode <> LastParentValue Then
                    GF1.ArrayAppend(mParent, mparentcode)
                End If
                If mparentcode = LastParentValue Then
                    Select Case ParentCodeCeiling
                        Case -1
                            GF1.ArrayAppend(mParent, LastParentValue)
                        Case Else
                            For ii = mParent.Count To ParentCodeCeiling - 1
                                GF1.ArrayAppend(mParent, LastParentValue)
                            Next
                    End Select
                    Return mParent
                    Exit Function
                End If
                GroupCodeValue = mparentcode
                GoTo StartAgain
            Else
                Select Case ParentCodeCeiling
                    Case -1
                        GF1.ArrayAppend(mParent, LastParentValue)
                    Case Else
                        For ii = mParent.Count To ParentCodeCeiling
                            GF1.ArrayAppend(mParent, LastParentValue)
                        Next
                End Select

            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetParentNodes(ByVal DtGroupMaster As DataTable, ByVal GroupCodeField As String, ByVal GroupCodeValue As String, ByVal ParentCodeField As String, ByVal LastParentValue As String, Optional ByVal ParentCodeCeiling As Integer = -1) As String()")
        End Try
        Return mParent
    End Function
    ''' <summary>
    ''' This Function converts dateTime.Now  of server into dateTime.Now of IST time zone.
    ''' </summary>
    ''' <returns>Current date in IST</returns>
    Public Function getDateTimeISTNow() As DateTime
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim serverdate As DateTime = DateTime.UtcNow()
        Dim tzi As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time")
        Dim localdateTime As DateTime = TimeZoneInfo.ConvertTimeFromUtc(serverdate, tzi)
        Return localdateTime
    End Function
    ''' <summary>
    ''' This function converts dateTime.Now of server into dateTime.Now of UTC time.
    ''' </summary>
    ''' <returns>Current date in UTC</returns>
    Public Function getDateTimeUTCNow() As DateTime
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim serverdate As DateTime = DateTime.UtcNow()
        Dim utcdate As DateTime = serverdate.ToUniversalTime()
        Return utcdate
    End Function

    ''' <summary>
    ''' This function will convert date from IST to UTC or viceversa. Date argument should be either in IST or UTC time zone.
    ''' </summary>
    ''' <param name="dt">Date to be converted. If Zone is IST then it should be in UTC timezone and viceversa </param>
    ''' <param name="zone">Zone is the time zone in which dt is to be converted. It can only be IST or UTC</param>
    ''' <returns></returns>
    Public Function ConvertDateTimeUTC_IST(ByVal dt As DateTime, ByVal zone As String) As DateTime
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim tempdate As DateTime = DateTime.Now
        If dt <> Nothing Then
            tempdate = dt
        End If
        Dim ConvertedDate As DateTime
        Select Case LCase(zone)
            Case "ist"
                Dim tzi As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time")
                ConvertedDate = TimeZoneInfo.ConvertTimeFromUtc(tempdate, tzi)
            Case "utc"
                ConvertedDate = tempdate.ToUniversalTime()
            Case Else
                ConvertedDate = Nothing
        End Select

        Return ConvertedDate
    End Function




    ''' <summary>
    ''' To add row in data table ,whoose column values are given as hash table
    ''' </summary>
    ''' <param name="LDataTable">Data table in which rows are to be added</param>
    ''' <param name="ColumnValues">ColumnValues as hashtable where key is columnname and value is its content</param>
    ''' <param name="CheckColumnValues">Check wether row exists for specified columns of separate hash table</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddRowInDataTable(ByVal LDataTable As DataTable, ByVal ColumnValues As Hashtable, Optional ByVal CheckColumnValues As Hashtable = Nothing) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mrow As DataRow = LDataTable.NewRow
            Dim mindex() As Integer = {}
            If Not CheckColumnValues Is Nothing Then
                mindex = SearchDataTableRowIndex(LDataTable, CheckColumnValues, True)
            End If
            Dim mColumns() As String = {}
            Dim mvalues() As Object = {}
            GF1.ConvertHashTableToArrays(ColumnValues, mColumns, mvalues)
            For i = 0 To mColumns.Count - 1
                If mindex.Count = 0 Then
                    mrow(mColumns(i)) = mvalues(i)
                Else
                    LDataTable.Rows(mindex(0))(mColumns(i)) = mvalues(i)
                End If
            Next
            If mindex.Count = 0 Then
                LDataTable.Rows.Add(mrow)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddRowInDataTable(ByVal LDataTable As DataTable, ByVal ColumnValues As Hashtable, Optional ByVal CheckColumnValues As Hashtable = Nothing) As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    ''' To add row in data table ,whoose column values are given as hash table
    ''' </summary>
    ''' <param name="LDataTable">Data table in which rows are to be added</param>
    ''' <param name="RepeatRowAsNew">RowNo whoose ColumnValues will be repeated</param>
    ''' <param name="ReplacedColumnValues">ColumnValues as hash table to be replaced,Key is columnname,Value is column value</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddRowInDataTable(ByVal LDataTable As DataTable, ByVal RepeatRowAsNew As Integer, Optional ByVal ReplacedColumnValues As Hashtable = Nothing) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mrow As DataRow = LDataTable.NewRow
            Dim orow As DataRow = LDataTable.Rows(RepeatRowAsNew)
            Dim mColumns() As String = {}
            Dim mvalues() As Object = {}
            GF1.ConvertHashTableToArrays(ReplacedColumnValues, mColumns, mvalues)
            For i = 0 To LDataTable.Columns.Count - 1
                mrow(i) = orow(i)
            Next
            For i = 0 To mColumns.Count - 1
                mrow(mColumns(i)) = mvalues(i)
            Next
            LDataTable.Rows.Add(mrow)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddRowInDataTable(ByVal LDataTable As DataTable, ByVal RepeatRowAsNew As Integer, Optional ByVal ReplacedColumnValues As Hashtable = Nothing) As DataTable")
        End Try

        Return LDataTable
    End Function
    ''' <summary>
    ''' To add row in data table ,whoose column values are given as hash table
    ''' </summary>
    ''' <param name="LDataTable">Data table in which rows are to be added</param>
    ''' <param name="RepeatRowAsNew">DataRow  whoose ColumnValues will be repeated</param>
    ''' <param name="ReplacedColumnValues">ColumnValues as hash table to be replaced,Key is columnname,Value is column value</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddRowInDataTable(ByVal LDataTable As DataTable, ByVal RepeatRowAsNew As DataRow, Optional ByVal ReplacedColumnValues As Hashtable = Nothing) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mrow As DataRow = LDataTable.NewRow
            For i = 0 To LDataTable.Columns.Count - 1
                mrow(i) = RepeatRowAsNew(i)
            Next
            If ReplacedColumnValues IsNot Nothing Then
                Dim mColumns() As String = {}
                Dim mvalues() As Object = {}
                GF1.ConvertHashTableToArrays(ReplacedColumnValues, mColumns, mvalues)
                For i = 0 To mColumns.Count - 1
                    mrow(mColumns(i)) = mvalues(i)
                Next
            End If
            LDataTable.Rows.Add(mrow)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddRowInDataTable(ByVal LDataTable As DataTable, ByVal RepeatRowAsNew As DataRow, Optional ByVal ReplacedColumnValues As Hashtable = Nothing) As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    ''' To add row in data table ,whoose column values are given as hash table
    ''' </summary>
    ''' <param name="LDataTable">Data table in which rows are to be added</param>
    ''' <param name="NewDataRow">DataRow  which is added to datatable.</param>
    ''' <param name="NoAdditionIfExists">If row exists, no datarow will be added</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddRowInDataTable(ByVal LDataTable As DataTable, ByVal NewDataRow As DataRow, ByVal NoAdditionIfExists As Boolean) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim checkrowno As Integer = -1
            If NoAdditionIfExists = True Then
                checkrowno = CheckRowInDataTable(NewDataRow, LDataTable)
            End If
            If checkrowno = -1 Then
                Dim mrow As DataRow = LDataTable.NewRow
                For i = 0 To LDataTable.Columns.Count - 1
                    mrow(i) = NewDataRow(i)
                Next
                LDataTable.Rows.Add(mrow)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddRowInDataTable(ByVal LDataTable As DataTable, ByVal RepeatRowAsNew As DataRow, Optional ByVal ReplacedColumnValues As Hashtable = Nothing) As DataTable")
        End Try
        Return LDataTable
    End Function
    ''' <summary>
    ''' To add row in data table ,whoose column values are given as hash table
    ''' </summary>
    ''' <param name="LDataTable">Data table in which rows are to be added</param>
    ''' <param name="DataRowsArray">An array of DataRows  which is added to datatable.</param>
    ''' <param name="NoAdditionIfExists">If row exists, no datarow will be added</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddRowInDataTable(ByVal LDataTable As DataTable, ByVal DataRowsArray() As DataRow, ByVal NoAdditionIfExists As Boolean) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For k = 0 To DataRowsArray.Count - 1
                Dim checkrowno As Integer = -1
                If NoAdditionIfExists = True Then
                    checkrowno = CheckRowInDataTable(DataRowsArray(k), LDataTable)
                End If
                If checkrowno = -1 Then
                    Dim mrow As DataRow = LDataTable.NewRow
                    mrow.ItemArray = DataRowsArray(k).ItemArray
                    LDataTable.Rows.Add(mrow)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddRowInDataTable(ByVal LDataTable As DataTable, ByVal RepeatRowAsNew As DataRow, Optional ByVal ReplacedColumnValues As Hashtable = Nothing) As DataTable")
        End Try
        Return LDataTable
    End Function


    ''' <summary>
    ''' To assign column values to a data row 
    ''' </summary>
    ''' <param name="BaseRow" >DataRow inwhich values to assigned</param>
    ''' <param name="ColumnValues" >ColumnValues as hashtable, with key as columnname and value as columnvalue</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AssignValuesToRow(ByVal BaseRow As DataRow, ByVal ColumnValues As Hashtable) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mColumns() As String = {}
            Dim mvalues() As Object = {}
            GF1.ConvertHashTableToArrays(ColumnValues, mColumns, mvalues)
            AssignValuesToRow(BaseRow, mColumns, mvalues)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AssignValuesToRow(ByVal BaseRow As DataRow, ByVal ColumnValues As Hashtable) As DataRow")
        End Try
        Return BaseRow
    End Function
    ''' <summary>
    ''' To assign column values to a data row 
    ''' </summary>
    ''' <param name="BaseRow" >DataRow inwhich values to assigned</param>
    ''' <param name="ColumnNames" >Column names as string array</param>
    ''' <param name="ColumnValues" >Column Values as object array</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AssignValuesToRow(ByVal BaseRow As DataRow, ByVal ColumnNames() As String, ByVal ColumnValues() As Object) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function


        Try

            For i = 0 To ColumnNames.Count - 1
                BaseRow(ColumnNames(i)) = ColumnValues(i)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AssignValuesToRow(ByVal BaseRow As DataRow, ByVal ColumnNames() As String, ByVal ColumnValues() As Object) As DataRow")

        End Try

        Return BaseRow
    End Function
    ''' <summary>
    ''' To add columns  in data table 
    ''' </summary>
    ''' <param name="LDataTable">Data table in which columns  are to be added</param>
    ''' <param name="ColumnNames">Comma separated ColumnNames  as string></param>
    ''' <param name="ColumnTypes">Comma separated system.types as string eg system.string,system.decimal,system.int16,system.int32,system.byte[] etc</param>
    ''' <param name="ColumnsAfter">Existing ColumnName after which columns added</param>
    ''' <param name="ColumnsBefore">Existing ColumnName before which columns added</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddColumnsInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, Optional ByVal ColumnTypes As String = "", Optional ByVal ColumnsAfter As String = "", Optional ByVal ColumnsBefore As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim acolumns() As String = Split(ColumnNames, ",")
            Dim atypes() As String = Split(ColumnTypes, ",")
            If acolumns.Count <> atypes.Count And ColumnTypes.Trim.Length > 0 Then
                QuitMessage("No. of columns differs no. of column types", "AddColumnsInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, Optional ByVal ColumnTypes As String = "", Optional ByVal ColumnsAfter As String = "", Optional ByVal ColumnsBefore As String = "") As DataTable ")
                Return Nothing
                Exit Function
            End If
            Dim mordinal As Integer = -1
            If ColumnsAfter.Trim.Length > 0 Then
                mordinal = CheckColumnInDataTable(ColumnsAfter, LDataTable) + 1
            End If
            If ColumnsBefore.Trim.Length > 0 Then
                mordinal = CheckColumnInDataTable(ColumnsBefore, LDataTable) - 1
                If mordinal = -1 Then
                    mordinal = 0
                End If
            End If

            For i = 0 To acolumns.Count - 1
                If CheckColumnInDataTable(acolumns(i), LDataTable) = -1 Then
                    Dim mtype As String = ""
                    If ColumnTypes.Trim.Length > 0 Then
                        mtype = atypes(i).Trim
                        If LCase(Left(mtype, 7)) <> "system." Then
                            mtype = "system." & mtype
                        End If
                    End If
                    Select Case True
                        Case mtype = "" And mordinal < 0
                            LDataTable.Columns.Add(acolumns(i).Trim)
                        Case mtype = "" And mordinal > -1
                            LDataTable.Columns.Add(acolumns(i)).SetOrdinal(mordinal)
                            mordinal = mordinal + 1
                        Case mtype > "" And mordinal < 0
                            LDataTable.Columns.Add(acolumns(i).Trim, System.Type.GetType(mtype, True, True))
                        Case mtype > "" And mordinal > -1
                            LDataTable.Columns.Add(acolumns(i).Trim, System.Type.GetType(mtype, True, True)).SetOrdinal(mordinal)
                            mordinal = mordinal - 1
                            If mordinal < 0 Then
                                mordinal = 0
                            End If
                        Case Else
                            LDataTable.Columns.Add(acolumns(i).Trim)
                    End Select
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.AddColumnsInDataTable(ByVal LDataTable As DataTable, ByVal ColumnNames As String, Optional ByVal ColumnTypes As String = "", Optional ByVal ColumnsAfter As String = "", Optional ByVal ColumnsBefore As String = "") As DataTable")
        End Try

        Return LDataTable
    End Function
    ''' <summary>
    ''' To get an array objects of given comma separated columns of a datarow 
    ''' </summary>
    ''' <param name="LdataRow">DataRow of columns</param>
    ''' <param name="LcolumnNames">Comma separated string of column names</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetColumnValuesFromRow(ByVal LdataRow As DataRow, ByVal LcolumnNames As String) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aColumns() As String = Split(LcolumnNames, ",")
        Dim aObject() As Object = {}
        Try
            For i = 0 To aColumns.Count - 1
                GF1.ArrayAppend(aObject, LdataRow(aColumns(i)))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetColumnValuesFromRow(ByVal LdataRow As DataRow, ByVal LcolumnNames As String) As Object()")
        End Try

        Return aObject
    End Function
    ''' <summary>
    ''' To get a comma separated string values  of comma separated columns of a datarow 
    ''' </summary>
    ''' <param name="LdataRow">DataRow of columns</param>
    ''' <param name="LcolumnNames">Comma separated string of column names</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetColumnValuesFromRowAsString(ByVal LdataRow As DataRow, ByVal LcolumnNames As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aColumns() As String = Split(LcolumnNames, ",")
        Dim aObject As String = ""
        Try
            For i = 0 To aColumns.Count - 1
                aObject = aObject & IIf(aObject.Length = 0, "", ",") & LdataRow(aColumns(i))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.GetColumnValuesFromRowAsString(ByVal LdataRow As DataRow, ByVal LcolumnNames As String) As String")
        End Try

        Return aObject
    End Function


    ''' <summary>
    ''' To compare a valueObject with ColumnValues oject of a datarow of given columns
    ''' </summary>
    ''' <param name="LdataRow">DataRow of columns</param>
    ''' <param name="LcolumnNames">Comma separated string of column names</param>
    ''' <param name="ValueObject">An array of obects to be compared with array of columnsvalue object of a datarow</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function MatchColumnValuesOfRow(ByVal LdataRow As DataRow, ByVal LcolumnNames As String, ByVal ValueObject() As Object) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aColumns() As String = Split(LcolumnNames, ",")
        Dim aObject() As Object = GetColumnValuesFromRow(LdataRow, LcolumnNames)
        If aColumns.Count <> ValueObject.Count Then
            QuitMessage("Value object size does not match column nos.", "MatchColumnValuesOfRow(ByVal LdataRow As DataRow, ByVal LcolumnNames As String, ByVal ValueObject() As Object) As Boolean  ")
        End If
        Dim matched As Boolean = True
        Try
            For i = 0 To aColumns.Count - 1
                If aObject(i).Equals(ValueObject(i)) = False Then
                    matched = False
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.MatchColumnValuesOfRow(ByVal LdataRow As DataRow, ByVal LcolumnNames As String, ByVal ValueObject() As Object) As Boolean")
        End Try
        Return matched
    End Function
   
    '    Public Sub CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, ByRef LProgressBar As UserProgressBar.ProgressBar, Optional ByVal LeaveColumns As String = "")
    '        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
    '        Dim misValue As Object = System.Reflection.Missing.Value
    '        Try
    '            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.ApplicationClass
    '            Dim wb As Excel.Workbook = xlApp.Workbooks.Add(misValue)
    '            Dim ws As Excel.Worksheet = wb.Sheets("Sheet1")
    '            wb = xlApp.Workbooks.Add(misValue)
    '            ws = wb.Sheets("Sheet1")
    '            Dim aLeaveColumns() As String = {}
    '            If LeaveColumns.Trim.Length > 0 Then
    '                aLeaveColumns = Split(LeaveColumns, ",")
    '                For i = 0 To aLeaveColumns.Count - 1
    '                    Try
    '                        Dim mcol As String = aLeaveColumns(i)
    '                        LdataTable.Columns.Remove(mcol)
    '                    Catch ex As Exception
    '                        QuitError(ex, Err, "Error in removing columns")
    '                    End Try
    '                Next
    '            End If
    '            For j As Integer = 0 To LdataTable.Columns.Count - 1
    '                If LdataTable.Columns(j).ColumnName IsNot Nothing Then
    '                    ws.Cells(1, j + 1) = LdataTable.Columns(j).ColumnName.ToString.Trim
    '                End If
    '            Next
    '            Dim totrows As Integer = LdataTable.Rows.Count
    '            Dim mFactor As Decimal = 0
    '            If LProgressBar IsNot Nothing Then
    '                LProgressBar.Visible = True
    '                If totrows > 0 Then
    '                    mFactor = 100 / totrows
    '                End If
    '            End If
    '            For i As Integer = 0 To LdataTable.Rows.Count - 1
    '                If LProgressBar IsNot Nothing Then
    '                    'System.Threading.Thread.Sleep(0)
    '                    Dim valPer As Decimal = ((i + 1) * mFactor / 100)
    '                    Dim valStr As String = Math.Round((i + 1), 0).ToString & " / " & totrows.ToString
    '                    Dim IsCancelPending As Boolean = LProgressBar.SetProgressBar(valPer, valStr)
    '                    If IsCancelPending Then           '--------When user clicks on the close button of the ProgressBar form then stop the process & Exit Sub---------'
    '                        LProgressBar.Visible = False
    '                        GoTo EndPara
    '                    End If
    '                End If
    '                For j As Integer = 0 To LdataTable.Columns.Count - 1
    '                    If IsDBNull(LdataTable.Rows(i)(j)) = False Then
    '                        Try
    '                            ws.Cells(i + 2, j + 1) = LdataTable.Rows(i)(j).ToString.Trim
    '                        Catch ex As Exception
    '                            Continue For
    '                        End Try
    '                    End If
    '                Next

    '            Next
    'EndPara:
    '            xlApp.DisplayAlerts = False
    '            ws.SaveAs(OutputExcelFile)
    '            '  ws.Delete()
    '            wb.Close()
    '            xlApp.Quit()
    '            If LProgressBar IsNot Nothing Then
    '                LProgressBar.Visible = False
    '                LProgressBar.Value = 0
    '            End If
    '        Catch ex As Exception
    '            QuitError(ex, Err, "Unable to execute DataFunction.CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, Optional ByVal LeaveColumns As String = "", Optional ByVal LProgressBar As UserProgressBar.ProgressBar = Nothing)")
    '        End Try
    '    End Sub

    ''' <summary>
    ''' Create an excel worksheet by a datatable
    ''' </summary>  
    ''' <param name="LdataTable">DataTable to be exported as work sheet</param>
    ''' <param name="OutputExcelFile">Full identifier of excel sheet</param>
    ''' <param name="LeaveColumns" >Comma separated columns name which are not transfered</param>
    ''' <remarks></remarks>
    Public Sub CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, Optional ByVal LeaveColumns As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim misValue As Object = System.Reflection.Missing.Value
        Try
            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.ApplicationClass
            Dim wb As Excel.Workbook = xlApp.Workbooks.Add(misValue)
            Dim ws As Excel.Worksheet = wb.Sheets("Sheet1")
            wb = xlApp.Workbooks.Add(misValue)
            ws = wb.Sheets("Sheet1")
            Dim aLeaveColumns() As String = {}
            If LeaveColumns.Trim.Length > 0 Then
                aLeaveColumns = Split(LeaveColumns, ",")
                For i = 0 To aLeaveColumns.Count - 1
                    Try
                        Dim mcol As String = aLeaveColumns(i)
                        LdataTable.Columns.Remove(mcol)
                    Catch ex As Exception
                        QuitError(ex, Err, "Error in removing columns")
                    End Try
                Next
            End If
            For j As Integer = 0 To LdataTable.Columns.Count - 1
                If LdataTable.Columns(j).ColumnName IsNot Nothing Then
                    ws.Cells(1, j + 1) = LdataTable.Columns(j).ColumnName.ToString.Trim
                End If
            Next
            Dim totrows As Integer = LdataTable.Rows.Count
            Dim mFactor As Decimal = 0
            '' If LProgressBar IsNot Nothing Then
            'LProgressBar.Visible = True
            'If totrows > 0 Then
            '    mFactor = 100 / totrows
            'End If
            'End If
            For i As Integer = 0 To LdataTable.Rows.Count - 1
                'If LProgressBar IsNot Nothing Then
                '    'System.Threading.Thread.Sleep(0)
                '    Dim valPer As Decimal = ((i + 1) * mFactor / 100)
                '    Dim valStr As String = Math.Round((i + 1), 0).ToString & " / " & totrows.ToString
                '    Dim IsCancelPending As Boolean = LProgressBar.SetProgressBar(valPer, valStr)
                '    If IsCancelPending Then           '--------When user clicks on the close button of the ProgressBar form then stop the process & Exit Sub---------'
                '        LProgressBar.Visible = False
                '        GoTo EndPara
                '    End If
                'End If
                For j As Integer = 0 To LdataTable.Columns.Count - 1
                    If IsDBNull(LdataTable.Rows(i)(j)) = False Then
                        Try
                            ws.Cells(i + 2, j + 1) = LdataTable.Rows(i)(j).ToString.Trim
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next

            Next
EndPara:
            xlApp.DisplayAlerts = False
            ws.SaveAs(OutputExcelFile)
            '  ws.Delete()
            wb.Close()
            xlApp.Quit()
            'If LProgressBar IsNot Nothing Then
            '    LProgressBar.Visible = False
            '    LProgressBar.Value = 0
            'End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, Optional ByVal LeaveColumns As String = "", Optional ByVal LProgressBar As UserProgressBar.ProgressBar = Nothing)")
        End Try
    End Sub


    ''' <summary>
    ''' Create a datatable from excel worksheet -1 without using ADO.NET
    ''' </summary>  
    ''' <param name="InputExcelFile">Full identifier of excel sheet</param>
    ''' <remarks></remarks>
    Public Function GetDataTableFromExcel(ByVal InputExcelFile As String) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim ldatatable As New DataTable
        Try
            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.ApplicationClass
            xlApp.Visible = True
            Dim wb As Excel.Workbook = xlApp.Workbooks.Open(InputExcelFile)
            Dim ws As Excel.Worksheet = wb.Sheets(0)
            '   wb = xlApp.Workbooks.Add(misValue)
            '  ws = wb.Sheets("Sheet1")


            For j As Integer = 1 To ws.Columns.Count
                If ws.Cells(1, j) IsNot Nothing Then
                    If ws.Cells(1, j).ToString.Length > 0 Then
                        ldatatable.Columns.Add(ws.Cells(1, j).ToString.Trim)
                    End If
                End If
            Next
            For i = 2 To ws.Rows.Count
                Dim mrow As DataRow = ldatatable.NewRow
                For j = 1 To ldatatable.Columns.Count
                    If ws.Cells(i, j) IsNot Nothing Then
                        mrow(j - 1) = ws.Cells(i, j)
                    End If
                Next
                ldatatable.Rows.Add(mrow)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.Public Function GetDataTableFromExcel(ByVal InputExcelFile As String) As DataTable")
        End Try
        Return ldatatable
    End Function


    ''' <summary>
    ''' Create a delimited text file from a datatable
    ''' </summary>  
    ''' <param name="LdataTable">DataTable to be exported as delimited text file.</param>
    ''' <param name="OutputTextFile">Full identifier of delimited text file</param>
    ''' <param name="Delimiter" >Delimiter as character default is ","</param>
    ''' <param name="LProgressBar" >Progress bar</param>
    ''' <param name="LeaveColumns" >Comma separated columns name which are not transfered</param>
    ''' <remarks></remarks>
    Public Sub CreateDelimitedTxtFileFromDataTable(ByVal LdataTable As DataTable, ByVal OutputTextFile As String, Optional ByVal Delimiter As String = ",", Optional ByVal LeaveColumns As String = "", Optional ByRef LProgressBar As UserProgressBar.ProgressBar = Nothing)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim misValue As Object = System.Reflection.Missing.Value
        Try
            Dim aLeaveColumns() As String = {}
            If LeaveColumns.Trim.Length > 0 Then
                aLeaveColumns = Split(LeaveColumns, ",")
                For i = 0 To aLeaveColumns.Count - 1
                    Try
                        Dim mcol As String = aLeaveColumns(i)
                        LdataTable.Columns.Remove(mcol)
                    Catch ex As Exception
                        QuitError(ex, Err, "Error in removing columns")
                    End Try
                Next
            End If
            Dim totrows As Integer = LdataTable.Rows.Count
            Dim mFactor As Decimal = 0
            If LProgressBar IsNot Nothing Then
                LProgressBar.Visible = True
                If totrows > 0 Then
                    mFactor = 100 / totrows
                End If
            End If

            Dim Stm As New System.IO.StreamWriter(OutputTextFile, False)

            For i As Integer = 0 To LdataTable.Rows.Count - 1
                If LProgressBar IsNot Nothing Then
                    'System.Threading.Thread.Sleep(0)
                    Dim valPer As Decimal = ((i + 1) * mFactor / 100)
                    Dim valStr As String = Math.Round((i + 1), 0).ToString & " / " & totrows.ToString
                    Dim IsCancelPending As Boolean = LProgressBar.SetProgressBar(valPer, valStr)
                    If IsCancelPending Then           '--------When user clicks on the close button of the ProgressBar form then stop the process & Exit Sub---------'
                        LProgressBar.Visible = False
                        GoTo EndPara
                    End If
                End If
                Dim str As String = ""
                For j As Integer = 0 To LdataTable.Columns.Count - 1
                    If IsDBNull(LdataTable.Rows(i)(j)) = False Then
                        Try
                            str = str & IIf(str.Length = 0, "", Delimiter) & LdataTable.Rows(i)(j).ToString.Trim
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next
                Stm.WriteLine(str)
            Next
EndPara:
            Stm.Close()
            Stm.Dispose()
            If LProgressBar IsNot Nothing Then
                LProgressBar.Visible = False
                LProgressBar.Value = 0
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, Optional ByVal LeaveColumns As String = "", Optional ByVal LProgressBar As UserProgressBar.ProgressBar = Nothing)")
        End Try
    End Sub
    ''' <summary>
    ''' Create an excel worksheet by a datatable
    ''' </summary>
    ''' <param name="LdataTable">DataTable to be exported as work sheet</param>
    ''' <param name="OutputExcelFile">Full identifier of excel sheet</param>
    ''' <param name="DtColumnNames" >DataTable having Column names descriptions </param>
    ''' <param name="KeyField1" >Key Field name of DtColumnNames for searching of description of columnns of ldatatable</param>
    ''' <param name="DescriptionField1" >Field name of dtcolumnnames holding column descriptions</param>
    ''' <param name="DtColumnValues" >DataTable having cell value corresponding to column code</param>
    ''' <param name="KeyField2" >Comma separated key fields for dtcolumnvalues,key is combination of two values i.e columnname and cellvalue of ldatatable</param>
    ''' <param name="DescriptionField2" >field name of description of columnvalues strored in cells</param>
    ''' <param name="LeaveColumns" >An array of column names witch are not transfered to excel worksheet</param>
    ''' <remarks></remarks>


    Public Sub CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, ByVal DtColumnNames As DataTable, ByVal KeyField1 As String, ByVal DescriptionField1 As String, ByVal DtColumnValues As DataTable, ByVal KeyField2 As String, ByVal DescriptionField2 As String, Optional ByVal LeaveColumns() As String = Nothing, Optional ByRef LProgressBar As ProgressBar = Nothing)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim mErrorString As String = ""

        Try
            Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.ApplicationClass
            Dim wb As Excel.Workbook = xlApp.Workbooks.Add(misValue)
            Dim ws As Excel.Worksheet = wb.Sheets("Sheet1")
            wb = xlApp.Workbooks.Add(misValue)
            ws = wb.Sheets("Sheet1")
            If LeaveColumns IsNot Nothing Then
                For i = 0 To LeaveColumns.Count - 1
                    Try
                        Dim mcol As String = LeaveColumns(i)
                        mErrorString = "error in column " & mcol.ToString
                        If CheckColumnInDataTable(mcol, LdataTable) > 0 Then
                            LdataTable.Columns.Remove(mcol)
                        End If
                    Catch ex As Exception
                        GF1.QuitError(ex, Err, mErrorString)
                    End Try
                Next
            End If
            For j As Integer = 0 To LdataTable.Columns.Count - 1
                If LdataTable.Columns(j).ColumnName IsNot Nothing Then
                    Dim mcol As String = LdataTable.Columns(j).ColumnName.ToString.Trim
                    Dim mindx() As Integer = SearchDataTableRowIndex(DtColumnNames, KeyField1, mcol, True)
                    mErrorString = "error in ldatatable " & LdataTable.Columns(j).ColumnName
                    Try
                        If mindx.Count > 0 Then
                            If IsDBNull(DtColumnNames.Rows(mindx(0)).Item(DescriptionField1)) = False Then
                                Try
                                    ws.Cells(1, j + 1) = DtColumnNames.Rows(mindx(0)).Item(DescriptionField1).ToString.Trim
                                Catch ex As Exception
                                    Continue For
                                End Try
                            End If
                        Else
                            ws.Cells(1, j + 1) = mcol.Trim
                        End If
                    Catch ex As Exception
                        GF1.QuitError(ex, Err, mErrorString)
                    End Try
                End If
            Next
            Dim totrows As Integer = LdataTable.Rows.Count
            Dim mfactor As Integer = 0
            If LProgressBar IsNot Nothing Then
                LProgressBar.Visible = True
                If totrows > 0 Then
                    mfactor = 100 / totrows
                End If
            End If
            For i As Integer = 0 To LdataTable.Rows.Count - 1
                If LProgressBar IsNot Nothing Then
                    'System.Threading.Thread.Sleep(0)
                    Dim valPer As Decimal = ((i + 1) * mfactor / 100)
                    Dim valStr As String = Math.Round((i + 1), 0).ToString & " / " & totrows.ToString
                    Dim IsCancelPending As Boolean = LProgressBar.SetProgressBar(valPer, valStr)
                    If IsCancelPending Then           '--------When user clicks on the close button of the ProgressBar form then stop the process & Exit Sub---------'
                        LProgressBar.Visible = False
                        GoTo EndPara
                    End If
                End If

                For j As Integer = 0 To LdataTable.Columns.Count - 1
                    Try
                        If LdataTable.Columns(j).ColumnName IsNot Nothing Then
                            Dim mcol As String = LdataTable.Columns(j).ColumnName.ToString.Trim
                            Dim mcell As String = LdataTable.Rows(i).Item(mcol).ToString.Trim
                            Dim mindx() As Integer = SearchDataTableRowIndex(DtColumnValues, KeyField2, mcol & "," & mcell, True)
                            mErrorString = " error in " & LdataTable.TableName & " row " & i.ToString & " column " & j.ToString
                            If mindx.Count > 0 Then
                                If IsDBNull(DtColumnValues.Rows(mindx(0)).Item(DescriptionField2)) = False Then
                                    ws.Cells(i + 2, j + 1) = DtColumnValues.Rows(mindx(0)).Item(DescriptionField2).ToString.Trim
                                End If
                            Else
                                If IsDBNull(LdataTable.Rows(i).Item(mcol)) = False Then
                                    ws.Cells(i + 2, j + 1) = LdataTable.Rows(i).Item(mcol).ToString.Trim
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        GF1.QuitError(ex, Err, mErrorString)
                    End Try
                Next
            Next
EndPara:
            xlApp.DisplayAlerts = False
            ws.SaveAs(OutputExcelFile)
            '    ws.Delete()
            wb.Close()
            xlApp.Quit()
            If LProgressBar IsNot Nothing Then
                LProgressBar.Visible = False
                LProgressBar.Value = 0
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.CreateExcelFromDataTable(ByVal LdataTable As DataTable, ByVal OutputExcelFile As String, ByVal DtColumnNames As DataTable, ByVal KeyField1 As String, ByVal DescriptionField1 As String, ByVal DtColumnValues As DataTable, ByVal KeyField2 As String, ByVal DescriptionField2 As String, Optional ByVal LeaveColumns() As String = Nothing)")
        End Try
    End Sub
    ''' <summary>
    ''' To start an sql  transaction.
    ''' </summary>
    ''' <param name="ServerDataBase"></param>
    ''' <param name="mIsolationLevel" ></param>
    ''' <param name="MaxPoolSize"></param>
    ''' <param name="ConnectionTimeOut"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BeginTransaction(ByVal ServerDataBase As String, Optional ByVal mIsolationLevel As String = "ReadCommitted", Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlTransaction
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MyTrans As SqlTransaction = Nothing
        Try
            Dim lobj As SqlConnection = OpenSqlConnection(ServerDataBase, MaxPoolSize, ConnectionTimeOut)
            Select Case LCase(mIsolationLevel)
                Case "readcommitted"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.ReadCommitted, "SampleTransaction")
                Case "readuncommitted"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.ReadUncommitted, "SampleTransaction")
                Case "repeatableread"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.RepeatableRead, "SampleTransaction")
                Case "serializable"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.Serializable, "SampleTransaction")
                Case "snapshot"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.Snapshot, "SampleTransaction")
                Case "chaos"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.Chaos, "SampleTransaction")
                Case "unspecified"
                    MyTrans = lobj.BeginTransaction(IsolationLevel.Unspecified, "SampleTransaction")
            End Select
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.BeginTransaction(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlTransaction")
        End Try
        Return MyTrans
    End Function
    ''' <summary>
    ''' Function to commit transanction or rollback.
    ''' </summary>
    ''' <param name="Mtrans">Transanction object</param>
    ''' <param name="LastKeyPlusInTransaction" >An array of hashtable having LastKeyPlus  values involved in this transaction.</param>
    ''' <param name="TransName">Name of transanction supplied for identification</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CommitTransaction(ByRef Mtrans As SqlTransaction, Optional ByVal TransName As String = "", Optional ByRef LastKeyPlusInTransaction() As Hashtable = Nothing) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Mtrans.Commit()
            If LastKeyPlusInTransaction IsNot Nothing Then
                For i = 0 To LastKeyPlusInTransaction.Count - 1
                    LastKeyPlusInTransaction(i).Clear()
                Next
            End If
            Return True
        Catch ex As Exception
            Mtrans.Rollback()
            QuitError(ex, Err, "Transaction not committed :" & TransName)
        End Try
        Mtrans.Connection.Close()
        Mtrans.Connection.Dispose()
        Mtrans.Dispose()
        Return False
    End Function


    ''' <summary>
    ''' To start an sql transanction.
    ''' </summary>
    ''' <param name="LConnection"></param>
    ''' <param name="mIsolationLevel" ></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function BeginTransaction(ByRef LConnection As SqlConnection, Optional ByVal mIsolationLevel As String = "ReadCommitted") As SqlTransaction
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim MyTrans As SqlTransaction = Nothing
        Try
            Select Case LCase(mIsolationLevel)
                Case "readcommitted"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.ReadCommitted, "SampleTransaction")
                Case "readuncommitted"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.ReadUncommitted, "SampleTransaction")
                Case "repeatableread"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.RepeatableRead, "SampleTransaction")
                Case "serializable"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.Serializable, "SampleTransaction")
                Case "snapshot"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.Snapshot, "SampleTransaction")
                Case "chaos"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.Chaos, "SampleTransaction")
                Case "unspecified"
                    MyTrans = LConnection.BeginTransaction(IsolationLevel.Unspecified, "SampleTransaction")
            End Select
            '      MyTrans = LConnection.BeginTransaction("SampleTransaction")
            '    MyTrans = LConnection.BeginTransaction((IsolationLevel.ReadCommitted, "SampleTransaction")
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute DataFunction.BeginTransaction(ByVal ServerDataBase As String, Optional ByVal MaxPoolSize As Integer = 100, Optional ByVal ConnectionTimeOut As Integer = 30) As SqlTransaction")
        End Try
        Return MyTrans
    End Function
    ''' <summary>
    ''' Add upper node columns to a datatable,which has upper key in UpperNodeField  eg Town,District,State,Country.
    ''' </summary>
    ''' <param name="LdataTable">DataTable with upperNode field and on which columns to be added</param>
    ''' <param name="SourceNodes" >DataTable having primary keys of all uppernodes</param>
    ''' <param name="UpperNodeField">Name of field which has group value of datarow.</param>
    ''' <param name="TotalColumnsAdded">No. of columns to be added.</param>
    ''' <param name="NewColumnNames">Comma separated column Names to be added</param>
    ''' <returns>New data table with new columns and values</returns>
    ''' <remarks></remarks>
    Public Function AddUpperNodeColumns(ByVal LdataTable As DataTable, ByVal SourceNodes As DataTable, ByVal UpperNodeField As String, Optional ByRef TotalColumnsAdded As Int16 = 0, Optional ByVal NewColumnNames As String = "") As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim nodeno As Int16 = 0
        For i = 0 To LdataTable.Rows.Count - 1
            Dim mcol As Int16 = -1
            Dim munder As Integer = LdataTable(i).Item(UpperNodeField)
para:
            If munder > 0 Then
                mcol = mcol + 1
                If mcol > TotalColumnsAdded Then
                    TotalColumnsAdded = mcol
                End If
                Dim mcolumn As String = "Upper" & mcol
                LdataTable = AddColumnsInDataTable(LdataTable, mcolumn, "System.Int64")
                LdataTable.Rows(i).Item(mcolumn) = munder
                Dim mrow As DataRow = FindRowByPrimaryCols(SourceNodes, munder)
                If mrow IsNot Nothing Then
                    munder = mrow.Item(UpperNodeField)
                    GoTo para
                End If
            End If
        Next
        If NewColumnNames.Length > 0 Then
            Dim anewcols() As String = NewColumnNames.Split(",")
            For i = 0 To TotalColumnsAdded
                If i <= anewcols.Count - 1 Then
                    RenameDataTableColumn(LdataTable, "upper" & i.ToString, anewcols(i))
                End If
            Next
        End If
        Return LdataTable
    End Function
    ''' <summary>
    ''' Create a treeview object from a datatable having a hierarchical strucure fields,eg. ParentField,ChildField,NodeTextField and other fields attatched in tag.
    ''' </summary>
    ''' <param name="LdataTable">Data table from which tree created</param>
    ''' <param name="ChildField">Primary key field of datatable </param>
    ''' <param name="ParentField">Field name contains the parent key field of datarow</param>
    ''' <param name="NodeTextField">Text field name which is shown as node.text</param>
    ''' <param name="OtherDisplayfields">Other Fields value stored in treeview.tag</param>
    ''' <param name="TreeViewObject" >A TreeView Object sent by reference with other initial property settings.</param>
    ''' <returns>Tree View object</returns>
    ''' <remarks></remarks>
    Public Function CreateTreeView(ByVal LdataTable As DataTable, ByVal ChildField As String, ByVal ParentField As String, ByVal NodeTextField As String, Optional ByRef TreeViewObject As System.Windows.Forms.TreeView = Nothing, Optional ByVal OtherDisplayfields As String = "") As System.Windows.Forms.TreeView
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim TreeViewName As New System.Windows.Forms.TreeView
        If TreeViewObject IsNot Nothing Then
            TreeViewName = TreeViewObject
            TreeViewName.Nodes.Clear()
        End If
        Try
            Dim nodeParent, nodeChild As New System.Windows.Forms.TreeNode
            Dim rowParent, rowChild As DataRow
            Dim Ldatatablesort As DataTable = SortDataTable(LdataTable, ParentField)
            Dim ParentData As DataTable = SearchInSortedDataTable(Ldatatablesort, ParentField, 0, NodeTextField)
            For Each rowParent In ParentData.Rows
                nodeParent = TreeViewName.Nodes.Add(rowParent(NodeTextField))
                nodeParent.Name = rowParent(ChildField)
                nodeParent.Tag = rowParent
                Dim ChildData As DataTable = SearchInSortedDataTable(Ldatatablesort, ParentField, CInt(rowParent(ChildField)), NodeTextField)
                For Each rowChild In ChildData.Rows
                    If rowChild(ParentField) = rowParent(ChildField) Then
                        nodeChild = nodeParent.Nodes.Add(rowChild(NodeTextField))
                        nodeChild.Name = rowChild(ChildField)
                        nodeChild.Tag = rowChild
                        NodeFill(Ldatatablesort, rowChild(ChildField), nodeChild, ChildField, ParentField, NodeTextField)
                    End If
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunctions.CreateTreeView(ByVal LdataTable As DataTable, ByVal cFieldName As String, ByVal pFieldName As String, ByVal NodeFieldName As String, ByVal propline As String, Optional ByVal OtherDisplayfields As String = "") As System.Windows.Forms.TreeView")
        End Try
        Return TreeViewName
    End Function
    Public Function CreateTreeView_old(ByVal LdataTable As DataTable, ByVal ChildField As String, ByVal ParentField As String, ByVal NodeTextField As String, Optional ByRef TreeViewObject As System.Windows.Forms.TreeView = Nothing, Optional ByVal OtherDisplayfields As String = "") As System.Windows.Forms.TreeView
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim TreeViewName As New System.Windows.Forms.TreeView
        If TreeViewObject IsNot Nothing Then
            TreeViewName = TreeViewObject
            TreeViewName.Nodes.Clear()
        End If
        Try
            Dim nodeParent, nodeChild As New System.Windows.Forms.TreeNode
            Dim rowParent, rowChild As DataRow
            Dim Ldatatablesort As DataTable = SortDataTable(LdataTable, ParentField)
            Dim ParentData As DataTable = SearchInSortedDataTable(Ldatatablesort, ParentField, 0, NodeTextField)
            For Each rowParent In ParentData.Rows
                nodeParent = TreeViewName.Nodes.Add(rowParent(NodeTextField))
                nodeParent.Name = rowParent(ChildField)
                nodeParent.Tag = rowParent
                Dim ChildData As DataTable = SearchInSortedDataTable(Ldatatablesort, ParentField, CInt(rowParent(ParentField)))
                For Each rowChild In ChildData.Rows
                    If rowChild(ParentField) = rowParent(ChildField) Then
                        nodeChild = nodeParent.Nodes.Add(rowChild(NodeTextField))
                        nodeChild.Name = rowChild(ChildField)
                        nodeChild.Tag = rowChild
                        NodeFill(Ldatatablesort, CInt(rowChild(ChildField)), nodeChild, ChildField, ParentField, NodeTextField)
                    End If
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunctions.CreateTreeView(ByVal LdataTable As DataTable, ByVal cFieldName As String, ByVal pFieldName As String, ByVal NodeFieldName As String, ByVal propline As String, Optional ByVal OtherDisplayfields As String = "") As System.Windows.Forms.TreeView")
        End Try
        Return TreeViewName
    End Function

    Private Sub NodeFill(ByVal dt As DataTable, ByVal tt As Integer, ByVal oNode As System.Windows.Forms.TreeNode, ByVal ChildField As String, ByVal ParentField As String, ByVal NodeFieldName As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim ChildData As DataTable = SearchInSortedDataTable(dt, ParentField, tt, NodeFieldName)
            For Each dr In ChildData.Rows
                Dim NewNode As New System.Windows.Forms.TreeNode(dr(NodeFieldName))
                NewNode.Name = dr(ChildField)
                NewNode.Tag = dr
                oNode.Nodes.Add(NewNode)
                NodeFill(dt, CInt(dr(ChildField)), NewNode, ChildField, ParentField, NodeFieldName)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute datafunctions.NodeFill(ByVal dt As DataTable, ByVal tt As String, ByVal oNode As System.Windows.Forms.TreeNode, ByVal cFieldName As String, ByVal pFieldName As String, ByVal NodeFieldName As String")

        End Try
    End Sub
    ''' <summary>
    ''' Get associated Tag details  of a TreeNode object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetTreeNodeDetail(ByVal TreeNodeObject As System.Windows.Forms.TreeNode) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim str As String = ""
        Dim row As DataRow = CType(TreeNodeObject.Tag, DataRow)
        For i = 2 To row.ItemArray.Count - 1
            str = str & vbNewLine
            str = str & row.Table.Columns(i).ColumnName.ToString & " :  " & row.ItemArray(i).ToString & vbNewLine & vbNewLine
        Next
        Return str
    End Function
    ''' <summary>
    ''' Check wether a datarow has all fields empty.
    ''' </summary>
    ''' <param name="LDataRow">DataRow to be checked</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsDataRowEmpty(ByVal LDataRow As DataRow) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim flag As Boolean = True
        If LDataRow Is Nothing Then
            Return flag
            Exit Function
        End If
        For i = 0 To LDataRow.ItemArray.Count - 1
            If IsDBNull(LDataRow(i)) = False Then
                flag = False
                Exit For
            End If
        Next
        Return flag
    End Function
    ''' <summary>
    '''  Insert rows in CurrRowArray in table class object for new country,state,district,town for later sql stmt. execution.
    ''' 
    ''' </summary>
    ''' <param name="MyTrans">Sql Transanction</param>
    ''' <param name="ClsTownTable">Class of info table</param>
    ''' <param name="HashTownDetails">Hash table of inserting values keys are Country,CountryKey,State,StateKey,District,DistrictKey,HomwTown,HomeTownKey</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InsertForNewTownInInfoTable(ByRef MyTrans As SqlTransaction, ByRef ClsTownTable As Object, ByVal HashTownDetails As Hashtable) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mHomeTownKey As Integer = 0
        Try
            If HashTownDetails IsNot Nothing Then
                If GF1.GetValueFromHashTable(HashTownDetails, "NewRowFlag") = False Then
                    mHomeTownKey = GF1.GetValueFromHashTable(HashTownDetails, "HomeTownKey")
                    Return mHomeTownKey
                    Exit Function
                End If
                ClsTownTable.SqlUpdation = True
                Dim country As String = GF1.GetValueFromHashTable(HashTownDetails, "Country")
                Dim mKeyPlusGroups As String = IIf(ClsTownTable.RowStatusFlag = True, "Y,R,O,D", "Y,R,O,D")
                If country IsNot Nothing Then
                    If country.Trim.Length > 0 Then
                        ClsTownTable = LastKeysPlus(MyTrans, ClsTownTable, mKeyPlusGroups)
                        Dim mrow As DataRow = ClsTownTable.NewRow
                        mrow("InfoTable_Key") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "InfoTable_Key")
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable", Nothing) IsNot Nothing Then
                            mrow("P_InfoTable") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable")
                        Else
                            mrow("P_InfoTable") = mrow("InfoTable_Key")
                        End If
                        ' MsgBox(ClsTownTable.fieldsfinalvalues.count)
                        ' changes by Neha

                        If Not GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus", "Nothing") Is Nothing Then
                            mrow("RowStatus") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus")
                        Else
                            mrow("RowStatus") = 0
                        End If

                        mrow("InfoType") = "C"
                        mrow("NameOfInfo") = country
                        mrow("Under") = 0
                        mrow("InfoDes") = ""
                        mrow("GeneratedBy") = "U"
                        mrow("Verified") = "N"
                        GF1.ArrayAppend(ClsTownTable.CurrRowsArray, mrow)
                        HashTownDetails = GF1.AddItemToHashTable(HashTownDetails, "CountryKey", mrow("P_InfoTable"))
                    End If
                End If
                Dim state As String = GF1.GetValueFromHashTable(HashTownDetails, "state")
                If state IsNot Nothing Then
                    If state.Trim.Length > 0 Then
                        ClsTownTable = LastKeysPlus(MyTrans, ClsTownTable, mKeyPlusGroups)
                        Dim mrow As DataRow = ClsTownTable.NewRow
                        mrow("InfoTable_Key") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "InfoTable_Key")
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable") IsNot Nothing Then
                            mrow("P_InfoTable") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable")
                        Else
                            mrow("P_InfoTable") = mrow("InfoTable_Key")
                        End If
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus", "Nothing") IsNot Nothing Then
                            mrow("RowStatus") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus")
                        Else
                            mrow("RowStatus") = 0
                        End If

                        mrow("InfoType") = "S"
                        mrow("NameOfInfo") = state
                        mrow("Under") = GF1.GetValueFromHashTable(HashTownDetails, "CountryKey")
                        mrow("InfoDes") = ""
                        mrow("GeneratedBy") = "U"
                        mrow("Verified") = "N"
                        GF1.ArrayAppend(ClsTownTable.CurrRowsArray, mrow)
                        HashTownDetails = GF1.AddItemToHashTable(HashTownDetails, "StateKey", mrow("P_InfoTable"))
                    End If
                End If
                Dim District As String = GF1.GetValueFromHashTable(HashTownDetails, "District")
                If District IsNot Nothing Then
                    If District.Trim.Length > 0 Then
                        ClsTownTable = LastKeysPlus(MyTrans, ClsTownTable, mKeyPlusGroups)
                        Dim mrow As DataRow = ClsTownTable.NewRow
                        mrow("InfoTable_Key") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "InfoTable_Key")
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable") IsNot Nothing Then
                            mrow("P_InfoTable") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable")
                        Else
                            mrow("P_InfoTable") = mrow("InfoTable_Key")
                        End If
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus", "Nothing") IsNot Nothing Then
                            mrow("RowStatus") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus")
                        Else
                            mrow("RowStatus") = 0
                        End If
                        mrow("InfoType") = "D"
                        mrow("NameOfInfo") = District
                        mrow("Under") = GF1.GetValueFromHashTable(HashTownDetails, "StateKey")
                        mrow("InfoDes") = ""
                        mrow("GeneratedBy") = "U"
                        mrow("Verified") = "N"
                        GF1.ArrayAppend(ClsTownTable.CurrRowsArray, mrow)
                        HashTownDetails = GF1.AddItemToHashTable(HashTownDetails, "DistrictKey", mrow("P_InfoTable"))
                    End If
                End If
                Dim HomeTown As String = GF1.GetValueFromHashTable(HashTownDetails, "HomeTown")
                If HomeTown IsNot Nothing Then
                    If HomeTown.Trim.Length > 0 Then
                        ClsTownTable = LastKeysPlus(MyTrans, ClsTownTable, mKeyPlusGroups)
                        Dim mrow As DataRow = ClsTownTable.NewRow
                        mrow("InfoTable_Key") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "InfoTable_Key")
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable") IsNot Nothing Then
                            mrow("P_InfoTable") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "P_InfoTable")
                        Else
                            mrow("P_InfoTable") = mrow("InfoTable_Key")
                        End If
                        If GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus", "Nothing") IsNot Nothing Then
                            mrow("RowStatus") = GF1.GetValueFromHashTable(ClsTownTable.FieldsFinalValues, "RowStatus")
                        Else
                            mrow("RowStatus") = 0
                        End If
                        mrow("InfoType") = "T"
                        mrow("NameOfInfo") = HomeTown
                        mrow("Under") = GF1.GetValueFromHashTable(HashTownDetails, "DistrictKey")
                        mrow("InfoDes") = ""
                        mrow("GeneratedBy") = "U"
                        mrow("Verified") = "N"
                        GF1.ArrayAppend(ClsTownTable.CurrRowsArray, mrow)
                        HashTownDetails = GF1.AddItemToHashTable(HashTownDetails, "HomeTownKey", mrow("P_InfoTable"))
                    End If
                    mHomeTownKey = GF1.GetValueFromHashTable(HashTownDetails, "HomeTownKey")
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute  for infotable code town   InsertForNewTownInInfoTable(ByRef MyTrans As SqlTransaction, ByRef ClsTownTable As Object, ByVal HashTownDetails As Hashtable, Optional ByVal mLastLoginHits As Integer = 0) As Integer")
        End Try

        Return mHomeTownKey
    End Function
    ''' <summary>
    '''  Insert rows in CurrRowArray in table class object if new record is selected by the user for later sql stmt. execution.
    ''' </summary>
    ''' <param name="MyTrans">Sql Transanction </param>
    ''' <param name="ClsInfoTable">Class of table</param>
    ''' <param name="HashInfoDetails">Hash table of inserting values with extra keys are  "NewRowFlag" =true,and "KeyValue" as primary key value  </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetCurrRowArrayIfNewInsert(ByRef MyTrans As SqlTransaction, ByRef ClsInfoTable As Object, ByVal HashInfoDetails As Hashtable, ByVal mKeyField As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mInfoKey As Integer = 0
        Try
            If HashInfoDetails.Count > 0 Then
                If GF1.CheckIfKeyNameExistsinHashTable(HashInfoDetails, "NewRowFlag") = False Then Exit Function
                If GF1.GetValueFromHashTable(HashInfoDetails, "NewRowFlag") = False Then

                    Dim mInfoKey1 As Integer = GF1.GetValueFromHashTable(HashInfoDetails, mKeyField)
                    If mInfoKey1 > -1 Then
                        Return mInfoKey1
                        Exit Function
                    End If
                End If
                ' Dim mOthers As String = GF1.GetValueFromHashTable(HashInfoDetails, "NameOfInfo")
                'If mOthers IsNot Nothing Then
                'If mOthers.Trim.Length > 0 Then
                Dim mKeyPlusGroups As String = IIf(ClsInfoTable.RowStatusFlag = True, "Y,R,O,D", "Y,R,O,D")
                ClsInfoTable.SqlUpdation = True
                ClsInfoTable = LastKeysPlus(MyTrans, ClsInfoTable, mKeyPlusGroups)
                Dim mrow As DataRow = ClsInfoTable.NewRow
                For i = 0 To HashInfoDetails.Count - 1
                    Dim mkey As String = HashInfoDetails.Keys(i)
                    Dim mvalue As Object = HashInfoDetails.Item(mkey)
                    If CheckColumnInDataTable(mkey, ClsInfoTable.CurrDt) > -1 Then
                        If mvalue IsNot Nothing Then
                            mrow.Item(mkey) = mvalue
                        End If
                    End If
                Next
                For i = 0 To ClsInfoTable.FieldsFinalValues.Count - 1
                    Dim mFieldsFinalValues As Hashtable = ClsInfoTable.FieldsFinalValues
                    Dim mkey As String = mFieldsFinalValues.Keys(i)
                    Dim mvalue As Object = mFieldsFinalValues.Item(mkey)
                    If mvalue IsNot Nothing Then
                        mrow.Item(mkey) = mvalue
                    End If
                Next
                GF1.ArrayAppend(ClsInfoTable.CurrRowsArray, mrow)
                mInfoKey = mrow.Item(ClsInfoTable.PrimaryKey)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute   datafunctions.SetCurrRowArrayIfNewInsert(ByRef MyTrans As SqlTransaction, ByRef ClsInfoTable As Object, ByVal HashInfoDetails As Hashtable,mKeyField as string) As Integer")

        End Try
        Return mInfoKey
    End Function
    ''' <summary>
    '''  Insert rows in CurrRowArray in table class object if new record is selected by the user for later sql stmt. execution.
    ''' </summary>
    ''' <param name="MyTrans">Sql Transanction </param>
    ''' <param name="ClsInfoTable">Class of table</param>
    ''' <param name="HashInfoDetails">Hash table of inserting values with extra keys are  "NewRowFlag" =true,and "KeyValue" as primary key value  </param>
    ''' <param name="mKeyField" >Key name  of hashtable ,whose value assigned to maintable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SetCurrRowArrayIfNewUchkInsert(ByRef MyTrans As SqlTransaction, ByRef ClsInfoTable As Object, ByVal HashInfoDetails As Hashtable, ByVal mKeyField As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mInfoKey1 As String = ""
        Try
            If HashInfoDetails.Count > 0 Then
                If GF1.GetValueFromHashTable(HashInfoDetails, "NewRowFlag") = False Then
                    mInfoKey1 = GF1.GetValueFromHashTable(HashInfoDetails, mKeyField)
                    ' If mInfoKey1.Trim.Length > 0 Then
                    Return mInfoKey1
                    Exit Function
                    'End If
                End If
                ClsInfoTable.SqlUpdation = True
                Dim mKeyPlusGroups As String = IIf(ClsInfoTable.RowStatusFlag = True, "Y,R,O,D", "Y,R,O,D")
                ClsInfoTable = LastKeysPlus(MyTrans, ClsInfoTable, mKeyPlusGroups)
                Dim mrow As DataRow = ClsInfoTable.NewRow
                For i = 0 To HashInfoDetails.Count - 1
                    Dim mkey As String = HashInfoDetails.Keys(i)
                    Dim mvalue As Object = HashInfoDetails.Item(mkey)
                    If CheckColumnInDataTable(mkey, ClsInfoTable.CurrDt) > -1 Then
                        If mvalue IsNot Nothing Then
                            mrow.Item(mkey) = mvalue
                        End If
                    End If
                Next
                For i = 0 To ClsInfoTable.FieldsFinalValues.Count - 1
                    Dim mFieldsFinalValues As Hashtable = ClsInfoTable.FieldsFinalValues
                    Dim mkey As String = mFieldsFinalValues.Keys(i)
                    Dim mvalue As Object = ClsInfoTable.FieldsFinalValues.Item(mkey)
                    If mvalue IsNot Nothing Then
                        mrow.Item(mkey) = mvalue
                    End If
                Next
                GF1.ArrayAppend(ClsInfoTable.CurrRowsArray, mrow)
                Dim mInfoKey As Integer = mrow.Item(ClsInfoTable.PrimaryKey)
                If GF1.GetValueFromHashTable(HashInfoDetails, mKeyField) IsNot Nothing Then
                    mInfoKey1 = GF1.GetValueFromHashTable(HashInfoDetails, mKeyField)
                End If
                mInfoKey1 = mInfoKey1 & IIf(mInfoKey1.Length = 0, "", ",") & mInfoKey.ToString
                HashInfoDetails = GF1.AddItemToHashTable(HashInfoDetails, mKeyField, mInfoKey1)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute   datafunctions.SetCurrRowArrayIfNewInsert(ByRef MyTrans As SqlTransaction, ByRef ClsInfoTable As Object, ByVal HashInfoDetails As Hashtable,mKeyField as string) As Integer")
        End Try
        Return mInfoKey1
    End Function
    ''' <summary>
    '''  Insert rows in CurrRowArray in table class object if new record is selected by the user for later sql stmt. execution.
    ''' </summary>
    ''' <param name="MyTrans">Sql Transanction </param>
    ''' <param name="ClsTables">Class of table</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetKeyValueIfNewInsert(ByRef MyTrans As SqlTransaction, ByRef ClsTables() As Object) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        For i = 0 To ClsTables.Count - 1
            Try
                Dim HGroupCode As String = GF1.GetValueFromHashTable(ClsTables(i).GroupFieldsType, "H")
                If HGroupCode = "" Then  'changed by Neha
                    Continue For
                End If
                Dim aHgroup() As String = HGroupCode.Split("|")
                For j = 0 To aHgroup.Count - 1
                    Dim afield() As String = aHgroup(j).Split("#")
                    Dim mFiled As String = afield(0)
                    Dim msubtable As String = afield(1)
                    Dim minfocode As String = ""
                    Dim mKeyField As String = ""

                    Dim mCntrType As String = ""
                    If afield.Count > 2 Then
                        minfocode = afield(2)
                        mCntrType = UCase(afield(3))
                        mKeyField = afield(8)
                    End If
                    Dim xx As Integer = -1
                    For x = 0 To ClsTables.Count - 1
                        If LCase(ClsTables(x).TableName) = LCase(msubtable) Then
                            xx = x
                            Exit For
                        End If
                    Next
                    If xx > -1 Then
                        If ClsTables(i).MultyRowsSqlHandling = True Then
                            For y = 0 To ClsTables(i).CurrDt.Rows.Count - 1
                                If IsDBNull(ClsTables(i).CurrDt.Rows(y).Item("Hash" & mFiled)) = False Then
                                    Dim hashvalues As Hashtable = ClsTables(i).CurrDt.Rows(y).Item("Hash" & mFiled)
                                    Select Case mCntrType
                                        Case "ULOC"
                                            ClsTables(i).CurrDt.Rows(y).Item(mFiled) = InsertForNewTownInInfoTable(MyTrans, ClsTables(xx), hashvalues)
                                        Case "UCHK"
                                            ClsTables(i).CurrDt.Rows(y).Item(mFiled) = SetCurrRowArrayIfNewUchkInsert(MyTrans, ClsTables(xx), hashvalues, mKeyField)
                                        Case Else
                                            ClsTables(i).CurrDt.Rows(y).Item(mFiled) = SetCurrRowArrayIfNewInsert(MyTrans, ClsTables(xx), hashvalues, mKeyField)
                                    End Select
                                End If
                            Next
                        Else
                            If IsDBNull(ClsTables(i).CurrRow.Item("Hash" & mFiled)) = False Then
                                Dim hashvalues As Hashtable = ClsTables(i).CurrRow.Item("Hash" & mFiled)
                                Select Case mCntrType
                                    Case "ULOC"
                                        ClsTables(i).CurrRow.Item(mFiled) = InsertForNewTownInInfoTable(MyTrans, ClsTables(xx), hashvalues)
                                    Case "UCHK"
                                        ClsTables(i).CurrRow.Item(mFiled) = SetCurrRowArrayIfNewUchkInsert(MyTrans, ClsTables(xx), hashvalues, mKeyField)
                                    Case Else
                                        ClsTables(i).CurrRow.Item(mFiled) = SetCurrRowArrayIfNewInsert(MyTrans, ClsTables(xx), hashvalues, mKeyField)
                                End Select
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute  for table " & ClsTables(i).TableName & "  datafunctions.SetKeyValueIfNewInsert(ByRef MyTrans As SqlTransaction, ByRef ClsTables() As Object) As Object()")
            End Try

        Next
        Return ClsTables
    End Function

    ''' <summary>
    ''' Convert fieldvalue to formatted value in diffrent FRMT prefixed fieldname.
    ''' </summary>
    ''' <param name="ClsObject"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormattedDisplayValue(ByRef ClsObject As Object) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
        Dim mvalue As String = GF1.GetValueFromHashTable(mGroupFieldsType, "F")
        If mvalue Is Nothing Then
            Return ClsObject
        End If
        Dim afields() As String = mvalue.Split("|")
        If ClsObject.MultyRowsSqlHandling = True Then
            For i = 0 To ClsObject.CurrDt.Rows.Count - 1
                ' GF1.AddItemToHashTable(mtemp, "F", "JobFrom#D=dd/MM/yyyy|JobTo#D=dd/MM/yyyy")
                For j = 0 To afields.Count - 1
                    Dim xField() As String = afields(j).Split("#")
                    Dim mvalue0 As Object = ClsObject.CurrDt.Rows(i).item(xField(0))
                    If IsDBNull(mvalue0) = False Then
                        Dim mField As String = "Frmt" & xField(0)
                        Dim mFormat As String = xField(1)
                        ClsObject.CurrDt.Rows(i).item(mField) = GF1.FormattedValue(ClsObject.CurrDt.Rows(i).item(xField(0)), mFormat)
                    End If
                Next
            Next
        Else
            For j = 0 To afields.Count - 1
                Dim xField() As String = afields(j).Split("#")
                Dim mField As String = "Frmt" & xField(0)
                Dim mFormat As String = xField(1)
                ClsObject.CurrRow.Item(mField) = GF1.FormattedValue(ClsObject.CurrRow.Item(xField(0)), mFormat)
            Next
        End If
        Return ClsObject
    End Function
    ''' <summary>
    ''' Validate  fieldvalues of a table class.
    ''' </summary>
    ''' <param name="ClsObject"></param>
    ''' <param name="HashUnloadValues" >A hashtable having keys as CondionVariable names and value is variablevalue </param>
    ''' <returns>Return controlname to be focussed</returns>
    ''' <remarks></remarks>
    Public Function ValidateFields(ByRef ClsObject As Object, ByVal HashUnloadValues As Hashtable, Optional ByRef MessageString As String = "", Optional ByVal ApplicationType As String = "VB") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mGroupFieldsType As Hashtable = ClsObject.GroupFieldsType
        Dim mvalue As String = GF1.GetValueFromHashTable(mGroupFieldsType, "V")
        Dim mName As String = ""
        If mvalue Is Nothing Then
            Return mName
        End If
        Dim afields() As String = mvalue.Split("|")
        MessageString = ""
        For j = 0 To afields.Count - 1
            Dim xField() As String = afields(j).Split("#")
            Dim mField As String = xField(0)
            Dim mCntrType As String = xField(1)
            Dim mValueProperty As String = xField(2)
            Dim mValueType As String = xField(3)
            Dim mHashKey As String = xField(4)
            Dim mHashKeyType As String = LCase(xField(5))
            Dim mCondition As String = xField(6)
            Dim mmessage As String = xField(7)
            Dim aVar() As String = GF1.ExtractVariables(mCondition)
            Dim aValueHash As New Hashtable
            For z = 0 To aVar.Count - 1
                Dim ivalue As Object = GF1.GetValueFromHashTable(HashUnloadValues, mField)
                If ivalue Is Nothing Then
                    QuitMessage("Field Value of key " & mField & " not defined in hashtable", "Public Function ValidateFields(ByRef ClsObject As Object, ByVal ValidationVariables As Hashtable) As String")
                End If
                If mValueType = "hash" Then
                    If GF1.GetValueFromHashTable(ivalue, mHashKey) Is Nothing Then
                        Select Case mHashKeyType
                            Case "int", "dec"
                                ivalue = 0
                            Case "date"
                                ivalue = #1/1/1900#
                            Case "string"
                                ivalue = ""
                        End Select
                    Else
                        ivalue = GF1.GetValueFromHashTable(ivalue, mHashKey)
                    End If
                End If
                aValueHash = GF1.AddItemToHashTable(aValueHash, mField, ivalue)
            Next
            Dim rexpr As String = GF1.ReplaceValuesInExpression(mCondition, aValueHash, "VB")
            Dim merror As Boolean = False
            Dim mflag As Boolean = GF1.EvaluateExpression(rexpr, merror)
            If mflag = True Then
                MessageString = MessageString & IIf(MessageString.Length = 0, "", vbCrLf) & mmessage
                If mName.Length = 0 Then
                    mName = mCntrType & mField
                End If
            End If
        Next
        If ApplicationType = "VB" Then
            If mName.Length > 0 Then
                MsgBox(MessageString)
            End If
        End If
        Return mName
    End Function
    ''' <summary>
    ''' Reverse Rows of a DataTable. i.e. LastRow becomes FirstRow.
    ''' </summary>
    ''' <param name="mDataTabale">DataTable to be reversed</param>
    ''' <param name="Row_IdColumn">True if datatable contains Row_Id column.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReverseDataTable(ByVal mDataTabale As DataTable, Optional ByVal Row_IdColumn As Boolean = True) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim newdt As DataTable = mDataTabale.Clone
        Dim i As Integer = 0
        For ll = mDataTabale.Rows.Count - 1 To 0 Step -1
            Dim mRow As DataRow = newdt.NewRow
            mRow.ItemArray = mDataTabale.Rows(ll).ItemArray
            If Row_IdColumn = True Then
                mRow.Item("Row_Id") = i
                i = i + 1
            End If
            newdt.Rows.Add(mRow)
        Next
        Return newdt
    End Function
    ''' <summary>
    ''' To create an sql filter querry by tables FilterDetails,FilterSelected
    ''' </summary>
    ''' <param name="FilterDetails"></param>
    ''' <param name="FilterSelected"></param>
    ''' <param name="ParentDetailKey"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetUserFilterClause(ByVal FilterDetails As DataTable, ByVal FilterSelected As DataTable, ByVal ParentDetailKey As Integer) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If FilterDetails.Rows.Count = 0 Then
            Return ""
            Exit Function
        End If
        Dim FilterString As String = ""
        Dim DialogueStr As String = ""
        Dim GridStr As String = ""
        'Dim InsertSquareBracket As String = ""
        Dim FurtherNestedGridStr As String = ""
        Dim TempStr As String = ""
        Dim MultiSelectAlias As String = ""
        Dim SubTableCount As Int16 = 0
        Try
            Dim BaseTableFilter As DataTable = SearchDataTable(FilterDetails, "ParentFilterDetailsKey,DefineFlag", ParentDetailKey.ToString & ",True")
            Dim GridContionTable As DataTable = SearchDataTable(BaseTableFilter, "ConditionType", "G") '----get rows which contain gridColumns
            Dim DialogueConditionTable As DataTable = SearchDataTable(BaseTableFilter, "ConditionType = 'r'") '---get rows which contains dialogue columns
            Dim TableName As String = GridContionTable(0)("MainTable").ToString.Trim
            Dim KeyId As String = GridContionTable(0)("MainTable_Key")
            If DialogueConditionTable.Rows.Count > 0 Then
                DialogueStr = GetDialogueString(BaseTableFilter, "m1", FilterSelected) '---Create filter string for all the dialogue rows
            End If
            Dim MultiSelectedConditionTable As DataTable = SearchDataTable(GridContionTable, "FilterSelectedFlag", True) '----Get that grid column for which multiple selection is done
            If MultiSelectedConditionTable.Rows.Count > 0 Then
                '------Get only those rows which are slected by multiple Selection because when multiple selection is done all other conditions doesnt have any importance so not needed.
                SubTableCount += 1
                MultiSelectAlias = "s" & SubTableCount
                Dim MultiSelectedTableName As String = TableName & "MultiSelected"
                FilterString = "Select * from " & TableName & " m1 inner join " & MultiSelectedTableName & " " & MultiSelectAlias & " on m1." & KeyId & "=" & MultiSelectAlias & ".FilterSelected_Key"
                Return FilterString
                Exit Function
            Else
                '-----wehen multiple selection is not done on any table then get filter string for all other conditions
                FilterString = "Select * from " & TableName & " m1 where " & DialogueStr
            End If
            If GridContionTable.Rows.Count > 0 Then '----to get Grid Condition in filterString
                'InsertSquareBracket = "["
                For rr = 0 To GridContionTable.Rows.Count - 1
                    Dim ParentKey As Integer = GridContionTable(rr)("FilterDetails_Key")
                    Dim gridtable As DataTable = SearchDataTable(FilterDetails, "ParentFilterDetailsKey,DefineFlag", ParentKey & ",True") '---Check if further Filteration exit for the gridColumn
                    If gridtable.Rows.Count = 0 Then
                        Return FilterString
                        Exit Function
                    Else '-----Get Futher Filterartion Condition for the GridColumn if Exist.
                        GridStr = GetGridString(GridContionTable.Rows(rr), "m1")
                        FurtherNestedGridStr = GetSubGridString(FilterDetails, gridtable, FilterSelected, SubTableCount)
                        'FurtherNestedGridStr = GetSubGridString(FilterDetails, FilterSelected, ParentKey)
                        FurtherNestedGridStr = FurtherNestedGridStr & ")"
                        TempStr = IIf(TempStr = "", GridStr & FurtherNestedGridStr, TempStr & GridStr & FurtherNestedGridStr)
                    End If
                Next
                FilterString = FilterString & TempStr
                'FilterString = FilterString & InsertSquareBracket & TempStr
                'FilterString = FilterString & " ]"
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Public Function GetFilterQuery(ByVal FilterDetails As DataTable, ByVal FilterSelected As DataTable, ByVal ParentDetailKey As Integer) As String")
        End Try
        Return FilterString
    End Function
    Private Function GetDialogueString(ByVal BaseTableFilter As DataTable, ByVal TableNameAlias As String, ByVal FilterSelected As DataTable) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mFinalStatus As String = ""
            Dim GridContionTable As DataTable = SearchDataTable(BaseTableFilter, "ConditionType,DefineFlag", "G,True")
            For Each rw As DataRow In BaseTableFilter.Rows
                Dim CondType As String = rw("ConditionType")
                Select Case LCase(CondType)
                    Case "r"
                        Dim optr As String = rw("LogicalOperator")
                        'Dim ConditionColumnName As String = rw("ConditionColumn")
                        'Dim cond As String = Replace(rw("Condition"), rw("ConditionColumn"), TableNameAlias & "." & rw("ConditionColumn"))
                        Dim cond As String = CreateConditionString(rw, TableNameAlias, FilterSelected)
                        mFinalStatus = mFinalStatus & IIf(mFinalStatus = "", cond, " " & optr & " " & cond)
                End Select
            Next
            Return mFinalStatus
        Catch ex As Exception
            QuitError(ex, Err, "Private Function GetDialogueString(ByVal BaseTableFilter As DataTable, ByVal TableNameAlias As String, ByVal FilterSelected As DataTable) As String")
            Return Nothing
        End Try
    End Function
    Private Function GetGridString(ByVal GridRow As DataRow, ByVal TableNameAlias As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mFinalStatus As String = ""
            Dim tableName As String = GridRow("MainTable")
            Dim KeyId As String = GridRow("MainTable_Key")
            Dim ColName As String = GridRow("ConditionColumn")
            'mFinalStatus = "AND " & tableName & "." & ColName & " In "
            mFinalStatus = " AND " & TableNameAlias & "." & ColName & " In "
            Return mFinalStatus
        Catch ex As Exception
            QuitError(ex, Err, "Private Function GetGridString(ByVal GridRow As DataRow, ByVal TableNameAlias As String) As String")
            Return Nothing
        End Try
    End Function
    Private Function GetSubGridString(ByVal FilterDetails As DataTable, ByVal GridTable As DataTable, ByVal FilterSelected As DataTable, ByRef SubTableCount As Int16) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            SubTableCount += 1
            Dim TableAlias As String = "s" & SubTableCount
            Dim mFinalStatus As String = ""
            Dim FurtherGridStr As String = ""
            'Dim InsertSquareBracket As String = ""
            Dim TempStr As String = ""
            Dim GridContionTable As DataTable = SearchDataTable(GridTable, "ConditionType", "G") '----get rows which contain gridColumns
            Dim DialogueConditionTable As DataTable = SearchDataTable(GridTable, "ConditionType = 'r'") '---get rows which contains dialogue columns
            Dim DialogueStr As String = ""
            Dim GridStr As String = ""
            If DialogueConditionTable.Rows.Count > 0 Then
                DialogueStr = GetDialogueString(DialogueConditionTable, TableAlias, FilterSelected) '---Create filter string for all the dialogue rows
            End If
            Dim tableName As String = GridTable(0)("MainTable")
            Dim KeyId As String = GridTable(0)("MainTable_Key")
            Dim MultiSelectedConditionTable As DataTable = SearchDataTable(GridContionTable, "FilterSelectedFlag", True) '----Get that grid column for which multiple selection is done
            If MultiSelectedConditionTable.Rows.Count > 0 Then
                '------Get only those rows which are slected by multiple Selection because when multiple selection is done all other conditions doesnt have any importance so not needed.
                SubTableCount += 1
                Dim MultiSelectAlias As String = "s" & SubTableCount
                Dim MultiSelectedTableName As String = tableName & "MultiSelected"
                mFinalStatus = "(Select " & KeyId & " from " & tableName & " " & TableAlias & " inner join " & MultiSelectedTableName & " " & MultiSelectAlias & " on " & TableAlias & "." & KeyId & "=" & MultiSelectAlias & ".FilterSelected_Key"
                Return mFinalStatus
                Exit Function
            Else
                '-----wehen multiple selection is not done on any table then get filter string for all other conditions
                mFinalStatus = "(Select " & KeyId & " from " & tableName & " " & TableAlias & " where " & DialogueStr
            End If
            If GridContionTable.Rows.Count > 0 Then '----to get Grid Condition in filterString
                'InsertSquareBracket = "["
                For rr = 0 To GridContionTable.Rows.Count - 1
                    Dim ParentKey As Integer = GridContionTable(rr)("FilterDetails_Key")
                    GridTable = SearchDataTable(FilterDetails, "ParentFilterDetailsKey,DefineFlag", ParentKey & ",True") '---Check if further Filteration exit for the gridColumn
                    If GridTable.Rows.Count = 0 Then
                        Return mFinalStatus
                        Exit Function
                    Else '-----Get Futher Filterartion Condition for the GridColumn if Exist.
                        GridStr = GetGridString(GridContionTable.Rows(rr), TableAlias)
                        FurtherGridStr = GetSubGridString(FilterDetails, GridTable, FilterSelected, SubTableCount)
                        FurtherGridStr = FurtherGridStr & ")"
                        TempStr = IIf(TempStr = "", GridStr & FurtherGridStr, TempStr & GridStr & FurtherGridStr)
                    End If
                Next
                mFinalStatus = mFinalStatus & TempStr
                'mFinalStatus = mFinalStatus & InsertSquareBracket & TempStr
                'mFinalStatus = mFinalStatus & " ]"
            End If
            Return mFinalStatus
        Catch ex As Exception
            QuitError(ex, Err, "Private Function GetSubGridString(ByVal FilterDetails As DataTable, ByVal GridTable As DataTable, ByVal FilterSelected As DataTable, ByRef SubTableCount As Int16) As String")
            Return Nothing
        End Try
    End Function


    Private Function CreateConditionString(ByVal GridRow As DataRow, ByVal TableNameAlias As String, ByVal FilterSelected As DataTable) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mstatus As String = ""
            Dim mConditionColoumn As String = ""
            Dim mEndValue As String = ""
            Dim mStartValue As String = ""
            Dim mDialogType0 As String = ""
            Dim mLogicalOperator As String = ""
            Dim mDefineFlag As String = ""
            Dim mConditionType As String = GridRow("ConditionType")
            Dim mFilterBy As String = GridRow("FilterBy")
            Dim mRelationalOperator As String = GridRow("RelationalOperator")


            mDefineFlag = GridRow("DefineFlag").ToString
            If mDefineFlag = True Then
                mConditionColoumn = TableNameAlias & "." & GridRow("ConditionColumn")
                mStartValue = IIf(IsDBNull(GridRow("StartValue")) = True, "", GridRow("StartValue"))
                mEndValue = IIf(IsDBNull(GridRow("EndValue")) = True, "", GridRow("EndValue"))
                mDialogType0 = GridRow("DialogType")
                mLogicalOperator = GridRow("LogicalOperator").ToString


                Select Case True
                    Case LCase(mFilterBy) = "r"
                        Dim str As String = mConditionColoumn & " " & " >= "

                        Select Case LCase(mDialogType0)
                            Case "string"
                                str = str & "'" & mStartValue.ToString & "'  and  " & mConditionColoumn & " <= '" & mEndValue.ToString & "'"
                            Case "integer", "integer1", "decimal", "decimal1"
                                str = str & mStartValue.ToString & "  and  " & mConditionColoumn & " <= " & mEndValue.ToString
                            Case "datetime", "time"
                                str = str & Convert.ToDateTime(mStartValue) & " and  " & mConditionColoumn & " <=  " & Convert.ToDateTime(mEndValue)

                        End Select
                        str = "(" & str & ")"
                        mLogicalOperator = IIf(mLogicalOperator.Trim.Length = 0, " and ", mLogicalOperator)
                        mstatus = mstatus & IIf(mstatus.Length = 0, "", " " & mLogicalOperator & " ") & str


                    Case LCase(mFilterBy) = "s"

                        Dim str As String = ""
                        If LCase(mDialogType0) = "string" Then
                            str = ""
                            Select Case LCase(mRelationalOperator.Replace(" ", ""))
                                Case "equalsto"
                                    str = mConditionColoumn & " " & " = '" & mStartValue.ToString & "'"
                                Case "notequalsto"
                                    str = mConditionColoumn & " " & " <> '" & mStartValue.ToString & "'"
                                Case "like%"
                                    str = mConditionColoumn & " " & " Like '" & mStartValue.ToString & "%'"
                                Case "instring"
                                    str = "CHARINDEX ( '" & mStartValue.ToString & "'," & mConditionColoumn & ") > 0 "
                                Case "containedby"
                                    str = "Contains (" & mConditionColoumn & ", '" & mStartValue.ToString & "')"
                            End Select
                        Else
                            str = ""
                            Select Case LCase(mRelationalOperator)
                                Case "equalsto"
                                    str = mConditionColoumn & " " & " = "
                                Case "notequalsto"
                                    str = mConditionColoumn & " " & " <> "
                                Case "greaterthan"
                                    str = mConditionColoumn & " " & " > "
                                Case "lessthan"
                                    str = mConditionColoumn & " " & " < "
                            End Select
                            Select Case LCase(mDialogType0)
                                Case "integer", "integer1", "decimal", "decimal1"
                                    str = str & mStartValue.ToString
                                Case "datetime", "time"
                                    str = str & Convert.ToDateTime(mStartValue)
                            End Select
                        End If

                        str = "(" & str & ")"
                        mLogicalOperator = IIf(mLogicalOperator.Trim.Length = 0, " and ", mLogicalOperator)
                        mstatus = str
                        'ConditionGrid.SourceTable.Rows(mrowno).Item("Condition") = mstatus

                    Case LCase(mFilterBy) = "v"
                        mstatus = ""
                        'mConditionColoumn = TableNameAlias & "." & GridRow("ConditionColumn")
                        mstatus = GridRow("Condition")
                        'mstatus = mstatus.Replace(mConditionColoumn, TableNameAlias & "." & mConditionColoumn)
                        Dim ByValuesSelectedItems As DataTable = SearchDataTable(FilterSelected, "FilterDetailsKey", GridRow("FilterDetails_Key").ToString)
                        Dim str As String = ""
                        For ww = 0 To ByValuesSelectedItems.Rows.Count - 1
                            Select Case LCase(mDialogType0)
                                Case "string"
                                    str = str & IIf(str = "", "'" & ByValuesSelectedItems(ww)("ValuesString").ToString.Trim & "'", ", " & "'" & ByValuesSelectedItems(ww)("ValuesString").ToString.Trim & "'")
                                Case "integer", "integer1", "decimal", "decimal1"
                                    str = str & IIf(str = "", ByValuesSelectedItems(ww)("ValuesString").ToString.Trim, ", " & ByValuesSelectedItems(ww)("ValuesString").ToString.Trim)
                                Case "datetime", "time"
                                    str = str & IIf(str = "", Convert.ToDateTime(ByValuesSelectedItems(ww)("ValuesString").ToString.Trim), ", " & Convert.ToDateTime(ByValuesSelectedItems(ww)("ValuesString").ToString.Trim))
                            End Select
                        Next
                        str = mConditionColoumn & " In " & "(" & str & ")"
                        mLogicalOperator = IIf(mLogicalOperator.Trim.Length = 0, " and ", mLogicalOperator)
                        'mstatus = mstatus & IIf(mstatus.Length = 0, "", " " & mLogicalOperator & " ") & str
                        mstatus = str
                        'End If

                End Select
            End If
            Return mstatus
        Catch ex As Exception
            QuitError(ex, Err, "Private Function CreateConditionString(ByVal GridRow As DataRow, ByVal TableNameAlias As String, ByVal FilterSelected As DataTable) As String")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Convert a datatable into string with or without ColumnNames.String started with chrw(248) if columnheaders present in the string. 
    ''' </summary>
    ''' <param name="FileDt">DataTable for which output string is to be obtained</param>
    ''' <param name="mColumnsSeparator">Separator used to separate different column values.Default is |</param>
    ''' <param name="mRowsSeparator" >Separator used to separate different column values.Default is chrw(13)</param>
    ''' <param name="ExcludeFields">, separated string of Columns to be Excluded while creating string</param>
    ''' <param name="mHeader">Indicates whether columnnames of the data table  to be saved on not on first row position.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DataTableToString(ByVal FileDt As DataTable, Optional ByVal mColumnsSeparator As Object = "|", Optional ByVal mRowsSeparator As Object = ChrW(13), Optional ByVal ExcludeFields As String = "", Optional ByVal mHeader As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If FileDt Is Nothing Then
            Return Nothing
        End If

        Dim ColArr As String() = {}
        If ExcludeFields.Length > 0 Then
            ColArr = LCase(ExcludeFields).Split(",")
        End If

        Dim HeaderStr As String = ""
        If mHeader = True Then
            For k = 0 To FileDt.Columns.Count - 1
                If ExcludeFields.Length > 0 Then
                    If ColArr.Contains(LCase(FileDt.Columns(k).ColumnName)) Then
                        Continue For
                    End If
                End If
                HeaderStr = HeaderStr & mColumnsSeparator & FileDt.Columns(k).ColumnName
            Next
            HeaderStr = HeaderStr.Remove(0, 1)
            HeaderStr = ChrW(248) & HeaderStr
        End If

        Dim FinalStr As String = ""
        For Each rw As DataRow In FileDt.Rows
            Dim str As String = ""
            For k = 0 To FileDt.Columns.Count - 1
                If ExcludeFields.Length > 0 Then
                    If ColArr.Contains(LCase(rw.Table.Columns(k).ColumnName)) Then
                        Continue For
                    End If
                End If
                If IsDBNull(rw(k)) = False Then
                    If rw(k).GetType.ToString = "System.DateTime" Then
                        str = str & mColumnsSeparator & CDate(rw(k))
                    Else
                        str = str & mColumnsSeparator & CStr(rw(k)).ToString
                    End If
                Else
                    FinalStr = FinalStr & mColumnsSeparator & ""
                End If

            Next
            If str.Length > 0 Then
                str = str.Remove(0, 1)
            End If
            FinalStr = FinalStr & mRowsSeparator & str
        Next
        If FinalStr.Length > 0 Then
            FinalStr = FinalStr.Remove(0, 1)
        End If
        FinalStr = IIf(Len(HeaderStr) = 0, "", HeaderStr & mRowsSeparator) & FinalStr
        Return FinalStr
    End Function

    ''' <summary>
    ''' Convert a datatRow into string with or without ColumnNames.Columnames and row itemarray  seperated by carriage return. 
    ''' </summary>
    ''' <param name="FileRw">DataRow for which output string is to be obtained</param>
    ''' <param name="mColumnsSeparator">Separator used to separate different column values.Default is |</param>
    ''' <param name="mRowsSeparator" >Separator used to separate different row values.Default is chrw(13)</param>
    ''' <param name="ExcludeFields">Comma separated string of Columns to be Excluded while creating string</param>
    ''' ''' <param name="mHeader">Indicates whether Heading of the table is to be saved on not.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DataRowToString(ByVal FileRw As DataRow, Optional ByVal mColumnsSeparator As Object = "|", Optional ByVal mRowsSeparator As Object = ChrW(13), Optional ByVal ExcludeFields As String = "", Optional ByVal mHeader As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If FileRw Is Nothing Then
            Return Nothing
        End If

        Dim ColArr As String() = {}
        If ExcludeFields.Length > 0 Then
            ColArr = LCase(ExcludeFields).Split(",")
        End If

        Dim HeaderStr As String = ""
        If mHeader = True Then
            For k = 0 To FileRw.Table.Columns.Count - 1
                If ExcludeFields.Length > 0 Then
                    If ColArr.Contains(LCase(FileRw.Table.Columns(k).ColumnName)) Then
                        Continue For
                    End If
                End If
                HeaderStr = HeaderStr & mColumnsSeparator & FileRw.Table.Columns(k).ColumnName
            Next
            HeaderStr = HeaderStr.Remove(0, 1)
            HeaderStr = ChrW(248) & HeaderStr
        End If

        Dim FinalStr As String = ""
        For k = 0 To FileRw.Table.Columns.Count - 1
            If ExcludeFields.Length > 0 Then
                If ColArr.Contains(LCase(FileRw.Table.Columns(k).ColumnName)) Then
                    Continue For
                End If
            End If
            If IsDBNull(FileRw(k)) = False Then
                If FileRw(k).GetType.ToString = "System.DateTime" Then
                    FinalStr = FinalStr & mColumnsSeparator & CDate(FileRw(k))
                Else
                    FinalStr = FinalStr & mColumnsSeparator & CStr(FileRw(k))
                End If
            Else
                FinalStr = FinalStr & mColumnsSeparator & ""
            End If


        Next
        If FinalStr.Length > 0 Then
            FinalStr = FinalStr.Remove(0, 1)
        End If
        FinalStr = IIf(Len(HeaderStr) = 0, "", HeaderStr & mRowsSeparator) & FinalStr

        Return FinalStr
    End Function
    ''' <summary>
    ''' To make a datatable object from a string protocol,string starts with chrw(248),if header exists
    ''' </summary>
    ''' <param name="InputString">Input string from which data is to be converted into datatable</param>
    ''' <param name="mColumnsSeparator" >Separator used to separate different column values.Default is |</param>
    ''' <param name="mRowsSeparator" >Separator used to separate different column values.Default is chrw(13)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StringToDataTable(ByVal InputString As String, Optional ByVal mColumnsSeparator As String = "|", Optional ByVal mRowsSeparator As String = ChrW(13), Optional ByVal mDataTable As DataTable = Nothing) As DataTable

        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If InputString.Length = 0 Then
            Return Nothing
        End If
        Dim dtIMS As New DataTable
        If mDataTable Is Nothing Then
            mDataTable = dtIMS.Copy
        End If

        Dim HeaderCheck As Char = InputString.Substring(0, 1)
        Dim IMSVal() As String = {}
        If HeaderCheck = ChrW(248) Then
            InputString = InputString.Remove(0, 1)
            IMSVal = Split(InputString, mRowsSeparator)
            Dim mHeading() As String = Split(IMSVal(0), mColumnsSeparator)
            For hh = 0 To mHeading.Count - 1
                mDataTable = AddColumnsInDataTable(dtIMS, mHeading(hh).Trim)
            Next
            For ee = 1 To IMSVal.Count - 1
                Dim RowVal() As String = Split(IMSVal(ee), mColumnsSeparator)
                Dim dtRw As DataRow = dtIMS.NewRow
                For cc = 0 To RowVal.Count - 1
                    dtRw(cc) = RowVal(cc).Trim
                Next
                AddRowInDataTable(mDataTable, dtRw, True)
            Next
        Else
            IMSVal = Split(InputString, mRowsSeparator)
            For ee = 0 To IMSVal.Count - 1
                Dim RowVal() As String = Split(IMSVal(ee), mColumnsSeparator)
                Dim mcols As Int16 = mDataTable.Columns.Count
                For hh = mcols To RowVal.Count - 1
                    mDataTable = AddColumnsInDataTable(mDataTable, "Col" & hh)
                Next
                Dim dtRw As DataRow = mDataTable.NewRow
                For cc = 0 To RowVal.Count - 1
                    dtRw(cc) = RowVal(cc).Trim
                Next
                AddRowInDataTable(mDataTable, dtRw, True)
            Next
        End If
        Return mDataTable
    End Function
    ''' <summary>
    ''' To make a datarow object from a string protocol,string starts with chrw(248),if header exists
    ''' </summary>
    ''' <param name="InputString">Input string from which data is to be converted into datatable</param>
    ''' <param name="mColumnsSeparator" >Separator used to separate different column values.Default is |</param>
    ''' <param name="mRowsSeparator" >Separator used to separate different column values.Default is chrw(13)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StringToDataRow(ByVal InputString As String, Optional ByVal mColumnsSeparator As String = "|", Optional ByVal mRowsSeparator As String = ChrW(13)) As DataRow
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If InputString.Length = 0 Then
            Return Nothing
        End If

        Dim dtIMS As New DataTable
        Dim HeaderCheck As Char = InputString.Substring(0, 1)
        Dim IMSVal() As String = {}
        If HeaderCheck = ChrW(248) Then
            InputString = InputString.Remove(0, 1)
            IMSVal = InputString.Split(mRowsSeparator)
            Dim mHeading() As String = IMSVal(0).Split(mColumnsSeparator)
            For hh = 0 To mHeading.Count - 1
                dtIMS = AddColumnsInDataTable(dtIMS, mHeading(hh))
            Next
            For ee = 1 To IMSVal.Count - 1
                Dim RowVal() As String = IMSVal(ee).Split(mColumnsSeparator)
                Dim dtRw As DataRow = dtIMS.NewRow
                For cc = 0 To RowVal.Count - 1
                    dtRw(cc) = RowVal(cc)
                Next
                AddRowInDataTable(dtIMS, dtRw, True)
            Next
        Else
            IMSVal = InputString.Split(mRowsSeparator)
            Dim ColCount() As String = IMSVal(0).Split(mColumnsSeparator)
            For hh = 0 To ColCount.Count - 1
                dtIMS = AddColumnsInDataTable(dtIMS, "Col" & hh)
            Next
            For ee = 0 To IMSVal.Count - 1
                Dim RowVal() As String = IMSVal(ee).Split(mColumnsSeparator)
                Dim dtRw As DataRow = dtIMS.NewRow
                For cc = 0 To RowVal.Count - 1
                    dtRw(cc) = RowVal(cc)
                Next
                AddRowInDataTable(dtIMS, dtRw, True)
            Next
        End If
        Dim dtrow As DataRow = Nothing
        If dtIMS.Rows.Count > 0 Then
            dtrow = dtIMS.Rows(0)
        End If
        Return dtrow
    End Function
    ''' <summary>
    ''' To create a textfile of insert queries as per the records fetched by database table
    ''' </summary>
    ''' <param name="filePath">Full file path including name where text file is to be created</param>
    ''' <param name="serverDataBaseName">Serverdatabase name in 0_srv_0.0_mdf_0 format</param>
    ''' <param name="tableName">Name of table</param>
    ''' <param name="condition">Condition in SQL query</param>
    ''' <param name="NoOfRows" >No. of rows to be taken in one script</param>
    ''' <param name="WithFieldNames" >False ,if all fields values supplied in value clause</param>
    ''' <remarks></remarks>
    Public Sub WriteInsertQueryScript(ByVal filePath As String, ByVal serverDataBaseName As String, ByVal tableName As String, Optional ByVal condition As String = "", Optional ByVal NoOfRows As Integer = 0, Optional ByVal WithFieldNames As Boolean = False)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim dt As New DataTable
        dt = GetDataFromSql(serverDataBaseName, tableName, "*", "", condition, "", "")
        Dim TotRows As Integer = dt.Rows.Count
        NoOfRows = IIf(NoOfRows = 0, TotRows, NoOfRows)
        Dim nofiles As Integer = Math.Ceiling(TotRows / NoOfRows)
        Dim file2 As List(Of String) = GF1.FullFileNameToList(filePath)
        If System.IO.Directory.Exists(file2(0)) = False Then
            System.IO.Directory.CreateDirectory(file2(0))
        End If
        Dim newname As String = file2(1)
        For k = 0 To nofiles - 1
            file2(1) = newname & "_" & k.ToString
            Dim newfile As String = GF1.GetFullFileName(file2)
            Dim file1 As System.IO.StreamWriter
            file1 = My.Computer.FileSystem.OpenTextFileWriter(newfile, False)
            Dim kk As Integer = k * NoOfRows
            For i = kk To kk + NoOfRows - 1
                If i > TotRows - 1 Then
                    Continue For
                End If
                Dim col1 As String = ""
                Dim seperatedColName As String = ""
                Dim row1 As String = ""
                For j = 0 To dt.Columns.Count - 1
                    Dim ColDataType As String = dt.Columns(j).DataType.ToString
                    col1 = dt.Columns(j).ColumnName
                    If Not IsDBNull(dt.Rows(i).Item(col1)) Then
                        seperatedColName += dt.Columns(j).ColumnName + ","
                        Dim t As Type = dt.Columns(j).DataType
                        If ColDataType = "System.String" Or ColDataType = "System.DateTime" Or ColDataType = "System.Boolean" Then
                            row1 += "'" & Trim(dt.Rows(i).Item(col1).ToString) & "'" + ","
                        ElseIf ColDataType = "System.DateTime" Then
                            Dim tempdate As Date = dt.Rows(i).Item(col1)
                            Dim a As String = tempdate.ToString("yyyy-MM-dd hh:mm:ss ")
                            'End If
                            row1 += "'" & a & "'" + ","
                        Else
                            row1 += dt.Rows(i).Item(col1).ToString + ","
                        End If
                    End If
                Next
                seperatedColName = seperatedColName.TrimEnd(",")
                row1 = row1.TrimEnd(",")
                If WithFieldNames = True Then
                    file1.WriteLine("INSERT INTO " & tableName & "(" & seperatedColName & ") VALUES (" & row1 & ")")
                Else
                    file1.WriteLine("INSERT INTO " & tableName & "  VALUES (" & row1 & ")")
                End If

            Next
            file1.Close()
            file1.Dispose()
        Next

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub


End Class
