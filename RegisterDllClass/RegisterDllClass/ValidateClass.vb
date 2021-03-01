Imports System.Windows.Forms
Imports System.Management
Imports System.IO
Imports System
Imports Microsoft.Win32
Imports System.Net.Mail
Imports System.Web
Imports System.Net
Imports System.Security.Permissions

Public Class ValidateClass
    '    Dim GF1 As New GlobalFunction1.GlobalFunction1
    Private Shared RegistryFolder As String = GlobalControl.Variables.RegistryFolder
    Public aumess As String = "deliaf noitacitnehtua"
    Dim sep1 As String = ChrW(211)
    Dim sep2 As String = ChrW(212)
    Dim WmiClassProperty0() As String = {"Win32_BASEBOARD", "manufacturer", "model", "product", "serialnumber", "installdate", "partnumber", "sku"}
    Dim WmiClassProperty1() As String = {"Win32_BIOS", "identificationcode", "biosversion", "releasedate", "serialnumber", "SMBIOSBIOSVersion", "installdate", "manufacturer", "softwareelementid", "version"}
    Dim WmiClassProperty2() As String = {"Win32_ComputerSystem", "InstallDate", "manufacturer", "name", "model"}
    Dim WmiClassProperty3() As String = {"Win32_Processor", "deviceid", "CurrentClockSpeed", "name", "ProcessorId", "version", "uniqueid", "installdate"}
    Dim WmiClassProperty4() As String = {"Win32_MOTHERBOARDDEVICE", "installdate"}
    Dim WmiClassProperty5() As String = {"Win32_ONBOARDDEVICE", "installdate", "serialnumber", "sku"}
    Dim WmiClassProperty6() As String = {"Win32_diskdrive", "model", "installdate", "totalheads"}
    Dim WmiClassProperty7() As String = {"Win32_videocontroller", "installdate"}

    Public Function AdminVault() As Integer
        If LCase("WebAzure,WebGodaddy,WebCloud,WebLocal").Contains(LCase(GlobalControl.Variables.SaralType)) = True And Not GlobalControl.Variables.SaralType = "" Then
            GlobalControl.Variables.AuthenticationChecked = GlobalControl.Variables.xControl + 130858
            Return GlobalControl.Variables.AuthenticationChecked
            Exit Function
        End If
        If GlobalControl.Variables.AuthenticationChecked <> GlobalControl.Variables.xControl + 130858 Then
            Try

                Dim KeyFolder As String = Environment.GetEnvironmentVariable("userprofile") & "\SaralKeyFolder"
                Dim logfile As String = KeyFolder & "\abc.txt"
                Dim RegistryDemo As String = "HARDWARE\\DESCRIPTION\\SYSTEM\\BIOS"
                If Not My.Computer.FileSystem.DirectoryExists(KeyFolder) Then
                    Try
                        My.Computer.FileSystem.CreateDirectory(KeyFolder)
                       
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to create folder " & KeyFolder)
                    End Try
                End If
                Dim LCurrentDate As String = DateTostringFunction(Now.Date)
                Dim HrdDayFile As String = KeyFolder & "\HrdTemp.Txt", Newhrd As Boolean = True

                Dim TempHash As New Hashtable
                Dim RegDLLHash As New Hashtable
                If My.Computer.FileSystem.FileExists(HrdDayFile) Then
                    Try
                        Dim TempStr As String = StringDecript(StringRead(HrdDayFile), 35)
                        TempHash = StringToHashTable(TempStr, sep1, sep2)
                        If ValueFromHashTable(TempHash, "currdt") = LCurrentDate Then
                            Newhrd = False
                        End If
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to read file  " & HrdDayFile)
                    End Try
                End If
                If Newhrd Then
                    Try
                        Dim TempStr As String = CreateHRDTemp(RegistryFolder)
                        TempHash = StringToHashTable(TempStr, sep1, sep2)
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to regenerate hardware file  in folder  " & RegistryFolder)
                    End Try
                End If
                Dim ARegFiles As List(Of String) = SearchFiles(KeyFolder, "*.XDL")
                If ARegFiles.Count > 1 Then
                    QuitMessage("More Than one *.XDL files in the folder " & KeyFolder & " not allowed")
                    Return GlobalControl.Variables.AuthenticationChecked
                    Exit Function
                End If
                Dim HrdEmlFile As String = KeyFolder & "\HrdEml.Txt"
                Dim NewEml As Boolean = True
                Dim emaildt As String = DateTostringFunction(Now.Date)
                If My.Computer.FileSystem.FileExists(HrdEmlFile) Then
                    Try
                        Dim TempStr As String = StringDecript(StringRead(HrdEmlFile), 35)
                        Dim TempHash0 As Hashtable = StringToHashTable(TempStr, sep1, sep2)
                        emaildt = ValueFromHashTable(TempHash0, "emaildt")
                        If LCurrentDate < emaildt Then
                            NewEml = False
                        End If
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to read file  " & HrdEmlFile)
                    End Try
                End If
                If NewEml = True Then
                    If My.Computer.Network.IsAvailable = True Then
                        emaildt = DateTostringFunction(Now.AddDays(10))
                        Dim emailstr = "emaildt" & sep2 & emaildt
                        Dim finalstr1 As String = StringEncript(emailstr, 35)
                        HrdEmlFile = StringWrite(HrdEmlFile, finalstr1)
                        Dim maregfile As String = ""
                        If ARegFiles.Count > 0 Then
                            maregfile = ARegFiles(0)
                        End If
                        Dim mattach As String = IIf(maregfile.Length > 0, maregfile & ",", "") & HrdEmlFile & "," & HrdDayFile
                        SendingEmail("hcgupta@saralerp.com", "Software running machine " & My.Computer.Name, "Hardware details of software running machine ip address :" & GetIPAddress(), mattach)
                    End If
                End If
                'Read Regfile
                If ARegFiles.Count = 0 Then
                    Dim sdate As String = GetRegistryValues(RegistryDemo, "SystemSaral")
                    If sdate.Trim.Length = 0 Then
                        sdate = DateTostringFunction(Now.AddDays(15))
                        SetRegistryValues(RegistryDemo, "SystemSaral", sdate)
                    End If
                    Dim nowdate As String = DateTostringFunction(Now)
                    Dim mdate As Date = StringToDate(sdate)
                    Select Case True
                        Case nowdate > sdate
                            SetRegistryValues(RegistryDemo, "SystemSaral", "10010101")
                            sdate = "10010101"
                        Case Now < mdate.AddDays(-20)
                            SetRegistryValues(RegistryDemo, "SystemSaral", "10010101")
                            sdate = "10010101"
                        Case Else
                            '   GlobalControl.Variables.AuthenticationChecked = GlobalControl.Variables.xControl + 130858
                    End Select
                    GlobalControl.Variables.DemoDate = StringToDate(sdate)
                    MsgBox("It is a Demo Version of application")
                    Return GlobalControl.Variables.AuthenticationChecked
                    Exit Function
                End If
                Dim RegStr As String = StringRead(ARegFiles(0))
                If Not Left(RegStr, 1) = ChrW(15) Then
                    QuitMessage("Invalid registration file")
                    Return GlobalControl.Variables.AuthenticationChecked
                    Exit Function
                End If
                Dim LRegStr As String = StringDecript(Mid(RegStr, 2, RegStr.Length - 1), 40)
                Dim RegHash As Hashtable = StringToHashTable(LRegStr, sep1, sep2)
                Dim mclass As String = WmiClassProperty0(0)
                For i = 1 To WmiClassProperty0.Count - 1
                    Dim mkeyname As String = LCase(mclass & "_" & WmiClassProperty0(i))
                    If Not ValueFromHashTable(TempHash, mkeyname) = ValueFromHashTable(RegHash, mkeyname) Then
                        Dim tvalue As String = ValueFromHashTable(TempHash, mkeyname)
                        Dim rvalue As String = ValueFromHashTable(RegHash, mkeyname)
                        QuitMessage(mkeyname & " mismatched value (" & tvalue & " and " & rvalue & "), Not authenticated")
                        Return GlobalControl.Variables.AuthenticationChecked
                        Exit Function
                    End If
                Next
                GlobalControl.Variables.AuthenticationChecked = GlobalControl.Variables.xControl + 130858
                Dim mallowdt As String = ValueFromHashTable(RegHash, LCase("allowdate"))
                If mallowdt IsNot Nothing Then
                    GlobalControl.Variables.AllowDate = New Date(CInt(Left(mallowdt, 4)), CInt(Mid(mallowdt, 5, 2)), CInt(Right(mallowdt, 2)))
                End If

            Catch ex As Exception
                QuitError(ex, Err, "Unable to execute ValidateClass.AdminVault")
            End Try
        End If
        '   GlobalControl.Variables.AuthenticationChecked = VaultFlag
        Return GlobalControl.Variables.AuthenticationChecked
    End Function
    Public Function ReverseString(ByVal InputString As String, ByVal Seedno As Integer) As String
        Dim kk As String = ""
        For i = InputString.Length To 1 Step -1 * Seedno
            kk = kk & Mid(InputString, i, Seedno)
        Next
        Return kk
    End Function

    Private Function CreateHRDTemp(ByVal RegistryFolder As String) As String
        Dim finalstr As String = ""
        Try
            Dim LCurrentDate As String = DateTostringFunction(Now.Date)
            Dim KeyFolder As String = Environment.GetEnvironmentVariable("userprofile") & "\SaralKeyFolder"
            If Not System.IO.Directory.Exists(KeyFolder) Then
                System.IO.Directory.CreateDirectory(KeyFolder)
            End If
            Dim HardwareFile As String = KeyFolder & "\HRDTEMP.TXT"
            finalstr = "currdt" & sep2 & LCurrentDate & sep1 & MachRead(RegistryFolder)
            Dim finalstr1 As String = StringEncript(finalstr, 35)
            HardwareFile = StringWrite(HardwareFile, finalstr1)
            Return finalstr
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.CreateHrdTemp")
        End Try
        Return finalstr
    End Function

    Private Function MachRead(ByVal RegistryFolder As String) As String
        Dim MachStr As String = ""
        Try

            Dim LRegDrive As String = GetRegistryValues(RegistryFolder, "RegDrive")
            Dim LDriveType As Integer = 0
            If LRegDrive.Length = 0 Then
                LRegDrive = Left(Environment.CurrentDirectory, 2)
                SetRegistryValues(RegistryFolder, "RegDrive", LRegDrive)
            End If
            Dim finalStr As String = ""
            Dim result As Hashtable = identifier0(WmiClassProperty0)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty1)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty2)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty3)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty4)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty5)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty6)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            result = identifier0(WmiClassProperty7)
            If result.Count > 0 Then
                finalStr = finalStr & IIf(finalStr.Length = 0, "", sep1) & HashTableToString(result, sep1, sep2)
            End If
            finalStr = RemoveChar(finalStr, " ")
            MachStr = finalStr
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.MachRead")
        End Try
        Return MachStr
    End Function


    Private Function StringEncript(ByVal InputStr As String, ByVal SeedNo As Integer) As String
        Dim Ostr As String = "", Ls1 As String = "", Lsa1 As String = "", nlen As Integer, nrem As Integer, nquot As Integer
        Try
            nquot = Math.DivRem(Len(InputStr), 10, nrem)
            nlen = IIf(Len(InputStr) <= 10, 1, ((Len(InputStr) + 10 - nrem) / 10))

            For i = 0 To nlen - 1
                Ls1 = Mid(InputStr, i * 10 + 1, 10)
                For j = 0 To Len(Ls1) - 1
                    Ostr = Ostr + ChrW(AscW(Mid(Ls1, j + 1, 1)) + (j + 1) * (j + 1) + SeedNo)
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.StringEncript")
        End Try
        Return Ostr
    End Function

    Private Function StringDecript(ByVal InputStr As String, ByVal SeedNo As Integer) As String
        Dim Ostr As String = "", Ls1 As String = "", Lsa1 As String = "", nlen As Integer, nrem As Integer, nquot As Integer
        Try
            nquot = Math.DivRem(Len(InputStr), 10, nrem)
            nlen = IIf(Len(InputStr) <= 10, 1, ((Len(InputStr) + 10 - nrem) / 10))

            For i = 0 To nlen - 1
                Ls1 = Mid(InputStr, i * 10 + 1, 10)
                For j = 0 To Len(Ls1) - 1
                    Ostr = Ostr + ChrW(AscW(Mid(Ls1, j + 1, 1)) - (j + 1) * (j + 1) - SeedNo)
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.StringEncript")
        End Try
        Return Ostr
    End Function


    Private Function identifier0(ByVal wmiClassProperty() As String) As Hashtable
        Dim result As New Hashtable
        Try
            Dim kk As New System.Management.ManagementObjectSearcher("select * from " & wmiClassProperty(0))
            For Each ii In kk.Get
                Try
                    Dim jj As System.Management.PropertyDataCollection = ii.Properties
                    For i = 1 To wmiClassProperty.Count - 1
                        Try
                            If wmiClassProperty(i).ToString.Length > 0 Then
                                Dim keyname As String = LCase(wmiClassProperty(0) & "_" & wmiClassProperty(i))
                                Dim xvalue As Object = ii(wmiClassProperty(i).ToString)
                                If Not xvalue Is Nothing Then
                                    result.Add(keyname, ii(wmiClassProperty(i)).ToString.Trim)
                                End If
                            End If
                        Catch ex As Exception
                            '   MsgBox("Exception1")
                            Continue For
                        End Try
                    Next
                Catch ex As Exception
                    '   MsgBox("Exception2")
                    Continue For
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.identifier0(ByVal wmiClassProperty() As String)")


            ' MsgBox("Exception3")
        End Try
        Return result
    End Function

    Private Sub QuitError(ByVal ex As Exception, ByVal err As ErrObject, ByVal Mess As String)
        If LCase("WebAzure,WebServer,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            MsgBox(ex.Message & " Stacktrace " & ex.StackTrace.ToString & " Procedure " & err.Erl.ToString & vbCrLf & "Message: " & Mess)
            System.Windows.Forms.Application.ExitThread()
            System.Windows.Forms.Application.Exit()
            Process.GetCurrentProcess.Kill()
        End If
    End Sub
    Private Function ArrayAppend(ByRef ArrayName() As Control, ByVal LastValue As Control) As Control()
        Try
            Dim ii As Integer = ArrayName.Length
            ReDim Preserve ArrayName(ii)
            ArrayName.SetValue(LastValue, ii)
            Return ArrayName
            Exit Function
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.ArrayAppend)")
        End Try
        Return ArrayName
    End Function
    Private Function DateTostringFunction(ByVal InputDate As Date, Optional ByVal DateFormatString As String = "", Optional ByRef DisplayDateString As String = "") As String
        DateTostringFunction = ""
        Try
            Dim lday As String = Microsoft.VisualBasic.Right(CStr(100 + InputDate.Day), 2)
            Dim lmon As String = Microsoft.VisualBasic.Right(CStr(100 + InputDate.Month), 2)
            Dim lyear As String = CStr(InputDate.Year)
            DateTostringFunction = lyear & lmon & lday
            If Not DateFormatString = "" Then
                DisplayDateString = InputDate.ToString(DateFormatString)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.DateTostringFunction)")
        End Try
    End Function
    Private Function StringToDate(ByVal InputDateString As String) As Date
        Dim mdate As Date = #1/1/1900#
        Try
            If InputDateString.Trim.Length = 8 Then
                mdate = New Date(CInt(Left(InputDateString, 4)), CInt(Mid(InputDateString, 5, 2)), CInt(Right(InputDateString, 2)))
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.DateTostringFunction)")
        End Try
        Return mdate
    End Function

    Private Function StringWrite(ByVal TxtFile As String, ByVal TxtString As String) As String
        Dim retstr As String = ""
        Try
            Dim fsw As System.IO.FileStream
            fsw = New System.IO.FileStream(TxtFile, System.IO.FileMode.Create, IO.FileAccess.ReadWrite)
            Dim sw As New System.IO.StreamWriter(fsw, System.Text.Encoding.UTF8)
            sw.Write(TxtString)
            sw.Flush()
            sw.Close()
            fsw.Close()
            retstr = TxtFile
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.StringWrite " & TxtFile)
        End Try
        Return retstr
    End Function
    Private Function StringRead(ByVal TxtFile As String) As String
        Dim outstr As String = ""
        Try
            Dim fs As System.IO.FileStream = New System.IO.FileStream(TxtFile, FileMode.Open, FileAccess.ReadWrite)
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(fs)
            Dim NBuff(0) As Char
            Do While sr.Peek() >= 0
                sr.Read(NBuff, 0, 1)
                If Asc(NBuff(0)) > 0 Then
                    outstr = outstr & IIf(Asc(NBuff(0)) > 0, NBuff(0), "")
                End If
            Loop
            sr.Close()
            fs.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.StringRead " & TxtFile)
        End Try
        Return outstr
    End Function


    Private Function SearchFiles(ByVal SourcePath As String, Optional ByVal WildCard As String = "*.*", Optional ByVal TopLevel As Boolean = True) As List(Of String)
        Dim filelist As New List(Of String)
        Try
            Dim aa As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(SourcePath)
            aa.GetFiles()
            For Each fi1 As System.IO.FileInfo In aa.GetFiles(WildCard)
                filelist.Add(UCase(Trim(fi1.FullName)))
            Next
            If TopLevel = False Then
                For Each dir1 As System.IO.DirectoryInfo In aa.GetDirectories
                    SearchFiles(dir1.FullName.ToString, WildCard, TopLevel)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.SearchFiles " & SourcePath)
        End Try

        Return filelist
    End Function
    Private Function RemoveChar(ByVal InputString As String, Optional ByVal RemChar As String = "") As String
        'to remove letters
        Dim NewString As String = ""
        Try
            For i = 1 To InputString.Length
                If AscW(Mid(InputString, i, 1)) > 0 Then
                    If (Not Mid(InputString, i, 1) = RemChar) Or RemChar.Length = 0 Then
                        NewString = NewString & Mid(InputString, i, 1)
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.RemoveChar ")
        End Try
        Return NewString
    End Function
    Private Function BreakFileName(ByVal FullFileName As String) As List(Of String)
        '0=patth,1=filename,2=extension
        Dim mpath As String = ""
        Dim mfilename As String = ""
        Dim mext As String = ""
        Dim mpathnameext As New List(Of String)
        Dim dd As Integer, ee As Integer
        Try
            If FullFileName.Trim.Length > 0 Then
                dd = FullFileName.LastIndexOfAny("\")
                mpath = Left(FullFileName, dd)
                ee = FullFileName.LastIndexOfAny(".")
                mext = Right(FullFileName, FullFileName.Length - ee - 1)
                mfilename = Mid(FullFileName, dd + 2, FullFileName.Length - mpath.Length - mext.Length - 2)
                mpathnameext.Add(mpath)
                mpathnameext.Add(mfilename)
                mpathnameext.Add(mext)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.BreakFileName ")

        End Try

        Return mpathnameext
    End Function
    Private Function GetFullFileName(ByVal PathNameExt As List(Of String)) As String
        '0=patth,1=filename,2=extension
        Dim RetStr As String = ""
        Try

            If PathNameExt.Count = 3 Then
                Dim mfolder As String = PathNameExt(0)
                Dim mfilename As String = PathNameExt(1)
                Dim mext As String = PathNameExt(2)
                mfolder = IIf(mfolder = ".", Environment.CurrentDirectory, mfolder)
                mfolder = mfolder & IIf(mfolder.Trim.Length > 0, IIf(Right(mfolder.Trim, 1) = "\", "", "\"), "")
                RetStr = mfolder & mfilename & "." & mext
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.GetFullFileName ")

        End Try

        Return RetStr
    End Function
    Private Function StringToHashTable(ByVal InputString As String, Optional ByVal VarHook As String = "~", Optional ByVal ValHook As String = "=") As Hashtable
        Dim LHashTable As New Hashtable
        Try
            Dim ArrayVar() As String = InputString.Split(VarHook)
            For i = 0 To ArrayVar.Count - 1
                Dim LArray() As String = ArrayVar(i).Split(ValHook)
                If LArray.Count > 1 Then
                    LHashTable.Add(LCase(LArray(0)), LArray(1))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.StringToHashTable ")
        End Try
        Return LHashTable
    End Function
    Private Function HashTableToString(ByVal InputHashTable As Hashtable, Optional ByVal VarHook As String = "~", Optional ByVal ValHook As String = "=") As String
        Dim Lstr As String = ""
        Try

            For i = 0 To InputHashTable.Count - 1
                Lstr = Lstr & IIf(Lstr.Length = 0, "", VarHook) & InputHashTable.Keys(i) & ValHook & InputHashTable.Item(LCase(InputHashTable.Keys(i)))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.HashTableToString ")
        End Try
        Return Lstr
    End Function
    Private Function ValueFromHashTable(ByVal AHashTable As Hashtable, ByVal AKeyName As String) As String
        Dim StrVal As String = ""
        Try
            StrVal = AHashTable.Item(LCase(AKeyName))
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.ValueFromHashTable ")
        End Try
        Return StrVal
    End Function
    Private Function GetRegistryValues(ByVal RegistryFolder As String, ByVal RegKeyName As String) As String
        GetRegistryValues = ""
        Try
            Dim flag As Boolean = CreateRegistryFolder(RegistryFolder)
            If flag Then
                If Registry.LocalMachine.OpenSubKey(RegistryFolder, True).GetValue(RegKeyName) Is Nothing Then
                    Registry.LocalMachine.OpenSubKey(RegistryFolder, True).SetValue(RegKeyName, "", RegistryValueKind.String)
                End If
                GetRegistryValues = Registry.LocalMachine.OpenSubKey(RegistryFolder, True).GetValue(RegKeyName)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.GetRegistryValues ")
        End Try
        Registry.LocalMachine.Close()
    End Function
    Private Function CreateRegistryFolder(ByVal RegistryFolder As String) As Boolean
        CreateRegistryFolder = False
        Try
            If Registry.LocalMachine.OpenSubKey(RegistryFolder, True) Is Nothing Then
                Registry.LocalMachine.CreateSubKey(RegistryFolder, RegistryKeyPermissionCheck.ReadWriteSubTree)
            End If
            Registry.LocalMachine.Close()
            CreateRegistryFolder = True
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.CreateRegistryFolder ")
        End Try
    End Function
    Private Function SetRegistryValues(ByVal RegistryFolder As String, ByVal RegKeyName As String, ByVal RegKeyValue As String) As Boolean
        SetRegistryValues = False
        Try
            Dim flag As Boolean = CreateRegistryFolder(RegistryFolder)
            If flag = True Then
                Registry.LocalMachine.OpenSubKey(RegistryFolder, True).SetValue(RegKeyName, RegKeyValue, RegistryValueKind.String)
                SetRegistryValues = True
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.SetRegistryValues")
        End Try
        Registry.LocalMachine.Close()
    End Function
    Public Sub QuitMessage(ByVal MessageString As String)
        If LCase("WebAzure,WebServer,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            If MessageString.Length > 0 Then
                MsgBox(MessageString)
            End If
            System.Windows.Forms.Application.ExitThread()
            System.Windows.Forms.Application.Exit()
            Process.GetCurrentProcess.Kill()
        End If
    End Sub
    ''' <summary>
    ''' Convert Character string into ascii string
    ''' </summary>
    ''' <param name="TextString "></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TextToAscii(ByVal TextString As String) As String
        Dim Formatedvalue As String = ""
        Try
            For i = 0 To TextString.Count - 1
                Formatedvalue = Formatedvalue & Asc(TextString(i)).ToString.PadLeft(3, "0")
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.TextToAscii ")
        End Try
        Return Formatedvalue
    End Function
    ''' <summary>
    ''' Convert ascii string into character string
    ''' </summary>
    ''' <param name="AsciiString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AsciiToText(ByVal AsciiString As String) As String
        Dim Formatedvalue As String = ""
        Try
            For i = 0 To AsciiString.Count - 1
                If AsciiString.Length > 0 Then
                    Dim s As Integer = CInt(Mid(AsciiString, 1, 3))
                    Formatedvalue = Formatedvalue & Chr(s)
                    AsciiString = AsciiString.Remove(0, 3)
                Else
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.AsciiToText ")
        End Try
        Return Formatedvalue
    End Function
    ''' <summary>
    ''' To send emails with attatchments
    ''' </summary>
    ''' <param name="ToMailId">Email id to which mail to be sent</param>
    ''' <param name="MailSubject">Subject of email</param>
    ''' <param name="MyMailText">Body of email</param>
    ''' <param name="AttachmentFileName">Comma separated FileNames with path to be attatched</param>
    ''' <remarks></remarks>
    Public Sub SendingEmail(ByVal ToMailId As String, ByVal MailSubject As String, ByVal MyMailText As String, Optional ByVal AttachmentFileName As String = "")
        Try
            Dim password As String = GlobalControl.Variables.EmailPassword
            Dim FromMailId As String = GlobalControl.Variables.EmailId
            Dim SMTPServer As New SmtpClient()
            SMTPServer.Timeout = 300000
            SMTPServer.Host = GlobalControl.Variables.LocalSMTPServerHost
            SMTPServer.Port = GlobalControl.Variables.LocalSMTPServerPort
            SMTPServer.EnableSsl = GlobalControl.Variables.LocalMTPServerEnableSsl

            If LCase(GlobalControl.Variables.SaralType) = LCase("WebServer") Then
                FromMailId = GlobalControl.Variables.WebEmailId
                SMTPServer.Host = GlobalControl.Variables.WebSMTPServerHost
                password = GlobalControl.Variables.WebEmailPwd
            End If


            'From requires an instance of the MailAddress type
            Dim MyMailMessage As New MailMessage()
            MyMailMessage.From = New MailAddress(FromMailId)
            'To is a collection of MailAddress types
            MyMailMessage.To.Add(ToMailId)
            MyMailMessage.Subject = MailSubject

            'Dim tempstr As String = CreateInvoiceFile()
            MyMailMessage.Body = MyMailText
            MyMailMessage.IsBodyHtml = True

            If AttachmentFileName.Trim.Length > 0 Then
                Dim aattached() As String = AttachmentFileName.Split(",")
                For i = 0 To aattached.Count - 1
                    'Dim k0 As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(System.Web.HttpContext.Current.Server.MapPath(AttachmentFileName))
                    Dim k As New System.Net.Mail.Attachment(Trim(aattached(i)))
                    MyMailMessage.Attachments.Add(k)
                Next
            End If


            'Create the SMTPClient object and specify the SMTP GMail server
            'If LocalServerFlag = True Then
            '    SMTPServer.Host = "smtp.gmail.com"
            '    SMTPServer.Port = 587
            '    SMTPServer.EnableSsl = True
            'ElseIf LocalServerFlag = False Then
            '    SMTPServer.Host = "relay-hosting.secureserver.net"
            '    password = "123456"
            'End If
            SMTPServer.Credentials = New System.Net.NetworkCredential(FromMailId, password)
            SMTPServer.Send(MyMailMessage)
            MyMailMessage.Dispose()
        Catch ex As SmtpException
            QuitError(ex, Err, "Email not send")
        End Try
    End Sub
    Private Function GetIPAddress() As String
        Dim strHostName As String
        Dim strIPAddress As String
        strHostName = System.Net.Dns.GetHostName()
        strIPAddress = System.Net.Dns.GetHostEntry(strHostName).AddressList(1).ToString()
        ' MessageBox.Show("Host Name: " & strHostName & "; IP Address: " & strIPAddress)
        Return strIPAddress
    End Function

    Public Function AdminVault1() As String
        Try

            Dim RegStr As String = StringRead("d:\dll_folder\pc100000.xdl")
            Dim LRegStr As String = StringDecript(Mid(RegStr, 2, RegStr.Length - 1), 40)
            Dim RegHash As Hashtable = StringToHashTable(LRegStr, sep1, sep2)

            Dim hrdstr As String = StringRead("d:\edl_folder\hrdtemp.txt")
            Dim TempStr As String = StringDecript(hrdstr, 35)
            Dim TempHash As Hashtable = StringToHashTable(TempStr, sep1, sep2)

            'Dim hrdtemp As String = StringRead("d:\edl_folder\pc100000.edl")
            'Dim TempStr As String = StringDecript(mid(hrdstr, 2, hrdstr.length - 1), 30)
            'Dim TempHash As Hashtable = StringToHashTable(TempStr, sep1, sep2)


            Dim mclass As String = WmiClassProperty0(0)
            For i = 1 To WmiClassProperty0.Count - 1
                Dim mkeyname As String = LCase(mclass & "_" & WmiClassProperty0(i))
                If Not ValueFromHashTable(TempHash, mkeyname) = ValueFromHashTable(RegHash, mkeyname) Then
                    Dim tvalue As String = ValueFromHashTable(TempHash, mkeyname)
                    Dim rvalue As String = ValueFromHashTable(RegHash, mkeyname)
                    QuitMessage(mkeyname & " mismatched value (" & tvalue & " and " & rvalue & "), Not authenticated")
                    Return GlobalControl.Variables.AuthenticationChecked
                    Exit Function
                End If
            Next
            GlobalControl.Variables.AuthenticationChecked = GlobalControl.Variables.xControl + 130858
            Dim mallowdt As String = ValueFromHashTable(RegHash, LCase("allowdate"))
            If mallowdt IsNot Nothing Then
                GlobalControl.Variables.AllowDate = New Date(CInt(Left(mallowdt, 4)), CInt(Mid(mallowdt, 5, 2)), CInt(Right(mallowdt, 2)))
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute ValidateClass.AdminVault")
        End Try
        '   GlobalControl.Variables.AuthenticationChecked = VaultFlag
        Return GlobalControl.Variables.AuthenticationChecked
    End Function



End Class
