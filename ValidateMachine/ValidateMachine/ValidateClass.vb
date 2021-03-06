﻿Imports System.Windows.Forms
Imports System.Reflection
Imports System.Management
Imports System.IO
Imports System
Imports Microsoft.Win32
Imports System.Security.Cryptography


Public Class ValidateClass
    '    Dim GF1 As New GlobalFunction1.GlobalFunction1
    'Private Shared RegistryFolder As String = GlobalClassLibrary.GlobalVarClass.RegistryFolder
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
            QuitError(ex, Err)
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
            QuitError(ex, Err)
        End Try
    End Function
    Private Function SetRegistryValues(ByVal RegistryFolder As String, ByVal RegKeyName As String, ByVal RegKeyValue As String) As Boolean
        SetRegistryValues = False
        Try
            Dim flag As Boolean = CreateRegistryFolder(RegistryFolder)
            If flag Then
                Registry.LocalMachine.OpenSubKey(RegistryFolder, True).SetValue(RegKeyName, RegKeyValue, RegistryValueKind.String)
                SetRegistryValues = True
            End If
        Catch ex As Exception
            QuitError(ex, Err)
        End Try
        Registry.LocalMachine.Close()
    End Function

    Public Function CreateHRDTemp(ByVal RegistryFolder As String) As String
        Dim finalstr As String = ""
        Try
            Dim LCurrentDate As String = DateTostringFunction(Now.Date)
            Dim KeyFolder As String = Environment.GetEnvironmentVariable("userprofile") & "\SaralKeyFolder"
            If Not System.IO.Directory.Exists(KeyFolder) Then
                System.IO.Directory.CreateDirectory(KeyFolder)
            End If
            Dim HardwareFile As String = KeyFolder & "\HRDTEMP.TXT"
            finalstr = "currdt" & sep1 & LCurrentDate & sep1 & MachRead(RegistryFolder)
            Dim finalstr1 As String = StringEncrypt(finalstr, 35)
            HardwareFile = StringWriteNew(HardwareFile, finalstr)
            Return HardwareFile
        Catch ex As Exception
            QuitError(ex, Err)
        End Try
        Return finalstr
    End Function
    Public Function PickDLLClient(ByVal RegistryFolder As String, ByVal LComputerName As String, ByVal OutPutFolder As String) As String
        PickDLLClient = ""
        Try
            Dim KeyFolder As String = Environment.GetEnvironmentVariable("userprofile") & "\SaralKeyFolder"
            If Not System.IO.Directory.Exists(KeyFolder) Then
                System.IO.Directory.CreateDirectory(KeyFolder)
            End If
            Dim ClientFile As String = OutPutFolder & IIf(Right(OutPutFolder, 1) = "\", "", "\") & PickFileName(LComputerName, 8) & ".EDL"
            Dim finalStr = "dllclient" & sep2 & LComputerName & sep1 & MachRead(RegistryFolder)
            finalStr = RemoveChar(finalStr, " ")
            '  Dim kk As String = finalStr
            'Dim kk1 As String = StringDecript(finalStr, 30)
            'MsgBox(kk & Environment.NewLine & kk1)
            finalStr = StringEncrypt(finalStr, 30)
            StringWriteNew(ClientFile, ChrW(14) + finalStr)
            PickDLLClient = ClientFile

        Catch ex As Exception
            QuitError(ex, Err)
        End Try
    End Function


    Public Function CreateDLLClient(ByVal InputClientFile As String, ByVal OutputFolder As String, Optional ByVal LDate As Date = Nothing) As String
        CreateDLLClient = ""
        Try
            If Not System.IO.File.Exists(InputClientFile) Then
                MsgBox(InputClientFile & " not found ")
                Exit Function
            End If

            If Not System.IO.Directory.Exists(OutputFolder) Then
                System.IO.Directory.CreateDirectory(OutputFolder)
            End If
            Dim ldot As Integer = InStr(InputClientFile, ".")
            Dim LPathNameExt As List(Of String) = BreakFileName(InputClientFile)
            LPathNameExt(0) = OutputFolder
            LPathNameExt(2) = "XDL"
            Dim OutputFile As String = GetFullFileName(LPathNameExt)
            Dim DLLMachine As String = StringReadNew(InputClientFile)
            If Not Left(DLLMachine, 1) = ChrW(14) Then
                MsgBox(InputClientFile & " Invalid client file")
                Exit Function
            End If
            Dim LDLLMachine As String = StringDecrypt(Mid(DLLMachine, 2, DLLMachine.Length - 1), 30)
            If Not LDate = Nothing Then
                Dim sdate As String = DateTostringFunction(LDate)
                LDLLMachine = LDLLMachine & sep1 & "AllowDate" & sep2 & sdate
            End If
            Dim finalStr As String = ChrW(15) + StringEncrypt(LDLLMachine, 40)
            CreateDLLClient = StringWriteNew(OutputFile, finalStr)
        Catch ex As Exception
            QuitError(ex, Err)
        End Try

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
            QuitError(ex, Err)
        End Try
        Return MachStr
    End Function


    Private Function StringEncrypt(ByVal InputStr As String, ByVal SeedNo As Integer) As String
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
            QuitError(ex, Err)
        End Try
        Return Ostr
    End Function

    Public Function StringEncryptWeb(ByVal InputStr As String) As String
        Dim seedno As Integer = 20
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
            QuitError(ex, Err)
        End Try
        Return Ostr
    End Function


    Private Function StringEncriptOld(ByVal InputStr As String, ByVal SeedNo As Integer) As String
        ' Dim  Method_Name As String = "StringEncriptOld",param_names() As String = {"InputStr","SeedNo"}, param_values() As Object = {InputStr,SeedNo} : rdc.StartMethod(Method_Name, param_names, param_values)
        Dim Ostr As String = "", Ls1 As String = "", Lsa1 As String = "", nlen As Integer, nrem As Integer, nquot As Integer
        '  Try
        nquot = Math.DivRem(Len(InputStr), 10, nrem)
        nlen = IIf(Len(InputStr) <= 10, 1, ((Len(InputStr) + 10 - nrem) / 10))
        For i = 0 To nlen - 1
            Ls1 = Mid(InputStr, i * 10 + 1, 10)
            For j = 0 To Len(Ls1) - 1
                Dim k As Integer = Asc(Mid(Ls1, j + 1, 1)) + (j + 1) * 5 + SeedNo
                If k >= 254 Then
                    k = 254
                End If
                Ostr = Ostr + Chr(k)
            Next
        Next
        '   Catch ex As Exception
        ' rdc.QuitError(Ex,Err,New StackTrace(True))
        '  End Try
        '   rdc.EndMethod()
        Return Ostr
    End Function




    Private Function StringDecrypt(ByVal InputStr As String, ByVal SeedNo As Integer) As String
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
            QuitError(ex, Err)
        End Try
        Return Ostr
    End Function

    Public Function StringDecryptWeb(ByVal InputStr As String) As String
        Dim seedno As Integer = 20
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
            QuitError(ex, Err)
        End Try
        Return Ostr
    End Function

    Private Function StringDecriptOld(ByVal InputStr As String, ByVal SeedNo As Integer) As String
        '  Dim  Method_Name As String = "StringDecriptOld",param_names() As String = {"InputStr","SeedNo"}, param_values() As Object = {InputStr,SeedNo} : rdc.StartMethod(Method_Name, param_names, param_values)
        Dim Ostr As String = "", Ls1 As String = "", Lsa1 As String = "", nlen As Integer, nrem As Integer, nquot As Integer
        '  Try
        nquot = Math.DivRem(Len(InputStr), 10, nrem)
        nlen = IIf(Len(InputStr) <= 10, 1, ((Len(InputStr) + 10 - nrem) / 10))
        For i = 0 To nlen - 1
            Ls1 = Mid(InputStr, i * 10 + 1, 10)
            For j = 0 To Len(Ls1) - 1
                Dim k As String = AscW(Mid(Ls1, j + 1, 1))
                If k < 254 Then
                    Ostr = Ostr + ChrW(AscW(Mid(Ls1, j + 1, 1)) - (j + 1) * 5 - SeedNo)
                Else
                    Ostr = Ostr + ChrW(254)
                End If
            Next
        Next
        ' Catch ex As Exception
        '  rdc.QuitError(ex, Err, New StackTrace(True))
        ' End Try
        '  rdc.EndMethod()
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
            ' MsgBox("Exception3")
        End Try
        Return result
    End Function



    Public Sub QuitError(ByVal ex As Exception, ByVal err As ErrObject)
        MsgBox(ex.Message & " Stacktrace " & ex.StackTrace.ToString & " Procedure " & err.Erl.ToString)
        Application.Exit()
    End Sub
    Public Function ArrayAppend(ByRef ArrayName() As Control, ByVal LastValue As Control) As Control()
        Dim ii As Integer = ArrayName.Length
        ReDim Preserve ArrayName(ii)
        ArrayName.SetValue(LastValue, ii)
        Return ArrayName
    End Function
    Public Function DateTostringFunction(ByVal InputDate As Date, Optional ByVal DateFormatString As String = "", Optional ByRef DisplayDateString As String = "") As String
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
            QuitError(ex, Err)
        End Try
    End Function
    Public Function ReverseString(ByVal InputString As String, ByVal Seedno As Integer) As String
        Dim kk As String = ""
        For i = InputString.Length To 1 Step -1 * Seedno
            kk = kk & Mid(InputString, i, Seedno)
        Next
        Return kk
    End Function


    Public Function PickFileName(ByVal LName As String, ByVal LSize As Integer) As String
        Dim mpart As String = ""
        LName = UCase(LName)
        Dim Gstring As String = "ABCDEFGHIJKLMNOPQRSTUVWXY0123456789_"
        For i = 1 To LSize
            mpart = mpart + IIf(InStr(Gstring, Mid(LName, i, 1)) > 0, Mid(LName, i, 1), "")
        Next
        mpart = Left(mpart & StrDup(LSize, "0"), LSize)
        Return mpart
    End Function
    Public Function StringWrite(ByVal TxtFile As String, ByVal TxtString As String) As String

        Dim retstr As String = ""
        Try
            Dim fsw As System.IO.FileStream
            fsw = New System.IO.FileStream(TxtFile, System.IO.FileMode.Create, IO.FileAccess.ReadWrite)
            Dim sw As New System.IO.StreamWriter(fsw, System.Text.Encoding.Default)
            sw.Write(TxtString)
            sw.Flush()
            sw.Close()
            fsw.Close()
            retstr = TxtFile
        Catch ex As Exception
            QuitError(ex, Err)
        End Try
        Return retstr
    End Function
    Public Function StringWriteNew(ByVal TxtFile As String, ByVal TxtString As String, Optional ByVal Mencoding As System.Text.Encoding = Nothing) As String

        Dim retstr As String = ""
        Try
            Dim fsw As System.IO.FileStream
            fsw = New System.IO.FileStream(TxtFile, System.IO.FileMode.Create, IO.FileAccess.ReadWrite)
            Dim sw As New System.IO.StreamWriter(fsw)
            If Mencoding IsNot Nothing Then
                sw = New System.IO.StreamWriter(fsw, Mencoding)
            End If
            sw.Write(TxtString)
            sw.Flush()
            sw.Close()
            fsw.Close()
            retstr = TxtFile
        Catch ex As Exception
            QuitError(ex, Err)
        End Try
        Return retstr
    End Function
    Public Function StringReadall(ByVal TxtFile As String) As String
        Dim outstr As String = ""
        'Try
        ' Dim fs As System.IO.FileStream = New System.IO.FileStream(TxtFile, FileMode.Open, FileAccess.ReadWrite)
        'Dim sr As System.IO.StreamReader = New System.IO.StreamReader(fs, System.Text.Encoding.UTF8)
        outstr = File.ReadAllText(TxtFile)
        ' sr = New System.IO.StreamReader(fs)
        '    Dim NBuff(0) As Char
        '    Dim ascstr As String = ""
        '    Do While sr.Peek() >= 0
        '        sr.Read(NBuff, 0, 1)
        '        If Asc(NBuff(0)) > 0 Then
        '            outstr = outstr & IIf(AscW(NBuff(0)) > 0, ChrW(AscW(NBuff(0))), "")
        '        End If
        '        ascstr = ascstr & IIf(ascstr.Length > 0, ",", "") & AscW(NBuff(0)).ToString
        '    Loop
        '    sr.Close()
        '    fs.Close()
        'Catch ex As Exception
        '    QuitError(ex, Err)
        'End Try
        Return outstr
    End Function




    Public Function StringRead(ByVal TxtFile As String) As String
        Dim outstr As String = ""
        Try
            Dim fs As System.IO.FileStream = New System.IO.FileStream(TxtFile, FileMode.Open, FileAccess.ReadWrite)
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(fs, System.Text.Encoding.Default)

            ' sr = New System.IO.StreamReader(fs)
            Dim NBuff(0) As Char
            Dim ascstr As String = ""
            Do While sr.Peek() >= 0
                sr.Read(NBuff, 0, 1)
                If Asc(NBuff(0)) > 0 Then
                    outstr = outstr & IIf(Asc(NBuff(0)) > 0, ChrW(Asc(NBuff(0))), "")
                End If
                ascstr = ascstr & IIf(ascstr.Length > 0, ",", "") & Asc(NBuff(0)).ToString
            Loop
            sr.Close()
            fs.Close()
        Catch ex As Exception
            QuitError(ex, Err)
        End Try
        Return outstr
    End Function
    Public Function StringReadNew(ByVal TxtFile As String) As String
        Dim outstr As String = ""
        Try
            Dim fs As System.IO.FileStream = New System.IO.FileStream(TxtFile, FileMode.Open, FileAccess.ReadWrite)
            Dim sr As System.IO.StreamReader = New System.IO.StreamReader(fs)
            ' sr = New System.IO.StreamReader(fs)
            Dim NBuff(0) As Char
            Dim ascstr As String = ""
            Do While sr.Peek() >= 0
                sr.Read(NBuff, 0, 1)
                If AscW(NBuff(0)) > 0 Then
                    outstr = outstr & IIf(AscW(NBuff(0)) > 0, ChrW(AscW(NBuff(0))), "")
                End If
                ascstr = ascstr & IIf(ascstr.Length > 0, ",", "") & AscW(NBuff(0)).ToString
            Loop
            sr.Close()
            fs.Close()
        Catch ex As Exception
            QuitError(ex, Err)
        End Try
        Return outstr
    End Function


    Public Function SearchFiles(ByVal SourcePath As String, Optional ByVal WildCard As String = "*.*", Optional ByVal TopLevel As Boolean = True) As List(Of String)
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
            QuitError(ex, Err)
        End Try

        Return filelist
    End Function
    Public Function RemoveChar(ByVal InputString As String, Optional ByVal RemChar As String = "") As String
        'to remove letters
        Dim NewString As String = ""
        For i = 1 To InputString.Length
            If AscW(Mid(InputString, i, 1)) > 0 Then
                If (Not Mid(InputString, i, 1) = RemChar) Or RemChar.Length = 0 Then
                    NewString = NewString & Mid(InputString, i, 1)
                End If
            End If
        Next
        Return NewString
    End Function
    Public Function BreakFileName(ByVal FullFileName As String) As List(Of String)
        '0=patth,1=filename,2=extension
        Dim mpath As String = ""
        Dim mfilename As String = ""
        Dim mext As String = ""
        Dim mpathnameext As New List(Of String)
        Dim dd As Integer, ee As Integer
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
        Return mpathnameext
    End Function
    Public Function GetFullFileName(ByVal PathNameExt As List(Of String)) As String
        '0=patth,1=filename,2=extension
        Dim RetStr As String = ""
        If PathNameExt.Count = 3 Then
            Dim mfolder As String = PathNameExt(0)
            Dim mfilename As String = PathNameExt(1)
            Dim mext As String = PathNameExt(2)
            mfolder = IIf(mfolder = ".", Environment.CurrentDirectory, mfolder)
            mfolder = mfolder & IIf(mfolder.Trim.Length > 0, IIf(Right(mfolder.Trim, 1) = "\", "", "\"), "")
            RetStr = mfolder & mfilename & "." & mext
        End If
        Return RetStr
    End Function
    Public Sub VisibleControls(ByRef FormName As Object, ByVal ControlNames As String, Optional ByVal ExceptControls As String = "", Optional ByVal VisibleTrue As Boolean = True, Optional ByVal OnlyTopLevel As Boolean = True)
        '* for all controls
        If ControlNames = "*" Then
            ControlNames = GetAllControlNames(FormName)
        End If
        Dim Acontrol() As String = LCase(ControlNames).Split(",")
        Dim Econtrol() As String = Lcase(ExceptControls).Split(",")
        For i = 0 To Acontrol.Count - 1
            '  MsgBox(Array.IndexOf(Econtrol, Acontrol(i)).ToString)

            If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                Try
                    lcontrol.visible = VisibleTrue
                Catch ex As Exception
                    Continue For
                End Try
            End If
        Next
    End Sub

    Public Function GetAllControlNames(ByVal lform As Object, Optional ByVal Sep0 As String = ",", Optional ByVal OnlyTopLevel As Boolean = False) As String
        Dim ChildArr() As Control = {}
        Dim result As String = ""
        For Each ctrl1 As Control In lform.Controls
            result = result & IIf(result.Length = 0, "", Sep0) & ctrl1.Name
            If OnlyTopLevel = False Then
                If ctrl1.HasChildren Then '===========preserve the controls on form, having child controls=== 
                    ArrayAppend(ChildArr, ctrl1)
                End If
            End If
        Next

        '===========checking for extended controls or inner controls===========
sloop:
        '----copying parent controls in an array to check for child controls of each of them.
        Dim ExtChildArr() As Control = ChildArr.Clone
        If ChildArr.Length = 0 Then
            GoTo rloop
        End If
        Array.Clear(ChildArr, 0, 0)
        Array.Resize(ChildArr, 0)
        For i = 0 To ExtChildArr.Length - 1
            For Each ctrl2 As Control In ExtChildArr(i).Controls '--------testing inner controls--------
                result = result & IIf(result.Length = 0, "", Sep0) & ctrl2.Name
                If OnlyTopLevel = False Then
                    If ctrl2.HasChildren Then
                        ArrayAppend(ChildArr, ctrl2)
                    End If
                End If
            Next
        Next
        GoTo sloop
rloop:
        Return result
    End Function

    Public Function ControlNameToObject(ByVal lform As Object, ByVal ControlName As String, Optional ByRef TypeOfControlName As String = "") As Object
        Dim ChildArr() As Control = {}
        Dim Val As Object = Nothing
        If LCase("me") = LCase(Trim(ControlName)) Then
            'TypeOfControlName = "F"
            Return lform.Name   '===========returns form====
        End If
        For Each ctrl1 As Control In lform.Controls
            If LCase(ctrl1.Name) = LCase(Trim(ControlName)) Then
                Val = ctrl1
                'TypeOfControlName = "C"
                GoTo rloop     '============returns control on form=============
            End If
            If ctrl1.HasChildren Then '===========preserve the controls on form, having child controls=== 
                ArrayAppend(ChildArr, ctrl1)
            End If
        Next
        '===========checking for extended controls or inner controls===========
sloop:
        '----copying parent controls in an array to check for child controls of each of them.
        Dim ExtChildArr() As Control = ChildArr.Clone
        If ChildArr.Length = 0 Then
            GoTo rloop
        End If
        Array.Clear(ChildArr, 0, 0)
        Array.Resize(ChildArr, 0)
        For i = 0 To ExtChildArr.Length - 1
            For Each ctrl2 As Control In ExtChildArr(i).Controls '--------testing inner controls--------
                If LCase(ctrl2.Name) = LCase(Trim(ControlName)) Then
                    Val = ctrl2
                    'TypeOfControlName = "C"
                    GoTo rloop    '=====returns controls present in control on form=======
                End If
                If ctrl2.HasChildren Then
                    ArrayAppend(ChildArr, ctrl2)
                End If
            Next
        Next
        GoTo sloop
rloop:
        If Val Is Nothing Then
            Dim myFieldInfo() As FieldInfo
            '------------- Get the type and fields of FieldInfoClass-------------
            myFieldInfo = lform.GetType.GetFields(BindingFlags.NonPublic Or BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)
            For i = 0 To myFieldInfo.Length - 1  '------- Display the field information of FieldInfoClass----
                If LCase(myFieldInfo(i).Name) = LCase(Trim(ControlName)) Then
                    'TypeOfControlName = "V"
                    Return myFieldInfo(i)    '========== returns variable present on form ========
                End If
            Next
        End If
        If Val Is Nothing Then
            '===========checking for a external variable i.e. variable on another class==============
            Dim pp As String = Application.ProductName & "." & ControlName
            Dim ControlName_1 As Object = Activator.CreateInstance(Type.GetType(pp)) '----to obtain instance of that type--
            'TypeOfControlName = "E"
            Return ControlName_1
        End If
        Return Val
    End Function
    Public Function StringToHashTable(ByVal InputString As String, Optional ByVal VarHook As String = "~", Optional ByVal ValHook As String = "=") As Hashtable
        Dim ArrayVar() As String = InputString.Split(VarHook)
        Dim LHashTable As New Hashtable
        For i = 0 To ArrayVar.Count - 1
            Dim LArray() As String = ArrayVar(i).Split(ValHook)
            LHashTable.Add(LCase(LArray(0)), LArray(1))
        Next
        Return LHashTable
    End Function
    Public Function HashTableToString(ByVal InputHashTable As Hashtable, Optional ByVal VarHook As String = "~", Optional ByVal ValHook As String = "=") As String
        Dim Lstr As String = ""
        For i = 0 To InputHashTable.Count - 1
            Lstr = Lstr & IIf(Lstr.Length = 0, "", VarHook) & InputHashTable.Keys(i) & ValHook & InputHashTable.Item(LCase(InputHashTable.Keys(i)))
        Next
        Return Lstr
    End Function
    Public Function ValueFromHashTable(ByVal AHashTable As Hashtable, ByVal AKeyName As String) As String
        Dim StrVal As String = AHashTable.Item(LCase(AKeyName))
        Return StrVal
    End Function
    Public Sub QuitMessage(ByVal MessageString As String)
        If MessageString.Length > 0 Then
            MsgBox(MessageString)
        End If
        Application.Exit()
    End Sub

    ''' <summary>
    ''' Convert Character string into ascii string
    ''' </summary>
    ''' <param name="TextString "></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TextToAscii(ByVal TextString As String) As String
        Dim Formatedvalue As String = ""
        For i = 0 To TextString.Count - 1
            Formatedvalue = Formatedvalue & Asc(TextString(i)).ToString.PadLeft(3, "0")
        Next
        Return Formatedvalue
    End Function
    ''' <summary>
    ''' Convert Ascii string into character string
    ''' </summary>
    ''' <param name="AsciiString"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AsciiToText(ByVal AsciiString As String) As String
        Dim Formatedvalue As String = ""
        For i = 0 To AsciiString.Count - 1
            If AsciiString.Length > 0 Then
                Dim s As Integer = CInt(Mid(AsciiString, 1, 3))
                Formatedvalue = Formatedvalue & Chr(s)
                AsciiString = AsciiString.Remove(0, 3)
            Else
                Exit For
            End If
        Next
        Return Formatedvalue
    End Function
    ''' <summary>
    ''' Read Client EPL/EGR/EAA files and return hashtable
    ''' </summary>
    ''' <param name="RegFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetHashTableFromClientReg(ByVal RegFileName As String) As Hashtable
        Dim mhash As New Hashtable
        '  Try

        '   Dim mstring As String = StringRead(RegFileName)
        Dim mstring As String = StringRead(RegFileName)
        '  mstring = ConvertUniCodeToAscii(mstring)
        If Microsoft.VisualBasic.Left(mstring, 1) = ChrW(14) Then
            mstring = Microsoft.VisualBasic.Mid(mstring, 2, mstring.Length - 1)
        Else

            QuitMessage("Invalid client registration file " & RegFileName) ', "", New StackTrace(True))
        End If
        mstring = StringDecriptOld(mstring, 30)
        Dim astring As String() = mstring.Split("$")
        Dim aHeader() As String = astring(0).Split("!")
        Dim adatarow() As String = astring(1).Split(ChrW(24))
        Dim mTable As New DataTable
        For i = 0 To aHeader.Length - 1
            Dim afield() As String = aHeader(i).Split(ChrW(23))
            mTable.Columns.Add(afield(0))
        Next
        mTable.Columns.Add("MachineLine")
        mTable.Columns.Add("FolderStamps")
        mTable.Columns.Add("Pcode")
        mTable.Columns.Add("RegDate")
        mTable.Columns.Add("CurrentCeilingDate")
        mTable.Columns.Add("PreviousCeilingDate")
        mTable.Columns.Add("SaralHybrid")
        'mTable.Columns.Add("Nodes")
        'mTable.Columns.Add("GSTFlag")
        'mTable.Columns.Add("ImportSalesFlag")
        Dim mrow As DataRow = mTable.NewRow
        Dim mcolval() As String = adatarow(0).Split(ChrW(23))
        For i = 0 To mcolval.Count - 1
            mrow(i) = mcolval(i)
        Next
        Dim mline1 As String = IIf(astring.Count > 1, astring(2), "")
        Dim mline2 As String = IIf(astring.Count > 2, astring(3), "")
        Dim aline2() As String = mline2.Split("~")
        'dateline:=md1+[~]+md2+[~]+md3+[~]+md4+[~]+md5+[~]+md6+[~]+[SaralHybrid]+[~]+mgstin
        Dim SaralHybrid As String = ""
        If aline2.Length > 6 Then
            SaralHybrid = "Y"
        End If
        mrow("MachineLine") = mline1
        mrow("FolderStamps") = mline2
        mrow("Pcode") = ""
        mrow("SaralHybrid") = SaralHybrid
        mTable.Rows.Add(mrow)
        For i = 0 To mTable.Columns.Count - 1
            mhash.Add(LCase(mTable.Columns(i).ColumnName), mTable.Rows(0)(i))
        Next
        '  rdc.EndMethod()
        Return mhash

        Return mhash
    End Function
    ''' <summary>
    ''' Read Client EPL/EGR/EAA files and return hashtable
    ''' </summary>
    ''' <param name="RegFileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetHashTableFromXReg(ByVal RegFileName As String) As Hashtable
        Dim mhash As New Hashtable
        Dim mstring As String = StringRead(RegFileName)
        '  mstring = ConvertUniCodeToAscii(mstring)
        If Microsoft.VisualBasic.Left(mstring, 1) = ChrW(15) Then
            mstring = Microsoft.VisualBasic.Mid(mstring, 2, mstring.Length - 1)
        Else
            QuitMessage("Invalid client registration file " & RegFileName)
        End If
        '   Dim mstring As String = mFolderStamp & "$" & mLan & "~" & mNodes.ToString & "~" & mL_Date & "~" & mPcode & "~" & mHospital & "~" & mGst & "~" & mSaleImp & "~" & mclient & "~" & mcity & "$" & machineline
        mstring = StringDecriptOld(mstring, 35)
        Dim astring As String() = mstring.Split("$")
        Dim mline1 As String = IIf(astring.Count > 0, astring(1), "")
        Dim mline2 As String = IIf(astring.Count > 1, astring(2), "")
        mhash.Add(LCase("FolderStamp"), astring(0))
        mhash.Add(LCase("machineline"), astring(2))
        Dim bstring() As String = Split(astring(1), "~")
        mhash.Add(LCase("Lan"), bstring(0))
        mhash.Add(LCase("Nodes"), CInt(bstring(1)))
        mhash.Add(LCase("L_Date"), bstring(2))
        mhash.Add(LCase("Pcode"), bstring(3))
        mhash.Add(LCase("Hospital"), bstring(4))
        mhash.Add(LCase("GST"), bstring(5))
        mhash.Add(LCase("SaleImp"), bstring(6))
  Dim cstring() As String = Split(astring(3), "~")
        mhash.Add(LCase("Client"), bstring(7))
        mhash.Add(LCase("City"), bstring(8))
  mhash.Add(LCase("Mobile"), bstring(9))



        Return mhash
    End Function



    ''' <summary>
    ''' Create Client registration file.
    ''' </summary>
    ''' <param name="ClientRegHashTable" ></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function CreateRegFileFromHashTable_new(ByVal ClientRegHashTable As Hashtable, ByVal mFolder As String, ByVal regfileName As String) As String
        Try
            '   Dim Method_Name As String = "CreateRegFileFromHashTable", param_names() As String = {"ClientRegHashTable", "mFolder", "regfileName"}, param_values() As Object = {ClientRegHashTable, mFolder, regfileName} : rdc.StartMethod(Method_Name, param_names, param_values)
            Dim machineline As String = ClientRegHashTable.Item(LCase("MachineLine"))
            Dim mFolderStamp As String = ClientRegHashTable.Item(LCase("FolderStamps"))
            Dim mLan As String = ClientRegHashTable.Item(LCase("LAN"))
            Dim mNodes As Int16 = ClientRegHashTable.Item(LCase("Nodes"))
            Dim mPcode As String = ClientRegHashTable.Item(LCase("Pcode"))
            Dim mGst As String = ClientRegHashTable.Item(LCase("gst"))
            Dim mHospital As String = ClientRegHashTable.Item(LCase("Hospital"))
            Dim mSaleImp As String = ClientRegHashTable.Item(LCase("Sale_Imp"))
            Dim mclient As String = ClientRegHashTable.Item(LCase("clname"))
            Dim mcity As String = ClientRegHashTable.Item(LCase("city"))
            Dim mmobile As String = ClientRegHashTable.Item(LCase("mobile"))
            Dim mcustcode As String = ClientRegHashTable.Item(LCase("custcode"))
            Dim mSvol As String = ClientRegHashTable.Item(LCase("svol"))
            Dim mLastDate As DateTime = ClientRegHashTable.Item(LCase("L_date"))
            Dim mL_Date As String = mLastDate.ToString("yyyyMMdd")
            Dim mExt As String = IIf(mSvol = "0", "XPL", IIf(mSvol = "1", "XAA", "XGR"))

            ' Dim mSize As Int16 = IIf(CInt(mPcode) > 9999, 3, 4)
            Dim mRegFile As String = mFolder & regfileName & "." & mExt
            mGst = IIf(mGst = "1", "2", mGst)
            mSaleImp = IIf(mSaleImp = "1", "2", mSaleImp)
            Dim mstring As String = mFolderStamp & "$" & mLan & "~" & mNodes.ToString & "~" & mL_Date & "~" & mPcode & "~" & mHospital & "~" & mGst & "~" & mSaleImp & "$" & machineline & "$" & mclient & "~" & mcity & "~" & mmobile & "~" & mcustcode
            Dim SeedNo As Integer = GetSeedNo() + 25
            Dim regstring1 As String = StringEncriptOld(mstring, SeedNo)
            Dim regstring As String = ChrW(20) & regstring1
            regstring = regstring.Insert(50, ChrW(200 + SeedNo))
            ' mRegFile = StringWrite("d:\temp3.txt", regasc)
            mRegFile = StringWrite(mRegFile, regstring)
            '  rdc.EndMethod()
            Return mRegFile
        Catch Ex As Exception
            ' rdc.QuitError(Ex, Err, New StackTrace(True))
        End Try
        '  rdc.EndMethod()
        Return Nothing
    End Function
    Private Function GetSeedNo() As Int16
        Dim nn As Integer = 0
        Dim mtime As String = Now().ToString
        For h = 1 To mtime.Count - 1
            nn = nn + Asc(Microsoft.VisualBasic.Mid(mtime, h, 1)) + h * h
        Next
        While nn > 9
            Dim snn As String = LTrim(nn.ToString)
            nn = 0
            For h = 1 To Len(snn)
                nn = nn + Val(Microsoft.VisualBasic.Mid(snn, h, 1))
            Next
        End While
        If nn = 0 Then
            nn = 10
        End If
        Return nn

    End Function











    ''' <summary>
    ''' Create Client registration file.
    ''' </summary>
    ''' <param name="ClientRegHashTable" ></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function CreateRegFileFromHashTable(ByVal ClientRegHashTable As Hashtable, ByVal mFolder As String, ByVal regfileName As String) As String
        Dim machineline As String = ClientRegHashTable.Item(LCase("MachineLine"))
        Dim mFolderStamp As String = ClientRegHashTable.Item(LCase("FolderStamps"))
        Dim mLan As String = ClientRegHashTable.Item(LCase("LAN"))
        Dim mNodes As Int16 = ClientRegHashTable.Item(LCase("Nodes"))
        Dim mPcode As String = ClientRegHashTable.Item(LCase("Pcode"))
        Dim mGst As String = ClientRegHashTable.Item(LCase("gst"))
        Dim mHospital As String = ClientRegHashTable.Item(LCase("Hospital"))
        Dim mSaleImp As String = ClientRegHashTable.Item(LCase("Sale_Imp"))
        Dim mclient As String = ClientRegHashTable.Item(LCase("clname"))
        Dim mcity As String = ClientRegHashTable.Item(LCase("city"))
        Dim mmobile As String = ClientRegHashTable.Item(LCase("mobile"))
        Dim mcustcode As String = ClientRegHashTable.Item(LCase("custcode"))
        Dim mSvol As String = ClientRegHashTable.Item(LCase("svol"))
        Dim mLastDate As DateTime = ClientRegHashTable.Item(LCase("L_date"))
        Dim mL_Date As String = mLastDate.ToString("yyyyMMdd")
        Dim mExt As String = IIf(mSvol = "0", "XPL", IIf(mSvol = "1", "XAA", "XGR"))
        ' Dim mSize As Int16 = IIf(CInt(mPcode) > 9999, 3, 4)
        Dim mRegFile As String = mFolder & regfileName & "." & mExt
        mGst = IIf(mGst = "1", "2", mGst)
        mSaleImp = IIf(mSaleImp = "1", "2", mSaleImp)
        Dim mstring As String = mFolderStamp & "$" & mLan & "~" & mNodes.ToString & "~" & mL_Date & "~" & mPcode & "~" & mHospital & "~" & mGst & "~" & mSaleImp & "$" & machineline & "$" & mclient & "~" & mcity & "~" & mmobile & "~" & mcustcode
        Dim regstring1 As String = StringEncriptOld(mstring, 35)
        Dim regasc As String = ""
        For i = 0 To regstring1.Length - 1
            regasc = regasc & IIf(regasc.Length = 0, "", ",") & Asc(Mid(regstring1, i + 1, 1).ToString)
        Next
        Dim regstring As String = ChrW(15) & regstring1
        ' mRegFile = StringWrite("d:\temp3.txt", regasc)
        mRegFile = StringWrite(mRegFile, regstring)
        Return mRegFile
    End Function

    Public Function GetName(ByVal ClName As String, ByVal PickSize As Int16) As String
        Dim mfname As String = ""
        For i = 1 To ClName.Length
            Dim mchar As String = Microsoft.VisualBasic.Mid(ClName, i, 1)
            If InStr("ABCDEFGHIJKLMNOPQRSTUVWXY0123456789_", mchar) > 0 Then
                mfname = mfname + mchar
                If mfname.Length = PickSize Then
                    Exit For
                End If
            End If
        Next
        mfname = Left(mfname + Microsoft.VisualBasic.StrDup(PickSize, "0"), PickSize)
        Return mfname
    End Function
    Public Function ConvertUniCodeToAscii(ByVal InputString As String) As String
        Dim ascii As System.Text.Encoding = System.Text.Encoding.ASCII
        Dim unicode0 As System.Text.Encoding = System.Text.Encoding.Unicode
        Dim unicodeBytes As Byte() = unicode0.GetBytes(InputString)

        ' Perform the conversion from one encoding to the other.
        Dim asciiBytes As Byte() = System.Text.Encoding.Convert(unicode0, ascii, unicodeBytes)
        Dim NewString As String = System.Text.Encoding.ASCII.GetString(asciiBytes)
        Return NewString
    End Function



End Class
