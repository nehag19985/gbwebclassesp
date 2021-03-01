Imports System.IO
Imports IWshRuntimeLibrary      'com-->windows script host object model
Imports Microsoft.Win32
Imports System.ComponentModel
Imports GlobalControl
Imports GlobalFunction1
Imports EnvDTE
Imports EnvDTE80
Imports System
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports Extensibility
Imports System.Runtime.InteropServices
Imports System.CodeDom.Compiler
Imports System.Net.Mail
Imports System.Web
Imports VslangProj90
Imports VSLangProj




Public Class GlobalFunction2
    Dim gf1 As New GlobalFunction1.GlobalFunction1
#Region "DimVariables"
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
    ''' Error Message box before quitting the application
    ''' </summary>
    ''' <param name="ex">Error on exception </param>
    ''' <param name="err"> error object </param>
    ''' <remarks></remarks>
    Public Sub QuitError(ByVal ex As Exception, ByVal err As ErrObject, ByVal ErrorString As String)
        If LCase("WebAzure,WebServer,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            MsgBox("ERROR_MESSAGE ( " & ex.Message & " )" & vbCrLf & vbCrLf & "STACK_TRACE  (" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString)
            'MsgBox(Application.ProductName)
            System.Windows.Forms.Application.ExitThread()
            System.Windows.Forms.Application.Exit()
            System.Diagnostics.Process.GetCurrentProcess.Kill()
        Else
            GlobalControl.Variables.ErrorString = "ERROR_MESSAGE ( " & ex.Message & " )" & vbCrLf & vbCrLf & "STACK_TRACE  (" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString
        End If
    End Sub
    ''' <summary>
    ''' Message box before quitting the application
    ''' </summary>
    ''' <param name="MessageString"> message as string</param>
    '''<param name="QuitProcedure" >Function or sunroutine name from exception thrown</param>
    ''' <remarks></remarks>
    Public Sub QuitMessage(ByVal MessageString As String, ByVal QuitProcedure As String)
        If LCase("WebAzure,WebServer,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            If MessageString.Length > 0 Then
                MsgBox(MessageString & vbCrLf & vbCrLf & QuitProcedure)
            End If
            System.Windows.Forms.Application.ExitThread()
            System.Windows.Forms.Application.Exit()
            System.Diagnostics.Process.GetCurrentProcess.Kill()
        End If
    End Sub

#End Region
    ''' <summary>
    ''' Function is used to edit existing Config.NT or Autoexec.NT files  separately
    ''' </summary>
    ''' <param name="FullNTFile">Name of NT file with path</param>
    ''' <param name="EditingLines">Comma separated Entries to be edited, such as files=145 or set clipper=F:145;swapk:32000,set l_n=Y,set ydi=c:\nod1 etc. </param>
    ''' <param name="RemoveFlag">True if EditingLines are to be removed from FullNTFile or False if editted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Edit_NT_Files(ByVal FullNTFile As String, ByVal EditingLines As String, Optional ByVal RemoveFlag As Boolean = False) As String



        '-------------if particular line already exist then this function will replace that line,if line does not exist then it will add that particular line at the end of the text file------'          Try
        '
        Try
            Dim AllLines As String() = EditingLines.Trim.Split(",")
            Dim fs As FileStream = New FileStream(FullNTFile, FileMode.Open, FileAccess.Read)
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(fs)
            Dim TempFileFullPath As String = GetTempFileName(FullNTFile, "_T")
            Dim sw As System.IO.StreamWriter = Nothing
            Dim fs1 As FileStream = Nothing
            fs1 = New FileStream(TempFileFullPath, FileMode.Create, FileAccess.Write)
            sw = New System.IO.StreamWriter(fs1)
            Dim FileLine As String = ""
            FileLine = sr.ReadLine()
            Dim replline As New List(Of Integer)
            Do Until FileLine Is Nothing
                Dim rplflag As Boolean = False
                For k = 0 To AllLines.Length - 1
                    Dim srchtxt As String() = AllLines(k).Trim.Split("=")
                    Dim substring As String = Microsoft.VisualBasic.Left(LCase(FileLine), srchtxt(0).Length)
                    If substring = srchtxt(0) Then
                        Dim value As String = FileLine.Replace(FileLine, AllLines(k))
                        If RemoveFlag = False Then
                            sw.WriteLine(value)
                        End If
                        replline.Add(k)
                        rplflag = True
                    End If
                Next
                If rplflag = False Then
                    sw.WriteLine(FileLine)
                End If
                FileLine = sr.ReadLine()
            Loop
            If RemoveFlag = False Then
                For k = 0 To AllLines.Length - 1
                    If replline.Contains(k) = False Then
                        sw.WriteLine(AllLines(k))
                    End If
                Next
            End If
            sw.Close()
            fs1.Close()
            sr.Close()
            fs.Close()
            System.IO.File.Delete(FullNTFile)
            System.IO.File.Move(TempFileFullPath, FullNTFile)
            Return Nothing
        Catch ex As Exception
            MsgBox("File does not exist")
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' To creates temporary file at particular path by adding any suffix in original file name-----------'
    ''' </summary>
    ''' <param name="FileFullPath"></param>
    ''' <param name="AddSuffix"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetTempFileName(ByVal FileFullPath As String, ByVal AddSuffix As String) As String
        '--------------Following function creates temporary file at particular path by adding any suffix in original file name-----------'
        Dim FileName As String = Path.GetFileName(FileFullPath)
        Dim FullPath As String = Path.GetDirectoryName(FileFullPath)
        Dim FileExtension As String = Path.GetExtension(FileFullPath)
        Dim val As String() = FileName.Split(".")
        Dim ss As String = val(0).Replace(val(0), val(0) & AddSuffix)
        Dim TempFileName As String = FullPath & "\" & ss & FileExtension
        Return TempFileName
    End Function
    '-------------Creates shortcut at desktop and Start Menu by taking icon from some Icon location.------------------'
    Public Function CreateShortcut(ByVal ExeLocation As String, ByVal ShortcutName As String, ByVal IconLoc As String)
        Dim shell As WshShell = New WshShellClass
        Dim Desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim StartUp As String = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
        Dim shortcut As WshShortcut = shell.CreateShortcut(Desktop & ShortcutName & ".lnk")
        Dim shortcut1 As WshShortcut = shell.CreateShortcut(StartUp & ShortcutName & ".lnk")
        shortcut.TargetPath = ExeLocation.Trim
        shortcut1.TargetPath = ExeLocation.Trim
        shortcut.IconLocation = IconLoc
        shortcut1.IconLocation = IconLoc
        Dim loc As String = Path.GetDirectoryName(ExeLocation).Trim
        shortcut.WorkingDirectory = loc
        shortcut1.WorkingDirectory = loc
        Registry.SetValue("HKEY_CURRENT_USER\Console\" & ShortcutName, "FullScreen", 1)
        Registry.SetValue("HKEY_CURRENT_USER\Console\" & ShortcutName, "ScreenBufferSize", 1638480)
        Registry.SetValue("HKEY_CURRENT_USER\Console\" & ShortcutName, "WindowSize", 1638480)
        shortcut.Save()
        shortcut1.Save()
        Return Nothing
    End Function
    '-----------Copy folder(containing files/folder) from Source path to destination path---------------'
    'Public Function CopyFolder(ByVal SourcePath As String, ByVal DestPath As String)
    '    If Not Directory.Exists(DestPath) Then
    '        Directory.CreateDirectory(DestPath)
    '    End If
    '    For Each fi As String In Directory.GetFiles(SourcePath)
    '        Dim dest As String = Path.Combine(DestPath, Path.GetFileName(fi))
    '        System.IO.File.Copy(fi, dest)
    '    Next
    '    For Each folder As String In Directory.GetDirectories(SourcePath)
    '        Dim dest As String = Path.Combine(DestPath, Path.GetFileName(folder))
    '        CopyFolder(folder, dest)
    '    Next
    '    Return Nothing
    'End Function
    Dim num As Integer = 1
    Dim NFolderName As String = ""
    '------------This function search a folder at particular path and if folder exist then creates a folder with other name--------'
    Public Function SearchAndCreateFolder(ByVal FolderFullPath As String) As String
        Dim DirName As String = Path.GetDirectoryName(FolderFullPath)
        Dim OFolderName As String = Path.GetFileName(FolderFullPath)
        NFolderName = OFolderName & "_old" & num.ToString.PadLeft(2, "0")
        Dim NewPath As String = ""
        NewPath = DirName & NFolderName
start:
        If Directory.Exists(NewPath) Then
            num += 1
            NFolderName = OFolderName & "_old" & num.ToString.PadLeft(2, "0")
            NewPath = DirName & NFolderName
            GoTo start
        Else
            Directory.CreateDirectory(NewPath)
        End If
        Return NewPath
    End Function
    Public Function ProcessProjectControl(ByVal ProjectVBFile As String, ByVal ProjectPath As String, ByVal FullFormNames As String, ByVal FullClassNames As String, ByVal FullDLLNames As String, ByVal ProjectType As String, ByVal BuildPath As String) As String
        Dim gf1 As New GlobalFunction1.GlobalFunction1
        Dim sw As StreamWriter
        Dim ProjectList As List(Of String) = gf1.FullFileNameToList(ProjectVBFile)
        Dim ProjectName As String = ProjectList(1)
        Dim IniDir As String = ProjectList(0)
        Dim VbProjArray As New List(Of String)
        ProjectType = LCase(ProjectType)
        Select Case ProjectType
            Case "class"
                CreateProjectFromTemplates("ClassLibrary.zip", ProjectName, ProjectPath)
            Case "form"
                CreateProjectFromTemplates("WindowsApplication.zip", ProjectName, ProjectPath)
            Case Else
                MsgBox("Invalid project type")
                Return Nothing
                Exit Function
        End Select
        Dim buildflag As Boolean = False
        VbProjArray = System.IO.File.ReadAllLines(ProjectPath & "\" & ProjectName & ".VbProj").ToList
        For i = 0 To VbProjArray.Count - 1
            Dim mline As String = VbProjArray(i).Trim
            If mline.Contains("'Release|AnyCPU'") = True And buildflag = False Then
                buildflag = True
            End If
            If Left(mline, 12) & " " = "<OutputPath> " And buildflag = True Then
                Dim mstr As String = Right(mline, mline.Length - 12).Trim
                buildflag = False
                If mstr.LastIndexOf("<") > -1 Then
                    Dim mguid As String = Left(mstr, mstr.IndexOf("<"))
                    VbProjArray(i) = VbProjArray(i).Replace(mguid, BuildPath)
                    Exit For
                End If
            End If

        Next
        For i = 0 To VbProjArray.Count - 1
            Dim mline As String = VbProjArray(i).Trim
            If mline.Contains("<OutputType>WinExe</OutputType>") = True Then
                VbProjArray(i) = VbProjArray(i).Replace("<OutputType>WinExe</OutputType>", "<OutputType>Library</OutputType>")
            End If
            If mline.Contains("<StartupObject>RCM_1.My.MyApplication</StartupObject>") = True Then
                VbProjArray(i) = VbProjArray(i).Replace("<StartupObject>RCM_1.My.MyApplication</StartupObject>", "<StartupObject></StartupObject>")
            End If
            If mline.Contains("<MyType>WindowsForms</MyType>") = True Then
                VbProjArray(i) = VbProjArray(i).Replace("<MyType>WindowsForms</MyType>", "<MyType>Windows</MyType>")
            End If
        Next

        Dim mGUID1 As String = "{" & System.Guid.NewGuid().ToString() & "}"
        For i = 0 To VbProjArray.Count - 1
            Dim mline As String = VbProjArray(i).Trim
            If Left(mline, 13) & " " = "<ProjectGuid> " Then
                Dim mstr As String = Right(mline, mline.Length - 13).Trim
                buildflag = False
                If mstr.LastIndexOf("<") > -1 Then
                    Dim mguid As String = Left(mstr, mstr.IndexOf("<"))
                    VbProjArray(i) = VbProjArray(i).Replace(mguid, mGUID1)
                    Exit For
                End If
            End If
        Next

        If ProjectType = "form" Then
            Dim ind As Integer = VbProjArray.Contains("Form1.vb")
            If ind > -1 Then
                System.IO.File.Copy(IniDir & "\" & ProjectName & ".vb", ProjectPath & "\" & ProjectName & ".vb", True)
                System.IO.File.Copy(IniDir & "\" & ProjectName & ".designer.vb", ProjectPath & "\" & ProjectName & ".designer.vb", True)
                System.IO.File.Copy(IniDir & "\" & ProjectName & ".resx", ProjectPath & "\" & ProjectName & ".resx", True)
                System.IO.File.Delete(ProjectPath & "\Form1.vb")
                System.IO.File.Delete(ProjectPath & "\Form1.designer.vb")
                System.IO.File.Delete(ProjectPath & "\Form1.resx")
                'Dim IndxC As Integer = Array.FindIndex(VbProjArray.ToArray, Function(x) x.ToString.Contains("</Compile>"))
                'VbProjArray.Insert(IndxC + 1, "<Compile Include=""" & ProjectName & ".Designer.vb"">")
                'VbProjArray.Insert(IndxC + 2, "<DependentUpon>" & ProjectName & ".vb</DependentUpon>")
                'VbProjArray.Insert(IndxC + 3, "</Compile>")
                'VbProjArray.Insert(IndxC + 4, "<Compile Include=""" & ProjectName & ".vb"">")
                'VbProjArray.Insert(IndxC + 5, "<SubType>Form</SubType>")
                'VbProjArray.Insert(IndxC + 6, "</Compile>")
                Dim IndxE As Integer = Array.FindIndex(VbProjArray.ToArray, Function(x) x.ToString.Contains("</EmbeddedResource>"))
                VbProjArray.Insert(IndxE + 1, "<EmbeddedResource Include=""" & ProjectName & ".resx"">")
                VbProjArray.Insert(IndxE + 2, "<DependentUpon>" & ProjectName & ".vb</DependentUpon>")
                VbProjArray.Insert(IndxE + 3, "<SubType>Designer</SubType>")
                VbProjArray.Insert(IndxE + 4, "</EmbeddedResource>")
                For i = 0 To VbProjArray.Count - 1
                    VbProjArray(i) = VbProjArray(i).Replace("Form1.", ProjectName & ".")
                    VbProjArray(i) = VbProjArray(i).Replace(ProjectName & ".Form1", ProjectName & "." & ProjectName)
                Next
            End If
        End If
        If ProjectType = "class" Then
            Dim ind As Integer = VbProjArray.Contains("Class1.vb")
            If ind > -1 Then
                System.IO.File.Copy(IniDir & "\" & ProjectName & ".vb", ProjectPath & "\" & ProjectName & ".vb", True)
                System.IO.File.Delete(ProjectPath & "\Class1.vb")

                For i = 0 To VbProjArray.Count - 1
                    VbProjArray(i) = VbProjArray(i).Replace("Class1.", ProjectName & ".")
                Next
            End If
        End If

        If FullFormNames.Trim.Length > 0 Then
            Dim FormStr() As String = FullFormNames.Split(",")
            For i = 0 To FormStr.Length - 1
                Dim afiles As List(Of String) = gf1.FullFileNameToList(FormStr(i))
                If Not VbProjArray.Contains(afiles(1)) Then
                    System.IO.File.Copy(afiles(0) & "\" & afiles(1) & ".vb", ProjectPath & "\" & afiles(1) & ".vb", True)
                    System.IO.File.Copy(afiles(0) & "\" & afiles(1) & ".Designer.vb", ProjectPath & "\" & afiles(1) & ".Designer.vb", True)
                    System.IO.File.Copy(afiles(0) & "\" & afiles(1) & ".resx", ProjectPath & "\" & afiles(1) & ".resx", True)
                    Dim IndxC As Integer = Array.FindIndex(VbProjArray.ToArray, Function(x) x.ToString.Contains("</Compile>"))
                    VbProjArray.Insert(IndxC + 1, "<Compile Include=""" & FormStr(i) & ".Designer.vb"">")
                    VbProjArray.Insert(IndxC + 2, "<DependentUpon>" & FormStr(i) & ".vb</DependentUpon>")
                    VbProjArray.Insert(IndxC + 3, "</Compile>")
                    VbProjArray.Insert(IndxC + 4, "<Compile Include=""" & FormStr(i) & ".vb"">")
                    VbProjArray.Insert(IndxC + 5, "<SubType>Form</SubType>")
                    VbProjArray.Insert(IndxC + 6, "</Compile>")
                    Dim IndxE As Integer = Array.FindIndex(VbProjArray.ToArray, Function(x) x.ToString.Contains("</EmbeddedResource>"))
                    VbProjArray.Insert(IndxE + 1, "<EmbeddedResource Include=""" & FormStr(i) & ".resx"">")
                    VbProjArray.Insert(IndxE + 2, "<DependentUpon>" & FormStr(i) & ".vb</DependentUpon>")
                    VbProjArray.Insert(IndxE + 3, "<SubType>Designer</SubType>")
                    VbProjArray.Insert(IndxE + 4, "</EmbeddedResource>")
                End If
            Next
        End If
        If FullDLLNames.Trim.Length > 0 Then

            Dim DllStr() As String = FullDLLNames.Split(",")
            Dim IndxD As Integer = Array.FindIndex(VbProjArray.ToArray, Function(x) x.ToString.Contains("</Reference>"))
            For i = 0 To DllStr.Length - 1
                Dim AdllStr As List(Of String) = gf1.FullFileNameToList(DllStr(i), True)
                VbProjArray.Insert(IndxD + 1, "<Reference Include=""" & AdllStr(1) & ", Version=1.0.0.0, Culture=neutral, processorArchitecture=MSIL"">")
                VbProjArray.Insert(IndxD + 2, "<SpecificVersion>False</SpecificVersion>")
                Dim mpath As String = "..\..\..\" & AdllStr(0) & "\" & AdllStr(1) & "." & AdllStr(2)
                VbProjArray.Insert(IndxD + 3, "<HintPath>" & mpath & "</HintPath>")
                VbProjArray.Insert(IndxD + 4, "</Reference>")
            Next
            Dim IndxZ As Integer
            For x = VbProjArray.Count - 1 To 0 Step -1
                If (VbProjArray(x).Equals("  </ItemGroup>")) Then
                    IndxZ = x
                    Exit For
                End If
            Next
            'VbProjArray.Insert(IndxZ + 1, "<ItemGroup>")
            'IndxZ = IndxZ + 2
            'Dim k As Integer
            'For k = 0 To DllStr.Count - 1
            '    If Not VbProjArray.Contains(DllStr(k)) Then
            '        VbProjArray.Insert(IndxZ + k, "<Content Include=""" & DllStr(k) & """/>")
            '    End If
            'Next
            'VbProjArray.Insert(IndxZ + DllStr.Count, " </ItemGroup>")
        End If
        If FullClassNames.Trim.Length > 0 Then

            Dim ClassStr() As String = FullClassNames.Split(",")
            For i = 0 To ClassStr.Count - 1
                Dim aClassStr As List(Of String) = gf1.FullFileNameToList(ClassStr(i))
                If Not VbProjArray.Contains(aClassStr(1)) Then
                    System.IO.File.Copy(aClassStr(0) & "\" & aClassStr(1) & ".vb", ProjectPath & "\" & aClassStr(1) & ".vb", True)
                    Dim IndxP As Integer = Array.FindIndex(VbProjArray.ToArray, Function(x) x.ToString.Contains("</Compile>"))
                    VbProjArray.Insert(IndxP + 1, "<Compile Include=""" & aClassStr(1) & ".vb"" />")
                End If
            Next
        End If
        sw = New StreamWriter(ProjectPath & "\" & ProjectName & ".VbProj", False, System.Text.Encoding.UTF8)
        For l As Integer = 0 To VbProjArray.Count - 1
            sw.WriteLine(VbProjArray(l))
        Next
        sw.Close()
        If ProjectType = "form" Then
            Dim mfile As String = ProjectPath & "\My Project\Application.Designer.vb"
            mfile = IO.Path.GetFullPath(mfile)
            If System.IO.File.Exists(mfile) = True Then
                Dim AppSettings As List(Of String) = System.IO.File.ReadAllLines(mfile).ToList
                'Dim omainfrm As String = "Me.MainForm = Global." & ProjectName & ".Form1"
                'Dim mmainfrm As String = "Me.MainForm = Global." & ProjectName & "." & ProjectName
                'For i = 0 To AppSettings.Count - 1
                '    Dim mline As String = AppSettings(i).Trim
                '    If mline.Contains(omainfrm) = True Then
                '        AppSettings(i) = AppSettings(i).Replace(omainfrm, mmainfrm)
                '    End If
                'Next
                sw = New StreamWriter(mfile, False, System.Text.Encoding.UTF8)
                'For l As Integer = 0 To AppSettings.Count - 1
                '    sw.WriteLine(AppSettings(l))
                'Next
                For l As Integer = 0 To 10
                    sw.WriteLine(AppSettings(l))
                Next


                sw.Close()
            End If
        End If


        Return Nothing
    End Function
    'Use templatename = "WindowsApplication.zip"
    Public Sub CreateProjectFromTemplates(ByVal TemplateName As String, ByVal ProjectName As String, ByVal ProjectPath As String)
        Try
            Dim dte As DTE = CType(Microsoft.VisualBasic.Interaction.CreateObject("VisualStudio.DTE", ""), DTE)
            Dim soln As Solution2 = CType(dte.Solution, Solution2)
            Dim vbTemplatePath As String = ""
            ' "vbproj: is the DefaultProjectExtension as seen in the registry.
            vbTemplatePath = soln.GetProjectTemplate(TemplateName, "vbproj")
            ' Create a new Visual Basic Visual Basic project using the template obtained above.
            ProjectPath = ProjectPath & "\"
            If System.IO.Directory.Exists(ProjectPath) = True And ProjectPath.Length > 3 Then
                System.IO.Directory.Delete(ProjectPath, True)
            End If
            System.IO.Directory.CreateDirectory(ProjectPath)
            'soln.AddFromTemplate(vbTemplatePath, ProjectPath, ProjectName, False)
            OleMessageFilter.Register()
            soln.AddFromTemplate(vbTemplatePath, ProjectPath, ProjectName, False)
            OleMessageFilter.Revoke()
            soln.Close()
        Catch ex As System.Exception
            gf1.QuitError(ex, Err, "CreateProjectFromTemplates(ByVal TemplateName As String, ByVal ProjectName As String, ByVal ProjectPath As String)")
        End Try
    End Sub
    Public Function GetImportsOfVB(ByVal VbFileName As String) As String()
        Dim VbLinesArray() As String = System.IO.File.ReadAllLines(VbFileName)
        Dim ImportsArray() As String = {}
        Dim gf1 As New GlobalFunction1.GlobalFunction1
        For i = 0 To VbLinesArray.Count - 1
            If Left(LCase(VbLinesArray(i)).Trim, 6) & " " = "import " Then
                Dim mstr As String = Right(VbLinesArray(i), VbLinesArray(i).Length - 7).Trim
                gf1.ArrayAppend(ImportsArray, mstr)
            End If
        Next
        Return ImportsArray
    End Function
    Public Function GetProjectProperties(ByVal VbProjFileName As String) As Hashtable
        Dim VbLinesArray() As String = System.IO.File.ReadAllLines(VbProjFileName)
        Dim ImportsArray() As String = {}
        Dim CompileArray() As String = {}
        Dim EResources() As String = {}
        Dim LResources() As String = {}
        Dim PropHash As New Hashtable
        Dim gf1 As New GlobalFunction1.GlobalFunction1
        Dim mref As Boolean = False
        Dim aRef() As String = {}
        Dim aRefdir As New Hashtable
        Dim buildflag As Boolean = False
        Dim mreference As String = ""
        For i = 0 To VbLinesArray.Count - 1
            Dim mline As String = LCase(VbLinesArray(i).Trim)
            Dim mline1 As String = VbLinesArray(i).Trim
            If Left(mline, 19) & " " = LCase("<Reference Include= ") Then
                Dim mstr As String = Right(mline1, mline1.Length - 19).Trim
                Select Case True
                    Case mstr.IndexOf(",") > -1
                        mreference = Left(mstr, mstr.IndexOf(","))
                        gf1.ArrayAppend(aRef, mreference)
                    Case mstr.IndexOf("/") > -1
                        mreference = Left(mstr, mstr.IndexOf("/"))
                        gf1.ArrayAppend(aRef, mreference)
                    Case mstr.IndexOf(">") > -1
                        mreference = Left(mstr, mstr.IndexOf(">"))
                        gf1.ArrayAppend(aRef, mreference)
                End Select
            End If
            If Left(mline, 10) & " " = LCase("<HintPath> ") And mreference.Length > 0 Then
                Dim mstr As String = Right(mline1, mline1.Length - 10).Trim
                If mstr.IndexOf("<") > -1 Then
                    Dim mdllf As String = Left(mstr, mstr.IndexOf("<"))
                    gf1.AddItemToHashTable(aRefdir, mreference, mdllf)
                    mreference = ""
                End If
            End If
            If Left(mline, 13) & " " = LCase("<ProjectGuid> ") Then
                Dim mstr As String = Right(mline1, mline1.Length - 13).Trim
                If mstr.IndexOf("<") > -1 Then
                    Dim mguid As String = Left(mstr, mstr.IndexOf("<"))
                    gf1.AddItemToHashTable(PropHash, "ProjectGuid", mguid)
                End If
            End If
            If mline.Contains(LCase("'Release|AnyCPU'")) = True And buildflag = False Then
                buildflag = True
            End If
            If Left(mline, 12) & " " = LCase("<OutputPath> ") And buildflag = True Then
                Dim mstr As String = Right(mline1, mline1.Length - 12).Trim
                ' buildflag = False
                If mstr.IndexOf("<") > -1 Then
                    Dim mguid As String = Left(mstr, mstr.IndexOf("<"))
                    gf1.AddItemToHashTable(PropHash, "BuildPath", mguid)
                End If
            End If
            If Left(mline, 8) & " " = LCase("<MyType> ") And buildflag = True Then
                Dim mstr As String = Right(mline1, mline1.Length - 8).Trim
                ' buildflag = False
                If mstr.IndexOf("<") > -1 Then
                    Dim mguid As String = Left(mstr, mstr.IndexOf("<"))
                    gf1.AddItemToHashTable(PropHash, "MyType", mguid)
                End If
            End If
            If Left(mline, 16) & " " = "<import include= " Then
                Dim mstr As String = Right(mline1, mline1.Length - 16).Trim
                Select Case True
                    Case mstr.IndexOf("/") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf("/"))
                        gf1.ArrayAppend(ImportsArray, mimp)
                    Case mstr.IndexOf(">") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf(">"))
                        gf1.ArrayAppend(ImportsArray, mimp)
                End Select
            End If
            If Left(mline, 17) & " " = LCase("<Compile Include= ") Then
                Dim mstr As String = Right(mline1, mline1.Length - 17).Trim
                Select Case True
                    Case mstr.IndexOf("/") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf("/"))
                        gf1.ArrayAppend(CompileArray, mimp)
                    Case mstr.IndexOf(">") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf(">"))
                        gf1.ArrayAppend(CompileArray, mimp)
                End Select
            End If

            If Left(mline, 26) & " " = LCase("<EmbeddedResource Include= ") Then
                Dim mstr As String = Right(mline1, mline1.Length - 26).Trim
                Select Case True
                    Case mstr.IndexOf("/") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf("/"))
                        gf1.ArrayAppend(EResources, mimp)
                    Case mstr.IndexOf(">") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf(">"))
                        gf1.ArrayAppend(EResources, mimp)
                End Select
            End If
            If Left(mline, 24) & " " = LCase("<LinkedResource Include= ") Then
                Dim mstr As String = Right(mline1, mline1.Length - 24).Trim
                Select Case True
                    Case mstr.IndexOf("/") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf("/"))
                        gf1.ArrayAppend(EResources, mimp)
                    Case mstr.IndexOf(">") > -1
                        Dim mimp As String = Left(mstr, mstr.IndexOf(">"))
                        gf1.ArrayAppend(EResources, mimp)
                End Select
            End If
        Next
        gf1.AddItemToHashTable(PropHash, "Compiles", CompileArray)
        gf1.AddItemToHashTable(PropHash, "EmbededResources", EResources)
        gf1.AddItemToHashTable(PropHash, "LinkedResources", LResources)
        gf1.AddItemToHashTable(PropHash, "Imports", ImportsArray)
        gf1.AddItemToHashTable(PropHash, "References", aRef)
        gf1.AddItemToHashTable(PropHash, "HintPath", aRefdir)
        Return PropHash
    End Function
    Public Function CreateVBSolution(ByVal VBProjFile As String, Optional ByVal ProjectType As String = "class") As String
        Dim ProjProperty As Hashtable = GetProjectProperties(VBProjFile)
        Dim VbProj As List(Of String) = gf1.FullFileNameToList(VBProjFile)
        Dim SlnFolder As String = Left(VbProj(0), VbProj(0).LastIndexOf("\"))
        Dim SlnFile As String = SlnFolder & "\" & VbProj(1) & ".sln"
        System.IO.File.Copy("D:\SaralWin\Templates\SlnTemplate.txt", SlnFile, True)
        Dim typeguid As String = ""
        If LCase(ProjectType) = "class" Then
            typeguid = "{F184B08F-C81C-45F6-A57F-5ABD9991F28F}"
        Else
            typeguid = "{F184B08F-C81C-45F6-A57F-5ABD9991F28F}"
        End If
        Dim VbLinesArray() As String = System.IO.File.ReadAllLines(SlnFile)
        'Project("ProjectTypeGUID") = "ProjectName", "ProjectFolder\ProjectName.vbproj", "ProjectGuid"
        For i = 0 To VbLinesArray.Count - 1
            If Left(VbLinesArray(i), 8) = "Project(" Then
                VbLinesArray(i) = VbLinesArray(i).Replace("ProjectTypeGUID", typeguid)
                VbLinesArray(i) = VbLinesArray(i).Replace("ProjectName", VbProj(1))
                VbLinesArray(i) = VbLinesArray(i).Replace("ProjectFolder", VbProj(1))
            End If
            VbLinesArray(i) = VbLinesArray(i).Replace("ProjectGuid", gf1.GetValueFromHashTable(ProjProperty, "ProjectGuid"))
        Next
        Dim sw As StreamWriter = New StreamWriter(SlnFile, False, System.Text.Encoding.UTF8)
        For l As Integer = 0 To VbLinesArray.Count - 1
            sw.WriteLine(VbLinesArray(l))
        Next
        sw.Close()
        Return Nothing
    End Function


    '---------------This function shows delay by making gif image visible (until loading process completes)DelayGif is a picturebox,LeftPos & TopPos are 
    '------------positions to set picturebox,frm is the form on which you want to display picturebox and Subname is name of subroutine(while processing this sub, show picturebox to cover delay)-------------'
    Dim wrkDeploy As New BackgroundWorker
    Public Delegate Sub wrk(ByVal sender As Object, ByVal e As DoWorkEventArgs)
    Public Function DelayLoop(ByVal DelayGif As PictureBox, ByVal LeftPos As Integer, ByVal TopPos As Integer, ByVal frm As Object, ByVal SubName As wrk) As Object
        frm.Controls.Add(DelayGif)
        DelayGif.Image = GlobalControl.Variables.DelayImage
        DelayGif.Left = LeftPos
        DelayGif.Top = TopPos
        System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = False
        Dim arg As DoWorkEventArgs = Nothing
        wrkDeploy.RunWorkerAsync()
        AddHandler wrkDeploy.DoWork, AddressOf SubName.Invoke
        Return Nothing
    End Function
    Public Function CreateDLL(ByVal VbProjFile As String) As String
        Dim ProjProperty As Hashtable = GetProjectProperties(VbProjFile)
        Dim mlist As List(Of String) = gf1.FullFileNameToList(VbProjFile)

        Dim RefDefaultFolder = "C:\Program Files\Reference Assemblies\Microsoft\Framework\v3.5"

        Dim Ref() As String = {}
        Dim aCompile() As String = gf1.GetValueFromHashTable(ProjProperty, "Compiles")
        Dim aEmbeded() As String = gf1.GetValueFromHashTable(ProjProperty, "EmbededResources")
        Dim aLink() As String = gf1.GetValueFromHashTable(ProjProperty, "LinkedResources")
        Dim OutputDLLPath As String = gf1.GetValueFromHashTable(ProjProperty, "BuildPath")
        If Left(OutputDLLPath, 9) = "..\..\..\" Then
            Dim mstr As String = Right(OutputDLLPath, OutputDLLPath.Length - 9)
            OutputDLLPath = "D:\saralwin\" & mstr & mlist(1).Trim & ".dll"
        End If

        Dim Aref() As String = gf1.GetValueFromHashTable(ProjProperty, "References")
        Dim RefFolder As Hashtable = gf1.GetValueFromHashTable(ProjProperty, "HintPath")
        Dim ReferenceRange() As String = {}
        Dim xfolder As String = "C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727,C:\Program Files\Reference Assemblies\Microsoft\Framework\v3.5"
        'Dim xfolder As String = "C:\WINDOWS\Microsoft.NET\Framework,C:\Program Files\Reference Assemblies\Microsoft\Framework"
        For j = 1 To Aref.Count - 1
            Dim mref As String = Aref(j)
            Dim mfolder As String = gf1.GetValueFromHashTable(RefFolder, mref)
            If mfolder = "" Then
                mref = mref.Replace("""", "").Trim & ".dll"
                Dim lfolder As List(Of String) = gf1.SearchFiles(xfolder, mref, False, True)
                mref = lfolder(0)
                mfolder = mref
                '       gf1.ArrayAppend(ReferenceRange, mref)
            End If
            If mfolder.Length = 0 Then
                mref = mref.Replace("""", "").Trim & ".dll"
                Dim lfolder As List(Of String) = gf1.SearchFiles(xfolder, mref, False, True)
                mref = lfolder(0)
                mfolder = mref
                '        gf1.ArrayAppend(ReferenceRange, mref)
            End If
            If Left(mfolder, 9) = "..\..\..\" Then
                Dim mstr As String = Right(mfolder, mfolder.Length - 9)
                mfolder = "D:\saralwin\" & mstr
                mref = mfolder.Replace("""", "").Trim
            End If
            gf1.ArrayAppend(ReferenceRange, mref)
        Next
        Dim InputVBFilePath() As String = {}
        For i = 0 To aCompile.Count - 1
            Dim mcompile As String = aCompile(i).Replace("""", "")
            If mcompile.Contains("My Project") = True Then
                ' mcompile = mcompile.Replace("My Project", "MyProj~1")
            Else
                gf1.ArrayAppend(InputVBFilePath, mlist(0) & "\" & mcompile.Trim)
            End If
        Next
        Dim AembeddResx() As String = {}
        For i = 0 To aEmbeded.Count - 1
            Dim mEmbeded As String = aEmbeded(i).Replace("""", "")
            If mEmbeded.Contains("My Project") = True Then
                ' mEmbeded = mEmbeded.Replace("My Project", "MyProj~1")
            Else
                gf1.ArrayAppend(AembeddResx, mlist(0) & "\" & mEmbeded.Trim)
            End If
        Next
        Dim ALinkResx() As String = {}
        For i = 0 To aLink.Count - 1
            Dim mlink As String = aLink(i).Replace("""", "")
            If mlink.Contains("My Project") = True Then
                ' mlink = mlink.Replace("My Project", "MyProj~1")
            Else
                gf1.ArrayAppend(ALinkResx, mlink.Trim)
            End If
        Next



        Dim strReturn As String = ""
        Try
            Dim CompilerParams = New CompilerParameters


            With CompilerParams
                .TreatWarningsAsErrors = False
                .WarningLevel = 4
                .GenerateInMemory = False
                .IncludeDebugInformation = True
                .ReferencedAssemblies.AddRange(ReferenceRange)
                .OutputAssembly = OutputDLLPath.Trim
                .CompilerOptions = "/doc"
                If AembeddResx.Count > 0 Then
                    .EmbeddedResources.AddRange(AembeddResx)
                End If
                If ALinkResx.Count > 0 Then
                    .LinkedResources.AddRange(ALinkResx)
                End If
                .MainClass = mlist(1).Trim
            End With




            Try
                Dim provOpt = New Dictionary(Of String, String)
                provOpt.Add("CompilerVersion", "v3.5")
                Dim Compiler = New VBCodeProvider(provOpt)
                Dim CompileResults = Compiler.CompileAssemblyFromFile(CompilerParams, InputVBFilePath)
                If CompileResults.Errors.HasErrors Then
                    For Each Err As CodeDom.Compiler.CompilerError In CompileResults.Errors
                        strReturn &= Err.ErrorText & " @Line: " & Err.Line & vbCrLf
                    Next
                    gf1.QuitMessage(strReturn, "CreateDll " & VbProjFile)
                End If
                Compiler.Dispose()


            Catch ex As Exception
                gf1.QuitError(ex, Err, "compileexecute")
                'Dim pdbfile As String = ""
                'Dim pdbdir = Path.GetDirectoryName(OutputDLLPath)
                'pdbfile = Path.GetFileNameWithoutExtension(OutputDLLPath)
                'File.Delete(pdbdir & "\" & pdbfile & ".PDB")
            End Try
        Catch ex As Exception
            gf1.QuitError(ex, Err, "compileparam")

        End Try
        Return Nothing
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
        'Fixed variables
        Dim password As String = GlobalControl.Variables.EmailPassword
        Dim FromMailId As String = GlobalControl.Variables.EmailId
        Dim SendGridId As String = GlobalControl.Variables.EmailId
        Dim SendGridPwd As String = GlobalControl.Variables.EmailPassword
        Dim SMTPServer As New SmtpClient()
        '   SMTPServer.UseDefaultCredentials = False
        '  SMTPServer.DeliveryMethod = SmtpDeliveryMethod.Network
        SMTPServer.Timeout = 300000
        SMTPServer.Host = GlobalControl.Variables.LocalSMTPServerHost
        SMTPServer.Port = GlobalControl.Variables.LocalSMTPServerPort
        SMTPServer.EnableSsl = GlobalControl.Variables.LocalMTPServerEnableSsl
        'If LCase(GlobalControl.Variables.SaralType) = LCase("WebGodaddy") Then
        If GlobalControl.Variables.RunningAtWeb = True Then
            If LCase(GlobalControl.Variables.SaralType) = "webgodaddy" Then
                FromMailId = "info@saralerp.com"
                SMTPServer.Timeout = 300000
                ' sg2nlvphout(-v01.shr.prod.sin2.secureserver.net)
                SMTPServer.Host = "sg2nlvphout-v01.shr.prod.sin2.secureserver.net"
                SMTPServer.Port = 25
                SMTPServer.EnableSsl = False
                SendGridId = "info@saralerp.com"
                SendGridPwd = "Saral@1234!"
            End If
        End If
        'From requires an instance of the MailAddress type
        Dim MyMailMessage As New MailMessage()
        MyMailMessage.From = New MailAddress(FromMailId, "Saral Team")
        'To is a collection of MailAddress types
        MyMailMessage.To.Add(ToMailId)
        MyMailMessage.Subject = MailSubject

        'Dim tempstr As String = CreateInvoiceFile()
        MyMailMessage.Body = MyMailText
        MyMailMessage.IsBodyHtml = True

        If AttachmentFileName.Trim.Length > 0 Then
            Dim aattached() As String = AttachmentFileName.Split(",")
            For i = 0 To aattached.Count - 1
                ' Dim k As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(System.Web.HttpContext.Current.Server.MapPath(AttachmentFileName))
                Dim k As New System.Net.Mail.Attachment(Trim(aattached(i)))
                MyMailMessage.Attachments.Add(k)
            Next
        End If
        SMTPServer.Credentials = New System.Net.NetworkCredential(SendGridId, SendGridPwd)


        If GF1.CheckInternetConnection = True Then
            Try
                SMTPServer.Send(MyMailMessage)
            Catch ex As SmtpException
                gf1.QuitError(ex, Err, "Email not send")
            End Try
        Else
            GF1.QuitMessage("No Internet connection open", "")
        End If

        MyMailMessage.Dispose()
    End Sub

    Public Sub VBIndenter(ByVal VbFileName As String)
        Dim _NameSpaceIndents As New List(Of Integer)
        Dim _classIndents As New List(Of Integer)
        Dim _moduleIndents As New List(Of Integer)
        Dim _subIndents As New List(Of Integer)
        Dim _functionIndents As New List(Of Integer)
        Dim _propertyIndents As New List(Of Integer)
        Dim _structureIndents As New List(Of Integer)
        Dim _enumIndents As New List(Of Integer)
        Dim _usingIndents As New List(Of Integer)
        Dim _withIndents As New List(Of Integer)
        Dim _ifIndents As New List(Of Integer)
        Dim _tryIndents As New List(Of Integer)
        Dim _getIndents As New List(Of Integer)
        Dim _setIndents As New List(Of Integer)
        Dim _forIndents As New List(Of Integer)
        Dim _selectIndents As New List(Of Integer)
        Dim _doIndents As New List(Of Integer)
        Dim _whileIndents As New List(Of Integer)

        Dim IndentWidth As Integer = 4
        Dim IndentChar As Char = " "c


        Dim lastLabelIndent As Integer = 0
        Dim lastRegionIndent As Integer = 0
        Dim currentIndent As Integer = 0
        Dim inProperty As Boolean = False
        Dim lineText As String
        Dim newLineIndent As Integer
        Dim lines As String() = gf1.StringReadLine(VbFileName)

        For i As Integer = 0 To lines.Count - 1
            Dim line = lines(i)

            'get the trimmed line without any comments
            lineText = LCase(StripComments(line))
            Dim LineWithoutSpace As String = lineText.Replace(" ", "")
            'only change the indent on lines that are code
            If lineText.Length > 0 Then
                'special case for regions and labels - they always have zero indent
                If lineText.StartsWith("#") Then
                    lastRegionIndent = currentIndent
                    currentIndent = 0
                ElseIf lineText.EndsWith(":") Then
                    lastLabelIndent = currentIndent
                    currentIndent = 0
                End If

                'if we are in a property and we see something 


                If (_propertyIndents.Count > 0) Then
                    If Not lineText.StartsWith("end") Then
                        If lineText.StartsWith("namespace ") OrElse SearchStatement(lineText, " namespace ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("class ") OrElse SearchStatement(lineText, " class ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("module ") OrElse SearchStatement(lineText, " module ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("sub ") OrElse SearchStatement(lineText, " sub ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("function ") OrElse SearchStatement(lineText, " function ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("property ") OrElse SearchStatement(lineText, " property ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("structure ") OrElse SearchStatement(lineText, " structure ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        ElseIf lineText.StartsWith("enum ") OrElse SearchStatement(lineText, " enum ") Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                            currentIndent -= 1
                        End If
                    Else
                        If LineWithoutSpace = "endclass" Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                        End If
                        If LineWithoutSpace = "endnamespace" Then
                            _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                        End If
                    End If
                End If

                If LineWithoutSpace = "endnamespace" Then
                    currentIndent = _NameSpaceIndents.Item(_NameSpaceIndents.Count - 1)
                    _NameSpaceIndents.RemoveAt(_NameSpaceIndents.Count - 1)
                ElseIf LineWithoutSpace = "endclass" Then
                    currentIndent = _classIndents.Item(_classIndents.Count - 1)
                    _classIndents.RemoveAt(_classIndents.Count - 1)
                ElseIf LineWithoutSpace = "endmodule" Then
                    currentIndent = _moduleIndents.Item(_moduleIndents.Count - 1)
                    _moduleIndents.RemoveAt(_moduleIndents.Count - 1)
                ElseIf LineWithoutSpace = "endsub" Then
                    currentIndent = _subIndents.Item(_subIndents.Count - 1)
                    _subIndents.RemoveAt(_subIndents.Count - 1)
                ElseIf LineWithoutSpace = "endfunction" Then
                    currentIndent = _functionIndents.Item(_functionIndents.Count - 1)
                    _functionIndents.RemoveAt(_functionIndents.Count - 1)
                ElseIf LineWithoutSpace = "endproperty" Then
                    currentIndent = _propertyIndents.Item(_propertyIndents.Count - 1)
                    _propertyIndents.RemoveAt(_propertyIndents.Count - 1)
                ElseIf LineWithoutSpace = "endtry" Then
                    currentIndent = _tryIndents.Item(_tryIndents.Count - 1)
                    _tryIndents.RemoveAt(_tryIndents.Count - 1)
                ElseIf LineWithoutSpace = "endwith" Then
                    currentIndent = _withIndents.Item(_withIndents.Count - 1)
                    _withIndents.RemoveAt(_withIndents.Count - 1)
                ElseIf LineWithoutSpace = "endget" Then
                    currentIndent = _getIndents.Item(_getIndents.Count - 1)
                    _getIndents.RemoveAt(_getIndents.Count - 1)
                ElseIf LineWithoutSpace = "endset" Then
                    currentIndent = _setIndents.Item(_setIndents.Count - 1)
                    _setIndents.RemoveAt(_setIndents.Count - 1)
                ElseIf LineWithoutSpace = "endif" Then
                    currentIndent = _ifIndents.Item(_ifIndents.Count - 1)
                    _ifIndents.RemoveAt(_ifIndents.Count - 1)
                ElseIf LineWithoutSpace = "endusing" Then
                    currentIndent = _usingIndents.Item(_usingIndents.Count - 1)
                    _usingIndents.RemoveAt(_usingIndents.Count - 1)
                ElseIf LineWithoutSpace = "endstructure" Then
                    currentIndent = _structureIndents.Item(_structureIndents.Count - 1)
                    _structureIndents.RemoveAt(_structureIndents.Count - 1)
                ElseIf LineWithoutSpace = "endselect" Then
                    currentIndent = _selectIndents.Item(_selectIndents.Count - 1)
                    _selectIndents.RemoveAt(_selectIndents.Count - 1)
                ElseIf LineWithoutSpace = "endenum" Then
                    currentIndent = _enumIndents.Item(_enumIndents.Count - 1)
                    _enumIndents.RemoveAt(_enumIndents.Count - 1)
                ElseIf LineWithoutSpace = "endwhile" OrElse lineText = "wend" Then
                    currentIndent = _whileIndents.Item(_whileIndents.Count - 1)
                    _whileIndents.RemoveAt(_whileIndents.Count - 1)
                ElseIf lineText = "next" OrElse lineText.StartsWith("next ") Then
                    currentIndent = _forIndents.Item(_forIndents.Count - 1)
                    _forIndents.RemoveAt(_forIndents.Count - 1)
                ElseIf lineText = "loop" OrElse lineText.StartsWith("loop ") Then
                    currentIndent = _doIndents.Item(_doIndents.Count - 1)
                    _doIndents.RemoveAt(_doIndents.Count - 1)
                ElseIf lineText.StartsWith("else") Then
                    currentIndent = _ifIndents.Item(_ifIndents.Count - 1)
                ElseIf lineText.StartsWith("catch") Then
                    currentIndent = _tryIndents.Item(_tryIndents.Count - 1)
                ElseIf lineText.StartsWith("case") Then
                    currentIndent = _selectIndents.Item(_selectIndents.Count - 1) + 1
                ElseIf lineText = "finally" Then
                    currentIndent = _tryIndents.Item(_tryIndents.Count - 1)
                End If

            End If

            'find the current indent
            newLineIndent = currentIndent * IndentWidth
            'change the indent of the current line 
            line = New String(IndentChar, newLineIndent) & line.TrimStart
            lines(i) = line

            If lineText.Length > 0 Then
                If lineText.StartsWith("#") Then
                    currentIndent = lastRegionIndent
                ElseIf lineText.EndsWith(":") Then
                    currentIndent = lastLabelIndent
                End If

                If Not lineText.StartsWith("end") Then
                    If (lineText.StartsWith("namespace ") OrElse SearchStatement(lineText, " namespace ")) Then
                        _NameSpaceIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("class ") OrElse SearchStatement(lineText, " class ")) Then
                        _classIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("module ") OrElse SearchStatement(lineText, " module ")) Then
                        _moduleIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("sub ") OrElse SearchStatement(lineText, " sub ")) Then
                        _subIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("function ") OrElse SearchStatement(lineText, " function ")) Then
                        _functionIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("property ") OrElse SearchStatement(lineText, " property ")) Then
                        _propertyIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("structure ") OrElse SearchStatement(lineText, " structure ")) Then
                        _structureIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf (lineText.StartsWith("enum ") OrElse SearchStatement(lineText, " enum ")) Then
                        _enumIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf SearchStatement(lineText, "using ") Then
                        _usingIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf LineWithoutSpace.StartsWith("selectcase") Then
                        _selectIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText = "try" Then
                        _tryIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText = "get" Then
                        _getIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText.StartsWith("set") AndAlso Not SearchStatement(lineText, "=") Then
                        _setIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText.StartsWith("with") Then
                        _withIndents.Add(currentIndent)
                        currentIndent += 1
                        ' ElseIf lineText.StartsWith("if") AndAlso lineText.EndsWith("then") Then
                    ElseIf lineText.StartsWith("if") Then
                        _ifIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText.StartsWith("for") Then
                        _forIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText.StartsWith("while") Then
                        _whileIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText.StartsWith("do") Then
                        _doIndents.Add(currentIndent)
                        currentIndent += 1
                    ElseIf lineText.StartsWith("case") Then
                        currentIndent += 1
                    ElseIf lineText.StartsWith("else") Then
                        currentIndent = _ifIndents.Item(_ifIndents.Count - 1) + 1
                    ElseIf lineText.StartsWith("catch") Then
                        currentIndent = _tryIndents.Item(_tryIndents.Count - 1) + 1
                    ElseIf lineText = "finally" Then
                        currentIndent = _tryIndents.Item(_tryIndents.Count - 1) + 1
                    End If
                End If
            End If
        Next
        'update the textbox
        gf1.StringWriteLine(VbFileName, lines)
    End Sub

    Private Function StripComments(ByVal code As String) As String
        If code.IndexOf("'"c) >= 0 Then
            code = code.Substring(0, code.IndexOf("'"c))
        End If
        Return code.Trim
    End Function
    Public Function SearchStatement(ByVal StatementStr As String, ByVal SearchStr As String) As Boolean
        Dim i As Int16 = StatementStr.IndexOf(SearchStr)
        Dim mflag As Boolean = False
        If i > -1 Then
            Dim astr() As Char = StatementStr.ToCharArray
            Dim n As Int16 = 0
            For k = 0 To i
                If astr(k) = """" Then
                    n = n + 1
                End If
            Next
            If n Mod 2 = 0 Then
                mflag = True
            End If
        End If
        Return mflag
    End Function
End Class
