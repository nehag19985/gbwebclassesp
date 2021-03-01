Imports System.IO

Public Class CreateProjectClass
    'Dim df1 As New DataFunctions.DataFunctions
    ''' <summary>
    ''' To create the project.
    ''' </summary>
    ''' <param name="SubFolder"></param>Represent the Folder in which project is to be saved
    ''' <param name="ProjectType"></param>Represent the type of Project i.e. windows application,control library etc.
    ''' <param name="ProjectName"></param>Represent the name of the project
    ''' <param name="ProjectTypeID"></param>Represent the type.i.e.it is of visual studio 8 or visual studio 15 project
    Public Sub CreateProject(ByVal SubFolder As String, ByVal ProjectType As String, ByVal ProjectName As String, ByVal ProjectTypeID As String)
        'To create project
        Dim type As System.Type = System.Type.GetTypeFromProgID(ProjectTypeID)
        'for visual studio 8 ,in gettypefromprogid - "VisualStudio.DTE.9.0"-this would be given in arugemnt.
        Dim obj As Object = System.Activator.CreateInstance(type, True)
        Dim DTE2 As EnvDTE80.DTE2 = CType(obj, EnvDTE80.DTE2)
        Try
            Dim soln As EnvDTE80.Solution2 = CType(DTE2.Solution, EnvDTE80.Solution2)
            Dim vbTemplatePath As String
            Dim vbPrjPath As String = SubFolder
            vbTemplatePath = soln.GetProjectTemplate _
            ("" & ProjectType & ".zip", "vbproj")
            soln.AddFromTemplate(vbTemplatePath, vbPrjPath, "" & ProjectName & "", False)
        Catch ex As System.Exception
            MsgBox("ERROR: " & ex.ToString)
            'End
        End Try
    End Sub
    ''' <summary>
    ''' To create Windows application project
    ''' </summary>
    ''' <param name="ProjectPath"></param>The path of the project to be saved..eg..E:\Newfolder
    ''' <param name="ProjectName"></param>The name given to the project.
    ''' <param name="FormName"></param>It represent comma separated values of path of form file with extension...eg..E:\Testing\WindowsApp\WindowsApp\Form1.vb
    ''' <param name="ClassName"></param>It represent comma separated values of path of class file with extension...eg..E:\Testing\WindowsApp\WindowsApp\Class1.vb
    ''' <param name="DllName"></param>It represent comma separated values of path of dll file with extension...eg..E:\Testing\WindowsApp\WindowsApp\Golbal1.dll
    ''' <param name="ProjectTypeID"></param>It represent the projectId i.e. For project 2008-"VisualStudio.DTE.9.0" and for project 2015-"VisualStudio.DTE.14.0"
    ''' <param name="BuildPath"></param>It represent the path where build dll have to be saved
    Public Sub CreateWindowsApplicationProject(ByVal ProjectPath As String, ByVal ProjectName As String, ByVal FormName As String, ByVal ClassName As String, ByVal DllName As String, ByVal ProjectTypeID As String, ByVal BuildPath As String)
        Dim MainFolder As String = "" & ProjectPath & "\" & ProjectName & ""
        Dim SubFolder As String = "" & MainFolder & "\" & ProjectName & ""

        If (Not Directory.Exists(MainFolder)) Then
            Directory.CreateDirectory(MainFolder)
            'To create project
            CreateProject(SubFolder, "WindowsApplication", ProjectName, ProjectTypeID)

            'Delete the class1.vb
            File.Delete("" & SubFolder & "\Form1.Designer.vb")
            File.Delete("" & SubFolder & "\Form1.vb")

            'For sln file
            Dim UniqueId As String = GetUniqueId("" & SubFolder & "\" & ProjectName & ".vbproj")
            Select Case ProjectTypeID
                Case "VisualStudio.DTE.14.0"
                    CreateSlNFile15(UniqueId, ProjectName, MainFolder)
                Case "VisualStudio.DTE.9.0"
                    CreateSlNFile08(UniqueId, ProjectName, MainFolder)
            End Select

            'For Forms
            Dim FilesNameValue() As String = {}
            Dim FilesPath() As String = {}
            If Not FormName.Trim = "" Then
                FilesPath = GetSpitString(FormName)
                FilesNameValue = GetFileName(FilesPath)
                For i = 0 To FilesPath.Length - 1
                    Dim FormPathValues As String() = GetValueArrayfromValuePath(FilesPath(i), "Form")
                    If FormPathValues.Length > 0 Then
                        CopyFilefromOnepathtoOther(FormPathValues, SubFolder)
                    End If
                Next
            End If
            'for Clas
            Dim ClassNameValue As String() = {}
            Dim ClassPath As String() = {}
            If Not ClassName.Trim = "" Then
                ClassPath = GetSpitString(ClassName)
                ClassNameValue = GetFileName(ClassPath)
                For i = 0 To ClassPath.Length - 1
                    Dim ClassPathValues As String() = GetValueArrayfromValuePath(ClassPath(i), "Class")
                    If ClassPathValues.Length > 0 Then
                        CopyFilefromOnepathtoOther(ClassPathValues, SubFolder)
                    End If
                Next
            End If
            'for Dll
            Dim DllNameValue As String() = {}
            Dim DllPath As String() = {}
            If Not DllName.Trim = "" Then
                DllPath = GetSpitString(DllName)
                DllNameValue = GetFileName(DllPath)
                For i = 0 To DllPath.Length - 1
                    Dim DllPathValues As String() = GetValueArrayfromValuePath(DllPath(i), "Dll")
                    If DllPathValues.Length > 0 Then
                        CopyFilefromOnepathtoOther(DllPathValues, "" & SubFolder & "\bin\Debug")
                    End If
                Next
            End If
            'for vbproj

            'for form vbproj code
            Dim VbProjCodeFormlist As New List(Of String)
            VbProjCodeFormlist = CreateVbprojFormCode(FilesNameValue)
            'for Class vbproj code
            Dim VbProjCodeClasslist As New List(Of String)
            VbProjCodeClasslist = CreateVbprojClassCode(ClassNameValue)
            'To combine form and class code
            Dim VbProjCodelist As New List(Of String)(VbProjCodeClasslist.Concat(VbProjCodeFormlist))
            'for resx vbproj code
            Dim VbProjCodeResxlist As New List(Of String)
            VbProjCodeResxlist = CreateVbprojResxCode(FilesNameValue)
            'for dll vbproj code
            Dim VbProjCodeDlllist As New List(Of String)
            VbProjCodeDlllist = CreateVbprojDllCode(DllNameValue, DllPath)

            'toget vbprojfile as list of string
            Dim Vbprojpath As String = "" & SubFolder & "\" & ProjectName & ".vbproj"
            AddCodeinVbprojWindowsApp(Vbprojpath, VbProjCodelist, VbProjCodeResxlist, VbProjCodeDlllist, BuildPath)

            'To delete the Form1 code in vbproj file
            DeleteCodeinVbprojWindowsApp(Vbprojpath)

            'To Set App Dsigner Value
            If FilesNameValue.Length > 0 Then
                Dim AppPath As String = "" & SubFolder & "\My Project\Application.Designer.vb"
                ChangeAppDesignerVal(FilesNameValue(0), AppPath)
            End If
        Else
            'for upadte or inculde
            'For Forms
            Dim FilesNameValue() As String = {}
            Dim FilesPath() As String = {}
            If Not FormName.Trim = "" Then
                FilesPath = GetSpitString(FormName)
                FilesNameValue = GetFileName(FilesPath)
                For i = 0 To FilesPath.Length - 1
                    Dim FormPathValues As String() = GetValueArrayfromValuePath(FilesPath(i), "Form")
                    If File.Exists("" & SubFolder & "\" & FilesNameValue(i) & ".vb") = True Then
                        If FormPathValues.Length > 0 Then
                            ReplaceFilefromOnepathtoOther(FormPathValues, SubFolder)
                        End If
                    End If
                Next
            End If
            'for Class
            Dim ClassNameValue As String() = {}
            Dim ClassPath As String() = {}
            If Not ClassName.Trim = "" Then
                ClassPath = GetSpitString(ClassName)
                ClassNameValue = GetFileName(ClassPath)
                For i = 0 To ClassPath.Length - 1
                    Dim ClassPathValues As String() = GetValueArrayfromValuePath(ClassPath(i), "Class")
                    If File.Exists("" & SubFolder & "\" & ClassNameValue(i) & ".vb") = True Then
                        If ClassPathValues.Length > 0 Then
                            ReplaceFilefromOnepathtoOther(ClassPathValues, SubFolder)
                        End If
                    End If
                Next
            End If
            'for Dll
            Dim DllNameValue As String() = {}
            Dim DllPath As String() = {}
            If Not DllName.Trim = "" Then
                DllPath = GetSpitString(DllName)
                DllNameValue = GetFileName(DllPath)
                For i = 0 To DllPath.Length - 1
                    Dim DllPathValues As String() = GetValueArrayfromValuePath(DllPath(i), "Dll")
                    If File.Exists("" & SubFolder & "\bin\Debug\" & DllNameValue(i) & ".dll") = True Then
                        If DllPathValues.Length > 0 Then
                            ReplaceFilefromOnepathtoOther(DllPathValues, "" & SubFolder & "\bin\Debug")
                        End If
                    End If
                Next
            End If
            'Update BuildPath
            Dim Vbprojpath As String = "" & SubFolder & "\" & ProjectName & ".vbproj"
            UpdateCodeinVbprojClassLibraryOrWindowsApp(Vbprojpath, BuildPath)
        End If
    End Sub
    ''' <summary>
    ''' to add code in vbproj file of windows application
    ''' </summary>
    ''' <param name="Vbprojpath"></param>Path of Vbproj file
    ''' <param name="VbProjCodelist"></param>Code of Form and class that has to be added in vbproj file
    ''' <param name="VbProjCodeResxlist"></param>Code of Resx that has to be added in vbproj file
    ''' <param name="VbProjCodeDlllist"></param>Code of dll that has to be added in vbproj file
    ''' <param name="BuildPath"></param> Represent the path where build dll have to be saved
    Public Sub AddCodeinVbprojWindowsApp(ByVal Vbprojpath As String, ByVal VbProjCodelist As List(Of String), ByVal VbProjCodeResxlist As List(Of String), ByVal VbProjCodeDlllist As List(Of String), ByVal BuildPath As String)
        Dim Compliecount As Integer = 0
        Dim Referencecount As Integer = 0
        Dim OutputPathcount As Integer = 0
        If File.Exists(Vbprojpath) Then
            Dim lines As String() = File.ReadAllLines(Vbprojpath)
            Dim projectVbprojFile As New List(Of String)(lines)
            For i = 0 To projectVbprojFile.Count - 1
                If projectVbprojFile(i).Contains("<OutputPath") Then
                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                    projectVbprojFile(i) = "<OutputPath>" & BuildPath & "\</OutputPath>"
                    lines = projectVbprojFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Vbprojpath, lines)
                ElseIf projectVbprojFile(i).Contains("<Reference") Then
                    Referencecount += 1
                    If Referencecount = 1 Then
                        Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                        Dim index As Integer = 0
                        If VbProjCodeDlllist.Count > 0 Then
                            For j = 0 To VbProjCodeDlllist.Count - 1
                                projectVbprojFile.Insert((i + index), VbProjCodeDlllist(j))
                                index += 1
                            Next
                        End If
                        lines = projectVbprojFile.ToArray()
                        Vbprojfile.Close()
                        File.WriteAllLines(Vbprojpath, lines)
                    End If
                ElseIf projectVbprojFile(i).Contains("</Compile>") Then
                    Compliecount += 1
                    If Compliecount = 2 Then
                        Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                        Dim index As Integer = 0
                        If VbProjCodelist.Count > 0 Then
                            For j = 0 To VbProjCodelist.Count - 1
                                projectVbprojFile.Insert((i + index) + 1, VbProjCodelist(j))
                                index += 1
                            Next
                        End If
                        lines = projectVbprojFile.ToArray()
                        Vbprojfile.Close()
                        File.WriteAllLines(Vbprojpath, lines)
                    End If
                Else
                    If projectVbprojFile(i).Contains("<EmbeddedResource") Then
                        Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                        Dim index As Integer = 0
                        If VbProjCodeResxlist.Count > 0 Then
                            For j = 0 To VbProjCodeResxlist.Count - 1
                                projectVbprojFile.Insert((i + index), VbProjCodeResxlist(j))
                                index += 1
                            Next
                        End If
                        lines = projectVbprojFile.ToArray()
                        Vbprojfile.Close()
                        File.WriteAllLines(Vbprojpath, lines)
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Code That has to be deleted from vbproj
    ''' </summary>
    ''' <param name="Vbprojpath"></param>Path of Vbproj file
    Public Sub DeleteCodeinVbprojWindowsApp(ByVal Vbprojpath As String)
        If File.Exists(Vbprojpath) Then
            Dim lines As String() = File.ReadAllLines(Vbprojpath)
            Dim projectVbprojFile As New List(Of String)(lines)
            For i = 0 To projectVbprojFile.Count - 1
                If projectVbprojFile(i).Contains("<Compile") AndAlso projectVbprojFile(i).Contains("Form1.vb") Then
                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                    projectVbprojFile.RemoveRange(i, 7)
                    lines = projectVbprojFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Vbprojpath, lines)
                    Exit For
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' To set the Value to the current form
    ''' </summary>
    ''' <param name="FormName"></param>The form name that has to be set
    ''' <param name="Apppath"></param>The app path of file
    Public Sub ChangeAppDesignerVal(ByVal FormName As String, ByVal Apppath As String)
        If File.Exists(Apppath) Then
            Dim lines As String() = File.ReadAllLines(Apppath)
            Dim AppFile As New List(Of String)(lines)
            For i = 0 To AppFile.Count - 1
                If AppFile(i).Contains("Me.MainForm") Then
                    Dim value1 As String() = AppFile(i).Split("=")
                    Dim value2 As String() = value1(1).Split(".")
                    Dim FinalVal As String = "" & value1(0) & "=" & value2(0) & "." & value2(1) & "." & FormName & ""
                    Dim Vbprojfile As New IO.StreamWriter(Apppath, True)
                    AppFile(i) = FinalVal
                    lines = AppFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Apppath, lines)
                    Exit For
                End If
            Next
        End If
    End Sub
    'Public Sub CreateUserControlProject(ByVal ProjectPath As String, ByVal ProjectName As String, ByVal ClassFileFolder As String, ByVal ClassName As String, ByVal DllFileFolder As String, ByVal DllName As String, ByVal ProjectTypeID As String, ByVal BuildPath As String)
    '    Dim MainFolder As String = "" & ProjectPath & "\" & ProjectName & ""
    '    Dim SubFolder As String = "" & MainFolder & "\" & ProjectName & ""
    '    If Not Directory.Exists(MainFolder) Then
    '        Directory.CreateDirectory(MainFolder)
    '        ''To create project
    '        'To create project
    '        CreateProject(SubFolder, "WindowsControl", ProjectName, ProjectTypeID)

    '        'For sln file
    '        Dim UniqueId As String = GetUniqueId("" & SubFolder & "\" & ProjectName & ".vbproj")
    '        Select Case ProjectTypeID
    '            Case "VisualStudio.DTE.14.0"
    '                CreateSlNFile15(UniqueId, ProjectName, MainFolder)
    '            Case "VisualStudio.DTE.9.0"
    '                CreateSlNFile08(UniqueId, ProjectName, MainFolder)
    '        End Select

    '        'for Class
    '        Dim ClassNameValue As String() = GetSpitString(ClassName)
    '        If ClassNameValue.Length > 0 Then
    '            For i = 0 To ClassNameValue.Length - 1
    '                Dim ClassPathValues As String() = GetValueArrayfromValuePath(ClassFileFolder, ClassNameValue(i), "Class")
    '                If ClassPathValues.Length > 0 Then
    '                    CopyFilefromOnepathtoOther(ClassPathValues, SubFolder)
    '                End If
    '            Next
    '        End If
    '        'for Dll
    '        Dim DllNameValue As String() = GetSpitString(DllName)
    '        If DllNameValue.Length > 0 Then
    '            For i = 0 To DllNameValue.Length - 1
    '                Dim DllPathValues As String() = GetValueArrayfromValuePath(DllFileFolder, DllNameValue(i), "Dll")
    '                If DllPathValues.Length > 0 Then
    '                    CopyFilefromOnepathtoOther(DllPathValues, "" & SubFolder & "\bin\Debug")
    '                End If
    '            Next
    '        End If
    '        'for vbproj
    '        'for Class vbproj code
    '        Dim VbProjCodeClasslist As New List(Of String)
    '        VbProjCodeClasslist = CreateVbprojClassCode(ClassNameValue)
    '        'for dll vbproj code
    '        Dim VbProjCodeDlllist As New List(Of String)
    '        VbProjCodeDlllist = CreateVbprojDllCodeUserControl(DllNameValue, DllFileFolder)

    '        'toget vbprojfile as list of string
    '        Dim Vbprojpath As String = "" & SubFolder & "\" & ProjectName & ".vbproj"
    '        AddCodeinVbprojUserControl(Vbprojpath, VbProjCodeClasslist, VbProjCodeDlllist, BuildPath)
    '    End If
    'End Sub
    'Public Sub AddCodeinVbprojUserControl(ByVal Vbprojpath As String, ByVal VbProjCodelist As List(Of String), ByVal VbProjCodeDlllist As List(Of String), ByVal BuildPath As String)
    '    Dim Referencecount As Integer = 0
    '    Dim OutputPathcount As Integer = 0
    '    If File.Exists(Vbprojpath) Then
    '        Dim lines As String() = File.ReadAllLines(Vbprojpath)
    '        Dim projectVbprojFile As New List(Of String)(lines)
    '        For i = 0 To projectVbprojFile.Count - 1
    '            If projectVbprojFile(i).Contains("<OutputPath") Then
    '                OutputPathcount += 1
    '                If OutputPathcount = 1 Then
    '                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
    '                    projectVbprojFile(i) = "<OutputPath>bin\" & BuildPath & "\</OutputPath>"
    '                    lines = projectVbprojFile.ToArray()
    '                    Vbprojfile.Close()
    '                    File.WriteAllLines(Vbprojpath, lines)
    '                End If
    '            ElseIf projectVbprojFile(i).Contains("<Reference") Then
    '                Referencecount += 1
    '                If Referencecount = 1 Then
    '                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
    '                    Dim index As Integer = 0
    '                    If VbProjCodeDlllist.Count > 0 Then
    '                        For j = 0 To VbProjCodeDlllist.Count - 1
    '                            projectVbprojFile.Insert((i + index), VbProjCodeDlllist(j))
    '                            index += 1
    '                        Next
    '                    End If
    '                    lines = projectVbprojFile.ToArray()
    '                    Vbprojfile.Close()
    '                    File.WriteAllLines(Vbprojpath, lines)
    '                End If
    '            ElseIf projectVbprojFile(i).Contains("<Compile") Then
    '                Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
    '                Dim index As Integer = 0
    '                If VbProjCodelist.Count > 0 Then
    '                    For j = 0 To VbProjCodelist.Count - 1
    '                        projectVbprojFile.Insert((i + index), VbProjCodelist(j))
    '                        index += 1
    '                    Next
    '                End If
    '                lines = projectVbprojFile.ToArray()
    '                Vbprojfile.Close()
    '                File.WriteAllLines(Vbprojpath, lines)
    '                Exit For
    '            End If
    '        Next
    '    End If
    'End Sub
    ''' <summary>
    ''' To create Class library project
    ''' </summary>
    ''' <param name="ProjectPath"></param>The path of the project to be saved..eg..E:\Newfolder
    ''' <param name="ProjectName"></param>The name given to the project.
    ''' <param name="FormName"></param>It represent comma separated values of path of form file with extension...eg..E:\Testing\WindowsApp\WindowsApp\Form1.vb
    ''' <param name="ClassName"></param>It represent comma separated values of path of class file with extension...eg..E:\Testing\WindowsApp\WindowsApp\Class1.vb
    ''' <param name="DllName"></param>It represent comma separated values of path of dll file with extension...eg..E:\Testing\WindowsApp\WindowsApp\Golbal1.dll
    ''' <param name="ProjectTypeID"></param>It represent the projectId i.e. For project 2008-"VisualStudio.DTE.9.0" and for project 2015-"VisualStudio.DTE.14.0"
    ''' <param name="BuildPath"></param>It represent the path where build dll have to be saved
    Public Sub CreateClassLibraryProject(ByVal ProjectPath As String, ByVal ProjectName As String, Optional ByVal FormName As String = "", Optional ByVal ClassName As String = "", Optional ByVal DllName As String = "", Optional ByVal ProjectTypeID As String = "VisualStudio.DTE.9.0", Optional ByVal BuildPath As String = "")
        Dim MainFolder As String = "" & ProjectPath & "\" & ProjectName & ""
        Dim SubFolder As String = "" & MainFolder & "\" & ProjectName & ""
        If Not Directory.Exists(MainFolder) Then
            Directory.CreateDirectory(MainFolder)

            'To create project
            CreateProject(SubFolder, "ClassLibrary", ProjectName, ProjectTypeID)

            'Delete the class1.vb
            File.Delete("" & SubFolder & "\Class1.vb")

            'For sln file
            Dim UniqueId As String = GetUniqueId("" & SubFolder & "\" & ProjectName & ".vbproj")
            Select Case ProjectTypeID
                Case "VisualStudio.DTE.14.0"
                    CreateSlNFile15(UniqueId, ProjectName, MainFolder)
                Case "VisualStudio.DTE.9.0"
                    CreateSlNFile08(UniqueId, ProjectName, MainFolder)
            End Select

            'For Forms
            Dim FilesNameValue() As String = {}
            Dim FilesPath() As String = {}
            If Not FormName.Trim = "" Then
                FilesPath = GetSpitString(FormName)
                FilesNameValue = GetFileName(FilesPath)
                For i = 0 To FilesPath.Length - 1
                    Dim FormPathValues As String() = GetValueArrayfromValuePath(FilesPath(i), "Form")
                    If FormPathValues.Length > 0 Then
                        CopyFilefromOnepathtoOther(FormPathValues, SubFolder)
                    End If
                Next
            End If
            'for Class
            Dim ClassNameValue As String() = {}
            Dim ClassPath As String() = {}
            If Not ClassName.Trim = "" Then
                ClassPath = GetSpitString(ClassName)
                ClassNameValue = GetFileName(ClassPath)
                For i = 0 To ClassPath.Length - 1
                    Dim ClassPathValues As String() = GetValueArrayfromValuePath(ClassPath(i), "Class")
                    If ClassPathValues.Length > 0 Then
                        CopyFilefromOnepathtoOther(ClassPathValues, SubFolder)
                    End If
                Next

            End If
            'for Dll
            Dim DllNameValue As String() = {}
            Dim DllPath As String() = {}
            If Not DllName.Trim = "" Then
                DllPath = GetSpitString(DllName)
                DllNameValue = GetFileName(DllPath)
                For i = 0 To DllPath.Length - 1
                    Dim DllPathValues As String() = GetValueArrayfromValuePath(DllPath(i), "Dll")
                    If DllPathValues.Length > 0 Then
                        CopyFilefromOnepathtoOther(DllPathValues, "" & SubFolder & "\bin\Debug")
                    End If
                Next
            End If
            'for vbproj

            'for form vbproj code
            Dim VbProjCodeRefFormlist As New List(Of String)
            VbProjCodeRefFormlist = CreateVbprojFormexcessCode()

            'for form vbproj code
            Dim VbProjCodeFormlist As New List(Of String)
            VbProjCodeFormlist = CreateVbprojFormCode(FilesNameValue)
            'for Class vbproj code
            Dim VbProjCodeClasslist As New List(Of String)
            VbProjCodeClasslist = CreateVbprojClassCode(ClassNameValue)
            'To combine form and class code
            Dim VbProjCodelist As New List(Of String)(VbProjCodeClasslist.Concat(VbProjCodeFormlist))

            'for resx vbproj code
            Dim VbProjCodeResxlist As New List(Of String)
            VbProjCodeResxlist = CreateVbprojResxCode(FilesNameValue)

            'for dll vbproj code
            Dim VbProjCodeDlllist As New List(Of String)
            VbProjCodeDlllist = CreateVbprojDllCode(DllNameValue, DllPath)

            'toget vbprojfile as list of string
            Dim Vbprojpath As String = "" & SubFolder & "\" & ProjectName & ".vbproj"
            AddCodeinVbprojClassLibrary(Vbprojpath, VbProjCodelist, VbProjCodeRefFormlist, VbProjCodeResxlist, VbProjCodeDlllist, BuildPath)

            'To delete the class1.vb code in vbproj file
            DeleteCodeinVbprojClassLibrary(Vbprojpath)

        Else
            'for upadte or inculde
            'For Forms
            Dim FilesNameValue() As String = {}
            Dim FilesPath() As String = {}
            If Not FormName.Trim = "" Then
                FilesPath = GetSpitString(FormName)
                FilesNameValue = GetFileName(FilesPath)
                For i = 0 To FilesPath.Length - 1
                    Dim FormPathValues As String() = GetValueArrayfromValuePath(FilesPath(i), "Form")
                    If File.Exists("" & SubFolder & "\" & FilesNameValue(i) & ".vb") = True Then
                        If FormPathValues.Length > 0 Then
                            ReplaceFilefromOnepathtoOther(FormPathValues, SubFolder)
                        End If
                    End If
                Next
            End If
            'for Class
            Dim ClassNameValue As String() = {}
            Dim ClassPath As String() = {}
            If Not ClassName.Trim = "" Then
                ClassPath = GetSpitString(ClassName)
                ClassNameValue = GetFileName(ClassPath)
                For i = 0 To ClassPath.Length - 1
                    Dim ClassPathValues As String() = GetValueArrayfromValuePath(ClassPath(i), "Class")
                    If File.Exists("" & SubFolder & "\" & ClassNameValue(i) & ".vb") = True Then
                        If ClassPathValues.Length > 0 Then
                            ReplaceFilefromOnepathtoOther(ClassPathValues, SubFolder)
                        End If
                    End If
                Next
            End If
            'for Dll
            Dim DllNameValue As String() = {}
            Dim DllPath As String() = {}
            If Not DllName.Trim = "" Then
                DllPath = GetSpitString(DllName)
                DllNameValue = GetFileName(DllPath)
                For i = 0 To DllPath.Length - 1
                    Dim DllPathValues As String() = GetValueArrayfromValuePath(DllPath(i), "Dll")
                    If File.Exists("" & SubFolder & "\bin\Debug\" & DllNameValue(i) & ".dll") = True Then
                        If DllPathValues.Length > 0 Then
                            ReplaceFilefromOnepathtoOther(DllPathValues, "" & SubFolder & "\bin\Debug")
                        End If
                    End If
                Next
            End If
            'Update BuildPath
            Dim Vbprojpath As String = "" & SubFolder & "\" & ProjectName & ".vbproj"
            UpdateCodeinVbprojClassLibraryOrWindowsApp(Vbprojpath, BuildPath)
        End If
    End Sub
    ''' <summary>
    ''' to add code in vbproj file of class library
    ''' </summary>
    ''' <param name="Vbprojpath"></param>Path of Vbproj file
    ''' <param name="VbProjCodelist"></param>Code of Form and class that has to be added in vbproj file
    ''' <param name="VbProjCodeRefFormlist"></param>Code of Reference that has to be added in vbproj file
    ''' <param name="VbProjCodeResxlist"></param>Code of Resx that has to be added in vbproj file
    ''' <param name="VbProjCodeDlllist"></param>Code of dll that has to be added in vbproj file
    ''' <param name="BuildPath"></param> Represent the path where build dll have to be saved
    Public Sub AddCodeinVbprojClassLibrary(ByVal Vbprojpath As String, ByVal VbProjCodelist As List(Of String), ByVal VbProjCodeRefFormlist As List(Of String), ByVal VbProjCodeResxlist As List(Of String), ByVal VbProjCodeDlllist As List(Of String), ByVal BuildPath As String)
        Dim Compliecount As Integer = 0
        Dim Referencecount As Integer = 0
        Dim value1 As String = "~System.Data~"
        Dim Value1Final As String = value1.Replace("~", """")
        Dim OutputPathcount As Integer = 0
        If File.Exists(Vbprojpath) Then
            Dim lines As String() = File.ReadAllLines(Vbprojpath)
            Dim projectVbprojFile As New List(Of String)(lines)
            For i = 0 To projectVbprojFile.Count - 1
                If projectVbprojFile(i).Contains("<OutputPath") Then
                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                    projectVbprojFile(i) = "<OutputPath>" & BuildPath & "\</OutputPath>"
                    lines = projectVbprojFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Vbprojpath, lines)
                ElseIf projectVbprojFile(i).Contains("<Compile") Then
                    Compliecount += 1
                    If Compliecount = 1 Then
                        Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                        Dim index As Integer = 0
                        If VbProjCodelist.Count > 0 Then
                            For j = 0 To VbProjCodelist.Count - 1
                                projectVbprojFile.Insert((i + index) + 1, VbProjCodelist(j))
                                index += 1
                            Next
                        End If
                        lines = projectVbprojFile.ToArray()
                        Vbprojfile.Close()
                        File.WriteAllLines(Vbprojpath, lines)
                    End If
                ElseIf projectVbprojFile(i).Contains("<Reference") Then
                    Referencecount += 1
                    If Referencecount = 1 Then
                        Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                        Dim index As Integer = 0
                        If VbProjCodeDlllist.Count > 0 Then
                            For j = 0 To VbProjCodeDlllist.Count - 1
                                projectVbprojFile.Insert((i + index), VbProjCodeDlllist(j))
                                index += 1
                            Next
                        End If
                        lines = projectVbprojFile.ToArray()
                        Vbprojfile.Close()
                        File.WriteAllLines(Vbprojpath, lines)
                    ElseIf projectVbprojFile(i).Contains(Value1Final) AndAlso projectVbprojFile(i).Contains("<Reference") Then
                        Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                        Dim index As Integer = 0
                        If VbProjCodeRefFormlist.Count > 0 Then
                            For j = 0 To VbProjCodeRefFormlist.Count - 1
                                projectVbprojFile.Insert((i + index) + 1, VbProjCodeRefFormlist(j))
                                index += 1
                            Next
                        End If
                        lines = projectVbprojFile.ToArray()
                        Vbprojfile.Close()
                        File.WriteAllLines(Vbprojpath, lines)
                    End If
                ElseIf projectVbprojFile(i).Contains("<EmbeddedResource") Then
                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                    Dim index As Integer = 0
                    If VbProjCodeResxlist.Count > 0 Then
                        For j = 0 To VbProjCodeResxlist.Count - 1
                            projectVbprojFile.Insert((i + index), VbProjCodeResxlist(j))
                            index += 1
                        Next
                    End If
                    lines = projectVbprojFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Vbprojpath, lines)
                    Exit For
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' Code that has to be deleted from the class library vbproj file
    ''' </summary>
    ''' <param name="Vbprojpath"></param>Path of vbproj file
    Public Sub DeleteCodeinVbprojClassLibrary(ByVal Vbprojpath As String)
        If File.Exists(Vbprojpath) Then
            Dim lines As String() = File.ReadAllLines(Vbprojpath)
            Dim projectVbprojFile As New List(Of String)(lines)
            For i = 0 To projectVbprojFile.Count - 1
                If projectVbprojFile(i).Contains("<Compile") AndAlso projectVbprojFile(i).Contains("Class1.vb") Then
                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                    projectVbprojFile.Remove(projectVbprojFile(i))
                    lines = projectVbprojFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Vbprojpath, lines)
                    Exit For
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' The code that need to be added in vbproj file
    ''' </summary>
    ''' <returns></returns>
    Public Function CreateVbprojFormexcessCode() As List(Of String)
        Dim VbProjCodeClasslist As New List(Of String)
        Dim Value1 As String = "<Reference Include=~System.Drawing~ />"
        Dim Value1Final As String = Value1.Replace("~", """")
        VbProjCodeClasslist.Add(Value1Final)
        Dim Value2 As String = "<Reference Include=~System.Windows.Forms~ />"
        Dim Value2Final As String = Value2.Replace("~", """")
        VbProjCodeClasslist.Add(Value2Final)
        Return VbProjCodeClasslist
    End Function
    ''' <summary>
    ''' The resx code that need to be added in vbproj file
    ''' </summary>
    ''' <param name="Filenames"></param>The name of the form file
    ''' <returns></returns>
    Public Function CreateVbprojResxCode(ByVal Filenames As String()) As List(Of String)
        Dim VbProjResxCodelist As New List(Of String)
        For j = 0 To Filenames.Length - 1
            Dim Value1 As String = "<EmbeddedResource Include=~" & Filenames(j) & ".resx~>"
            Dim Value1Final As String = Value1.Replace("~", """")
            VbProjResxCodelist.Add(Value1Final)
            VbProjResxCodelist.Add("<DependentUpon>" & Filenames(j) & ".vb</DependentUpon>")
            VbProjResxCodelist.Add("</EmbeddedResource>")
        Next
        Return VbProjResxCodelist
    End Function
    ''' <summary>
    '''  The Class code that need to be added in vbproj file
    ''' </summary>
    ''' <param name="Classnames"></param>The name of the class file
    ''' <returns></returns>
    Public Function CreateVbprojClassCode(ByVal Classnames As String()) As List(Of String)
        Dim VbProjCodeClasslist As New List(Of String)
        For j = 0 To Classnames.Length - 1
            Dim Value1 As String = "<Compile Include=~" & Classnames(j) & ".vb~ />"
            Dim Value1Final As String = Value1.Replace("~", """")
            VbProjCodeClasslist.Add(Value1Final)
        Next
        Return VbProjCodeClasslist
    End Function
    ''' <summary>
    ''' The Form code that need to be added in vbproj file
    ''' </summary>
    ''' <param name="Filenames"></param>The name of the form file
    ''' <returns></returns>
    Public Function CreateVbprojFormCode(ByVal Filenames As String()) As List(Of String)
        Dim VbProjCodeFormlist As New List(Of String)
        For j = 0 To Filenames.Length - 1
            Dim Value1 As String = "<Compile Include=~" & Filenames(j) & ".vb~>"
            Dim Value1Final As String = Value1.Replace("~", """")
            VbProjCodeFormlist.Add(Value1Final)
            VbProjCodeFormlist.Add("<SubType>Form</SubType>")
            VbProjCodeFormlist.Add("</Compile>")
            Dim Value2 As String = "<Compile Include=~" & Filenames(j) & ".Designer.vb~>"
            Dim Value2Final As String = Value2.Replace("~", """")
            VbProjCodeFormlist.Add(Value2Final)
            VbProjCodeFormlist.Add("<DependentUpon>" & Filenames(j) & ".vb</DependentUpon>")
            VbProjCodeFormlist.Add("<SubType>Form</SubType>")
            VbProjCodeFormlist.Add("</Compile>")
        Next
        Return VbProjCodeFormlist
    End Function
    ''' <summary>
    ''' The dll code that need to be added in vbproj file
    ''' </summary>
    ''' <param name="Dllnames"></param>The name of the dll file
    ''' <param name="DllPath"></param>The path of the dll file
    ''' <returns></returns>
    Public Function CreateVbprojDllCode(ByVal Dllnames As String(), ByVal DllPath As String()) As List(Of String)
        Dim VbProjCodeDlllist As New List(Of String)
        For j = 0 To Dllnames.Length - 1
            Dim Value1 As String = "<Reference Include=~" & Dllnames(j) & "~>"
            Dim Value1Final As String = Value1.Replace("~", """")
            VbProjCodeDlllist.Add(Value1Final)
            VbProjCodeDlllist.Add("<HintPath>" & DllPath(j) & "</HintPath>")
            VbProjCodeDlllist.Add("</Reference>")
        Next
        Return VbProjCodeDlllist
    End Function
    ''' <summary>
    ''' The dll code that need to be added in vbproj file for usercontrol
    ''' </summary>
    ''' <param name="Dllnames"></param>The name of the dll file
    ''' <param name="DllFileFolder"></param>The path of the dll file
    ''' <returns></returns>
    Public Function CreateVbprojDllCodeUserControl(ByVal Dllnames As String(), ByVal DllFileFolder As String) As List(Of String)
        Dim VbProjCodeDlllist As New List(Of String)
        For j = 0 To Dllnames.Length - 1
            Dim Value1 As String = "<Reference Include=~" & Dllnames(j) & ", Version=1.0.0.0, Culture=neutral, processorArchitecture=MSIL~>"
            Dim Value1Final As String = Value1.Replace("~", """")
            VbProjCodeDlllist.Add(Value1Final)
            VbProjCodeDlllist.Add("<SpecificVersion>False</SpecificVersion>")
            VbProjCodeDlllist.Add("<HintPath>" & DllFileFolder & "\" & Dllnames(j) & ".dll</HintPath>")
            VbProjCodeDlllist.Add("</Reference>")
        Next
        Return VbProjCodeDlllist
    End Function
    ''' <summary>
    ''' To update BuildPathcode in Vbproj file
    ''' </summary>
    ''' <param name="Vbprojpath"></param>Path of Vbproj file
    ''' <param name="BuildPath"></param>Represent the path where build dll have to be saved
    Public Sub UpdateCodeinVbprojClassLibraryOrWindowsApp(ByVal Vbprojpath As String, ByVal BuildPath As String)
        If File.Exists(Vbprojpath) Then
            Dim lines As String() = File.ReadAllLines(Vbprojpath)
            Dim projectVbprojFile As New List(Of String)(lines)
            For i = 0 To projectVbprojFile.Count - 1
                If projectVbprojFile(i).Contains("<OutputPath") Then
                    Dim Vbprojfile As New IO.StreamWriter(Vbprojpath, True)
                    projectVbprojFile(i) = "<OutputPath>" & BuildPath & "\</OutputPath>"
                    lines = projectVbprojFile.ToArray()
                    Vbprojfile.Close()
                    File.WriteAllLines(Vbprojpath, lines)
                End If
            Next
        End If
    End Sub
    ''' <summary>
    ''' to get the spit separated string array
    ''' </summary>
    ''' <param name="Value"></param>the value that has to be split
    ''' <returns></returns>
    Public Function GetSpitString(ByVal Value As String) As String()
        'Dim SpitString As String() = Value.Split(",")
        'Return SpitString
        Dim SpitString As String() = {}
        If Value.Contains(",") Then
            SpitString = Value.Split(",")
        Else
            Array.Resize(SpitString, 1)
            SpitString(0) = Value
        End If
        Return SpitString
    End Function
    ''' <summary>
    ''' To get a array of file that has to coped or replace.
    ''' </summary>
    ''' <param name="ValueName"></param>The name of path of the file
    ''' <param name="ValType"></param>To get the type of the file i.e. "Form","Class","Dll"
    ''' <returns></returns>
    Function GetValueArrayfromValuePath(ByVal ValueName As String, ByVal ValType As String) As String() 'Changed by divya
        Dim valueArray As String() = {}
        Select Case ValType
            Case "Form"
                Dim FilterStr As String() = {".resx", ".Designer.vb"}
                valueArray = GetFilePathArr(FilterStr, ValueName)
            Case "Class"
                Dim FilterStr As String() = {}
                valueArray = GetFilePathArr(FilterStr, ValueName)
            Case "Dll"
                Dim FilterStr As String() = {".pdb"".xml"}
                valueArray = GetFilePathArr(FilterStr, ValueName)
        End Select
        Return valueArray
    End Function
    ''' <summary>
    ''' To get the array of the exist file
    ''' </summary>
    ''' <param name="FilterStr"></param>Contain extension for the files
    ''' <param name="ValueName"></param>The name of path of the file
    ''' <returns></returns>
    Function GetFilePathArr(ByVal FilterStr As String(), ByVal ValueName As String) 'add by divya
        Dim TempArr As New ArrayList()
        Dim TempArrTemp As New ArrayList()
        TempArrTemp.Add(ValueName)
        Dim val As String = Path.GetExtension(ValueName)
        Dim FileNoExt As String = ValueName.Replace(val, "")
        For i = 0 To FilterStr.Length - 1
            TempArrTemp.Add(FileNoExt & FilterStr(i))
        Next
        For i = 0 To TempArrTemp.Count - 1
            If File.Exists(TempArrTemp(i)) Then
                TempArr.Add(TempArrTemp(i))
            End If
        Next
        Return (CType(TempArr.ToArray(GetType(String)), String()))
    End Function
    ''' <summary>
    ''' To get the name of the file from the path
    ''' </summary>
    ''' <param name="Filepath"></param>The path of the file
    ''' <returns></returns>
    Function GetFileName(ByVal Filepath As String()) As String()
        Dim TempArr As New ArrayList()
        For i = 0 To Filepath.Length - 1
            TempArr.Add(Path.GetFileNameWithoutExtension(Filepath(i)))
        Next
        Return CType(TempArr.ToArray(GetType(String)), String())
    End Function
    ''' <summary>
    ''' To copy one file to another position
    ''' </summary>
    ''' <param name="InitialPath"></param>The initial path of the file
    ''' <param name="FinalPath"></param>The path where the file has to be copied
    Public Sub CopyFilefromOnepathtoOther(ByVal InitialPath As String(), ByVal FinalPath As String)
        For i = 0 To InitialPath.Length - 1
            Dim InitialPathvalue As String = InitialPath(i)
            Dim FileName As String = Path.GetFileName(InitialPath(i))
            Dim FinalPathvalue As String = " " & FinalPath & "\" & FileName & ""
            File.Copy(InitialPathvalue, FinalPathvalue)
        Next
    End Sub
    ''' <summary>
    '''  To replaced one file to another position
    ''' </summary>
    ''' <param name="InitialPath"></param>The initial path of the file
    ''' <param name="FinalPath"></param>The path where the file has to be replaced
    Public Sub ReplaceFilefromOnepathtoOther(ByVal InitialPath As String(), ByVal FinalPath As String)
        For i = 0 To InitialPath.Length - 1
            Dim InitialPathvalue As String = InitialPath(i)
            Dim FileName As String = Path.GetFileName(InitialPath(i))
            Dim FinalPathvalue As String = " " & FinalPath & "\" & FileName & ""
            Dim FileNameExt As String = Path.GetFileNameWithoutExtension(FinalPathvalue)
            Dim currdt As String = Now.ToString("yyyy_MM_dd_hh_mm_ss")
            FileSystem.Rename(FinalPathvalue, " " & FinalPath & "\" & FileNameExt & "_" & (i + 1) & currdt & "txt")
            File.Copy(InitialPathvalue, FinalPathvalue)
        Next
    End Sub
    ''' <summary>
    ''' To get the unique id
    ''' </summary>
    ''' <param name="VbprojfilePath"></param>The path of vbproj file 
    ''' <returns></returns>
    Public Function GetUniqueId(ByVal VbprojfilePath As String) As String
        Dim Uniqueid As String = ""
        Dim lines() As String = File.ReadAllLines(VbprojfilePath)
        For j = 0 To lines.Count - 1
            If lines(j).Contains("ProjectGuid") Then
                Dim Split1 As String() = lines(j).Split("{")
                Dim Split2 As String() = Split1(1).Split("}")
                Uniqueid = "{" & Split2(0) & "}"
            End If
        Next
        Return Uniqueid
    End Function
    ''' <summary>
    ''' To cretae the sln of 2015
    ''' </summary>
    ''' <param name="Uniqueid"></param>represent the unique id
    ''' <param name="ProjectName"></param>To get the name of the file
    ''' <param name="MainFolderPath"></param>To get the main folder of the file
    Public Sub CreateSlNFile15(ByVal Uniqueid As String, ByVal ProjectName As String, ByVal MainFolderPath As String)
        Dim slnfilepath = "" & MainFolderPath & "\" & ProjectName & ".sln"
        Dim Slnfile As New IO.StreamWriter(slnfilepath, True)
        Slnfile.WriteLine("Microsoft Visual Studio Solution File, Format Version 12.00")
        Slnfile.WriteLine("# Visual Studio 14")
        Slnfile.WriteLine("VisualStudioVersion = 14.0.25420.1")
        Slnfile.WriteLine("MinimumVisualStudioVersion = 10.0.40219.1")
        Dim SlnString As String = "Project(~{F184B08F-C81C-45F6-A57F-5ABD9991F28F}~) = ~" & ProjectName & "~, ~" & ProjectName & "\" & ProjectName & ".vbproj~, ~" & Uniqueid & "~"
        Dim SlnStringfinal As String = SlnString.Replace("~", """")
        Slnfile.WriteLine(SlnStringfinal)
        Slnfile.WriteLine("EndProject")
        Slnfile.WriteLine("Global")
        Slnfile.WriteLine("GlobalSection(SolutionConfigurationPlatforms) = preSolution")
        Slnfile.WriteLine("Debug|Any CPU = Debug|Any CPU")
        Slnfile.WriteLine("Release|Any CPU = Release|Any CPU")
        Slnfile.WriteLine("EndGlobalSection")
        Slnfile.WriteLine("GlobalSection(ProjectConfigurationPlatforms) = postSolution")
        Slnfile.WriteLine("" & Uniqueid & ".Debug|Any CPU.ActiveCfg = Debug|Any CPU")
        Slnfile.WriteLine("" & Uniqueid & ".Debug|Any CPU.Build.0 = Debug|Any CPU")
        Slnfile.WriteLine("" & Uniqueid & ".Release|Any CPU.ActiveCfg = Release|Any CPU")
        Slnfile.WriteLine("" & Uniqueid & ".Release|Any CPU.Build.0 = Release|Any CPU")
        Slnfile.WriteLine("EndGlobalSection")
        Slnfile.WriteLine("GlobalSection(SolutionProperties) = preSolution")
        Slnfile.WriteLine("HideSolutionNode = FALSE")
        Slnfile.WriteLine("EndGlobalSection")
        Slnfile.WriteLine("EndGlobal")
        Slnfile.Close()
    End Sub
    ''' <summary>
    ''' To cretae the sln of 2008
    ''' </summary>
    ''' <param name="Uniqueid"></param>represent the unique id
    ''' <param name="ProjectName"></param>To get the name of the file
    ''' <param name="MainFolderPath"></param>to get the main folder of the file
    Public Sub CreateSlNFile08(ByVal Uniqueid As String, ByVal ProjectName As String, ByVal MainFolderPath As String)
        Dim slnfilepath = "" & MainFolderPath & "\" & ProjectName & ".sln"
        Dim Slnfile As New IO.StreamWriter(slnfilepath, True)
        Slnfile.WriteLine("Microsoft Visual Studio Solution File, Format Version 10.00")
        Slnfile.WriteLine("# Visual Studio 2008")
        Dim SlnString As String = "Project(~{F184B08F-C81C-45F6-A57F-5ABD9991F28F}~) = ~" & ProjectName & "~, ~" & ProjectName & "\" & ProjectName & ".vbproj~, ~" & Uniqueid & "~"
        Dim SlnStringfinal As String = SlnString.Replace("~", """")
        Slnfile.WriteLine(SlnStringfinal)
        Slnfile.WriteLine("EndProject")
        Slnfile.WriteLine("Global")
        Slnfile.WriteLine("GlobalSection(SolutionConfigurationPlatforms) = preSolution")
        Slnfile.WriteLine("Debug|Any CPU = Debug|Any CPU")
        Slnfile.WriteLine("Release|Any CPU = Release|Any CPU")
        Slnfile.WriteLine("EndGlobalSection")
        Slnfile.WriteLine("GlobalSection(ProjectConfigurationPlatforms) = postSolution")
        Slnfile.WriteLine("" & Uniqueid & ".Debug|Any CPU.ActiveCfg = Debug|Any CPU")
        Slnfile.WriteLine("" & Uniqueid & ".Debug|Any CPU.Build.0 = Debug|Any CPU")
        Slnfile.WriteLine("" & Uniqueid & ".Release|Any CPU.ActiveCfg = Release|Any CPU")
        Slnfile.WriteLine("" & Uniqueid & ".Release|Any CPU.Build.0 = Release|Any CPU")
        Slnfile.WriteLine("EndGlobalSection")
        Slnfile.WriteLine("GlobalSection(SolutionProperties) = preSolution")
        Slnfile.WriteLine("HideSolutionNode = FALSE")
        Slnfile.WriteLine("EndGlobalSection")
        Slnfile.WriteLine("EndGlobal")
        Slnfile.Close()
    End Sub

    ''' <summary>
    ''' To build the the Project of the vbproj file provided.
    ''' </summary>
    ''' <param name="VbprojFullPath">Represent the path of vbproj file.. eg-"E:\laveena\DemoClassLib\DemoClassLib\DemoClassLib.vbproj</param>
    ''' <param name="OutputPath">Path where dll will be created</param>
    ''' <param name="BuildType">Type of BuildType i.e "Coded" or "Normal".Bydefault, BuildType is coded,then OutputPath can be "",As output path is alraedy set in vbproj.</param>
    ''' <param name="MsBuildExePath">Path of MSBuildExe</param>
    ''' <remarks></remarks>
    Public Sub BuildProject(ByVal VbprojFullPath As String, ByVal OutputPath As String, Optional ByVal BuildType As String = "Coded", Optional ByVal MsBuildExePath As String = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\msbuild.exe")
        Dim p As New Process()
        p.StartInfo = New ProcessStartInfo(MsBuildExePath)
        p.StartInfo.RedirectStandardOutput = True
        p.StartInfo.UseShellExecute = False
        p.StartInfo.Arguments = String.Format("""" & VbprojFullPath & """" & " /t:rebuild")
        p.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        p.Start()
        Dim MsBuildOutput As String = p.StandardOutput.ReadToEnd()
        If LCase(BuildType) = "normal" Then
            'get dll from debug
            Dim DllName As String = Path.GetFileNameWithoutExtension(VbprojFullPath)
            Dim PathVal As String = Path.GetDirectoryName(VbprojFullPath)
            Dim DllFileToCopy As New List(Of String)
            DllFileToCopy.Add("" & PathVal & "\bin\Debug\" & DllName & ".dll")
            DllFileToCopy.Add("" & PathVal & "\bin\Debug\" & DllName & ".pdb")
            DllFileToCopy.Add("" & PathVal & "\bin\Debug\" & DllName & ".xml")

            'copy dll to outpath
            Dim DllFileToDlt As New List(Of String)
            DllFileToDlt.Add("" & OutputPath & "\" & DllName & ".dll")
            DllFileToDlt.Add("" & OutputPath & "\" & DllName & ".pdb")
            DllFileToDlt.Add("" & OutputPath & "\" & DllName & ".xml")
            For i = 0 To DllFileToDlt.Count - 1
                If File.Exists(DllFileToDlt(i)) Then
                    My.Computer.FileSystem.DeleteFile(DllFileToDlt(i))
                End If
            Next
            For i = 0 To DllFileToCopy.Count - 1
                If File.Exists(DllFileToCopy(i)) Then
                    File.Copy(DllFileToCopy(i), DllFileToDlt(i))
                End If
            Next
        End If
    End Sub


End Class

