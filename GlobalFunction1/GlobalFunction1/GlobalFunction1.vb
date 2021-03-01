#Region "Imports"
Imports System.Reflection
Imports System.Windows.Forms
Imports System.IO
Imports RegisterDllClass
Imports System.Drawing
Imports System.CodeDom.Compiler
Imports System.Resources
Imports GlobalControl.Variables
Imports System.DirectoryServices
Imports System.Net
Imports System.Security.Permissions
'Imports Microsoft.SolverFoundation.Common
'Imports Microsoft.SolverFoundation.Services
Imports DataFunctions

'Imports System.ComponentModel
#End Region
Public Class GlobalFunction1
    Dim _InstanceName As String = ""
    ''' <summary>
    ''' Instance variable name as string, if taken as constructor of new class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property InstanceName() As String
        Get
            Return _InstanceName
        End Get
        Set(ByVal value As String)
            value = _InstanceName
        End Set
    End Property


#Region "DimVariables"
    Dim dllInteger As Integer = 280196 + 130858
    ' Dim sp1 As New SetPropertyClass.SetProperties
    Public Sub New()
        Try
            If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then
                Dim rc As New RegisterDllClass.ValidateClass
                Dim AuFlag As Integer = rc.AdminVault
                If AuFlag <> dllInteger Then
                    QuitMessage(Me.ToString & " " & rc.ReverseString(rc.aumess, 1), "New")
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.New")
        End Try
    End Sub
#End Region

#Region "Functions"
    ''' <summary>
    ''' To add resources to *.resx file from a hash table. 
    ''' </summary>
    ''' <param name="ResourceEntries">Resource Entries as hash table , where key as name, and value as content</param>
    ''' <param name="ResourceFilePath">Full path Name of resource file</param>
    ''' <remarks></remarks>
    Public Sub AddItemToResourceFile(ByVal ResourceEntries As Hashtable, Optional ByVal ResourceFilePath As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            If My.Computer.FileSystem.FileExists(ResourceFilePath) = True Then
                Dim mhash As New Hashtable
                Dim ResxReader As New ResourceReader(ResourceFilePath)
                For Each d As DictionaryEntry In ResxReader
                    Try
                        mhash = AddItemToHashTable(mhash, d.Key, d.Value)
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to read resource " & d.Key.ToString)
                    End Try

                Next
                For i = 0 To ResourceEntries.Count - 1
                    Dim mkey As String = ResourceEntries.Keys(i).ToString
                    Try
                        mhash = AddItemToHashTable(mhash, mkey, ResourceEntries.Item(mkey))
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to add resource " & mkey)
                    End Try
                Next
                Dim ResxWriter As New ResXResourceWriter(ResourceFilePath)
                For i = 0 To mhash.Count - 1
                    Dim mkey As String = mhash.Keys(i).ToString
                    Try
                        ResxWriter.AddResource(mkey, mhash.Item(mkey))
                    Catch ex As Exception
                        QuitError(ex, Err, "Unable to write resource " & mkey)
                    End Try
                Next
                ResxReader.Close()
                ResxWriter.Close()
            Else
                Dim ResxWriter As New ResXResourceWriter(ResourceFilePath)
                For i = 0 To ResourceEntries.Count - 1
                    Dim mkey As String = ResourceEntries.Keys(i).ToString
                    ResxWriter.AddResource(mkey, ResourceEntries.Item(mkey))
                Next
                ResxWriter.Close()
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute AddItemToResourceFile(ByVal ResourceFilePath As String, ByVal ResourceEntries As Hashtable)")
        End Try
    End Sub

    ''' <summary>
    ''' Function to get value of a keyword from a composite fields value string
    ''' </summary>
    ''' <param name="CompositeFieldString">String value of composite fields, where default pair separator is chrw(200) and default key=value separtor is chrw(210)</param>
    ''' <param name="Keyword">Keyword/field name to be searched</param>
    ''' <param name="ValuePairSeparator">separator between two fields,default is chrw(200)</param>
    ''' <param name="ValueKeySeparator">separator between  field and value ,default is chrw(201)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetKeywordValueFromCompositeField(ByVal CompositeFieldString As String, ByVal Keyword As String, Optional ByVal ValuePairSeparator As String = ChrW(201), Optional ByVal ValueKeySeparator As String = ChrW(200)) As String
        Dim mvalue As String = ""
        ' Dim Method_Name As String = "GetKeywordValue", param_names() As String = {"CompositeFieldString", "Keyword", "ValuePairSeparator", "ValueKeySeparator"}, param_values() As Object = {CompositeFieldString, Keyword, ValuePairSeparator, ValueKeySeparator} : rdc.StartMethod(Method_Name, param_names, param_values)
        '  Try
        Dim aCompositeString() As String = CompositeFieldString.Split(ValuePairSeparator)
        For i = 0 To aCompositeString.Count - 1
            Dim apair() As String = aCompositeString(i).Trim.Split(ValueKeySeparator)
            Select Case apair.Count
                Case 2, 3
                    If LCase(apair(0)) = LCase(Keyword) Then
                        mvalue = apair(1)
                        Exit For
                    End If
            End Select
        Next
        '    rdc.EndMethod()
        Return mvalue
        ' Catch Ex As Exception
        '    rdc.QuitError(Ex, Err, New StackTrace(True))
        '  End Try
        ' rdc.EndMethod()
        Return mvalue
    End Function




    ''' <summary>
    ''' Get item from resources (resx) file
    ''' </summary>
    ''' <param name="KeyName">Key name of resource to be found</param>
    ''' <param name="ResourceFilePath">Full path Name of resource file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetItemFromResources(ByVal KeyName As String, Optional ByVal ResourceFilePath As String = "") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim ResxReader As New ResourceReader(ResourceFilePath)
        Dim RetVal As New Object
        For Each d As DictionaryEntry In ResxReader
            Try
                If LCase(d.Key.ToString) = LCase(KeyName) Then
                    RetVal = d.Value
                    Exit For
                End If
            Catch ex As Exception
                QuitError(ex, Err, "Unable to read resource for " & d.Key.ToString)
            End Try
        Next
        Return RetVal
    End Function



    ''' <summary>
    ''' Add string element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">string element added at last position </param>
    ''' <returns> Output array after adding new string </returns>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <remarks></remarks>
    Public Function ArrayAppend(ByRef ArrayName() As String, ByVal LastValue As String, Optional ByVal IgnoreIfExists As Boolean = False, Optional ByVal IgnoreCase As Boolean = True) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = ArrayFind(ArrayName, LastValue, IgnoreCase)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    ''' Add Two arrays of string elements
    ''' </summary>
    ''' <param name="FirstArray"> FirstArray to be added</param>
    ''' <param name="SecondArray ">Second Array to be added </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoArrays(ByRef FirstArray() As String, ByVal SecondArray() As String, Optional ByVal IgnoreIfExists As Boolean = False, Optional ByVal IgnoreCase As Boolean = True) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondArray.Count - 1
                Dim mindx As Integer = -1
                Dim LastValue As String = SecondArray(i)
                If IgnoreIfExists = True Then
                    mindx = ArrayFind(FirstArray, LastValue, IgnoreCase)
                End If
                If mindx = -1 Then
                    Dim ii As Integer = FirstArray.Length
                    ReDim Preserve FirstArray(ii)
                    FirstArray.SetValue(SecondArray(i), ii)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As String, ByVal SecondArray() As String, Optional ByVal IgnoreIfExists As Boolean = False")
        End Try
        Return FirstArray
    End Function
    ''' <summary>
    ''' Add Two arrays of integer elements
    ''' </summary>
    ''' <param name="FirstArray"> FirstArray to be added</param>
    ''' <param name="SecondArray ">Second Array to be added </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoArrays(ByRef FirstArray() As Integer, ByVal SecondArray() As Integer, Optional ByVal IgnoreIfExists As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondArray.Count - 1
                Dim mindx As Integer = -1
                Dim LastValue As Integer = SecondArray(i)
                If IgnoreIfExists = True Then
                    mindx = Array.IndexOf(FirstArray, LastValue)
                End If
                If mindx = -1 Then
                    Dim ii As Integer = FirstArray.Length
                    ReDim Preserve FirstArray(ii)
                    FirstArray.SetValue(SecondArray(i), ii)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As Integer, ByVal SecondArray() As Integer, Optional ByVal IgnoreIfExists As Boolean = False")
        End Try
        Return FirstArray
    End Function
    ''' <summary>
    ''' Add Two arrays of decimal elements
    ''' </summary>
    ''' <param name="FirstArray"> FirstArray to be added</param>
    ''' <param name="SecondArray ">Second Array to be added </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoArrays(ByRef FirstArray() As Decimal, ByVal SecondArray() As Decimal, Optional ByVal IgnoreIfExists As Boolean = False) As Decimal()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondArray.Count - 1
                Dim mindx As Integer = -1
                Dim LastValue As Integer = SecondArray(i)
                If IgnoreIfExists = True Then
                    mindx = Array.IndexOf(FirstArray, LastValue)
                End If
                If mindx = -1 Then
                    Dim ii As Integer = FirstArray.Length
                    ReDim Preserve FirstArray(ii)
                    FirstArray.SetValue(SecondArray(i), ii)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As Integer, ByVal SecondArray() As Integer, Optional ByVal IgnoreIfExists As Boolean = False")
        End Try
        Return FirstArray
    End Function



    ''' <summary>
    ''' Add Two arrays of object elements
    ''' </summary>
    ''' <param name="FirstArray"> FirstArray to be added</param>
    ''' <param name="SecondArray ">Second Array to be added </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoArrays(ByRef FirstArray() As Object, ByVal SecondArray() As Object, Optional ByVal IgnoreIfExists As Boolean = False) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondArray.Count - 1
                Dim mindx As Integer = -1
                Dim LastValue As Object = SecondArray(i)
                If IgnoreIfExists = True Then
                    mindx = Array.IndexOf(FirstArray, LastValue)
                End If
                If mindx = -1 Then
                    Dim ii As Integer = FirstArray.Length
                    ReDim Preserve FirstArray(ii)
                    FirstArray.SetValue(SecondArray(i), ii)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As object, ByVal SecondArray() As object, Optional ByVal IgnoreIfExists As Boolean = False")
        End Try
        Return FirstArray
    End Function
    ''' <summary>
    ''' Add Two HashTables  
    ''' </summary>
    ''' <param name="FirstHashTable"> FirstHashTable to be added</param>
    ''' <param name="SecondHashTable ">SecondHashTable to be added </param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoHashTable(ByRef FirstHashTable As Hashtable, ByVal SecondHashTable As Hashtable, Optional ByVal ReplaceIfKeyExists As Boolean = True) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondHashTable.Count - 1
                Dim mkey As String = SecondHashTable.Keys(i)
                Dim mvalue As Object = SecondHashTable.Item(mkey)
                FirstHashTable = AddItemToHashTable(FirstHashTable, mkey, mvalue)

            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As object, ByVal SecondArray() As object, Optional ByVal IgnoreIfExists As Boolean = False")
        End Try
        Return FirstHashTable
    End Function

    ''' <summary>
    ''' Add Two arrays of fileinfo elements
    ''' </summary>
    ''' <param name="FirstArray"> FirstArray to be added</param>
    ''' <param name="SecondArray ">Second Array to be added </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoArrays(ByRef FirstArray() As FileInfo, ByVal SecondArray() As FileInfo, Optional ByVal IgnoreIfExists As Boolean = False) As FileInfo()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondArray.Count - 1
                Dim mindx As Integer = -1
                Dim LastValue As FileInfo = SecondArray(i)
                If IgnoreIfExists = True Then
                    mindx = Array.IndexOf(FirstArray, LastValue)
                End If
                If mindx = -1 Then
                    Dim ii As Integer = FirstArray.Length
                    ReDim Preserve FirstArray(ii)
                    FirstArray.SetValue(SecondArray(i), ii)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As object, ByVal SecondArray() As object, Optional ByVal IgnoreIfExists As Boolean = False")
        End Try
        Return FirstArray
    End Function


    ''' <summary>
    ''' Add Two arrays of Collection elements
    ''' </summary>
    ''' <param name="FirstArray"> FirstArray to be added</param>
    ''' <param name="SecondArray ">Second Array to be added </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after addition </returns>
    ''' <remarks></remarks>
    Public Function AddTwoArrays(ByRef FirstArray() As Collection, ByVal SecondArray() As Collection, Optional ByVal IgnoreIfExists As Boolean = False) As Collection()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To SecondArray.Count - 1
                Dim mindx As Integer = -1
                Dim LastValue As Object = SecondArray(i)
                If IgnoreIfExists = True Then
                    mindx = Array.IndexOf(FirstArray, LastValue)
                End If
                If mindx = -1 Then
                    Dim ii As Integer = FirstArray.Length
                    ReDim Preserve FirstArray(ii)
                    FirstArray.SetValue(SecondArray(i), ii)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTwoArrays(ByRef FirstArray() As Collection, ByVal SecondArray() As Collection, Optional ByVal IgnoreIfExists As Boolean = False) As Collection()")
        End Try
        Return FirstArray
    End Function


    ''' <summary>
    ''' Add an integer element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">Integer element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new integer </returns>
    ''' <remarks></remarks>

    Public Function ArrayAppend(ByRef ArrayName() As Integer, ByVal LastValue As Integer, Optional ByVal IgnoreIfExists As Boolean = False) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Integer, ByVal LastValue As Integer, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    ''' True,if hashtable has all keyvalues nothing/empty
    ''' </summary>
    ''' <param name="mhashTable"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsHashValuesEmpty(ByVal mhashTable As Hashtable) As Boolean
        Dim mflg As Boolean = True
        For i = 0 To mhashTable.Count - 1
            If Not mhashTable.Item(mhashTable.Keys(i)) = Nothing Then
                mflg = False
                Exit For
            End If
        Next
        Return mflg
    End Function
    ''' <summary>
    ''' Add a decimal element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">decimal element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new decimal </returns>
    ''' <remarks></remarks>   
    ''' 
    Public Function ArrayAppend(ByRef ArrayName() As Decimal, ByVal LastValue As Decimal, Optional ByVal IgnoreIfExists As Boolean = False) As Decimal()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Decimal, ByVal LastValue As Decimal, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    ''' Add a Single element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">Single element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new decimal </returns>
    ''' <remarks></remarks>   
    ''' 
    Public Function ArrayAppend(ByRef ArrayName() As Single, ByVal LastValue As Single, Optional ByVal IgnoreIfExists As Boolean = False) As Single()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Decimal, ByVal LastValue As Decimal, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    ''' Add a Color element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">Color element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new decimal </returns>
    ''' <remarks></remarks>   
    ''' 
    Public Function ArrayAppend(ByRef ArrayName() As Color, ByVal LastValue As Color, Optional ByVal IgnoreIfExists As Boolean = False) As Color()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Decimal, ByVal LastValue As Decimal, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    ''' Add a object element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">object element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new object </returns>
    ''' <remarks></remarks>   
    Public Function ArrayAppend(ByRef ArrayName() As Object, ByVal LastValue As Object, Optional ByVal IgnoreIfExists As Boolean = False) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Object, ByVal LastValue As Object, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    ''' Add a fileinfo element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">object element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new object </returns>
    ''' <remarks></remarks>   
    Public Function ArrayAppend(ByRef ArrayName() As FileInfo, ByVal LastValue As FileInfo, Optional ByVal IgnoreIfExists As Boolean = False) As FileInfo()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Object, ByVal LastValue As Object, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function



    ''' <summary>
    ''' Add a object element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">Collection element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new object </returns>
    ''' <remarks></remarks>   
    Public Function ArrayAppend(ByRef ArrayName() As Collection, ByVal LastValue As Collection, Optional ByVal IgnoreIfExists As Boolean = False) As Collection()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Collection, ByVal LastValue As Collection, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    ''' Add a object element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">HashTable element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new object </returns>
    ''' <remarks></remarks>   
    Public Function ArrayAppend(ByRef ArrayName() As Hashtable, ByVal LastValue As Hashtable, Optional ByVal IgnoreIfExists As Boolean = False) As Hashtable()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Hashtable, ByVal LastValue As Hashtable, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function



    '''' <summary>
    '''' Add a object element at the last of an array
    '''' </summary>
    '''' <param name="ArrayName"> Array to be added</param>
    '''' <param name="LastValue">object element added at last position </param>
    '''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    '''' <returns> Output array after adding new object </returns>
    '''' <remarks></remarks>   
    'Public Function ArrayAppend(ByRef ArrayName() As Decision, ByVal LastValue As Decision, Optional ByVal IgnoreIfExists As Boolean = False) As Object()
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    Try
    '        Dim mindx As Integer = -1
    '        If IgnoreIfExists = True Then
    '            mindx = Array.IndexOf(ArrayName, LastValue)
    '        End If
    '        If mindx = -1 Then
    '            Dim ii As Integer = ArrayName.Length
    '            ReDim Preserve ArrayName(ii)
    '            ArrayName.SetValue(LastValue, ii)
    '        End If
    '    Catch ex As Exception
    '        QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Decision, ByVal LastValue As Decision, Optional ByVal IgnoreIfExists As Boolean = False)")
    '    End Try
    '    Return ArrayName
    'End Function

    ''' <summary>
    ''' Add a object element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">DataColumn element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new DataColumn </returns>
    ''' <remarks></remarks>   
    Public Function ArrayAppend(ByRef ArrayName() As DataColumn, ByVal LastValue As DataColumn, Optional ByVal IgnoreIfExists As Boolean = False) As DataColumn()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As DataColumn, ByVal LastValue As DataColumn, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    ''' Add a datarow  element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">DataColumn element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new DataColumn </returns>
    ''' <remarks></remarks>   
    Public Function ArrayAppend(ByRef ArrayName() As DataRow, ByVal LastValue As DataRow, Optional ByVal IgnoreIfExists As Boolean = False) As DataRow()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As DataRow, ByVal LastValue As DataRow, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    ''' Add a control element at the last of an array
    ''' </summary>
    ''' <param name="ArrayName"> Array to be added</param>
    ''' <param name="LastValue">Control element added at last position </param>
    ''' <param name="IgnoreIfExists" >Ignore if item already exists in the array</param>
    ''' <returns> Output array after adding new control </returns>
    ''' <remarks></remarks>
    Public Function ArrayAppend(ByRef ArrayName() As Control, ByVal LastValue As Control, Optional ByVal IgnoreIfExists As Boolean = False) As Control()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mindx As Integer = -1
            If IgnoreIfExists = True Then
                mindx = Array.IndexOf(ArrayName, LastValue)
            End If
            If mindx = -1 Then
                Dim ii As Integer = ArrayName.Length
                ReDim Preserve ArrayName(ii)
                ArrayName.SetValue(LastValue, ii)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayAppend(ByRef ArrayName() As Control, ByVal LastValue As Control, Optional ByVal IgnoreIfExists As Boolean = False)")
        End Try
        Return ArrayName
    End Function




    ''' <summary>
    ''' Get index no of an element of an array 
    ''' </summary>
    ''' <param name="ArrayToSearch">Array of string to searched</param>
    ''' <param name="Element">Element to be searched in the above array</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayFind(ByVal ArrayToSearch() As String, ByVal Element As String, Optional ByVal IgnoreCase As Boolean = True) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For indx = 0 To ArrayToSearch.Count - 1
                Dim mItem As String = IIf(IgnoreCase = True, UCase(ArrayToSearch(indx)), ArrayToSearch(indx))
                Dim mElement As String = IIf(IgnoreCase = True, UCase(Element), Element)
                If mItem = mElement Then
                    Return indx
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayFind(ByVal ArrayToSearch() As String, ByVal Element As String, Optional ByVal IgnoreCase As Boolean = True)")
        End Try
        Return -1
    End Function
    ''' <summary>
    ''' Get collection from an array collection. 
    ''' </summary>
    ''' <param name="ArrayCollection">Array of string to searched</param>
    ''' <param name="KeyName" >KeyName of collection</param>
    ''' <param name="KeyValue" >Key value to find in array collection</param>
    ''' <returns>Collection matched for keyname=keyvalue</returns>
    ''' <remarks></remarks>
    Public Function ArrayCollectionFind(ByVal ArrayCollection() As Collection, ByVal KeyName As String, ByVal KeyValue As String, Optional ByVal Indx As Integer = 0) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aa As New Collection
        Try
            For i = 0 To ArrayCollection.Count - 1
                Dim mItem As String = GetValueFromCollection(ArrayCollection(i), KeyName)
                If LCase(mItem) = LCase(KeyValue) Then
                    Return ArrayCollection(i)
                    Indx = i
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayCollectionFind(ByVal ArrayCollection() As Collection, ByVal KeyName As String, ByVal KeyValue As String, Optional ByVal Indx As Integer = 0)")
        End Try
        Return aa
    End Function
    ''' <summary>
    ''' Get collection from an array collection. 
    ''' </summary>
    ''' <param name="ArrayCollection">Array of string to searched</param>
    ''' <param name="KeyName" >KeyName of collection</param>
    ''' <param name="KeyValue" >Key value to find in array collection</param>
    ''' <returns>Collection matched for keyname=keyvalue</returns>
    ''' <remarks></remarks>
    Public Function ArrayCollectionFind(ByVal ArrayCollection() As Collection, ByVal KeyName As String, ByVal KeyValue As Object, Optional ByVal Indx As Integer = 0) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aa As New Collection
        Try
            For i = 0 To ArrayCollection.Count - 1
                Dim mItem As Object = GetValueFromCollection(ArrayCollection(i), KeyName)
                If mItem.Equals(KeyValue) = True Then
                    Return ArrayCollection(i)
                    Indx = i
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayCollectionFind(ByVal ArrayCollection() As Collection, ByVal KeyName As String, ByVal KeyValue As object, Optional ByVal Indx As Integer = 0)")
        End Try
        Return aa
    End Function



    ''' <summary>
    ''' Get index no of an element of an array
    ''' </summary>
    ''' <param name="ArrayToSearch"></param>
    ''' <param name="Element">Element to be searched in the above array</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayFind(ByVal ArrayToSearch() As Integer, ByVal Element As Integer) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For indx = 0 To ArrayToSearch.Count - 1
                If ArrayToSearch(indx) = Element Then
                    Return indx
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayFind(ByVal ArrayToSearch() As Integer, ByVal Element As Integer)")
        End Try
        Return -1
    End Function
    ''' <summary>
    ''' Get index no of an element of an array
    ''' </summary>
    ''' <param name="ArrayToSearch"></param>
    ''' <param name="Element">Element to be searched in the above array</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayFind(ByVal ArrayToSearch() As Decimal, ByVal Element As Decimal) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For indx = 0 To ArrayToSearch.Count - 1
                If ArrayToSearch(indx) = Element Then
                    Return indx
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayFind(ByVal ArrayToSearch() As Decimal, ByVal Element As Decimal)")
        End Try
        Return -1
    End Function

    ''' <summary>
    ''' Get index no of an element of an array
    ''' </summary>
    ''' <param name="ArrayToSearch"></param>
    ''' <param name="Element">Element to be searched in the above array</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayFind(ByVal ArrayToSearch() As Object, ByVal Element As Object) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For indx = 0 To ArrayToSearch.Count - 1
                If ArrayToSearch(indx) = Element Then
                    Return indx
                    Exit Function
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayFind(ByVal ArrayToSearch() As Decimal, ByVal Element As Decimal)")
        End Try
        Return -1
    End Function
    ''' <summary>
    ''' Error Message box before quitting the application
    ''' </summary>
    ''' <param name="ex">Error on exception </param>
    ''' <param name="err"> error object </param>
    ''' <remarks></remarks>
    Public Sub QuitError(ByVal ex As Exception, ByVal err As ErrObject, ByVal ErrorString As String)
        If LCase("WebAzure,WebGodaddy,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            MsgBox("ERROR_MESSAGE ( " & ex.Message & " )" & vbCrLf & vbCrLf & "STACK_TRACE  (" & ex.StackTrace.ToString & " )" & vbCrLf & vbCrLf & "Procedure " & err.Erl.ToString & vbCrLf & vbCrLf & "ERROR STMT:  " & ErrorString)
            'MsgBox(Application.ProductName)
            Application.ExitThread()
            Application.Exit()
            Process.GetCurrentProcess.Kill()
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
    Public Sub QuitMessage(ByVal MessageString As String, ByVal QuitProcedure As String, Optional ByVal MessageTitle As String = "Warning")
        If LCase("WebAzure,WebGodaddy,WebCloud").Contains(LCase(GlobalControl.Variables.SaralType)) = False Or GlobalControl.Variables.SaralType.Trim.Length = 0 Then
            If MessageString.Length > 0 Then
                MsgBox(MessageString & vbCrLf & vbCrLf & QuitProcedure, MsgBoxStyle.DefaultButton1, MessageTitle)
            End If
            Application.ExitThread()
            Application.Exit()
            Process.GetCurrentProcess.Kill()
        End If
    End Sub
    ''' <summary>
    ''' To convert date to string as per given format
    ''' </summary>
    ''' <param name="InputDate"> Date to be converted</param>
    ''' <param name="DateFormat ">Custom date format,eg "dd/MM/yyyy" </param>
    ''' <param name="DisplayDate "> display date as string in above custome format</param>
    ''' <returns>String output the the format "yyyyMMdd" </returns>
    ''' <remarks></remarks>
    Public Function DateTostring(ByVal InputDate As Date, Optional ByVal DateFormat As String = "", Optional ByRef DisplayDate As String = "") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        DateTostring = ""
        Try
            Dim lday As String = Microsoft.VisualBasic.Right(CStr(100 + InputDate.Day), 2)
            Dim lmon As String = Microsoft.VisualBasic.Right(CStr(100 + InputDate.Month), 2)
            Dim lyear As String = CStr(InputDate.Year)
            DateTostring = lyear & lmon & lday
            If Not DateFormat = "" Then
                DisplayDate = InputDate.ToString(DateFormat)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.DateTostring(ByVal InputDate As Date, Optional ByVal DateFormat As String = "", Optional ByRef DisplayDate As String = "")")
        End Try
    End Function
    ''' <summary>
    ''' Get dos file name from input string name
    ''' </summary>
    ''' <param name="LName"> Input string to be converted to dos file name</param>
    ''' <param name="LSize">Size of dos file name</param>
    ''' <returns>Output file name</returns>
    ''' <remarks></remarks>
    Public Function GetDosFileName(ByVal LName As String, ByVal LSize As Integer) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mpart As String = ""
        Try
            LName = UCase(LName)
            Dim Gstring As String = "ABCDEFGHIJKLMNOPQRSTUVWXY0123456789_"
            For i = 1 To LSize
                mpart = mpart + IIf(InStr(Gstring, Mid(LName, i, 1)) > 0, Mid(LName, i, 1), "")
            Next
            mpart = Left(mpart & StrDup(LSize, "0"), LSize)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetDosFileName(ByVal LName As String, ByVal LSize As Integer)")
        End Try
        Return mpart
    End Function

    Public Function AddItemIfUniqueHashtable(ByVal ht1 As Hashtable, ByVal hvalue As Object) As Hashtable
        Dim ValFound As Boolean = False
        For i = 0 To ht1.Count - 1
            Dim HashVal As Object = ht1.Item(ht1.Keys(i))
            If hvalue.GetType.Equals(HashVal.GetType) Then
                If hvalue = HashVal Then
                    ValFound = True
                    Exit For
                End If
            End If
        Next
        If ValFound = False Then
            Dim KeyCol As ICollection = ht1.Keys
            Dim Keylist As New List(Of Integer)
            For i = 0 To KeyCol.Count - 1
                Dim keyval As Integer = Convert.ToInt32(KeyCol(i))
                Keylist.Add(keyval)
                Console.WriteLine(keyval)
            Next
            Console.WriteLine(Keylist.Count & " count")
            Dim keyvalfinal As String = ""
            If Keylist.Count = 0 Then
                Console.WriteLine(" if count 0")
                keyvalfinal = "1"
            Else
                Keylist.Sort()
                Console.WriteLine(" if count not 0")
                keyvalfinal = Keylist(Keylist.Count - 1) + 1.ToString
            End If


            '  IIf(Keylist.Count = 0, "1", ()
            ht1.Add(keyvalfinal, hvalue)
        End If
        Return ht1
    End Function



    Private Function StringCast(ByVal mparam As Object) As String
        Dim mvalue As String = "*****"

        Try
            Dim mtype As String = LCase(mparam.GetType.Name.ToString)
            If mtype = "string" Or mtype = "decimal" Or mtype = "int32" Or mtype = "date" Or mtype = "datetime" Or mtype = "int16" Or mtype = "int64" Or mtype = "integer" Then
                mvalue = mparam.ToString
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringCast(ByVal mparam As Object)")
        End Try
        Return mvalue
    End Function


    Public Function GetParametersValueLine(Optional ByVal param1 As Object = Nothing, Optional ByVal param2 As Object = Nothing, Optional ByVal param3 As Object = Nothing, Optional ByVal param4 As Object = Nothing, Optional ByVal param5 As Object = Nothing, Optional ByVal param6 As Object = Nothing, Optional ByVal param7 As Object = Nothing, Optional ByVal param8 As Object = Nothing, Optional ByVal param9 As Object = Nothing, Optional ByVal param10 As Object = Nothing, Optional ByVal param11 As Object = Nothing, Optional ByVal param12 As Object = Nothing, Optional ByVal param13 As Object = Nothing, Optional ByVal param14 As Object = Nothing, Optional ByVal param15 As Object = Nothing) As String
        Dim lstr As String = ""
        Try
            If Not param1 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param1)
            End If
            If Not param2 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param2)
            End If
            If Not param3 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param3)
            End If
            If Not param4 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param4)
            End If
            If Not param5 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param5)
            End If
            If Not param6 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param6)
            End If
            If Not param7 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param7)
            End If
            If Not param8 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param8)
            End If
            If Not param9 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param9)
            End If
            If Not param10 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param10)
            End If
            If Not param11 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param11)
            End If
            If Not param12 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param12)
            End If
            If Not param13 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param13)
            End If
            If Not param14 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param14)
            End If
            If Not param15 Is Nothing Then
                lstr = lstr & IIf(lstr.Length = 0, "", ChrW(218)) & StringCast(param15)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetParametersValueLine")

        End Try


        Return lstr
    End Function


    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="TxtFile">Output full file name </param>
    ''' <param name="TxtString">Sring to be written</param>
    '''<param name="AddLast">True,if string is added to an existing file </param>
    ''' <returns> string to be written</returns>
    ''' <remarks></remarks>
    Public Function StringWrite(ByVal TxtFile As String, ByVal TxtString As String, Optional ByVal AddLast As Boolean = False, Optional ByVal mencoding As System.Text.Encoding = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim retstr As String = ""
        Try
            Dim fsw As System.IO.FileStream
            Dim Openmode As System.IO.FileMode = FileMode.Create
            If AddLast = True And System.IO.File.Exists(TxtFile) Then
                Openmode = FileMode.Append
                fsw = New System.IO.FileStream(TxtFile, Openmode, IO.FileAccess.Write)
            Else
                fsw = New System.IO.FileStream(TxtFile, Openmode, IO.FileAccess.ReadWrite)
            End If
            Dim sw As New System.IO.StreamWriter(fsw)
            If mencoding IsNot Nothing Then
                sw = New System.IO.StreamWriter(fsw, mencoding)
            End If
            sw.Write(TxtString)
            sw.Flush()
            sw.Close()
            fsw.Close()
            retstr = TxtFile
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWrite")
        End Try
        Return retstr
    End Function
    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="LFileStream">FileStream as already defined</param>
    ''' <param name="TxtString">Text String to be written</param>
    ''' <returns>Text string has been written</returns>
    ''' <remarks></remarks>
    Public Function StringWrite(ByRef LFileStream As FileStream, ByVal TxtString As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim retstr As String = ""
        Try
            Dim sw As New System.IO.StreamWriter(LFileStream)
            sw.Write(TxtString)
            sw.Flush()
            sw.Close()
            retstr = TxtString
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWrite(ByRef LFileStream As FileStream, ByVal TxtString As String)")
        End Try
        Return retstr
    End Function

    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="TxtFile">Output full file name </param>
    ''' <param name="TxtString">Array of text lines to be written</param>
    '''<param name="AddLast">True,if string is added to an existing file </param>
    ''' <returns> string to be written</returns>
    ''' <remarks></remarks>
    Public Function StringWriteLine(ByVal TxtFile As String, ByVal TxtString() As String, Optional ByVal AddLast As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim retstr As String = ""
        Try
            Dim fsw As System.IO.FileStream
            Dim Openmode As System.IO.FileMode = FileMode.Create
            If AddLast = True And System.IO.File.Exists(TxtFile) Then
                Openmode = FileMode.Append
                fsw = New System.IO.FileStream(TxtFile, Openmode, IO.FileAccess.Write)
            Else
                fsw = New System.IO.FileStream(TxtFile, Openmode, IO.FileAccess.ReadWrite)
            End If
            Dim sw As New System.IO.StreamWriter(fsw)
            For i = 0 To TxtString.Count - 1
                sw.WriteLine(TxtString(i))
            Next
            sw.Flush()
            sw.Close()
            fsw.Close()
            retstr = TxtFile
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWriteLine(ByVal TxtFile As String, ByVal TxtString() As String, Optional ByVal AddLast As Boolean = False)")
        End Try
        Return retstr
    End Function
    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="TxtFile">Output full file name </param>
    ''' <param name="TxtString">Array of text lines to be written</param>
    ''' <param name="Mencoding" >Text Format as system.text.encoding</param>
    '''<param name="AddLast">True,if string is added to an existing file </param>
    ''' <returns> string to be written</returns>
    ''' <remarks></remarks>
    Public Function StringWriteAllLines(ByVal TxtFile As String, ByVal TxtString() As String, ByVal Mencoding As System.Text.Encoding, Optional ByVal AddLast As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If AddLast = True And System.IO.File.Exists(TxtFile) Then
                Dim txtstring1() As String = StringReadAllLines(TxtFile, Mencoding)
                TxtString = AddTwoArrays(TxtString, txtstring1)
            End If
            File.WriteAllLines(TxtFile, TxtString, Mencoding)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWriteLine(ByVal TxtFile As String, ByVal TxtString() As String, ByVal Mencoding As System.Text.Encoding, Optional ByVal AddLast As Boolean = False)")
        End Try
        Return TxtFile
    End Function




    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="LFileStream">FileStream as already defined</param>
    ''' <param name="TxtString">Array of text lines to be written</param>
    ''' <returns>Text string has been written</returns>
    ''' <remarks></remarks>
    Public Function StringWriteLine(ByRef LFileStream As FileStream, ByVal TxtString() As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim retstr As String = ""
        Try
            Dim sw As New System.IO.StreamWriter(LFileStream)
            For i = 0 To TxtString.Count - 1
                sw.WriteLine(TxtString)
            Next
            sw.Flush()
            sw.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWriteLine(ByRef LFileStream As FileStream, ByVal TxtString() As String)")
        End Try
        Return retstr
    End Function
    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="LStreamWriter">StreamWriter  already defined</param>
    ''' <param name="TxtList">List  of text lines to be written</param>
    ''' <returns>Text string has been written</returns>
    ''' <remarks></remarks>
    Public Function StringWriteLine(ByRef LStreamWriter As StreamWriter, ByVal TxtList As List(Of String)) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim retstr As Boolean = False
        Try
            For i = 0 To TxtList.Count - 1
                LStreamWriter.WriteLine(TxtList(i))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWriteLine(ByRef LStreamWriter As StreamWriter, ByVal TxtList As List(Of String)) As Boolean")
        End Try
        Return retstr
    End Function
    ''' <summary>
    ''' Write string to a file
    ''' </summary>
    ''' <param name="LStreamWriter">StreamWriter  already defined</param>
    ''' <param name="TxtLine">Text line to be written</param>
    ''' <returns>Text string has been written</returns>
    ''' <remarks></remarks>
    Public Function StringWriteLine(ByRef LStreamWriter As StreamWriter, ByVal TxtLine As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim retstr As Boolean = False
        Try
            LStreamWriter.WriteLine(TxtLine)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWriteLine(ByRef LStreamWriter As StreamWriter, ByVal TxtList As List(Of String)) As Boolean")
        End Try
        Return retstr
    End Function

    ''' <summary>
    ''' Read  contents of a file into string
    ''' </summary>
    ''' <param name="TxtFile">Full file name to be read </param>
    ''' <returns> String as output</returns>
    ''' <remarks></remarks>
    Public Function StringRead(ByVal TxtFile As String, Optional ByVal mEncoding As System.Text.Encoding = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim outstr As String = ""
        Try
            Dim fs As System.IO.FileStream = New System.IO.FileStream(TxtFile, FileMode.Open, FileAccess.ReadWrite)
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(fs)
            If mEncoding IsNot Nothing Then
                sr = New System.IO.StreamReader(fs, mEncoding)
            End If
            Dim NBuff(0) As Char
            Do While sr.Peek() >= 0
                sr.Read(NBuff, 0, 1)
                If AscW(NBuff(0)) > 0 Then
                    outstr = outstr & IIf(AscW(NBuff(0)) > 0, NBuff(0), "")
                End If
            Loop
            sr.Close()
            fs.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringWriteLine(ByRef LFileStream As FileStream, ByVal TxtString() As String)")
        End Try
        Return outstr
    End Function


    ''' <summary>
    ''' Read  line contents of a file into string array
    ''' </summary>
    ''' <param name="TxtFile">Full file name to be read </param>
    ''' <returns> String as output</returns>
    ''' <remarks></remarks>
    Public Function StringReadLine(ByVal TxtFile As String) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim outstr() As String = {}
        Try
            Dim fs As System.IO.FileStream = New System.IO.FileStream(TxtFile, FileMode.Open, FileAccess.ReadWrite)
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(fs)
            sr.Peek()
            Do While sr.Peek() >= 0
                Dim str As String = sr.ReadLine
                ArrayAppend(outstr, str)
            Loop
            sr.Close()
            fs.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringReadLine(ByVal TxtFile As String)")
        End Try
        Return outstr
    End Function


    ''' <summary>
    ''' Read  line contents of a file into string array
    ''' </summary>
    ''' <param name="TxtFile">Full file name to be read </param>
    ''' <param name="mEncoding" >Text Encoding Format eg. utf-8</param>
    ''' <returns> String as output</returns>
    ''' <remarks></remarks>
    Public Function StringReadAllLines(ByVal TxtFile As String, ByVal mEncoding As System.Text.Encoding) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim outstr() As String = {}
        Try
            outstr = File.ReadAllLines(TxtFile, mEncoding)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringReadAllLine(ByVal TxtFile As String ByVal mEncoding As System.Text.Encoding)")
        End Try
        Return outstr
    End Function



    ''' <summary>
    ''' Read  line contents of a file into string array
    ''' </summary>
    ''' <param name="LFileStream"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StringReadLine(ByVal LFileStream As FileStream) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim outstr() As String = {}
        Try
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(LFileStream)
            Do While sr.Peek() >= 0
                Dim str As String = sr.ReadLine
                ArrayAppend(outstr, str)
            Loop
            sr.Close()
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringReadLine(ByVal LFileStream As FileStream)")
        End Try
        Return outstr
    End Function


    ''' <summary>
    ''' Convert string layout (key1+chrw(217)+val1+chrw(218)+key2+chrw(217)+val2+chrw(217)+key3+chrw(217)+val3) into hashtable object
    ''' </summary>
    ''' <param name="InputString">Input string of hashtable layout eg .(key1+chrw(217)+val1+chrw(218)+key2+chrw(217)+val2+chrw(217)+key3+chrw(217)+val3)</param>
    ''' <param name="VarHook"> Hashtable Keys separator default is chrw(218)</param>
    ''' <param name="ValHook">Separator of key and its value,default is chrw(217)</param>
    ''' <returns> Hash Table object</returns>
    ''' <remarks></remarks>
    Public Function GetHashTableFromString(ByVal InputString As String, Optional ByVal VarHook As String = ChrW(218), Optional ByVal ValHook As String = ChrW(217)) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            ValHook = IIf(InStr(InputString, "=") > 0, "=", ValHook)
            VarHook = IIf(InStr(InputString, "~") > 0, "~", VarHook)
            Dim LHashTable As New Hashtable
            If InputString.Trim.Length > 0 Then
                Dim ArrayVar() As String = InputString.Split(VarHook)
                For i = 0 To ArrayVar.Count - 1
                    Dim LArray() As String = ArrayVar(i).Split(ValHook)
                    LHashTable.Add(LCase(LArray(0)), LArray(1))
                Next
            End If
            Return LHashTable
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GetHashTableFromString(ByVal InputString As String, Optional ByVal VarHook As String = ChrW(218), Optional ByVal ValHook As String = ChrW(217)) As Hashtable")
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Create a new HashTable by filtering a MainHashTable on an array of keys already exist in MainHashTable.
    ''' </summary>
    ''' <param name="MainHashTable">MainHashTable from which new hashtable extracted.</param>
    ''' <param name="HashKeys">An array of hashkeys to be filtered.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetHashTableFromKeys(ByVal MainHashTable As Hashtable, ByVal HashKeys() As String) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim LHashTable As New Hashtable
            If MainHashTable.Count > 0 Then
                For i = 0 To HashKeys.Count - 1
                    Dim mKey As String = HashKeys(i)
                    Dim mvalue As Object = GetValueFromHashTable(MainHashTable, mKey)
                    If mvalue IsNot Nothing Then
                        LHashTable = AddItemToHashTable(LHashTable, mKey, mvalue)
                    End If
                Next
            End If
            Return LHashTable
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetHashTableFromKeys(ByVal MainHashTable As Hashtable, ByVal HashKeys() As String) As Hashtable")
        End Try
        Return Nothing
    End Function


    ''' <summary>
    ''' Convert hashtable object into string layout (key1+chrw(217)+val1+chrw(218)+key2+chrw(217)+val2+chrw(217)+key3+chrw(217)+val3)
    ''' </summary>
    ''' <param name="InputHashTable"> Input hash table object</param>
    ''' <param name="VarHook">Separator of two hashtable keys, Default value= chrw(218) </param>
    ''' <param name="ValHook">Separator of hashtable key and its item, Default value =  chrw(217)</param>
    ''' <returns>Output string of layout (key1+chrw(217)+val1+chrw(218)+key2+chrw(217)+val2+chrw(217)+key3+chrw(217)+val3)) </returns>
    ''' <remarks></remarks>
    Public Function GetStringFromHashTable(ByVal InputHashTable As Hashtable, Optional ByVal VarHook As String = ChrW(218), Optional ByVal ValHook As String = ChrW(217)) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lstr As String = ""
        Try
            For i = 0 To InputHashTable.Count - 1
                Lstr = Lstr & IIf(Lstr.Length = 0, "", VarHook) & LCase(InputHashTable.Keys(i)) & ValHook & InputHashTable.Item(LCase(InputHashTable.Keys(i)))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.HashTableToString(ByVal InputHashTable As Hashtable, Optional ByVal VarHook As String = ~, Optional ByVal ValHook As String = " = ")")
        End Try
        Return Lstr
    End Function
    ''' <summary>
    ''' Convert hashtable object into string
    ''' </summary>
    ''' <param name="InputHashTable"> Input hash table object</param>
    ''' <param name="SqlStringFormat" >Optional if True, String values are converted with single quotes otherwise with double quotes.</param>
    ''' <param name="LogicalOperator" > any logical operator default value  "=" , not equal to symbol ,less than symbol,greater than symbol </param>
    ''' <param name="LogicGate" >Logic gate placed between two columns default is " and "</param>
    ''' <returns>Output condition string  </returns>
    ''' <remarks></remarks>
    Public Function GetStringConditionFromHashTable(ByVal InputHashTable As Hashtable, Optional ByVal SqlStringFormat As Boolean = False, Optional ByVal LogicalOperator As String = "=", Optional ByVal LogicGate As String = " And ") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Lstr As String = ""
        If InputHashTable Is Nothing Then
            Return Lstr
            Exit Function
        End If
        Try
            For i = 0 To InputHashTable.Count - 1
                Dim mkey As String = InputHashTable.Keys(i)
                Dim mvalue As Object = InputHashTable.Item(mkey)
                If mvalue Is Nothing Then
                    Continue For
                End If
                Dim mtype As String = LCase(mvalue.GetType.Name)
                mvalue = mvalue.ToString
                If mtype = "string" Then
                    If SqlStringFormat = True Then
                        mvalue = "'" & mvalue & "'"
                    Else
                        mvalue = """" & mvalue & """"
                    End If
                End If
                If mtype = "datetime" Or mtype = "date" Then
                    If SqlStringFormat = True Then
                        mvalue = "'" & mvalue & "'"
                    Else
                        mvalue = "#" & mvalue & "#"
                    End If
                End If
                Lstr = Lstr & IIf(Lstr.Length = 0, "", " " & LogicGate & " ") & mkey & " " & LogicalOperator & " " & mvalue
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetStringConditionFromHashTable(ByVal InputHashTable As Hashtable, Optional ByVal LogicGate As String = " And ") As String")
        End Try
        Return Lstr
    End Function


    ''' <summary>
    ''' Get ServerDatabase from globalcontrol hashtable by server and database keys eg. 0_srv_0 and 0_mdf_0
    ''' </summary>
    ''' <param name="ServerKey">eg 0_srv_0 or 1_srv_1</param>
    ''' <param name="MdfKey">eg 0_mdf_0 or 1_mdf_1</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetServerDataBase(ByVal ServerKey As String, ByVal MdfKey As String) As String
        Dim mserver As String = GetValueFromHashTable(GlobalControl.Variables.AllServers, "ServerKey")
        Dim mdatabase As String = GetValueFromHashTable(GlobalControl.Variables.MDFFiles, "MdfKey")
        Return mserver & "." & mdatabase
    End Function


    ''' <summary>
    ''' Get Value from hash table of given key
    ''' </summary>
    ''' <param name="AHashTable">Hash table to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AlternateKeyName" >Alternate key if first key value not found, If Alternate Key is "Nothing" then returns nothing object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValueFromHashTable(ByVal AHashTable As Hashtable, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal As Object = AHashTable.Item(LCase(AKeyName))
        Try
            If IsDBNull(StrVal) Or StrVal Is Nothing Then
                Select Case True
                    Case AlternateKeyName = "Nothing"
                        StrVal = Nothing
                    Case AlternateKeyName <> ""
                        StrVal = AHashTable.Item(LCase(AlternateKeyName))
                    Case Else
                        StrVal = ""
                End Select
            End If

        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromHashTable(ByVal AHashTable As Hashtable, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "")")
        End Try
        Return StrVal
    End Function
    ''' <summary>
    ''' Get missing keys of  hash table
    ''' </summary>
    ''' <param name="AHashTable">Hash table to be searched</param>
    ''' <param name="KeyNames">Comma separated Keynames to find in hash table</param>
    ''' <returns>Return comma separated missing keynames in hash table </returns>
    ''' <remarks></remarks>
    Public Function MissingKeysInHashTable(ByVal AHashTable As Hashtable, ByVal KeyNames As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim missingkeys As String = ""
        Dim akeys() As String = KeyNames.Trim.Split(",")
        Try
            For k = 0 To akeys.Count - 1
                Dim akeyname As String = akeys(k).Trim
                Dim mflag As Boolean = False
                For i = 0 To AHashTable.Keys.Count - 1
                    Dim mkey As String = AHashTable.Keys(i).ToString.Trim
                    If LCase(mkey) = LCase(akeyname.Trim) Then
                        mflag = True
                        Exit For
                    End If
                Next
                If mflag = False Then
                    missingkeys = missingkeys & "," & akeyname
                End If
            Next
            missingkeys = missingkeys.Remove(0, 1)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.CheckKeysInHashTable(ByVal AHashTable As Hashtable, ByVal KeyNames As String) As String")
        End Try
        Return missingkeys
    End Function



    ''' <summary>
    ''' Get hash table key of a given given value
    ''' </summary>
    ''' <param name="AHashTable">Hash table to be searched</param>
    ''' <param name="mValue">Value to find in hash table</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetKeyFromHashTable(ByVal AHashTable As Hashtable, ByVal mValue As Object) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mkey As String = ""
        Try
            For i = 0 To AHashTable.Keys.Count - 1
                mkey = AHashTable.Keys(i)
                If AHashTable.Item(LCase(mkey)) = mValue Then
                    Return mkey
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetKeyFromHashTable(ByVal AHashTable As Hashtable, ByVal mValue As Object) As Object)")
        End Try
        Return mkey
    End Function
    ''' <summary>
    ''' Create a one key hashtable.
    ''' </summary>
    ''' <param name="mkey">Key of item</param>
    ''' <param name="mValue">Value of item</param>
    ''' <returns>A new HashTable </returns>
    ''' <remarks></remarks>
    Public Function CreateHashTable(ByVal mkey As String, ByVal mValue As Object) As Hashtable
        Dim mHash As New Hashtable
        mHash = AddItemToHashTable(mHash, mkey, mValue)
        Return mHash
    End Function
    ''' <summary>
    ''' Create a hashtable from all columns of a datarow. Where ColumnName is key and ColumnValue is its value
    ''' </summary>
    ''' <param name="mRow" >A DataRow which is converted into hashtable</param>
    ''' <param name="ExcludeColumns" >Name of rowcolumns to be excluded</param>
    ''' <param name="IncludeColumns" >Name of rowcolumns to be excluded,default is all</param>
    ''' <returns>A new HashTable </returns>
    ''' <remarks></remarks>
    Public Function CreateHashTable(ByVal mRow As DataRow, Optional ByVal ExcludeColumns As String = "", Optional ByVal IncludeColumns As String = "") As Hashtable
        Dim mHash As New Hashtable
        Dim mDt As DataTable = mRow.Table
        For i = 0 To mDt.Columns.Count - 1
            Dim mcolumn As String = LCase(mDt.Columns(i).ColumnName)
            If IncludeColumns.Length > 0 Then
                If InStr(LCase(IncludeColumns), mcolumn) = 0 Then
                    Continue For
                End If
            End If
            If ExcludeColumns.Length > 0 Then
                If InStr(LCase(ExcludeColumns), mcolumn) > 0 Then
                    Continue For
                End If
            End If
            mHash = AddItemToHashTable(mHash, mcolumn, mRow(i))
        Next
        Return mHash
    End Function
    ''' <summary>
    ''' Create a hashtable from all columns of a datarow. Where ColumnName is key and ColumnValue is its value
    ''' </summary>
    ''' <param name="mTable" >A DataTable whoose first row , which is converted into hashtable</param>
    ''' <param name="ExcludeColumns" >Name of rowcolumns to be excluded</param>
    ''' <param name="IncludeColumns" >Name of rowcolumns to be excluded,default is all</param>
    ''' <returns>A new HashTable </returns>
    ''' <remarks></remarks>
    Public Function CreateHashTable(ByVal mTable As DataTable, Optional ByVal ExcludeColumns As String = "", Optional ByVal IncludeColumns As String = "") As Hashtable
        Dim mHash As New Hashtable
        If mTable.Rows.Count > 0 Then
            For i = 0 To mTable.Columns.Count - 1
                Dim mcolumn As String = LCase(mTable.Columns(i).ColumnName)
                If IncludeColumns.Length > 0 Then
                    If InStr(LCase(IncludeColumns), mcolumn) = 0 Then
                        Continue For
                    End If
                End If
                If ExcludeColumns.Length > 0 Then
                    If InStr(LCase(ExcludeColumns), mcolumn) > 0 Then
                        Continue For
                    End If
                End If
                mHash = AddItemToHashTable(mHash, mcolumn, mTable.Rows(0).Item(i))
            Next
        End If
        Return mHash
    End Function

    ''' <summary>
    ''' Convert a Hashtable into datatable where key is ColumnNames and value is column value. 
    ''' </summary>
    ''' <param name="mHashTable">HashTable for which output datatable to be obtained</param>
    ''' <param name="ExtraFields">, separated string of Columns to be included in the string</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateDataTableFromHashTable(ByVal mHashTable As Hashtable, Optional ByVal ExtraFields As String = "") As DataTable
        Dim mdtable As New DataTable

        If mHashTable Is Nothing Then
            Return mdtable
        End If
        Dim df1 As New DataFunctions.DataFunctions
        For i = 0 To mHashTable.Count - 1
            Dim mkey As String = mHashTable.Keys(i)
            '   Dim mvalue As String = mHashTable.Item(mkey)
            mdtable = df1.AddColumnsInDataTable(mdtable, mkey)
        Next
        Dim mrow As DataRow = mdtable.NewRow
        For i = 0 To mHashTable.Count - 1
            Dim mkey As String = mHashTable.Keys(i)
            Dim mvalue As New Object
            mvalue = GetValueFromHashTable(mHashTable, mkey)   'mHashTable.Item(mkey)
            If Not mvalue Is Nothing Then
                mrow(mkey) = mvalue
            End If
        Next
        mdtable.Rows.Add(mrow)
        If ExtraFields.Length > 0 Then
            Dim aFields() As String = ExtraFields.Split(",")
            For i = 0 To aFields.Count - 1
                mdtable = df1.AddColumnsInDataTable(mdtable, aFields(i))
            Next
        End If
        Return mdtable
    End Function
    ''' <summary>
    ''' Get Value from a collection of given key
    ''' </summary>
    ''' <param name="ACollection">Collection to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AlternateKeyName" >Alternate key if first key value not found</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValueFromCollection(ByVal ACollection As Collection, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal As Object = ACollection.Item(LCase(AKeyName))
        Try
            If StrVal Is Nothing And AlternateKeyName <> "" Then
                StrVal = ACollection.Item(LCase(AlternateKeyName))
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromCollection(ByVal ACollection As Collection, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "")")
        End Try
        Return StrVal
    End Function
    ''' <summary>
    ''' Get colletion object from an array of collections by keyname and its value.
    ''' </summary>
    ''' <param name="ACollection">array of Collections to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AkeyValue" >Key Value to be searched in array of collections</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValueFromCollection(ByVal ACollection() As Collection, ByVal AKeyName As String, ByVal AkeyValue As String) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal As New Collection
        Try
            For i = 0 To ACollection.Count - 1
                If LCase(ACollection(i).Item(AKeyName)) = LCase(AkeyValue) Then
                    StrVal = ACollection(i)
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromCollection(ByVal ACollection As Collection, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "")")
        End Try
        Return StrVal
    End Function
    ''' <summary>
    ''' Get colletion object from an array of collections by keyname and its value.
    ''' </summary>
    ''' <param name="ACollection">array of Collections to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AkeyValue" >Key Value/comma separated keyvalues  to be searched in array of collections</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCollectionFromCollection(ByVal ACollection() As Collection, ByVal AKeyName As String, ByVal AkeyValue As String) As Collection()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal() As Collection = {}
        Dim keyvalues() As String = AkeyValue.Split(",")
        Try
            For i = 0 To ACollection.Count - 1
                For j = 0 To keyvalues.Count - 1
                    If LCase(ACollection(i).Item(AKeyName)) = LCase(keyvalues(j)) Then
                        ArrayAppend(StrVal, ACollection(i))
                    End If
                Next
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromCollection(ByVal ACollection() As Collection, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "")")
        End Try
        Return StrVal
    End Function



    ''' <summary>
    ''' Get colletion object from an array of collections by keyname and its value.
    ''' </summary>
    ''' <param name="ACollection">Array of Collections to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AkeyValue" >Key Value to be searched in array of collections</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValueFromCollection(ByVal ACollection() As Collection, ByVal AKeyName As String, ByVal AkeyValue As Integer) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal As New Collection
        Try
            For i = 0 To ACollection.Count - 1
                If ACollection(i).Item(AKeyName) = AkeyValue Then
                    StrVal = ACollection(i)
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromCollection(ByVal ACollection As Collection, ByVal AKeyName As String, Optional ByVal AlternateKeyName As String = "")")
        End Try
        Return StrVal
    End Function
    ''' <summary>
    ''' Get HashTable object from an array of HashTables by keyname and its value.
    ''' </summary>
    ''' <param name="AHashTable">Array of HashTables to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AkeyValue" >Key Value to be searched in array of hashtable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValueFromHashTable(ByVal AHashTable() As Hashtable, ByVal AKeyName As String, ByVal AkeyValue As Integer) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal As New Hashtable
        Try
            For i = 0 To AHashTable.Count - 1
                If AHashTable(i).Item(LCase(AKeyName)) = AkeyValue Then
                    StrVal = AHashTable(i)
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromHashTable(ByVal AHashTable() As Hashtable, ByVal AKeyName As String, ByVal AkeyValue As Integer) As Hashtable")
        End Try
        Return StrVal
    End Function
    ''' <summary>
    ''' Get HashTable object from an array of HashTables by keyname and its value.
    ''' </summary>
    ''' <param name="AHashTable">Array of HashTables to be searched</param>
    ''' <param name="AKeyName">Key name to find</param>
    ''' <param name="AkeyValue" >Key Value to be searched in array of hashtable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetValueFromHashTable(ByVal AHashTable() As Hashtable, ByVal AKeyName As String, ByVal AkeyValue As String) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim StrVal As New Hashtable
        Try
            For i = 0 To AHashTable.Count - 1
                If LCase(AHashTable(i).Item(LCase(AKeyName))) = LCase(AkeyValue) Then
                    StrVal = AHashTable(i)
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetValueFromHashTable(ByVal AHashTable() As Hashtable, ByVal AKeyName As String, ByVal AkeyValue As Integer) As Hashtable")
        End Try
        Return StrVal
    End Function




    ''' <summary>
    ''' Get Network folder accessible to the system.
    ''' </summary>
    ''' <param name="oFolderBrowserDialog"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetNetworkFolders(ByVal oFolderBrowserDialog As FolderBrowserDialog) As String
        Dim type As Type = oFolderBrowserDialog.[GetType]
        Dim fieldInfo As Reflection.FieldInfo = type.GetField("rootFolder", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
        fieldInfo.SetValue(oFolderBrowserDialog, DirectCast(18, Environment.SpecialFolder))
        If oFolderBrowserDialog.ShowDialog() = DialogResult.OK Then
            Return oFolderBrowserDialog.SelectedPath
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    '''Return a list of string having the items of full file name matching the wildcard criteria (*.*)  
    ''' </summary>
    ''' <param name="SourcePath"> Path to be searched </param>
    ''' <param name="WildCard"> wild card string such as "*.*" ,"*.dat" etc.,default value="*.* </param>
    ''' <param name="TopLevel"> False , if searched for child folder also,default value=False</param>
    ''' <returns>List object of fullfile name</returns>
    ''' <remarks></remarks>
    Public Function SearchFiles0(ByVal SourcePath As String, Optional ByVal WildCard As String = "*.*", Optional ByVal TopLevel As Boolean = False) As List(Of String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim filelist As New List(Of String)
        Try
            Dim aa As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(SourcePath)

            aa.GetFiles()
            For Each fi1 As System.IO.FileInfo In aa.GetFiles(WildCard)
                filelist.Add(UCase(Trim(fi1.FullName)))
            Next
            If TopLevel = False Then
                For Each dir1 As System.IO.DirectoryInfo In aa.GetDirectories
                    SearchFiles0(dir1.FullName.ToString, WildCard, TopLevel)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SearchFiles(ByVal SourcePath As String, Optional ByVal WildCard As String = *.*, Optional ByVal TopLevel As Boolean = False")
        End Try

        Return filelist
    End Function

    ''' <summary>
    '''Return a list of string having the items of full file name matching the wildcard criteria (*.*)  
    ''' </summary>
    ''' <param name="SourcePath"> Comma separated Folders or drives to be searched </param>
    ''' <param name="SearchPattern"> wild card string such as "*.*" ,"*.dat" etc.,default value="*.* </param>
    ''' <param name="TopLevel"> False , if searched for child folder also,default value=False</param>
    ''' <returns>List object of fullfile name</returns>
    ''' <remarks></remarks>
    Public Function SearchFiles(ByVal SourcePath As String, Optional ByVal SearchPattern As String = "*.*", Optional ByVal TopLevel As Boolean = False, Optional ByVal OnlyFirstFind As Boolean = True) As List(Of String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim filelist As New List(Of String)
        Dim aFolder() As String = SourcePath.Split(",")
        Dim wildcard As Boolean = IIf(SearchPattern.Contains("*") = True, True, False)
        wildcard = IIf(SearchPattern.Contains("?") = True, True, wildcard)
        Dim msource As String = ""
        Try
            For i = 0 To aFolder.Count - 1
                msource = aFolder(i)
                Dim aa As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(msource)
                ListFiles(filelist, SearchPattern, aa, OnlyFirstFind)
                If OnlyFirstFind = True And filelist.Count > 0 Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SearchFiles(ByVal SourcePath As String, Optional ByVal WildCard As String = *.*, Optional ByVal TopLevel As Boolean = False, Optional ByVal FirstFind As Boolean = False) As List(Of String)")
        End Try
        Return filelist
    End Function
    Private Sub ListFiles(ByVal lst As List(Of String), ByVal pattern As String, ByVal dir_info As DirectoryInfo, ByVal OnlyFirstFind As Boolean)
        ' Get the files in this directory. 
        Dim fs_infos() As FileInfo
        Try
            fs_infos = dir_info.GetFiles(pattern)
            For Each fs_info As FileInfo In fs_infos
                lst.Add(fs_info.FullName)
                If OnlyFirstFind = True Then
                    Exit Sub
                End If
            Next fs_info
        Catch ex As UnauthorizedAccessException
        End Try
        If OnlyFirstFind = True And lst.Count > 0 Then
            Exit Sub
        End If
        fs_infos = Nothing
        ' Search subdirectories. 
        Dim subdirs() As DirectoryInfo
        subdirs = dir_info.GetDirectories()
        Try
            For Each subdir As DirectoryInfo In subdirs
                ListFiles(lst, pattern, subdir, OnlyFirstFind)
            Next subdir
        Catch ex As UnauthorizedAccessException
        End Try
    End Sub

    ''' <summary>
    ''' To backward searching of a directory path for existence of a filename. 
    ''' </summary>
    ''' <param name="FullPath">Full path of searching backward</param>
    ''' <param name="FileName">File name to be searched, * wild cards are permissible </param>
    ''' <returns>Folder name / location where seaching file exists </returns>
    ''' <remarks></remarks>
    Public Function BackwardFileSearching(ByVal FullPath As String, ByVal FileName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DirName As String = Path.GetDirectoryName(FullPath.Trim)
        Try
            For i = DirName.Length To FullPath.Trim.IndexOf("\") Step -1
start:
                Dim arr() As String = Directory.GetFiles(DirName, FileName.Trim)
                If arr.Length > 0 Then
                    Return DirName
                    Exit Function
                Else
                    If DirName.Contains("\") = True Then
                        DirName = Microsoft.VisualBasic.Left(DirName, DirName.LastIndexOf("\"))
                        GoTo start
                    Else
                        Exit For
                    End If
                End If
            Next
            If File.Exists(DirName) = False Then
                Return ""
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.BackwardFileSearching(ByVal FullPath As String, ByVal FileName As String)")
        End Try
        Return ""
    End Function


    '''' <summary>
    '''' Remove character from a string
    '''' </summary>
    '''' <param name="InputString">Input string</param>
    '''' <param name="RemChar"> Character to be removed</param>
    '''' <returns>Output string</returns>
    '''' <remarks></remarks>
    'Public Function RemoveChar(ByVal InputString As String, Optional ByVal RemChar As String = "") As String
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
    '    'to remove letters
    '    Dim NewString As String = ""
    '    Try
    '        For i = 1 To InputString.Length
    '            If AscW(Mid(InputString, i, 1)) > 0 Then
    '                If (Not Mid(InputString, i, 1) = RemChar) Or RemChar.Length = 0 Then
    '                    NewString = NewString & Mid(InputString, i, 1)
    '                End If
    '            End If
    '        Next
    '    Catch ex As Exception
    '        QuitError(ex, Err, "Unable to execute GlobalFunction1.RemoveChar(ByVal InputString As String, Optional ByVal RemChar As String = "")")
    '    End Try
    '    Return NewString
    'End Function

    ''' <summary>
    ''' Remove characters from a string
    ''' </summary>
    ''' <param name="InputString">Input string</param>
    ''' <param name="RemChar"> comma separated characters to be removed</param>
    ''' <param name="RemoveAllExcept">comma separated characters which are left in the new string</param>
    ''' <returns>Output string</returns>
    ''' <remarks></remarks>
    Public Function RemoveChar(ByVal InputString As String, Optional ByVal RemChar As String = "", Optional ByVal RemoveAllExcept As String = "") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'to remove letters
        Dim NewString As String = ""
        Dim aremove() As String = RemChar.Split(",")
        Dim aexcept() As String = RemoveAllExcept.Split(",")
        Try
            For i = 1 To InputString.Length
                If AscW(Mid(InputString, i, 1)) > 0 Then
                    For j = 0 To aremove.Count - 1
                        If ((Not Mid(InputString, i, 1) = aremove(j)) Or aremove(j).Length = 0) And (InStr(RemoveAllExcept, Mid(InputString, i, 1)) > 0 Or RemoveAllExcept.Length = 0) Then
                            NewString = NewString & Mid(InputString, i, 1)
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "RemoveChar(ByVal InputString As String, Optional ByVal RemChar As String = "", Optional ByVal RemoveAllExcept As String = "") As String")
        End Try
        Return NewString
    End Function

    ''' <summary>
    ''' Convert fullfileName into a list object '0=path(dot(.)=for current directory,1=filename,2=extension,3=DriveLetter, if ExcludeDriveLetter =True
    ''' </summary>
    ''' <param name="FullFileName">Full file name as string </param>
    ''' <returns></returns>
    ''' <remarks> List object</remarks>
    Public Function FullFileNameToList(ByVal FullFileName As String, Optional ByVal ExcludeDriveLetter As Boolean = False) As List(Of String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        '0=patth,1=filename,2=extension
        Dim mpathnameext As New List(Of String)
        Try
            Dim mpath As String = ""
            Dim mfilename As String = ""
            Dim mext As String = ""
            Dim mdrive As String = ""
            Dim dd As Integer, ee As Integer
            If FullFileName.Trim.Length > 0 Then
                dd = FullFileName.LastIndexOf("\")
                mpath = Left(FullFileName, dd)
                ee = FullFileName.LastIndexOf(".")
                mext = Right(FullFileName, FullFileName.Length - ee - 1)
                mfilename = Mid(FullFileName, dd + 2, FullFileName.Length - mpath.Length - mext.Length - 2)
                If ExcludeDriveLetter = True Then
                    dd = mpath.IndexOf(":")
                    If dd > 0 Then
                        mdrive = Left(mpath, dd)
                        Dim mpath1 As String = Right(mpath, mpath.Length - dd - 1)
                        If Left(mpath1, 1) = "\" Then
                            mpath1 = Right(mpath1, mpath1.Length - 1)
                        End If
                        mpath = mpath1
                    End If
                End If
                mpathnameext.Add(mpath)
                mpathnameext.Add(mfilename)
                mpathnameext.Add(mext)
                mpathnameext.Add(mdrive)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.FullFileNameToList(ByVal FullFileName As String)")
        End Try
        Return mpathnameext
    End Function

    ''' <summary>
    ''' Convert fullfileName into a collection object of three string columns keys are "folder" ,"filename","extension" . If folder value is (.) then take current directory.
    ''' </summary>
    ''' <param name="FullFileName">Full file name as string </param>
    ''' <returns></returns>
    ''' <remarks> Collection object</remarks>
    Public Function FullFileNameToCollection(ByVal FullFileName As String) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        '0=patth,1=filename,2=extension
        Dim mpathnameext As New Collection
        Try
            If FullFileName.Trim.Length > 0 Then
                Dim mpath As String = Path.GetDirectoryName(FullFileName)
                Dim mfilename As String = Path.GetFileName(FullFileName)
                Dim mext As String = Path.GetExtension(FullFileName)
                mpathnameext.Add(mpath, "folder")
                mpathnameext.Add(mfilename, "filename")
                mpathnameext.Add(mext, "extension")
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.FullFileNameToCollection(ByVal FullFileName As String) As Collection")
        End Try

        Return mpathnameext
    End Function


    ''' <summary>
    ''' To convert a Formatted Date String to other formatted date string .
    ''' </summary>
    ''' <param name="InputDateString"> Date to be converted</param>
    ''' <param name="InputFormat ">Input custom date format,eg "dd/MM/yyyy" </param>
    ''' <param name="OutputFormat "> display date as string in above custome format</param>
    ''' <returns>String output the date  </returns>
    ''' <remarks></remarks>
    Public Function DateFormatConversion(ByVal InputDateString As String, ByVal InputFormat As String, ByVal OutputFormat As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DateTostring As String = ""
        Try
            Select Case LCase(InputFormat)
                Case "dd/mm/yyyy"
                    Dim ldate() As String = InputDateString.Split("/")
                    DateTostring = ldate(2) & ldate(1) & ldate(0)
                Case "dd-mm-yyyy"
                    Dim ldate() As String = InputDateString.Split("-")
                    DateTostring = ldate(2) & ldate(1) & ldate(0)

            End Select
        Catch ex As Exception
            QuitError(ex, Err, "DateFormatConversion(ByVal InputDateString As String, ByVal InputFormat As String, ByVal OutputFormat As String ")
        End Try

        Try
            Select Case LCase(OutputFormat)
                Case "yyymmdd"
                    DateTostring = DateTostring
            End Select
        Catch ex As Exception
            QuitError(ex, Err, "DateFormatConversion(ByVal InputDateString As String, ByVal InputFormat As String, ByVal OutputFormat As String ")
        End Try
        Return DateTostring
    End Function





    ''' <summary>
    ''' Add an item to a collection
    ''' </summary>
    ''' <param name="mCollection">Collection to be added</param>
    ''' <param name="ValueItem">Value of Item as object</param>
    ''' <param name="keyItem">Key of Item</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddItemToCollection(ByRef mCollection As Collection, ByVal ValueItem As Object, ByVal keyItem As String) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If mCollection.Contains(keyItem) = True Then
                mCollection.Remove(keyItem)
            End If
            mCollection.Add(ValueItem, keyItem)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.FullFileNameToCollection(ByVal FullFileName As String) As Collection")
        End Try
        Return mCollection
    End Function
    ''' <summary>
    ''' Add an item to a collection
    ''' </summary>
    ''' <param name="mCollection">Collection to be added</param>
    ''' <param name="ValueItem">Value of Item as string</param>
    ''' <param name="keyItem">Key of Item</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddItemToCollection(ByRef mCollection As Collection, ByVal ValueItem As String, ByVal keyItem As String) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If mCollection.Contains(keyItem) = True Then
                mCollection.Remove(keyItem)
            End If
            mCollection.Add(ValueItem, keyItem)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToCollection(ByRef mCollection As Collection, ByVal ValueItem As String, ByVal keyItem As String) As Collection")
        End Try

        Return mCollection
    End Function
    ''' <summary>
    ''' Add an item to a collection
    ''' </summary>
    ''' <param name="mCollection">Collection to be added</param>
    ''' <param name="ValueItem">Value of Item as integer</param>
    ''' <param name="keyItem">Key of Item</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddItemToCollection(ByRef mCollection As Collection, ByVal ValueItem As Integer, ByVal keyItem As String) As Collection
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try

            If mCollection.Contains(keyItem) = True Then
                mCollection.Remove(keyItem)
            End If
            mCollection.Add(ValueItem, keyItem)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToCollection(ByRef mCollection As Collection, ByVal ValueItem As Integer, ByVal keyItem As String) As Collection")

        End Try

        Return mCollection
    End Function


    ''' <summary>
    ''' Get file name with extension and path from a list control  '0=path(dot(.)=for current directory,1=filename,2=extension
    ''' </summary>
    ''' <param name="PathNameExt"> List control of size 3  '0=path(dot(.)=for current directory,1=filename,2=extension </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFullFileName(ByVal PathNameExt As List(Of String)) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        '0=patth,1=filename,2=extension
        Dim RetStr As String = ""
        Try
            If PathNameExt.Count >= 3 Then
                Dim mfolder As String = PathNameExt(0)
                Dim mfilename As String = PathNameExt(1)
                Dim mext As String = PathNameExt(2)
                mfolder = IIf(mfolder = ".", Environment.CurrentDirectory, mfolder)
                mfolder = mfolder & IIf(mfolder.Trim.Length > 0, IIf(Right(mfolder.Trim, 1) = "\", "", "\"), "")
                RetStr = mfolder & mfilename & "." & mext
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetFullFileName(ByVal PathNameExt As List(Of String)) As String")
        End Try

        Return RetStr
    End Function


    ''' <summary>
    ''' Get file name with extension and path from a collection object keys are "folder" ,"filename","extension" . If folder value is (.) then take current directory.
    ''' </summary>
    ''' <param name="PathNameExt"> List control of size 3  '0=path(dot(.)=for current directory,1=filename,2=extension </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFullFileName(ByVal PathNameExt As Collection) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        'Keys are folder,filename,extension
        Dim RetStr As String = ""
        Try
            If PathNameExt.Count >= 3 Then
                Dim mfolder As String = PathNameExt.Item("folder")
                Dim mfilename As String = PathNameExt.Item("filename")
                Dim mext As String = PathNameExt.Item("extension")
                mfolder = IIf(mfolder = ".", Environment.CurrentDirectory, mfolder)
                mfolder = mfolder & IIf(mfolder.Trim.Length > 0, IIf(Right(mfolder.Trim, 1) = "\", "", "\"), "")
                RetStr = mfolder & mfilename & "." & mext
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetFullFileName(ByVal PathNameExt As List(Of String)) As String")
        End Try
        Return RetStr
    End Function

    ''' <summary>
    ''' To make previous form controls enable with respect to a controlName
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet">Comma separated string of control names ,(*) for all controls of the form</param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="EnableTrue">Enable value ,Default value =True</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>
    Public Sub EnablePreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LformControlsSet As String = FormControlsSet
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")

            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next

            For i = 0 To Acontrol.Count - 1
                If Acontrol(i) = LCase(ControlName) Then
                    If IncludeCurrentControl = True Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.enabled = EnableTrue
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                    Exit For
                Else
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.enabled = EnableTrue
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.EnablePreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try
    End Sub

    ''' <summary>
    ''' To get previous form control as object with respect to a controlName
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet">Comma separated string of control names ,(*) for all controls of the form</param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="NoOfControls" >No of control to be fetched</param>
    '''<param name="OnlyVisibleEnabledControls" >To get the previous control as VisibleEnabled ,default is false </param>
    ''' <returns >An array of objects </returns>
    ''' <remarks></remarks>
    Public Function GetPreviousControlObject(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, ByVal NoOfControls As Integer, Optional ByVal OnlyVisibleEnabledControls As Boolean = False) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcontrol() As Object = {}
        Try
            Dim LformControlsSet As String = FormControlsSet
            Dim Vlcontrol As Object = ControlNameToObject(FormName, ControlName)
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next

            Dim ii As Integer = -1
            For i = 0 To Acontrol.Count - 1
                If Acontrol(i) = LCase(ControlName) Then
                    ii = i
                    Exit For
                End If
            Next
            If ii > 0 Then
                Dim k As Integer = 0
                For iii = ii - 1 To 0 Step -1
                    k = k + 1
                    If k > NoOfControls Then
                        Exit For
                    End If
                    Vlcontrol = ControlNameToObject(FormName, Acontrol(iii))
                    If OnlyVisibleEnabledControls = True Then
                        If Vlcontrol.Visible = False And Vlcontrol.Enabled = False Then
                            Continue For
                        End If
                    End If
                    ArrayAppend(lcontrol, Vlcontrol)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetPreviousControlObject(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, ByVal NoOfControls As Integer, Optional ByVal OnlyVisibleEnabledControls As Boolean = False) As Object()")
        End Try

        Return lcontrol
    End Function
    ''' <summary>
    ''' To get previous form control as object with respect to a controlName
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet">Comma separated string of control names ,(*) for all controls of the form</param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    '''<param name="OnlyVisibleEnabledControls" >To get the previous control as VisibleEnabled ,default is false </param>
    ''' <returns >A Control object returned </returns>
    ''' <remarks></remarks>
    Public Function GetPreviousControlObject(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal OnlyVisibleEnabledControls As Boolean = False) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim LformControlsSet As String = FormControlsSet
        Dim lcontrol As Object = ControlNameToObject(FormName, ControlName)
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            Dim ii As Integer = -1
            For i = 0 To Acontrol.Count - 1
                If Acontrol(i) = LCase(ControlName) Then
                    ii = i
                    Exit For
                End If
            Next
            If ii > 0 Then
                For iii = ii - 1 To 0 Step -1
                    lcontrol = ControlNameToObject(FormName, Acontrol(iii))
                    If OnlyVisibleEnabledControls = True Then
                        If lcontrol.Visible = False Or lcontrol.Enabled = False Then
                            Continue For
                        End If
                    End If
                    Exit For
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetPreviousControlObject(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal OnlyVisibleEnabledControls As Boolean = False) As Object()")
        End Try

        Return lcontrol
    End Function

    ''' <summary>
    ''' To get previous passed control as object with respect to a controlName
    ''' </summary>
    '''<param name="PassingControls" >A hash table having the controls which has passed by the user</param>
    '''<param name="ControlName" >Control name from which previous seraching executed</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <returns ></returns>
    Public Function GetPreviousControlObject(ByVal PassingControls As Hashtable, ByVal ControlName As String, Optional ByVal ControlSequence As String = "L") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcontrol As New Object
        Try
            Dim PrevKey As String = LCase(ControlName)
            Dim mControls As New List(Of Collection)
            For i = 0 To PassingControls.Keys.Count - 1
                Dim mkey As String = PassingControls.Keys(i)
                Dim ctrl1 As Object = PassingControls.Item(mkey)
                Dim lCollection As New Collection
                lCollection.Add(ctrl1.Name, "name")
                lCollection.Add(ctrl1.Left, "left")
                lCollection.Add(ctrl1.Top, "top")
                lCollection.Add(ctrl1.TabIndex, "tabindex")
                lCollection.Add(ctrl1, "object")
                mControls.Add(lCollection)
            Next
            Dim mControls1() As Collection = mControls.ToArray
            If ControlSequence = "L" Then
                Dim mSortKeys() As String = {"left", "top"}
                mControls1 = SortCollection(mControls1, mSortKeys)
            End If
            If ControlSequence = "T" Then
                mControls1 = SortCollection(mControls1, "tabindex")
            End If

            For i = 0 To mControls1.Count - 1
                If LCase(mControls1(i).Item("name")).ToString = LCase(ControlName) Then
                    Exit For
                End If
                PrevKey = mControls1(i).Item("name").ToString
            Next
            lcontrol = GetValueFromHashTable(PassingControls, PrevKey)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetPreviousControlObject(ByVal PassingControls As Hashtable, ByVal ControlName As String, Optional ByVal ControlSequence As String = L) As Object")
        End Try
        Return lcontrol
    End Function
    ''' <summary>
    ''' To get previous passed controls() as object with respect to a controlName
    ''' </summary>
    '''<param name="PassingControls" >A hash table having the controls which has passed by the user</param>
    '''<param name="ControlName" >Control name from which previous seraching executed</param>
    ''' <param name="NoOfControls" >No. of controls returned</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <returns >An array of controls</returns>
    Public Function GetPreviousControlObject(ByVal PassingControls As Hashtable, ByVal ControlName As String, ByVal NoOfControls As Integer, Optional ByVal ControlSequence As String = "L", Optional ByVal OnlyVisibleEnabled As Boolean = False) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcontrol() As Object = {}
        Try
            Dim PrevKey As String = LCase(ControlName)
            Dim mControls As New List(Of Collection)
            For i = 0 To PassingControls.Keys.Count - 1
                Dim mkey As String = PassingControls.Keys(i)
                Dim ctrl1 As Object = PassingControls.Item(mkey)
                Dim lCollection As New Collection
                lCollection.Add(ctrl1.Name, "name")
                lCollection.Add(ctrl1.Left, "left")
                lCollection.Add(ctrl1.Top, "top")
                lCollection.Add(ctrl1.TabIndex, "tabindex")
                lCollection.Add(ctrl1, "object")
                mControls.Add(lCollection)
            Next
            Dim mControls1() As Collection = mControls.ToArray
            If ControlSequence = "L" Then
                Dim mSortKeys() As String = {"left", "top"}
                mControls1 = SortCollection(mControls1, mSortKeys)
            End If
            If ControlSequence = "T" Then
                mControls1 = SortCollection(mControls1, "tabindex")
            End If
            Dim ii As Integer = 0
            For i = 0 To mControls1.Count - 1
                If LCase(mControls1(i).Item("name")).ToString = LCase(ControlName) Then
                    ii = i
                    Exit For
                End If
            Next
            If ii > 0 Then
                Dim k As Integer = 0
                For iii = ii - 1 To 0 Step -1
                    k = k + 1
                    If k > NoOfControls Then
                        Exit For
                    End If
                    PrevKey = LCase(mControls1(iii).Item("name").ToString)
                    Dim vlcontrol As Object = GetValueFromHashTable(PassingControls, PrevKey)
                    If OnlyVisibleEnabled = True Then
                        If vlcontrol.visible = False Or vlcontrol.enabled = False Then
                            Continue For
                        End If
                    End If
                    ArrayAppend(lcontrol, vlcontrol)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetPreviousControlObject(ByVal PassingControls As Hashtable, ByVal ControlName As String, Optional ByVal ControlSequence As String = L) As Object")
        End Try
        Return lcontrol
    End Function
    ''' <summary>
    ''' To get next passed control as object with respect to a controlName
    ''' </summary>
    '''<param name="PassingControls" >A hash table having the controls which has passed by the user</param>
    '''<param name="ControlName" >Control name from which previous seraching executed</param>
    Public Function GetNextControlObject(ByVal PassingControls As Hashtable, ByVal ControlName As String) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcontrol As Object = Nothing
        Try
            Dim NextKey As String = LCase(ControlName)
            For i = PassingControls.Keys.Count - 1 To 0 Step -1
                If LCase(PassingControls.Keys(i).ToString) = NextKey Then
                    Exit For
                End If
                NextKey = PassingControls.Keys(i)
            Next
            lcontrol = GetValueFromHashTable(PassingControls, NextKey)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetNextControlObject(ByVal PassingControls As Hashtable, ByVal ControlName As String) As Object")
        End Try

        Return lcontrol
    End Function

    ''' <summary>
    ''' To Add item in a hashtable with key and its value
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string of hashtable item</param>
    ''' <param name="ItemValue" >Item value of hash table item</param>
    Public Function AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As Object, Optional ByVal ReplaceIfKeyExists As Boolean = True) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return HashTableControl : Exit Function
        If HashTableControl Is Nothing Then Return Nothing : Exit Function
        Try
            Dim mKey As String = LCase(KeyValue)
            Dim ii As Integer = -1
            For i = 0 To HashTableControl.Keys.Count - 1
                If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                    ii = i
                    Exit For
                End If
            Next
            If ii = -1 Then
                HashTableControl.Add(mKey, ItemValue)
            Else
                ' HashTableControl.Item(LCase(KeyValue)) = ItemValue
                If ReplaceIfKeyExists = True Then
                    HashTableControl.Item(mKey) = ItemValue
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As Object) As Hashtable")
        End Try

        Return HashTableControl
    End Function

    ''' <summary>
    ''' To Add item in a hashtable with key and its value
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string hashtable item</param>
    ''' <param name="ItemValue" >Item value of hash table item</param>
    Public Function AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As Hashtable, Optional ByVal ReplaceIfKeyExists As Boolean = True) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return HashTableControl : Exit Function
        If HashTableControl Is Nothing Then Return Nothing : Exit Function

        Try
            Dim mKey As String = LCase(KeyValue)
            Dim ii As Integer = -1
            For i = 0 To HashTableControl.Keys.Count - 1
                If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                    ii = i
                    Exit For
                End If
            Next
            If ii = -1 Then
                HashTableControl.Add(mKey, ItemValue)
            Else
                ' HashTableControl.Item(LCase(KeyValue)) = ItemValue
                If ReplaceIfKeyExists = True Then
                    HashTableControl.Item(mKey) = ItemValue
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As Object) As Hashtable")
        End Try

        Return HashTableControl
    End Function


    ''' <summary>
    ''' To Add item in a hashtable with key and its value
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string of hashtable item</param>
    ''' <param name="ItemValue" >Item value of hash table item</param>
    Public Function AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As String, Optional ByVal ReplaceIfKeyExists As Boolean = True) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return HashTableControl : Exit Function
        If HashTableControl Is Nothing Then Return Nothing : Exit Function
        Try
            Dim mKey As String = LCase(KeyValue)
            Dim ii As Integer = -1
            For i = 0 To HashTableControl.Keys.Count - 1
                If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                    ii = i
                    Exit For
                End If
            Next
            If ii = -1 Then
                HashTableControl.Add(mKey, ItemValue)
            Else
                '  HashTableControl.Item(LCase(KeyValue)) = ItemValue
                If ReplaceIfKeyExists = True Then
                    HashTableControl.Item(mKey) = ItemValue
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As String) As Hashtable")
        End Try
        Return HashTableControl
    End Function
    ''' <summary>
    ''' To Add item in a hashtable with key and its value
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string of hashtable item</param>
    ''' <param name="ItemValue" >Item value of hash table item</param>
    Public Function AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As Integer, Optional ByVal ReplaceIfKeyExists As Boolean = True) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return HashTableControl : Exit Function
        If HashTableControl Is Nothing Then Return Nothing : Exit Function
        Try
            Dim mKey As String = LCase(KeyValue)
            Dim ii As Integer = -1
            For i = 0 To HashTableControl.Keys.Count - 1
                If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                    ii = i
                    Exit For
                End If
            Next
            If ii = -1 Then
                HashTableControl.Add(mKey, ItemValue)
            Else
                If ReplaceIfKeyExists = True Then
                    HashTableControl.Item(mKey) = ItemValue
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String, ByVal ItemValue As String) As Hashtable")
        End Try

        Return HashTableControl
    End Function
    ''' <summary>
    ''' Set Text to radiobutton items  from global MasterOptions datatable of radiobuttons Group
    ''' </summary>
    ''' <param name="OptionKeyAndChecked">Key of MasterOptions and checked position separated by ~ sign</param>
    ''' <param name="RadioButtonsContainer">Panel/Groupbox  of radiobuttons</param>
    ''' <param name="DtMasterOptions" >Datatable of master options</param>
    ''' <returns>A hashtable containing key as name of radiobutton and value as radiobutton controls in the panel</returns>
    ''' <remarks></remarks>

    Public Function AddTextToRadioButtonsGroup(ByVal OptionKeyAndChecked As String, ByRef RadioButtonsContainer As Control, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mHash As New Hashtable
        Try
            If DtMasterOptions Is Nothing Then
                DtMasterOptions = GlobalControl.Variables.MasterOptions
            End If
            If OptionKeyAndChecked.Length = 0 Then
                QuitMessage("No option key defined", "AddItemsToRadioButtonsPanel(ByVal OptionKey As Int16, ByRef mRadioButtonPanel As Panel, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable")
            End If
            Dim df1 As New DataFunctions.DataFunctions
            Dim aOptions() As String = OptionKeyAndChecked.Split("~")
            Dim OptionKey As Int16 = CInt(aOptions(0))
            Dim mpos As Int16 = 0
            If aOptions.Count > 1 Then
                mpos = CInt(aOptions(1))
            End If

            Dim mrow As DataRow = df1.FindRowByPrimaryCols(DtMasterOptions, OptionKey)
            If mrow IsNot Nothing Then
                Dim avaluesSet() As String = mrow("ValuesSet").ToString.Split("~")
                Dim rdbcount As Int16 = RadioButtonsContainer.Controls.OfType(Of RadioButton)().Count
                If avaluesSet.Count <> rdbcount Then
                    QuitMessage("No. of radiobuttons in the panel mismatched by radiobutton texts.", "AddItemsToRadioButtonsPanel(ByVal OptionKey As Int16, ByRef mRadioButtonPanel As Panel, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable")
                End If
                '  Dim aradb As Control.ControlCollection = mRadioButtonPanel.Controls
                Dim j As Int16 = 0
                For Each Ctr As RadioButton In RadioButtonsContainer.Controls.OfType(Of RadioButton)()
                    Ctr.Text = avaluesSet(j)
                    Ctr.Tag = j
                    Ctr.Checked = False
                    If j = mpos Then
                        Ctr.Checked = True
                    End If
                    mHash = AddItemToHashTable(mHash, Ctr.Name, Ctr)
                    j = j + 1
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemsToRadioButtonsPanel(ByVal OptionKey As Int16, ByRef mRadioButtonPanel As Panel, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable")
        End Try
        Return mHash
    End Function
    ''' <summary>
    ''' Get Radiobutton checked in a container.(Panel/Groupbox)
    ''' </summary>
    ''' <param name="RadioButtonsContainer">Panel/Groupbox  of radiobuttons</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetRadioButtonChecked(ByRef RadioButtonsContainer As Control) As RadioButton
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mradio As New RadioButton
        Try
            For Each Ctr As RadioButton In RadioButtonsContainer.Controls.OfType(Of RadioButton)()
                If Ctr.Checked = True Then
                    mradio = Ctr
                    Exit For
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetRadioButtonChecked(ByRef RadioButtonsContainer As Control) As RadioButton")
        End Try
        Return mradio
    End Function
    ''' <summary>
    ''' Set  Radiobutton checked in a container.(Panel/Groupbox)
    ''' </summary>
    ''' <param name="RadioButtonsContainer">Panel/Groupbox  of radiobuttons</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function SetRadioButtonChecked(ByRef RadioButtonsContainer As Control, ByVal PropertyIndex As Int16) As RadioButton
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mradio As New RadioButton
        Try
            For Each Ctr As RadioButton In RadioButtonsContainer.Controls.OfType(Of RadioButton)()
                If Ctr.Tag = PropertyIndex Then
                    Ctr.Checked = True
                    mradio = Ctr
                Else
                    Ctr.Checked = False
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SetRadioButtonChecked(ByRef RadioButtonsContainer As Control, ByVal PropertyIndex As Int16) As RadioButton")
        End Try
        Return mradio
    End Function



    ''' <summary>
    ''' Get CheckBoxes checked in a container.(Panel/Groupbox)
    ''' </summary>
    ''' <param name="RadioButtonsContainer">Panel/Groupbox  of Checkboxes</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetCheckBoxesChecked(ByRef RadioButtonsContainer As Control) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mradio As New Hashtable
        Try
            For Each Ctr As CheckBox In RadioButtonsContainer.Controls.OfType(Of CheckBox)()
                If Ctr.Checked = True Then
                    mradio = AddItemToHashTable(mradio, Ctr.Name, Ctr)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetCheckBoxesChecked(ByRef RadioButtonsContainer As Control) As Hashtable")
        End Try
        Return mradio
    End Function
    ''' <summary>
    ''' Set Text to radiobutto items  from global MasterOptions datatable of radiobuttons Group
    ''' </summary>
    ''' <param name="RadioButtonsContainer">Panel/Groupbox  of radiobuttons</param>
    ''' <param name="TextOptions" > A ~ separated string of containing of radiobutton texts</param>
    ''' <param name="RbnPosition" >positions in textoption of radiobutton checked </param>
    ''' <returns>A hashtable containing key as name of radiobutton and value as radiobutton controls in the panel</returns>
    ''' <remarks></remarks>
    ''' 


    Public Function AddTextToRadioButtonsGroup(ByRef RadioButtonsContainer As Control, ByVal TextOptions As String, ByVal RbnPosition As Int16) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mHash As New Hashtable
        Try
            Dim avaluesSet() As String = TextOptions.Split("~")
            Dim rdbcount As Int16 = RadioButtonsContainer.Controls.OfType(Of RadioButton)().Count
            If avaluesSet.Count <> rdbcount Then
                QuitMessage("No. of Radiobuttons in the panel mismatched by checkboxes texts.", "AddTextToRadioButtonsGroup(ByRef RadioButtonsContainer As Control, ByVal TextOptions As String, ByVal RbnPosition As Int16) As Hashtable")
            End If
            Dim j As Int16 = 0

            For Each Ctr As RadioButton In RadioButtonsContainer.Controls.OfType(Of RadioButton)()
                Ctr.Text = avaluesSet(j)
                Ctr.Tag = j
                Ctr.Checked = False
                If RbnPosition = j Then
                    Ctr.Checked = True
                End If
                mHash = AddItemToHashTable(mHash, Ctr.Name, Ctr)
                j = j + 1
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTextToRadioButtonsGroup(ByRef RadioButtonsContainer As Control, ByVal TextOptions As String, ByVal RbnPosition As Int16) As Hashtable")
        End Try
        Return mHash
    End Function

    ''' <summary>
    ''' This function returns reverse array of strings
    ''' </summary>
    ''' <param name="str">Array of string to reverse</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayReverse(ByVal str() As String) As String()
        Dim str2(str.Length - 1) As String
        If str.Length > 0 Then
            For i = 0 To str.Length - 1
                Dim any As Integer = (str.Length - 1) - i
                str2(i) = str((str.Length - 1) - i)
            Next
        End If
        Return str2
    End Function


    ''' <summary>
    ''' Set text to Checkbox items  from global MasterOptions datatable of CheckBox Group
    ''' </summary>
    ''' <param name="OptionKeyAndChecked">Key of MasterOptions and checked position separated by ~ sign,multiple checked positions are separated by comma(,)</param>
    ''' <param name="CheckBoxesContainer">Panel/Groupbox  of checkboxes</param>
    ''' <param name="DtMasterOptions" >Datatable of master options</param>
    ''' <returns>A hashtable containing key as name of checkboxes and values as checkbox controls in the panel/groupbox</returns>
    ''' <remarks></remarks>

    Public Function AddTextToCheckBoxesGroup(ByVal OptionKeyAndChecked As String, ByRef CheckBoxesContainer As Control, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mHash As New Hashtable
        Try
            If DtMasterOptions Is Nothing Then
                DtMasterOptions = GlobalControl.Variables.MasterOptions
            End If
            If OptionKeyAndChecked.Length = 0 Then
                QuitMessage("No option key defined", "AddTextToCheckBoxesGroup(ByVal OptionKeyAndChecked As String, ByRef CheckBoxesContainer As Control, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable")
            End If
            Dim df1 As New DataFunctions.DataFunctions
            Dim aOptions() As String = OptionKeyAndChecked.Split("~")
            Dim OptionKey As Int16 = CInt(aOptions(0))
            Dim apos() As Integer = {}
            If aOptions.Count > 1 Then
                Dim spos() As String = aOptions(1).Split(",")
                For u = 0 To spos.Count - 1
                    apos = ArrayAppend(apos, CInt(spos(u)))
                Next
            End If

            Dim mrow As DataRow = df1.FindRowByPrimaryCols(DtMasterOptions, OptionKey)
            If mrow IsNot Nothing Then
                Dim avaluesSet() As String = mrow("ValuesSet").ToString.Split("~")
                Dim rdbcount As Int16 = CheckBoxesContainer.Controls.OfType(Of CheckBox)().Count
                If avaluesSet.Count <> rdbcount Then
                    QuitMessage("No. of checkboxes in the panel/groupbox mismatched by checkbox texts.", "AddTextToCheckBoxesGroup(ByVal OptionKeyAndChecked As String, ByRef CheckBoxesContainer As Control, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable")
                End If
                '  Dim aradb As Control.ControlCollection = mRadioButtonPanel.Controls
                Dim j As Int16 = 0
                For Each Ctr As CheckBox In CheckBoxesContainer.Controls.OfType(Of CheckBox)()
                    Ctr.Text = avaluesSet(j)
                    Ctr.Tag = j
                    Ctr.Checked = False
                    If ArrayFind(apos, j) > -1 Then
                        Ctr.Checked = True
                    End If
                    mHash = AddItemToHashTable(mHash, Ctr.Name, Ctr)
                    j = j + 1
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTextToCheckBoxesGroup(ByVal OptionKeyAndChecked As String, ByRef CheckBoxesContainer As Control, Optional ByVal DtMasterOptions As DataTable = Nothing) As Hashtable")
        End Try
        Return mHash
    End Function

    ''' <summary>
    ''' Set text to Checkbox items  from global MasterOptions datatable of CheckBox Group
    ''' </summary>
    ''' <param name="CheckBoxesContainer">Panel/Groupbox  of radiobuttons</param>
    ''' <param name="TextOptions" > A ~ separated string of containing of radiobutton texts</param>
    ''' <param name="CheckedPositions" >Comma separated positions in textoption of check boxes checked </param>
    ''' <returns>A hashtable containing key as name of radiobutton and value as radiobutton controls in the panel</returns>
    ''' <remarks></remarks>

    Public Function AddTextToCheckBoxesGroup(ByRef CheckBoxesContainer As Control, ByVal TextOptions As String, ByVal CheckedPositions As String) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mHash As New Hashtable
        Try
            Dim avaluesSet() As String = TextOptions.Split("~")
            Dim rdbcount As Int16 = CheckBoxesContainer.Controls.OfType(Of CheckBox)().Count
            If avaluesSet.Count <> rdbcount Then
                QuitMessage("No. of checkboxes in the panel mismatched by checkboxes texts.", "AddTextToCheckBoxesGroup(ByRef CheckBoxesContainer As Control, ByVal TextOptions As String, ByVal CheckedPositions As String) As Hashtable")
            End If
            Dim j As Int16 = 0
            Dim apos() As Integer = {}
            If CheckedPositions.Length > 0 Then
                Dim aCheckedPosition() As String = CheckedPositions.Split(",")
                For u = 0 To aCheckedPosition.Count - 1
                    apos = ArrayAppend(apos, CInt(aCheckedPosition(u)))
                Next
            End If

            For Each Ctr As CheckBox In CheckBoxesContainer.Controls.OfType(Of CheckBox)()
                Ctr.Text = avaluesSet(j)
                Ctr.Tag = j
                Ctr.Checked = False
                If ArrayFind(apos, j) > -1 Then
                    Ctr.Checked = True
                End If
                mHash = AddItemToHashTable(mHash, Ctr.Name, Ctr)
                j = j + 1
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddTextToCheckBoxesGroup(ByRef CheckBoxesContainer As Control, ByVal TextOptions As String, ByVal CheckedPositions As String) As Hashtable")
        End Try
        Return mHash
    End Function


    ''' <summary>
    ''' To Add item in a hashtable with keys() and its values()
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string of hashtable item</param>
    ''' <param name="ItemValue" >Item value of hash table item</param>
    Public Function AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue() As String, ByVal ItemValue() As Object, Optional ByVal ReplaceIfKeyExists As Boolean = True) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return HashTableControl : Exit Function
        If HashTableControl Is Nothing Then Return Nothing : Exit Function
        If KeyValue.Count <> ItemValue.Count Then
            QuitMessage("Keyvalue and Itemvalue must be of same size array", "AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue() As String, ByVal ItemValue() As Object) As Hashtable ")
        End If
        Try
            For j = 0 To KeyValue.Count - 1
                Dim mKey As String = LCase(KeyValue(j))
                Dim ii As Integer = -1
                For i = 0 To HashTableControl.Keys.Count - 1
                    If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                        ii = i
                        Exit For
                    End If
                Next
                If ii = -1 Then
                    HashTableControl.Add(LCase(KeyValue(j)), ItemValue(j))
                Else
                    '   HashTableControl.Item(LCase(KeyValue(j))) = ItemValue(j)
                    If ReplaceIfKeyExists = True Then
                        HashTableControl.Item(LCase(KeyValue(j))) = ItemValue(j)
                    End If

                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.AddItemToHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue() As String, ByVal ItemValue() As Object) As Hashtable")
        End Try

        Return HashTableControl
    End Function
    ''' <summary>
    ''' To Remove item from a hashtable with key and its value
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string of hashtable item</param>
    Public Function RemoveItemFromHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mKey As String = LCase(KeyValue)
            Dim ii As Integer = -1
            For i = 0 To HashTableControl.Keys.Count - 1
                If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                    ii = i
                    Exit For
                End If
            Next
            If ii > -1 Then
                HashTableControl.Remove(mKey)
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.RemoveItemFromHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue As String) As Hashtable")
        End Try

        Return HashTableControl
    End Function
    ''' <summary>
    ''' To Remove item from a hashtable with keys() and  values()
    ''' </summary>
    '''<param name="HashTableControl" >A hash table control in which values to be added</param>
    '''<param name="KeyValue" >Key value as string of hashtable item</param>
    Public Function RemoveItemFromHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue() As String) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For j = 0 To KeyValue.Count - 1
                Dim mKey As String = LCase(KeyValue(j))
                Dim ii As Integer = -1
                For i = 0 To HashTableControl.Keys.Count - 1
                    If LCase(HashTableControl.Keys(i).ToString) = mKey Then
                        ii = i
                        Exit For
                    End If
                Next
                If ii > -1 Then
                    HashTableControl.Remove(mKey)
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.RemoveItemFromHashTable(ByRef HashTableControl As Hashtable, ByVal KeyValue() As String) As Hashtable")
        End Try
        Return HashTableControl
    End Function

    ''' <summary>
    ''' To get next form control as object with respect to a controlName
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet">Comma separated string of control names ,(*) for all controls of the form</param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    '''<param name="VisibleControl" >To get the previous control as visible </param>
    ''' <param name="EnabledControl" >To get the previous control as enabled</param>
    ''' <remarks></remarks>
    Public Function GetNextControlObject(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleControl As Boolean = True, Optional ByVal EnabledControl As Boolean = True) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim lcontrol As Object = Nothing
        Dim LformControlsSet As String = FormControlsSet
        lcontrol = ControlNameToObject(FormName, ControlName)
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next

            Dim ii As Integer = -1
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
                If Acontrol(i) = LCase(ControlName) Then
                    ii = i
                    Exit For
                End If
            Next
            If ii > 0 Then
                Dim iii As Integer = ii + 1
                While iii <= Acontrol.Count - 1
                    lcontrol = ControlNameToObject(FormName, Acontrol(iii))
                    If lcontrol.Visible = VisibleControl And lcontrol.Enabled = EnabledControl Then
                        Exit While
                    End If
                    iii = iii + 1
                End While
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetNextControlObject(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleControl As Boolean = True, Optional ByVal EnabledControl As Boolean = True) As Object")
        End Try
        Return lcontrol
    End Function

    ''' <summary>
    ''' To make next form controls enable with respect to a ControlName  
    ''' </summary>
    ''' <param name="FormName"> Parent form</param>
    ''' <param name="FormControlsSet">Comma separated string of control names,(*) for all controls of the form </param>
    ''' <param name="ControlName">Control name for which next controls are to be  considered</param>
    ''' <param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array.</param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="EnableTrue">Enable value ,Default value =True</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>
    Public Sub EnableNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            '* for all controls
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            Dim k As Integer = Array.IndexOf(Acontrol, LCase(ControlName))
            If k > -1 Then
                k = IIf(IncludeCurrentControl = True, k, k + 1)
                For i = k To Acontrol.Count - 1
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.enabled = EnableTrue
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.EnableNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try
    End Sub
    ''' <summary>
    ''' To make form controls enable
    ''' </summary>
    ''' <param name="FormName">Parent form , whose controls to make enable</param>
    ''' <param name="ControlNames">Comma separated string of control names to make enable,(*) for all controls of the form</param>
    ''' <param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array.</param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="EnableTrue">Enable value ,Default value =True</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not considered</param>
    ''' <remarks></remarks>
    Public Sub EnableControls(ByRef FormName As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try
            If ControlNames = "*" Then
                ControlNames = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(ControlNames).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next

            For i = 0 To Acontrol.Count - 1
                If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                    Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                    Try
                        lcontrol.enabled = EnableTrue
                    Catch ex As Exception
                        Continue For
                    End Try
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.EnableControls(ByRef FormName As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal ExceptControls As String = "")")
        End Try

    End Sub

    ''' <summary>
    ''' To make form controls enable
    ''' </summary>
    ''' <param name="FormName">Parent form , whose controls to make enable</param>
    ''' <param name="ControlNames">Comma separated string of control names to make enable,(*) for all controls of the form</param>
    ''' <param name="EnableTrue">Enable value ,Default value =True</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not considered</param>
    ''' <remarks></remarks>
    Public Sub EnableControlsTemp(ByRef FormName As Object, ByVal ControlNames As String, Optional ByVal EnableTrue As Boolean = True, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try

            Dim Acontrol() As String = LCase(ControlNames).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next

            For i = 0 To Acontrol.Count - 1
                If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then

                    Dim lcontrol As New Control

                    Dim ErpFormatRow As DataRow() = FormName.EntFormdtALL.Select("ControlName='" & Acontrol(i) & "'")
                    Dim Ctrlcollection As String = ErpFormatRow(0).Item("parentstring")
                    Dim Value1 As String() = Ctrlcollection.Split(",")
                    lcontrol = GetControlFromParentString(FormName, Value1)
                    Try
                        lcontrol.Enabled = EnableTrue
                    Catch ex As Exception
                        Continue For
                    End Try
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.EnableControls(ByRef FormName As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal EnableTrue As Boolean = True, Optional ByVal ExceptControls As String = "")")
        End Try

    End Sub





    ''' <summary>
    ''' To make previous form controls visible with respect to a controlName  
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet"> Comma separated string of control names ,(*) for all controls of the form </param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="VisibleTrue">Visible value ,Default value =True</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>
    Public Sub VisiblePreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LformControlsSet As String = FormControlsSet
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                If Acontrol(i) = LCase(ControlName) Then
                    If IncludeCurrentControl = True Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.visible = VisibleTrue
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                    Exit For
                Else
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.visible = VisibleTrue
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.VisiblePreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try
    End Sub
    ''' <summary>
    ''' To make next form controls visible with respect to a ControlName  
    ''' </summary>
    ''' <param name="FormName"> Parent form </param>
    ''' <param name="FormControlsSet"> Comma separated string of control names,(*) for all controls of the form </param>
    '''  <param name="ControlName"> Control name from  which next controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="VisibleTrue">Visible value ,Default value =True</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>

    Public Sub VisibleNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            Dim k As Integer = Array.IndexOf(Acontrol, LCase(ControlName))
            If k > -1 Then
                k = IIf(IncludeCurrentControl = True, k, k + 1)
                For i = k To Acontrol.Count - 1
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.visible = VisibleTrue
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.VisibleNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleTrue As Boolean = True, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try
    End Sub

    ''' <summary>
    ''' To make form controls visible
    ''' </summary>
    ''' <param name="LForm">Parent form , whose controls to make visible</param>
    ''' <param name="ControlNames">Comma separated string of control names to make visible,(*) for all controls of the form</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames </param>
    ''' <param name="VisibleTrue">Visible value ,Default value =True</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not considered </param>
    ''' <remarks></remarks>
    Public Sub VisibleControls(ByRef LForm As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleTrue As Boolean = True, Optional ByVal ExceptControls As String = "")
        '* for all controls
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            If ControlNames = "*" Then
                ControlNames = GetAllControlNames(LForm, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(ControlNames).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                '  MsgBox(Array.IndexOf(Econtrol, Acontrol(i)).ToString)

                If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                    Dim lcontrol As Object = ControlNameToObject(LForm, Acontrol(i))
                    Try
                        lcontrol.visible = VisibleTrue
                    Catch ex As Exception
                        Continue For
                    End Try
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.VisibleControls(ByRef LForm As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal VisibleTrue As Boolean = True, Optional ByVal ExceptControls As String = "")")
        End Try

    End Sub



    ''' <summary>
    ''' To set  form controls tab index sequentially.
    ''' </summary>
    ''' <param name="LForm">Parent form , whose controls to make visible</param>
    ''' <param name="ControlNames">Comma separated string of control names to make visible,(*) for all controls of the form</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames </param>
    ''' <remarks></remarks>
    Public Sub SetTabIndex(ByRef LForm As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, ByVal StartTabIndex As Integer)
        '* for all controls
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            If ControlNames = "*" Then
                ControlNames = GetAllControlNames(LForm, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(ControlNames).Split(",")
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Dim lcontrol As Object = ControlNameToObject(LForm, Acontrol(i))
                Try
                    lcontrol.TabIndex = StartTabIndex + i
                Catch ex As Exception
                    Continue For
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SetTabIndex(ByRef LForm As Object, ByVal ControlNames As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, ByVal StartTabIndex As Integer)")

        End Try

    End Sub
    ''' <summary>
    ''' To SendToBack() overlaping next controls in z-order.  
    ''' </summary>
    ''' <param name="FormName"> Parent form </param>
    ''' <param name="FormControlsSet"> Comma separated string of control names,(*) for all controls of the form </param>
    '''  <param name="ControlName"> Control name from  which next controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>

    Public Sub SendToBackNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            Dim k As Integer = Array.IndexOf(Acontrol, LCase(ControlName))
            If k > -1 Then
                k = IIf(IncludeCurrentControl = True, k, k + 1)
                For i = k To Acontrol.Count - 1
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.sendtoback()
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SendToBackNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try
    End Sub

    ''' <summary>
    ''' To BringToFront() overlaping next controls in z-order.  
    ''' </summary>
    ''' <param name="FormName"> Parent form </param>
    ''' <param name="FormControlsSet"> Comma separated string of control names,(*) for all controls of the form </param>
    '''  <param name="ControlName"> Control name from  which next controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>

    Public Sub BringToFrontNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            Dim k As Integer = Array.IndexOf(Acontrol, LCase(ControlName))
            If k > -1 Then
                k = IIf(IncludeCurrentControl = True, k, k + 1)
                For i = k To Acontrol.Count - 1
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.BringToFront()
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.BringToFrontNextControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try

    End Sub
    ''' <summary>
    ''' To BringToFront() overlaping controls in z-order.  
    ''' </summary>
    ''' <param name="FormName"> Parent form </param>
    ''' <param name="FormControlsSet"> Comma separated string of control names,(*) for all controls of the form </param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <remarks></remarks>

    Public Sub BringToFrontControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            '* for all controls
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                Try
                    lcontrol.BringToFront()
                Catch ex As Exception
                    Continue For
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.BringToFrontControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean)")
        End Try

    End Sub

    ''' <summary>
    ''' To SentToBack() overlaping controls in z-order.  
    ''' </summary>
    ''' <param name="FormName"> Parent form </param>
    ''' <param name="FormControlsSet"> Comma separated string of control names,(*) for all controls of the form </param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <remarks></remarks>

    Public Sub SentToBackControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                Try
                    lcontrol.SentToBack()
                Catch ex As Exception
                    Continue For
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SentToBackControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean)")
        End Try
    End Sub


    ''' <summary>
    '''   To SendToBack() overlaping previous controls in z-order.  
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet"> Comma separated string of control names ,(*) for all controls of the form </param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>
    Public Sub SentToBackPreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LformControlsSet As String = FormControlsSet
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                If Acontrol(i) = LCase(ControlName) Then
                    If IncludeCurrentControl = True Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.SendToBack()
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                    Exit For
                Else
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.SendToBack()
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SentToBackPreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try
    End Sub
    ''' <summary>
    '''   To SendToBack() overlaping previous controls in z-order.  
    ''' </summary>
    ''' <param name="FormName">Parent forms</param>
    ''' <param name="FormControlsSet"> Comma separated string of control names ,(*) for all controls of the form </param>
    ''' <param name="ControlName">Control name for which previous controls are to be  considered</param>
    '''<param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence,""=As in control array. </param>
    ''' <param name="OnlyTopLevel">False,If also executed for child controls of controlnames</param>
    ''' <param name="IncludeCurrentControl " >True if above control name included in visible controls list</param>
    ''' <param name="ExceptControls">Comma separated string of control names which are not to be considered</param>
    ''' <remarks></remarks>
    Public Sub BringToFrontPreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim LformControlsSet As String = FormControlsSet
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, ControlSequence, OnlyTopLevel)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            Dim Econtrol() As String = LCase(ExceptControls).Split(",")
            ControlName = ControlName.Trim
            For i = 0 To Econtrol.Count - 1
                Econtrol(i) = Econtrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                If Acontrol(i) = LCase(ControlName) Then
                    If IncludeCurrentControl = True Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.BringToFront()
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                    Exit For
                Else
                    If Array.IndexOf(Econtrol, Acontrol(i)) < 0 Then
                        Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                        Try
                            lcontrol.BringToFront()
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.BringToFrontPreviousControls(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal ControlName As String, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal IncludeCurrentControl As Boolean = False, Optional ByVal ExceptControls As String = "")")
        End Try


    End Sub

    ''' <summary>
    ''' Get all control names of a form as a string separated by sep0 default(",")
    ''' </summary>
    ''' <param name="lform"> Parent form</param>
    ''' <param name="ControlSequence" >Sequence of controls in the string, "T" =As per TabIndex,"L"=As per Location,"R"=As per Rectangle Sequence.</param>
    ''' <param name="OnlyTopLevel"> Included only top level controls ,Default value=False </param>
    ''' <param name="Sep0"> Control name's separator default is "," </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetAllControlNames(ByVal Lform As Object, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean, Optional ByVal Sep0 As String = ",") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim result As String = ""
        Try
            Dim mControls As New List(Of Collection)
            Dim ChildArr() As Control = {}
            For Each ctrl1 As Control In Lform.Controls
                '            result = result & IIf(result.Length = 0, "", Sep0) & ctrl1.Name
                Dim lCollection As New Collection
                lCollection.Add(ctrl1.Name, "name")
                lCollection.Add(ctrl1.Left, "left")
                lCollection.Add(ctrl1.Top, "top")
                lCollection.Add(ctrl1.TabIndex, "tabindex")
                mControls.Add(lCollection)
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
            Array.Copy(ChildArr, ExtChildArr, ChildArr.Count)
            If ChildArr.Length = 0 Then
                GoTo rloop
            End If
            Array.Clear(ChildArr, 0, 0)
            Array.Resize(ChildArr, 0)
            For i = 0 To ExtChildArr.Length - 1
                Dim ctrl1 As Control = ExtChildArr(i)
                For Each ctrl2 As Control In ExtChildArr(i).Controls '--------testing inner controls--------
                    ' result = result & IIf(result.Length = 0, "", Sep0) & ctrl2.Name
                    Dim lCollection As New Collection
                    lCollection.Add(ctrl2.Name, "name")
                    lCollection.Add(ctrl2.Left + ctrl1.Left, "left")
                    lCollection.Add(ctrl2.Top + ctrl1.Top, "top")
                    lCollection.Add(ctrl2.TabIndex, "tabindex")
                    mControls.Add(lCollection)
                    If OnlyTopLevel = False Then
                        If ctrl2.HasChildren Then
                            ArrayAppend(ChildArr, ctrl2)
                        End If
                    End If
                Next
            Next
            GoTo sloop
rloop:
            Dim mControls1() As Collection = mControls.ToArray
            If ControlSequence = "L" Then
                Dim mSortKeys() As String = {"left", "top"}
                mControls1 = SortCollection(mControls1, mSortKeys)
            End If
            If ControlSequence = "T" Then
                mControls1 = SortCollection(mControls1, "tabindex")
            End If
            result = ""
            For i = 0 To mControls1.Count - 1
                result = result & IIf(result.Length = 0, "", Sep0) & mControls1(i).Item("name").ToString
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.GetAllControlNames(ByVal Lform As Object, ByVal ControlSequence As String, ByVal OnlyTopLevel As Boolean) As String")
        End Try

        Return result
    End Function
    ''' <summary>
    ''' Get the object/control of the form  by its name (string)  
    ''' </summary>
    ''' <param name="lform"> Parent form of the object / control </param>
    ''' <param name="ControlName"> Name of control or object as string</param>
    ''' <param name="TypeOfControlName">Output as control type,F=Form,C=Control,V=Form variables,E=External variables</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ControlNameToObject(ByVal lform As Object, ByVal ControlName As String, Optional ByRef TypeOfControlName As String = "") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim Val As Object = Nothing
        If lform Is Nothing Then Return Val
        Try
            Dim ChildArr() As Control = {}
            If LCase("me") = LCase(Trim(ControlName)) Then
                TypeOfControlName = "F"
                Return lform.Name   '===========returns form====
            End If
            For Each ctrl1 As Control In lform.Controls
                If LCase(ctrl1.Name) = LCase(Trim(ControlName)) Then
                    Val = ctrl1
                    TypeOfControlName = "C"
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
            Array.Copy(ChildArr, ExtChildArr, ChildArr.Count)
            If ChildArr.Length = 0 Then
                GoTo rloop
            End If
            Array.Clear(ChildArr, 0, 0)
            Array.Resize(ChildArr, 0)
            For i = 0 To ExtChildArr.Length - 1
                For Each ctrl2 As Control In ExtChildArr(i).Controls '--------testing inner controls--------
                    If LCase(ctrl2.Name) = LCase(Trim(ControlName)) Then
                        Val = ctrl2
                        TypeOfControlName = "C"
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
                        TypeOfControlName = "V"
                        Return myFieldInfo(i)    '========== returns variable present on form ========
                    End If
                Next
            End If
            If Val Is Nothing Then
                '===========checking for a external variable i.e. variable on another class==============
                Dim pp As String = Application.ProductName & "." & ControlName
                Try
                    Dim ControlName_1 As Object = Activator.CreateInstance(Type.GetType(pp)) '----to obtain instance of that type--
                    TypeOfControlName = "E"
                    Return ControlName_1
                Catch ex As Exception
                    QuitMessage("Control Name not found on the form " & ControlName, "ControlNameToObject(ByVal lform As Object, ByVal ControlName As String, Optional ByRef TypeOfControlName As String = "") As Object  ")
                End Try
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ControlNameToObject(ByVal lform As Object, ByVal ControlName As String, Optional ByRef TypeOfControlName As String = "") As Object")
        End Try
        Return Val
    End Function
    ''' <summary>
    ''' To change control  size for active font accoarding to  Text string.  
    ''' </summary>
    ''' <param name="TextString"> Largest text string to be fitted in the control</param>
    ''' <param name="LControl"> Control-which size to be changed</param>
    ''' <remarks></remarks>

    Public Sub MakeControlSize(ByVal TextString As String, ByRef LControl As Control)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim instance As Graphics = LControl.CreateGraphics()
            Dim StringSize As New SizeF
            StringSize = instance.MeasureString(TextString, LControl.Font)
            StringSize.Width = IIf(StringSize.Width < 70, 70, StringSize.Width)
            StringSize.Height = IIf(StringSize.Height < 20, 20, StringSize.Height)
            LControl.Height = StringSize.Height
            LControl.Width = StringSize.Width
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.MakeControlSize(ByVal TextString As String, ByRef LControl As Control)")
        End Try

    End Sub

    '============ final evalution of mathematical expression or equation ============
    ''' <summary>
    ''' To evalute an numeric expression entered in text string
    ''' </summary>
    ''' <param name="NumericExpression">Text String of numeric expression </param>
    ''' <param name="ErrorInExpression" >Stores true if an error in found in numeric expression</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EvalNumericExpression(ByVal NumericExpression As String, ByRef ErrorInExpression As Boolean) As Decimal
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim MyProvider As New VBCodeProvider
            Dim cp As New CompilerParameters
            cp.GenerateExecutable = False
            cp.GenerateInMemory = True
            Dim TempModuleSource As String = "Imports System" & Environment.NewLine & _
                                             "Namespace ns " & Environment.NewLine & _
                                             "Public Class class1" & Environment.NewLine & _
                                             "Public Shared Function Evaluate()" & Environment.NewLine & _
                                             "Return " & NumericExpression & Environment.NewLine & _
                                             "End Function" & Environment.NewLine & _
                                             "End Class" & Environment.NewLine & _
                                             "End Namespace"
            Dim cr As CompilerResults = MyProvider.CompileAssemblyFromSource(cp, TempModuleSource)
            Dim methInfo As MethodInfo = cr.CompiledAssembly.GetType("ns.class1").GetMethod("Evaluate")
            Dim RetValue As Decimal = methInfo.Invoke(Nothing, Nothing)
            Return RetValue
        Catch ex As Exception
            MsgBox("Invalid " & NumericExpression)
            ErrorInExpression = True
            Return 0
        End Try
    End Function

    '============ final evalution of mathematical expression or equation ============
    ''' <summary>
    ''' To evalute an numeric expression entered in text string
    ''' </summary>
    ''' <param name="ExpressionString">Text String of numeric expression </param>
    ''' <param name="ErrorInExpression" >Stores true if an error in found in expression</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EvaluateExpression(ByVal ExpressionString As String, ByRef ErrorInExpression As Boolean) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim MyProvider As New VBCodeProvider
            Dim cp As New CompilerParameters
            cp.ReferencedAssemblies.Add("System.dll")
            cp.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            cp.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")
            cp.GenerateExecutable = False
            cp.GenerateInMemory = True
            Dim miif As Boolean = Microsoft.VisualBasic.IIf(LCase(Microsoft.VisualBasic.Left(ExpressionString, 4)) = "iif(", True, False)
            Dim TempModuleSource As String = ""
            If miif = False Then
                TempModuleSource = "Imports System" & Environment.NewLine & _
                                                 "Namespace ns " & Environment.NewLine & _
                                                 "Public Class class1" & Environment.NewLine & _
                                                 "Public Shared Function Evaluate() As Object" & Environment.NewLine & _
                                                 "Return " & ExpressionString & Environment.NewLine & _
                                                 "End Function" & Environment.NewLine & _
                                                 "End Class" & Environment.NewLine & _
                                                 "End Namespace"
            Else
                ExpressionString = ExpressionString.Replace("iif(", "Microsoft.VisualBasic.iif(")
                TempModuleSource = "Imports System" & Environment.NewLine & _
                                                         "Namespace ns " & Environment.NewLine & _
                                                         "Public Class class1" & Environment.NewLine & _
                                                         "Public Shared Function Evaluate() As Object" & Environment.NewLine & _
                                                         "Dim mObject as Object = " & ExpressionString & Environment.NewLine & _
                                                         "Return mObject " & Environment.NewLine & _
                                                         "End Function" & Environment.NewLine & _
                                                         "End Class" & Environment.NewLine & _
                                                         "End Namespace"
            End If
            Dim cr As CompilerResults = MyProvider.CompileAssemblyFromSource(cp, TempModuleSource)
            If cr.Errors.HasErrors Then
                MsgBox("Error: Line>" & cr.Errors(0).Line.ToString & ", " & cr.Errors(0).ErrorText)
                Return Nothing
                Exit Function
            End If
            Dim methInfo As MethodInfo = cr.CompiledAssembly.GetType("ns.class1").GetMethod("Evaluate")
            Dim RetValue As Object = methInfo.Invoke(Nothing, Nothing)
            Return RetValue
        Catch ex As Exception
            MsgBox("Invalid " & ExpressionString)
            ErrorInExpression = True
            Return Nothing
        End Try
    End Function
    '============ final evalution of mathematical expression or equation ============
    ''' <summary>
    ''' To evalute an numeric expression entered in text string
    ''' </summary>
    ''' <param name="ExpressionString">Text String of numeric expression </param>
    ''' <param name="ErrorInExpression" >Stores true if an error in found in expression</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EvaluateBooleanExpression(ByVal ExpressionString As String, ByRef ErrorInExpression As Boolean) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim MyProvider As New VBCodeProvider
            Dim cp As New CompilerParameters
            cp.GenerateExecutable = False
            cp.GenerateInMemory = True
            Dim TempModuleSource As String = "Imports System" & Environment.NewLine & _
                                             "Namespace ns " & Environment.NewLine & _
                                             "Public Class class1" & Environment.NewLine & _
                                             "Public Shared Function Evaluate()" & Environment.NewLine & _
                                             "Return " & ExpressionString & Environment.NewLine & _
                                             "End Function" & Environment.NewLine & _
                                             "End Class" & Environment.NewLine & _
                                             "End Namespace"
            Dim cr As CompilerResults = MyProvider.CompileAssemblyFromSource(cp, TempModuleSource)
            Dim methInfo As MethodInfo = cr.CompiledAssembly.GetType("ns.class1").GetMethod("Evaluate")
            Dim RetValue As Boolean = methInfo.Invoke(Nothing, Nothing)
            Return RetValue
        Catch ex As Exception
            MsgBox("Invalid " & ExpressionString)
            ErrorInExpression = True
            Return Nothing
        End Try
    End Function

    Public Sub SetSizeAsText(ByVal TextString As String, ByRef LControl As Control)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim instance As Graphics = LControl.CreateGraphics()
            Dim StringSize As New SizeF
            StringSize = instance.MeasureString(TextString, LControl.Font)
            StringSize.Width = IIf(StringSize.Width < 70, 70, StringSize.Width)
            StringSize.Height = IIf(StringSize.Height < 20, 20, StringSize.Height)
            LControl.Height = StringSize.Height
            LControl.Width = StringSize.Width
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SetSizeAsText(ByVal TextString As String, ByRef LControl As Control)")
        End Try
    End Sub
    ''' <summary>
    ''' This function sorts an array of collection on a column specified
    ''' </summary>
    ''' <param name="Lcollection">An array of collection to be sorted</param>
    ''' <param name="SortColumnKey">Key of collection column on which sorting done</param>
    ''' <param name="SortOrder">Order of sorting ASC or DESC,Default is ASC</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SortCollection(ByVal Lcollection As Collection(), ByVal SortColumnKey As String, Optional ByVal SortOrder As String = "ASC") As Array
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function

        Dim rcollection As Array = Lcollection.Clone
        Try
            If Lcollection.Length = 0 Then
                Return Lcollection
            End If
            Dim SortedList As New List(Of Collection)
            Dim LarrayList As New ArrayList(Lcollection)
startfor:
            For i = 0 To LarrayList.Count - 1
                Dim maxcollection As Collection = LarrayList.Item(i)
                Dim FirstValue As Object = LarrayList.Item(i).Item(SortColumnKey)
                Dim MaxInteger As Integer = i
                For j = 0 To LarrayList.Count - 1
                    If LarrayList.Item(j).Item(SortColumnKey) > FirstValue Then
                        maxcollection = LarrayList.Item(j)
                        FirstValue = LarrayList.Item(j).Item(SortColumnKey)
                        MaxInteger = j
                    End If
                Next
                SortedList.Add(maxcollection)
                LarrayList.RemoveAt(MaxInteger)
                GoTo startfor
            Next
            If UCase(SortOrder) = "ASC" Then
                For i = SortedList.Count - 1 To 0 Step -1
                    rcollection.SetValue(SortedList.Item(i), SortedList.Count - 1 - i)
                Next
            Else
                For i = 0 To SortedList.Count - 1
                    rcollection.SetValue(SortedList.Item(i), i)
                Next
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SortCollection(ByVal Lcollection As Collection(), ByVal SortColumnKey As String, Optional ByVal SortOrder As String = "") As Array")
        End Try
        Return rcollection
    End Function

    ''' <summary>
    ''' This function sorts an array of collection on a column specified
    ''' </summary>
    ''' <param name="Lcollection">An array of collection to be sorted</param>
    ''' <param name="SortColumnKeys">An Array of Keys of collection on which sorting done</param>
    ''' <param name="SortOrder">Order of sorting ASC or DESC,Default is ASC</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SortCollection(ByVal Lcollection As Collection(), ByVal SortColumnKeys() As String, Optional ByVal SortOrder As String = "ASC") As Array
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim rcollection As Array = Lcollection.Clone
        Try
            For s = 0 To SortColumnKeys.Count - 1
                Dim mSortColumnKey As String = SortColumnKeys(s)
                Try
                    If Lcollection.Length = 0 Then
                        Return Lcollection
                    End If
                    Dim SortedList As New List(Of Collection)
                    Dim LarrayList As New ArrayList(Lcollection)
startfor:
                    For i = 0 To LarrayList.Count - 1
                        Dim maxcollection As Collection = LarrayList.Item(i)
                        Dim FirstValue As Object = LarrayList.Item(i).Item(mSortColumnKey)
                        Dim MaxInteger As Integer = i
                        For j = 0 To LarrayList.Count - 1
                            If LarrayList.Item(j).Item(mSortColumnKey) > FirstValue Then
                                maxcollection = LarrayList.Item(j)
                                FirstValue = LarrayList.Item(j).Item(mSortColumnKey)
                                MaxInteger = j
                            End If
                        Next
                        SortedList.Add(maxcollection)
                        LarrayList.RemoveAt(MaxInteger)
                        GoTo startfor
                    Next
                    If UCase(SortOrder) = "ASC" Then
                        For i = SortedList.Count - 1 To 0 Step -1
                            rcollection.SetValue(SortedList.Item(i), SortedList.Count - 1 - i)
                        Next
                    Else
                        For i = 0 To SortedList.Count - 1
                            rcollection.SetValue(SortedList.Item(i), i)
                        Next
                    End If
                Catch ex As Exception
                    QuitError(ex, Err, "Unable to execute GlobalFunction1.SortCollection(ByVal Lcollection As Collection(), ByVal SortColumnKeys() As String, Optional ByVal SortOrder As String = ASC) As Array")
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SortCollection(ByVal Lcollection As Collection(), ByVal SortColumnKeys() As String, Optional ByVal SortOrder As String = ASC) As Array")
        End Try
        Return rcollection
    End Function

    ''' <summary>
    ''' Convert hashtable to two same size arrays i.e keyaray and value array
    ''' </summary>
    ''' <param name="LHashTable">Hashtable to be converted </param>
    ''' <param name="KeysArray">Array contains the keys of hashtable</param>
    ''' <param name="ValuesArray">Array contains the values of hashtable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertHashTableToArrays(ByVal LHashTable As Hashtable, ByRef KeysArray() As String, ByRef ValuesArray() As Object) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For k = 0 To LHashTable.Count - 1
                ArrayAppend(KeysArray, LHashTable.Keys(k))
                ArrayAppend(ValuesArray, LHashTable.Values(k))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ConvertHashTableToArrays(ByVal LHashTable As Hashtable, ByRef KeysArray() As String, ByRef ValuesArray() As Object) As Boolean")
        End Try
        Return True
    End Function


    ''' <summary>
    ''' Convert two same size arrays to hash table,where first array has unique values
    ''' </summary>
    ''' <param name="FirstArray">First string array to be used as key of hashtable</param>
    ''' <param name="SecondArray">Second object array to be used as value of hashtable</param>
    ''' <returns>New hash table</returns>
    ''' <remarks></remarks>

    Public Function ConvertTwoArraysToHashTable(ByVal FirstArray() As String, ByVal SecondArray() As Object) As Hashtable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If FirstArray.Count <> SecondArray.Count Then
            MsgBox("both Array size not equal " & "ConvertTwoArraysToHashTable")
        End If
        Dim mHashTable As New Hashtable
        Try
            For k = 0 To FirstArray.Count - 1
                mHashTable.Add(LCase(FirstArray(k)), SecondArray(k))
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ConvertTwoArraysToHashTable(ByVal FirstArray() As String, ByVal SecondArray() As Object) As Hashtable")
        End Try
        Return mHashTable
    End Function
    ''' <summary>
    ''' Remove item from array at specified Index and shrink
    ''' </summary>
    ''' <param name="ArrayName">An array of object to be shrink</param>
    ''' <param name="ItemIndex">Index no. to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrink(ByVal ArrayName As Object(), ByVal ItemIndex As Integer) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FinalArray As Object() = {}
        Try
            Array.Clear(ArrayName, ItemIndex, 1)
            For k = 0 To ArrayName.Length - 1
                If k <> ItemIndex Then
                    FinalArray = ArrayAppend(FinalArray, ArrayName.GetValue(k))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As Object(), ByVal ItemIndex As Integer) As Object()")
        End Try
        Return FinalArray
    End Function
    ''' <summary>
    ''' Remove item from array at specified Index and shrink
    ''' </summary>
    ''' <param name="ArrayName">An array of string to be shrink</param>
    ''' <param name="ItemIndex">Index no. to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrink(ByVal ArrayName As String(), ByVal ItemIndex As Integer) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FinalArray As String() = {}
        Try
            Array.Clear(ArrayName, ItemIndex, 1)
            For k = 0 To ArrayName.Count - 1
                If k <> ItemIndex Then
                    FinalArray = ArrayAppend(FinalArray, ArrayName.GetValue(k))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As String(), ByVal ItemIndex As String) As Object()")

        End Try
        Return FinalArray
    End Function
    ''' <summary>
    ''' Remove item from array at specified Index and shrink
    ''' </summary>
    ''' <param name="ArrayName">An array of string to be shrink</param>
    ''' <param name="ItemIndex">Array of Indexes  to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrink(ByVal ArrayName As String(), ByVal ItemIndex() As Integer) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FinalArray As String() = {}
        Try
            For k = 0 To ArrayName.Count - 1
                If Array.IndexOf(ItemIndex, k) < 0 Then
                    FinalArray = ArrayAppend(FinalArray, ArrayName.GetValue(k))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As String(), ByVal ItemIndex As String) As Object()")

        End Try
        Return FinalArray
    End Function
    ''' <summary>
    ''' Create a new array from a given array.
    ''' </summary>
    ''' <param name="ArrayName">An array of string be given</param>
    ''' <param name="ItemIndex">Array of Indexes  which will be elements of new array</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayNew(ByVal ArrayName As String(), ByVal ItemIndex() As Integer) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FinalArray As String() = {}
        Try
            For k = 0 To ArrayName.Count - 1
                If Array.IndexOf(ItemIndex, k) < 0 Then
                    FinalArray = ArrayAppend(FinalArray, ArrayName.GetValue(k))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As String(), ByVal ItemIndex As String) As Object()")

        End Try
        Return FinalArray
    End Function



    ''' <summary>
    ''' Remove item from array at specified Index and shrink
    ''' </summary>
    ''' <param name="ArrayName">An array of integer to be shrink</param>
    ''' <param name="ItemIndex">Index no. to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrink(ByVal ArrayName As Integer(), ByVal ItemIndex As Integer) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FinalArray As Integer() = {}
        Try
            Array.Clear(ArrayName, ItemIndex, 1)
            For k = 0 To ArrayName.Length - 1
                If k <> ItemIndex Then
                    FinalArray = ArrayAppend(FinalArray, ArrayName.GetValue(k))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As Integer(), ByVal ItemIndex As Integer) As Integer()")
        End Try
        Return FinalArray
    End Function
    ''' <summary>
    ''' Remove item from array at specified Index and shrink
    ''' </summary>
    ''' <param name="ArrayName">An array of decimal/numeric to be shrink</param>
    ''' <param name="ItemIndex">Index no. to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrink(ByVal ArrayName As Decimal(), ByVal ItemIndex As Integer) As Decimal()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FinalArray As Decimal() = {}
        Try
            Array.Clear(ArrayName, ItemIndex, 1)
            For k = 0 To ArrayName.Length - 1
                If k <> ItemIndex Then
                    FinalArray = ArrayAppend(FinalArray, ArrayName.GetValue(k))
                End If
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As Decimal(), ByVal ItemIndex As Decimal) As Decimal()")
        End Try
        Return FinalArray
    End Function
    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValue">Item value to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As Object(), ByVal ItemValue As Object) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValue)
            ArrayName = ArrayShrink(ArrayName, ItemIndex)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrink(ByVal ArrayName As Object(), ByVal ItemIndex As Object) As Object()")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValues">Item values array  to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As Object(), ByVal ItemValues() As Object) As Object()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To ItemValues.Count - 1
                Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValues(i))
                ArrayName = ArrayShrink(ArrayName, ItemIndex)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As Object(), ByVal ItemIndex As Object) As Object()")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValue">Item value to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As String(), ByVal ItemValue As String) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValue)
            ArrayName = ArrayShrink(ArrayName, ItemIndex)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As String(), ByVal ItemIndex As Object) As String()")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValues">Array of item values removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As String(), ByVal ItemValues() As String) As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To ItemValues.Count - 1
                Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValues(i))
                ArrayName = ArrayShrink(ArrayName, ItemIndex)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As String(), ByVal ItemValues() As Object) As String()")
        End Try
        Return ArrayName
    End Function


    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValue">Item value to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As Decimal(), ByVal ItemValue As Decimal) As Decimal()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValue)
            ArrayName = ArrayShrink(ArrayName, ItemIndex)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As Decimal(), ByVal ItemValue As Object) As String()")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValues">Item values array to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As Decimal(), ByVal ItemValues() As Decimal) As Decimal()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To ItemValues.Count - 1
                Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValues(i))
                ArrayName = ArrayShrink(ArrayName, ItemIndex)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As Decimal(), ByVal ItemValue() As Object) As String()")
        End Try
        Return ArrayName
    End Function


    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValue">Item value to be removed</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As Integer(), ByVal ItemValue As Integer) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValue)
            ArrayName = ArrayShrink(ArrayName, ItemIndex)
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As Integer(), ByVal ItemValue() As Integer) As String()")
        End Try
        Return ArrayName
    End Function
    ''' <summary>
    '''Remove item from array by specifying item value and shrink
    ''' </summary>
    ''' <param name="ArrayName">Array to be shrinked</param>
    ''' <param name="ItemValue">Array of Item Values </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ArrayShrinkByValue(ByVal ArrayName As Integer(), ByVal ItemValue() As Integer) As Integer()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            For i = 0 To ItemValue.Count - 1
                Dim ItemIndex As Integer = Array.IndexOf(ArrayName, ItemValue(i))
                ArrayName = ArrayShrink(ArrayName, ItemIndex)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ArrayShrinkByValue(ByVal ArrayName As Integer(), ByVal ItemValue() As Integer) As Integer()")
        End Try
        Return ArrayName
    End Function

    ''' <summary>
    ''' To Set a common value to a property to specified controls.  
    ''' </summary>
    ''' <param name="FormName"> Parent form </param>
    ''' <param name="FormControlsSet"> Comma separated string of control names,(*) for all controls of the form </param>
    '''<param name="PropertyName " >Property name which is set </param>
    ''' <param name="PropertyValue ">Property Value which to be set</param>
    ''' <remarks></remarks>

    Public Sub SetCommonPropertyValue(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal PropertyName As String, ByVal PropertyValue As Object)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        '* for all controls
        Try
            If FormControlsSet = "*" Then
                FormControlsSet = GetAllControlNames(FormName, "", True)
            End If
            Dim Acontrol() As String = LCase(FormControlsSet).Split(",")
            For i = 0 To Acontrol.Count - 1
                Acontrol(i) = Acontrol(i).Trim
            Next
            For i = 0 To Acontrol.Count - 1
                Dim lcontrol As Object = ControlNameToObject(FormName, Acontrol(i))
                Try

                    Dim PropArr() As PropertyInfo = lcontrol.GetType.GetProperties
                    For ii = 0 To PropArr.Count - 1
                        Dim aa As String = PropArr(ii).Name

                        If LCase(PropArr(ii).Name) = LCase(PropertyName) Then
                            '  Dim mn As Object = prop.GetValue(SFrm, Nothing)
                            Dim t As Type = lcontrol.GetType
                            Dim t2 As PropertyInfo = t.GetProperty(PropArr(ii).Name)
                            t2.SetValue(lcontrol, PropertyValue, BindingFlags.IgnoreCase, Nothing, Nothing, Nothing)

                            Exit For
                        End If
                    Next
                Catch ex As Exception
                    Continue For
                End Try
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SetCommonPropertyValue(ByRef FormName As Object, ByVal FormControlsSet As String, ByVal PropertyName As String, ByVal PropertyValue As Object)")
        End Try
    End Sub
    ''' <summary>
    ''' Convert decimal numerics in to words 
    ''' </summary>
    ''' <param name="InputNumber">Input decimal value to be converted into words</param>
    ''' <param name="CurrencyName">Currency name in words eg. "Rupees" </param>
    ''' <param name="FractionName">Name of fraction of currency eg. "Paise"</param>
    ''' <param name="CurrencyPosition">Position of currency prefix or suffix</param>
    ''' <param name="DigitsAfterDecimal ">Digits after decimal eg. 2</param>
    ''' <param name="FigureToWordsSystem">indian or british</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ConvertFiguresInToWords(ByVal InputNumber As Decimal, ByVal CurrencyName As String, ByVal FractionName As String, ByVal CurrencyPosition As String, ByVal DigitsAfterDecimal As Integer, ByVal FigureToWordsSystem As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim FullName As String = ""
        Try
            Dim TextStr As String = ""
            Dim WholeNum As String = ""
            If InputNumber.ToString.Contains(".") Then
                WholeNum = Microsoft.VisualBasic.Mid(InputNumber, 1, InputNumber.ToString.IndexOf("."))
            Else
                WholeNum = InputNumber
            End If
            Dim DecimalNum As String = ""
            Dim Decimalsize As String = IIf(DigitsAfterDecimal > 0, StrDup(DigitsAfterDecimal, "0"), "")
            If InputNumber.ToString.Contains(".") Then
                DecimalNum = Microsoft.VisualBasic.Mid(InputNumber, InputNumber.ToString.IndexOf(".") + 2, InputNumber.ToString.Length - InputNumber.ToString.IndexOf(".") + 1)
                DecimalNum = Microsoft.VisualBasic.Left(DecimalNum & "000000000", DigitsAfterDecimal)
            End If
            Dim WholeNumString As String = Microsoft.VisualBasic.Right("1000000000000000" + WholeNum, 15)
            Dim FirstTwo As String = Microsoft.VisualBasic.Right(WholeNumString, 2)
            Dim SecondOne As String = Microsoft.VisualBasic.Mid(WholeNumString, 13, 1)
            Dim ThirdtTwo As String = Microsoft.VisualBasic.Mid(WholeNumString, 11, 2)
            Dim FourthTwo As String = Microsoft.VisualBasic.Mid(WholeNumString, 9, 2)
            Dim FifthTwo As String = Microsoft.VisualBasic.Mid(WholeNumString, 7, 2)
            Dim sixthTwo As String = Microsoft.VisualBasic.Mid(WholeNumString, 5, 2)
            Dim seventhTwo As String = Microsoft.VisualBasic.Mid(WholeNumString, 3, 2)
            Dim eightTwo As String = Microsoft.VisualBasic.Mid(WholeNumString, 1, 2)

            Dim ThirdtThree As String = Microsoft.VisualBasic.Mid(WholeNumString, 10, 3)
            Dim FourthThree As String = Microsoft.VisualBasic.Mid(WholeNumString, 7, 3)
            Dim FifthThree As String = Microsoft.VisualBasic.Mid(WholeNumString, 4, 3)
            Dim sixthThree As String = Microsoft.VisualBasic.Mid(WholeNumString, 1, 3)
            If eightTwo > "00" Then
                eightTwo = IIf(eightTwo > "01", (BreakStringIntoWords(eightTwo, 2) & " Neel"), (BreakStringIntoWords(eightTwo, 2) & " Neel"))
            End If
            If seventhTwo > "00" Then
                seventhTwo = IIf(seventhTwo > "01", (BreakStringIntoWords(seventhTwo, 2) & " Kharab"), (BreakStringIntoWords(seventhTwo, 2) & " Kharab"))
            End If
            If sixthTwo > "00" Then
                sixthTwo = IIf(sixthTwo > "01", (BreakStringIntoWords(sixthTwo, 2) & " Arab"), (BreakStringIntoWords(sixthTwo, 2) & " Arab"))
            End If
            If FifthTwo > "00" Then
                FifthTwo = IIf(FifthTwo > "01", (BreakStringIntoWords(FifthTwo, 2) & " Crore"), (BreakStringIntoWords(FifthTwo, 2) & " Crore"))
            End If
            If FourthTwo > "00" Then
                FourthTwo = IIf(FourthTwo > "01", BreakStringIntoWords(FourthTwo, 2) & " Lakh", (BreakStringIntoWords(FourthTwo, 2) & " Lakh"))
            End If
            If ThirdtTwo > "00" Then
                ThirdtTwo = IIf(ThirdtTwo > "01", (BreakStringIntoWords(ThirdtTwo, 2) & " Thousand"), (BreakStringIntoWords(ThirdtTwo, 2) & " Thousand"))
            End If
            If SecondOne > "0" Then
                SecondOne = IIf(SecondOne > "01", (BreakStringIntoWords(SecondOne, 1) & " Hundred"), (BreakStringIntoWords(SecondOne, 1) & " Hundred"))
            End If
            If FirstTwo > "00" Then
                FirstTwo = BreakStringIntoWords(FirstTwo, 2) & " "
            End If
            If ThirdtThree > "000" Then
                ThirdtThree = IIf(ThirdtThree > "001", (BreakStringIntoWords(ThirdtThree, 3) & " Thousand"), (BreakStringIntoWords(ThirdtThree, 3) & " Thousand"))
            End If
            If FourthThree > "000" Then
                FourthThree = IIf(FourthThree > "001", (BreakStringIntoWords(FourthThree, 3) & " Million"), (BreakStringIntoWords(FourthThree, 3) & " Million"))
            End If
            If FifthThree > "000" Then
                FifthThree = IIf(FifthThree > "001", (BreakStringIntoWords(FifthThree, 3) & " Billion"), (BreakStringIntoWords(FifthThree, 3) & " Billion"))
            End If
            If sixthThree > "000" Then
                sixthThree = IIf(sixthThree > "001", (BreakStringIntoWords(sixthThree, 3) & " Trillion"), (BreakStringIntoWords(sixthThree, 3) & " Trillion"))
            End If
            If eightTwo = "00" Or eightTwo.Trim = "Neel" Then : eightTwo = "" : End If
            If seventhTwo = "00" Or seventhTwo.Trim = "Kharab" Then : seventhTwo = "" : End If
            If FifthTwo = "00" Or FifthTwo.Trim = "Crore" Then : FifthTwo = "" : End If
            If sixthThree = "000" Or sixthThree.Trim = "Trillion" Then : sixthThree = "" : End If
            If FifthThree = "000" Or FifthThree.Trim = "Billion" Then : FifthThree = "" : End If
            If FourthThree = "000" Or FourthThree.Trim = "Million" Then : FourthThree = "" : End If
            If ThirdtThree = "000" Or ThirdtThree.Trim = "Thousand" Then : ThirdtThree = "" : End If
            If FirstTwo = "00" Then : FirstTwo = "" : End If
            If SecondOne = "0" Or SecondOne.Trim = "Hundred" Then : SecondOne = "" : End If
            If ThirdtTwo = "00" Or ThirdtTwo.Trim = "Thousand" Then : ThirdtTwo = "" : End If
            If FourthTwo = "00" Or FourthTwo.Trim = "Lakh" Then : FourthTwo = "" : End If
            If sixthTwo = "00" Or sixthTwo.Trim = "Arab" Then : sixthTwo = "" : End If

            If LCase(FigureToWordsSystem) = "indian" Then
                TextStr = TextStr & " " & eightTwo & " " & seventhTwo & " " & sixthTwo & " " & FifthTwo & " " & FourthTwo & " " & ThirdtTwo & " " & SecondOne & " " & FirstTwo
            ElseIf LCase(FigureToWordsSystem) = "british" Then
                TextStr = TextStr & " " & sixthThree & " " & FifthThree & " " & FourthThree & " " & ThirdtThree & " " & SecondOne & " " & FirstTwo
            End If
            If TextStr.Length > 0 Then
                If LCase(CurrencyPosition) = "prefix" Then
                    TextStr = CurrencyName & " " & TextStr
                Else
                    TextStr = TextStr & " " & CurrencyName
                End If
                TextStr = TextStr.Trim
            End If
            Dim lword As String = ""
            If DecimalNum > Decimalsize Then
                If Decimalsize.Length > 2 Then
                    For j = 1 To DecimalNum.Length
                        Dim item As String = Mid(DecimalNum, j, 1)
                        lword = lword & BreakStringIntoWords(item, 1) & " "
                    Next
                Else
                    lword = BreakStringIntoWords(DecimalNum, 2) & " "
                End If
                If LCase(CurrencyPosition) = "prefix" Then
                    lword = FractionName & " " & lword
                Else
                    lword = lword & " " & FractionName
                End If
                lword = lword.Trim
            End If
            FullName = IIf(TextStr.Length > 0, TextStr, "")
            FullName = FullName & IIf(lword.Length > 0, IIf(FullName.Length > 0, " And ", "") & lword, "")
            If FullName.Length = 0 Then
                FullName = " Zero Only"
            Else
                FullName = FullName & " Only"
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ConvertFiguresInToWords(ByVal InputNumber As Decimal, ByVal CurrencyName As String, ByVal FractionName As String, ByVal CurrencyPosition As String, ByVal DigitsAfterDecimal As Integer, ByVal FigureToWordsSystem As String) As String")
        End Try
        Return FullName
    End Function

    Private Function BreakStringIntoWords(ByVal GivenNum As String, ByVal nosize As Integer) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim outstring As String = ""
        Try
            Dim NumWordArray1() As String = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "30", "40", "50", "60", "70", "80", "90"}
            Dim NumWordArray2() As String = {"", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninty"}
            Dim str As String = ""
            Select Case nosize
                Case 2
                    If GivenNum > "20" Then
                        Dim d1 As String = Microsoft.VisualBasic.Left(GivenNum, 1) & "0"
                        Dim d2 As String = Microsoft.VisualBasic.Right(GivenNum, 1)
                        outstring = " " & NumWordArray2(Array.IndexOf(NumWordArray1, d1)) & " " & NumWordArray2(Array.IndexOf(NumWordArray1, d2))
                    Else
                        outstring = " " & NumWordArray2(Array.IndexOf(NumWordArray1, CStr(CInt(GivenNum))))
                    End If
                Case 1
                    Return " " & NumWordArray2(Array.IndexOf(NumWordArray1, GivenNum))
                Case 3
                    Dim d1 As String = Microsoft.VisualBasic.Left(GivenNum, 1)
                    Dim d2 As String = Microsoft.VisualBasic.Mid(GivenNum, 2, 2)
                    If d1 > "0" Then
                        outstring = " " & NumWordArray2(Array.IndexOf(NumWordArray1, d1)) & " Hundred"
                    End If
                    If d2 > "20" Then
                        Dim d1a As String = Microsoft.VisualBasic.Left(d2, 1) & "0"
                        Dim d2a As String = Microsoft.VisualBasic.Right(d2, 1)
                        outstring = outstring & " " & NumWordArray2(Array.IndexOf(NumWordArray1, d1a)) & " " & NumWordArray2(Array.IndexOf(NumWordArray1, d2a))
                    Else
                        outstring = outstring & " " & NumWordArray2(Array.IndexOf(NumWordArray1, CStr(CInt(d2))))
                    End If
            End Select
            outstring = outstring.Trim
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.BreakStringIntoWords(ByVal GivenNum As String, ByVal nosize As Integer) As String")
        End Try

        Return outstring
    End Function
    Public Function ExcelColumnLetterToNumber(ByVal ExcelColumnLetter As String) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mnumber As Integer = 0
        Try
            For i = 1 To Len(ExcelColumnLetter)
                mnumber = mnumber * 26 + (Asc(UCase(Mid(ExcelColumnLetter, i, 1))) - 64)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ExcelColumnLetterToNumber(ByVal ExcelColumnLetter As String) As Integer")
        End Try
        Return mnumber - 1
    End Function
    Public Sub SetGlobalVariables(ByVal TxtFileFullPath As String, Optional ByVal Encripted As Boolean = False, Optional ByVal Delimeter As String = vbCrLf)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Try
            Dim aline() As String = {}
            If Delimeter = vbCrLf Then
                aline = StringReadAllLines(TxtFileFullPath, System.Text.Encoding.UTF8)
            Else
                Dim bline As String = StringRead(TxtFileFullPath)
                If Encripted = True Then
                    bline = StringDecript(bline, 20)
                End If
                aline = Split(bline, Delimeter)
            End If
            For i = 0 To aline.Length - 1
                Dim valuepair() As String = Split(aline(i), "=")
                If valuepair.Length = 2 Then
                    If valuepair(1).Trim.Length = 0 Or valuepair(1).Trim = """" Then
                        Continue For
                    End If
                    Select Case LCase(valuepair(0))
                        Case LCase("SqlVersion")
                            GlobalControl.Variables.RegistryFolder = valuepair(1).Trim
                        Case LCase("RegistryFolder")
                            GlobalControl.Variables.RegistryFolder = valuepair(1).Trim
                        Case LCase("EventLogger")
                            GlobalControl.Variables.EventLogger = valuepair(1).Trim
                        Case LCase("MDIHeight")
                            GlobalControl.Variables.MDIHeight = CInt(valuepair(1))
                        Case LCase("MDIWidth")
                            GlobalControl.Variables.MDIWidth = CInt(valuepair(1))
                        Case LCase("xBaseResolution")
                            GlobalControl.Variables.xBaseResolution = CInt(valuepair(1))
                        Case LCase("yBaseResolution")
                            GlobalControl.Variables.yBaseResolution = CInt(valuepair(1))
                        Case LCase("AuthenticationChecked")
                            GlobalControl.Variables.AuthenticationChecked = valuepair(1).Trim
                        Case LCase("0_srv_0")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "LocalSqlServer", valuepair(1).Trim)
                            AddItemToHashTable(GlobalControl.Variables.AllServers, "0_srv_0", valuepair(1).Trim)
                        Case LCase("1_srv_1")
                            AddItemToHashTable(GlobalControl.Variables.AllServers, "1_srv_1", valuepair(1).Trim)
                        Case LCase("0_mdf_0")
                            AddItemToHashTable(GlobalControl.Variables.MDFFiles, "0_mdf_0", valuepair(1).Trim)
                        Case LCase("1_mdf_1")
                            AddItemToHashTable(GlobalControl.Variables.MDFFiles, "1_mdf_1", valuepair(1).Trim)
                        Case LCase("2_mdf_2")
                            AddItemToHashTable(GlobalControl.Variables.MDFFiles, "2_mdf_2", valuepair(1).Trim)
                        Case LCase("3_mdf_3")
                            AddItemToHashTable(GlobalControl.Variables.MDFFiles, "3_mdf_3", valuepair(1).Trim)
                        Case LCase("4_mdf_4")
                            AddItemToHashTable(GlobalControl.Variables.MDFFiles, "4_mdf_4", valuepair(1).Trim)
                        Case LCase("5_mdf_5")
                            AddItemToHashTable(GlobalControl.Variables.MDFFiles, "5_mdf_5", valuepair(1).Trim)
                        Case LCase("TablesExcelControl")
                            GlobalControl.Variables.TablesExcelControl = valuepair(1).Trim
                        Case LCase("FieldsExcelControl")
                            GlobalControl.Variables.FieldsExcelControl = valuepair(1).Trim
                        Case LCase("Client")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "Client", valuepair(1).Trim)
                        Case LCase("CurrDt")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "CurrDt", valuepair(1).Trim)
                        Case LCase("CPUId")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "CPUId", valuepair(1).Trim)
                        Case LCase("BaseId")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "BaseId", valuepair(1).Trim)
                        Case LCase("BiosId")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "BiosId", valuepair(1).Trim)
                        Case LCase("LastDt")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "LastDt", valuepair(1).Trim)
                        Case LCase("NoDays")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "NoDays", CInt(valuepair(1)))
                        Case LCase("SaralKeyExists")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "SaralKeyExists", CBool(valuepair(1)))
                        Case LCase("SaralProduct")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "SaralProduct", valuepair(1).Trim)
                        Case LCase("SaralVersion")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "SaralVersion", valuepair(1).Trim)
                        Case LCase("NoOfVouch")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "NoOfVouch", CInt(valuepair(1)))
                        Case LCase("TypePC")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "TypePC", valuepair(1).Trim)
                        Case LCase("SaralType")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "SaralType", valuepair(1).Trim)
                            GlobalControl.Variables.SaralType = valuepair(1).Trim
                        Case LCase("AddlPCNo")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "AddlPCNo", CInt(valuepair(1)))
                        Case LCase("HomePCNo")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "HomePCNo", CInt(valuepair(1)))
                        Case LCase("RemoteNo")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteNo", CInt(valuepair(1)))
                        Case LCase("NodeNo")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "NodeNo", CInt(valuepair(1)))
                        Case LCase("MainClientCode")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "MainClientCode", valuepair(1).Trim)
                        Case LCase("AllowBusinessType")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "AllowBusinessType", valuepair(1).Trim)
                        Case LCase("ServicePhone")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "ServicePhone", valuepair(1).Trim)
                        Case LCase("AppFolder")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "AppFolder", valuepair(1).Trim)
                        Case LCase("DataFolder")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "dataFolder", valuepair(1).Trim)
                        Case LCase("ImageFolder")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "ImageFolder", valuepair(1).Trim)
                        Case LCase("ResourcesFile")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "ResourcesFile", valuepair(1).Trim)
                        Case LCase("ComputerType")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "ComputerType", valuepair(1).Trim)
                        Case LCase("LANSqlServer")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlServer", valuepair(1).Trim)
                        Case LCase("CloudSqlServer")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlServer", valuepair(1).Trim)
                        Case LCase("LocalSqlServer")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "LocalSqlServer", valuepair(1).Trim)
                            AddItemToHashTable(GlobalControl.Variables.AllServers, "0_srv_0", valuepair(1).Trim)
                        Case LCase("WebSqlServer")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlServer", valuepair(1).Trim)
                            AddItemToHashTable(GlobalControl.Variables.AllServers, "0_srv_0", valuepair(1).Trim)
                        Case LCase("SqlUserName")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserName", valuepair(1).Trim)
                        Case LCase("SqlUserPassword")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserPassword", valuepair(1).Trim)
                        Case LCase("WebSqlUserName")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserName", valuepair(1).Trim)
                        Case LCase("WebSqlUserPassword")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserPassword", valuepair(1).Trim)
                        Case LCase("LANSqlUserName")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserName", valuepair(1).Trim)
                        Case LCase("LANSqlUserPassword")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserPassword", valuepair(1).Trim)
                        Case LCase("CloudSqlUserName")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserName", valuepair(1).Trim)
                        Case LCase("CloudSqlUserPassword")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserPassword", valuepair(1).Trim)
                        Case LCase("DemoType")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "DemoType", valuepair(1).Trim)
                        Case LCase("ImageFolder")
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "ImageFolder", valuepair(1).Trim)
                        Case LCase("EmailId")
                            GlobalControl.Variables.EmailId = valuepair(1).Trim
                        Case LCase("EmailPassword")
                            GlobalControl.Variables.EmailPassword = valuepair(1).Trim
                        Case LCase("LocalSMTPServerHost")
                            GlobalControl.Variables.LocalSMTPServerHost = valuepair(1).Trim
                        Case LCase("LocalSMTPServerPort")
                            GlobalControl.Variables.LocalSMTPServerPort = CInt(valuepair(1).Trim)
                        Case LCase("LocalMTPServerEnableSsl")
                            GlobalControl.Variables.LocalMTPServerEnableSsl = CBool(valuepair(1).Trim)
                        Case LCase("WebSMTPServerPort")
                            GlobalControl.Variables.WebSMTPServerPort = CInt(valuepair(1).Trim)
                        Case LCase("WebSMTPServerEnableSsl")
                            GlobalControl.Variables.WebSMTPServerEnableSsl = CBool(valuepair(1).Trim)
                        Case LCase("WebEmailId")
                            GlobalControl.Variables.WebEmailId = valuepair(1).Trim
                        Case LCase("WebEmailPwd")
                            GlobalControl.Variables.WebEmailPwd = valuepair(1).Trim
                        Case LCase("WebSMTPServerHost")
                            GlobalControl.Variables.WebSMTPServerHost = valuepair(1).Trim
                        Case LCase("WebHostingUserName")
                            GlobalControl.Variables.WebHostingUserName = valuepair(1).Trim
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "WebHostingUserName", valuepair(1).Trim)
                        Case LCase("WebHostingUserPassword")
                            GlobalControl.Variables.WebHostingUserPassword = valuepair(1).Trim
                            AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "WebHostingUserPassword", valuepair(1).Trim)
                        Case LCase("WebHostingServer")
                            GlobalControl.Variables.WebHostingServer = valuepair(1).Trim
                        Case LCase("userid0_mdf_0")
                            GlobalControl.Variables.userid0_mdf_0 = valuepair(1).Trim
                        Case LCase("pwd0_mdf_0")
                            GlobalControl.Variables.pwd0_mdf_0 = valuepair(1).Trim
                        Case LCase("userid1_mdf_1")
                            GlobalControl.Variables.userid1_mdf_1 = valuepair(1).Trim
                        Case LCase("pwd1_mdf_1")
                            GlobalControl.Variables.pwd1_mdf_1 = valuepair(1).Trim
                        Case LCase("userid2_mdf_2")
                            GlobalControl.Variables.userid2_mdf_2 = valuepair(1).Trim
                        Case LCase("pwd2_mdf_2")
                            GlobalControl.Variables.pwd2_mdf_2 = valuepair(1).Trim
                        Case LCase("userid3_mdf_3")
                            GlobalControl.Variables.userid3_mdf_3 = valuepair(1).Trim
                        Case LCase("pwd3_mdf_3")
                            GlobalControl.Variables.pwd3_mdf_3 = valuepair(1).Trim
                        Case LCase("userid_0_srv_0")
                            GlobalControl.Variables.userid_0_srv_0 = valuepair(1).Trim

                        Case LCase("pwd_0_srv_0")
                            GlobalControl.Variables.pwd_0_srv_0 = valuepair(1).Trim

                        Case LCase("LocalHostNo")
                            GlobalControl.Variables.LocalHostNo = valuepair(1).Trim
                        Case LCase("DataFolderServerPhysicalPath")
                            GlobalControl.Variables.DataFolderServerPhysicalPath = valuepair(1).Trim
                        Case LCase("MainServerDatabase")
                            GlobalControl.Variables.MainServerDatabase = valuepair(1).Trim
                        Case LCase("TemplateServerDatabase")
                            GlobalControl.Variables.TemplateServerDatabase = valuepair(1).Trim
                        Case LCase("UserServerDatabase")
                            GlobalControl.Variables.UserServerDatabase = valuepair(1).Trim



                    End Select
                End If
            Next
            Dim msaraltype As String = LCase(GlobalControl.Variables.SaralType)
            Select Case msaraltype
                Case "lan"
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlServer").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserName").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserPassword", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlUserPassword").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "1_srv_1", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LANSqlServer").ToString.Trim)
                Case "cloud"
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlServer").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserName").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserPassword", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlUserPassword").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "1_srv_1", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "CloudSqlServer").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "0_srv_0", GlobalControl.Variables.WebHostingServer.Trim)
                Case "weblocal"
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlServer").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserName").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserPassword", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlUserPassword").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "1_srv_1", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "WebSqlServer").ToString.Trim)
                Case "webgodaddy", "webazure"
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerName", GlobalControl.Variables.WebHostingServer.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName", GlobalControl.Variables.WebHostingUserName.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserPassword", GlobalControl.Variables.WebHostingUserPassword.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "1_srv_1", GlobalControl.Variables.WebHostingServer.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "0_srv_0", GlobalControl.Variables.WebHostingServer.Trim)
                Case Else
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LocalSqlServer").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserName", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserName").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AppControlHashTable, "RemoteServerUserPassword", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SqlUserPassword").ToString.Trim)
                    AddItemToHashTable(GlobalControl.Variables.AllServers, "1_srv_1", GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "LocalSqlServer").ToString.Trim)
            End Select
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.SetGlobalVariables(ByVal TxtFileFullPath As String, Optional ByVal Encripted As Boolean = False, Optional ByVal Delimeter As String = vbCrLf)")
        End Try


    End Sub



    ''' <summary>
    ''' Returns time difference between two dates in days, hours, minute format
    ''' </summary>
    ''' <param name="date1"></param>
    ''' <param name="date2"></param>
    ''' <returns></returns>
    Public Function GetTimeDifferenceString(ByVal date1 As DateTime, ByVal date2 As DateTime) As String
        Dim a As Integer = (date1 - date2).Days
        Dim b As Integer = (date1 - date2).Hours
        Dim c As Integer = (date1 - date2).Minutes
        Dim NoDayComplaintRegister As String = a & " days " & b & " hrs " & c & " mins"
        Return NoDayComplaintRegister
    End Function

    Public Sub setGlobalDataTablesTemp(ByVal excelpath As String)
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
        Dim df1 As New DataFunctions.DataFunctions
        '  GlobalControl.Variables.
        GlobalControl.Variables.ENTformstru = df1.GetDataFromExcel(excelpath & "\EntFormsStru.xlsx")
        GlobalControl.Variables.MasterOptions = df1.GetDataFromExcel(excelpath & "\masteroptions.xlsx")

        Dim primaryKey1(0) As DataColumn
        primaryKey1(0) = GlobalControl.Variables.MasterOptions.Columns("MasterOptions_Key")
        GlobalControl.Variables.MasterOptions.PrimaryKey = primaryKey1

        GlobalControl.Variables.ERPControlsList = df1.GetDataFromExcel(excelpath & "\erpcontrolslist.xlsx")
        Dim primaryKey(1) As DataColumn
        primaryKey(0) = GlobalControl.Variables.ERPControlsList.Columns("ControlType")
        primaryKey(1) = GlobalControl.Variables.ERPControlsList.Columns("ControlTypeName")
        GlobalControl.Variables.ERPControlsList.PrimaryKey = primaryKey
        GlobalControl.Variables.ERPControlPropertiesList = df1.GetDataFromExcel(excelpath & "\erpcontrolpropertieslist.xlsx")
        GlobalControl.Variables.EntControlProperties = df1.GetDataFromExcel(excelpath & "\EntControlProperties.xlsx")
        GlobalControl.Variables.FormsProjectFiles = df1.GetDataFromExcel(excelpath & "\formsprojectfiles.xlsx")
        GlobalControl.Variables.ProcNAmeFile = df1.GetDataFromExcel(excelpath & "\procnamefile.xlsx")
        GlobalControl.Variables.GridCodeMain = df1.GetDataFromExcel(excelpath & "\gridcodemain.xlsx")
        GlobalControl.Variables.GridColumns = df1.GetDataFromExcel(excelpath & "\gridcolumns.xlsx")

    End Sub
    ''' <summary>
    ''' Return the previous or next control in a form using entrycontrol property
    ''' </summary>
    ''' <param name="MForm">Parent form on which controls are placed</param>
    ''' <param name="sender">Reference control for which next and previous control to be determined</param>
    ''' <param name="NavDirection">Navigation Direction. It can be either prev or fwd</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNextPrevcontrol(ByVal MForm As Object, ByVal sender As Object, ByVal NavDirection As String) As Control
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim ControlName As String = sender.name
        Dim Ctrl As New Control
        Select Case NavDirection
            Case "fwd"
                Dim NextControl As String = ""
                Dim i As Integer = FindIndexListOfString(MForm.EntryControlList, ControlName)
                If Not i = MForm.EntryControlList.Count - 1 Then NextControl = MForm.EntryControlList(i + 1)
                Dim ErpFormatRow As DataRow() = MForm.EntFormdtALL.Select("ControlName='" & NextControl & "'")
                If ErpFormatRow.Count > 0 Then
                    Dim Ctrlcollection As String = ErpFormatRow(0).Item("parentstring")
                    Dim Value1 As String() = Ctrlcollection.Split(",")
                    Ctrl = GetControlFromParentString(MForm, Value1)
                    If GetParentTabPage(MForm, sender.name).Name = GetParentTabPage(MForm, NextControl).Name Then
                        While Ctrl.Visible = False Or Ctrl.Enabled = False
                            Ctrl = GetNextPrevcontrol(MForm, Ctrl, "fwd")
                        End While
                    Else
                        Dim TBC As Object = MForm.Controls(Value1(1))
                        TBC.SelectedTab = GetParentTabPage(MForm, Ctrl.Name)
                        While Ctrl.Visible = False Or Ctrl.Enabled = False
                            Ctrl = GetNextPrevcontrol(MForm, Ctrl, "fwd")
                        End While
                    End If
                End If
            Case "prev"
                Dim PreControl As String = ""
                Dim i As Integer = FindIndexListOfString(MForm.EntryControlList, ControlName)
                If Not i = 0 Then PreControl = MForm.EntryControlList(i - 1)
                Dim ErpFormatRow As DataRow() = MForm.EntFormdtALL.Select("ControlName='" & PreControl & "'")
                If ErpFormatRow.Count > 0 Then
                    Dim Ctrlcollection As String = ErpFormatRow(0).Item("parentstring")
                    Dim Value1 As String() = Ctrlcollection.Split(",")
                    Ctrl = GetControlFromParentString(MForm, Value1)
                    If GetParentTabPage(MForm, sender.name).Name = GetParentTabPage(MForm, PreControl).Name Then
                        While Ctrl.Visible = False Or Ctrl.Enabled = False
                            Ctrl = GetNextPrevcontrol(MForm, Ctrl, "prev")
                        End While
                    Else
                        Dim TBC As Object = MForm.Controls(Value1(1))
                        TBC.SelectedTab = GetParentTabPage(MForm, Ctrl.Name)
                        While Ctrl.Visible = False Or Ctrl.Enabled = False
                            Ctrl = GetNextPrevcontrol(MForm, Ctrl, "prev")
                        End While
                    End If
                End If
        End Select
        Return Ctrl
    End Function
    Public Function GetParentTabPage(ByVal mform As Object, ByVal CtrlName As String) As Control
        Dim Tabpage As New Control
        Dim parentstring As String = ""
        Dim datarow As DataRow() = mform.entformdtALL.Select("ControlName='" & CtrlName & "'")
        If datarow.Length > 0 Then
            parentstring = datarow(0).Item("parentstring")
            If Not parentstring = "" Then
                Dim value As String() = parentstring.Split(",")
                If value.Length > 3 Then
                    Dim tabname As String = value(1)
                    Dim tabpagename As String = value(2)
                    Dim tab As Control = mform.Controls(tabname)
                    Tabpage = tab.Controls(tabpagename)
                End If
            End If
        End If
        Return Tabpage
    End Function


    Dim Ctrl1 As New Control
    ''' <summary>
    ''' Returns an object of type control if parent controls are given in a comma
    ''' </summary>
    ''' <param name="ParentControl">top level control e.g. form</param>
    ''' <param name="CtrlString"> String array of parent controls </param>
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetControlFromParentString(ByVal ParentControl As Control, ByVal CtrlString As String(), Optional ByVal index As Integer = 1) As Control
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If CtrlString.Length > index Then
            Ctrl1 = ParentControl.Controls(CtrlString(index))
            Ctrl1 = GetControlFromParentString(Ctrl1, CtrlString, index + 1)
        End If
        Return Ctrl1
    End Function

    ''' <summary>
    ''' Finds index of string in a list of string
    ''' </summary>
    ''' <param name="listStr">List of string in which element has to be searched</param>
    ''' <param name="StringVal">Element to be searched</param>
    ''' <param name="exactMatch">boolean option to ignore case. True for exact search</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FindIndexListOfString(ByVal listStr As List(Of String), ByVal StringVal As String, Optional ByVal exactMatch As Boolean = True) As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim l As Integer = -1
        For k = 0 To listStr.Count - 1
            If StringVal = listStr(k) And exactMatch Then
                l = k
                Return l
                Exit For
            ElseIf exactMatch = False Then
                If LCase(StringVal) = LCase(listStr(k)) Then
                    l = k
                    Return l
                    Exit For
                End If
            End If
        Next
        Return l
    End Function


    ''' <summary>
    '''  To get the absolute location of the given control.
    ''' </summary>
    ''' <param name="Mform">Form as object</param>
    ''' <param name="MainControl">Control whose location has to be determined</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetAbsoluteLocation(ByVal Mform As Object, ByVal MainControl As Control) As Point
        'Declare all the varaible
        Dim NewPoint As New Point
        Dim LeftVal As Integer = 0
        Dim TopVal As Integer = 0
        Dim AllCtrllist As List(Of Control)
        Dim EntFormdtALL As New DataTable
        'to get the require row corresponding to the control name.

        Dim ErpFormatRow As DataRow() = Mform.EntFormdtALL.Select("ControlName='" & MainControl.Name & "'")
        Dim Ctrlcollection As String = ErpFormatRow(0).Item("parentstring")

        'to get the parent string
        '   Dim Ctrlcollection As String = EntFormdtALL(0).Item("parentstring")
        'execute if parent string  is not null
        If IsDBNull(Ctrlcollection) = False Then
            'split he values
            Dim Value1 As String() = Ctrlcollection.Split(",")
            'to get a list of control from parent string values.
            AllCtrllist = GetControlListFromParentString(Mform, Value1)
            'to get a list in order as -main control,outer control and so on.
            AllCtrllist.Reverse()
            'to add form in the list
            AllCtrllist.Add(Mform)
            'to loop through all its control and access the location
            For i = 0 To AllCtrllist.Count - 1
                Dim Point1 As New Point
                If Not AllCtrllist(i) Is Nothing Then
                    Point1 = AllCtrllist(i).Location
                    'to add the left anad top val in the varaible
                    LeftVal += Point1.X
                    TopVal += Point1.Y
                End If
            Next
        End If
        'assign the values in the point varaible.
        NewPoint = New Point(LeftVal, TopVal)
        For p = 0 To Ctrllist.Count - 1
            Ctrllist.RemoveAt(0)
        Next
        Return NewPoint
    End Function
    Dim Ctrllist As New List(Of Control)
    ''' <summary>
    ''' To get a list of control from the Ctrlstring.
    ''' </summary>
    ''' <param name="ParentControl"></param>The main control-Generally Me.
    ''' <param name="CtrlString"></param>The parent string comma separatedin form of string array.
    ''' <param name="index"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetControlListFromParentString(ByVal ParentControl As Control, ByVal CtrlString As String(), Optional ByVal index As Integer = 1) As List(Of Control)
        Dim Ctrl1 As New Control
        'If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If CtrlString.Length > index Then
            'to get the control for its outer control.
            Ctrl1 = ParentControl.Controls(CtrlString(index))
            'add control in the list
            Ctrllist.Add(Ctrl1)
            'again called the function
            Ctrllist = GetControlListFromParentString(Ctrl1, CtrlString, index + 1)
        End If
        Return Ctrllist
    End Function
    ''' <summary>
    ''' Add hardcoded offsets to location so that grid opens overlaaping with correspondig control
    ''' </summary>
    ''' <param name="Loc">Absolute Location of parent control</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function setGridlocationCorrection(ByVal Loc As Point) As Point
        Dim k As New Point
        k.X = Loc.X + 8
        k.Y = Loc.Y + 31
        Return k
    End Function
    ''' <summary>
    ''' Get the computer name on which application is running
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetComputerName() As String
        Dim ComputerName As String
        ComputerName = System.Net.Dns.GetHostName
        Return ComputerName
    End Function
    'Public Sub SetGlobalDataTables()
    '    If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Exit Sub
    '    If UCase(Microsoft.VisualBasic.Left(GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "SaralType"), 3)) <> "WEB" Then
    '        Dim df1 As New DataFunctions.DataFunctions
    '        Dim mdf0 As String = GetValueFromHashTable(GlobalControl.Variables.MDFFiles, "0_mdf_0")
    '        If mdf0.Trim.Length > 0 Then
    '            Dim str0 As String = "SELECT avg_fragmentation_in_percent FROM sys.dm_db_index_physical_stats(DB_ID('" & mdf0 & "'),NULL, NULL, NULL , 'DETAILED') "
    '            Dim dt1 As DataTable = df1.SqlExecuteDataTable("1_srv_1.0_mdf_0", str0)
    '        End If
    '        Dim mdf1 As String = GetValueFromHashTable(GlobalControl.Variables.MDFFiles, "1_mdf_1")
    '        If mdf1.Trim.Length > 0 Then
    '            Dim str0 As String = "SELECT avg_fragmentation_in_percent FROM sys.dm_db_index_physical_stats(DB_ID('" & mdf1 & "'),NULL, NULL, NULL , 'DETAILED') "
    '            Dim dt1 As DataTable = df1.SqlExecuteDataTable("1_srv_1.1_mdf_1", str0)
    '        End If
    '        Dim lwhere As String = "ltrim(str(busitype,3))+','  in " & GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "AllowBusinessType") & ",  or busitype = 0 "
    '        Dim clsFormHeader As New FormHeader.FormHeader
    '        clsFormHeader.WhereClauseDefault = lwhere
    '        Dim ClsColorFontScheme As New ColorFontScheme.ColorFontScheme
    '        Dim ClsErpControlsList As New ErpControlsList.ErpControlsList
    '        Dim ClsMasterOptions As New MasterOptions.MasterOptions
    '        '  Dim ClsMenusList As New MenusList.MenusList
    '        ' Dim ClsFormControls As New FormControls.FormControls
    '        'Dim ClsFormControlsOl As New FormControlsOL.FormControlsOL
    '        'Dim aTables() As Object = {clsFormHeader, ClsColorFontScheme, ClsErpControlsList, ClsMasterOptions, ClsMenusList, ClsFormControls, ClsFormControlsOl}
    '        Dim aTables() As Object = {clsFormHeader, ClsColorFontScheme, ClsErpControlsList, ClsMasterOptions}
    '        Dim dtset As DataSet = df1.SqlExecuteDataSet(aTables)
    '        GlobalControl.Variables.FormsHeader = dtset.Tables(0)
    '        GlobalControl.Variables.ColorFontScheme = dtset.Tables(1)
    '        GlobalControl.Variables.ERPControlsList = dtset.Tables(2)
    '        GlobalControl.Variables.MasterOptions = dtset.Tables(3)
    '        '  GlobalControl.Variables.MenusList = dtset.Tables(4)
    '        ' GlobalControl.Variables.Form = dtset.Tables(0)
    '    End If

    'End Sub


    Private Function StringDecript(ByVal InputStr As String, ByVal SeedNo As Integer) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
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
            QuitError(ex, Err, "Unable to execute GlobalFunction1.StringDecript(ByVal InputStr As String, ByVal SeedNo As Integer) As String")
        End Try
        Return Ostr
    End Function
    ''' <summary>
    ''' Replace values into expression containing the elements such as @Var1,@Var2,@var3 etc. within the expression from a hashtable values with keys var1,var2,var3..
    ''' </summary>
    ''' <param name="ExpressionString">String Expression containing the elements such as @Var1,@Var2,@var3 etc.</param>
    ''' <param name="Variables">A hashtable object with keys var1,var2,var3 etc. and its values only numerics,strings and datetime are acceptable.</param>
    ''' <param name="StringFormat">Optional ,Permissible values SQL,VB,None, SQL=String values are enclosed in single quotes ,VB=String values enclosed in double quotes,None=String values without quotes.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReplaceValuesInExpression(ByVal ExpressionString As String, ByVal Variables As Hashtable, Optional ByVal StringFormat As String = "SQL") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If ExpressionString.Contains("@") = False Then
            Return ExpressionString
        End If
        Try

            If StringFormat = "VB" Then
                Dim findstr As String = Chr(34) & Chr(34)
                Dim rplstr As String = Chr(34)
                ExpressionString = Replace(ExpressionString, findstr, rplstr)
            End If
            For i = 0 To Variables.Count - 1
                Dim mkey As String = Variables.Keys(i)
                Dim mvalue As Object = GetValueFromHashTable(Variables, mkey)
                Dim mtype As String = LCase(mvalue.GetType.Name.ToString)
                Dim mvalue1 As String = mvalue.ToString
                Dim mvalue2 As String = mvalue1
                If mtype = "string" Then
                    Select Case UCase(StringFormat)
                        Case "SQL"
                            mvalue2 = "'" & mvalue1 & "'"
                        Case "VB"
                            mvalue2 = """" & mvalue1 & """"
                    End Select
                End If
                If mtype = "datetime" Then
                    Select Case UCase(StringFormat)
                        Case "SQL"
                            mvalue2 = "'" & mvalue1 & "'"
                        Case "VB"
                            mvalue2 = "#" & mvalue1 & "#"
                    End Select
                End If
                mkey = "@" & mkey
                ExpressionString = Replace(ExpressionString, mkey, mvalue2, 1, -1, CompareMethod.Text)
            Next
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute GlobalFunction1.ReplaceValuesInExpression(ByVal ExpressionString As String, ByVal Variables As Hashtable, Optional ByVal SqlStringFormat As Boolean = True) As String")
        End Try
        Return ExpressionString
    End Function
    ''' <summary>
    ''' Extract variables as string array from an expression . 
    ''' </summary>
    ''' <param name="Expression">Expression as string having constants,variables,operators</param>
    ''' <param name="NumericExpression" >True if it is a numeric expression.</param>
    ''' <param name="StartLetter">Start letter to prefix on variable to define , default is @</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExtractVariables(ByVal Expression As String, Optional ByVal NumericExpression As Boolean = True, Optional ByVal StartLetter As String = "@") As String()
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Expression = Expression.Replace(" ", "") 'removing WhiteSpaces
        Dim varval() As String = {}
        Dim aparam() As Char = {"+"c, "^"c, "="c, "*"c, "-"c, "/", "<"c, ">"c, "", "("c, ")"c, "}"c, "{"c, "]"c, "["c, "."c, "'"c, """"c, ","c, "!"c, "~"c, "#"c, "$"c, "^"c, "&"c, ":"c, ";"c, " "c}
        'If NumericExpression = True Then
        '   aparam= 
        'Else
        '   aparam = {"+"c, "^"c, "="c, "*"c, "-"c, "/", "<"c, ">"c, "", "("c, ")"c, "}"c, "{"c, "]"c, "["c, "."c, "'"c, """"c, ","c, "_"c, "!"c, "~"c, "#"c, "$"c, "^"c, "&"c, ":"c, ";"c, " "c}
        'End If
        Dim SplitArr() As String = Expression.Split(aparam) 'Split at Delimiters
        For i = 0 To SplitArr.Length - 1
            If LCase(Left(SplitArr(i).Trim, 1)) = LCase(StartLetter) Then
                Dim mvar As String = SplitArr(i).Trim
                If StartLetter = "@" Then
                    mvar = mvar.Replace("@", "")
                End If
                varval = ArrayAppend(varval, mvar)
            End If
        Next
        Return varval
    End Function



    ''' <summary>
    ''' To add a row to the given datatable with the values extracted from hashtable 
    ''' </summary>
    ''' <param name="dt"> Datatable in with datarow has to be appended</param>
    ''' <param name="mhash">Hashtable containing key value pairs of columns of datatable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateRowFromHashtable(ByVal dt As DataTable, ByVal mhash As Hashtable) As DataTable
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim a As DataRow = dt.NewRow
        Dim Rowapp As Boolean = False
        For i = 0 To mhash.Count - 1
            Dim Key1 As String = mhash.Keys(i)
            Dim mvalue As New Object
            mvalue = GetValueFromHashTable(mhash, Key1)
            If Not dt.Columns.Contains(Key1) Then
                Continue For
            Else
                If Not mvalue Is Nothing Then
                    Rowapp = True
                    a(Key1) = mvalue
                End If
            End If
        Next
        If Rowapp Then dt.Rows.Add(a)
        Return dt
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
        If IsDBNull(LDataRow) Then
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
    ''' Count pc 's on network
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CountNetworkPc() As Integer
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim childEntry As DirectoryEntry
        Dim ParentEntry As New DirectoryEntry()
        Dim str As String = ""
        Dim CountPC As Integer = 0
        Try
            ParentEntry.Path = "WinNT:"
            For Each childEntry In ParentEntry.Children
                Select Case childEntry.SchemaClassName
                    Case "Domain"
                        Dim SubChildEntry As DirectoryEntry
                        Dim SubParentEntry As New DirectoryEntry()
                        SubParentEntry.Path = "WinNT://" & childEntry.Name
                        For Each SubChildEntry In SubParentEntry.Children
                            Select Case SubChildEntry.SchemaClassName
                                Case "Computer"
                                    str += SubChildEntry.Name + " - " & GetIPAddress(SubChildEntry.Name) & vbNewLine
                                    CountPC += 1
                            End Select
                        Next
                End Select
            Next
            Return CountPC
        Catch Ex As Exception
            MsgBox("Error While Reading Directories")
        Finally
            ParentEntry = Nothing
        End Try
    End Function
    ''' <summary>
    ''' Get IP address of a computer name
    ''' </summary>
    ''' <param name="CompName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetIPAddress(ByVal CompName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim sAddr As String = ""
        Try
            Dim hostentry As New IPHostEntry
            hostentry = System.Net.Dns.GetHostEntry(CompName)
            sAddr = hostentry.AddressList(0).ToString
            Return sAddr
        Catch Ex As Exception
            MsgBox(Ex.Message)
        End Try
        Return sAddr
    End Function
    ''' <summary>
    ''' Check network connection on a computer
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckNetworkConnection() As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If My.Computer.Network.IsAvailable = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Check a pcname, wether it is on the network or not.
    ''' </summary>
    ''' <param name="PCName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckNetworkPC(ByVal PCName As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            If My.Computer.Network.Ping(PCName) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Download file from network.
    ''' </summary>
    ''' <param name="SourceFile"></param>
    ''' <param name="DestinationFile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DownloadFileFromNetwork(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            My.Computer.Network.DownloadFile(SourceFile, DestinationFile)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Upload file on network.
    ''' </summary>
    ''' <param name="SourceFile"></param>
    ''' <param name="DestinationFile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UploadFileOnNetwork(ByVal SourceFile As String, ByVal DestinationFile As String) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            My.Computer.Network.UploadFile(SourceFile, DestinationFile)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function FormattedValue(ByVal mValue As DateTime, ByVal mFormatString As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aFormat() As String = mFormatString.Split(">>")
        Dim mValue1 As DateTime = CDate(mValue)
        Dim sValue As String = mValue1.ToString
        If aFormat.Count > 1 Then
            Dim sFormat As String = aFormat(1)
            sValue = mValue1.ToString(sFormat)
        End If
        Return sValue
    End Function

    Public Function FormattedValue(ByVal mValue As Decimal, ByVal mFormatString As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aFormat() As String = mFormatString.Split(">>")
        Dim mValue1 As Decimal = CDec(mValue)
        Dim sValue As String = mValue1.ToString
        If aFormat.Count > 1 Then
            Dim sFormat As String = aFormat(1)
            sValue = mValue1.ToString(sFormat)
        End If
        Return sValue
    End Function
    Public Function FormattedValue(ByVal mValue As Integer, ByVal mFormatString As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim aFormat() As String = mFormatString.Split(">>")
        Dim mValue1 As Integer = CInt(mValue)
        Dim sValue As String = mValue1.ToString
        If aFormat.Count > 1 Then
            Dim sFormat As String = aFormat(1)
            sValue = mValue1.ToString(sFormat)
        End If
        Return sValue
    End Function
    ''' <summary>
    ''' To Convert a value in a formatted string or get value from DtMasterOptions or evaluate an expression
    ''' </summary>
    ''' <param name="mValue">Value to </param>
    ''' <param name="mFormatString"></param>
    ''' <param name="DtMasterOptions"></param>
    ''' <param name="PublicVariables"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FormattedValue(ByVal mValue As Object, ByVal mFormatString As String, Optional ByVal DtMasterOptions As DataTable = Nothing, Optional ByVal PublicVariables As Hashtable = Nothing, Optional ByRef ColumnValsList As List(Of String) = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If mValue Is Nothing Then
            Return ""
            Exit Function
        End If
        If IsDBNull(mValue) = True Then
            Return ""
            Exit Function
        End If

        Dim aFormat() As String = mFormatString.Split("~")
        Dim sValue As String = mValue.ToString
        Select Case UCase(aFormat(0))
            Case "D"
                Dim mValue1 As DateTime = CDate(mValue)
                If aFormat.Count > 1 Then
                    Dim sFormat As String = aFormat(1)
                    sValue = mValue1.ToString(sFormat)
                End If
            Case "N"
                Dim mValue1 As Decimal = CDec(mValue)
                If aFormat.Count > 1 Then
                    Dim sFormat As String = aFormat(1)
                    sValue = DecimalToString(mValue1, sFormat)
                End If
            Case "I"
                Dim mValue1 As Integer = CInt(mValue)
                If aFormat.Count > 1 Then
                    Dim sFormat As String = aFormat(1)
                    sValue = mValue1.ToString(sFormat, System.Globalization.CultureInfo.InvariantCulture)
                End If
            Case "A"
                'If aFormat.Count = 2 Then
                Dim svalarr() As String = Split(sValue, "~")
                Dim masteroption_key As Integer = -1
                If Not svalarr(0).Trim = "" Then masteroption_key = svalarr(0)
                '  Dim KeyIndex As String = MasterOption_key.ToString & "~" & mValue.ToString
                sValue = GetMasterOptionsValue(sValue)
                ' End If
            Case "E"
                'Evaluate expression
                If aFormat.Count > 1 Then
                    Dim mexpr As String = aFormat(1)
                    Dim rexpr As String = ReplaceValuesInExpression(mexpr, PublicVariables, "VB")
                    Dim merror As Boolean = False
                    sValue = EvaluateExpression(rexpr, merror)
                End If

            Case "M"
                Dim val As String = mValue.ToString
                Dim valArr() As String = Split(val, "~", -1)
                If valArr.Count = 2 Then
                    sValue = valArr(1)

                End If

            Case "O"
                'Dim methoarr() As String = Split(aFormat(1), "~")
                'If methoarr.Count = 2 Then
                '    FormattedValue(ColumnValsList(CInt(methoarr(0)) - 1), methoarr(1))
                'ElseIf methoarr.Count = 1 Then
                '    sValue = ColumnValsList(CInt(methoarr(0)) - 1)
                'End If

                '  Dim methoarr() As String = Split(aFormat(1), "~")
                If aFormat.Count = 3 Then
                    sValue = FormattedValue(ColumnValsList(CInt(aFormat(1)) - 1), aFormat(2))
                Else
                    Dim Kind As Integer = CInt(aFormat(1)) - 1
                    If ColumnValsList.Count = Kind + 1 Then sValue = ColumnValsList(CInt(aFormat(1)) - 1)
                End If


        End Select
        Return sValue
    End Function
    ''' <summary>
    ''' Set order of columns as per columnorder in GridColumns table
    ''' </summary>
    ''' <param name="DataGridView1"></param>
    ''' <remarks></remarks>
    Public Sub SetColumnOrder(ByRef mform As Object, ByRef DataGridView1 As DataGridView) 'add by divya
        Dim view As New DataView(GridColumns)
        view.Sort = "UserColumnOrder ASC"
        Dim GridColumnsTemp As DataTable = view.ToTable
        Dim DtColumnVal As String = ""
        For i = 0 To GridColumnsTemp.Rows.Count - 1
            If Not IsDBNull(GridColumnsTemp.Rows(i).Item("DtColumn")) Then
                DtColumnVal = GridColumnsTemp.Rows(i).Item("DtColumn")
                DataGridView1.Columns(DtColumnVal).DisplayIndex = i
            End If
        Next
    End Sub
    ''' <summary>
    ''' Set Width of columns as per column width in gridcolumns table
    ''' </summary>
    ''' <param name="DataGridView1"></param>
    ''' <remarks></remarks>
    Public Sub SetColumnWidth(ByRef mform As Object, ByRef DataGridView1 As DataGridView) 'add by divya
        Dim DtColumnVal As String = ""
        Dim ColumnWidthVal As Integer = 0
        For i = 0 To mform.GridColumns.Rows.Count - 1
            If Not IsDBNull(mform.GridColumns.Rows(i).Item("DtColumn")) Then
                DtColumnVal = mform.GridColumns.Rows(i).Item("DtColumn")
                If Not IsDBNull(mform.GridColumns.rows(i).item("UserColumnHeading")) Then
                    ColumnWidthVal = mform.GridColumns.Rows(i).Item("UserColumnWidth")
                    DataGridView1.Columns(DtColumnVal).MinimumWidth = 50
                    DataGridView1.Columns(DtColumnVal).Width = ColumnWidthVal
                End If
            End If
        Next
    End Sub
    ''' <summary>
    ''' Set heading of columns as per columnheading in gridcolumns table
    ''' </summary>
    ''' <param name="DataGridView1"></param>
    ''' <remarks></remarks>
    Public Sub SetColumnHeading(ByRef mform As Object, ByRef DataGridView1 As DataGridView) 'add by divya
        Dim DtColumnVal As String = ""
        Dim ColumnHeadingVal As String = ""
        For i = 0 To mform.GridColumns.Rows.Count - 1
            If Not IsDBNull(mform.GridColumns.Rows(i).Item("DtColumn")) Then
                DtColumnVal = mform.GridColumns.Rows(i).Item("DtColumn")
                If Not IsDBNull(mform.GridColumns.rows(i).item("UserColumnHeading")) Then
                    ColumnHeadingVal = mform.GridColumns.Rows(i).Item("ColumnHeading")
                    DataGridView1.Columns(DtColumnVal).HeaderText = ColumnHeadingVal
                End If
            End If
        Next
    End Sub

    Public Sub SetColumnVisiblity(ByRef mform As Object, ByRef DataGridView1 As DataGridView) 'add by divya
        Dim DtColumnVal As String = ""
        Dim ColumnVisibleVal As Boolean = True
        For i = 0 To mform.GridColumns.Rows.Count - 1
            If Not IsDBNull(mform.GridColumns.Rows(i).Item("DtColumn")) Then
                DtColumnVal = mform.GridColumns.Rows(i).Item("DtColumn")
                If Not IsDBNull(mform.GridColumns.rows(i).item("UserVisible")) Then
                    ColumnVisibleVal = mform.GridColumns.Rows(i).Item("UserVisible")
                    DataGridView1.Columns(DtColumnVal).Visible = ColumnVisibleVal
                End If
            End If
        Next
    End Sub




    Public Function ProcessOtherDetailsInfoTableGrid(ByVal otherdetails As String) As List(Of String)

        Dim a As New List(Of String)
        Dim OtherDetailsArr() As String = Split(otherdetails, ChrW(201))
        For k = 0 To OtherDetailsArr.Count - 1
            Dim ValPairArr() As String = Split(OtherDetailsArr(k), ChrW(200))
            If ValPairArr.Count = 2 Then a.Add(ValPairArr(1))

        Next
        Return a
    End Function

    ''' <summary>
    ''' Get Option Value as string from  MasterOptions datatable ,input string is MasterOption_key~Index,Conversion code "AFV0"
    ''' </summary>
    ''' <param name="KeyOfValuesAndIndex"> String type value as MasterOption_key~Index </param>
    ''' <returns>String correspnding to Index in MasterOption_key row of MasterOptions datatable</returns>
    ''' <remarks></remarks>
    Public Function GetMasterOptionsValue(ByVal KeyOfValuesAndIndex As String, Optional ByVal DtMasterOptions As DataTable = Nothing) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mstr As String = ""
        Try
            Dim Apart() As String = Split(KeyOfValuesAndIndex.Trim, "~")
            If DtMasterOptions Is Nothing Then
                DtMasterOptions = GlobalControl.Variables.MasterOptions
            End If

            If Apart.Count > 1 Then
                If Microsoft.VisualBasic.IsNumeric(Apart(1)) = True Then
                    Dim astr() As String = Split(DtMasterOptions.Rows.Find(Apart(0)).Item("ValuesSet").ToString, "~")
                    If Apart(1) >= 0 Then mstr = astr(CInt(Apart(1)))
                Else
                    mstr = Apart(1)
                End If
            Else
                mstr = KeyOfValuesAndIndex
            End If
        Catch ex As Exception
            QuitError(ex, Err, "Unable to execute Get Option Value as string from  MasterOptions datatable ,input string is MasterOption_key~Index,Conversion code AFV0")
        End Try
        Return mstr
    End Function


    ''' <summary>
    ''' Check Internet connection on a default port 80
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckInternetConnection(Optional ByVal mPortNo As Int16 = 80) As Boolean
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Try
            Dim mclient As System.Net.Sockets.TcpClient = New System.Net.Sockets.TcpClient("www.google.com", mPortNo)
            mclient.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    '---------This function converts decimal value to text according to given numeric format---------'
    ''' <summary>
    ''' Convert a decimal value in to a comma separated grouped value.
    ''' </summary>
    ''' <param name="NumericValue">Numeric value as decimal</param>
    ''' <param name="NumericFormat">NumericFormat as ##,##,###.##</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DecimalToString(ByVal NumericValue As Decimal, ByVal NumericFormat As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim decimaltotext As String = ""
        Dim FDecPosition As Integer = NumericFormat.IndexOf(".") + 1
        Dim FDecimalPart As String = ""
        Dim FWholePart As String = NumericFormat
        If FDecPosition > 1 Then
            FDecimalPart = Microsoft.VisualBasic.Right(NumericFormat, NumericFormat.Length - FDecPosition)
            FWholePart = Microsoft.VisualBasic.Left(NumericFormat, FDecPosition - 1)
        End If
        Dim nodec As Integer = FDecimalPart.Length
        Dim IDecPosition As Integer = NumericValue.ToString.IndexOf(".") + 1
        Dim IDecimalPart As String = ""
        Dim IWholePart As String = NumericValue.ToString
        If Microsoft.VisualBasic.Left(NumericValue.ToString, 1) = "-" Then
            IWholePart = NumericValue.ToString.Substring(1, NumericValue.ToString.Length - 1)
        End If
        If IDecPosition > 1 Then
            IDecimalPart = Microsoft.VisualBasic.Right(NumericValue.ToString, NumericValue.ToString.Length - IDecPosition)
            If Microsoft.VisualBasic.Left(NumericValue.ToString, 1) = "-" Then
                IWholePart = Microsoft.VisualBasic.Left(IWholePart.ToString, IDecPosition - 2)
            Else
                IWholePart = Microsoft.VisualBasic.Left(NumericValue.ToString, IDecPosition - 1)
            End If
        End If
        Dim FReverse As String = StringReverse(FWholePart)
        Dim IReverse As String = StringReverse(IWholePart)
        Dim mchar As String = ""
        For i = 1 To FReverse.Length
            mchar = Mid(FReverse, i, 1)
            If i > IReverse.Length Then
                Exit For
            End If
            If Not "-0123456789#.".Contains(mchar) Then
                IReverse = IReverse.Insert(i - 1, mchar)
            End If
        Next
        decimaltotext = IIf(IWholePart >= "0", StringReverse(IReverse), "")
        If FDecimalPart.Length > 0 Then
            decimaltotext = decimaltotext & "." & Microsoft.VisualBasic.Left(IDecimalPart & Microsoft.VisualBasic.StrDup(FDecimalPart.Length, "0"), FDecimalPart.Length)
        End If
        '----Increment in cursorindex when every single comma is increased--------'
        If decimaltotext.Contains("-") Then
            decimaltotext = decimaltotext.Replace("-", "(-)")
        End If
        Return decimaltotext
    End Function
    '--------This function reverses a string----------'
    Private Function StringReverse(ByRef inputstring As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim ss As String = ""
        For i = inputstring.Length To 1 Step -1
            ss = ss & Mid(inputstring, i, 1)
        Next
        Return ss
    End Function
    ''' <summary>
    ''' Create a new instance from a Master Object HashTable by its KeyName as objectname.
    ''' </summary>
    ''' <param name="ObjectName">KeyName of object of new instance in hashtable.</param>
    ''' <param name="NewInstanceName">Name of new instance </param>
    ''' <param name="ObjectsHashTable">Master Hashtable having all objects.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NewInstanceByObject(ByVal ObjectName As String, ByVal NewInstanceName As String, Optional ByVal ObjectsHashTable As Hashtable = Nothing) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mobj As New Object
        Try
            If ObjectsHashTable Is Nothing Then
                ObjectsHashTable = GlobalControl.Variables.ObjectsHashTable
            End If
            mobj = GetValueFromHashTable(ObjectsHashTable, ObjectName)
            Dim mtype As Type = mobj.GetType
            mobj = mtype.GetConstructor(New System.Type() {}).Invoke(New Object() {})
            mobj = Convert.ChangeType(mobj, mtype)
            Dim mProperty As PropertyInfo = mobj.GetType.GetProperty("Name", BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.Public)
            If mProperty IsNot Nothing Then
                mobj.Name = NewInstanceName
            End If
        Catch ex As Exception
            QuitError(ex, Err, ErrorString)
        End Try
        Return mobj
    End Function
    ''' <summary>
    ''' Add items to a new or existing combobox , according to MasterOptionsString
    ''' </summary>
    ''' <param name="MasterOptionsString">RangeCode~Index ,Where RangeCode is the row element of Global datatable variable </param>
    ''' <param name="ComboObject"></param>
    ''' <param name="DtMasterOptions"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddComboBoxItems(ByVal MasterOptionsString As String, Optional ByRef ComboObject As ComboBox = Nothing, Optional ByVal DtMasterOptions As DataTable = Nothing) As System.Windows.Forms.ComboBox
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If MasterOptionsString.Trim.Length = 0 Then
            Return ComboObject
            Exit Function
        End If
        If ComboObject Is Nothing Then
            ComboObject = New System.Windows.Forms.ComboBox
        End If
        If DtMasterOptions Is Nothing Then
            DtMasterOptions = GlobalControl.Variables.MasterOptions
        End If
        ComboObject.Items.Clear()
        Dim astr() As String = MasterOptionsString.Split("~")
        Dim RowIndex As Int16 = CInt(astr(0))
        Dim df1 As New DataFunctions.DataFunctions
        Dim bstr() As String = df1.FindRowByPrimaryCols(DtMasterOptions, RowIndex).Item("ValuesSet").ToString.Split("~")
        ComboObject.Items.AddRange(bstr)
        If astr.Length > 1 Then
            Dim mindex As Int16 = CInt(astr(1))
            ComboObject.SelectedIndex = mindex
            ComboObject.Text = CStr(ComboObject.Items(mindex))
        End If
        Return ComboObject
    End Function
    ''' <summary>
    ''' Add items to a new or existing ListBox , according to MasterOptionsString
    ''' </summary>
    ''' <param name="MasterOptionsString">RangeCode~Index ,Where RangeCode is the row element of Global datatable variable </param>
    ''' <param name="ListObject"></param>
    ''' <param name="DtMasterOptions"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function AddListBoxItems(ByVal MasterOptionsString As String, Optional ByRef ListObject As ListBox = Nothing, Optional ByVal DtMasterOptions As DataTable = Nothing) As System.Windows.Forms.ListBox
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If MasterOptionsString.Trim.Length = 0 Then
            Return ListObject
            Exit Function
        End If
        If ListObject Is Nothing Then
            ListObject = New System.Windows.Forms.ListBox
        End If
        If DtMasterOptions Is Nothing Then
            DtMasterOptions = GlobalControl.Variables.MasterOptions
        End If
        ListObject.Items.Clear()
        Dim astr() As String = MasterOptionsString.Split("~")
        Dim RowIndex As Int16 = CInt(astr(0))
        Dim df1 As New DataFunctions.DataFunctions
        Dim bstr() As String = df1.FindRowByPrimaryCols(DtMasterOptions, RowIndex).Item("ValuesSet").ToString.Split("~")
        ListObject.Items.AddRange(bstr)
        If astr.Length > 1 Then
            Dim mindex As Int16 = CInt(astr(1))
            ListObject.SelectedIndex = mindex
            ListObject.SelectedItem = ListObject.Items(mindex)
        End If
        Return ListObject
    End Function
    ''' <summary>
    ''' Return True if the given keyname is present in the given hashtable otherwise false.
    ''' </summary>
    ''' <param name="ht1"></param>Represent a Hashtable.
    ''' <param name="Keyname"></param>Represent the keyname value which has to be searched in the given hashtable.
    ''' <returns></returns>
    Public Function CheckIfKeyNameExistsinHashTable(ByVal ht1 As Hashtable, ByVal Keyname As String) As Boolean
        Dim bln As Boolean = False
        Dim KeyCol As ICollection = ht1.Keys
        For i = 0 To KeyCol.Count - 1
            Dim a As String = KeyCol(i)
            If LCase(Keyname) = LCase(KeyCol(i)) Then
                bln = True
            End If
        Next
        Return bln
    End Function



    ''' <summary>
    ''' Create a new instance from a Master Object HashTable by its KeyName as objectname.
    ''' </summary>
    ''' <param name="BaseObject">BaseObject for  new instance.</param>
    ''' <param name="NewInstanceName">Name of new instance </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NewInstanceByObject(ByVal BaseObject As Object, ByVal NewInstanceName As String) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mobj As New Object
        Try
            Dim mtype As Type = BaseObject.GetType
            mobj = mtype.GetConstructor(New System.Type() {}).Invoke(New Object() {})
            mobj = Convert.ChangeType(mobj, mtype)
            Dim mProperty As PropertyInfo = mobj.GetType.GetProperty("Name", BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.Public)
            If mProperty IsNot Nothing Then
                mobj.Name = NewInstanceName
            End If
        Catch ex As Exception
            QuitError(ex, Err, ErrorString)
        End Try
        Return mobj
    End Function



    ''' <summary>
    ''' Create a new instance from a Master Object HashTable by its KeyName as objectname.
    ''' </summary>
    ''' <param name="TypeName">Key of  type of new instance in hashtable.</param>
    ''' <param name="NewInstanceName">Name of new instance </param>
    ''' <param name="mTypesHashTable">Master Hashtable having all types key is typename and value as type object.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NewInstanceByType(ByVal TypeName As String, ByVal NewInstanceName As String, Optional ByVal mTypesHashTable As Hashtable = Nothing) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mobj As New Object
        Try
            If mTypesHashTable Is Nothing Then
                mTypesHashTable = GlobalControl.Variables.TypesHashTable
            End If
            Dim mtype As Type = GetValueFromHashTable(mTypesHashTable, TypeName)
            mobj = mtype.GetConstructor(New System.Type() {}).Invoke(New Object() {})
            mobj = Convert.ChangeType(mobj, mtype)
            Dim mProperty As PropertyInfo = mobj.GetType.GetProperty("Name", BindingFlags.IgnoreCase Or BindingFlags.Instance Or BindingFlags.Public)
            If mProperty IsNot Nothing Then
                mobj.Name = NewInstanceName
            End If
        Catch ex As Exception
            QuitError(ex, Err, "NewInstanceByType(ByVal TypeName As String, ByVal NewInstanceName As String, Optional ByVal TypesHashTable As Hashtable = Nothing) As Object")
        End Try
        Return mobj
    End Function
    ''' <summary>
    ''' Create a new instance from a Master Object HashTable by its KeyName as objectname.
    ''' </summary>
    ''' <param name="TypeCode">TypeCode of  new instance in hashtable.</param>
    ''' <param name="NewInstanceName">Name of new instance </param>
    ''' <param name="DtControlsList" >Full Control List as datatable</param>
    ''' <param name="mTypesHashTable">Master Hashtable having all types key is typename and value as type object.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NewInstanceByTypeCode(ByVal TypeCode As String, ByVal NewInstanceName As String, ByVal DtControlsList As DataTable, Optional ByVal mTypesHashTable As Hashtable = Nothing) As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim mobj As New Object
        Try
            If mTypesHashTable Is Nothing Then
                mTypesHashTable = GlobalControl.Variables.TypesHashTable
            End If
            Dim mrows() As DataRow = DtControlsList.Select("ControlType  = '" & TypeCode & "'")
            If mrows IsNot Nothing Then
                If mrows.Count > 0 Then
                    Dim mOwner As String = mrows(0).Item("Owner")
                    If mOwner = "S" Then
                        Dim mControlType As String = mrows(0).Item("ControlTypeName")
                        mControlType = IIf(mControlType = "Button1", "Button", mControlType)
                        mobj = NewInstanceByType(mControlType, NewInstanceName)
                    Else
                        Dim mDllfile As String = mrows(0).Item("DLLFolder")
                        Dim mFullControlTypeName As String = mrows(0).Item("FullControlTypeName")
                        mobj = NewInstanceByDLL(mDllfile, mFullControlTypeName)
                    End If
                End If
            End If
        Catch ex As Exception
            QuitError(ex, Err, "NewInstanceByTypeCode(ByVal TypeCode As String, ByVal NewInstanceName As String, ByVal DtControlsList As DataTable, Optional ByVal mTypesHashTable As Hashtable = Nothing) As Object")
        End Try
        Return mobj
    End Function



    ''' <summary>
    ''' Function to create an instance of a DLL assembly.
    ''' </summary>
    ''' <param name="DLLFullPath">Dll file full path</param>
    ''' <param name="NameOfClass">Optional, if main class name is different from DLL name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function NewInstanceByDLL(ByVal DLLFullPath As String, Optional ByVal NameOfClass As String = "") As Object
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim otype As System.Type = Nothing
        Dim oAssembly As System.Reflection.Assembly
        Dim oObject As System.Object
        oAssembly = Assembly.LoadFrom(DLLFullPath)
        Dim aPath As List(Of String) = FullFileNameToList(DLLFullPath)
        If NameOfClass.Length = 0 Then
            NameOfClass = aPath(1).Trim
        End If
        Dim aTypes() As Type = oAssembly.GetTypes
        For i = 0 To aTypes.Length - 1
            If LCase(oAssembly.GetTypes(i).FullName) = LCase(NameOfClass) Then
                otype = aTypes(i)
                Exit For
            End If
        Next
        If otype Is Nothing Then
            QuitMessage("Instance name not found in assembly", "NewInstanceByDLL(ByVal DLLFullPath As String, Optional ByVal NameOfClass As String = "") As Object")
        End If
        oObject = Activator.CreateInstance(otype)
        oObject = Convert.ChangeType(oObject, otype)
        Return oObject
    End Function
    ''' <summary>
    ''' System copy of a file into given destination if file already exists in given destination.New FileName created by subscript.
    ''' </summary>
    ''' <param name="FileToBeCopied">FileName with folder to be copied</param>
    ''' <param name="TargetPath" >Folder location A=In the application bin, D=In the DataFolder of application ,T=In TempFolder of application or any other fixed path.</param>
    ''' <param name="ConfirmCopy" >Show message and ask for confirmation ,if new file name is created.</param>
    ''' <returns>New File Name</returns>
    ''' <remarks></remarks>
    Public Function CopyFileNewNameIfExist(ByVal FileToBeCopied As String, ByVal TargetPath As String, Optional ByVal DoNotCopy As Boolean = False, Optional ByVal ConfirmCopy As Boolean = False) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim sourcepath = FileToBeCopied
        Dim DestPath = CreateChildPathInFolderLocation(TargetPath)
        Dim fileDetails = FullFileNameToList(FileToBeCopied)
        Dim FileName = fileDetails(1)
        Dim flag As Boolean = False
        Dim newPath As String = DestPath & "\" & fileDetails(1) & "." & fileDetails(2)
        Dim i = 0
        Do
            flag = My.Computer.FileSystem.FileExists(newPath)
            If flag = True Then
                i = i + 1
                FileName = fileDetails(1) & "-Copy(" & i & ")"
                newPath = DestPath & "\" & FileName & "." & fileDetails(2)
            End If
        Loop While (flag = True)
        Try
            If DoNotCopy = False Then
                If i > 0 And ConfirmCopy = True Then
                    If MessageBox.Show("File will be copied as " & newPath) = DialogResult.OK Then
                        FileCopy(sourcepath, newPath)
                    End If
                Else
                    FileCopy(sourcepath, newPath)
                End If
            End If
        Catch ex As Exception
            QuitMessage("Unable to copy file " & newPath, "Public Function CopyFileNewNameIfExist(ByVal FileToBeCopied As String, ByVal TargetPath As String, Optional ByVal DoNotCopy As Boolean = False, Optional ByVal ConfirmCopy As Boolean = False) As String")
        End Try
        Return newPath
    End Function
    ''' <summary>
    ''' Create SubFolder in some given data folder.
    ''' </summary>
    ''' <param name="FolderLocation" >Folder location A=In the application bin, D=In the DataFolder of application ,T=In TempFolder of application or any other fixed path.</param>
    ''' <param name="ChildFolderName">SubFolderName to be added in FolderLocation as child</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateChildPathInFolderLocation(ByVal FolderLocation As String, Optional ByVal ChildFolderName As String = "") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim DPath As String = GetChildPathInFolderLocation(FolderLocation, ChildFolderName)
        If My.Computer.FileSystem.DirectoryExists(DPath) = False Then
            Try
                My.Computer.FileSystem.CreateDirectory(DPath)
            Catch ex As Exception
                QuitMessage("Unable to create folder " & DPath, "CreateChildPathInFolderLocation(ByVal FolderLocation As String, Optional ByVal ChildFolderName As String = "") As String")
            End Try
        End If
        Return DPath
    End Function
    ''' <summary>
    ''' Get child folder path as string in FolderLocation.
    ''' </summary>
    ''' <param name="FolderLocation">Folder location A=In the application bin, D=In the DataFolder of application ,T=In TempFolder of application or any other fixed path.</param>
    ''' <param name="ChildFolderName">SubFolderName to be added in FolderLocation as child</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetChildPathInFolderLocation(ByVal FolderLocation As String, Optional ByVal ChildFolderName As String = "") As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        If FolderLocation.Length = 0 Then
            QuitMessage("FolderLocation  could not be blank", "CreateSubFolderInFolder(ByVal FolderName As String, ByVal DataStore As String) As String")
        End If
        Dim datastore As String = FolderLocation
        If FolderLocation.Length = 1 Then
            Select Case FolderLocation
                Case "A"
                    datastore = System.Windows.Forms.Application.ExecutablePath
                    Dim PathList As List(Of String) = FullFileNameToList(datastore)
                    datastore = PathList(0)
                Case "D"
                    datastore = GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "DataFolder")
                Case "T"
                    datastore = GetValueFromHashTable(GlobalControl.Variables.AppControlHashTable, "TempFolder")
            End Select
        End If
        Dim DPath As String = IIf(ChildFolderName.Length = 0, datastore, datastore & "\" & ChildFolderName)
        Return DPath
    End Function
    ''' <summary>
    ''' Create folder in application.
    ''' </summary>
    ''' <param name="FolderName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateFolderInApplication(ByVal FolderName As String) As String
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        Dim PathList As List(Of String) = FullFileNameToList(System.Windows.Forms.Application.ExecutablePath)
        Dim DPath As String = PathList(0) & "\" & FolderName
        If My.Computer.FileSystem.DirectoryExists(DPath) = False Then
            Try
                My.Computer.FileSystem.CreateDirectory(DPath)
            Catch ex As Exception
                QuitMessage("Unable to create folder " & DPath, "")
            End Try
        End If
        Return DPath
    End Function
    ''' <summary>
    ''' Permissible Form Location TopLeft,TopCenter,TopRight,MiddleLeft,MiddleCenter,MiddleRight,BottomLeft,BottomCenter,BottomRight.
    ''' </summary>
    ''' <param name="FormObj">Form object whose location to be set</param>
    ''' <param name="FormPosition">Permissible Form Location TopLeft,TopCenter,TopRight,MiddleLeft,MiddleCenter,MiddleRight,BottomLeft,BottomCenter,BottomRight</param>
    ''' <param name="LocationReference">If with respect to screen , then left nothing</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetFormLocation(ByRef FormObj As Form, ByVal FormPosition As String, Optional ByVal LocationReference As Object = Nothing) As Point
        If GlobalControl.Variables.AuthenticationChecked <> dllInteger Then Return Nothing : Exit Function
        FormObj.StartPosition = FormStartPosition.Manual
        Dim ContainerTop As Integer = 0
        Dim ContainerLeft As Integer = 0
        Dim ContainerWidth As Integer = My.Computer.Screen.WorkingArea.Width
        Dim ContainerHeight As Integer = My.Computer.Screen.WorkingArea.Height
        Dim BorderWidth As Integer = 0
        Dim TitlebarHeight As Integer = 0
        If LocationReference IsNot Nothing Then
            Dim comparer As Comparer(Of Form) = comparer.Default
            If comparer.Equals(LocationReference, LocationReference.FindForm) = False Then
                BorderWidth = (LocationReference.FindForm.Width - LocationReference.FindForm.ClientSize.Width) / 2
                TitlebarHeight = LocationReference.FindForm.Height - LocationReference.FindForm.ClientSize.Height - 2 * BorderWidth
                ContainerTop = LocationReference.FindForm.Location.Y + TitlebarHeight + 2 * BorderWidth + LocationReference.Top - 2
                ContainerLeft = LocationReference.FindForm.Location.X + 2 * BorderWidth + LocationReference.Left - 1
                ContainerWidth = LocationReference.FindForm.Width
                ContainerHeight = LocationReference.FindForm.Height
            Else
                BorderWidth = (LocationReference.Width - LocationReference.ClientSize.Width) / 2
                TitlebarHeight = LocationReference.Height - LocationReference.ClientSize.Height - 2 * BorderWidth
                ContainerTop = LocationReference.Location.Y + TitlebarHeight + BorderWidth
                ContainerLeft = LocationReference.Location.X + BorderWidth
                ContainerWidth = LocationReference.Width
                ContainerHeight = LocationReference.Height
            End If
        End If
        Dim FormLocation As New Point
        Select Case LCase(FormPosition)

            Case "topleft"
                FormLocation = New Point(ContainerLeft, ContainerTop)
            Case "topcenter"
                FormLocation = New Point(ContainerLeft + (LocationReference.Width - FormObj.Width) / 2, ContainerTop)
            Case "topright"
                FormLocation = New Point(ContainerLeft + (LocationReference.Width - FormObj.Width), ContainerTop)
            Case "middleleft"
                FormLocation = New Point(ContainerLeft, ContainerTop + (LocationReference.Height - FormObj.Height) / 2)
            Case "middlecenter"
                FormLocation = New Point(ContainerLeft + (LocationReference.Width - FormObj.Width) / 2, ContainerTop + (LocationReference.Height - FormObj.Height) / 2)
            Case "middleright"
                FormLocation = New Point(ContainerLeft + (LocationReference.Width - FormObj.Width), ContainerTop + (LocationReference.Height - FormObj.Height) / 2)
            Case "bottomleft"
                FormLocation = New Point(ContainerLeft, ContainerTop + LocationReference.Height)
            Case "bottomcenter"
                FormLocation = New Point(ContainerLeft + (LocationReference.Width - FormObj.Width) / 2, ContainerTop + LocationReference.Height)
            Case "bottomright"
                FormLocation = New Point(ContainerLeft + (LocationReference.Width - FormObj.Width), ContainerTop + LocationReference.Height)
        End Select
        If FormLocation.X + FormObj.Width > My.Computer.Screen.WorkingArea.Width Then
            FormLocation.X = My.Computer.Screen.WorkingArea.Width - FormObj.Width
        End If
        If FormLocation.Y + FormObj.Height > My.Computer.Screen.WorkingArea.Height Then
            FormLocation.Y = My.Computer.Screen.WorkingArea.Height - FormObj.Height
        End If
        Return FormLocation
    End Function
#End Region
End Class









