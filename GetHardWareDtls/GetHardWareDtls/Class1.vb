
'Imports 
Public Class GetHardWareDetails

    Dim df1 As New DataFunctions.DataFunctions
    Public Function getHardwareDtsDatatable(ByVal machineLineEPL As String, ByVal p_customers As Integer) As DataTable
        Dim dtHash1 As New DataTable
        Dim prevhrdStr As String = GetPreviousHardwareMachineStr(p_customers)
        dtHash1 = df1.AddColumnsInDataTable(dtHash1, "p_customers,drive,driveno,windowsver,baseboard,bios,cpu,processor")
        dtHash1 = MAchineLineDBToDatatable(prevhrdStr, dtHash1)
        dtHash1 = processMachineLine(machineLineEPL, dtHash1)
        For o = 0 To dtHash1.Rows.Count - 1
            dtHash1.Rows(o).Item("p_customers") = CStr(p_customers)
        Next
        Return dtHash1
    End Function
    Public Function GetPreviousHardwareMachineStr(ByVal p_customers As Integer) As String
        Dim clsregis As New Registrations.Registrations.Registrations
        Dim SqlStr As String = "select * from registrations where rowstatus = 0 and p_customers=" & p_customers
        Dim dtreg As DataTable = df1.SqlExecuteDataTable(clsregis.ServerDatabase, SqlStr)
        Dim machineline As String = ""
        For u = 0 To dtreg.Rows.Count - 1
            machineline = df1.GetCellValue(dtreg.Rows(u), "hardwarestring")

        Next

        'If machineline = "" Then
        '    Dim sqlregtran As String = "select * from registrationtran where p_customers = " & p_customers & " order by registrationtran_key desc"
        '    Dim dtregtran As DataTable = df1.SqlExecuteDataTable(clsregis.ServerDatabase, sqlregtran)
        '    If dtregtran.Rows.Count > 0 Then
        '        machineline = df1.GetCellValue(dtregtran.Rows(0), "machineline")
        '    End If

        'End If
        Return machineline
    End Function

    Public Function processMachineLine(ByVal machineLine As String, ByRef dthash As DataTable) As DataTable
        '   Dim Dthash As New DataTable


        Dim hrdarr() As String = Split(machineLine, "~")
        If hrdarr.Count = 9 Then
            dthash.Rows.Add()
            dthash.Rows(dthash.Rows.Count - 1).Item("drive") = hrdarr(0)
            dthash.Rows(dthash.Rows.Count - 1).Item("driveno") = hrdarr(2)
            dthash.Rows(dthash.Rows.Count - 1).Item("windowsver") = hrdarr(3)
            dthash.Rows(dthash.Rows.Count - 1).Item("baseboard") = hrdarr(5)


            dthash.Rows(dthash.Rows.Count - 1).Item("bios") = hrdarr(6)
            dthash.Rows(dthash.Rows.Count - 1).Item("cpu") = hrdarr(7)
            dthash.Rows(dthash.Rows.Count - 1).Item("processor") = hrdarr(8)



        End If

        Return dthash

    End Function



    Public Function MAchineLineDBToDatatable(ByVal machineline As String, ByRef machhashtab As DataTable, Optional ByVal machtype As String = "main") As DataTable
        Dim machinarr() As String = Split(machineline, Chr(201))
        ' Dim machHAshTab As New DataTable

        For o = 0 To machinarr.Count - 1
            Dim mchinind As String = machinarr(o)
            Dim finalhrdStrArr() As String = Split(mchinind, Chr(200))
            If finalhrdStrArr.Count = 2 Then
                Select Case machtype
                    Case "main"
                        If finalhrdStrArr(0) = "main" Then
                            machhashtab = processMachineLine(finalhrdStrArr(1), machhashtab)
                        End If
                    Case "home"
                        If finalhrdStrArr(0).Substring(0, 4) = "home" Then
                            machhashtab = processMachineLine(finalhrdStrArr(1), machhashtab)

                        End If

                        '  Case Else



                End Select

            End If
        Next

        Return machhashtab

    End Function


End Class