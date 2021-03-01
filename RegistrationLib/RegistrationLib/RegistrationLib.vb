﻿Imports System.Data.SqlClient
Imports System.IO
Imports System.IO.Compression

Public Class RegistrationLib
    Dim df1 As New DataFunctions.DataFunctions
    Dim GF1 As New GlobalFunction1.GlobalFunction1
    Dim GF2 As New GlobalFunction2.GlobalFunction2
    Dim cfc As New CommonFunctionsCloud.CommonFunctionsCloud
    Dim libcustomerfeature As New CustomerFeatureLib.CustomerFeatureFunctions
    Dim libSaralAuth As New SaralAuthLib.LoginFunctions

#Region "CSVMT"
    ''' <summary>
    ''' This Function fetches service no from CustomerService table. 
    ''' </summary>
    ''' <param name="P_customers"></param>
    ''' <returns>Service no</returns>
    Public Function GetServiceNoCservices(ByVal P_customers As Integer) As Integer
        Dim clscustomerservices As New CustomerServices.CustomerServices.CustomerServices
        Dim serviceno As Int16 = 0
        Dim dtable As DataTable = df1.SqlExecuteDataTable(clscustomerservices.ServerDatabase, "Select top (1) * from customerservices where P_customers = " & P_customers & " order by serviceno desc")
        If dtable.Rows.Count > 0 Then
            serviceno = CInt(dtable.Rows(0).Item("serviceno"))
        End If
        Return serviceno + 1
        '   Dim dtrow As DataTable = df1.GetDataFromSql(clscustomerservices.ServerDatabase, clscustomerservices.TableName, "serviceno", "", "P_customers=" & P_customers)
    End Function


    ''' <summary>
    ''' FUNCTION TO GET THE COUNT OF ROWS of unpaid 15 days registartions
    ''' </summary>
    ''' <returns></returns>
    Public Function getCountofUnpaidReg() As Integer
        Dim fromDate As DateTime = df1.getDateTimeISTNow.AddDays(-30)
        Dim Todate As DateTime = df1.getDateTimeISTNow
        Dim FromdateStr As String = fromDate.ToString("yyyy-MM-dd 00:00:00.00")
        Dim TodateStr As String = Todate.ToString("yyyy-MM-dd 23:59:59.999")

        Dim Query As String = String.Format("select Distinct P_Customers from [chargingheader] where Mtimestamp between '" & FromdateStr & "' and '" & TodateStr & "'  and PaYmentFlag='U' AND rOWSTATUS=0")
        Dim ClsPayment As New Payment.Payment.Payment
        Dim chDetails As DataTable = df1.SqlExecuteDataTable(ClsPayment.ServerDatabase, Query)
        Dim P_Customers As String = ""
        For i = 0 To chDetails.Rows.Count - 1
            P_Customers += "," & chDetails.Rows(i).Item("P_Customers").ToString.Trim
        Next
        If P_Customers.StartsWith(",") Then P_Customers = P_Customers.Substring(1)
        Dim q1 As String = String.Format("Select count(*) as Rcount from Customers where P_Customers in (" & P_Customers & ") and Rowstatus=0")
        Dim RegistrationDetails As DataTable = df1.SqlExecuteDataTable(ClsPayment.ServerDatabase, q1)
        Return RegistrationDetails.Rows(0).Item("Rcount")
    End Function
    ''' <summary>
    ''' paginated function to get dt of 15days open regstrations
    ''' </summary>
    ''' <param name="start"></param>
    ''' <param name="psize"></param>
    ''' <param name="Dtinfotable"></param>
    ''' <returns></returns>
    Public Function RegistrationOpenUnpaiddt(start As Integer, psize As Integer, Dtinfotable As DataTable) As DataTable
        Dim fromDate As DateTime = df1.getDateTimeISTNow.AddDays(-30)
        Dim Todate As DateTime = df1.getDateTimeISTNow
        Dim FromdateStr As String = fromDate.ToString("yyyy-MM-dd 00:00:00.00")
        Dim TodateStr As String = Todate.ToString("yyyy-MM-dd 23:59:59.999")

        Dim Query As String = String.Format("select Distinct P_Customers from [chargingheader] where Mtimestamp between '" & FromdateStr & "' and '" & TodateStr & "'  and PaYmentFlag='U' AND rOWSTATUS=0")
        Dim ClsPayment As New Payment.Payment.Payment
        Dim chDetails As DataTable = df1.SqlExecuteDataTable(ClsPayment.ServerDatabase, Query)
        Dim P_Customers As String = ""
        For i = 0 To chDetails.Rows.Count - 1
            P_Customers += "," & chDetails.Rows(i).Item("P_Customers").ToString.Trim
        Next
        If P_Customers.StartsWith(",") Then P_Customers = P_Customers.Substring(1)
        ' Dim q1 As String = String.Format("Select P_Customers,CustCode,CustName,MobNo,PostalAddress1,PostalAddress2,PostalAddress3,PostalAddress4,ServicingAgentCode,HometOWN,CurrRegDate from Customers where  Order by CurrRegDate Desc")
        Dim RegistrationDetails As DataTable = df1.GetDataFromSqlFixedRows(ClsPayment.ServerDatabase, "Customers", "P_Customers,CustCode,CustName,MobNo,PostalAddress1,PostalAddress2,PostalAddress3,PostalAddress4,ServicingAgentCode,HometOWN,CurrRegDate", "", "P_Customers in (" & P_Customers & ") and Rowstatus=0", "", "CurrRegDate Desc", start, psize, -1)
        'Dim RegistrationDetails As DataTable = df1.SqlExecuteDataTable(ClsPayment.ServerDatabase, q1)

        RegistrationDetails = df1.AddingNameForCodesPrimamryCols(RegistrationDetails, "HomeTown", "TextHomeTown", Dtinfotable, "NameOfInfo")
        RegistrationDetails = df1.AddColumnsInDataTable(RegistrationDetails, "S.no",,, "CustCode")
        RegistrationDetails = df1.AddColumnsInDataTable(RegistrationDetails, "CombinedAddress,OpenedUptoStr,BilledUptoStr,CurrRegDate1,TextServicingAgentCode", "System.String,System.String,System.String,System.String,System.String")
        If RegistrationDetails.Rows.Count > 0 Then
            For i = 0 To RegistrationDetails.Rows.Count - 1
                RegistrationDetails.Rows(i).Item("S.no") = i + 1
                RegistrationDetails.Rows(i).Item("CombinedAddress") = df1.GetCellValue(RegistrationDetails.Rows(i), "PostalAddress1", "String") & " " & df1.GetCellValue(RegistrationDetails.Rows(i), "PostalAddress2", "String") & " " & df1.GetCellValue(RegistrationDetails.Rows(i), "PostalAddress3", "String") & " " & df1.GetCellValue(RegistrationDetails.Rows(i), "PostalAddress4", "String")
                Dim tempdtop As DateTime = GetOpenedUptoDate(df1.GetCellValue(RegistrationDetails.Rows(i), "P_Customers", "integer"))
                Dim tempdtopstr As String = tempdtop.ToString("yyyy-MM-dd")
                RegistrationDetails.Rows(i).Item("OpenedUptoStr") = tempdtopstr
                Dim tempdtRegSendDate As DateTime = RegistrationDetails.Rows(i).Item("CurrRegDate")
                Dim tempRegSendDatstr As String = tempdtRegSendDate.ToString("yyyy-MM-dd")
                RegistrationDetails.Rows(i).Item("CurrRegDate1") = tempRegSendDatstr

                Dim tempbillupto As DateTime = libcustomerfeature.GetBilledUpToDate(df1.GetCellValue(RegistrationDetails.Rows(i), "P_Customers", "integer"), "main")
                RegistrationDetails.Rows(i).Item("BilledUptoStr") = tempbillupto.ToString("yyyy-MM-dd")

                Dim mDealer1 As DataRow = libSaralAuth.getAccMasterRowForp_acccode(df1.GetCellValue(RegistrationDetails.Rows(i), "ServicingAgentCode", "integer"))
                If Not mDealer1 Is Nothing Then
                    RegistrationDetails.Rows(i).Item("TextServicingAgentCode") = mDealer1("AccName").ToString.Trim
                End If

            Next
            RegistrationDetails = df1.AlterDataTable(RegistrationDetails, "", "ServicingAgentCode,CurrRegDate,PostalAddress1,PostalAddress2,PostalAddress3,PostalAddress4")
        End If
        Return RegistrationDetails
    End Function



    ''' <summary>
    ''' This Function Creates datatable from Epl file after Epl is Uploaded for Customer Registration.
    ''' </summary>
    ''' <param name="CustDT">datatable created from  GF1.CreateDataTableFromHashTable(abc) function.</param>
    ''' <param name="Lfilenames">Full path and file name of the uploaded epl file(s)</param>
    ''' <param name="Dealerrow">linkcode of Current Login user from Websessions table.</param>
    ''' <returns>datatable from epl with calculated date columns.</returns>
    Public Function CreateDtFromEpl(CustDT As DataTable, Lfilenames As String, dealerRow As DataRow) As DataTable
        Dim ClsLinkOldCodes As New LinkOldCodes.LinkOldCodes.LinkOldCodes
        Dim dt As New DataTable
        dt = df1.AddColumnsInDataTable(dt, "Customers_Key,P_Customers,CustName,CustNameFromDb,HomeTown,MobNo,ProductCode,CustCode,AllowUpto,AllowUpto1,BilledUpto,BilledUptoStr,CalculatedAllowUpto,CalculatedAllowUptoStr,ChangeAllowUpto,ChangeAllowUptoStr,Lan,Nodes,regtype,Changedate,ChangeNodes,ChangeCustCode,ChangeLan,Changeregtype,eplpathserver,CurrRegDate,Custtype,isvalid,Filename, postaladdress1,postaladdress2,postaladdress3,postaladdress4,pincode,phone,email,mainbusscode,LastOpenDt,LastOpenDtStr,allowregdownload,nodesdb,gstin,saralhybrid", "System.Int32,System.Int32,System.String,System.String,System.Int32,System.String,System.Int32,System.String,System.DateTime,System.String,System.DateTime,System.String,System.DateTime,System.String,System.DateTime,System.String,System.String,System.Int32,System.String,System.String,System.Int32,System.String,System.String,System.String,System.String,System.DateTime,System.String,System.String, System.String, System.String,System.String,System.String,System.String,System.String,System.String,System.String,System.Int32,System.DateTime,System.String,System.String,system.int32,system.string,system.string")
        If Lfilenames.First() = "," Then Lfilenames = Lfilenames.Substring(1, Lfilenames.Length - 1)
        Dim lfile() As String = Split(Lfilenames, ",")
        For i = 0 To CustDT.Rows.Count - 1
            Dim HomeTownHash As Hashtable = GF1.GetHashTableFromString("OldCode=" & CustDT.Rows(i).Item("citycode"))
            Dim clslink1 As New LinkOldCodes.LinkOldCodes.LinkOldCodes
            Dim aClsObject() As Object = {ClsLinkOldCodes}
            Dim mserverdb As String = df1.GetServerMDFForTransanction(aClsObject)
            Dim mytrans As SqlTransaction = df1.BeginTransaction(mserverdb)
            Dim LinkOldCodeRow As DataRow = df1.SeekRecord(mytrans, ClsLinkOldCodes, HomeTownHash)
            Dim Busscodehash As Hashtable = GF1.GetHashTableFromString("OldCode=" & CustDT.Rows(i).Item("BussCd"))

            Dim mytrans1 As SqlTransaction = df1.BeginTransaction(mserverdb)
            Dim LinkCodeBusscode As DataRow = df1.SeekRecord(mytrans1, clslink1, Busscodehash)

            Dim newRow = dt.NewRow
            newRow("CustName") = CustDT.Rows(i).Item("clName")
            newRow("eplpathserver") = lfile(i)
            newRow("Filename") = Path.GetFileName(lfile(i))
            If LinkOldCodeRow IsNot Nothing Then
                newRow("HomeTown") = LinkOldCodeRow("InfoCode")
            Else
                newRow("HomeTown") = DBNull.Value
            End If

            Dim Mob As String = df1.GetCellValue(CustDT.Rows(i), "mobile")

            If String.IsNullOrEmpty(Mob) = False Then
                If Mob.First = "0" Then Mob = Mob.Substring(1)
                If Mob.Length > 10 Then Mob = ""
            End If
            newRow("MobNo") = Mob
            newRow("gstin") = df1.GetCellValue(CustDT.Rows(i), "gstin")
            newRow("PostalAddress1") = CustDT.Rows(i).Item("add1")
            newRow("PostalAddress2") = CustDT.Rows(i).Item("add2")
            newRow("PostalAddress3") = CustDT.Rows(i).Item("add3")
            newRow("PostalAddress4") = CustDT.Rows(i).Item("add4")
            newRow("email") = CustDT.Rows(i).Item("email")
            newRow("phone") = CustDT.Rows(i).Item("ph")
            newRow("saralhybrid") = CustDT.Rows(i).Item("saralhybrid")
            If LinkCodeBusscode IsNot Nothing Then
                newRow("mainbusscode") = LinkCodeBusscode("InfoCode")
            Else
                newRow("mainbusscode") = DBNull.Value
            End If

            If CustDT.Rows(i).Item("regtype") = 0 Then
                newRow("regtype") = "new"
            Else
                newRow("regtype") = "amc"
            End If

            If CustDT.Rows(i).Item("lan") = 0 Then
                newRow("Lan") = "N"
            Else
                newRow("Lan") = "Y"
            End If

            If CustDT.Rows(i).Item("svol") = 0 Then
                newRow("ProductCode") = 906
            ElseIf CustDT.Rows(i).Item("svol") = 1 Then
                newRow("ProductCode") = 908
            ElseIf CustDT.Rows(i).Item("svol") = 2 Then
                newRow("ProductCode") = 907
            End If
            Dim knod As Integer = CInt(CustDT.Rows(i).Item("nodes").ToString) - 1
            newRow("Nodes") = IIf(knod < 0, 0, knod)
            newRow("ChangeNodes") = IIf(knod < 0, 0, knod)
            dt.Rows.Add(newRow)
        Next
        'Set Allowupto date if pcode is coming from EPL
        Dim dt1 As New DataTable
        For qw = 0 To dt.Rows.Count - 1
            If IsDBNull(dt.Rows(qw).Item("CustCode")) = False And dt.Rows(qw).Item("CustCode") IsNot "" Then
                dt1 = CalculateAllowUptodate(dt.Rows(qw).Item("CustCode"), dt.Rows(qw).Item("regtype"), "", dealerRow)
                dt.Rows(qw).Item("Customers_Key") = dt1.Rows(0).Item("Customers_Key")
                dt.Rows(qw).Item("P_Customers") = dt1.Rows(0).Item("P_Customers")
                dt.Rows(qw).Item("CustNameFromDb") = dt1.Rows(0).Item("CustName")
                dt.Rows(qw).Item("CurrRegDate") = dt1.Rows(0).Item("CurrRegDate")
                dt.Rows(qw).Item("CalculatedAllowUpto") = dt1.Rows(0).Item("calculatedallowupto")
                dt.Rows(qw).Item("ChangeAllowUpto") = dt1.Rows(0).Item("calculatedallowupto")
                Dim temp As Date = dt.Rows(qw).Item("ChangeAllowUpto")
                dt.Rows(qw).Item("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")

            ElseIf dt.Rows(qw).Item("regtype") = "new" Then
                dt1 = CalculateAllowUptodate("", dt.Rows(qw).Item("regtype"), "", dealerRow)
                dt.Rows(qw).Item("CurrRegDate") = dt1.Rows(0).Item("CurrRegDate")
                dt.Rows(qw).Item("CalculatedAllowUpto") = dt1.Rows(0).Item("calculatedallowupto")
                dt.Rows(qw).Item("ChangeAllowUpto") = dt1.Rows(0).Item("calculatedallowupto")
                Dim temp As Date = dt1.Rows(0).Item("calculatedallowupto")
                dt.Rows(qw).Item("ChangeAllowUptoStr") = temp.ToString("yyyy-MM-dd")
            End If
        Next
        Return dt
    End Function
    ''' <summary>
    ''' Function Calculates AllowUpto date when Customer registration is being opened.
    ''' </summary>
    ''' <param name="CUSTCODE">Custcode from customer epl file/Customer table.</param>
    ''' <param name="regtype">Regtype of Customer i.e amc or new</param>
    ''' <param name="regtype2">regtype2 of customer i.e main or home</param>
    ''' <param name="DealerRow">linkcode of Current Login user from Websessions table.</param>
    ''' <returns></returns>
    Public Function CalculateAllowUptodate(CUSTCODE As String, regtype As String, regtype2 As String, DealerRow As DataRow) As DataTable
        Dim CustomerRow As New DataTable
        Dim TempCustomerRow As New DataTable
        Dim ClsCustomers As New Customers.Customers.Customers
        Dim s As New DataTable
        CustomerRow = df1.AddColumnsInDataTable(CustomerRow, "Customers_Key, P_Customers, CurrRegDate, calculatedallowupto, ChangeAllowUpto, ChangeAllowUptoStr,Custtype,CustNameFromDb,lan,nodes,IsAssigned,IsCustomer", "System.Int32, System.Int32, System.DateTime, System.DateTime, System.DateTime, System.String,System.String, System.String, System.String, System.Int32,System.String,System.String")
        Dim CurrentDate As Date = df1.getDateTimeISTNow()

        If CUSTCODE = "" Then
            Dim CustRow As DataRow = CustomerRow.NewRow()
            CustRow("CurrRegDate") = CurrentDate
            CustRow("ChangeAllowUpto") = CurrentDate.AddDays(365)
            CustRow("calculatedallowupto") = CurrentDate.AddDays(365)
            Dim temp As Date = CustRow("ChangeAllowUpto")
            CustRow("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")
            CustRow("CustNameFromDb") = ""
            CustRow("IsAssigned") = ""
            CustRow("IsCustomer") = ""
            CustomerRow.Rows.Add(CustRow)
        Else
            s = df1.AddColumnsInDataTable(s, "Customers_Key, P_Customers, CurrRegDate, calculatedallowupto, ChangeAllowUpto, ChangeAllowUptoStr,Custtype,CustNameFromDb,lan,nodes,IsAssigned,IsCustomer,ActiveF,Mobno", "System.Int32, System.Int32, System.DateTime, System.DateTime, System.DateTime, System.String,System.String, System.String, System.String, System.Int32,System.String,System.String,System.String,system.string")
            Dim TempRow As DataRow = s.NewRow()

            TempCustomerRow = df1.GetDataFromSql(ClsCustomers.ServerDatabase, ClsCustomers.TableName, "*", "", "custcode = '" & CUSTCODE & "' and rowstatus = 0", "", "")
            CustomerRow = TempCustomerRow.Clone
            If TempCustomerRow.Rows.Count > 0 Then
                ' If df1.GetCellValue(TempCustomerRow.Rows(0), "servicingagentcode") = df1.GetCellValue(DealerRow, "P_dealers") Then
                '   CustomerRow = df1.GetDataFromSql(ClsCustomers.ServerDatabase, ClsCustomers.TableName, "*", "", "custcode = '" & CUSTCODE & "' and rowstatus = 0 and servicingagentcode = " & df1.GetCellValue(DealerRow, "P_dealers"), "", "")
                'CustomerRow.Rows.Add()

                'CustomerRow.Rows(0) = df1.UpdateDataRows(CustomerRow.Rows(0), TempCustomerRow.Rows(0))
                CustomerRow = TempCustomerRow
                CustomerRow = df1.AddColumnsInDataTable(CustomerRow, "IsAssigned,IsCustomer", "System.String,System.String")

                If df1.GetCellValue(TempCustomerRow.Rows(0), "servicingagentcode") = df1.GetCellValue(DealerRow, "P_acccode") Then
                    CustomerRow.Rows(0).Item("IsCustomer") = "true"
                    CustomerRow.Rows(0).Item("IsAssigned") = "true"
                Else
                    TempRow("IsCustomer") = "true"
                    TempRow("IsAssigned") = "false"
                    s.Rows.Add(TempRow)
                    CustomerRow = s
                End If
            Else
                TempRow("CurrRegDate") = CurrentDate
                TempRow("ChangeAllowUpto") = CurrentDate.AddDays(365)
                TempRow("calculatedallowupto") = CurrentDate.AddDays(365)
                Dim temp As Date = TempRow("ChangeAllowUpto")
                TempRow("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")
                TempRow("CustNameFromDb") = ""
                TempRow("IsCustomer") = "false"
                TempRow("IsAssigned") = "false"
                s.Rows.Add(TempRow)
                CustomerRow = s
            End If


            'added billeduptoStr,LastOpened,LastOpenedStr,RegOpenedUpto,RegOpenedUptoStr
            CustomerRow = df1.AddColumnsInDataTable(CustomerRow, "CurrRegDate,calculatedallowupto,ChangeAllowUpto,ChangeAllowUptoStr,CustNameFromDb,lan,nodes,LastOpendt,LastOpendtStr,billedUptoStr,RegOpenedUpto,RegOpenedUptoStr", "system.datetime,system.datetime,system.datetime,system.string,system.string,system.string,system.int32,system.datetime,system.string,system.string,system.datetime,system.string")

            If CustomerRow.Rows.Count > 0 Then
                If CustomerRow.Rows(0).Item("isassigned") = "true" Then
                    Dim Nodes As Integer = GetLanandNodeFromChargingItems(CustomerRow(0).Item("P_customers"))
                    If Nodes = 0 Then
                        CustomerRow.Rows(0).Item("Lan") = "N"
                    Else
                        CustomerRow.Rows(0).Item("Lan") = "Y"
                    End If
                    CustomerRow.Rows(0).Item("Nodes") = Nodes

                    CustomerRow.Rows(0).Item("CustNameFromDb") = CustomerRow.Rows(0).Item("custname")
                    Dim BilledUptoValue As Date = "1990-01-01 00:00:00"
                    If Not DBNull.Value.Equals(CustomerRow.Rows(0).Item("BilledUpto")) Then
                        BilledUptoValue = CustomerRow.Rows(0).Item("BilledUpto")


                    End If



                    If regtype = "amc" And regtype2 = "main" Then

                        'For lastRegDate Calculation
                        Dim RegTranQuery As String = "select top 1 * from RegistrationTran where P_Customers=" & CustomerRow.Rows(0).Item("P_Customers") & " Order by RegSendDate desc"
                        Dim RegTrandt As DataTable = df1.SqlExecuteDataTable(ClsCustomers.ServerDatabase, RegTranQuery)
                        If RegTrandt IsNot Nothing Then
                            If RegTrandt.Rows.Count > 0 Then
                                If IsDBNull(RegTrandt.Rows(0).Item("RegSendDate")) = False Then
                                    Dim LastOpen As Date = RegTrandt.Rows(0).Item("RegSendDate")
                                    CustomerRow.Rows(0).Item("LastOpendt") = LastOpen
                                    CustomerRow.Rows(0).Item("LastOpendtStr") = LastOpen.ToString("dd/MM/yyyy")
                                    Dim RegOpenUpTo As Date = RegTrandt.Rows(0).Item("OpenedUpto")
                                    CustomerRow.Rows(0).Item("RegOpenedUpto") = RegOpenUpTo
                                    CustomerRow.Rows(0).Item("RegOpenedUptoStr") = RegOpenUpTo.ToString("dd/MM/yyyy")
                                End If
                            End If
                        End If


                        If Not IsDBNull(CustomerRow.Rows(0).Item("amcmonth")) Then
                            Dim amcmonth As String = CustomerRow.Rows(0).Item("amcmonth").ToString
                            ' Dim mday As Integer = CInt(Strings.(amcmonth,))


                            Dim mday As Integer = CInt(amcmonth.Substring(1, 2))
                            Dim mmonth As Integer = CInt(Right(amcmonth, 2))
                            Dim currdate As Date = df1.getDateTimeISTNow
                            Dim ryear As Integer = currdate.Year
                            Dim lyear As Integer = IIf(mmonth <= currdate.Month + 1, ryear + IIf(mmonth = 1 And currdate.Month = 12, 2, 1), ryear)
                            Dim adt As New DateTime(lyear, mmonth, mday)

                            CustomerRow.Rows(0).Item("calculatedallowupto") = adt
                            CustomerRow.Rows(0).Item("ChangeAllowUpto") = adt
                            Dim temp As Date = CustomerRow.Rows(0).Item("ChangeAllowUpto")
                            CustomerRow.Rows(0).Item("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")
                            CustomerRow.Rows(0).Item("CurrRegDate") = CurrentDate

                            CustomerRow.Rows(0).Item("billedUptoStr") = BilledUptoValue.ToString("dd/MM/yyyy")
                        End If

                    ElseIf regtype = "new" And regtype2 = "main" Then
                        CustomerRow.Rows(0).Item("CurrRegDate") = CurrentDate
                        CustomerRow.Rows(0).Item("ChangeAllowUpto") = CurrentDate.AddDays(365)
                        CustomerRow.Rows(0).Item("calculatedallowupto") = CurrentDate.AddDays(365)
                        Dim temp As Date = CustomerRow.Rows(0).Item("ChangeAllowUpto")
                        CustomerRow.Rows(0).Item("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")

                    ElseIf regtype = "new" And regtype2 = "home" Then
                        CustomerRow.Rows(0).Item("CurrRegDate") = CurrentDate
                        CustomerRow.Rows(0).Item("ChangeAllowUpto") = CurrentDate.AddDays(365)
                        CustomerRow.Rows(0).Item("calculatedallowupto") = CurrentDate.AddDays(365)
                        Dim temp As Date = CustomerRow.Rows(0).Item("ChangeAllowUpto")
                        CustomerRow.Rows(0).Item("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")

                    ElseIf regtype = "amc" And regtype2 = "home" Then
                        Dim mbilledupto As DateTime = libcustomerfeature.GetBilledUpToDate(CustomerRow.Rows(0).Item("P_Customers"), regtype2)

                        Dim mday As Integer = mbilledupto.Day    'CInt(amcmonth.Substring(1, 2))
                        Dim mmonth As Integer = mbilledupto.Month  ' CInt(Right(amcmonth, 2))
                        Dim currdate As Date = df1.getDateTimeISTNow
                        Dim ryear As Integer = currdate.Year
                        Dim lyear As Integer = IIf(mmonth <= currdate.Month + 1, ryear + IIf(mmonth = 1 And currdate.Month = 12, 2, 1), ryear)
                        Dim adt As New DateTime(lyear, mmonth, mday)




                        CustomerRow.Rows(0).Item("calculatedallowupto") = adt
                        CustomerRow.Rows(0).Item("ChangeAllowUpto") = adt
                        Dim temp As Date = CustomerRow.Rows(0).Item("ChangeAllowUpto")
                        CustomerRow.Rows(0).Item("ChangeAllowUptostr") = temp.ToString("yyyy-MM-dd")
                        CustomerRow.Rows(0).Item("CurrRegDate") = CurrentDate

                        'CustomerRow.Rows(0).Item("billedUptoStr") = BilledUptoValue.ToString("dd/MM/yyyy")



                    End If

                    'End If
                End If
            End If
        End If
        Return CustomerRow
    End Function
    'Public Function GetBilledUpToDate(ByVal P_Customer As Integer, regtype As String) As DateTime
    '    Dim clschargingheader As New ChargingHeader.ChargingHeader.ChargingHeader
    '    Dim clschargingItems As New ChargingItems.ChargingItems.ChargingItems

    '    Dim Amcdate As DateTime
    '    Dim chdt As New DataTable
    '    Dim chidt As New DataTable
    '    Dim headerno As Integer = 0
    '    Dim BuildUpTo As DateTime
    '    chdt = df1.GetDataFromSql(clschargingheader.ServerDatabase, clschargingheader.TableName, "billedupto,headerno", "", "P_Customers = " & P_Customer & " and PaymentFlag = 'P' and rowstatus = 0", "", "BilledUpto DESC")
    '    For i = 0 To chdt.Rows.Count - 1
    '        headerno = df1.GetCellValue(chdt.Rows(i), "HeaderNo")
    '        BuildUpTo = df1.GetCellValue(chdt.Rows(i), "BilledUpto")
    '        Dim entered As Boolean = False
    '        chidt = df1.GetDataFromSql(clschargingItems.ServerDatabase, clschargingItems.TableName, "servicecode,chargingtodate", "", "P_Customers = " & P_Customer & " and HeaderNo = " & headerno & " and rowstatus = 0", "", "")
    '        For j = 0 To chidt.Rows.Count - 1
    '            Dim servicecode As Integer = df1.GetCellValue(chidt.Rows(j), "ServiceCode")
    '            If regtype = "main" Then

    '                If servicecode = 2412 Or servicecode = 2401 Or servicecode = 3019 Then
    '                    entered = True
    '                    Amcdate = BuildUpTo
    '                    Exit For
    '                End If
    '            ElseIf regtype = "home" Then

    '                If servicecode = 2403 Or servicecode = 2774 Then
    '                    entered = True
    '                    Amcdate = df1.GetCellValue(chidt.Rows(j), "chargingtodate")
    '                    Exit For
    '                End If

    '            End If
    '        Next
    '        If entered = True Then Exit For
    '    Next
    '    Return Amcdate
    'End Function
    ''' <summary>
    ''' Function gets the values of Lan and nodes from ChargingItems Table.
    ''' </summary>
    ''' <param name="P_Customers"></param>
    ''' <returns>Lan and nodes value</returns>
    Public Function GetLanandNodeFromChargingItems(P_Customers As Integer) As Integer

        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtChargingheader As DataTable = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "paymentflag = 'P' and rowstatus=0 and p_customers = " & P_Customers, "", "")

        Dim clschargingItem As New ChargingItems.ChargingItems.ChargingItems
        Dim P_chargingHeader As String = ""
        For t = 0 To dtChargingheader.Rows.Count - 1
            P_chargingHeader = P_chargingHeader & "," & df1.GetCellValue(dtChargingheader.Rows(t), "headerno")
        Next
        Dim qty As Integer = 0
        If P_chargingHeader.StartsWith(",") Then P_chargingHeader = P_chargingHeader.Substring(1)
        If Not P_chargingHeader = "" Then
            Dim dtchargingItems As DataTable = df1.GetDataFromSql(clschargingItem.ServerDatabase, clschargingItem.TableName, "*", "", "rowstatus = 0 and headerno in ( " & P_chargingHeader & ")", "", "")
            qty = 0
            For kl = 0 To dtchargingItems.Rows.Count - 1
                Dim mServiceCode As Integer = df1.GetCellValue(dtchargingItems.Rows(kl), "servicecode")
                If mServiceCode = 2402 Then
                    Dim quantity As Decimal = df1.GetCellValue(dtchargingItems.Rows(kl), "quantity")
                    qty = qty + quantity
                End If
            Next
        End If
        Return qty
    End Function
    Public Function populatecustomerMismatch(ByVal customers As DataTable, ByVal customersMachData As DataTable, ByVal customerRowDb As DataTable)

        Dim dtcustinfomismatch As New DataTable
        dtcustinfomismatch = df1.AddColumnsInDataTable(dtcustinfomismatch, "p_customers,custname,mobno,email,postaladdress1,postaladdress2,postaladdress3,postaladdress4,pincode,hometown,mainbusscode,gstin,machinemismatch,nodesmismatch")

        For k = 0 To customers.Rows.Count - 1
            If customers.Rows(k).Item("isvalid") = "false" Then Continue For
            If customers.Rows(k).Item("chargingtrue") = "false" And customers.Rows(k).Item("paymentflag") = True Then Continue For
            Dim newCust As Boolean = False
            '  End If
            Dim p_customers As Integer = df1.GetCellValue(customers.Rows(k), "p_customers")
            Dim dtcust As New DataTable
            Dim dtrow As DataRow = dtcustinfomismatch.NewRow
            If Not customerRowDb Is Nothing Then dtcust = customerRowDb '      Session("customerRowDb")

            If dtcust.Rows.Count > 0 Then

                Dim dtcustarr() As DataRow = dtcust.Select("P_customers=" & p_customers) '.CopyToDataTable
                If dtcustarr.Count > 0 Then
                    If df1.GetCellValue(dtcustarr(0), "customers_key", "integer") <> 0 Then
                        Dim custname As String = df1.GetCellValue(dtcustarr(0), "custname", "string")
                        Dim mobno As String = df1.GetCellValue(dtcustarr(0), "mobno", "string")
                        Dim email As String = df1.GetCellValue(dtcustarr(0), "email", "string")
                        Dim postaladdress1 As String = df1.GetCellValue(dtcustarr(0), "postaladdress1", "string")
                        Dim postaladdress2 As String = df1.GetCellValue(dtcustarr(0), "postaladdress2", "string")
                        Dim postaladdress3 As String = df1.GetCellValue(dtcustarr(0), "postaladdress3", "string")
                        Dim postaladdress4 As String = df1.GetCellValue(dtcustarr(0), "postaladdress4", "string")
                        Dim pincode As String = df1.GetCellValue(dtcustarr(0), "pincode", "string")
                        Dim hometown As Integer = df1.GetCellValue(dtcustarr(0), "hometown")
                        Dim mainbusscode As Integer = df1.GetCellValue(dtcustarr(0), "mainbusscode")
                        Dim gstin As String = df1.GetCellValue(dtcustarr(0), "gstin")
                        '  Dim eplhash As New Hashtable

                        Dim custnameepl As String = df1.GetCellValue(customers.Rows(k), "custname", "string")
                        Dim mobnoepl As String = df1.GetCellValue(customers.Rows(k), "mobno", "string")
                        Dim emailepl As String = df1.GetCellValue(customers.Rows(k), "email", "string")
                        Dim postaladdress1epl As String = df1.GetCellValue(customers.Rows(k), "postaladdress1", "string")
                        Dim postaladdress2epl As String = df1.GetCellValue(customers.Rows(k), "postaladdress2", "string")
                        Dim postaladdress3epl As String = df1.GetCellValue(customers.Rows(k), "postaladdress3", "string")
                        Dim postaladdress4epl As String = df1.GetCellValue(customers.Rows(k), "postaladdress4", "string")
                        Dim hometownepl As Integer = df1.GetCellValue(customers.Rows(k), "hometown")
                        Dim mainbusscodeepl As Integer = df1.GetCellValue(customers.Rows(k), "mainbusscode")
                        Dim gstinepl As String = df1.GetCellValue(customers.Rows(k), "gstin")

                        dtrow.Item("p_customers") = p_customers
                        If LCase(custname.Trim) = LCase(custnameepl.Trim) Then
                            dtrow.Item("custname") = "N"
                        Else
                            dtrow.Item("custname") = "Y"
                        End If

                        If LCase(mobno.Trim) = LCase(mobnoepl.Trim) Then
                            dtrow.Item("mobno") = "N"
                        Else
                            dtrow.Item("mobno") = "Y"
                        End If

                        If LCase(email.Trim) = LCase(emailepl.Trim) Then
                            dtrow.Item("email") = "N"
                        Else
                            dtrow.Item("email") = "Y"
                        End If


                        If LCase(postaladdress1.Trim) = LCase(postaladdress1epl.Trim) Then
                            dtrow.Item("postaladdress1") = "N"
                        Else
                            dtrow.Item("postaladdress1") = "Y"
                        End If

                        If LCase(postaladdress2.Trim) = LCase(postaladdress2epl.Trim) Then
                            dtrow.Item("postaladdress2") = "N"
                        Else
                            dtrow.Item("postaladdress2") = "Y"
                        End If


                        If LCase(postaladdress3.Trim) = LCase(postaladdress3epl.Trim) Then
                            dtrow.Item("postaladdress3") = "N"
                        Else
                            dtrow.Item("postaladdress3") = "Y"
                        End If

                        If LCase(postaladdress4.Trim) = LCase(postaladdress4epl.Trim) Then
                            dtrow.Item("postaladdress4") = "N"
                        Else
                            dtrow.Item("postaladdress4") = "Y"
                        End If

                        If hometown = hometownepl Then
                            dtrow.Item("hometown") = "N"
                        Else
                            dtrow.Item("hometown") = "Y"
                        End If
                        If mainbusscode = mainbusscodeepl Then
                            dtrow.Item("mainbusscode") = "N"
                        Else
                            dtrow.Item("mainbusscode") = "Y"
                        End If
                        If gstin = gstinepl Then
                            dtrow.Item("gstin") = "N"
                        Else
                            dtrow.Item("gstin") = "Y"
                        End If
                        Dim prevMachineStr As String = ""
                        Dim currMachineStr As String = ""
                        Dim dtcustmacharr() As DataRow = customersMachData.Select("p_customers ='" & CStr(p_customers) & "'")
                        If dtcustmacharr.Count = 2 Then
                            prevMachineStr = df1.GetCellValue(dtcustmacharr(0), "baseboard") & df1.GetCellValue(dtcustmacharr(0), "cpu") & df1.GetCellValue(dtcustmacharr(0), "bios") & df1.GetCellValue(dtcustmacharr(0), "processor")
                            currMachineStr = df1.GetCellValue(dtcustmacharr(1), "baseboard") & df1.GetCellValue(dtcustmacharr(1), "cpu") & df1.GetCellValue(dtcustmacharr(1), "bios") & df1.GetCellValue(dtcustmacharr(1), "processor")
                        ElseIf dtcustmacharr.Count = 1 Then
                            currMachineStr = df1.GetCellValue(dtcustmacharr(0), "baseboard") & df1.GetCellValue(dtcustmacharr(0), "cpu") & df1.GetCellValue(dtcustmacharr(0), "bios") & df1.GetCellValue(dtcustmacharr(0), "processor")
                        End If
                        If prevMachineStr = currMachineStr Then
                            dtrow.Item("machinemismatch") = "N"
                        Else
                            dtrow.Item("machinemismatch") = "Y"
                        End If
                        If customers.Rows(k).Item("changenodes") <> customers.Rows(k).Item("nodesDb") Then
                            dtrow.Item("nodesmismatch") = "Y"
                        Else
                            dtrow.Item("nodesmismatch") = "N"
                        End If
                    Else
                        newCust = True
                    End If
                Else
                    newCust = True
                End If
            Else
                newCust = True
            End If
            If newCust = True Then
                dtrow.Item("p_customers") = p_customers
                dtrow.Item("custname") = "Y"
                dtrow.Item("mobno") = "Y"
                dtrow.Item("email") = "Y"
                dtrow.Item("postaladdress1") = "Y"
                dtrow.Item("postaladdress2") = "Y"
                dtrow.Item("postaladdress3") = "Y"
                dtrow.Item("postaladdress4") = "Y"
                dtrow.Item("hometown") = "Y"
                dtrow.Item("mainbusscode") = "Y"
                dtrow.Item("gstin") = "Y"
                dtrow.Item("machinemismatch") = "N"
                dtrow.Item("nodesmismatch") = "N"
                dtcustinfomismatch.Rows.Add(dtrow)
            Else
                dtcustinfomismatch.Rows.Add(dtrow)
            End If

        Next

        Return dtcustinfomismatch
    End Function


    ''' <summary>
    ''' This function stores the relevant data into tables(Customers,CustomerServices,Registration,RegistrationTrans and websessions) in database after CustomerRegistration is opened.
    ''' </summary>
    ''' <param name="dt1">datatable with data to be used while saving in tables in database.</param>
    ''' <param name="sessionRow">Session Row contain details of the current login user.</param>
    ''' <param name="dealerRow">Session Row which contain details of main dealer as per loggedin user</param>
    ''' <returns></returns>
    Public Function ProcessCustomersforRegistrationPreOrder(ByVal dt1 As DataTable, sessionRow As DataRow, dealerRow As DataRow, custmachData As DataTable, newbutold As String) As DataTable
        Dim LoginType As String = ""

        Dim sessionkey As Integer = 0
        Dim Loginkey As Integer = 0
        Dim clswebsession As New WebSessions.WebSessions.WebSessions
        If sessionRow IsNot Nothing Then
            clswebsession.PrevRow = df1.UpdateDataRows(clswebsession.PrevRow, sessionRow)
            LoginType = sessionRow("LinkType")
            Loginkey = sessionRow("LinkCode")
            sessionkey = sessionRow("websessions_key")
        End If
        Dim DealerloginKey As Integer = -1
        If dealerRow IsNot Nothing Then
            DealerloginKey = df1.GetCellValue(dealerRow, "p_acccode")
        End If
        For i = 0 To dt1.Rows.Count - 1

            If Not UCase(dt1.Rows(i).Item("isvalid")) = UCase("true") Then Continue For
            If LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "new" And UCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = UCase("main") Then
                If newbutold = True Then Continue For
                Dim ClsCustomers As New Customers.Customers.Customers
                Dim ClsCustomerServices As New CustomerServices.CustomerServices.CustomerServices
                Dim clsRegistrations As New Registrations.Registrations.Registrations



                Dim dtro As DataRow = dt1.Rows(i)
                Dim dthash As Hashtable = GF1.CreateHashTable(dtro)
                dthash = GF1.AddItemToHashTable(dthash, "accname", dt1.Rows(i).Item("custname"))
                dthash = GF1.AddItemToHashTable(dthash, "acctype", 3041)
                dthash = GF1.AddItemToHashTable(dthash, "mobile", dt1.Rows(i).Item("mobno"))

                Dim libsaralauth As New SaralAuthLib.LoginFunctions
                Dim p_acccode As Integer = libsaralauth.InsertUpdateInAcc_Master(-1, dthash, sessionRow)

                Dim dthashUserLogin As New Hashtable
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "linktype", "C")
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "linkcode", p_acccode)
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "userid", dt1.Rows(i).Item("CustCode"))
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "pwd", dt1.Rows(i).Item("CustCode"))
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "email", dt1.Rows(i).Item("Email"))
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "address", df1.GetCellValue(dt1.Rows(i), "postaladdress1", "string") & " " & df1.GetCellValue(dt1.Rows(i), "postaladdress2", "string") & " " & df1.GetCellValue(dt1.Rows(i), "postaladdress3", "string") & " " & df1.GetCellValue(dt1.Rows(i), "postaladdress4", "string"))
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "name", dt1.Rows(i).Item("custname"))
                dthashUserLogin = GF1.AddItemToHashTable(dthashUserLogin, "mobile", dt1.Rows(i).Item("mobno"))

                Dim userlogin_key As Integer = libsaralauth.InsertUpdateUserLogin(-1, dthashUserLogin)

                Dim dthashUserRoles As New Hashtable
                dthashUserRoles = GF1.AddItemToHashTable(dthashUserRoles, "userlogin_key", userlogin_key)

                dthashUserRoles = GF1.AddItemToHashTable(dthashUserRoles, "approles", "46,47,48,49")
                libsaralauth.InsertUpdateUserRoles(-1, dthashUserRoles)

                ClsCustomers.CurrRow.Item("p_acccode") = p_acccode

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CustCode")) Then
                    ClsCustomers.CurrRow("CustCode") = dt1.Rows(i).Item("CustCode")
                End If
                Dim bln As Boolean = checkForDuplicateCustCode(dt1.Rows(i).Item("CustCode"))
                If bln = True Then Continue For

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CustName")) Then
                    ClsCustomers.CurrRow("CustName") = dt1.Rows(i).Item("CustName").ToString.Trim
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("MobNo")) Then
                    Dim mobNo As String = dt1.Rows(i).Item("MobNo")
                    If String.IsNullOrEmpty(mobNo) = False Then
                        If mobNo.Length <= 10 Then
                            ClsCustomers.CurrRow("MobNo") = dt1.Rows(i).Item("MobNo")
                        Else
                            ClsCustomers.CurrRow("Phone") = dt1.Rows(i).Item("MobNo")
                        End If
                    End If

                End If
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("HomeTown")) Then
                    ClsCustomers.CurrRow("HomeTown") = dt1.Rows(i).Item("HomeTown")
                End If
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("ProductCode")) Then
                    ClsCustomers.CurrRow("ProductCode") = dt1.Rows(i).Item("ProductCode")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("PostalAddress1")) Then
                    ClsCustomers.CurrRow("PostalAddress1") = dt1.Rows(i).Item("PostalAddress1")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("PostalAddress2")) Then
                    ClsCustomers.CurrRow("PostalAddress2") = dt1.Rows(i).Item("PostalAddress2")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("PostalAddress3")) Then
                    ClsCustomers.CurrRow("PostalAddress3") = dt1.Rows(i).Item("PostalAddress3")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("PostalAddress4")) Then
                    ClsCustomers.CurrRow("PostalAddress4") = dt1.Rows(i).Item("PostalAddress4")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("Pincode")) Then
                    ClsCustomers.CurrRow("Pincode") = dt1.Rows(i).Item("Pincode")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("Email")) Then
                    ClsCustomers.CurrRow("Email") = dt1.Rows(i).Item("Email")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("MainBussCode")) Then
                    ClsCustomers.CurrRow("MainBussCode") = dt1.Rows(i).Item("MainBussCode")
                End If

                ClsCustomers.CurrRow("SaleAgent") = "D"
                ClsCustomers.CurrRow("SaleAgentCode") = DealerloginKey 'Session("Key")
                ClsCustomers.CurrRow("TrainingAgent") = "D"  'LoginType
                ClsCustomers.CurrRow("TrainingAgentCode") = DealerloginKey 'Session("Key")
                ClsCustomers.CurrRow("ServicingAgent") = "D"
                ClsCustomers.CurrRow("ServicingAgentCode") = DealerloginKey 'Session("Key")
                ClsCustomers.CurrRow("verified") = "N"
                ClsCustomers.CurrRow("customerstatus") = "P"
                'If Not DBNull.Value.Equals(dt1.Rows(i).Item("BilledUpto")) Then
                '    ClsCustomers.CurrRow("BilledUpto") = dt1.Rows(i).Item("BilledUpto")
                'End If
                'Copied from regtype = amc main case 'ceilling date is same in that as well
                'AskFrom Mam
                'Nisha
                'If Not DBNull.Value.Equals(dt1.Rows(i).Item("CalculatedAllowUpto")) Then
                '    ClsCustomers.CurrRow("BilledUpto") = dt1.Rows(i).Item("CalculatedAllowUpto")
                'End If
                'End Nisha
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    ClsCustomers.CurrRow("CurrRegDate") = dt1.Rows(i).Item("CurrRegDate")
                    ClsCustomers.CurrRow("FirstInstallDate") = dt1.Rows(i).Item("CurrRegDate")
                End If

                If DBNull.Value.Equals(dt1.Rows(i).Item("changeallowupto")) Then
                    ClsCustomers.CurrRow("AllowUpto") = "1990-01-01 00:00:00.000"
                Else
                    ClsCustomers.CurrRow("AllowUpto") = dt1.Rows(i).Item("changeallowupto")
                End If

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CalculatedAllowUpto")) Then
                    ClsCustomers.CurrRow("Ceilingdate") = dt1.Rows(i).Item("CalculatedAllowUpto")
                End If
                'changes by neha
                ClsCustomers.CurrRow("RecurrChgDueAfter") = 365
                ' Dim mamcmnth As String = "M" & CDate(ClsCustomers.CurrRow("FirstInstallDate")).Day.ToString.PadLeft(2, "0") & CDate(ClsCustomers.CurrRow("FirstInstallDate")).Month.ToString.PadLeft(2, "0")
                Dim mamcmnth As String = "M" & CDate(ClsCustomers.CurrRow("FirstInstallDate")).Day.ToString.PadLeft(2, "0") & CDate(ClsCustomers.CurrRow("FirstInstallDate")).Month.ToString.PadLeft(2, "0")

                If DealerloginKey = 3 Then
                    mamcmnth = "M" & "25" & CDate(ClsCustomers.CurrRow("FirstInstallDate")).Month.ToString.PadLeft(2, "0")
                End If
                ClsCustomers.CurrRow("amcmonth") = mamcmnth
                ' for custromer service
                Dim dt2 As New DataTable
                dt2 = ClsCustomerServices.CurrDt

                ClsCustomerServices.PrevDt = df1.UpdateDataTables(ClsCustomerServices.PrevDt, dt2)  'add by Shweta

                Dim a As DataRow = dt2.NewRow
                a("CustomerServices_key") = -1
                a("ServiceInstalledBy") = "D" 'LoginType
                a("ServiceInstallingAgent") = DealerloginKey 'Session("Key")
                a("ServiceCode") = 2401
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    a("ServiceStartDate") = dt1.Rows(i).Item("CurrRegDate")
                End If
                a("Quantity") = 1
                a("RecurrServiceCode") = 2412
                a("RecurrChgDueAfter") = 365
                a("Verified") = "N"
                a("serviceno") = 1
                dt2.Rows.Add(a)

                If dt1.Rows(i).Item("changenodes") > 0 Then
                    Dim b As DataRow = dt2.NewRow
                    b("CustomerServices_key") = -2
                    b("ServiceInstalledBy") = "D" 'LoginType
                    b("ServiceInstallingAgent") = DealerloginKey
                    b("ServiceCode") = 2402
                    If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                        b("ServiceStartDate") = dt1.Rows(i).Item("CurrRegDate")
                    End If
                    b("Quantity") = dt1.Rows(i).Item("changenodes")
                    b("RecurrServiceCode") = 2773
                    b("RecurrChgDueAfter") = 365
                    b("Verified") = "N"
                    b("serviceno") = 2
                    dt2.Rows.Add(b)
                End If

                ClsCustomerServices.CurrDt = dt2

                Dim HashPublicValues As New Hashtable
                Dim aClsObject() As Object = {ClsCustomers, ClsCustomerServices, clswebsession}
                '       Dim success As Boolean = SaveIntodb(aClsObject, "D")
                Dim mserverdb As String = df1.GetServerMDFForTransanction(aClsObject)
                Dim mytrans As SqlTransaction = df1.BeginTransaction(mserverdb)
                Dim aLastKeysValues As New Hashtable
                aClsObject = df1.SetKeyValueIfNewInsert(mytrans, aClsObject)
                Dim sqlexec As Boolean = df1.CheckTableClassUpdations(aClsObject)

                aClsObject = df1.LastKeysPlus(mytrans, aClsObject, aLastKeysValues)
                HashPublicValues = GF1.AddItemToHashTable(HashPublicValues, "mTypeCode", "D")
                aClsObject = df1.SetFinalFieldsValues(aClsObject, HashPublicValues)
                Dim P_customers As Integer = -1, customers_key As Integer = -1
                Dim aam As Integer = 0
                Try
                    If sqlexec = True Then
                        aam = df1.InsertUpdateDeleteSqlTables(mytrans, aClsObject, aam)

                        Dim CustomerHash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValues, "Customers")
                        P_customers = GF1.GetValueFromHashTable(CustomerHash, "p_customers")
                        customers_key = GF1.GetValueFromHashTable(CustomerHash, "customers_key")
                        dt1.Rows(i).Item("p_customers") = P_customers
                        dt1.Rows(i).Item("customers_key") = customers_key


                        Dim cuslib As New CustomerFeatureLib.CustomerFeatureFunctions
                        cuslib.SendOTP(dt1, sessionRow)

                        mytrans.Commit()
                    End If
                Catch ex As Exception
                    mytrans.Rollback()
                End Try

                'added by Shweta
            End If
        Next

        Return dt1
    End Function
    Public Function checkForDuplicateCustCode(ByVal custcode As String) As Boolean
        Dim ClsCustomers As New Customers.Customers.Customers
        Dim bln As Boolean = False
        Dim custstrSql As String = "select count(*) from customers where custcode = '" & custcode & "' And rowstatus = 0"
        Dim dt As DataTable = df1.SqlExecuteDataTable(ClsCustomers.ServerDatabase, custstrSql)
        If dt.Rows.Count > 0 Then
            Dim val As Integer = dt.Rows(0).Item(0)
            If val > 0 Then
                bln = True
                Return bln
            End If
        Else
            bln = True
        End If

        Return bln
    End Function
    Public Sub addrowIncustomerVerification(ByVal dt1 As DataTable, ByVal dtcustomerMAch As DataTable, ByVal sessionrow As DataRow, ByVal abctran As Hashtable, ByVal CustomerInfoMismatch As DataTable)
        Dim custverifyHAshTable As New Hashtable
        Dim logintype As String = ""
        Dim loginkey As Integer = 0
        Dim sessionkey As Integer = 0
        If sessionrow IsNot Nothing Then
            logintype = sessionrow("linktype")
            loginkey = sessionrow("linkcode")
            sessionkey = sessionrow("websessions_key")
        End If
        For l = 0 To dt1.Rows.Count - 1
            If Not UCase(dt1.Rows(l).Item("isvalid")) = UCase("true") Then Continue For
            If dt1.Rows(l).Item("chargingtrue") = "false" And dt1.Rows(l).Item("paymentflag") = True Then Continue For

            Dim custinfo() As DataRow = CustomerInfoMismatch.Select("P_customers='" & CStr(dt1.Rows(l).Item("p_customers")) & "'")

            If custinfo.Count > 0 Then

                Dim createNewRow As Boolean = False

                Dim custname As String = custinfo(0).Item("custname")
                Dim machinemismatch As String = custinfo(0).Item("machinemismatch")

                Dim mobno As String = custinfo(0).Item("mobno")
                Dim email As String = custinfo(0).Item("email")

                Dim postaladdress1 As String = custinfo(0).Item("postaladdress1")
                Dim postaladdress2 As String = custinfo(0).Item("postaladdress2")

                Dim postaladdress3 As String = custinfo(0).Item("postaladdress3")
                Dim postaladdress4 As String = custinfo(0).Item("postaladdress4")

                Dim hometown As String = custinfo(0).Item("hometown")
                Dim mainbusscode As String = custinfo(0).Item("mainbusscode")

                Dim gstin As String = custinfo(0).Item("gstin")
                Dim nodesmismatch As String = custinfo(0).Item("nodesmismatch")

                If custname = "Y" Then createNewRow = True
                If mobno = "Y" Then createNewRow = True
                If machinemismatch = "Y" Then machinemismatch = True

                If email = "Y" Then createNewRow = True
                If postaladdress1 = "Y" Then createNewRow = True
                If postaladdress2 = "Y" Then createNewRow = True
                If postaladdress3 = "Y" Then createNewRow = True
                If postaladdress4 = "Y" Then createNewRow = True
                If hometown = "Y" Then createNewRow = True
                If mainbusscode = "Y" Then createNewRow = True
                If gstin = "Y" Then createNewRow = True
                If nodesmismatch = "Y" Then createNewRow = True


                If createNewRow = True Then
                    Dim clscustVerify As New CustomerVerification.CustomerVerification.CustomerVerification
                    Dim regtran_key As Integer = 0
                    Try
                        regtran_key = GF1.GetValueFromHashTable(abctran, CStr(dt1.Rows(l).Item("p_customers")))

                    Catch ex As Exception
                        ' Dim regtran_key As Integer
                    End Try
                    clscustVerify.CurrRow.Item("p_customers") = dt1.Rows(l).Item("p_customers")
                    clscustVerify.CurrRow.Item("logincode") = loginkey
                    clscustVerify.CurrRow.Item("logintype") = logintype
                    clscustVerify.CurrRow.Item("mtimestamp") = df1.getDateTimeISTNow
                    clscustVerify.CurrRow.Item("regtran_key") = regtran_key
                    clscustVerify.CurrRow.Item("status") = "P"
                    clscustVerify.CurrRow.Item("websessions_key") = sessionkey
                    Dim lcustverifyKey As Integer = 0
                    dt1.Rows(l).Item("custverifyflag") = "true"

                    Dim aClsObject() As Object = {clscustVerify}
                    Dim p_customerverification As Integer = cfc.SaveIntodbGetKey(aClsObject, "customerverification", "p_customerverification")
                    Dim clsCRMTasks As New CRMTasks.CRMTasks.CRMTasks
                    clsCRMTasks.CurrRow("Logintype") = sessionrow("linktype")
                    clsCRMTasks.CurrRow("Logincode") = sessionrow("linkcode")
                    clsCRMTasks.CurrRow("TaskTitle") = "Verify_" & dt1.Rows(l).Item("custname")
                    clsCRMTasks.CurrRow("Tasktype") = "S"
                    clsCRMTasks.CurrRow("TaskDescription") = "" 'RegCalls.IssueDescription
                    clsCRMTasks.CurrRow("Taskstatus") = 3008
                    clsCRMTasks.CurrRow("LinkCode") = p_customerverification
                    clsCRMTasks.CurrRow("LinkType") = "R"
                    clsCRMTasks.CurrRow("StartDate") = df1.getDateTimeISTNow()
                    clsCRMTasks.CurrRow("mtimestamp") = df1.getDateTimeISTNow()
                    clsCRMTasks.CurrRow("DueDate") = df1.getDateTimeISTNow().AddDays(5)
                    clsCRMTasks.CurrRow("websessions_key") = sessionkey
                    Dim aclsobj() As Object = {clsCRMTasks}

                    cfc.SaveIntodb(aclsobj)


                End If



                '    If LCase(dt1.Rows(l).Item("regtype").ToString.Trim) = "new" And UCase(dt1.Rows(l).Item("regtype2").ToString.Trim) = UCase("main") Then

            End If
        Next
        '  Return dt1
    End Sub

    ''' <summary>
    ''' This function stores the relevant data into tables(Customers,CustomerServices,Registration,RegistrationTrans and websessions) in database after CustomerRegistration is opened.
    ''' </summary>
    ''' <param name="dt1">datatable with data to be used while saving in tables in database.</param>
    ''' <param name="sessionRow">Session Row contain details of the current login user.</param>
    ''' <param name="DealerRow">Session Row contain details of the main dealer row as per logged in user</param>
    ''' <returns></returns>
    Public Function ProcessCustomersforRegistrationPostOrder(ByVal dt1 As DataTable, sessionRow As DataRow, DealerRow As DataRow, ByVal newButOld As String) As DataTable
        Dim LoginType As String = ""
        '  Dim jk As New List(Of Hashtable)
        Dim sessionkey As Integer = 0
        Dim Loginkey As Integer = 0
        Dim DealerLoginKey As Integer = -1
        Dim clswebsession As New WebSessions.WebSessions.WebSessions
        If sessionRow IsNot Nothing Then
            clswebsession.PrevRow = df1.UpdateDataRows(clswebsession.PrevRow, sessionRow)
            LoginType = sessionRow("linktype")
            Loginkey = sessionRow("linkcode")
        End If
        If DealerRow IsNot Nothing Then
            DealerLoginKey = df1.GetCellValue(DealerRow, "P_acccode")
        End If

        For i = 0 To dt1.Rows.Count - 1

            If Not UCase(dt1.Rows(i).Item("isvalid")) = UCase("true") Then Continue For

            If dt1.Rows(i).Item("chargingtrue").ToString.Trim = "false" And dt1.Rows(i).Item("paymentflag") = True Then Continue For

            If LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "new" And UCase(dt1.Rows(i).Item("regtype2")) = UCase("home") Then
                Dim ClsCustomers As New Customers.Customers.Customers

                Dim ClsCustomerServices As New CustomerServices.CustomerServices.CustomerServices
                '   ClsCustomers.CurrRow("curregdate") = dt1.Rows(i).Item("curregdate")
                Dim Mcustomer As DataTable = df1.GetDataFromSql(ClsCustomers.ServerDatabase, ClsCustomers.TableName, "*", "", "rowstatus = 0 and P_customers =" & dt1.Rows(i).Item("P_Customers"), "", "")
                Dim dtrow As DataRow = Nothing
                If Mcustomer.Rows.Count > 0 Then
                    dtrow = Mcustomer.Rows(0)
                Else
                    Continue For
                End If
                Dim dt2 As New DataTable
                dt2 = ClsCustomerServices.CurrDt
                ClsCustomers.PrevRow = df1.UpdateDataRows(ClsCustomers.PrevRow, dtrow)
                ClsCustomers.CurrRow("currregdate") = dt1.Rows(i).Item("currregdate")
                ClsCustomerServices.PrevDt = df1.UpdateDataTables(ClsCustomerServices.PrevDt, dt2)  'add by Shweta

                Dim a As DataRow = dt2.NewRow
                a("CustomerServices_key") = -1
                a("P_customers") = dtrow.Item("P_customers")
                a("ServiceInstalledBy") = "D"
                a("ServiceInstallingAgent") = DealerLoginKey  'Session("Key")
                a("ServiceCode") = 2403
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    a("ServiceStartDate") = dt1.Rows(i).Item("CurrRegDate")
                End If
                a("Quantity") = 1
                a("RecurrServiceCode") = 2774
                a("RecurrChgDueAfter") = 365
                a("Verified") = "N"

                a("serviceno") = GetServiceNoCservices(dt1.Rows(i).Item("P_customers"))
                dt2.Rows.Add(a)

                ClsCustomerServices.CurrDt = dt2


                Dim HashPublicValues As New Hashtable
                Dim aClsObject() As Object = {ClsCustomers, ClsCustomerServices, clswebsession}
                Dim success As Integer = cfc.SaveIntodb(aClsObject)

            ElseIf (LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "amc" And LCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = "main") Or newButOld = "true" Then
                Dim ClsCustomers As New Customers.Customers.Customers

                Dim ClsCustomerServices As New CustomerServices.CustomerServices.CustomerServices
                Dim BilledUpto As New Date(1901, 1, 1)
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CalculatedAllowUpto")) Then
                    BilledUpto = dt1.Rows(i).Item("CalculatedAllowUpto")
                End If
                Dim CurrRegDate As New Date(1901, 1, 1)
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    CurrRegDate = dt1.Rows(i).Item("CurrRegDate")
                End If

                Dim ldt As New DataTable
                Dim mLogin As DataTable = df1.GetDataFromSql(ClsCustomers.ServerDatabase, ClsCustomers.TableName, "*", "", "rowstatus = 0 and P_customers =" & dt1.Rows(i).Item("P_Customers"), "", "")
                Dim dtrow As DataRow = mLogin.Rows(0)

                ClsCustomers.PrevRow = df1.UpdateDataRows(ClsCustomers.PrevRow, dtrow)
                ClsCustomers.CurrRow("CeilingDate") = BilledUpto
                ClsCustomers.CurrRow("CurrRegDate") = CurrRegDate
                '  ClsCustomers.CurrRow("verified") = "N"
                If DBNull.Value.Equals(dt1.Rows(i).Item("changeallowupto")) Then
                    ClsCustomers.CurrRow("AllowUpto") = "1900-01-01 00:00:00.000"
                Else
                    ClsCustomers.CurrRow("AllowUpto") = dt1.Rows(i).Item("changeallowupto")
                End If

                Dim nodes As Integer = 0
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("nodesDb")) Then
                    nodes = dt1.Rows(i).Item("nodesDB")
                End If

                Dim changenode As Integer = 0
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("changenodes")) Then
                    changenode = dt1.Rows(i).Item("changenodes")
                End If

                If nodes <> changenode Then

                    Dim dt2 As New DataTable
                    dt2 = ClsCustomerServices.CurrDt

                    ClsCustomerServices.PrevDt = df1.UpdateDataTables(ClsCustomerServices.PrevDt, dt2)  'add by Shweta

                    Dim a As DataRow = dt2.NewRow
                    a("CustomerServices_key") = -1
                    a("P_customers") = dtrow.Item("P_customers")
                    a("ServiceInstalledBy") = "D"
                    a("ServiceInstallingAgent") = DealerLoginKey    'Session("Key")
                    a("ServiceCode") = 2402
                    If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                        a("ServiceStartDate") = dt1.Rows(i).Item("CurrRegDate")
                    End If

                    a("serviceno") = GetServiceNoCservices(dt1.Rows(i).Item("P_customers"))
                    a("Quantity") = changenode - nodes
                    a("RecurrServiceCode") = 2773
                    a("RecurrChgDueAfter") = 365
                    a("Verified") = "N"
                    dt2.Rows.Add(a)
                    ClsCustomerServices.CurrDt = dt2
                End If


                Dim aClsObject() As Object = {ClsCustomers, ClsCustomerServices, clswebsession}
                Dim success As Integer = cfc.SaveIntodb(aClsObject)


            ElseIf LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "amc" And LCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = "home" Then
                Dim ClsCustomers As New Customers.Customers.Customers

                Dim ClsCustomerServices As New CustomerServices.CustomerServices.CustomerServices

                Dim mLogin As DataTable = df1.GetDataFromSql(ClsCustomers.ServerDatabase, ClsCustomers.TableName, "*", "", "rowstatus = 0 and P_customers =" & dt1.Rows(i).Item("P_Customers"), "", "")
                Dim dtrow As DataRow = mLogin.Rows(0)

                ClsCustomers.PrevRow = df1.UpdateDataRows(ClsCustomers.PrevRow, dtrow)
                ClsCustomers.CurrRow("Currregdate") = dt1.Rows(i).Item("currregdate")
                Dim aClsObject() As Object = {ClsCustomers, ClsCustomerServices, clswebsession}
                Dim success As Integer = cfc.SaveIntodb(aClsObject)

            End If
        Next











        Return dt1
    End Function
    Public Sub UpdateRegistrationsOnVerification(ByVal dtregTran As DataRow, ByVal sessionRow As DataRow)

        Dim regtype As String = Trim(df1.GetCellValue(dtregTran, "regtype", "string"))
        Dim regtype2 As String = Trim(df1.GetCellValue(dtregTran, "regtype2", "string"))

        If Trim(regtype) = "new" And Trim(regtype2) = "main" Then

            Dim clsRegistrations As New Registrations.Registrations.Registrations

            clsRegistrations.CurrRow.Item("FolderStamp") = df1.GetCellValue(dtregTran, "folderstamps")
            clsRegistrations.CurrRow.Item("HardwareString") = "main" & Chr(200) & df1.GetCellValue(dtregTran, "machineline")
            clsRegistrations.CurrRow.Item("RegistrationString") = df1.GetCellValue(dtregTran, "registrationstring")
            clsRegistrations.CurrRow.Item("verified") = "N"
            clsRegistrations.CurrRow.Item("P_customers") = df1.GetCellValue(dtregTran, "p_customers") 'dt1(i).Item("P_customers")
            clsRegistrations.CurrRow.Item("websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer")
            '   dt3.Rows.Add(rg)
            '  clsRegistrations.CurrDt = df1.UpdateDataTables(clsRegistrations.CurrDt, dt3)  'add by shweta

            Dim HashPublicValues As New Hashtable
            Dim aClsObject() As Object = {clsRegistrations}
            Dim success As Integer = cfc.SaveIntodb(aClsObject)

            'added by Shweta
        ElseIf regtype = "new" And regtype2 = "home" Then

            Dim ClsRegistrations As New Registrations.Registrations.Registrations
            '   ClsCustomers.CurrRow("curregdate") = dt1.Rows(i).Item("curregdate")
            Dim blnmachVerify As Boolean = False
            Dim dtReg As New DataTable
            dtReg = df1.GetDataFromSql(ClsRegistrations.ServerDatabase, ClsRegistrations.TableName, "*", "", "rowstatus = 0 and P_customers=" & df1.GetCellValue(dtregTran, "p_customers"), "", "")
            If dtReg.Rows.Count > 0 Then
                Dim dtre As DataRow = dtReg.Rows(0)
                ClsRegistrations.PrevRow = df1.UpdateDataRows(ClsRegistrations.PrevRow, dtre)
                '  If UCase(dt1.Rows(i).Item("regtype2")) = "HOME" Then

                Dim prevHardware As String = ""
                Dim prevFolderstamp As String = ""
                Dim prevRegistrationString As String = ""
                prevHardware = df1.GetCellValue(dtre, "hardwarestring")
                prevFolderstamp = df1.GetCellValue(dtre, "folderstamp")
                prevRegistrationString = df1.GetCellValue(dtre, "registrationstring")

                Dim hardWrArr() As String = Split(prevHardware, Chr(201))
                For op = 0 To hardWrArr.Count - 1

                    If InStr(hardWrArr(op), "home^") <= 0 Then
                        prevHardware = Chr(201) & "home^1" & Chr(200) & prevHardware
                    Else
                        Dim instlTypArr() As String = Split(hardWrArr(op), Chr(200))
                        If instlTypArr.Count = 2 Then
                            Dim HomeNoArr() As String = Split(instlTypArr(0), "^")
                            Dim instum As Integer = CInt(HomeNoArr(1))
                            prevHardware = Chr(201) & "home^" & CStr(instum + 1) & Chr(200) & prevHardware
                        End If
                    End If




                Next

                '  If prevHardware = "" Then prevHardware = "" Else prevHardware = Chr(201) & "home" & Chr(200) & prevHardware
                If prevFolderstamp = "" Then prevFolderstamp = "" Else prevFolderstamp = "#" & prevFolderstamp
                If prevRegistrationString = "" Then prevRegistrationString = "" Else prevRegistrationString = "#" & prevRegistrationString
                ClsRegistrations.CurrRow("hardwarestring") = df1.GetCellValue(dtregTran, "machineline") & prevHardware
                ClsRegistrations.CurrRow("folderstamp") = df1.GetCellValue(dtregTran, "folderstamp") & prevFolderstamp
                ClsRegistrations.CurrRow("registrationstring") = df1.GetCellValue(dtregTran, "registrationstring") & prevRegistrationString
                ClsRegistrations.CurrRow("verified") = "N"
                ClsRegistrations.CurrRow("Websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer")


            Else



                ClsRegistrations.CurrRow.Item("FolderStamp") = df1.GetCellValue(dtregTran, "folderstamps") 'dt1(i).Item("folderstamp")
                ClsRegistrations.CurrRow.Item("HardwareString") = "home^1" & Chr(200) & df1.GetCellValue(dtregTran, "machineline") 'dt1(i).Item("machineline")
                ClsRegistrations.CurrRow.Item("RegistrationString") = df1.GetCellValue(dtregTran, "registrationstring") ' dt1(i).Item("RegistrationString")
                ClsRegistrations.CurrRow.Item("verified") = "N"
                ClsRegistrations.CurrRow.Item("P_customers") = df1.GetCellValue(dtregTran, "p_customers") 'dt1(i).Item("P_customers")
                ClsRegistrations.CurrRow.Item("websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer")


            End If



            Dim HashPublicValues As New Hashtable
            Dim aClsObject() As Object = {ClsRegistrations}
            Dim success As Integer = cfc.SaveIntodb(aClsObject)

        ElseIf regtype = "amc" And regtype2 = "main" Then '(LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "amc" And LCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = "main") Or newButOld = True Then

            Dim clsReg As New Registrations.Registrations.Registrations


            Dim dtReg As New DataTable
            dtReg = df1.GetDataFromSql(clsReg.ServerDatabase, clsReg.TableName, "*", "", "P_customers=" & df1.GetCellValue(dtregTran, "p_customers"), "", "")
            Dim blnmachVerify As Boolean = False
            If dtReg.Rows.Count > 0 Then
                Dim dtre As DataRow = dtReg.Rows(0)
                clsReg.PrevRow = df1.UpdateDataRows(clsReg.PrevRow, dtre)
                '  If UCase(dt1.Rows(i).Item("regtype2")) = "HOME" Then
                Dim prevHardware As String = ""
                Dim prevFolderstamp As String = ""
                Dim prevRegistrationString As String = ""
                prevHardware = df1.GetCellValue(dtre, "hardwarestring")
                prevFolderstamp = df1.GetCellValue(dtre, "folderstamp")
                prevRegistrationString = df1.GetCellValue(dtre, "registrationstring")

                Dim hrdStr As String = df1.GetCellValue(dtregTran, "machineline")  ' dt1.Rows(i).Item("machineline").ToString
                Dim hrdStrarr() As String = Split(hrdStr, "~")
                Dim strprevhardware1 As String = ""     ' effective hardware string from uploaded epl
                For tr = 6 To hrdStrarr.Count - 1
                    strprevhardware1 = strprevhardware1 & hrdStrarr(tr)
                Next
                Dim hrdwrArr() As String = Split(prevHardware, Chr(201))
                Dim posOfMatchStr As Int16 = -1
                For rt = 0 To hrdwrArr.Count - 1


                    'getting the effective hardware string from epl uploaded -- 
                    Dim prevhrdStr() As String = Split(hrdwrArr(rt), Chr(200))
                    If prevhrdStr.Count = 2 Then
                        If LCase(prevhrdStr(0)) = "main" Then
                            Dim actHRDStr() As String = Split(prevhrdStr(1), "~")
                            '  Strings.Join()
                            Dim strprevhardware As String = ""        'effective hardware string from strings stored in db when convention includes label such as main
                            For tr = 6 To actHRDStr.Count - 1
                                strprevhardware = strprevhardware & actHRDStr(tr)
                            Next
                            posOfMatchStr = rt
                            If LCase(Trim(strprevhardware)) = LCase(Trim(strprevhardware1)) Then
                                blnmachVerify = True
                            End If
                        End If

                    End If
                Next
                Dim finalHRdStr As String = ""
                If posOfMatchStr >= 0 Then
                    Dim kls As String = "main" & Chr(200) & hrdStr
                    hrdwrArr(posOfMatchStr) = kls
                    finalHRdStr = Strings.Join(hrdwrArr, Chr(201))
                Else
                    finalHRdStr = "main" & Chr(200) & hrdStr
                End If
                ' Dim finalHRdStr As String = 

                '  If prevHardware = "" Then prevHardware = "" Else prevHardware = "#" & prevHardware
                If prevFolderstamp = "" Then prevFolderstamp = "" Else prevFolderstamp = "#" & prevFolderstamp
                If prevRegistrationString = "" Then prevRegistrationString = "" Else prevRegistrationString = "#" & prevRegistrationString





                clsReg.CurrRow("hardwarestring") = finalHRdStr 'dt1.Rows(i).Item("machineline").ToString & prevHardware
                '     End If
                clsReg.CurrRow("folderstamp") = df1.GetCellValue(dtregTran, "folderstamps") & prevFolderstamp
                clsReg.CurrRow("registrationstring") = df1.GetCellValue(dtregTran, "registrationstring") & prevRegistrationString
                clsReg.CurrRow("verified") = "N"
                clsReg.CurrRow("Websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer")

            Else
                clsReg.CurrRow("P_Customers") = df1.GetCellValue(dtregTran, "p_customers")  'dt1(i).Item("P_customers")
                clsReg.CurrRow("hardwarestring") = "main" & Chr(200) & df1.GetCellValue(dtregTran, "machineline") ' dt1.Rows(i).Item("machineline").ToString
                clsReg.CurrRow("folderstamp") = df1.GetCellValue(dtregTran, "folderstamps") ' & prevFolderstamp
                clsReg.CurrRow("registrationstring") = df1.GetCellValue(dtregTran, "registrationstring")  'prevRegistrationString
                clsReg.CurrRow("Websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer")
                clsReg.CurrRow("verified") = "N"

            End If

            Dim HashPublicValues As New Hashtable
            Dim aClsObject() As Object = {clsReg}
            Dim success As Integer = cfc.SaveIntodb(aClsObject)


        ElseIf regtype = "amc" And regtype2 = "home" Then
            Dim clsReg As New Registrations.Registrations.Registrations
            Dim blnmachVerify As Boolean = False
            Dim dtReg As New DataTable
            dtReg = df1.GetDataFromSql(clsReg.ServerDatabase, clsReg.TableName, "*", "", "P_customers=" & df1.GetCellValue(dtregTran, "p_customers"), "", "")
            If dtReg.Rows.Count > 0 Then
                Dim dtre As DataRow = dtReg.Rows(0)
                clsReg.PrevRow = df1.UpdateDataRows(clsReg.PrevRow, dtre)
                '  If UCase(dt1.Rows(i).Item("regtype2")) = "HOME" Then
                Dim prevHardware As String = ""
                Dim prevFolderstamp As String = ""
                Dim prevRegistrationString As String = ""
                prevHardware = df1.GetCellValue(dtre, "hardwarestring")
                prevFolderstamp = df1.GetCellValue(dtre, "folderstamp")
                prevRegistrationString = df1.GetCellValue(dtre, "registrationstring")
                Dim hrdwrArr() As String = Split(prevHardware, Chr(201))
                Dim hrdStr As String = df1.GetCellValue(dtregTran, "machineline")
                Dim hrdStrarr() As String = Split(hrdStr, "~")
                Dim strprevhardware1 As String = ""  ' effective hardware string from uploaded epl
                Dim posOfMatchStr As Int16 = -1
                For tr = 6 To hrdStrarr.Count - 1
                    strprevhardware1 = strprevhardware1 & hrdStrarr(tr)
                Next
                Dim homeFound As Boolean = False
                For rt = 0 To hrdwrArr.Count - 1
                    Dim IndHrdwrArr() As String = Split(hrdwrArr(rt), Chr(200))
                    If IndHrdwrArr.Count = 2 Then
                        Dim abc() As String = Split(IndHrdwrArr(0), "^")
                        If Trim(LCase(abc(0))) = "home" Then
                            homeFound = True
                            Dim actHRDStr() As String = Split(IndHrdwrArr(1), "~")
                            '  Strings.Join()
                            Dim strprevhardware As String = ""        'effective hardware string from strings stored in db when convention includes label such as main
                            For tr = 6 To actHRDStr.Count - 1
                                strprevhardware = strprevhardware & actHRDStr(tr)
                            Next
                            '  posOfMatchStr = rt
                            If LCase(Trim(strprevhardware)) = LCase(Trim(strprevhardware1)) Then
                                blnmachVerify = True
                                posOfMatchStr = rt
                            End If
                        End If
                    End If




                Next

                Dim finalHRdStr As String = ""
                If posOfMatchStr >= 0 Then

                    Dim str1 As String = hrdwrArr(posOfMatchStr)
                    Dim str1arr() As String = Split(str1, Chr(200))
                    hrdwrArr(posOfMatchStr) = str1arr(0) & Chr(200) & hrdStr
                    finalHRdStr = Strings.Join(hrdwrArr, Chr(201))
                Else
                    If homeFound = False Then
                        finalHRdStr = prevHardware & Chr(201) & "home^1" & Chr(200) & hrdStr
                    Else
                        Dim homeInt As Int16 = 0
                        Dim posoFMtch As Int16 = -1
                        For rt = 0 To hrdwrArr.Count - 1
                            Dim IndHrdwrArr() As String = Split(hrdwrArr(rt), Chr(200))
                            If IndHrdwrArr.Count = 2 Then
                                Dim abc() As String = Split(IndHrdwrArr(0), "^")
                                If abc.Count = 2 Then
                                    If CInt(abc(1)) > homeInt Then
                                        homeInt = CInt(abc(1))
                                        posoFMtch = rt
                                    End If
                                End If
                            End If
                        Next
                        hrdwrArr(posoFMtch) = "home^" & CStr(homeInt + 1) & Chr(200) & hrdStr
                        blnmachVerify = False
                        finalHRdStr = Strings.Join(hrdwrArr, Chr(201))
                    End If


                End If

                If prevFolderstamp = "" Then prevFolderstamp = "" Else prevFolderstamp = "#" & prevFolderstamp

                If prevRegistrationString = "" Then prevRegistrationString = "" Else prevRegistrationString = "#" & prevRegistrationString
                ' If Not blnmachVerify Then     'if hardware string is found , then don't add
                clsReg.CurrRow("hardwarestring") = finalHRdStr ' dt1.Rows(i).Item("machineline").ToString & prevHardware
                ' End If
                clsReg.CurrRow("folderstamp") = df1.GetCellValue(dtregTran, "folderstamps") & prevFolderstamp
                clsReg.CurrRow("registrationstring") = df1.GetCellValue(dtregTran, "registrationstring") & prevRegistrationString
                clsReg.CurrRow("verified") = "N"
                clsReg.CurrRow("Websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer")

            Else





                clsReg.CurrRow("P_Customers") = df1.GetCellValue(dtregTran, "p_customers")  'dt1(i).Item("P_customers")
                clsReg.CurrRow("hardwarestring") = "home^1" & Chr(200) & df1.GetCellValue(dtregTran, "hardwarestring")
                clsReg.CurrRow("folderstamp") = df1.GetCellValue(dtregTran, "folderstamps") ' & prevFolderstamp
                clsReg.CurrRow("registrationstring") = df1.GetCellValue(dtregTran, "registrationstring")  'prevRegistrationString


                clsReg.CurrRow("Websessions_key") = clsReg.CurrRow("Websessions_key") = df1.GetCellValue(sessionRow, "websessions_key", "integer") ' sessionRow("websessions_key")
                clsReg.CurrRow("verified") = "N"






            End If


            Dim HashPublicValues As New Hashtable
            Dim aClsObject() As Object = {clsReg}
            Dim success As Integer = cfc.SaveIntodb(aClsObject)

        End If







    End Sub
    Public Function AddRowInRegistrationsTran_new(ByVal dt1 As DataTable, sessionRow As DataRow, ByVal EPLHashList As List(Of Hashtable), ByVal newButOld As String) As Hashtable
        Dim registrationTranHAsh As New Hashtable

        Dim LoginType As String = ""

        Dim sessionkey As Integer = 0
        Dim Loginkey As Integer = 0
        Dim clswebsession As New WebSessions.WebSessions.WebSessions
        If sessionRow IsNot Nothing Then
            clswebsession.PrevRow = df1.UpdateDataRows(clswebsession.PrevRow, sessionRow)
            LoginType = sessionRow("linktype")
            Loginkey = sessionRow("linkcode")
        End If


        For i = 0 To dt1.Rows.Count - 1

            If Not UCase(dt1.Rows(i).Item("isvalid")) = UCase("true") Then Continue For
            If dt1.Rows(i).Item("chargingtrue") = "false" And dt1.Rows(i).Item("paymentflag") = True Then Continue For
            If LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "new" And UCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = UCase("main") Then
                'If newButOld = True Then
                '    Continue For
                'End If

                Dim ClsRegTran As New RegistrationTran.RegistrationTran.RegistrationTran
                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    ClsRegTran.CurrRow.Item("RegsendDate") = dt1.Rows(i).Item("CurrRegDate")
                End If
                ClsRegTran.CurrRow.Item("P_customers") = dt1.Rows(i).Item("P_customers")
                ClsRegTran.CurrRow.Item("Regtype") = dt1.Rows(i).Item("regtype").ToString.Trim
                ClsRegTran.CurrRow.Item("Regtype2") = dt1.Rows(i).Item("regtype2").ToString.Trim
                ClsRegTran.CurrRow.Item("AllowuptoDate") = dt1.Rows(i).Item("CalculatedAllowUpto")
                ClsRegTran.CurrRow.Item("Openedupto") = dt1.Rows(i).Item("changeallowupto")
                ClsRegTran.CurrRow.Item("Lan") = dt1.Rows(i).Item("changelan")
                ClsRegTran.CurrRow.Item("Node") = dt1.Rows(i).Item("changenodes")
                ClsRegTran.CurrRow.Item("Websessions_key") = sessionRow("websessions_key")
                ClsRegTran.CurrRow.Item("timestamp") = df1.getDateTimeISTNow
                Dim eplhash As New Hashtable
                For y = 0 To EPLHashList.Count - 1
                    Dim p_cust As Integer = GF1.GetValueFromHashTable(EPLHashList(y), "p_customers")
                    If p_cust = dt1.Rows(i).Item("P_customers") Then
                        eplhash = EPLHashList(y)
                        Exit For
                    End If
                Next

                ClsRegTran.CurrRow("add4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add4"), 200)

                ClsRegTran.CurrRow("mobile") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "mobile"), 10)

                ClsRegTran.CurrRow("pmtr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtr"), 50)




                ClsRegTran.CurrRow("folderstamps") = GF1.GetValueFromHashTable(eplhash, "folderstamps")
                ClsRegTran.CurrRow("machineline") = GF1.GetValueFromHashTable(eplhash, "machineline")
                ClsRegTran.CurrRow("pmtd2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd2"), 500)
                ClsRegTran.CurrRow("busscd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "busscd"), 6)
                ClsRegTran.CurrRow("pmtd1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd1"), 500)
                If GF1.GetValueFromHashTable(eplhash, "regdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("regdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "regdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("regdate") = New Date(1900, 1, 1)
                    End Try
                End If


                ClsRegTran.CurrRow("machverifyflag") = "Y"
                ClsRegTran.CurrRow("chkd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "chkd"), 10)

                ClsRegTran.CurrRow("add3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add3"), 200)
                ClsRegTran.CurrRow("d3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d3"), 50)
                ClsRegTran.CurrRow("gst") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "gst"), 15)
                ClsRegTran.CurrRow("email") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "email"), 100)


                If GF1.GetValueFromHashTable(eplhash, "currentceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("currentceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "currentceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("currentceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If


                ClsRegTran.CurrRow("city") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "city"), 25)
                ClsRegTran.CurrRow("d6") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d6"), 50)
                If GF1.GetValueFromHashTable(eplhash, "previousceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("previousceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "previousceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("previousceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("add1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add1"), 200)
                ClsRegTran.CurrRow("pcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pcode"), 8)

                ClsRegTran.CurrRow("buss") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "buss"), 50)
                ClsRegTran.CurrRow("add2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add2"), 200)
                ClsRegTran.CurrRow("svol") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "svol"), 5)
                ClsRegTran.CurrRow("hospital") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "hospital"), 10)
                ClsRegTran.CurrRow("pterm") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pterm"), 50)
                ClsRegTran.CurrRow("sdir") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sdir"), 100)
                ClsRegTran.CurrRow("d1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d1"), 10)
                ClsRegTran.CurrRow("sale_imp") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sale_imp"), 50)
                ClsRegTran.CurrRow("dealer") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "dealer"), 50)
                ClsRegTran.CurrRow("lanepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "lan"), 5)

                ClsRegTran.CurrRow("statec") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "statec"), 8)
                ClsRegTran.CurrRow("clname") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "clname"), 100)
                ClsRegTran.CurrRow("d4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d4"), 50)
                ClsRegTran.CurrRow("nodesepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "nodes"), 3)
                ClsRegTran.CurrRow("state") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "state"), 20)
                ClsRegTran.CurrRow("printr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "printr"), 50)
                ClsRegTran.CurrRow("charges") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "charges"), 20)
                ClsRegTran.CurrRow("regtypeepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "regtype"), 5)
                ClsRegTran.CurrRow("ph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "ph"), 50)
                ClsRegTran.CurrRow("person") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "person"), 50)

                ClsRegTran.CurrRow("cpu") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "cpu"), 50)
                ClsRegTran.CurrRow("oldcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "oldcode"), 8)
                ClsRegTran.CurrRow("instb") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instb"), 50)
                ClsRegTran.CurrRow("d2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d2"), 50)
                ClsRegTran.CurrRow("citycode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "citycode"), 8)
                ClsRegTran.CurrRow("instd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instd"), 15)
                ClsRegTran.CurrRow("d5") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d5"), 50)
                If GF1.GetValueFromHashTable(eplhash, "amcdt") <> "" Then
                    Try
                        ClsRegTran.CurrRow("amcdt") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "amcdt"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("amcdt") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("file0") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "file0"), 50)
                ClsRegTran.CurrRow("serviceph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "serviceph"), 50)




                Dim HashPublicValues As New Hashtable
                Dim aClsObject() As Object = {ClsRegTran}
                Dim regtran_key As Integer = cfc.SaveIntodbGetKey(aClsObject, "registrationtran", "registrationtran_key")
                registrationTranHAsh = GF1.AddItemToHashTable(registrationTranHAsh, CStr(dt1.Rows(i).Item("P_customers")), regtran_key)
                'added by Shweta
            ElseIf LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "new" And UCase(dt1.Rows(i).Item("regtype2")) = UCase("home") Then

                '  Dim ClsRegistrations As New Registrations.Registrations.Registrations
                '   ClsCustomers.CurrRow("curregdate") = dt1.Rows(i).Item("curregdate")
                Dim blnmachVerify As Boolean = False
                Dim dtReg As New DataTable

                Dim ClsRegTran As New RegistrationTran.RegistrationTran.RegistrationTran

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    ClsRegTran.CurrRow.Item("RegsendDate") = dt1.Rows(i).Item("CurrRegDate")
                End If
                'If blnmachVerify = True Then
                '    ClsRegTran.CurrRow.Item("machverifyflag") = "Y"
                'Else
                '    ClsRegTran.CurrRow.Item("machverifyflag") = "N"
                'End If
                ClsRegTran.CurrRow.Item("Regtype") = dt1.Rows(i).Item("regtype").ToString.Trim
                ClsRegTran.CurrRow.Item("Regtype2") = dt1.Rows(i).Item("regtype2").ToString.Trim
                ClsRegTran.CurrRow.Item("AllowuptoDate") = dt1.Rows(i).Item("CalculatedAllowUpto")
                ClsRegTran.CurrRow.Item("Openedupto") = dt1.Rows(i).Item("changeallowupto")
                ClsRegTran.CurrRow.Item("timestamp") = df1.getDateTimeISTNow
                ClsRegTran.CurrRow.Item("P_customers") = dt1.Rows(i).Item("P_customers")
                ClsRegTran.CurrRow.Item("websessions_key") = sessionRow("websessions_key")


                Dim eplhash As New Hashtable
                For y = 0 To EPLHashList.Count - 1
                    Dim p_cust As Integer = GF1.GetValueFromHashTable(EPLHashList(y), "p_customers")
                    If p_cust = dt1.Rows(i).Item("P_customers") Then
                        eplhash = EPLHashList(y)
                        Exit For
                    End If
                Next


                ClsRegTran.CurrRow("add4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add4"), 200)

                ClsRegTran.CurrRow("mobile") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "mobile"), 10)

                ClsRegTran.CurrRow("pmtr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtr"), 50)




                ClsRegTran.CurrRow("folderstamps") = GF1.GetValueFromHashTable(eplhash, "folderstamps")
                ClsRegTran.CurrRow("machineline") = GF1.GetValueFromHashTable(eplhash, "machineline")
                ClsRegTran.CurrRow("pmtd2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd2"), 500)
                ClsRegTran.CurrRow("busscd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "busscd"), 6)
                ClsRegTran.CurrRow("pmtd1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd1"), 500)
                If GF1.GetValueFromHashTable(eplhash, "regdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("regdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "regdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("regdate") = New Date(1900, 1, 1)
                    End Try
                End If


                ClsRegTran.CurrRow("machverifyflag") = "Y"
                ClsRegTran.CurrRow("chkd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "chkd"), 10)

                ClsRegTran.CurrRow("add3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add3"), 200)
                ClsRegTran.CurrRow("d3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d3"), 50)
                ClsRegTran.CurrRow("gst") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "gst"), 15)
                ClsRegTran.CurrRow("email") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "email"), 100)


                If GF1.GetValueFromHashTable(eplhash, "currentceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("currentceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "currentceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("currentceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If


                ClsRegTran.CurrRow("city") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "city"), 25)
                ClsRegTran.CurrRow("d6") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d6"), 50)
                If GF1.GetValueFromHashTable(eplhash, "previousceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("previousceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "previousceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("previousceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("add1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add1"), 200)
                ClsRegTran.CurrRow("pcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pcode"), 8)

                ClsRegTran.CurrRow("buss") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "buss"), 50)
                ClsRegTran.CurrRow("add2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add2"), 200)
                ClsRegTran.CurrRow("svol") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "svol"), 5)
                ClsRegTran.CurrRow("hospital") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "hospital"), 10)
                ClsRegTran.CurrRow("pterm") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pterm"), 50)
                ClsRegTran.CurrRow("sdir") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sdir"), 100)
                ClsRegTran.CurrRow("d1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d1"), 10)
                ClsRegTran.CurrRow("sale_imp") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sale_imp"), 50)
                ClsRegTran.CurrRow("dealer") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "dealer"), 50)
                ClsRegTran.CurrRow("lanepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "lan"), 5)

                ClsRegTran.CurrRow("statec") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "statec"), 8)
                ClsRegTran.CurrRow("clname") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "clname"), 100)
                ClsRegTran.CurrRow("d4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d4"), 50)
                ClsRegTran.CurrRow("nodesepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "nodes"), 3)
                ClsRegTran.CurrRow("state") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "state"), 20)
                ClsRegTran.CurrRow("printr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "printr"), 50)
                ClsRegTran.CurrRow("charges") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "charges"), 20)
                ClsRegTran.CurrRow("regtypeepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "regtype"), 5)
                ClsRegTran.CurrRow("ph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "ph"), 50)
                ClsRegTran.CurrRow("person") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "person"), 50)

                ClsRegTran.CurrRow("cpu") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "cpu"), 50)
                ClsRegTran.CurrRow("oldcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "oldcode"), 8)
                ClsRegTran.CurrRow("instb") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instb"), 50)
                ClsRegTran.CurrRow("d2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d2"), 50)
                ClsRegTran.CurrRow("citycode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "citycode"), 8)
                ClsRegTran.CurrRow("instd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instd"), 15)
                ClsRegTran.CurrRow("d5") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d5"), 50)
                If GF1.GetValueFromHashTable(eplhash, "amcdt") <> "" Then
                    Try
                        ClsRegTran.CurrRow("amcdt") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "amcdt"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("amcdt") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("file0") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "file0"), 50)
                ClsRegTran.CurrRow("serviceph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "serviceph"), 50)


                Dim HashPublicValues As New Hashtable


                Dim aClsObject() As Object = {ClsRegTran}


                Dim regtran_key As Integer = cfc.SaveIntodbGetKey(aClsObject, "registrationtran", "registrationtran_key")
                registrationTranHAsh = GF1.AddItemToHashTable(registrationTranHAsh, CStr(dt1.Rows(i).Item("P_customers")), regtran_key)

                '   Dim success As Integer = SaveIntodb(aClsObject, "D")

            ElseIf (LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "amc" And LCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = "main") Or newButOld = True Then


                Dim ClsRegTran As New RegistrationTran.RegistrationTran.RegistrationTran

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    ClsRegTran.CurrRow.Item("RegsendDate") = dt1.Rows(i).Item("CurrRegDate")
                End If

                ClsRegTran.CurrRow.Item("machverifyflag") = "Y"

                ClsRegTran.CurrRow.Item("Regtype") = dt1.Rows(i).Item("regtype").ToString.Trim
                ClsRegTran.CurrRow.Item("Regtype2") = dt1.Rows(i).Item("regtype2").ToString.Trim
                ClsRegTran.CurrRow.Item("AllowuptoDate") = dt1.Rows(i).Item("CalculatedAllowUpto")
                ClsRegTran.CurrRow.Item("Openedupto") = dt1.Rows(i).Item("changeallowupto")
                ClsRegTran.CurrRow.Item("Lan") = dt1.Rows(i).Item("changelan")
                ClsRegTran.CurrRow.Item("Node") = dt1.Rows(i).Item("changenodes")
                ClsRegTran.CurrRow.Item("timestamp") = df1.getDateTimeISTNow
                ClsRegTran.CurrRow.Item("P_customers") = dt1.Rows(i).Item("p_customers")
                ClsRegTran.CurrRow.Item("websessions_key") = sessionRow("websessions_key")

                Dim eplhash As New Hashtable
                For y = 0 To EPLHashList.Count - 1
                    Dim p_cust As Integer = GF1.GetValueFromHashTable(EPLHashList(y), "p_customers")
                    If p_cust = dt1.Rows(i).Item("P_customers") Then
                        eplhash = EPLHashList(y)
                    End If
                Next





                ClsRegTran.CurrRow("add4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add4"), 200)

                ClsRegTran.CurrRow("mobile") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "mobile"), 10)

                ClsRegTran.CurrRow("pmtr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtr"), 50)




                ClsRegTran.CurrRow("folderstamps") = GF1.GetValueFromHashTable(eplhash, "folderstamps")
                ClsRegTran.CurrRow("machineline") = GF1.GetValueFromHashTable(eplhash, "machineline")
                ClsRegTran.CurrRow("pmtd2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd2"), 500)
                ClsRegTran.CurrRow("busscd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "busscd"), 6)
                ClsRegTran.CurrRow("pmtd1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd1"), 500)
                If GF1.GetValueFromHashTable(eplhash, "regdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("regdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "regdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("regdate") = New Date(1900, 1, 1)
                    End Try
                End If



                ClsRegTran.CurrRow("chkd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "chkd"), 10)

                ClsRegTran.CurrRow("add3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add3"), 200)
                ClsRegTran.CurrRow("d3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d3"), 50)
                ClsRegTran.CurrRow("gst") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "gst"), 15)
                ClsRegTran.CurrRow("email") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "email"), 100)


                If GF1.GetValueFromHashTable(eplhash, "currentceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("currentceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "currentceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("currentceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If


                ClsRegTran.CurrRow("city") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "city"), 25)
                ClsRegTran.CurrRow("d6") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d6"), 50)
                If GF1.GetValueFromHashTable(eplhash, "previousceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("previousceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "previousceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("previousceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("add1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add1"), 200)
                ClsRegTran.CurrRow("pcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pcode"), 8)

                ClsRegTran.CurrRow("buss") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "buss"), 50)
                ClsRegTran.CurrRow("add2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add2"), 200)
                ClsRegTran.CurrRow("svol") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "svol"), 5)
                ClsRegTran.CurrRow("hospital") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "hospital"), 10)
                ClsRegTran.CurrRow("pterm") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pterm"), 50)
                ClsRegTran.CurrRow("sdir") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sdir"), 100)
                ClsRegTran.CurrRow("d1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d1"), 10)
                ClsRegTran.CurrRow("sale_imp") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sale_imp"), 50)
                ClsRegTran.CurrRow("dealer") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "dealer"), 50)
                ClsRegTran.CurrRow("lanepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "lan"), 5)

                ClsRegTran.CurrRow("statec") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "statec"), 8)
                ClsRegTran.CurrRow("clname") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "clname"), 100)
                ClsRegTran.CurrRow("d4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d4"), 50)
                ClsRegTran.CurrRow("nodesepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "nodes"), 3)
                ClsRegTran.CurrRow("state") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "state"), 20)
                ClsRegTran.CurrRow("printr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "printr"), 50)
                ClsRegTran.CurrRow("charges") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "charges"), 20)
                ClsRegTran.CurrRow("regtypeepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "regtype"), 5)
                ClsRegTran.CurrRow("ph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "ph"), 50)
                ClsRegTran.CurrRow("person") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "person"), 50)

                ClsRegTran.CurrRow("cpu") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "cpu"), 50)
                ClsRegTran.CurrRow("oldcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "oldcode"), 8)
                ClsRegTran.CurrRow("instb") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instb"), 50)
                ClsRegTran.CurrRow("d2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d2"), 50)
                ClsRegTran.CurrRow("citycode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "citycode"), 8)
                ClsRegTran.CurrRow("instd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instd"), 15)
                ClsRegTran.CurrRow("d5") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d5"), 50)
                If GF1.GetValueFromHashTable(eplhash, "amcdt") <> "" Then
                    Try
                        ClsRegTran.CurrRow("amcdt") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "amcdt"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("amcdt") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("file0") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "file0"), 50)
                ClsRegTran.CurrRow("serviceph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "serviceph"), 50)
                Dim HashPublicValues As New Hashtable
                Dim aClsObject() As Object = {ClsRegTran}

                Dim regtran_key As Integer = cfc.SaveIntodbGetKey(aClsObject, "registrationtran", "registrationtran_key")
                registrationTranHAsh = GF1.AddItemToHashTable(registrationTranHAsh, CStr(dt1.Rows(i).Item("P_customers")), regtran_key)


                '  Dim success As Integer = SaveIntodb(aClsObject, "D")
            ElseIf LCase(dt1.Rows(i).Item("regtype").ToString.Trim) = "amc" And LCase(dt1.Rows(i).Item("regtype2").ToString.Trim) = "home" Then

                Dim ClsRegTran As New RegistrationTran.RegistrationTran.RegistrationTran

                If Not DBNull.Value.Equals(dt1.Rows(i).Item("CurrRegDate")) Then
                    ClsRegTran.CurrRow.Item("RegsendDate") = dt1.Rows(i).Item("CurrRegDate")
                End If
                ' If blnmachVerify = True Then
                ClsRegTran.CurrRow.Item("machverifyflag") = "Y"
                'Else
                ClsRegTran.CurrRow.Item("machverifyflag") = "N"
                'End If
                ClsRegTran.CurrRow.Item("Regtype") = dt1.Rows(i).Item("regtype").ToString.Trim
                ClsRegTran.CurrRow.Item("Regtype2") = dt1.Rows(i).Item("regtype2").ToString.Trim
                ClsRegTran.CurrRow.Item("AllowuptoDate") = dt1.Rows(i).Item("CalculatedAllowUpto")
                ClsRegTran.CurrRow.Item("Openedupto") = dt1.Rows(i).Item("changeallowupto")
                ClsRegTran.CurrRow.Item("P_customers") = dt1.Rows(i).Item("p_customers")
                ClsRegTran.CurrRow.Item("websessions_key") = sessionRow("websessions_key")

                ClsRegTran.CurrRow.Item("timestamp") = df1.getDateTimeISTNow


                Dim eplhash As New Hashtable
                For y = 0 To EPLHashList.Count - 1
                    Dim p_cust As Integer = GF1.GetValueFromHashTable(EPLHashList(y), "p_customers")
                    If p_cust = dt1.Rows(i).Item("P_customers") Then
                        eplhash = EPLHashList(y)
                    End If
                Next





                ClsRegTran.CurrRow("add4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add4"), 200)

                ClsRegTran.CurrRow("mobile") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "mobile"), 10)

                ClsRegTran.CurrRow("pmtr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtr"), 50)




                ClsRegTran.CurrRow("folderstamps") = GF1.GetValueFromHashTable(eplhash, "folderstamps")
                ClsRegTran.CurrRow("machineline") = GF1.GetValueFromHashTable(eplhash, "machineline")
                ClsRegTran.CurrRow("pmtd2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd2"), 500)
                ClsRegTran.CurrRow("busscd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "busscd"), 6)
                ClsRegTran.CurrRow("pmtd1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pmtd1"), 500)
                If GF1.GetValueFromHashTable(eplhash, "regdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("regdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "regdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("regdate") = New Date(1900, 1, 1)
                    End Try
                End If



                ClsRegTran.CurrRow("chkd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "chkd"), 10)

                ClsRegTran.CurrRow("add3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add3"), 200)
                ClsRegTran.CurrRow("d3") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d3"), 50)
                ClsRegTran.CurrRow("gst") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "gst"), 15)
                ClsRegTran.CurrRow("email") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "email"), 100)


                If GF1.GetValueFromHashTable(eplhash, "currentceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("currentceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "currentceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("currentceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If


                ClsRegTran.CurrRow("city") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "city"), 25)
                ClsRegTran.CurrRow("d6") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d6"), 50)
                If GF1.GetValueFromHashTable(eplhash, "previousceilingdate") <> "" Then
                    Try
                        ClsRegTran.CurrRow("previousceilingdate") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "previousceilingdate"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("previousceilingdate") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("add1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add1"), 200)
                ClsRegTran.CurrRow("pcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pcode"), 8)

                ClsRegTran.CurrRow("buss") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "buss"), 50)
                ClsRegTran.CurrRow("add2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "add2"), 200)
                ClsRegTran.CurrRow("svol") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "svol"), 5)
                ClsRegTran.CurrRow("hospital") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "hospital"), 10)
                ClsRegTran.CurrRow("pterm") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "pterm"), 50)
                ClsRegTran.CurrRow("sdir") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sdir"), 100)
                ClsRegTran.CurrRow("d1") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d1"), 10)
                ClsRegTran.CurrRow("sale_imp") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "sale_imp"), 50)
                ClsRegTran.CurrRow("dealer") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "dealer"), 50)
                ClsRegTran.CurrRow("lanepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "lan"), 5)

                ClsRegTran.CurrRow("statec") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "statec"), 8)
                ClsRegTran.CurrRow("clname") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "clname"), 100)
                ClsRegTran.CurrRow("d4") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d4"), 50)
                ClsRegTran.CurrRow("nodesepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "nodes"), 3)
                ClsRegTran.CurrRow("state") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "state"), 20)
                ClsRegTran.CurrRow("printr") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "printr"), 50)
                ClsRegTran.CurrRow("charges") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "charges"), 20)
                ClsRegTran.CurrRow("regtypeepl") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "regtype"), 5)
                ClsRegTran.CurrRow("ph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "ph"), 50)
                ClsRegTran.CurrRow("person") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "person"), 50)

                ClsRegTran.CurrRow("cpu") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "cpu"), 50)
                ClsRegTran.CurrRow("oldcode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "oldcode"), 8)
                ClsRegTran.CurrRow("instb") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instb"), 50)
                ClsRegTran.CurrRow("d2") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d2"), 50)
                ClsRegTran.CurrRow("citycode") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "citycode"), 8)
                ClsRegTran.CurrRow("instd") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "instd"), 15)
                ClsRegTran.CurrRow("d5") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "d5"), 50)
                If GF1.GetValueFromHashTable(eplhash, "amcdt") <> "" Then
                    Try
                        ClsRegTran.CurrRow("amcdt") = DateTime.ParseExact(GF1.GetValueFromHashTable(eplhash, "amcdt"), "dd/MM/yy", Nothing)
                    Catch
                        ClsRegTran.CurrRow("amcdt") = New Date(1900, 1, 1)
                    End Try
                End If
                ClsRegTran.CurrRow("file0") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "file0"), 50)
                ClsRegTran.CurrRow("serviceph") = cfc.getStringFixedSize(GF1.GetValueFromHashTable(eplhash, "serviceph"), 50)

                Dim HashPublicValues As New Hashtable
                Dim aClsObject() As Object = {ClsRegTran}
                Dim regtran_key As Integer = cfc.SaveIntodbGetKey(aClsObject, "registrationtran", "registrationtran_key")
                registrationTranHAsh = GF1.AddItemToHashTable(registrationTranHAsh, CStr(dt1.Rows(i).Item("P_customers")), regtran_key)

                ' Dim success As Integer = SaveIntodb(aClsObject, "D")

            End If
        Next
        '  Return dt1

        Return registrationTranHAsh
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dtcust"></param>
    Public Sub updateBilleduptoPostPayment(ByVal dtcust As DataTable)
        Dim ClsCustomers As New Customers.Customers.Customers
        For u = 0 To dtcust.Rows.Count - 1
            Dim billedupto As New DateTime
            billedupto = libcustomerfeature.GetBilledUpToDate(dtcust.Rows(u).Item("p_customers"), "main").ToString("yyyy-MM-dd hh:mm:ss")
            Dim upw As String = "update customers set billedupto ='" & billedupto & "' where rowstatus = 0 and p_customers = " & dtcust.Rows(u).Item("p_customers")
            df1.SqlExecuteNonQuery(ClsCustomers.ServerDatabase, upw)
        Next

    End Sub
    Public Sub UpdateAllowuptoSTRInCust(ByVal dt1 As DataTable)
        Dim clscust As New Customers.Customers.Customers
        For k = 0 To dt1.Rows.Count - 1
            If dt1.Rows(k).Item("isvalid") = "false" Then Continue For
            If dt1.Rows(k).Item("chargingtrue") = "false" And dt1.Rows(k).Item("paymentflag") = True Then Continue For
            Dim strSql As String = "update customers set allowupto = '" & CDate(df1.GetCellValue(dt1.Rows(k), "changeallowupto")).ToString("yyyy-MM-dd hh:mm:ss") & "' where rowstatus = 0 and p_customers =" & df1.GetCellValue(dt1.Rows(k), "P_customers")
            df1.SqlExecuteNonQuery(clscust.ServerDatabase, strSql)

        Next
    End Sub
    ''' <summary>
    ''' function to get data from payments for dealer after creating dealer entry
    ''' </summary>
    ''' <param name="lcondition">condition from fronend </param>
    ''' <param name="Lorder">Order by string</param>
    ''' <param name="start">starting row no</param>
    ''' <param name="pSize">No of rows to fetch</param>
    ''' <param name="DtInfoTable">Infotable</param>
    ''' <returns></returns>
    Function getPaymentDataForDealer(lcondition As String, Lorder As String, start As Integer, pSize As Integer, DtInfoTable As DataTable) As DataTable
        Dim dt As New DataTable
        Dim clscust As New Customers.Customers.Customers
        dt = df1.GetDataFromSqlFixedRows(clscust.ServerDatabase, "Payment", "*", "", lcondition, "", Lorder, start, pSize, -1)

        df1.AddColumnsInDataTable(dt, "TextAmount,TextVerifyCode,PaymentDate1,TextStatus,AvailableAmount,TextCommissionTo,Textdiscount")

        dt = df1.AddingNameForCodesPrimamryCols(dt, "PaymentMode,BenAccount", "TextPaymentMode,TextBenAccount", DtInfoTable, "NameOfInfo")

        For i = 0 To dt.Rows.Count - 1

            If IsDBNull(dt.Rows(i).Item("Amount")) = False Then
                dt.Rows(i).Item("TextAmount") = Math.Round(dt.Rows(i).Item("Amount"), 2)
            End If

            If IsDBNull(dt.Rows(i).Item("VerifyCode")) = False Then
                If dt.Rows(i).Item("VerifyCode") = "R" Then
                    dt.Rows(i).Item("TextVerifyCode") = "Rejected"
                ElseIf dt.Rows(i).Item("VerifyCode") = "P" Then
                    dt.Rows(i).Item("TextVerifyCode") = "Pending"
                ElseIf dt.Rows(i).Item("VerifyCode") = "V" Then
                    dt.Rows(i).Item("TextVerifyCode") = "Verified"
                End If
            End If
            If IsDBNull(dt.Rows(i).Item("Status")) = False Then
                If dt.Rows(i).Item("Status") = "F" Then
                    dt.Rows(i).Item("TextStatus") = "Fully Adjusted"
                ElseIf dt.Rows(i).Item("Status") = "P" Then
                    dt.Rows(i).Item("TextStatus") = "Partially Adjusted"
                ElseIf dt.Rows(i).Item("Status") = "U" Then
                    dt.Rows(i).Item("TextStatus") = "Unadjusted"
                End If
            End If
            If IsDBNull(dt.Rows(i).Item("discount")) = False Then
                dt.Rows(i).Item("Textdiscount") = Math.Round(dt.Rows(i).Item("discount"), 2)
            Else
                dt.Rows(i).Item("Textdiscount") = "N/A"
            End If
            Dim amount As Decimal = df1.GetCellValue(dt.Rows(i), "Amount")
            Dim paymentDate As Date = df1.GetCellValue(dt.Rows(i), "PaymentDate")
            dt.Rows(i).Item("PaymentDate1") = paymentDate.ToString("dd-MM-yyyy")
            Dim discount As Decimal = df1.GetCellValue(dt.Rows(i), "discount")

            Dim TempAvailableAmoount As Decimal = GetBalanceFromPaymentVoucher(dt.Rows(i).Item("P_Payment"))
            Dim AvailableAmoount As Decimal = Math.Round(((discount + amount) - TempAvailableAmoount), 2)
            dt.Rows(i).Item("AvailableAmount") = AvailableAmoount.ToString
            If IsDBNull(dt.Rows(i).Item("CommissionTo")) = False Then
                Dim Query1 As String = String.Format("Select AccName from AccMaster where P_acccode=" & dt.Rows(i).Item("CommissionTo"))
                Dim AccNamedt As DataTable = df1.SqlExecuteDataTable(clscust.ServerDatabase, Query1)
                If AccNamedt.Rows.Count > 0 Then
                    dt.Rows(i).Item("TextCommissionTo") = AccNamedt.Rows(0).Item("AccName")
                Else
                    dt.Rows(i).Item("TextCommissionTo") = "Not Available"
                End If
            Else
                dt.Rows(i).Item("TextCommissionTo") = "Not Available"
            End If
        Next
        Return dt
    End Function
    Public Function GetBalanceFromPaymentVoucher(ByVal P_payment As Integer) As Decimal
        Dim clsbillpay As New BillPayFlag.BillPayFlag.BillPayFlag
        Dim dtbillpay As New DataTable
        dtbillpay = df1.GetDataFromSql(clsbillpay.ServerDatabase, clsbillpay.TableName, "*", "", "rowstatus = 0 and p_payment = " & P_payment, "", "")
        Dim totpaid As Decimal = 0.0
        For l = 0 To dtbillpay.Rows.Count - 1
            totpaid = totpaid + df1.GetCellValue(dtbillpay.Rows(l), "amountadjusted")
        Next
        Return totpaid
    End Function
    Public Function RemovePaymentRowInChargingItem(ByVal orderheader As Integer, ByVal TotalAmt As Decimal) As Integer
        Dim clschargingItems As New ChargingItems.ChargingItems.ChargingItems
        Dim clschargingheader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtchrgingItem As DataTable = df1.GetDataFromSql(clschargingItems.ServerDatabase, clschargingItems.TableName, "*", "", "rowstatus =0 and servicecode = 3024 and orderheader = " & orderheader, "", "")
        '  Dim sumTot As Decimal = 0.0
        For l = 0 To dtchrgingItem.Rows.Count - 1
            Dim headerno As Integer = df1.GetCellValue(dtchrgingItem.Rows(l), "headerno")
            Dim amount As Decimal = df1.GetCellValue(dtchrgingItem.Rows(l), "taxableamount")
            '   sumTot = sumTot + amount
            Dim strsq As String = "update chargingheader set grandtotal = grandtotal - " & amount & " where rowstatus = 0 and headerno = " & headerno
            Dim kl As Int16 = df1.SqlExecuteNonQuery(clschargingheader.ServerDatabase, strsq)
        Next
        Dim strsq1 As String = "update orderheader set totalamount = " & TotalAmt & " where rowstatus = 0 and orderheader = " & orderheader
        Dim lq As Int16 = df1.SqlExecuteNonQuery(clschargingItems.ServerDatabase, strsq1)
        Dim strsql As String = "delete from chargingitems where rowstatus =0 and servicecode = 3024 and orderheader = " & orderheader ' "' chargingitems_key in(" & chrgingitemStr & ")"
        Dim k As Int16 = df1.SqlExecuteNonQuery(clschargingItems.ServerDatabase, strsql)
        Return k
    End Function
    Public Sub addInRegBlock(ByVal customers As DataTable, ByVal sessionrow As DataRow, ByVal NodeIncreasedHashTable As Hashtable)
        For l = 0 To customers.Rows.Count - 1
            If Not UCase(customers.Rows(l).Item("isvalid")) = UCase("true") Then Continue For
            Dim paymentFlag As Boolean = df1.GetCellValue(customers.Rows(l), "paymentflag")
            If paymentFlag = True And customers.Rows(l).Item("chargingtrue") = "false" Then Continue For

            Dim nodeincreasedStr As String = GF1.GetValueFromHashTable(NodeIncreasedHashTable, CStr(df1.GetCellValue(customers.Rows(l), "p_customers")))
            If nodeincreasedStr = "true" Then Continue For
            If customers.Rows(l).Item("regtype") = "new" And customers.Rows(l).Item("regtype2") = "home" Then Continue For

            If paymentFlag = True Then
                Dim clsRegBlock As New Regblock.Regblock.Regblock
                Dim dtregBlock As DataTable = df1.GetDataFromSql(clsRegBlock.ServerDatabase, clsRegBlock.TableName, "*", "", "rowstatus = 0 and p_customers =" & df1.GetCellValue(customers.Rows(l), "p_customers"), "", "")
                If dtregBlock.Rows.Count > 0 Then Continue For
                clsRegBlock.CurrRow.Item("P_customers") = customers.Rows(l).Item("P_customers")
                clsRegBlock.CurrRow.Item("regsenddate") = customers.Rows(l).Item("currregdate")
                clsRegBlock.CurrRow.Item("regtype") = customers.Rows(l).Item("regtype")
                clsRegBlock.CurrRow.Item("regtype2") = customers.Rows(l).Item("regtype2")
                clsRegBlock.CurrRow.Item("paymentflag") = "U"
                clsRegBlock.CurrRow.Item("mtimestamp") = df1.getDateTimeISTNow 'customers.Rows(l).Item("mtimestamp")
                clsRegBlock.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                Dim aclsobject() As Object = {clsRegBlock}
                cfc.SaveIntodb(aclsobject)
            End If
            '  End If
        Next
    End Sub
    Public Function AllowuptoProcessingLogicWithoutPayment(ByVal customers As DataTable, ByVal dealerrow As DataRow, ByVal DealerStaffRow As DataRow, ByVal newbutOld As String, ByVal custinfoMismatch As DataTable) As DataTable
        Dim dateceiling As New DateTime
        Dim daysceiling As New Integer
        Dim registrationGraceDays As Integer
        Dim allowuptoHash As New DataTable
        allowuptoHash = df1.AddColumnsInDataTable(allowuptoHash, "p_customers,allowupto", "system.int32,system.datetime")
        registrationGraceDays = df1.GetCellValue(dealerrow, "reggracedays")
        dateceiling = df1.GetCellValue(dealerrow, "dateceilling")
        daysceiling = df1.GetCellValue(dealerrow, "daysceilling")
        For k = 0 To customers.Rows.Count - 1
            Dim regtype2 As String = df1.GetCellValue(customers.Rows(k), "regtype2")
            Dim billeduptodt As DateTime = libcustomerfeature.GetBilledUpToDate(customers.Rows(k).Item("p_customers"), regtype2)
            Dim billeduptodtStr As String = billeduptodt.ToString("yyyy-MM-dd")
            If customers.Rows(k).Item("isvalid") = "true" Then
                Dim LeastDate As New List(Of DateTime)
                If customers.Rows(k).Item("chargingtrue") = "false" And customers.Rows(k).Item("paymentflag") = True Then Continue For
                Dim mL_date As DateTime = customers.Rows(k).Item("ChangeAllowUpto")
                If daysceiling > 0 Then
                    Dim mtimeSpan As TimeSpan = mL_date.Subtract(customers.Rows(k).Item("currRegDate"))
                    Dim duedays As Integer = Convert.ToInt32(mtimeSpan.TotalDays)
                    If duedays > daysceiling Then
                        Dim CurrentDate As Date = df1.getDateTimeISTNow()
                        LeastDate.Add(CurrentDate.AddDays(daysceiling))
                    End If
                End If
                If Not dateceiling.Year.ToString = "1" Then
                    Dim mtimeSpan As TimeSpan = dateceiling.Subtract(mL_date)
                    If mtimeSpan.TotalDays < 0 Then
                        LeastDate.Add(dateceiling)
                    End If
                End If
                If LeastDate.Count = 1 Then
                    customers.Rows(k).Item("ChangeAllowUptoStr") = LeastDate(0).ToString("yyyy-MM-dd")
                    customers.Rows(k).Item("ChangeAllowupto") = LeastDate(0)
                ElseIf LeastDate.Count = 2 Then
                    Dim date1 As New DateTime
                    date1 = IIf(LeastDate(0) < LeastDate(1), LeastDate(0), LeastDate(1))
                    customers.Rows(k).Item("ChangeAllowUptoStr") = date1.ToString("yyyy-MM-dd")
                    customers.Rows(k).Item("ChangeAllowupto") = date1
                End If
                If customers.Rows(k).Item("paymentflag") = True Then
                    Dim clsregblock As New Regblock.Regblock.Regblock
                    Dim curdate As DateTime = df1.getDateTimeISTNow
                    Dim regsenddate As New DateTime
                    Dim bool As Boolean = False
                    Dim regblockdt As DataTable = df1.GetDataFromSql(clsregblock.ServerDatabase, clsregblock.TableName, "regsendDate", "", "P_customers=" & customers.Rows(k).Item("P_Customers") & " And rowstatus = 0", "", "")
                    If regblockdt.Rows.Count > 0 Then
                        regsenddate = regblockdt.Rows(0).Item("regsendDate")
                        LeastDate.Add(regsenddate.AddDays(registrationGraceDays))
                    Else
                        Dim adt As DateTime = curdate   'customers.Rows(k).Item("currregdate")
                        LeastDate.Add(adt.AddDays(registrationGraceDays))
                    End If
                    If LeastDate.Count = 1 Then
                        customers.Rows(k).Item("ChangeAllowUptoStr") = LeastDate(0).ToString("yyyy-MM-dd")
                        customers.Rows(k).Item("ChangeAllowupto") = LeastDate(0)
                    ElseIf LeastDate.Count = 2 Then
                        Dim date1 As New DateTime
                        date1 = IIf(LeastDate(0) < LeastDate(1), LeastDate(0), LeastDate(1))
                        customers.Rows(k).Item("ChangeAllowUptoStr") = date1.ToString("yyyy-MM-dd")
                        customers.Rows(k).Item("ChangeAllowupto") = date1
                    ElseIf LeastDate.Count = 3 Then
                        Dim date2 As New DateTime
                        date2 = IIf(LeastDate(0) < LeastDate(1), IIf(LeastDate(0) < LeastDate(2), LeastDate(0), LeastDate(2)), IIf(LeastDate(1) < LeastDate(2), LeastDate(1), LeastDate(2)))
                        customers.Rows(k).Item("ChangeAllowUptoStr") = date2.ToString("yyyy-MM-dd")
                        customers.Rows(k).Item("ChangeAllowupto") = date2
                    End If

                    If newbutOld = "true" Then
                        Dim dtinfomismatch As DataTable = custinfoMismatch.Select("p_customers =" & customers.Rows(k).Item("p_customers")).CopyToDataTable
                        If dtinfomismatch.Rows.Count > 0 Then
                            Dim custnameINfo As String = dtinfomismatch.Rows(0).Item("custname")
                            Dim machinemismatch As String = dtinfomismatch.Rows(0).Item("machinemismatch")
                            ' Dim nodeMismatch As String = dtinfomismatch.Rows(0).Item("nodemismatch")
                            If custnameINfo = "Y" And machinemismatch = "Y" Then
                            Else
                                Dim tempdt As New DateTime
                                tempdt = customers.Rows(k).Item("currRegDate").adddays(5)

                                If tempdt < customers.Rows(k).Item("ChangeAllowupto") Then
                                    'customers.Rows(k).Item("ChangeAllowUptoStr") = tempdt.ToString("yyyy-MM-dd")
                                    'customers.Rows(k).Item("ChangeAllowupto") = tempdt
                                End If
                            End If
                        End If

                    End If

                ElseIf customers.Rows(k).Item("paymentflag") = False Then

                    Dim dtinfomismatch As DataTable = custinfoMismatch.Select("p_customers =" & customers.Rows(k).Item("p_customers")).CopyToDataTable
                    If dtinfomismatch.Rows.Count > 0 Then
                        Dim custnameINfo As String = dtinfomismatch.Rows(0).Item("custname")
                        Dim machinemismatch As String = dtinfomismatch.Rows(0).Item("machinemismatch")
                        Dim nodeMismatch As String = dtinfomismatch.Rows(0).Item("nodesmismatch")
                        If custnameINfo = "Y" And machinemismatch = "Y" And nodeMismatch = "Y" Then
                        Else
                            Dim tempdt As New DateTime
                            'tempdt = customers.Rows(k).Item("currRegDate").adddays(5)
                            'customers.Rows(k).Item("ChangeAllowUptoStr") = tempdt.ToString("yyyy-MM-dd")
                            'customers.Rows(k).Item("ChangeAllowupto") = tempdt
                        End If
                    End If

                    If customers.Rows(k).Item("ChangeAllowupto") > billeduptodt Then
                        customers.Rows(k).Item("ChangeAllowUptoStr") = billeduptodt.ToString("yyyy-MM-dd")
                        customers.Rows(k).Item("ChangeAllowupto") = billeduptodt
                    End If

                    '     If customers.Rows(k).Item("servicingagentcode") = 2 Or customers.Rows(k).Item("servicingagentcode") = 3 Then
                    If dealerrow.Item("p_dealers") = 2 Or dealerrow.Item("p_dealers") = 3 Then
                        Dim tempdt As DateTime = customers.Rows(k).Item("ChangeAllowupto")
                        Dim tempdtday As Integer = tempdt.Day
                        Dim tempdt1 As DateTime = New DateTime(tempdt.Year, tempdt.Month, 25)
                        customers.Rows(k).Item("ChangeAllowupto") = tempdt1
                        customers.Rows(k).Item("ChangeAllowUptoStr") = tempdt1.ToString("yyyy-MM-dd")
                    End If
                    'If regtype2 = "home" Then
                    '    customers.Rows(k).Item("ChangeAllowUptoStr") = billeduptodt.ToString("yyyy-MM-dd")
                    '    customers.Rows(k).Item("ChangeAllowupto") = billeduptodt
                    'End If
                End If
                ' allowuptoHash = GF1.AddItemToHashTable(allowuptoHash, CStr(),)
            End If
        Next
        Return customers
    End Function
    Public Function AllowuptoProcessingLogicWithPayment(ByVal customers As DataTable, ByVal dealerrow As DataRow, ByVal DealerStaffRow As DataRow, ByVal newbutOld As String) As DataTable
        Dim dateceiling As New DateTime
        Dim daysceiling As New Integer
        Dim registrationGraceDays As Integer
        registrationGraceDays = df1.GetCellValue(dealerrow, "reggracedays")
        dateceiling = df1.GetCellValue(dealerrow, "dateceilling")
        daysceiling = df1.GetCellValue(dealerrow, "daysceilling")
        For k = 0 To customers.Rows.Count - 1
            Dim regtype2 As String = customers.Rows(k).Item("regtype2")
            If customers.Rows(k).Item("isvalid") = "true" Then
                If customers.Rows(k).Item("chargingtrue") = "false" And customers.Rows(k).Item("paymentflag") = True Then Continue For
                Dim mL_date As DateTime = customers.Rows(k).Item("ChangeAllowupto")
                If daysceiling > 0 Then
                    Dim mtimeSpan As TimeSpan = mL_date.Subtract(customers.Rows(k).Item("currRegDate"))
                    Dim duedays As Integer = Convert.ToInt32(mtimeSpan.TotalDays)
                    If duedays > daysceiling Then
                        Dim CurrentDate As Date = df1.getDateTimeISTNow()
                        customers.Rows(k).Item("ChangeAllowUptoStr") = CurrentDate.AddDays(daysceiling).ToString("yyyy-MM-dd")
                        customers.Rows(k).Item("ChangeAllowupto") = CurrentDate.AddDays(daysceiling)

                    End If
                End If
                If Not dateceiling.Year.ToString = "1" Then
                    Dim mtimeSpan As TimeSpan = dateceiling.Subtract(mL_date)
                    If mtimeSpan.TotalDays < 0 Then
                        customers.Rows(k).Item("ChangeAllowUptoStr") = dateceiling.ToString("yyyy-MM-dd")
                        customers.Rows(k).Item("ChangeAllowupto") = dateceiling
                    End If
                End If

                Dim billedupto As New DateTime
                billedupto = libcustomerfeature.GetBilledUpToDate(customers.Rows(k).Item("p_customers"), regtype2)
                If billedupto < customers.Rows(k).Item("ChangeAllowupto") Then
                    customers.Rows(k).Item("ChangeAllowUptoStr") = billedupto.ToString("yyyy-MM-dd")
                    customers.Rows(k).Item("ChangeAllowupto") = billedupto

                End If
                '  If customers.Rows(k).Item("servicingagentcode") = 2 Or customers.Rows(k).Item("servicingagentcode") = 3 Then
                If dealerrow.Item("p_dealers") = 2 Or dealerrow.Item("p_dealers") = 3 Then
                    Dim tempdt As DateTime = customers.Rows(k).Item("ChangeAllowupto")
                    Dim tempdtday As Integer = tempdt.Day
                    Dim tempdt1 As DateTime = New DateTime(tempdt.Year, tempdt.Month, 25)
                    customers.Rows(k).Item("ChangeAllowupto") = tempdt1
                    customers.Rows(k).Item("ChangeAllowUptoStr") = tempdt1.ToString("yyyy-MM-dd")
                End If
            End If
        Next
        Return customers
    End Function
    Public Function IfHomeInstallationPresent(ByVal P_customers As Integer) As Boolean
        Dim ClsCustomers As New Customers.Customers.Customers
        Dim bln As Boolean = False
        Dim chrgingHeaderdt As DataTable = df1.SqlExecuteDataTable(ClsCustomers.ServerDatabase, "select * from chargingheader where paymentflag = 'P' and rowstatus = 0 and p_customers = " & P_customers)
        For j = 0 To chrgingHeaderdt.Rows.Count - 1
            Dim chidt As DataTable = df1.SqlExecuteDataTable(ClsCustomers.ServerDatabase, "select * from chargingitems where servicecode = 2403 and headerno =" & chrgingHeaderdt.Rows(j).Item("headerno"))
            If chidt.Rows.Count > 0 Then
                bln = True
            End If
            If chidt.Rows.Count <= 0 Then Continue For

        Next
        Return bln
    End Function
    Public Sub UpdateChargingHEaderWithRegTran(ByVal orderHEader As Integer, ByVal abctran As Hashtable)
        Dim ClsCustomers As New Customers.Customers.Customers
        Dim strsql As String = "select * from chargingheader where rowstatus = 0 and orderheader=" & orderHEader
        Dim dtch As DataTable = df1.SqlExecuteDataTable(ClsCustomers.ServerDatabase, strsql)
        For k = 0 To dtch.Rows.Count - 1
            Dim headerno As Integer = df1.GetCellValue(dtch.Rows(k), "headerno")
            Dim p_customers As Integer = df1.GetCellValue(dtch.Rows(k), "p_customers")
            Dim regtran_key As Integer = GF1.GetValueFromHashTable(abctran, p_customers)
            Dim updstr As String = "update chargingheader set regtran_key = " & regtran_key & " where rowstatus = 0 and headerno=" & headerno
            df1.SqlExecuteNonQuery(ClsCustomers.ServerDatabase, updstr)

        Next
    End Sub
    Public Function PopulateEPLHashList(ByVal dtcust As DataTable)
        Dim eplhashlist As New List(Of Hashtable)
        Dim clsvm As New ValidateMachine.ValidateClass
        For o = 0 To dtcust.Rows.Count - 1
            If dtcust.Rows(o).Item("isvalid") = "true" Then
                If dtcust.Rows(o).Item("chargingtrue") = "false" And dtcust.Rows(o).Item("paymentflag") = True Then Continue For
                Dim lPath As String = dtcust.Rows(o).Item("eplpathserver")
                Dim abc As New Hashtable
                abc = clsvm.GetHashTableFromClientReg(lPath)
                If dtcust.Rows(o).Item("ChangeNodes") = 0 Then abc = GF1.AddItemToHashTable(abc, "lan", "0") Else abc = GF1.AddItemToHashTable(abc, "lan", "1")
                abc = GF1.AddItemToHashTable(abc, "CustCode", dtcust.Rows(o).Item("ChangeCustCode"))
                If dtcust.Rows(o).Item("regtype2") = "main" Then
                    If Not CInt(dtcust.Rows(o).Item("changeNodes").ToString) = 0 Then
                        Dim knod As Integer = CInt(dtcust.Rows(o).Item("ChangeNodes").ToString) + 1
                        abc = GF1.AddItemToHashTable(abc, "nodes", knod)
                    End If
                End If
                abc = GF1.AddItemToHashTable(abc, "P_customers", dtcust.Rows(o).Item("p_customers"))
                eplhashlist.Add(abc)
            End If
        Next
        Return eplhashlist
    End Function
    Function GetPaymentMergeData(ByVal dtpayment As DataTable, ByVal dtSelectedPaymentRow As DataTable)
        Dim dt1 As New DataTable
        dt1 = df1.AddColumnsInDataTable(dt1, "Payment_key,P_Payment,RowStatus,BenAccount,PaymentDate,Amount,Discount,Status,custcode", "System.Int32,System.Int32,System.Int32,System.Int32,System.DateTime,System.Decimal,System.Decimal,System.String,System.String")
        For l = 0 To dtpayment.Rows.Count - 1
            Dim a As DataRow = dt1.NewRow
            a.Item("Payment_key") = dtpayment.Rows(l).Item("Payment_key")
            a.Item("P_Payment") = dtpayment.Rows(l).Item("P_Payment")
            a.Item("RowStatus") = dtpayment.Rows(l).Item("RowStatus")
            a.Item("BenAccount") = dtpayment.Rows(l).Item("BenAccount")
            a.Item("PaymentDate") = dtpayment.Rows(l).Item("PaymentDate")
            a.Item("Amount") = dtpayment.Rows(l).Item("Amount")
            a.Item("Status") = dtpayment.Rows(l).Item("Status")
            a.Item("custcode") = dtpayment.Rows(l).Item("custcode")
            a.Item("discount") = dtpayment.Rows(l).Item("discount")
            dt1.Rows.Add(a)
        Next
        For k = 0 To dtSelectedPaymentRow.Rows.Count - 1
            Dim a As DataRow = dt1.NewRow
            a.Item("Payment_key") = dtSelectedPaymentRow.Rows(k).Item("Payment_key")
            a.Item("P_Payment") = dtSelectedPaymentRow.Rows(k).Item("P_Payment")
            a.Item("RowStatus") = dtSelectedPaymentRow.Rows(k).Item("RowStatus")
            a.Item("BenAccount") = dtSelectedPaymentRow.Rows(k).Item("BenAccount")
            a.Item("PaymentDate") = dtSelectedPaymentRow.Rows(k).Item("PaymentDate")
            a.Item("Amount") = dtSelectedPaymentRow.Rows(k).Item("Amount")
            a.Item("Status") = dtSelectedPaymentRow.Rows(k).Item("Status")
            a.Item("custcode") = dtSelectedPaymentRow.Rows(k).Item("custcode")
            a.Item("discount") = dtSelectedPaymentRow.Rows(k).Item("discount")
            dt1.Rows.Add(a)
        Next
        Return dt1
    End Function
    Public Function getRateforDealer(ByVal RateParameters As Hashtable) As Hashtable
        Dim CLsRateTable As New RateTable.RateTable.RateTable
        Dim RateAmt As New Hashtable
        Dim dtRtTable As New DataTable
        dtRtTable = df1.GetDataFromSql(CLsRateTable.ServerDatabase, CLsRateTable.TableName, "*", "", "", "", "")
        ' dtRtTable = df1.GetDataFromSql(CLsRateTable.ServerDatabase, CLsRateTable.TableName, "*", "", "")
        ' Dim lcondition As String = ""
        Dim keyset() As String = {}

        Dim dttable As New DataTable
        Dim preceedencepara As New Hashtable
        preceedencepara = CreatePrecedenceParameters()

        Dim lcon As String = ""
        ' lcon =getStrfromprecedence(preceedencepara, RateParameters)
        Dim kl As New DataTable
        kl = df1.GetDataFromSql(CLsRateTable.ServerDatabase, CLsRateTable.TableName, "*", "", lcon, "", "")
        ' Dim lc As String = gf1.GetStringConditionFromHashTable(bhj)
        Dim adg As New DataTable
        getRTRecursivepost(preceedencepara, dtRtTable, RateParameters, lcon)

        Dim FinalDT As DataTable = df1.GetDataFromSql(CLsRateTable.ServerDatabase, CLsRateTable.TableName, "*", "", lcon, "", "")
        If FinalDT.Rows.Count = 1 Then
            RateAmt = GF1.AddItemToHashTable(RateAmt, "rate", FinalDT.Rows(0).Item("rate"))
            RateAmt = GF1.AddItemToHashTable(RateAmt, "ratemode", FinalDT.Rows(0).Item("ratemode"))
        ElseIf FinalDT.Rows.Count > 0 Then
            For l1 = 0 To FinalDT.Rows.Count - 1
                If FinalDT(l1).Item("DateOfEffect") <= df1.getDateTimeISTNow() Then
                    RateAmt = GF1.AddItemToHashTable(RateAmt, "rate", FinalDT(l1).Item("rate"))
                    RateAmt = GF1.AddItemToHashTable(RateAmt, "ratemode", FinalDT(l1).Item("ratemode"))
                End If
            Next
        End If

        Return RateAmt
    End Function
    Public Function CreatePrecedenceParameters(Optional ByVal ServiceCode As Integer = 0, Optional ByVal ProductCode As Integer = 0, Optional ByVal Dealercode As Integer = 0, Optional ByVal EmployeeCode As Integer = 0, Optional ByVal countrycode As Integer = 0, Optional ByVal statecode As Integer = 0, Optional ByVal districtcode As Integer = 0, Optional ByVal towncode As Integer = 0, Optional ByVal customerType As Integer = 0, Optional ByVal CustomerCode As Integer = 0, Optional ByVal BusinessCode As Integer = 0, Optional ByVal PromotionCode As Integer = 0) As Hashtable
        Dim RateParameters As New Hashtable
        RateParameters = GF1.AddItemToHashTable(RateParameters, "7", "EmployeeCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "5", "DealerCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "1", "CountryCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "2", "StateCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "3", "DistrictCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "4", "TownCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "6", "CustomerType")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "9", "CustomerCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "11", "ProductCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "10", "BusinessCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "12", "ServiceCode")
        RateParameters = GF1.AddItemToHashTable(RateParameters, "8", "PromoCode")
        Return RateParameters
    End Function
    Public Sub getRTRecursivepost(ByRef parameterPrecendence As Hashtable, ByRef dtRtTble As DataTable, ByRef rateparameters As Hashtable, Optional ByRef lcon As String = "")
        Dim ade As New Hashtable
        For op = 0 To parameterPrecendence.Keys.Count - 1
            Dim ghj As String = GF1.GetValueFromHashTable(parameterPrecendence, CStr(op + 1))
            If ghj <> "Nothing" Then
                'Dim op As Integer = 0
                Dim lvar As Object = GF1.GetValueFromHashTable(rateparameters, ghj)
                If CStr(lvar) <> "" Then
                    Dim hj As New DataTable
                    Dim llcon As String = LCase(lcon)
                    Dim lghj As String = LCase(ghj)
                    If llcon.Contains(lghj) Then Continue For

                    Dim lstr As String = ""
                    If llcon = "" Then
                        lstr = parameterPrecendence.Item(CStr(op + 1)) & "=" & CInt(lvar)
                    Else
                        lstr = lcon & " and " & parameterPrecendence.Item(CStr(op + 1)) & "=" & CInt(lvar)
                    End If
                    hj = df1.SearchDataTable(dtRtTble, lstr)
                    If hj.Rows.Count <= 0 Then
                        rateparameters = GF1.AddItemToHashTable(rateparameters, parameterPrecendence.Item(CStr(op + 1)), 0)
                        dtRtTble = df1.SearchDataTable(dtRtTble, parameterPrecendence.Item(CStr(op + 1)) & "=" & 0)
                        If llcon = "" Then
                            lcon = parameterPrecendence.Item(CStr(op + 1)) & "=" & 0
                        Else
                            lcon = lcon & " and " & parameterPrecendence.Item(CStr(op + 1)) & "=" & 0
                        End If
                        parameterPrecendence = GF1.AddItemToHashTable(parameterPrecendence, CStr(op + 1), "Nothing")
                        getRTRecursivepost(parameterPrecendence, dtRtTble, rateparameters, lcon)
                        Exit Sub
                    ElseIf hj.Rows.Count > 0 Then
                        If llcon = "" Then
                            lcon = parameterPrecendence.Item(CStr(op + 1)) & "=" & CInt(lvar)
                        Else
                            lcon = lcon & " and " & parameterPrecendence.Item(CStr(op + 1)) & "=" & CInt(lvar)
                        End If
                        Continue For
                        Exit Sub
                    End If
                End If
            End If
        Next
    End Sub
    Public Function CreateRateParameters(Optional ByVal ServiceCode As Integer = 0, Optional ByVal ProductCode As Integer = 0, Optional ByVal Dealercode As Integer = 0, Optional ByVal EmployeeCode As Integer = 0, Optional ByVal countrycode As Integer = 0, Optional ByVal statecode As Integer = 0, Optional ByVal districtcode As Integer = 0, Optional ByVal towncode As Integer = 0, Optional ByVal customerType As Integer = 0, Optional ByVal CustomerCode As Integer = 0, Optional ByVal BusinessCode As Integer = 0, Optional ByVal PromotionCode As Integer = 0) As Hashtable
        Dim RateParameters As New Hashtable
        RateParameters = GF1.AddItemToHashTable(RateParameters, "EmployeeCode", EmployeeCode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "DealerCode", Dealercode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "CountryCode", countrycode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "StateCode", statecode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "DistrictCode", districtcode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "TownCode", towncode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "CustomerType", customerType)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "CustomerCode", CustomerCode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "ProductCode", ProductCode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "BusinessCode", BusinessCode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "ServiceCode", ServiceCode)
        RateParameters = GF1.AddItemToHashTable(RateParameters, "PromoCode", PromotionCode)
        Return RateParameters
    End Function
    Public Function GetPaymentFlagFromChargingHeader(ByVal p_customers As Integer, ByVal GracePeriod As Integer, ByVal regtype As String) As Boolean
        Dim clsChargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim paymentFlag As Boolean = False
        Dim dt1 As DataTable = df1.GetDataFromSql(clsChargingHeader.ServerDatabase, clsChargingHeader.TableName, "*", "", "rowstatus = 0 and p_customers = " & p_customers & " and  paymentflag='P'", "", "billedupto desc")
        Dim currDate As DateTime = df1.getDateTimeISTNow
        If dt1.Rows.Count > 0 Then
            For op = 0 To dt1.Rows.Count - 1
                Dim paymentFl1 As String = df1.GetCellValue(dt1.Rows(op), "paymentflag")
                If paymentFl1 <> "P" Then Continue For
                Dim entered As Boolean = False
                Dim dtchargingitemsDt As DataTable = df1.SqlExecuteDataTable(clsChargingHeader.ServerDatabase, "select * from chargingitems where headerno =" & df1.GetCellValue(dt1.Rows(op), "headerno"))
                For o = 0 To dtchargingitemsDt.Rows.Count - 1
                    Dim srvccode As Integer = df1.GetCellValue(dtchargingitemsDt.Rows(o), "servicecode")
                    If regtype = "main" Then
                        If srvccode = 2412 Or srvccode = 2401 Or srvccode = 3019 Then
                            Dim billedUptoDt As DateTime = df1.GetCellValue(dt1.Rows(op), "BilledUpto")

                            Dim TimeSpn As TimeSpan = billedUptoDt.Subtract(currDate)
                            Dim paymentFl As String = df1.GetCellValue(dt1.Rows(op), "paymentflag")
                            If TimeSpn.Days <= GracePeriod Then
                                paymentFlag = True
                                entered = True
                                Exit For
                                'Exit For
                            Else
                                entered = True
                                paymentFlag = False
                                Exit For
                                '  Exit For
                            End If
                        End If
                    ElseIf regtype = "home" Then
                        If srvccode = 2403 Or srvccode = 2774 Then
                            Dim billedUptoDt As DateTime = df1.GetCellValue(dtchargingitemsDt.Rows(o), "chargingtodate")
                            Dim TimeSpn As TimeSpan = billedUptoDt.Subtract(currDate)
                            Dim paymentFl As String = df1.GetCellValue(dt1.Rows(op), "paymentflag")
                            If TimeSpn.Days <= GracePeriod Then
                                paymentFlag = True
                                entered = True
                                Exit For
                                '  Exit For
                            Else
                                entered = True
                                paymentFlag = False
                                Exit For
                                ' Exit For
                            End If
                        End If
                    End If
                    '   End If

                Next

                If entered = True Then Exit For


            Next
        Else
            paymentFlag = True
        End If
        Return paymentFlag
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Orderdt">Input Order datatable to which payment gateway is added</param>
    ''' <param name="AMTwithPGcharges">Amount to be paid by paymentgateway . Payment gateway charges not included</param>
    ''' <param name="paymentGatewayPercentage"></param>
    ''' <param name="GrandTotalWithTax">OrderValue including tax</param>
    ''' <param name="grandtotalWithoutTax">Ordervalue without including tax</param>
    ''' <param name="WT">Tax applicable or not</param>
    ''' <param name="AvalBal">Available Balance of main dealer of logged in user</param>
    ''' <returns></returns>
    Public Function CalculatePaymentGatewayChargesRowinOrder(ByVal Orderdt As DataTable, ByVal AMTwithPGcharges As Decimal, ByVal paymentGatewayPercentage As Decimal, ByVal GrandTotalWithTax As Decimal, ByVal grandtotalWithoutTax As Decimal, ByVal WT As String, ByVal AvalBal As Decimal) As DataTable
        If Orderdt.Rows.Count <= 0 Then
            Return Orderdt
        End If
        Dim ListOFCustCode As New List(Of String)
        Dim grandtotalval As Decimal = 0.0
        '    grandtotalval = df1.GetCellValue(Orderdt.Rows(0), "grandtotal")
        'To get list of unique custcode and grand total
        '  Dim denopayment As Decimal = 1 + paymentGatewayPercentage / 100
        '   Dim paymentgatewaycharges As Decimal = TotalAmt - (TotalAmt / denopayment)
        '  grandtotalval = grandtotalval + paymentgatewaycharges
        Dim tot As Decimal = 0.0
        If WT = "Y" Then grandtotalval = GrandTotalWithTax Else grandtotalval = grandtotalWithoutTax
        '  Dim TotalAMtWithPGNTax As Decimal = TotalAmt * 1.02 + 0.18 * TotalAmt * 0.02
        Dim paymentgatewayCharge As Decimal = AMTwithPGcharges * paymentGatewayPercentage + 0.18 * AMTwithPGcharges * paymentGatewayPercentage
        grandtotalval = grandtotalval + paymentgatewayCharge
        Dim SubTotalCustCode As Decimal = 0.0
        For i = 0 To Orderdt.Rows.Count - 1
            If df1.GetCellValue(Orderdt.Rows(i), "custcode") IsNot Nothing Then
                Dim custcode As String = Orderdt.Rows(i).Item("custcode")
                If GF1.FindIndexListOfString(ListOFCustCode, custcode) < 0 Then
                    ListOFCustCode.Add(custcode)
                End If
            End If
        Next
        '  Dim indAmt As Decimal = Math.Round(paymentgatewaycharges / ListOFCustCode.Count, MidpointRounding.AwayFromZero)
        ' Dim subtotal As Decimal = 
        ' grandtotalval = df1.GetCellValue(Orderdt.Rows(0), "grandtotal")
        'grandtotalval += 0.02 * grandtotalval
        'To compute  operations  on the other field related to the unique cust code
        For i = 0 To ListOFCustCode.Count - 1
            Dim subtotalval As Decimal = 0.0
            Dim indamt1 As Decimal = 0.0
            Dim dtcustrow = Orderdt.Select("custcode = '" & ListOFCustCode(i) & "'")
            If dtcustrow.Count > 0 Then
                indamt1 = paymentgatewayCharge * df1.GetCellValue(dtcustrow(0), "subtotal") / df1.GetCellValue(dtcustrow(0), "grandtotal")
                subtotalval = indamt1 + df1.GetCellValue(dtcustrow(0), "subtotal")
            End If
            Dim OrderRow As DataRow = Orderdt.NewRow
            If df1.GetCellValue(dtcustrow(0), "P_customers") IsNot Nothing Then
                OrderRow("P_customers") = dtcustrow(0).Item("P_customers")
            End If
            If df1.GetCellValue(dtcustrow(0), "custcode") IsNot Nothing Then
                OrderRow("custcode") = dtcustrow(0).Item("custcode")
            End If
            If df1.GetCellValue(dtcustrow(0), "CustName") IsNot Nothing Then
                OrderRow("CustName") = dtcustrow(0).Item("CustName")
            End If
            OrderRow("servicecode") = 3024
            OrderRow("TextServiceCode") = "Payment Gateway Charges"
            OrderRow("Amount") = indamt1
            OrderRow("subtotal") = subtotalval
            OrderRow("grandtotal") = grandtotalval
            OrderRow("igst") = 0
            ' insert updated subtotal and grand total in orderdt
            For k = 0 To Orderdt.Rows.Count - 1
                If ListOFCustCode(i) = Orderdt.Rows(k).Item("custcode") Then
                    Orderdt.Rows(k).Item("subtotal") = subtotalval
                    Orderdt.Rows(k).Item("grandtotal") = grandtotalval
                End If
            Next
            Orderdt.Rows.Add(OrderRow)
        Next
        Return Orderdt
    End Function
    Public Function PopulateGrandTotal(ByVal dt As DataTable) As DataTable
        Dim ListOFCustCode As New List(Of String)
        'Dim ListOFCust As New List(Of String) 
        For UI1 = 0 To dt.Rows.Count - 1
            Dim custcode As String = dt.Rows(UI1).Item("custcode")
            If GF1.FindIndexListOfString(ListOFCustCode, custcode) < 0 Then
                ListOFCustCode.Add(custcode)
            End If
        Next
        Dim k As Integer = 0
        Dim dtcustco As New DataTable
        Dim total_amount As Decimal = 0
        For lp = 0 To ListOFCustCode.Count - 1
            dtcustco = dt.Select("custcode = '" & ListOFCustCode(lp) & "'").CopyToDataTable
            Dim ListOfCust As New List(Of String)
            Dim subtot As Integer = 0
            For po = 0 To dtcustco.Rows.Count - 1
                Dim qty As Integer = df1.GetCellValue(dtcustco.Rows(po), "quantity")
                If qty <= 0 Then Continue For
                Dim subtotnew As Integer = 0
                subtotnew = dtcustco.Rows(po).Item("amount")
                total_amount += subtotnew
                subtot = subtot + subtotnew
            Next
            For i = 0 To dt.Rows.Count - 1

                If ListOFCustCode(lp) = dt.Rows(i).Item("custcode") Then
                    dt.Rows(i).Item("subtotal") = subtot
                End If
            Next
        Next
        For i = 0 To dt.Rows.Count - 1

            dt.Rows(i)("grandtotal") = total_amount
        Next
        Return dt
    End Function
    Public Function OrderPageWarningForDuplicateUnpaidChargingHeader(ByVal Customers As DataTable) As DataTable
        Customers = df1.AddColumnsInDataTable(Customers, "duplicatechheader", "system.string")
        '  Dim duplicate As DataTable = checkForDuplicateChargingHeader(Customers(k).Item("p_customers"))
        Dim custrec As New DataTable
        custrec = df1.AddColumnsInDataTable(custrec, "headerno,orderheader,p_customers,servicecode,custname,warningmsg", "system.int32,system.int32,system.int32,system.int32,system.string,system.string")
        For k = 0 To Customers.Rows.Count - 1
            Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
            If Customers.Rows(k).Item("isvalid") = "false" Then Continue For
            If Customers.Rows(k).Item("paymentflag") = False Then Continue For
            Dim regtype As String = Customers(k).Item("regtype")
            Dim regtype2 As String = Customers(k).Item("regtype2")
            If regtype = "amc" And regtype2 = "main" Then
                Dim dtchrhe As DataTable = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "paymentflag = 'U' and rowstatus=0 and p_customers = " & Customers.Rows(k).Item("p_customers"), "", "billdate desc")
                If dtchrhe.Rows.Count > 0 Then Customers.Rows(k).Item("duplicatechheader") = "true" Else Customers.Rows(k).Item("duplicatechheader") = "false"
                Dim P_chargingHeader1 As String = ""
                For t = 0 To dtchrhe.Rows.Count - 1
                    P_chargingHeader1 = P_chargingHeader1 & "," & df1.GetCellValue(dtchrhe.Rows(t), "headerno")
                Next
                If P_chargingHeader1.StartsWith(",") Then P_chargingHeader1 = P_chargingHeader1.Substring(1)
                If P_chargingHeader1 = "" Then Continue For
                Dim clschargingitems As New ChargingItems.ChargingItems.ChargingItems
                Dim dtchargingItems As DataTable = df1.GetDataFromSql(clschargingitems.ServerDatabase, clschargingitems.TableName, "*", "", "rowstatus = 0 and headerno in ( " & P_chargingHeader1 & ") and p_customers =" & Customers.Rows(k).Item("P_customers"), "", "chargingdate desc")
                Dim ListOFCust As New List(Of String)
                For lm = 0 To dtchargingItems.Rows.Count - 1
                    Dim srvccode As Integer = df1.GetCellValue(dtchargingItems.Rows(lm), "servicecode")
                    Dim P_cust As String = df1.GetCellValue(Customers.Rows(k), "P_customers")
                    If GF1.FindIndexListOfString(ListOFCust, P_cust) >= 0 Then Continue For
                    Select Case srvccode
                        Case 2402
                            custrec.Rows.Add()
                            custrec.Rows(custrec.Rows.Count - 1).Item("p_customers") = Customers.Rows(k).Item("P_customers")
                            custrec.Rows(custrec.Rows.Count - 1).Item("servicecode") = srvccode  '
                            custrec.Rows(custrec.Rows.Count - 1).Item("custname") = Customers.Rows(k).Item("custname")
                            custrec.Rows(custrec.Rows.Count - 1).Item("warningmsg") = "Already an existing unpaid order for new node license. Proceeding to create new order. Yes to Continue, No to Decline" ' Customers.Rows(k).Item("P_customers")
                            custrec.Rows(custrec.Rows.Count - 1).Item("orderheader") = dtchargingItems.Rows(lm).Item("orderheader")
                            custrec.Rows(custrec.Rows.Count - 1).Item("headerno") = dtchargingItems.Rows(lm).Item("headerno")
                            ListOFCust.Add(Customers.Rows(k).Item("P_customers"))
                        Case 2412 Or 3019
                            custrec.Rows.Add()
                            custrec.Rows(custrec.Rows.Count - 1).Item("p_customers") = Customers.Rows(k).Item("P_customers")
                            custrec.Rows(custrec.Rows.Count - 1).Item("servicecode") = srvccode  '
                            custrec.Rows(custrec.Rows.Count - 1).Item("custname") = Customers.Rows(k).Item("custname")
                            custrec.Rows(custrec.Rows.Count - 1).Item("warningmsg") = "Already an existing unpaid order for AMC of main license. Proceeding to create new order. Yes to Continue, No to Decline" ' Customers.Rows(k).Item("P_customers")
                            custrec.Rows(custrec.Rows.Count - 1).Item("orderheader") = dtchargingItems.Rows(lm).Item("orderheader")
                            custrec.Rows(custrec.Rows.Count - 1).Item("headerno") = dtchargingItems.Rows(lm).Item("headerno")
                            ListOFCust.Add(Customers.Rows(k).Item("P_customers"))
                        Case 2773
                            custrec.Rows.Add()
                            custrec.Rows(custrec.Rows.Count - 1).Item("p_customers") = Customers.Rows(k).Item("P_customers")
                            custrec.Rows(custrec.Rows.Count - 1).Item("servicecode") = srvccode  '
                            custrec.Rows(custrec.Rows.Count - 1).Item("custname") = Customers.Rows(k).Item("custname")
                            custrec.Rows(custrec.Rows.Count - 1).Item("warningmsg") = "Already an existing unpaid order for amc of existing nodes. Proceeding to create new order. Yes to Continue, No to Decline" ' Customers.Rows(k).Item("P_customers")
                            custrec.Rows(custrec.Rows.Count - 1).Item("orderheader") = dtchargingItems.Rows(lm).Item("orderheader")
                            custrec.Rows(custrec.Rows.Count - 1).Item("headerno") = dtchargingItems.Rows(lm).Item("headerno")
                            ListOFCust.Add(Customers.Rows(k).Item("P_customers"))
                    End Select
                Next
            End If
        Next
        Return custrec
    End Function
    ''' <summary>
    ''' Populates Order table corresponding to uploaded registration file as per payment status according to logged in Dealer
    ''' </summary>
    ''' <param name="Customers"></param>
    ''' <param name="DealerRow "></param>
    ''' <param name="NodeIncreasedHashTable"></param>
    ''' <returns></returns>
    Public Function ProcessCustomersDtForRate(ByVal Customers As DataTable, ByVal DealerRow As DataRow, ByVal NodeIncreasedHashTable As Hashtable) As DataTable
        Dim OrderFinal As New DataTable
        '   Dim clswebsession As New WebSessions.WebSessions.WebSessions
        '  Dim logintype As String = ""
        Dim loginkey As Integer = -1
        loginkey = df1.GetCellValue(DealerRow, "p_dealers")
        Dim currdate As DateTime = df1.getDateTimeISTNow

        OrderFinal = df1.AddColumnsInDataTable(OrderFinal, "P_customers, custcode, CustName, servicecode, Quantity, Rate, RateNarr, QuantityNarr, Amount, chargingFromDate,chargingFromDateStr, chargingToDate,chargingToDateStr, numberOfDays,subtotal,grandtotal,allowregdownload,WT", "system.int32,System.String,System.String,System.int32,System.int32,system.decimal,system.string,system.string,system.decimal,system.datetime,system.string,system.datetime,system.string,system.int16,system.decimal,system.decimal,system.string,system.string")

        For k = 0 To Customers.Rows.Count - 1
            If Not UCase(Customers.Rows(k).Item("isvalid")) = UCase("true") Then Continue For
            Dim custcode As String = Customers(k).Item("custcode")
            Dim regtype As String = Customers(k).Item("regtype")
            Dim regtype2 As String = Customers(k).Item("regtype2")
            Dim nodes As String = Customers(k).Item("nodes")
            Dim custname As String = Customers(k).Item("custname")
            Dim p_customers As Integer = Customers(k).Item("p_customers")
            Dim quantity As Integer = 0


            '  If Customers.Rows(k).Item("paymentflag") = False Then Continue For
            If LCase(regtype.Trim) = "new" And UCase(regtype2.Trim) = UCase("main") Then
                OrderFinal.Rows.Add()
                quantity = 1
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2401
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity
                ' OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("ratenarr") = "one time"
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = currdate
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = "NA"
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = currdate.AddDays(365)
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = "NA" ' currdate.AddDays(365)
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = 365 'currdate.AddDays(365)

                Dim clsRateParameters As New Hashtable
                clsRateParameters = CreateRateParameters(2401,, loginkey,,,,,,,,,)
                Dim ratehashTable As New Hashtable
                ratehashTable = getRateforDealer(clsRateParameters)
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable, "rate")
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable, "rate") * quantity

                If Customers.Rows(k).Item("changenodes") > 0 Then

                    OrderFinal.Rows.Add()
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2402
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                    quantity = df1.GetCellValue(Customers.Rows(k), "changenodes")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity
                    Dim clsRateParameters1 As New Hashtable
                    Dim ratehashTable1 As New Hashtable
                    clsRateParameters1 = CreateRateParameters(2402,, loginkey,,,,,,,,,)
                    ratehashTable1 = getRateforDealer(clsRateParameters1)
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = currdate
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = currdate.AddDays(365)
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = currdate.ToString("dd-MM-yyyy")

                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = currdate.AddDays(365).ToString("dd-MM-yyyy")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = 365 'currdate.AddDays(365)


                End If



            End If

            If LCase(regtype.Trim) = "new" And UCase(regtype2.Trim) = UCase("home") Then

                '  Dim DrchrgingItem As DataRow = checkForDuplicateChargingHeader(p_customers, 2403)
                '    If DrchrgingItem Is Nothing Then
                OrderFinal.Rows.Add()

                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2403
                quantity = 1
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity
                Dim clsRateParameters As New Hashtable
                Dim ratehashTable As New Hashtable
                clsRateParameters = CreateRateParameters(2403,, loginkey,,,,,,,,,)
                ratehashTable = getRateforDealer(clsRateParameters)
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable, "rate")
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable, "rate") * quantity
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = currdate
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = currdate.AddDays(365)
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = 365 'currdate.AddDays(365)
                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = "NA" 'currdate.ToString("dd-MM-yyyy")

                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = "NA" 'currdate.AddDays(365).ToString("dd-MM-yyyy")
                '  Else



            End If

            '  End If


            'If LCase(regtype.Trim) = "amc" And UCase(regtype2.Trim) = UCase("home") Then
            '    If df1.GetCellValue(Customers.Rows(k), "paymentflag") = False Then Continue For
            '    OrderFinal.Rows.Add()


            '    Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
            '    Dim dtChargingheader As DataTable = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "paymentflag = 'P' and rowstatus=0 and p_customers = " & p_customers, "", "billdate desc")
            '    Dim mbilledupto As New DateTime
            '    mbilledupto = GetBilledUptoDate(p_customers)



            '    Dim clschargingItem As New ChargingItems.ChargingItems.ChargingItems
            '    Dim P_chargingHeader As String = ""
            '    For t = 0 To dtChargingheader.Rows.Count - 1
            '        P_chargingHeader = P_chargingHeader & "," & df1.GetCellValue(dtChargingheader.Rows(t), "headerno")
            '    Next
            '    If P_chargingHeader.StartsWith(",") Then P_chargingHeader = P_chargingHeader.Substring(1)
            '    Dim dtchargingItems As DataTable = df1.GetDataFromSql(clschargingItem.ServerDatabase, clschargingItem.TableName, "*", "", "rowstatus = 0 and headerno in ( " & P_chargingHeader & ")", "", "chargingdate desc")
            '    Dim qtytemp As Integer = 0
            '    Dim chargingTo As New DateTime
            '    Dim mainQty As Integer = 0
            '    If currdate.AddDays(30) > mbilledupto.AddDays(365) And currdate.AddDays(30) < mbilledupto.AddDays(730) Then
            '        mainQty = 2
            '    ElseIf mbilledupto.AddDays(730) < currdate.AddDays(30) Then
            '        mainQty = 3
            '    Else
            '        mainQty = 1
            '    End If
            '    Dim clsRateParameters As New Hashtable
            '    Dim ratehashTable As New Hashtable
            '    clsRateParameters = CreateRateParameters(2774,, loginkey,,,,,,,,,)
            '    ratehashTable = getRateforDealer(clsRateParameters)
            '    Dim noofdays As Integer = mainQty * 365
            '    For u = 0 To dtchargingItems.Rows.Count - 1
            '        Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(u), "servicecode")
            '        quantity = df1.GetCellValue(dtchargingItems(u), "quantity")
            '        ' If quantity <= 0 Then Continue For
            '        If mserviceCode1 = "2403" Then
            '            chargingTo = df1.GetCellValue(dtchargingItems(u), "chargingtodate")
            '            Dim timespn As TimeSpan = chargingTo - currdate
            '            If timespn.Days >= 0 And timespn.Hours >= 0 And timespn.Minutes >= 0 Then
            '                Dim daysBilled As Integer = 365 - timespn.Days
            '                ' If Not daysBilled > 0 Then QtyTemp = QtyTemp + quantity
            '            Else

            '                qtytemp = 1 ' QtyTemp + quantity
            '            End If
            '        End If

            '    Next



            '    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
            '    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
            '    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
            '    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2774
            '    If qtytemp = 1 Then
            '        quantity = 1 * mainQty
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable, "rate") * quantity
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noofdays ' mbilledupto.AddDays(noofdays).Subtract(chargingTo).Days

            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noofdays)

            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noofdays).ToString("dd-MM-yyyy")



            '    Else


            '        Dim multifc As Decimal = mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days / 365
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable, "rate") * quantity * multifc, 2)

            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noofdays).Subtract(chargingTo).Days
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noofdays)

            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo.ToString("dd-MM-yyyy")
            '        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noofdays).ToString("dd-MM-yyyy")


            '    End If

            '    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity

            '    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable, "rate")

            'End If

            If LCase(regtype.Trim) = "amc" Then


                Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
                Dim dtChargingheader As DataTable = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "paymentflag = 'P' and rowstatus=0 and p_customers = " & p_customers, "", "billdate desc")
                Dim mbilledupto As New DateTime
                mbilledupto = libcustomerfeature.GetBilledUpToDate(p_customers, "main")
                '    If dtChargingheader.Rows.Count > 0 Then mbilledupto = df1.GetCellValue(dtChargingheader.Rows(0), "billedupto")

                Dim nodes1 As Integer = 0
                If Not DBNull.Value.Equals(Customers.Rows(k).Item("nodesDB")) Then
                    nodes1 = Customers.Rows(k).Item("nodesDB")
                End If
                Dim changenode As Integer = 0
                If Not DBNull.Value.Equals(Customers.Rows(k).Item("changenodes")) Then
                    changenode = Customers.Rows(k).Item("changenodes")
                End If
                Dim payfl As Boolean = GetPaymentFlagFromChargingHeader(p_customers, 40, "main")    'df1.GetCellValue(Customers.Rows(k), "paymentflag")
                ' If payfl = False And changenode <= nodes1 Then Continue For
                Dim nodeIncreasedStr As String = GF1.GetValueFromHashTable(NodeIncreasedHashTable, CStr(Customers.Rows(k).Item("P_customers")))
                If nodeIncreasedStr = "true" Then

                    OrderFinal.Rows.Add()
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2402
                    quantity = changenode - nodes1
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity   'changenode - nodes1
                    Dim clsRateParameters2 As New Hashtable
                    Dim ratehashTable2 As New Hashtable
                    clsRateParameters2 = CreateRateParameters(2402,, loginkey,,,,,,,,,)
                    ratehashTable2 = getRateforDealer(clsRateParameters2)
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable2, "rate")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable2, "rate") * quantity
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = currdate
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = currdate.AddDays(365)

                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = currdate.ToString("dd-MM-yyyy")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = currdate.AddDays(365).ToString("dd-MM-yyyy")


                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = 365


                End If


                Dim payflHome As Boolean = GetPaymentFlagFromChargingHeader(p_customers, 40, "home")
                If payflHome = True And payfl = False Then

                    Dim dtchargingItems As New DataTable

                    Dim homefound As Boolean = False
                    Dim clsRateParameters3 As New Hashtable
                    Dim ratehashTable3 As New Hashtable
                    Dim chargingTo1 As New DateTime
                    Dim qtytemp1 As Integer = 0
                    Dim mainQty As Integer = 0

                    clsRateParameters3 = CreateRateParameters(2774,, loginkey,,,,,,,,,)
                    ratehashTable3 = getRateforDealer(clsRateParameters3)
                    '  Dim noofdays As Integer = mainQty * 365
                    For u = 0 To dtchargingItems.Rows.Count - 1
                        Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(u), "servicecode")
                        quantity = df1.GetCellValue(dtchargingItems(u), "quantity")
                        ' If quantity <= 0 Then Continue For
                        If mserviceCode1 = "2403" Then
                            OrderFinal.Rows.Add()
                            homefound = True
                            chargingTo1 = df1.GetCellValue(dtchargingItems(u), "chargingtodate")
                            Dim timespn As TimeSpan = chargingTo1 - currdate
                            If timespn.Days >= 0 And timespn.Hours >= 0 And timespn.Minutes >= 0 Then
                                Dim daysBilled As Integer = 365 - timespn.Days
                                ' If Not daysBilled > 0 Then QtyTemp = QtyTemp + quantity
                            Else

                                qtytemp1 = 1 ' QtyTemp + quantity
                            End If
                            Exit For
                        End If

                    Next

                    If homefound = True Then
                        If currdate.AddDays(30) > chargingTo1.AddDays(365) And currdate.AddDays(30) < chargingTo1.AddDays(730) Then
                            mainQty = 2
                        ElseIf chargingTo1.AddDays(730) < currdate.AddDays(30) Then
                            mainQty = 3
                        Else
                            mainQty = 1
                        End If
                        Dim noOFDays As Integer = 365 * mainQty

                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2774
                        If qtytemp1 = 1 Then
                            quantity = 1 * mainQty
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable3, "rate") * quantity
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noOFDays ' mbilledupto.AddDays(noofdays).Subtract(chargingTo).Days

                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)

                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")



                        Else

                            quantity = 1 * mainQty
                            Dim multifc As Decimal = mbilledupto.AddDays(noOFDays).Subtract(chargingTo1).Days / 365
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable3, "rate") * quantity * multifc, 2)

                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noOFDays).Subtract(chargingTo1).Days
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo1
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)

                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo1.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")


                        End If

                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity

                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable3, "rate")


                    End If


                End If





                If payfl = True Then
                    OrderFinal.Rows.Add()



                    'If mbilledupto.AddDays(365) <> Customers.Rows(k).Item("calculatedallowupto") Then
                    '    Continue For
                    'End If
                    Dim servicecode As Integer = -1
                    Dim onsiteFlag As String = df1.GetCellValue(Customers.Rows(k), "onsiteflag")
                    If LCase(onsiteFlag) = "true" Then servicecode = 3019 Else servicecode = 2412
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = servicecode

                    Dim mainQty As Integer = 0
                    If currdate.AddDays(30) > mbilledupto.AddDays(365) And currdate.AddDays(30) < mbilledupto.AddDays(730) Then
                        mainQty = 2
                    ElseIf mbilledupto.AddDays(730) < currdate.AddDays(30) Then
                        mainQty = 3
                    Else
                        mainQty = 1
                    End If

                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = mainQty
                    Dim clsRateParameters As New Hashtable

                    clsRateParameters = CreateRateParameters(servicecode,, loginkey,,,,,,,,,)
                    Dim ratehashTable As New Hashtable
                    ratehashTable = getRateforDealer(clsRateParameters)
                    Dim mRate As Decimal = GF1.GetValueFromHashTable(ratehashTable, "rate")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = mRate
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = mRate * mainQty
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
                    Dim noOFDays As Integer = 365 * mainQty

                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)



                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noOFDays  'currdate.AddDays(365)
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                    OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("allowregdownload") = Customers.Rows(k).Item("allowregdownload")

                    Dim clschargingItem As New ChargingItems.ChargingItems.ChargingItems
                    Dim P_chargingHeader As String = ""
                    For t = 0 To dtChargingheader.Rows.Count - 1
                        P_chargingHeader = P_chargingHeader & "," & df1.GetCellValue(dtChargingheader.Rows(t), "headerno")
                    Next
                    If P_chargingHeader.StartsWith(",") Then P_chargingHeader = P_chargingHeader.Substring(1)
                    Dim dtchargingItems As DataTable = df1.GetDataFromSql(clschargingItem.ServerDatabase, clschargingItem.TableName, "*", "", "rowstatus = 0 and headerno in ( " & P_chargingHeader & ")", "", "chargingdate desc")

                    'BillingForHomeInstallation
                    Dim homefound As Boolean = False
                    Dim clsRateParameters3 As New Hashtable
                    Dim ratehashTable3 As New Hashtable
                    Dim chargingTo1 As New DateTime
                    Dim qtytemp1 As Integer = 0
                    clsRateParameters3 = CreateRateParameters(2774,, loginkey,,,,,,,,,)
                    ratehashTable3 = getRateforDealer(clsRateParameters3)
                    '  Dim noofdays As Integer = mainQty * 365
                    For u = 0 To dtchargingItems.Rows.Count - 1
                        Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(u), "servicecode")
                        quantity = df1.GetCellValue(dtchargingItems(u), "quantity")
                        ' If quantity <= 0 Then Continue For
                        If mserviceCode1 = "2403" Then
                            OrderFinal.Rows.Add()
                            homefound = True
                            chargingTo1 = df1.GetCellValue(dtchargingItems(u), "chargingtodate")
                            Dim timespn As TimeSpan = chargingTo1 - currdate
                            If timespn.Days >= 0 And timespn.Hours >= 0 And timespn.Minutes >= 0 Then
                                Dim daysBilled As Integer = 365 - timespn.Days
                                ' If Not daysBilled > 0 Then QtyTemp = QtyTemp + quantity
                            Else
                                qtytemp1 = 1 ' QtyTemp + quantity
                            End If
                            Exit For
                        End If
                    Next
                    If homefound = True Then
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2774
                        If qtytemp1 = 1 Then
                            quantity = 1 * mainQty
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable3, "rate") * quantity
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noOFDays ' mbilledupto.AddDays(noofdays).Subtract(chargingTo).Days
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                        Else
                            quantity = 1 * mainQty
                            Dim multifc As Decimal = mbilledupto.AddDays(noOFDays).Subtract(chargingTo1).Days / 365
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable3, "rate") * quantity * multifc, 2)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noOFDays).Subtract(chargingTo1).Days
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo1
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo1.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                        End If
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable3, "rate")
                    End If
                    If changenode = 0 Then Continue For
                    Dim lan As String = ""
                    If Not DBNull.Value.Equals(Customers.Rows(k).Item("changelan")) Then
                        lan = Customers.Rows(k).Item("changelan")
                    End If

                    If nodes1 > changenode Then
                        'Customer is reducing the number of nodes
                        Dim clsRateParameters1 As New Hashtable
                        Dim ratehashTable1 As New Hashtable
                        clsRateParameters1 = CreateRateParameters(2773,, loginkey,,,,,,,,,)
                        ratehashTable1 = getRateforDealer(clsRateParameters1)
                        Dim QtyTemp As Integer = 0
                        For op = 0 To dtchargingItems.Rows.Count - 1
                            Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(op), "servicecode")
                            quantity = df1.GetCellValue(dtchargingItems(op), "quantity")
                            If quantity <= 0 Then Continue For
                            If mserviceCode1 = "2402" Then
                                Dim chargingTo As DateTime = df1.GetCellValue(dtchargingItems(op), "chargingtodate")
                                Dim timespn As TimeSpan = chargingTo - currdate
                                If timespn.Days >= 0 And timespn.Hours >= 0 And timespn.Minutes >= 0 Then
                                    Dim daysBilled As Integer = 365 - timespn.Days
                                    If Not daysBilled > 0 Then QtyTemp = QtyTemp + quantity
                                Else

                                    QtyTemp = QtyTemp + quantity
                                End If
                            End If
                        Next
                        If QtyTemp >= changenode Then
                            OrderFinal.Rows.Add()
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                            quantity = changenode
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty
                            Dim mRate1 As Decimal = GF1.GetValueFromHashTable(ratehashTable1, "rate")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = mRate1
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = mRate1 * quantity
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noOFDays
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                            '   Dim Multifierct As Decimal = 365 / 365



                        Else
                            Dim qtyleft As Integer = changenode - QtyTemp
                            For op = 0 To dtchargingItems.Rows.Count - 1
                                Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(op), "servicecode")
                                If mserviceCode1 = "2402" Then
                                    Dim chargingTo As DateTime = df1.GetCellValue(dtchargingItems(op), "chargingtodate")
                                    Dim timespn As TimeSpan = chargingTo - currdate
                                    Dim daysBilled As Integer = 365 - timespn.Days
                                    If daysBilled > 0 Then
                                        If qtyleft = 0 Then Continue For
                                        OrderFinal.Rows.Add()
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                                        quantity = df1.GetCellValue(dtchargingItems(op), "quantity")
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")

                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo.ToString("dd-MM-yyyy")
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                                        Dim multifc As Decimal = mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days / 365
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity * multifc, 2)
                                        qtyleft = qtyleft - quantity
                                    End If
                                End If
                            Next

                        End If

                        If changenode - nodes1 > 0 Then
                            OrderFinal.Rows.Add()
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2402
                            quantity = changenode - nodes1
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity
                            Dim mRate2 As Decimal = GF1.GetValueFromHashTable(ratehashTable1, "rate")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = mRate2
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = mRate2 * quantity
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = currdate
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = currdate.AddDays(365)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = 365
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = currdate.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = currdate.AddDays(365).ToString("dd-MM-yyyy")
                            '   Dim Multifierct As Decimal = 365 / 365

                        End If

                    ElseIf nodes1 < changenode Then
                        'Customer is increasing the numbe rof nodes
                        If nodes1 > 0 Then
                            Dim clsRateParameters1 As New Hashtable
                            Dim ratehashTable1 As New Hashtable
                            clsRateParameters1 = CreateRateParameters(2773,, loginkey,,,,,,,,,)
                            ratehashTable1 = getRateforDealer(clsRateParameters1)
                            Dim QtyTemp As Integer = 0
                            For op = 0 To dtchargingItems.Rows.Count - 1
                                Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(op), "servicecode")
                                If mserviceCode1 = "2402" Then
                                    quantity = df1.GetCellValue(dtchargingItems(op), "quantity")
                                    If quantity <= 0 Then Continue For
                                    Dim chargingTo As DateTime = df1.GetCellValue(dtchargingItems(op), "chargingtodate")
                                    Dim timespn As TimeSpan = chargingTo - mbilledupto
                                    Dim daysBilled As Integer = 365 - timespn.Days
                                    If daysBilled > 0 And daysBilled < 365 Then
                                        'Adding row in orderfinal whose installation date is less than 365 days
                                        OrderFinal.Rows.Add()
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")

                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                                        '  OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(365) - chargingTo
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days
                                        Dim multifcd As Decimal = mbilledupto.AddDays(365).Subtract(chargingTo).Days / 365
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity * multifcd, 2)
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo.ToString("dd-MM-yyyy")
                                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")

                                    Else
                                        QtyTemp = QtyTemp + quantity
                                    End If
                                End If
                            Next
                            'adding row in orderfinal for nodes whose installation date is older than 365
                            OrderFinal.Rows.Add()
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                            quantity = QtyTemp
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noOFDays
                        End If
                        'Adding row in orderfinal for nodes license
                        OrderFinal.Rows.Add()
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2402
                        quantity = changenode - nodes1
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity   'changenode - nodes1
                        Dim clsRateParameters2 As New Hashtable
                        Dim ratehashTable2 As New Hashtable
                        clsRateParameters2 = CreateRateParameters(2402,, loginkey,,,,,,,,,)
                        ratehashTable2 = getRateforDealer(clsRateParameters2)
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable2, "rate")
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable2, "rate") * quantity
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = currdate
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = currdate.AddDays(365)
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = currdate.ToString("dd-MM-yyyy")
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = currdate.AddDays(365).ToString("dd-MM-yyyy")
                        OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = 365
                    ElseIf nodes1 = changenode Then
                        Dim clsRateParameters1 As New Hashtable
                        Dim ratehashTable1 As New Hashtable
                        clsRateParameters1 = CreateRateParameters(2773,, loginkey,,,,,,,,,)
                        ratehashTable1 = getRateforDealer(clsRateParameters1)
                        Dim QtyTemp As Integer = 0

                        Dim effecnodes As Integer = 0
                        For y = 0 To dtchargingItems.Rows.Count - 1
                            Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(y), "servicecode")
                            If mserviceCode1 = "2402" Then
                                effecnodes = effecnodes + df1.GetCellValue(dtchargingItems.Rows(y), "quantity", "integer")
                            End If

                        Next



                        Dim dteffecci As New DataTable
                        dteffecci = dtchargingItems.Clone
                        For f = 0 To dtchargingItems.Rows.Count - 1
                            Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(f), "servicecode")
                            Dim qtyu As Integer = df1.GetCellValue(dtchargingItems.Rows(f), "quantity", "integer")
                            If mserviceCode1 = "2402" And qtyu > 0 Then
                                ' dteffecci.Rows.Add()
                                dteffecci.ImportRow(dtchargingItems.Rows(f))
                                'dteffecci.Rows(dteffecci.Rows.Count - 1).ItemArray = dtchargingItems.Rows(f).ItemArray
                            End If
                        Next


                        Dim dtview As New DataView(dteffecci)
                        dtview.Sort = "chargingtodate desc"
                        dteffecci = dtview.ToTable



                        'for loop to add different rows for nodes installed on different dates and not has been completed one year
                        'For op = 0 To dtchargingItems.Rows.Count - 1

                        '    Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(op), "servicecode")
                        '    If mserviceCode1 = "2402" Then
                        '        Dim chargingTo As DateTime = df1.GetCellValue(dtchargingItems(op), "chargingtodate")
                        '        Dim timespn As TimeSpan = chargingTo.Subtract(mbilledupto)
                        '        Dim daysBilled As Integer = 365 - timespn.Days
                        '        quantity = df1.GetCellValue(dtchargingItems(op), "quantity")
                        '        If quantity <= 0 Then Continue For
                        '        If daysBilled > 0 Then
                        '            ' for nodes whose installation date is less than 365 days at amc main date
                        '            OrderFinal.Rows.Add()
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")

                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                        '            Dim multiplierFct As Decimal = (mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days) / 365
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo.ToString("dd-MM-yyyy")
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity * multiplierFct, 2)
                        '            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days
                        '        Else
                        '            '    for nodes whose installation date is equal to 365 days or more at amc main date
                        '            QtyTemp = QtyTemp + quantity
                        '        End If
                        '    End If
                        'Next



                        Dim qtysum As Integer = 0
                        For op = 0 To dteffecci.Rows.Count - 1
                            If qtysum >= effecnodes Then Exit For
                            '  Dim mserviceCode1 As String = df1.GetCellValue(dtchargingItems.Rows(op), "servicecode")
                            ' If mserviceCode1 = "2402" Then
                            Dim chargingTo As DateTime = df1.GetCellValue(dteffecci(op), "chargingtodate")
                            Dim timespn As TimeSpan = chargingTo.Subtract(mbilledupto)
                            Dim daysBilled As Integer = 365 - timespn.Days
                            quantity = df1.GetCellValue(dteffecci(op), "quantity")
                            qtysum = qtysum + quantity
                            If quantity <= 0 Then Continue For
                            If daysBilled > 0 And daysBilled < 365 Then
                                ' for nodes whose installation date is less than 365 days at amc main date
                                OrderFinal.Rows.Add()
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                                ' OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")

                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = chargingTo
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                                Dim multiplierFct As Decimal = (mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days) / 365
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty * multiplierFct
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = chargingTo.ToString("dd-MM-yyyy")
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = Math.Round(GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity * multiplierFct, 2)
                                OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = mbilledupto.AddDays(noOFDays).Subtract(chargingTo).Days
                            Else
                                '    for nodes whose installation date is equal to 365 days or more at amc main date
                                QtyTemp = QtyTemp + quantity
                            End If
                            ' End If
                        Next



                        ' for nodes whose installation date is more than or equal to  365 days at amc main date
                        If Not QtyTemp = 0 Then
                            OrderFinal.Rows.Add()
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("p_customers") = p_customers
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("custcode") = custcode
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("CustName") = custname
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("servicecode") = 2773
                            quantity = QtyTemp
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("quantity") = quantity * mainQty
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("rate") = GF1.GetValueFromHashTable(ratehashTable1, "rate")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("amount") = GF1.GetValueFromHashTable(ratehashTable1, "rate") * quantity
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDate") = mbilledupto
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDate") = mbilledupto.AddDays(noOFDays)
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingFromDateStr") = mbilledupto.ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("chargingToDateStr") = mbilledupto.AddDays(noOFDays).ToString("dd-MM-yyyy")
                            OrderFinal.Rows(OrderFinal.Rows.Count - 1).Item("numberofdays") = noOFDays
                        End If
                    End If
                End If
            End If
        Next
        Dim dt As DataTable = PopulateGrandTotal(OrderFinal)
        '  Dim dt1 As DataTable = AddPaymentGatewayRowinOrder(OrderFinal)
        Return dt
    End Function
    Public Function getAccCodefromLogincodetype(ByVal logintype As String, ByVal logincode As Integer) As Integer

        Dim clsCustomers As New Customers.Customers.Customers
        Dim P_acccode As Int16 = 0
        Dim dt As New DataTable
        If UCase(logintype) = "D" Then
            dt = df1.SqlExecuteDataTable(clsCustomers.ServerDatabase, "Select  P_acccode from Dealers where P_Dealers = " & logincode & " and Rowstatus=0")
        ElseIf UCase(logintype) = "E" Then
            dt = df1.SqlExecuteDataTable(clsCustomers.ServerDatabase, "Select  P_acccode from Employees where P_Employees = " & logincode & " and rowstatus=0")
        End If

        If dt.Rows.Count > 0 Then
            If IsDBNull(dt.Rows(0).Item("P_acccode")) = False Then
                P_acccode = dt.Rows(0).Item("P_acccode")
            End If
        End If
        Return P_acccode
    End Function
    ''' <summary>
    ''' Processes order populating orderheader , chargingheader, chargingitems tables 
    ''' </summary>
    ''' <param name="GrandTotalAmt"> total ordervalue inclusive/exclusive tax</param>
    ''' <param name="customersDt">Datatable containing details of customers</param>
    ''' <param name="OrderFinal">Datatable containing details of Order lines</param>
    ''' <param name="sessionRow">Datarow containing login user details</param>
    ''' <param name="paymentby">Payment done by user or customer</param>
    ''' <param name="custrec1">In case of duplicate order, containing details related to previous order </param>
    ''' <param name="WT">if Tax is applicable or not</param>
    ''' <returns></returns>
    Public Function ProcessOrder(ByVal GrandTotalAmt As Decimal, ByVal customersDt As DataTable, ByVal OrderFinal As DataTable, ByVal sessionRow As DataRow, ByVal paymentby As String, ByVal custrec1 As DataTable, ByVal WT As String, ByVal nodeIncreasedHashTable As Hashtable, ByVal DealerRow As DataRow) As Integer
        ' If chargingFlag = False Then Exit Function
        Dim chargingTrue As String = "false"
        For lp = 0 To customersDt.Rows.Count - 1
            Dim chargingtruflg As String = df1.GetCellValue(customersDt.Rows(lp), "chargingtrue")
            If LCase(chargingtruflg) = LCase("true") Then
                chargingTrue = "true"
                Exit For
            End If
        Next
        '   Dim DealerLoginKey As Integer = -1
        If chargingTrue = "false" Then
            Return -1
            Exit Function
        End If

        '  DealerLoginKey = df1.GetCellValue(DealerRow, "p_dealers")
        Dim clsOrderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim currdate As New DateTime
        currdate = df1.getDateTimeISTNow
        ' Dim logintype As String = sessionRow("linktype")
        Dim logincode1 As Integer = sessionRow("linkcode")
        Dim lorderheader As Integer = -1
        Dim P_acccode As Integer = df1.GetCellValue(DealerRow, "p_acccode") 'getAccCodefromLogincodetype("D", DealerLoginKey)
        Dim custSTR As String = ""
        '  Dim p_customers As Integer = -1
        If paymentby = "C" Then
            custSTR = customersDt.Rows(0).Item("p_customers")
        End If


        For rt = 0 To customersDt.Rows.Count - 1
            custSTR = custSTR & "," & df1.GetCellValue(customersDt.Rows(rt), "p_customers")
        Next
        If custSTR.First = "," Then custSTR = custSTR.Substring(1)

        Dim webSessions_key As Integer = sessionRow("webSessions_key")
        clsOrderHeader.CurrRow.Item("p_customers") = custSTR
        clsOrderHeader.CurrRow.Item("orderdate") = df1.getDateTimeISTNow
        clsOrderHeader.CurrRow.Item("logintype") = "D"
        clsOrderHeader.CurrRow.Item("LoginCode") = logincode1
        clsOrderHeader.CurrRow.Item("P_acccode") = P_acccode
        clsOrderHeader.CurrRow.Item("OrderSeries") = "D" & P_acccode
        clsOrderHeader.CurrRow.Item("paymentflag") = "U"
        clsOrderHeader.CurrRow.Item("totalamount") = GrandTotalAmt ' IIf(OrderFinal.Rows.Count <= 0, 0, OrderFinal.Rows(0).Item("grandtotal"))
        clsOrderHeader.CurrRow.Item("calledFrom") = "D" & paymentby
        clsOrderHeader.CurrRow.Item("mtimestamp") = currdate
        clsOrderHeader.CurrRow.Item("WT") = WT
        clsOrderHeader.CurrRow.Item("websessions_key") = webSessions_key
        Dim aClsObject() As Object = {clsOrderHeader}
        Dim Success As Boolean = True
        Dim mserverdb As String = df1.GetServerMDFForTransanction(aClsObject)
        Dim mytrans As SqlTransaction = df1.BeginTransaction(mserverdb)
        Dim aLastKeysValues As New Hashtable
        aClsObject = df1.SetKeyValueIfNewInsert(mytrans, aClsObject)
        Dim HashPublicValues As New Hashtable
        Dim sqlexec As Boolean = df1.CheckTableClassUpdations(aClsObject)
        aClsObject = df1.LastKeysPlus(mytrans, aClsObject, aLastKeysValues)
        aClsObject = df1.SetFinalFieldsValues(aClsObject, HashPublicValues)
        Dim aam As Integer = -1
        Try
            If sqlexec = True Then
                aam = df1.InsertUpdateDeleteSqlTables(mytrans, aClsObject, aam)
                Dim OrderHeaderhash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValues, "OrderHeader")
                lorderheader = GF1.GetValueFromHashTable(OrderHeaderhash, "p_orderheader")
                mytrans.Commit()
                Success = True
            End If
        Catch ex As Exception
            mytrans.Rollback()
            Success = False
        End Try
        mytrans.Dispose()
        If lorderheader < 0 And aam <= 0 Then Exit Function
        Dim OrderGrndTotal As Decimal = 0.0
        For p = 0 To customersDt.Rows.Count - 1
            If df1.GetCellValue(customersDt.Rows(p), "chargingTrue") = "true" Then
                Dim lChargingHeader As Integer = -1
                Dim mbilledupto As New DateTime
                Dim OrdercustCo As New DataTable
                Dim regtype As String = ""
                Dim regtyp2 As String = ""
                regtyp2 = customersDt.Rows(p).Item("regtype2")
                '   Dim custohx() As DataRow = custrec1.Select("p_customers=" & df1.GetCellValue(customersDt.Rows(p), "P_customers"))
                OrdercustCo = OrderFinal.Select("custcode ='" & customersDt.Rows(p).Item("custcode") & "'").CopyToDataTable
                Dim nodes As Integer = 0
                Dim changenodes As Integer = 0
                nodes = customersDt.Rows(p).Item("nodesdB")
                changenodes = customersDt.Rows(p).Item("changenodes")
                Dim chargingHeaderDt As New DataTable
                Dim clschrg As New ChargingHeader.ChargingHeader.ChargingHeader
                chargingHeaderDt = df1.GetDataFromSql(clschrg.ServerDatabase, clschrg.TableName, "billedupto", "", "paymentflag = 'P' and rowstatus = 0 and p_customers=" & df1.GetCellValue(customersDt.Rows(p), "P_customers"), "", "billdate desc")
                mbilledupto = libcustomerfeature.GetBilledUpToDate(customersDt.Rows(p).Item("p_customers"), "main")
                regtype = df1.GetCellValue(customersDt.Rows(p), "regtype")
                regtyp2 = df1.GetCellValue(customersDt.Rows(p), "regtype2")
                Dim clsChargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
                Dim dtchargingheader As New DataTable
                '
                ' create a new chargingheader if user does not chose to overwrite existing chargingheader. will be applicable in both cases if there is a charging header or not
                Dim clschargingHeader1 As New ChargingHeader.ChargingHeader.ChargingHeader
                clschargingHeader1.CurrRow.Item("OrderHeader") = lorderheader
                If chargingHeaderDt.Rows.Count > 0 Then
                    '  mbilledupto = chargingHeaderDt.Rows(0).Item("billedupto")
                    If regtype = "new" And regtyp2 = "home" Then
                        ' clschargingHeader1.CurrRow.Item("billedupto") =
                        mbilledupto = currdate.AddDays(365)
                        clschargingHeader1.CurrRow.Item("billedupto") = mbilledupto
                    ElseIf regtype = "amc" Then
                        Dim nodeincreaseStr As String = GF1.GetValueFromHashTable(nodeIncreasedHashTable, CStr(customersDt.Rows(p).Item("P_customers")))
                        If nodeincreaseStr = "true" Then
                            clschargingHeader1.CurrRow.Item("billedupto") = mbilledupto
                        Else
                            Dim mbilledupto1 As New DateTime
                            '   mbilledupto1 = mbilledupto.AddDays(365)

                            For o = 0 To OrdercustCo.Rows.Count - 1
                                Dim srvcCode As Integer = df1.GetCellValue(OrdercustCo.Rows(o), "servicecode")
                                If srvcCode = 2412 Or srvcCode = 3019 Then
                                    mbilledupto1 = df1.GetCellValue(OrdercustCo.Rows(o), "chargingtodate")
                                    Exit For
                                End If
                            Next

                            clschargingHeader1.CurrRow.Item("billedupto") = mbilledupto1

                            'clschargingHeader1.CurrRow.Item("billedupto") = calculatedallowUpto

                        End If

                    End If
                Else
                    mbilledupto = currdate.AddDays(365)
                    clschargingHeader1.CurrRow.Item("billedupto") = mbilledupto
                End If
                clschargingHeader1.CurrRow.Item("billdate") = currdate
                clschargingHeader1.CurrRow.Item("billtype") = "I"
                '  clschargingHeader1.CurrRow.Item("regtran_key") = GF1.GetValueFromHashTable(regtranHashTable, df1.GetCellValue(customersDt.Rows(p), "p_customers"))
                clschargingHeader1.CurrRow.Item("BillSeries") = "D" ' getAccCodefromLogincodetype(logintype, logincode)
                clschargingHeader1.CurrRow.Item("P_customers") = df1.GetCellValue(customersDt.Rows(p), "p_customers")
                clschargingHeader1.CurrRow.Item("p_acccode") = P_acccode
                clschargingHeader1.CurrRow.Item("logintype") = "D" 'dtCustomerOrderHeader.Rows(i).Item("P_customers")
                clschargingHeader1.CurrRow.Item("loginCode") = logincode1
                clschargingHeader1.CurrRow.Item("PaymentFlag") = "U"
                clschargingHeader1.CurrRow.Item("Grandtotal") = OrdercustCo.Rows(0).Item("subtotal")
                clschargingHeader1.CurrRow.Item("Websessions_key") = sessionRow("websessions_key")
                clschargingHeader1.CurrRow.Item("mtimestamp") = currdate
                clschargingHeader1.CurrRow.Item("roundoffamt") = Math.Ceiling(df1.GetCellValue(OrdercustCo.Rows(0), "subtotal"))
                If WT = "Y" Then
                    clschargingHeader1.CurrRow.Item("wt") = "Y"
                    ' clschargingHeader1.CurrRow.Item("Grandtotal") = 1.18 * OrdercustCo.Rows(0).Item("subtotal")
                    'clschargingHeader1.CurrRow.Item("roundoffamt") = Math.Ceiling(1.18 * df1.GetCellValue(OrdercustCo.Rows(0), "subtotal"))
                ElseIf WT = "N" Then
                    ' clschargingHeader1.CurrRow.Item("Grandtotal") = OrdercustCo.Rows(0).Item("subtotal")
                    'clschargingHeader1.CurrRow.Item("roundoffamt") = Math.Ceiling(df1.GetCellValue(OrdercustCo.Rows(0), "subtotal"))
                    clschargingHeader1.CurrRow.Item("wt") = "N"
                End If


                OrderGrndTotal = OrderGrndTotal + clschargingHeader1.CurrRow.Item("roundoffamt")

                Dim aClsObject1() As Object = {clschargingHeader1}
                Dim mserverdb1 As String = df1.GetServerMDFForTransanction(aClsObject1)
                Dim mytrans1 As SqlTransaction = df1.BeginTransaction(mserverdb1)
                Dim aLastKeysValues1 As New Hashtable
                aClsObject1 = df1.SetKeyValueIfNewInsert(mytrans1, aClsObject1)
                Dim HashPublicValues1 As New Hashtable
                Dim sqlexec1 As Boolean = df1.CheckTableClassUpdations(aClsObject1)
                aClsObject1 = df1.LastKeysPlus(mytrans1, aClsObject1, aLastKeysValues1)
                aClsObject1 = df1.SetFinalFieldsValues(aClsObject1, HashPublicValues1)
                Try
                    If sqlexec1 = True Then
                        aam = df1.InsertUpdateDeleteSqlTables(mytrans1, aClsObject1, aam)
                        Dim ChargingHeaderhash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValues1, "Chargingheader")
                        lChargingHeader = GF1.GetValueFromHashTable(ChargingHeaderhash, "headerno")
                        mytrans1.Commit()
                    End If
                Catch ex As Exception
                    mytrans1.Rollback()
                End Try
                mytrans1.Dispose()

                If lChargingHeader = -1 And aam <= 0 Then Continue For

                Dim nodeIsReduced As Boolean = False
                ' If LCase(df1.GetCellValue(customersDt.Rows(p), "overwriteorder")) = "false" Then
                If regtype = "amc" Then
                    If changenodes < nodes Then

                        nodeIsReduced = True

                        Dim am1 As New ChargingItems.ChargingItems.ChargingItems
                        am1.CurrRow.Item("headerno") = lChargingHeader
                        am1.CurrRow.Item("ItemSno") = 1
                        am1.CurrRow.Item("orderheader") = lorderheader
                        am1.CurrRow.Item("chargingdate") = currdate
                        am1.CurrRow.Item("logintype") = "D"
                        am1.CurrRow.Item("logincode") = logincode1
                        am1.CurrRow.Item("p_acccode") = P_acccode
                        am1.CurrRow.Item("p_customers") = df1.GetCellValue(customersDt.Rows(p), "p_customers")
                        am1.CurrRow.Item("productcode") = customersDt.Rows(p).Item("ProductCode")
                        am1.CurrRow.Item("servicecode") = 2402
                        am1.CurrRow.Item("chargingfromdate") = currdate
                        am1.CurrRow.Item("chargingtodate") = currdate.AddDays(365)
                        am1.CurrRow.Item("quantity") = changenodes - nodes
                        am1.CurrRow.Item("mtimestamp") = currdate
                        am1.CurrRow.Item("WebSessions_Key") = sessionRow("websessions_key")
                        Dim aclsobj1() As Object = {am1}
                        cfc.SaveIntodb(aclsobj1)
                    End If
                End If
                For l = 0 To OrdercustCo.Rows.Count - 1
                    Dim clschargingItems As New ChargingItems.ChargingItems.ChargingItems
                    clschargingItems.CurrRow.Item("headerno") = lChargingHeader
                    If nodeIsReduced = True Then
                        clschargingItems.CurrRow.Item("ItemSno") = l + 2
                    Else
                        clschargingItems.CurrRow.Item("ItemSno") = l + 1
                    End If
                    clschargingItems.CurrRow.Item("orderheader") = lorderheader
                    clschargingItems.CurrRow.Item("chargingdate") = currdate
                    clschargingItems.CurrRow.Item("logintype") = "D"
                    clschargingItems.CurrRow.Item("logincode") = logincode1
                    clschargingItems.CurrRow.Item("p_acccode") = P_acccode
                    clschargingItems.CurrRow.Item("p_customers") = df1.GetCellValue(customersDt.Rows(p), "p_customers")
                    clschargingItems.CurrRow.Item("productcode") = customersDt.Rows(p).Item("ProductCode")
                    clschargingItems.CurrRow.Item("servicecode") = OrdercustCo.Rows(l).Item("Servicecode")
                    If regtype = "new" And regtyp2 = "main" Then
                        clschargingItems.CurrRow.Item("chargingfromdate") = currdate
                        clschargingItems.CurrRow.Item("chargingtodate") = currdate.AddDays(365)
                    ElseIf regtype = "new" And regtyp2 = "home" Then
                        clschargingItems.CurrRow.Item("chargingfromdate") = currdate
                        clschargingItems.CurrRow.Item("chargingtodate") = currdate.AddDays(365)

                    ElseIf regtype = "amc" Then
                        Select Case df1.GetCellValue(OrdercustCo.Rows(l), "servicecode")
                            Case "2402"
                                clschargingItems.CurrRow.Item("chargingfromdate") = currdate
                                clschargingItems.CurrRow.Item("chargingtodate") = currdate.AddDays(365)
                            Case "2412", "3019"
                                Dim nodeincreaseStr As String = GF1.GetValueFromHashTable(nodeIncreasedHashTable, CStr(customersDt.Rows(p).Item("P_customers")))
                                If nodeincreaseStr = "true" Then
                                    clschargingItems.CurrRow.Item("chargingfromdate") = currdate
                                    clschargingItems.CurrRow.Item("chargingtodate") = currdate.AddDays(365)
                                Else
                                    clschargingItems.CurrRow.Item("chargingfromdate") = OrdercustCo.Rows(l).Item("chargingfromdate")
                                    clschargingItems.CurrRow.Item("chargingtodate") = OrdercustCo.Rows(l).Item("chargingtodate")
                                End If
                            Case "2773"
                                clschargingItems.CurrRow.Item("chargingfromdate") = OrdercustCo.Rows(l).Item("chargingfromdate")
                                clschargingItems.CurrRow.Item("chargingtodate") = OrdercustCo.Rows(l).Item("chargingtodate")
                            Case "2774"
                                clschargingItems.CurrRow.Item("chargingfromdate") = OrdercustCo.Rows(l).Item("chargingfromdate")
                                clschargingItems.CurrRow.Item("chargingtodate") = OrdercustCo.Rows(l).Item("chargingtodate")
                        End Select
                    End If
                    clschargingItems.CurrRow.Item("quantity") = OrdercustCo.Rows(l).Item("quantity")
                    clschargingItems.CurrRow.Item("productrate") = OrdercustCo.Rows(l).Item("rate")
                    clschargingItems.CurrRow.Item("mtimestamp") = currdate
                    clschargingItems.CurrRow.Item("WebSessions_Key") = sessionRow("websessions_key")

                    If WT = "Y" Then
                        clschargingItems.CurrRow.Item("baseamount") = OrdercustCo.Rows(l).Item("amount") - OrdercustCo.Rows(l).Item("igst")
                        clschargingItems.CurrRow.Item("taxableamount") = OrdercustCo.Rows(l).Item("amount") - OrdercustCo.Rows(l).Item("igst")

                        clschargingItems.CurrRow.Item("igst") = OrdercustCo.Rows(l).Item("igst")
                        clschargingItems.CurrRow.Item("Totalamt") = OrdercustCo.Rows(l).Item("amount")
                    ElseIf WT = "N" Then
                        clschargingItems.CurrRow.Item("igst") = 0 '0.18 * OrdercustCo.Rows(l).Item("amount")
                        clschargingItems.CurrRow.Item("Totalamt") = OrdercustCo.Rows(l).Item("amount")

                        clschargingItems.CurrRow.Item("baseamount") = OrdercustCo.Rows(l).Item("amount")
                        clschargingItems.CurrRow.Item("taxableamount") = OrdercustCo.Rows(l).Item("amount")
                    End If
                    Dim aclsObject2() As Object = {clschargingItems}
                    cfc.SaveIntodb(aclsObject2)
                Next
            End If
        Next
        If GrandTotalAmt <> OrderGrndTotal Then upDateGrndTotalInOrderHeader(lorderheader, OrderGrndTotal)
        Return lorderheader
    End Function
    Public Sub upDateGrndTotalInOrderHeader(ByVal orderheader As Integer, ByVal grandTotal As Decimal)
        Dim strSQL As String = "update orderheader set totalamount =" & grandTotal & " where rowstatus = 0 and p_orderheader = " & orderheader
        Dim clsorderHdr As New OrderHeader.OrderHeader.OrderHeader
        Dim Lbool As Boolean = df1.SqlExecuteNonQuery(clsorderHdr.ServerDatabase, strSQL)
    End Sub

    Public Function GetOpenedUptoDate(ByVal p_customers As Integer) As DateTime
        Dim OpenedUpto As New DateTime
        Dim clsRegTran As New RegistrationTran.RegistrationTran.RegistrationTran
        Dim Query As String = String.Format("select Top(1) Openedupto from RegistrationTran where P_Customers=" & p_customers & " and Rowstatus=0 Order By RegistrationTran_Key desc")
        Dim regTranDt As DataTable = df1.SqlExecuteDataTable(clsRegTran.ServerDatabase, Query)
        If regTranDt.Rows.Count > 0 Then
            OpenedUpto = regTranDt.Rows(0).Item("OpenedUpto")
        End If
        Return OpenedUpto
    End Function

    Public Function updateOrderAsPerTax(ByVal orderDt As DataTable, ByVal WT As String, ByVal Taxpercentage As Decimal) As DataTable

        If WT = "N" Then Taxpercentage = 0 'Return orderDt
        For k = 0 To orderDt.Rows.Count - 1
            orderDt.Rows(k).Item("igst") = Taxpercentage * orderDt.Rows(k).Item("amount")
            orderDt.Rows(k).Item("amount") = (1 + Taxpercentage) * orderDt.Rows(k).Item("amount")
            orderDt.Rows(k).Item("subtotal") = (1 + Taxpercentage) * orderDt.Rows(k).Item("subtotal")
            orderDt.Rows(k).Item("grandtotal") = (1 + Taxpercentage) * orderDt.Rows(k).Item("grandtotal")
        Next
        Return orderDt
    End Function
    ''' <summary>
    '''  logic to update billpay flag table as per FIFO logic and user selected payment voucher
    ''' </summary>
    ''' <param name="amtPG"> total amount to </param>
    ''' <param name="orderheader"></param>
    ''' <param name="sessionrow"></param>
    ''' <param name="p_acccode"></param>
    ''' <returns></returns>
    Public Function updateBillPayFlag(ByVal amtPG As Decimal, ByVal orderheader As Integer, ByVal sessionrow As DataRow, ByVal p_acccode As Integer, ByVal DealerRow As DataRow, ByVal WT As String, Optional dtPayment As DataTable = Nothing, Optional ByVal benaccount As Integer = -1) As Hashtable
        Dim dtpay As New DataTable
        Dim dtpayCopy As New DataTable
        Dim dtpayCopy1 As New DataTable
        Dim currDate As DateTime = df1.getDateTimeISTNow
        Dim CustTransactionStatus As New Hashtable
        Dim logintype As String = sessionrow("linktype")
        Dim logincode As Integer = sessionrow("linkcode")
        Dim dtchrgin As New DataTable
        Dim clschargingItem As New ChargingHeader.ChargingHeader.ChargingHeader
        dtchrgin = df1.GetDataFromSql(clschargingItem.ServerDatabase, clschargingItem.TableName, "headerno,grandtotal,p_customers", "", "rowstatus = 0 and orderheader = " & orderheader, "", "")
        dtchrgin = df1.AddColumnsInDataTable(dtchrgin, "adjamt", "system.decimal")
        Dim orderval As Decimal = 0
        For j = 0 To dtchrgin.Rows.Count - 1
            orderval = df1.GetCellValue(dtchrgin.Rows(j), "grandtotal") + orderval
        Next
        If dtPayment.Rows.Count = 0 Then
            Dim clsPay As New Payment.Payment.Payment
            Dim benaccountcondition As String = ""
            Select Case WT
                Case "Y"
                    If benaccount = -1 Then
                        benaccountcondition = " benaccount in ( 2751,3033) "
                    Else
                        benaccountcondition = " benaccount in ( 2751,3033) "
                    End If
                Case "N"
                    If benaccount = -1 Then
                        benaccountcondition = " benaccount <> 2751 and benaccount <> 3033 "
                        ' End If
                    Else
                        benaccountcondition = " benaccount = " & benaccount & "  or benaccount = 3031 "
                    End If
            End Select
            Select Case WT
                Case "Y"
                    dtpay = df1.GetDataFromSql(clsPay.ServerDatabase, clsPay.TableName, "Payment_key,P_payment,amount,status,rowstatus,benaccount,paymentdate,custcode,discount", "", "rowstatus = 0 and " & benaccountcondition & " and paymentdate >= '2019-07-01 00:00:00.000' and verifycode = 'V' and status <>'" & "F" & "' and p_acccode = " & p_acccode & "", "", "paymentdate")
                Case "N"
                    dtpay = df1.GetDataFromSql(clsPay.ServerDatabase, clsPay.TableName, "Payment_key,P_payment,amount,status,rowstatus,benaccount,paymentdate,custcode,discount", "", "rowstatus = 0 and " & benaccountcondition & " and paymentdate >= '2019-07-01 00:00:00.000' and verifycode = 'V' and status <>'" & "F" & "' and p_acccode = " & p_acccode & "", "", "paymentdate")
            End Select
            dtpay = df1.AddColumnsInDataTable(dtpay, "balance,editflag", "system.decimal,system.string")
            dtpayCopy = dtpay.Copy
        Else
            dtpay = dtPayment.Copy
            dtpay = df1.AddColumnsInDataTable(dtpay, "balance,editflag", "system.decimal,system.string")
            dtpayCopy = dtpay.Copy
        End If
        Dim customerwiseTracking As String = DealerRow.Item("customerwisepaymenttrack")
        Select Case customerwiseTracking
            Case "Y"
                For k = 0 To dtchrgin.Rows.Count - 1
                    Dim headerno As Integer = df1.GetCellValue(dtchrgin.Rows(k), "headerno")
                    Dim grdtot As Decimal = df1.GetCellValue(dtchrgin.Rows(k), "grandtotal")
                    dtchrgin.Rows(k).Item("adjamt") = grdtot
                    Dim mWT As String = df1.GetCellValue(dtchrgin.Rows(k), "WT")
                    Dim custcode As String = getcustcodeFromP_customers(dtchrgin.Rows(k).Item("p_customers"))
                    Dim dtPayselectcust1() As DataRow = dtpay.Select("custcode ='" & custcode & "'")
                    If dtPayselectcust1.Count = 0 Then Continue For
                    Dim dtPayselectcust As DataTable = dtPayselectcust1.CopyToDataTable
                    For e = 0 To dtPayselectcust.Rows.Count - 1 ' While grdtot <= 0
                        If df1.GetCellValue(dtPayselectcust.Rows(e), "editflag") = "C" Then Continue For
                        Dim benact As Int16 = df1.GetCellValue(dtPayselectcust.Rows(e), "BenAccount") '  2751
                        Dim clsbillPay As New BillPayFlag.BillPayFlag.BillPayFlag
                        '   Dim clsPay1 As New Payment.Payment.Payment
                        Dim sttus As String = df1.GetCellValue(dtPayselectcust.Rows(e), "status")
                        Dim p_pmt As Integer = df1.GetCellValue(dtPayselectcust.Rows(e), "p_payment")
                        Dim amt As Decimal = df1.GetCellValue(dtPayselectcust.Rows(e), "amount") + df1.GetCellValue(dtPayselectcust.Rows(e), "discount")
                        Dim amt1 As Decimal = 0.0
                        If sttus = "U" Then
                            amt1 = amt
                        ElseIf sttus = "P" Then
                            amt1 = amt - GetBalanceFromPaymentVoucher(p_pmt)
                        End If
                        '    clsPay1.PrevRow = df1.UpdateDataRows(clsPay1.PrevRow, dtpayCopy(e))
                        If df1.GetCellValue(dtPayselectcust.Rows(e), "editflag") = "P" Then amt1 = df1.GetCellValue(dtPayselectcust.Rows(e), "balance")
                        grdtot = grdtot - amt1
                        If grdtot = 0 Then
                            dtPayselectcust.Rows(e).Item("editflag") = "C"
                            clsbillPay.CurrRow.Item("p_payment") = df1.GetCellValue(dtPayselectcust.Rows(e), "p_payment")
                            clsbillPay.CurrRow.Item("headerno") = headerno 'df1.GetCellValue(dtpay.Rows(e), "headerno")
                            clsbillPay.CurrRow.Item("amountadjusted") = amt1
                            dtchrgin.Rows(k).Item("adjamt") = dtchrgin.Rows(k).Item("adjamt") - amt1
                            clsbillPay.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            clsbillPay.CurrRow.Item("mtimestamp") = currDate
                            clsbillPay.CurrRow.Item("logincode") = logincode
                            clsbillPay.CurrRow.Item("logintype") = logintype
                            dtPayselectcust.Rows(e).Item("status") = "F"

                            Dim clsobj1() As Object = {clsbillPay}
                            cfc.SaveIntodb(clsobj1)
                            Exit For
                            '  Continue For
                        ElseIf grdtot > 0 Then
                            dtPayselectcust.Rows(e).Item("editflag") = "C"
                            clsbillPay.CurrRow.Item("p_payment") = df1.GetCellValue(dtPayselectcust.Rows(e), "p_payment")
                            clsbillPay.CurrRow.Item("headerno") = headerno ' df1.GetCellValue(dtpay.Rows(e), "headerno")
                            clsbillPay.CurrRow.Item("amountadjusted") = amt1
                            dtchrgin.Rows(k).Item("adjamt") = dtchrgin.Rows(k).Item("adjamt") - amt1
                            clsbillPay.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            clsbillPay.CurrRow.Item("logincode") = logincode
                            clsbillPay.CurrRow.Item("logintype") = logintype
                            clsbillPay.CurrRow.Item("mtimestamp") = currDate
                            dtPayselectcust.Rows(e).Item("status") = "F"
                            Dim clsobj1() As Object = {clsbillPay}
                            cfc.SaveIntodb(clsobj1)
                            '   Continue For
                        ElseIf grdtot < 0 Then
                            dtPayselectcust.Rows(e).Item("editflag") = "P"   ' P for partial
                            dtPayselectcust.Rows(e).Item("balance") = amt1 - dtchrgin.Rows(k).Item("adjamt")
                            clsbillPay.CurrRow.Item("p_payment") = df1.GetCellValue(dtPayselectcust.Rows(e), "p_payment")
                            clsbillPay.CurrRow.Item("headerno") = headerno ' df1.GetCellValue(dtpay.Rows(e), "headerno")
                            clsbillPay.CurrRow.Item("amountadjusted") = dtchrgin.Rows(k).Item("adjamt")  'df1.GetCellValue(dtpay.Rows(cnt1), "amt1")
                            clsbillPay.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            clsbillPay.CurrRow.Item("mtimestamp") = currDate
                            clsbillPay.CurrRow.Item("logincode") = logincode
                            clsbillPay.CurrRow.Item("logintype") = logintype
                            dtPayselectcust.Rows(e).Item("status") = "P"
                            dtchrgin.Rows(k).Item("adjamt") = 0
                            Dim clsobj1() As Object = {clsbillPay}
                            cfc.SaveIntodb(clsobj1)
                            Exit For
                            '  Continue For
                        End If
                    Next
                    For u = 0 To dtPayselectcust.Rows.Count - 1
                        Dim edtflg As String = df1.GetCellValue(dtPayselectcust.Rows(u), "editflag")
                        If UCase(edtflg) = "C" Or UCase(edtflg) = "P" Then
                            Dim clsPay1 As New Payment.Payment.Payment
                            Dim abc As DataTable = df1.GetDataFromSql(clsPay1.ServerDatabase, clsPay1.TableName, "*", "", "rowstatus = 0 and p_payment = " & dtPayselectcust.Rows(u).Item("p_payment"), "", "")
                            clsPay1.PrevRow = df1.UpdateDataRows(clsPay1.PrevRow, abc.Rows(0))
                            '  clsPay1.PrevRow = df1.UpdateDataRows(clsPay1.PrevRow, dtpayCopy(u))
                            'clsPay1.CurrRow.Item("status") = edtflg  'dtpay.Rows(u).Item("status")
                            clsPay1.CurrRow.Item("status") = dtPayselectcust.Rows(u).Item("status")
                            clsPay1.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            Dim clsob1() As Object = {clsPay1}
                            cfc.SaveIntodb(clsob1)
                        End If
                    Next
                    Dim txnstatus As Boolean = False
                    If dtchrgin.Rows(k).Item("adjamt") = 0 Then txnstatus = True
                    CustTransactionStatus = GF1.AddItemToHashTable(CustTransactionStatus, CStr(dtchrgin.Rows(k).Item("p_customers")), txnstatus)
                Next
            Case Else
                For k = 0 To dtchrgin.Rows.Count - 1
                    Dim headerno As Integer = df1.GetCellValue(dtchrgin.Rows(k), "headerno")
                    Dim grdtot As Decimal = df1.GetCellValue(dtchrgin.Rows(k), "grandtotal")
                    dtchrgin.Rows(k).Item("adjamt") = grdtot
                    Dim mWT As String = df1.GetCellValue(dtchrgin.Rows(k), "WT")
                    For e = 0 To dtpay.Rows.Count - 1 ' While grdtot <= 0
                        If df1.GetCellValue(dtpay.Rows(e), "editflag") = "C" Then Continue For
                        Dim benact As Int16 = df1.GetCellValue(dtpay.Rows(e), "BenAccount") '  2751
                        'If benact = 2751 Then
                        '    If mWT = "N" Then Continue For
                        'End If
                        'If mWT = "Y" Then
                        '    If benact <> 2751 Then Continue For
                        'End If
                        Dim clsbillPay As New BillPayFlag.BillPayFlag.BillPayFlag
                        Dim clsPay1 As New Payment.Payment.Payment
                        Dim sttus As String = df1.GetCellValue(dtpay.Rows(e), "status")
                        Dim p_pmt As Integer = df1.GetCellValue(dtpay.Rows(e), "p_payment")
                        Dim amt As Decimal = df1.GetCellValue(dtpay.Rows(e), "amount") + df1.GetCellValue(dtpay.Rows(e), "discount")
                        Dim amt1 As Decimal = 0.0
                        If sttus = "U" Then
                            amt1 = amt
                        ElseIf sttus = "P" Then
                            amt1 = amt - GetBalanceFromPaymentVoucher(p_pmt)
                        End If
                        clsPay1.PrevRow = df1.UpdateDataRows(clsPay1.PrevRow, dtpayCopy(e))
                        If df1.GetCellValue(dtpay.Rows(e), "editflag") = "P" Then amt1 = df1.GetCellValue(dtpay.Rows(e), "balance")
                        grdtot = grdtot - amt1
                        If grdtot = 0 Then
                            dtpay.Rows(e).Item("editflag") = "C"
                            clsbillPay.CurrRow.Item("p_payment") = df1.GetCellValue(dtpay.Rows(e), "p_payment")
                            clsbillPay.CurrRow.Item("headerno") = headerno 'df1.GetCellValue(dtpay.Rows(e), "headerno")
                            clsbillPay.CurrRow.Item("amountadjusted") = amt1
                            dtchrgin.Rows(k).Item("adjamt") = dtchrgin.Rows(k).Item("adjamt") - amt1
                            clsbillPay.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            clsbillPay.CurrRow.Item("mtimestamp") = currDate
                            clsbillPay.CurrRow.Item("logincode") = logincode
                            clsbillPay.CurrRow.Item("logintype") = logintype
                            dtpay.Rows(e).Item("status") = "F"
                            Dim clsobj1() As Object = {clsbillPay}
                            cfc.SaveIntodb(clsobj1)
                            Exit For
                            Continue For
                        ElseIf grdtot > 0 Then
                            dtpay.Rows(e).Item("editflag") = "C"
                            clsbillPay.CurrRow.Item("p_payment") = df1.GetCellValue(dtpay.Rows(e), "p_payment")
                            clsbillPay.CurrRow.Item("headerno") = headerno ' df1.GetCellValue(dtpay.Rows(e), "headerno")
                            clsbillPay.CurrRow.Item("amountadjusted") = amt1
                            dtchrgin.Rows(k).Item("adjamt") = dtchrgin.Rows(k).Item("adjamt") - amt1
                            clsbillPay.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            clsbillPay.CurrRow.Item("mtimestamp") = currDate
                            clsbillPay.CurrRow.Item("logincode") = logincode
                            clsbillPay.CurrRow.Item("logintype") = logintype
                            dtpay.Rows(e).Item("status") = "F"
                            Dim clsobj1() As Object = {clsbillPay}
                            cfc.SaveIntodb(clsobj1)
                            Continue For
                        ElseIf grdtot < 0 Then
                            dtpay.Rows(e).Item("editflag") = "P"   ' P for partial
                            dtpay.Rows(e).Item("balance") = amt1 - dtchrgin.Rows(k).Item("adjamt")
                            clsbillPay.CurrRow.Item("p_payment") = df1.GetCellValue(dtpay.Rows(e), "p_payment")
                            clsbillPay.CurrRow.Item("headerno") = headerno ' df1.GetCellValue(dtpay.Rows(e), "headerno")
                            clsbillPay.CurrRow.Item("amountadjusted") = dtchrgin.Rows(k).Item("adjamt")  'df1.GetCellValue(dtpay.Rows(cnt1), "amt1")
                            clsbillPay.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                            clsbillPay.CurrRow.Item("mtimestamp") = currDate
                            clsbillPay.CurrRow.Item("logincode") = logincode
                            clsbillPay.CurrRow.Item("logintype") = logintype
                            dtpay.Rows(e).Item("status") = "P"
                            dtchrgin.Rows(k).Item("adjamt") = 0
                            Dim clsobj1() As Object = {clsbillPay}
                            cfc.SaveIntodb(clsobj1)
                            Exit For
                            Continue For
                        End If
                    Next
                Next
                For u = 0 To dtchrgin.Rows.Count - 1
                    Dim txnstatus As Boolean = False
                    If dtchrgin.Rows(u).Item("adjamt") = 0 Then txnstatus = True
                    CustTransactionStatus = GF1.AddItemToHashTable(CustTransactionStatus, CStr(dtchrgin.Rows(u).Item("p_customers")), txnstatus)
                Next
                For u = 0 To dtpayCopy.Rows.Count - 1
                    Dim edtflg As String = df1.GetCellValue(dtpay.Rows(u), "editflag")
                    If UCase(edtflg) = "C" Or UCase(edtflg) = "P" Then
                        Dim clsPay1 As New Payment.Payment.Payment
                        Dim abc As DataTable = df1.GetDataFromSql(clsPay1.ServerDatabase, clsPay1.TableName, "*", "", "rowstatus = 0 and p_payment = " & dtpayCopy.Rows(u).Item("p_payment"), "", "")
                        clsPay1.PrevRow = df1.UpdateDataRows(clsPay1.PrevRow, abc.Rows(0))
                        '  clsPay1.PrevRow = df1.UpdateDataRows(clsPay1.PrevRow, dtpayCopy(u))
                        'clsPay1.CurrRow.Item("status") = edtflg  'dtpay.Rows(u).Item("status")
                        clsPay1.CurrRow.Item("status") = dtpay.Rows(u).Item("status")
                        clsPay1.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                        Dim clsob1() As Object = {clsPay1}
                        cfc.SaveIntodb(clsob1)
                    End If
                Next
        End Select
        Return CustTransactionStatus
    End Function
    Public Function getcustcodeFromP_customers(ByVal p_customers As Integer) As String
        Dim ClsCustomers As New Customers.Customers.Customers
        Dim custcode As String = ""
        Dim abcsql As String = "select custcode from customers where p_customers =" & p_customers
        Dim dt As DataTable = df1.SqlExecuteDataTable(ClsCustomers.ServerDatabase, abcsql)
        If dt.Rows.Count > 0 Then
            custcode = df1.GetCellValue(dt.Rows(0), "custcode")
        End If
        Return custcode
    End Function
    ''' <summary>
    ''' To get available balance for a dealer of a particular acccode 
    ''' </summary>
    ''' <param name="p_acccode"></param>
    ''' <param name="WT"> With Tax yes or No</param>
    ''' <returns></returns>
    Public Function getLedgerBalance(ByVal p_acccode As Integer, ByVal WT As String, Optional ByVal CustomerWiseTracking As String = "N", Optional ByVal CustomersDt As DataTable = Nothing, Optional ByVal benaccount As Integer = -1)
        Dim clspaym As New Payment.Payment.Payment
        Dim dtpayment As New DataTable
        Dim custcodeStr As String = "", p_customersStr As String = ""
        If CustomerWiseTracking = "Y" Then
            For l = 0 To CustomersDt.Rows.Count - 1
                custcodeStr = custcodeStr & "," & "'" & df1.GetCellValue(CustomersDt.Rows(l), "custcode") & "'"
                p_customersStr = p_customersStr & "," & df1.GetCellValue(CustomersDt.Rows(l), "p_customers")
            Next
            If custcodeStr.First = "," Then custcodeStr = custcodeStr.Substring(1)
            If p_customersStr.First = "," Then p_customersStr = p_customersStr.Substring(1)
        End If
        Dim benaccountCondition As String = ""
        Select Case WT
            Case "Y"
                If benaccount = -1 Then
                    benaccountCondition = " benaccount in ( 2751,3033) "
                Else
                    benaccountCondition = " benaccount in ( 2751,3033) "
                End If
            Case "N"
                If benaccount = -1 Then
                    benaccountCondition = " benaccount <> 2751 and benaccount <> 3033 "
                Else
                    benaccountCondition = " benaccount = " & benaccount & "  or benaccount = 3031 "
                End If
        End Select
        Select Case WT
            Case "Y"
                If CustomerWiseTracking = "Y" Then
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 And verifycode = 'V' and " & benaccountCondition & " and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & " and custcode in (" & custcodeStr & ")", "", "")
                Else
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 and verifycode = 'V' and " & benaccountCondition & " and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode, "", "")
                End If
            Case "N"
                If CustomerWiseTracking = "Y" Then
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 and verifycode = 'V' and " & benaccountCondition & "  and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & " and custcode in (" & custcodeStr & ")", "", "")
                Else
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 And verifycode = 'V' and " & benaccountCondition & " and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode, "", "")
                End If
        End Select
        Dim totalP As Decimal = 0
        For k = 0 To dtpayment.Rows.Count - 1
            totalP = totalP + df1.GetCellValue(dtpayment.Rows(k), "amount") + df1.GetCellValue(dtpayment.Rows(k), "discount")
        Next
        Dim TotalC As Decimal = 0
        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtChargingHeader As New DataTable
        Select Case WT
            Case "Y"
                Select Case CustomerWiseTracking
                    Case "Y"
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'Y' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & " and p_customers in (" & p_customersStr & ")", "", "")
                    Case Else
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'Y' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode, "", "")
                End Select
            Case "N"
                Select Case CustomerWiseTracking
                    Case "Y"
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'N' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & " and p_customers in (" & p_customersStr & ")", "", "")
                    Case Else
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'N' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode, "", "")
                End Select
        End Select
        For k = 0 To dtChargingHeader.Rows.Count - 1
            TotalC = TotalC + df1.GetCellValue(dtChargingHeader.Rows(k), "grandtotal")
        Next
        Return IIf(totalP - TotalC >= 0, totalP - TotalC, 0)
    End Function
    Public Function GetLedgerBalanceOnDate(ByVal p_acccode As Integer, ByVal avlbldt As DateTime, ByVal WT As String, Optional ByVal customerwisetracking As String = "N", Optional customersdt As DataTable = Nothing)
        customerwisetracking = "N"
        Dim clspaym As New Payment.Payment.Payment
        Dim dtpayment As New DataTable
        Dim custcodeStr As String = "", p_customersStr As String = ""
        If customerwisetracking = "Y" Then
            For l = 0 To customersdt.Rows.Count - 1
                custcodeStr = custcodeStr & ", " & "'" & df1.GetCellValue(customersdt.Rows(l), "custcode") & "'"
                p_customersStr = p_customersStr & "," & df1.GetCellValue(customersdt.Rows(l), "p_customers")
            Next
            If custcodeStr.First = "," Then custcodeStr = custcodeStr.Substring(1)
            If p_customersStr.First = "," Then p_customersStr = p_customersStr.Substring(1)
        End If
        Select Case Trim(WT)
            Case "Y"
                If customerwisetracking = "Y" Then
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 and verifycode = 'V' and benaccount in (2751,3033) and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and verifydate < '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'" & " and custcode in (" & custcodeStr & ")", "", "")
                Else
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 and verifycode = 'V' and benaccount in (2751,3033) and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and verifydate < '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'", "", "")
                End If
            Case "N"
                If customerwisetracking = "Y" Then
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, " *", "", " rowstatus= 0 and verifycode='V' and benaccount <> 2751 and benaccount <> 3033 and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and verifydate < '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'" & " and custcode in (" & custcodeStr & ")", "", "")
                Else
                    dtpayment = df1.GetDataFromSql(clspaym.ServerDatabase, clspaym.TableName, "*", "", "rowstatus = 0 And verifycode = 'V' and benaccount <> 2751 and benaccount <> 3033 and paymentdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and verifydate < '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'", "", "")
                End If
        End Select
        Dim totalP As Decimal = 0
        For k = 0 To dtpayment.Rows.Count - 1
            totalP = totalP + df1.GetCellValue(dtpayment.Rows(k), "amount") + df1.GetCellValue(dtpayment.Rows(k), "discount")
        Next
        Dim TotalC As Decimal = 0
        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtChargingHeader As New DataTable
        Select Case WT
            Case "Y"
                Select Case customerwisetracking
                    Case "Y"
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, " *", "", " rowstatus= 0 and paymentflag='P' and WT='Y' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and billdate <='" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'" & " p_customers in (" & p_customersStr & ")", "", "")
                    Case Else
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'Y' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and billdate <= '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'", "", "")
                End Select
            Case "N"
                Select Case customerwisetracking
                    Case "Y"
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'N' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and billdate <= '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'" & "p_customers in (" & p_customersStr & ")", "", "")
                    Case Else
                        dtChargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus = 0 and paymentflag = 'P' and WT = 'N' and billdate >= '2019-07-01 00:00:00.000' and p_acccode = " & p_acccode & "  and billdate <= '" & avlbldt.ToString("yyyy-MM-dd H:mm:ss") & "'", "", "")
                End Select
        End Select
        For k = 0 To dtChargingHeader.Rows.Count - 1
            TotalC = TotalC + df1.GetCellValue(dtChargingHeader.Rows(k), "grandtotal")
        Next
        Return IIf(totalP - TotalC >= 0, totalP - TotalC, 0)
    End Function

    ''' <summary>
    ''' Function creates the Email structure i.e,Subject,Body title etc. from datatable and sends the email to the desired email id.
    ''' </summary>
    ''' <param name="RegistrationDetails">Registrationdetails datatable from which email is to be prepared</param>
    ''' <param name="fromDate"></param>
    ''' <returns></returns>
    Public Function EmailFormatOfMasterRegistrationDetails(RegistrationDetails As DataTable, fromDate As Date) As String
        Dim b As String = ""
        b = "<html><head></head><body><h2 align='center'>Details of Registation Opened On " & fromDate.ToString("dd-MM-yyyy") & "</h2>"
        If RegistrationDetails.Rows.Count <= 0 Then
            b += "<p>No Registration opened today</p>"
        Else
            b += "<table style='width:100%;'><thead><tr><th align='left'>S no</th><th align='left'>CustomerName</th><th align='left'>CustomerCode</th><th align='left'>RegSendDate</th><th align='left'>RegType</th><th align='left'>RegType2</th><th align='left'>OpenedUpto</th><th align='left'>Lan</th><th align='left'>Node</th><th align='left'>Payment Status</th><th align='left'>Amount</th><th align='left'>HomeTown</th><th align='left'>Dealer</th><th align='left'>OpenedBy</th>"
            b += "</tr></thead><tbody>"
            For j = 0 To RegistrationDetails.Rows.Count - 1
                b += "<tr style='border-bottom: 1px solid black;'>"
                b += "<td style='width:5%;'>" & j + 1 & "</td>"
                b += "<td style='width:25%;'>" & RegistrationDetails.Rows(j).Item("CustName").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("CustCode").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("TxtRegSendDate").ToString.Substring(1) & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("RegType").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("RegType2").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("TxtOpenedUpto").substring(1) & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("Lan").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("Node").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("PaymentStatus").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("HomeTown").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("Amount").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("Dealer").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("OpenedBy").ToString & "</td>"
                b += "</tr>"
            Next
            b += "</tbody></table></body></html>"
        End If
        Return b
    End Function

    ''' <summary>
    '''  Prepares the Payment Verified Today details datatable for all dealers to be sent through email or excel.
    ''' </summary>
    ''' <param name="fromDate"></param>
    ''' <param name="Todate"></param>
    ''' <returns></returns>
    Public Function PayVerifiedMasterReport(fromDate As Date, Todate As Date, Dtinfotable As DataTable) As DataTable
        If fromDate = Nothing Then fromDate = df1.getDateTimeISTNow
        If Todate = Nothing Then Todate = df1.getDateTimeISTNow
        Dim FromdateStr As String = fromDate.ToString("yyyy-MM-dd 00:00:00.00")
        Dim TodateStr As String = Todate.ToString("yyyy-MM-dd 23:59:59.999")
        Dim text As String = "Payment Verifed Function started on :" & fromDate.ToString("dd-MM-yyyy")
        cfc.WriteToFile(text, "C:\DailyReportLog.txt")
        Dim Query As String = String.Format("select P_Payment,CustCode,Proceedings,P_acccode,PaymentDate,Amount,Discount,PaymentMode,Status from payment where VerifyCode='V' and Rowstatus=0 and VerifyDate between '" & FromdateStr & "' and '" & TodateStr & "'
    Order By
    Payment_Key desc")
        Dim ClsPayment As New Payment.Payment.Payment
        Dim PaymentDetails As DataTable = df1.SqlExecuteDataTable(ClsPayment.ServerDatabase, Query)
        PaymentDetails = df1.AddColumnsInDataTable(PaymentDetails, "S.no",,, "P_Payment")
        PaymentDetails = df1.AddingNameForCodesPrimamryCols(PaymentDetails, "PaymentMode", "TxtPaymentMode", Dtinfotable, "NameOfInfo")
        PaymentDetails = df1.AddColumnsInDataTable(PaymentDetails, "Dealer,  Txtdiscount, TxtAmount, TxtPaymentDate, TxtStatus", "System.String, System.String, System.String, System.String, System.String")

        If PaymentDetails.Rows.Count > 0 Then
            Dim text1 As String = "PaymentVerifiedRows on :" & fromDate.ToString("dd-MM-yyyy") & PaymentDetails.Rows.Count
            cfc.WriteToFile(text1, "C:\DailyReportLog.txt")
            For i = 0 To PaymentDetails.Rows.Count - 1
                PaymentDetails.Rows(i).Item("S.no") = i + 1

                Dim mDealer As DataRow = libSaralAuth.getAccMasterRowForp_acccode(df1.GetCellValue(PaymentDetails.Rows(i), "P_acccode", "integer"))
                If Not mDealer Is Nothing Then
                    PaymentDetails.Rows(i).Item("Dealer") = mDealer("AccName").ToString.Trim
                End If

                If IsDBNull(PaymentDetails.Rows(i).Item("discount")) = False Then
                    PaymentDetails.Rows(i).Item("Txtdiscount") = Math.Round(PaymentDetails.Rows(i).Item("discount"), 2)
                Else
                    PaymentDetails.Rows(i).Item("Txtdiscount") = "N/A"
                End If
                If IsDBNull(PaymentDetails.Rows(i).Item("Amount")) = False Then
                    PaymentDetails.Rows(i).Item("TxtAmount") = Math.Round(PaymentDetails.Rows(i).Item("Amount"), 2)
                End If
                If IsDBNull(PaymentDetails.Rows(i).Item("Status")) = False Then
                    If PaymentDetails.Rows(i).Item("Status") = "F" Then
                        PaymentDetails.Rows(i).Item("TxtStatus") = "Fully Adjusted"
                    ElseIf PaymentDetails.Rows(i).Item("Status") = "P" Then
                        PaymentDetails.Rows(i).Item("TxtStatus") = "Partially Adjusted"
                    ElseIf PaymentDetails.Rows(i).Item("Status") = "U" Then
                        PaymentDetails.Rows(i).Item("TxtStatus") = "Unadjusted"
                    End If
                End If
                Dim paymentDate As Date = df1.GetCellValue(PaymentDetails.Rows(i), "PaymentDate")
                PaymentDetails.Rows(i).Item("TxtPaymentDate") = "'" & paymentDate.ToString("yyyy-MM-dd")
            Next

            PaymentDetails = df1.AlterDataTable(PaymentDetails, "", "PaymentMode,Amount,PaymentDate,Status,discount,P_acccode")
        Else
            Dim text2 As String = "no of entries: 0"
            cfc.WriteToFile(text2, "C:\DailyReportLog.txt")
        End If
        Dim text3 As String = "total rows returned: " & PaymentDetails.Rows.Count
        cfc.WriteToFile(text3, "C:\DailyReportLog.txt")
        Return PaymentDetails
    End Function


    ''' <summary>
    ''' Function creates the Email structure i.e,Subject,Body title etc. from datatable and sends the email to the desired email id.
    ''' </summary>
    ''' <param name="Paymentdt"> datatable from which email is to be prepared</param>
    ''' <param name="fromDate"></param>
    ''' <returns></returns>
    Public Function EmailFormatOfPaymentVerifiedDetails(Paymentdt As DataTable, fromDate As Date) As String
        Dim b As String = ""
        b = "<html><head></head><body><h2 align='center'>Verified Payment On: " & fromDate.ToString("dd-MM-yyyy") & "</h2>"
        If Paymentdt.Rows.Count <= 0 Then
            b += "<p>No Payment Verified today</p>"
        Else
            b += "<table style='width:100%;'><thead><tr><th align='left'>S no</th><th align='left'>PaymentId</th><th align='left'>CustCode</th><th align='left'>Proceedings</th><th align='left'>Date</th><th align='left'>Amount</th><th align='left'>Discount</th><th align='left'>PaymentMode</th><th align='left'>Status</th><th align='left'>DealerName</th>"
            b += "</tr></thead><tbody>"
            For j = 0 To Paymentdt.Rows.Count - 1
                b += "<tr style='border-bottom: 1px solid black;'>"
                b += "<td style='width:5%;'>" & j + 1 & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("P_Payment").ToString & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("CustCode").ToString & "</td>"
                b += "<td style='width:25%;'>" & Paymentdt.Rows(j).Item("Proceedings").ToString & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("TxtPaymentDate").ToString.Substring(1) & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("TxtAmount").ToString & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("Txtdiscount").ToString & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("TxtPaymentMode").ToString & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("TxtStatus").ToString & "</td>"
                b += "<td style='width:10%;'>" & Paymentdt.Rows(j).Item("Dealer").ToString & "</td>"
                b += "</tr>"
            Next
            b += "</tbody></table></body></html>"
        End If
        Return b
    End Function


    ''' <summary>
    '''  Prepares the CustomerRegistration details datatable for all dealers to be sent through email or excel.
    ''' </summary>
    ''' <param name="fromDate"></param>
    ''' <param name="Todate"></param>
    ''' <returns></returns>
    Public Function RegCreatedMasterReport(fromDate As Date, Todate As Date) As DataTable
        If fromDate = Nothing Then fromDate = df1.getDateTimeISTNow
        If Todate = Nothing Then Todate = df1.getDateTimeISTNow
        Dim FromdateStr As String = fromDate.ToString("MM-dd-yyyy 00:00:00.00")
        Dim TodateStr As String = Todate.ToString("MM-dd-yyyy 23:59:59.999")

        Dim Query As String = String.Format("Select Customers.CustCode,Customers.CustName,Customers.ServicingAgentCode,Customers.WebSessions_key,WebSessions.Linktype,WebSessions.Linkcode,InfoTable.NameOfInfo as HomeTown,RegistrationTran.RegSendDate,RegistrationTran.RegType,RegistrationTran.RegType2,RegistrationTran.OpenedUpto,RegistrationTran.Lan,RegistrationTran.Node,RegistrationTran.P_customers
    from
        RegistrationTran RegistrationTran
            inner join
        Customers Customers
            on Customers.P_Customers  = RegistrationTran.P_customers
  inner join
        WebSessions WebSessions
            on Customers.WebSessions_key  = WebSessions.WebSessions_key
      inner join
        InfoTable InfoTable
            on Customers.HomeTown = InfoTable.P_InfoTable      
    where
    RegistrationTran.RowStatus=0 and Customers.RowStatus=0 and (RegistrationTran.RegsendDate between '" & FromdateStr & "' and '" & TodateStr & "')
    Order By
    RegistrationTran.RegsendDate asc")
        Dim ClsChargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim RegistrationDetails As DataTable = df1.SqlExecuteDataTable(ClsChargingHeader.ServerDatabase, Query)
        RegistrationDetails = df1.AddColumnsInDataTable(RegistrationDetails, "S.no",,, "CustCode")
        RegistrationDetails = df1.AddColumnsInDataTable(RegistrationDetails, "TxtOpenedUpto,TxtRegSendDate,PaymentStatus,Amount,OpenedBy,Dealer", "System.String,System.String,System.String,System.String,System.String,System.String")
        If RegistrationDetails.Rows.Count > 0 Then
            For i = 0 To RegistrationDetails.Rows.Count - 1
                RegistrationDetails.Rows(i).Item("S.no") = i + 1
                Dim tempdtop As DateTime = RegistrationDetails.Rows(i).Item("openedupto")
                Dim tempdtopstr As String = "'" & tempdtop.ToString("yyyy-MM-dd")
                RegistrationDetails.Rows(i).Item("TxtOpenedUpto") = tempdtopstr
                Dim tempdtRegSendDate As DateTime = RegistrationDetails.Rows(i).Item("regSendDate")
                Dim tempRegSendDatstr As String = "'" & tempdtRegSendDate.ToString("yyyy-MM-dd")
                RegistrationDetails.Rows(i).Item("TxtRegSendDate") = tempRegSendDatstr
                Dim chdt As DataTable = df1.GetDataFromSql(ClsChargingHeader.ServerDatabase, ClsChargingHeader.TableName, "*", "", "RowStatus=0 and P_Customers=" & RegistrationDetails.Rows(i).Item("P_customers") & " and (BillDate between '" & FromdateStr & "' and '" & TodateStr & "')", "", "")
                If chdt.Rows.Count > 0 Then
                    If chdt.Rows(0).Item("PaymentFlag").ToString.Trim = "P" Then
                        RegistrationDetails.Rows(i).Item("PaymentStatus") = "Paid"
                    Else
                        RegistrationDetails.Rows(i).Item("PaymentStatus") = "UnPaid"
                    End If
                    RegistrationDetails.Rows(i).Item("Amount") = Math.Round(chdt.Rows(0).Item("GrandTotal"), 2)
                Else
                    RegistrationDetails.Rows(i).Item("PaymentStatus") = "N/A"
                End If
                Dim mDealer1 As DataRow = libSaralAuth.getAccMasterRowForp_acccode(df1.GetCellValue(RegistrationDetails.Rows(i), "ServicingAgentCode", "integer"))
                If Not mDealer1 Is Nothing Then
                    RegistrationDetails.Rows(i).Item("Dealer") = mDealer1("AccName").ToString.Trim
                End If
                If df1.GetCellValue(RegistrationDetails.Rows(i), "Linktype") = "U" Then

                    Dim empRow As DataRow = libSaralAuth.getUserLoginRowFromUserLoginKey(df1.GetCellValue(RegistrationDetails.Rows(i), "LinkCode", "integer"))
                    If Not empRow Is Nothing Then
                        RegistrationDetails.Rows(i).Item("OpenedBy") = empRow("Name").ToString.Trim
                    End If
                Else
                    Dim mDealer As DataRow = libSaralAuth.getAccMasterRowForp_acccode(df1.GetCellValue(RegistrationDetails.Rows(i), "LinkCode", "integer"))
                    If Not mDealer Is Nothing Then
                        RegistrationDetails.Rows(i).Item("OpenedBy") = mDealer("AccName").ToString.Trim
                    End If
                End If
            Next
            RegistrationDetails = df1.AlterDataTable(RegistrationDetails, "", "P_customers,ServicingAgentCode,WebSessions_key,linktype,linkcode,OpenedUpto,regSendDate")
        End If
        Return RegistrationDetails
    End Function



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="mClname"></param>
    ''' <returns></returns>
    Public Function getLastcharFromClname(ByVal mClname As String) As String
        Dim clsCust As New Customers.Customers.Customers
        Dim DtRow As New DataTable
        DtRow = df1.GetDataFromSql(clsCust.ServerDatabase, clsCust.TableName, "P_customers,Customers_key,CustCode,Custname", "", "custcode like '" & mClname & "%'", "", "custcode desc")
        If DtRow.Rows.Count = 0 Then
            Return mClname & "0001"
        Else
            Dim lcustc As String = DtRow(0).Item("custcode")
            Dim lcust1 As String = ""
            Dim mpl As Integer = 0
            Dim custleft As String = ""
            custleft = Strings.Left(lcustc, 4)
            lcust1 = Strings.Right(lcustc, 4)
            mpl = CInt(lcust1)
            mpl = mpl + 1
            Dim k As String = CStr(mpl)
            k = k.PadLeft(4, "0")
            Return custleft & k
        End If
    End Function
    Public Function GetProductsRateForDealers(ByVal LoginKey As Integer) As DataTable
        Dim ServiceCodesStr As String = "2401,2402,2412,2773,2774,2403,3019"
        Dim ServiceCodes As String() = ServiceCodesStr.Split(",")
        Dim dt As New DataTable
        dt = df1.AddColumnsInDataTable(dt, "ServiceCode,Rate", "System.Int32,System.String")
        For i = 0 To ServiceCodes.Length - 1
            Dim clsRateParameters As New Hashtable
            clsRateParameters = CreateRateParameters(ServiceCodes(i),, LoginKey,,,,,,,,,)
            Dim ratehashTable1 As New Hashtable
            ratehashTable1 = getRateforDealer(clsRateParameters)
            Dim rate As Double = GF1.GetValueFromHashTable(ratehashTable1, "rate")
            Dim ServiceRow As DataRow = dt.NewRow
            ServiceRow("ServiceCode") = ServiceCodes(i)
            ServiceRow("Rate") = rate
            dt.Rows.Add(ServiceRow)
        Next
        Return dt
    End Function
    Public Function ProcessRegistrationFilesWithoutPayment_new(ByVal customers As DataTable, ByVal sessionRow As DataRow, ByVal dealerrow As DataRow, ByVal DealerStaffRow As DataRow, ByVal newbutOld As String) As String
        Dim clsvm As New ValidateMachine.ValidateClass
        Dim EPLHashList As New List(Of Hashtable)

        Dim filename As String = df1.getDateTimeISTNow.ToString("yyyy_MM_dd_hh_mm_ss")

        ' Dim fpath As String = ""

        Dim linktype As String = sessionRow.Item("linktype")
        Dim dname As String = ""
        If linktype = "U" Then
            dname = df1.GetCellValue(DealerStaffRow, "name", "string")
        Else
            dname = df1.GetCellValue(dealerrow, "accname", "string")
        End If
        'fpath = "~/Dealer/" & dname.ToString.Trim & "/PendingReg/REG_" & filename & ".zip"
        ' fpath = GlobalControl.Variables.DataFolderServerPhysicalPath & "/Dealer/" & dname.ToString.Trim & "/PendingReg/REG_" & filename & ".zip"
        Dim xplpath As String = ""
        Dim mailtext As String = ""
        Dim mcustcode As String = ""
        Dim lpatc As String = ""
        Dim mailTextForView As String = ""
        For k = 0 To customers.Rows.Count - 1
            Dim regtype2 As String = df1.GetCellValue(customers.Rows(k), "regtype2")
            Dim billeduptodt As DateTime = libcustomerfeature.GetBilledUpToDate(customers.Rows(k).Item("p_customers"), regtype2)
            Dim billeduptodtStr As String = billeduptodt.ToString("yyyy-MM-dd")
            If customers.Rows(k).Item("isvalid") = "true" Then
                If customers.Rows(k).Item("chargingtrue") = "false" And customers.Rows(k).Item("paymentflag") = True Then Continue For
                Dim lPath As String = customers(k).Item("eplpathserver")
                lpatc = lPath
                Dim abc As New Hashtable
                abc = clsvm.GetHashTableFromClientReg(lPath)
                'If customers.Rows(k).Item("custverifyflag") = "true" Then
                '    If customers.Rows(k).Item("regtype") = "amc" Then
                '        customers.Rows(k).Item("ChangeAllowUptoStr") = df1.getDateTimeISTNow().AddDays(5).ToString("yyyy-MM-dd")

                '    End If
                'End If
                abc = GF1.AddItemToHashTable(abc, "l_date", customers.Rows(k).Item("ChangeAllowUptoStr"))
                abc = GF1.AddItemToHashTable(abc, "CustCode", customers.Rows(k).Item("ChangeCustCode"))

                If customers.Rows(k).Item("regtype2") = "main" Then
                    If Not CInt(customers.Rows(k).Item("changeNodes").ToString) = 0 Then

                        Dim knod As Integer = CInt(customers.Rows(k).Item("ChangeNodes").ToString) + 1
                        abc = GF1.AddItemToHashTable(abc, "nodes", knod)

                    End If
                End If
                If customers.Rows(k).Item("ChangeNodes") = 0 Then abc = GF1.AddItemToHashTable(abc, "lan", "0") Else abc = GF1.AddItemToHashTable(abc, "lan", "1")

                mcustcode = UCase(customers.Rows(k).Item("changecustcode"))

                Dim lpatc1 As String = Path.GetDirectoryName(lpatc) & "\"
                Dim xplfile As String = ""


                xplfile = clsvm.CreateRegFileFromHashTable_new(abc, lpatc1, mcustcode)
                abc = GF1.AddItemToHashTable(abc, "P_customers", customers.Rows(k).Item("p_customers"))
                EPLHashList.Add(abc)

                xplpath = xplpath & "," & xplfile
                Dim mregtype As String = ""
                If customers.Rows(k).Item("regtype") = "amc" Then mregtype = "old"

                mailtext = mailtext & "<br/>" & dealerrow.Item("P_acccode").ToString & "  &nbsp; &nbsp;&nbsp; &nbsp; " & customers.Rows(k).Item("CustCode") & "&nbsp; &nbsp; &nbsp;  " & customers.Rows(k).Item("custname") & " &nbsp; &nbsp;  " & customers.Rows(k).Item("texthometown") & "  &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;              " & customers.Rows(k).Item("ChangeAllowUptoStr") & "   &nbsp; &nbsp;   " & customers.Rows(k).Item("regtype") & " &nbsp; &nbsp;      " & customers.Rows(k).Item("regtype2") & " &nbsp; &nbsp;  " & customers.Rows(k).Item("changeNodes") & " &nbsp; &nbsp; " & Path.GetFileName(xplfile)
                mailTextForView += "," & dealerrow.Item("P_acccode").ToString & "~" & customers.Rows(k).Item("CustCode") & "~" & customers.Rows(k).Item("custname") & "~" & customers.Rows(k).Item("texthometown") & "~" & customers.Rows(k).Item("ChangeAllowUptoStr") & "~" & customers.Rows(k).Item("regtype") & "~" & customers.Rows(k).Item("ChangeNodes") & "~" & Path.GetFileName(xplfile)
            End If
        Next

        'UpdateAllowuptoSTRInCust(customers)
        'Dim regtranhash As Hashtable = AddRowInRegistrationsTran_new(customers, sessionRow, EPLHashList, newbutOld)
        'addrowIncustomerVerification(customers, Nothing, sessionRow)
        If xplpath = "" Then
            Return ""
            Exit Function
        End If
        If Not Directory.Exists(Path.GetDirectoryName(lpatc) & "\xplfiles") Then Directory.CreateDirectory(Path.GetDirectoryName(lpatc) & "\xplfiles")
        'If Not Directory.Exists(Path.GetDirectoryName(lpatc)) Then Directory.CreateDirectory(Path.GetDirectoryName(lpatc))

        If xplpath.First = "," Then xplpath = xplpath.Substring(1, xplpath.Length - 1)
        If mailtext.First = "," Then mailtext = mailtext.Substring(1, mailtext.Length - 1)
        If mailTextForView.First = "," Then mailTextForView = mailTextForView.Substring(1, mailTextForView.Length - 1)

        Dim abc1 As String = Path.GetDirectoryName(lpatc) & "\REG" & filename & ".zip"
        Dim xplpatch1() As String = Split(xplpath, ",")
        For op = 0 To xplpatch1.Count - 1
            FileIO.FileSystem.CopyFile(xplpatch1(op), Path.GetDirectoryName(lpatc) & "\xplfiles" & "\" & Path.GetFileName(xplpatch1(op)), True)
        Next

        ZipFile.CreateFromDirectory(Path.GetDirectoryName(lpatc) & "\xplfiles", Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")
        'Dim fileName_final As String = Path.GetDirectoryName(fpath) & "\REG_" & filename & ".zip"
        Dim zipfilename As String = Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip"
        Dim fileName_final As String = zipfilename.Replace("PendingReg", "xpl")

        FileIO.FileSystem.CopyFile(Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip", fileName_final, True)
        Dim gf2 As New GlobalFunction2.GlobalFunction2
        Dim emaillist As String = ""
        emaillist = dealerrow.Item("email").trim & ",hcgupta@saralerp.com,nehagupta@saralerp.com"
        If mailTextForView.First = "," Then mailTextForView = mailTextForView.Substring(1, mailTextForView.Length - 1)
        Dim msgcontent As String = mailtext & Chr(201) & "Registration Files " & filename
        emaillist = dealerrow.Item("email").trim & ",hcgupta@saralerp.com,nehagupta@saralerp.com"
        Dim cfc1 As New CommonFunctionsCloud.CommonFunctionsCloud
        Dim dtmsg As DataTable = cfc1.CreateMsgQueueDt("E", msgcontent, fileName_final, emaillist, "", "N")
        'Dim dtmsg As DataTable = cfc1.CreateMsgQueueDt("E", msgcontent, fileName_final, "93nishayadav@gmail.com", "", "N")
        cfc1.InsertIntoMsgQueue(dtmsg, sessionRow)
        'gf2.SendingEmail(emaillist, "Registration Files " & filename, mailtext, Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")   ' for server
        ' gf2.SendingEmail(emaillist, "Registration Files " & filename, mailtext, Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")   ' for server
        'HomeController.SendingEmail("93nishayadav@gmail.com", "Registration Files " & filename, mailtext, Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")
        Return fileName_final & "#" & "REG_" & filename & ".zip" & "#" & mailTextForView
    End Function
    Public Function ProcessRegistrationFilesWithPayment_new(ByVal customers As DataTable, ByVal sessionRow As DataRow, ByVal dealerrow As DataRow, ByVal DealerStaffRow As DataRow, ByVal newbutold As String)
        Dim EPLHashList As New List(Of Hashtable)
        Dim clsvm As New ValidateMachine.ValidateClass
        Dim filename As String = df1.getDateTimeISTNow.ToString("yyyy_MM_dd_hh_mm_ss")
        ' Dim fpath As String = ""
        Dim linktype As String = sessionRow.Item("linktype")
        Dim dname As String = ""
        If linktype = "U" Then
            dname = df1.GetCellValue(DealerStaffRow, "name", "string")
        Else
            dname = df1.GetCellValue(dealerrow, "accname", "string")
        End If

        'fpath = "~/Dealer/" & dname.ToString.Trim & "/PendingReg/REG_" & filename & ".zip"
        ' fpath = GlobalControl.Variables.DataFolderServerPhysicalPath & "/Dealer/" & dname.ToString.Trim & "/PendingReg/REG_" & filename & ".zip"
        '  fpath = "C:/inetpub/CRMData/Dealer/" & dname.ToString.Trim & "/PendingReg/REG_" & filename & ".zip"

        '    fpath = "~/Dealer/" & DealerStaffRow.Item("name").ToString.Trim & "/PendingReg/REG_" & filename & ".zip"

        Dim xplpath As String = ""
        Dim mailtext As String = ""
        Dim mcustcode As String = ""
        Dim lpatc As String = ""
        Dim mailTextForView As String = ""
        For k = 0 To customers.Rows.Count - 1
            If customers.Rows(k).Item("isvalid") = "true" Then
                If customers.Rows(k).Item("chargingtrue") = "false" And customers.Rows(k).Item("paymentflag") = True Then Continue For
                Dim lPath As String = customers(k).Item("eplpathserver")
                lpatc = lPath
                Dim regtype2 As String = df1.GetCellValue(customers.Rows(k), "regtype2")
                Dim abc As New Hashtable
                abc = clsvm.GetHashTableFromClientReg(lPath)
                abc = GF1.AddItemToHashTable(abc, "l_date", customers.Rows(k).Item("ChangeAllowUptoStr"))
                abc = GF1.AddItemToHashTable(abc, "CustCode", customers.Rows(k).Item("ChangeCustCode"))
                If customers.Rows(k).Item("regtype2") = "main" Then
                    If Not CInt(customers.Rows(k).Item("changeNodes").ToString) = 0 Then
                        Dim knod As Integer = CInt(customers.Rows(k).Item("ChangeNodes").ToString) + 1
                        abc = GF1.AddItemToHashTable(abc, "nodes", knod)
                    End If
                End If
                If customers.Rows(k).Item("ChangeNodes") = 0 Then abc = GF1.AddItemToHashTable(abc, "lan", "0") Else abc = GF1.AddItemToHashTable(abc, "lan", "1")
                mcustcode = UCase(customers.Rows(k).Item("changecustcode"))
                Dim lpatc1 As String = Path.GetDirectoryName(lpatc) & "\"
                Dim xplfile As String = ""
                xplfile = clsvm.CreateRegFileFromHashTable_new(abc, lpatc1, mcustcode)
                abc = GF1.AddItemToHashTable(abc, "p_customers", customers.Rows(k).Item("p_customers"))
                EPLHashList.Add(abc)
                xplpath = xplpath & "," & xplfile
                Dim mregtype As String = ""
                If customers.Rows(k).Item("regtype") = "amc" Then mregtype = "old"
                mailtext = mailtext & "<br/>" & dealerrow.Item("p_acccode").ToString & "  &nbsp; &nbsp;&nbsp; &nbsp; " & customers.Rows(k).Item("CustCode") & "&nbsp; &nbsp; &nbsp;  " & customers.Rows(k).Item("custname") & " &nbsp; &nbsp;  " & customers.Rows(k).Item("texthometown") & "  &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp;              " & customers.Rows(k).Item("ChangeAllowUptoStr") & "   &nbsp; &nbsp;   " & customers.Rows(k).Item("regtype") & " &nbsp; &nbsp;      " & customers.Rows(k).Item("regtype2") & " &nbsp; &nbsp;  " & customers.Rows(k).Item("changeNodes") & " &nbsp; &nbsp; " & Path.GetFileName(xplfile)
                mailTextForView += "," & dealerrow.Item("p_acccode").ToString & "~" & customers.Rows(k).Item("CustCode") & "~" & customers.Rows(k).Item("custname") & "~" & customers.Rows(k).Item("texthometown") & "~" & customers.Rows(k).Item("ChangeAllowUptoStr") & "~" & customers.Rows(k).Item("regtype") & "~" & customers.Rows(k).Item("ChangeNodes") & "~" & Path.GetFileName(xplfile)
            End If
        Next
        '  UpdateAllowuptoSTRInCust(customers)
        ' Dim regtranHAsh As Hashtable = AddRowInRegistrationsTran_new(customers, sessionRow, EPLHashList, newbutold)
        If xplpath = "" Then
            Return "" ' RedirectToAction("EditCustomerReg", New With {.eplUpload = True})
            Exit Function
        End If
        If Not Directory.Exists(Path.GetDirectoryName(lpatc) & "\xplfiles") Then Directory.CreateDirectory(Path.GetDirectoryName(lpatc) & "\xplfiles")
        If xplpath.First = "," Then xplpath = xplpath.Substring(1, xplpath.Length - 1)
        If mailtext.First = "," Then mailtext = mailtext.Substring(1, mailtext.Length - 1)
        If mailTextForView.First = "," Then mailTextForView = mailTextForView.Substring(1, mailTextForView.Length - 1)

        Dim abc1 As String = Path.GetDirectoryName(lpatc) & "\REG" & filename & ".zip"
        Dim xplpatch1() As String = Split(xplpath, ",")
        For op = 0 To xplpatch1.Count - 1
            FileIO.FileSystem.CopyFile(xplpatch1(op), Path.GetDirectoryName(lpatc) & "\xplfiles" & "\" & Path.GetFileName(xplpatch1(op)), True)
        Next

        ZipFile.CreateFromDirectory(Path.GetDirectoryName(lpatc) & "\xplfiles", Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")
        'Dim fileName_final As String = Path.GetDirectoryName(fpath) & "\REG_" & filename & ".zip"
        Dim zipfilename As String = Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip"
        Dim fileName_final As String = zipfilename.Replace("PendingReg", "xpl")

        FileIO.FileSystem.CopyFile(Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip", fileName_final, True)

        Dim gf2 As New GlobalFunction2.GlobalFunction2
        Dim emaillist As String = ""
        Dim cfc1 As New CommonFunctionsCloud.CommonFunctionsCloud
        Dim msgcontent As String = mailtext & Chr(201) & "Registration Files " & filename
        emaillist = dealerrow.Item("email").trim & ",hcgupta@saralerp.com,nehagupta@saralerp.com"
        ' Dim dtmsg As DataTable = cfc1.CreateMsgQueueDt("E", msgcontent, fileName_final, emaillist, "", "N")
        'cfc1.InsertIntoMsgQueue(dtmsg, sessionRow)
        'gf2.SendingEmail(emaillist, "Registration Files " & filename, mailtext, Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")   ' for server
        '  gf2.SendingEmail("nehagupta@saralerp.com", "Registration Files " & filename, mailtext, Path.GetDirectoryName(lpatc) & "\REG_" & filename & ".zip")   ' for server

        Return fileName_final & "#" & "REG_" & filename & ".zip" & "#" & mailTextForView
    End Function
#End Region
#Region "Payment Processor"
    ''' <summary>
    ''' Used for Processing and updating order after payment is done from registration create functionality and is successfull
    ''' </summary>
    ''' <param name="lorderHeader"></param>
    ''' <param name="totalamt">Total amount paid by payment gateway</param>
    ''' <param name="TotalAmtLessPaymentGateway">Total amount as sum total of chargingheaders</param>
    ''' <param name="amtwithoutPGCharges">Total available amount offset by dealer/emp ledger</param>
    ''' <param name="calledFrom"></param>
    ''' <param name="sessionrow"></param>
    ''' <param name="p_acccode"></param>
    ''' <returns></returns>
    Public Function ProcessOrderPostpaymentSuccessFull(ByVal lorderHeader As Integer, ByVal TotalAmt As Decimal, ByVal TotalAmtLessPaymentGateway As Decimal, ByVal amtwithoutPGCharges As Decimal, ByVal calledFrom As String, ByVal sessionrow As DataRow, ByVal p_acccode As Integer, ByVal DealerRow As DataRow) As DataTable
        Dim p_payment As Integer = -1
        Dim clsPayMent As New Payment.Payment.Payment
        Dim dtpayment As New DataTable
        Dim clsorderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim dtorderHeader As New DataTable
        dtorderHeader = df1.GetDataFromSql(clsorderHeader.ServerDatabase, clsorderHeader.TableName, "*", "", "rowstatus=0 and p_orderheader =" & lorderHeader, "", "")
        Dim p_customer As Integer = -1
        If dtorderHeader.Rows.Count < 0 Then
            Return dtpayment
            Exit Function
        Else
            Try
                p_customer = CInt(df1.GetCellValue(dtorderHeader.Rows(0), "p_customers"))
            Catch
                p_customer = -1
            End Try

        End If


        If TotalAmt > 0 Then
            clsPayMent.CurrRow.Item("paymentmode") = 1
            clsPayMent.CurrRow.Item("PaymentDate") = df1.getDateTimeISTNow
            clsPayMent.CurrRow.Item("amount") = TotalAmt
            clsPayMent.CurrRow.Item("VerifyCode") = "V"
            clsPayMent.CurrRow.Item("VerifyEmp") = -1
            clsPayMent.CurrRow.Item("VerifyDate") = df1.getDateTimeISTNow
            clsPayMent.CurrRow.Item("P_acccode") = p_acccode
            clsPayMent.CurrRow.Item("VerifyDate") = df1.getDateTimeISTNow
            clsPayMent.CurrRow.Item("benaccount") = 2751
            clsPayMent.CurrRow.Item("handleby") = "D"
            clsPayMent.CurrRow.Item("status") = "U"
            clsPayMent.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
            Dim customertrack As String = ""
            customertrack = df1.GetCellValue(DealerRow, "customerwisepaymenttrack")
            If customertrack = "Y" Then
                clsPayMent.CurrRow.Item("custcode") = getcustcodeFromP_customers(p_customer)
            End If
            Dim aClsObject() As Object = {clsPayMent}
            Dim mserverdb As String = df1.GetServerMDFForTransanction(aClsObject)
            Dim mytrans As SqlTransaction = df1.BeginTransaction(mserverdb)
            Dim aLastKeysValues As New Hashtable
            aClsObject = df1.SetKeyValueIfNewInsert(mytrans, aClsObject)
            Dim HashPublicValues As New Hashtable
            Dim sqlexec As Boolean = df1.CheckTableClassUpdations(aClsObject)
            aClsObject = df1.LastKeysPlus(mytrans, aClsObject, aLastKeysValues)
            aClsObject = df1.SetFinalFieldsValues(aClsObject, HashPublicValues)
            Dim aam As Integer = 0
            Try
                If sqlexec = True Then
                    aam = df1.InsertUpdateDeleteSqlTables(mytrans, aClsObject, aam)
                    Dim paymentHash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValues, "payment")
                    p_payment = GF1.GetValueFromHashTable(paymentHash, "p_payment")
                    mytrans.Commit()
                End If
            Catch ex As Exception
                mytrans.Rollback()
            End Try
            mytrans.Dispose()
        End If

        clsorderHeader.PrevRow = df1.UpdateDataRows(clsorderHeader.PrevRow, dtorderHeader.Rows(0))
        ' clsorderHeader.CurrRow.Item("P_payment") = p_payment
        clsorderHeader.CurrRow.Item("paymentflag") = "P"
        Dim aclsobj() As Object = {clsorderHeader}

        cfc.SaveIntodb(aclsobj)
        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtchargingHeader As New DataTable
        dtchargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus=0 and orderheader =" & lorderHeader, "", "")
        For k = 0 To dtchargingHeader.Rows.Count - 1
            Dim clschargingHeader1 As New ChargingHeader.ChargingHeader.ChargingHeader
            clschargingHeader1.PrevRow = df1.UpdateDataRows(clschargingHeader1.PrevRow, dtchargingHeader.Rows(k))
            clschargingHeader1.CurrRow.Item("P_payment") = p_payment
            clschargingHeader1.CurrRow.Item("paymentflag") = "P"
            Dim aClsObject1() As Object = {clschargingHeader1}
            cfc.SaveIntodb(aClsObject1)
        Next
        dtpayment = df1.GetDataFromSql(clsPayMent.ServerDatabase, clsPayMent.TableName, "", "", "rowstatus = 0 and p_payment = " & p_payment, "", "")
        Return dtpayment
    End Function



    ''' <summary>
    ''' Function to send dt for verify payment in CRM controller
    ''' </summary>
    ''' <param name="start">integer indicating from which rows are to be brought</param>
    ''' <param name="psize">No. of rows to be brought</param>
    ''' <param name="lcondition"> condition</param>
    ''' <param name="DtInfoTable">dt containing infotable</param>
    ''' <returns></returns>
    Public Function VerifyPaymentDataGrid(start As Integer, psize As Integer, lcondition As String, DtInfoTable As DataTable) As DataTable
        Dim dt As New DataTable
        Dim clsdealers As New Dealers.Dealers.Dealers
        dt = df1.GetDataFromSqlFixedRows(clsdealers.ServerDatabase, "Payment", "*", "", lcondition, "", "Payment_Key desc", start, psize, -1)
        dt = df1.AddColumnsInDataTable(dt, "TextAmount,TextVerifyCode,PaymentDate1,TextP_acccode,TextStatus,AvailableAmount,TextCommissionTo,Textdiscount", "System.String,System.String,System.String,System.String,System.String,System.String,System.String,System.String")

        dt = df1.AddingNameForCodesPrimamryCols(dt, "PaymentMode,BenAccount", "TextPaymentMode,TextBenAccount", DtInfoTable, "NameOfInfo")

        For i = 0 To dt.Rows.Count - 1

            If IsDBNull(dt.Rows(i).Item("discount")) = False Then
                dt.Rows(i).Item("Textdiscount") = Math.Round(dt.Rows(i).Item("discount"), 2)
            Else
                dt.Rows(i).Item("Textdiscount") = "N/A"
            End If



            If IsDBNull(dt.Rows(i).Item("Amount")) = False Then
                dt.Rows(i).Item("TextAmount") = Math.Round(dt.Rows(i).Item("Amount"), 2).ToString
            End If
            If IsDBNull(dt.Rows(i).Item("VerifyCode")) = False Then
                If dt.Rows(i).Item("VerifyCode") = "R" Then
                    dt.Rows(i).Item("TextVerifyCode") = "Rejected"
                ElseIf dt.Rows(i).Item("VerifyCode") = "P" Then
                    dt.Rows(i).Item("TextVerifyCode") = "Pending"
                ElseIf dt.Rows(i).Item("VerifyCode") = "V" Then
                    dt.Rows(i).Item("TextVerifyCode") = "Verified"
                End If
            End If
            If IsDBNull(dt.Rows(i).Item("Status")) = False Then
                If dt.Rows(i).Item("Status") = "F" Then
                    dt.Rows(i).Item("TextStatus") = "Fully Adjusted"
                ElseIf dt.Rows(i).Item("Status") = "P" Then
                    dt.Rows(i).Item("TextStatus") = "Partially Adjusted"
                ElseIf dt.Rows(i).Item("Status") = "U" Then
                    dt.Rows(i).Item("TextStatus") = "Unadjusted"
                End If
            End If
            Dim paymentDate As Date = df1.GetCellValue(dt.Rows(i), "PaymentDate")
            dt.Rows(i).Item("PaymentDate1") = paymentDate.ToString("dd-MM-yyyy")
            If IsDBNull(dt.Rows(i).Item("P_acccode")) = False Then
                Dim Query As String = String.Format("Select AccName from AccMaster where P_acccode=" & dt.Rows(i).Item("P_acccode"))
                Dim AccMasterDt As DataTable = df1.SqlExecuteDataTable(clsdealers.ServerDatabase, Query)
                If AccMasterDt.Rows.Count > 0 Then
                    dt.Rows(i).Item("TextP_acccode") = AccMasterDt.Rows(0).Item("AccName")
                Else
                    dt.Rows(i).Item("TextP_acccode") = "N/A"
                End If
            Else
                dt.Rows(i).Item("TextP_acccode") = "N/A"
            End If
            If IsDBNull(dt.Rows(i).Item("CommissionTo")) = False Then

                Dim Query1 As String = String.Format("Select AccName from AccMaster where P_acccode=" & dt.Rows(i).Item("CommissionTo"))
                Dim AccNamedt As DataTable = df1.SqlExecuteDataTable(clsdealers.ServerDatabase, Query1)
                If AccNamedt.Rows.Count > 0 Then
                    dt.Rows(i).Item("TextCommissionTo") = AccNamedt.Rows(0).Item("AccName")
                Else
                    dt.Rows(i).Item("TextCommissionTo") = "N/A"
                End If
            Else
                dt.Rows(i).Item("TextCommissionTo") = "N/A"
            End If
            Dim amount As Decimal = df1.GetCellValue(dt.Rows(i), "Amount")

            Dim discount As Decimal = df1.GetCellValue(dt.Rows(i), "discount")

            Dim TempAvailableAmoount As Decimal = GetBalanceFromPaymentVoucher(dt.Rows(i).Item("P_Payment"))
            Dim AvailableAmoount As Decimal = Math.Round(((discount + amount) - TempAvailableAmoount), 2)
            dt.Rows(i).Item("AvailableAmount") = AvailableAmoount.ToString
        Next
        Return dt
    End Function




    Public Sub RollbackOrderPostpaymentFailure(ByVal lorderHeader As Integer)
        Dim clsorderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim dtorderHeader As New DataTable
        dtorderHeader = df1.GetDataFromSql(clsorderHeader.ServerDatabase, clsorderHeader.TableName, "*", "", "rowstatus=0 and p_orderheader =" & lorderHeader, "", "")
        Dim p_customer As Integer = -1
        If dtorderHeader.Rows.Count < 0 Then
            Exit Sub
        End If
        clsorderHeader.PrevRow = df1.UpdateDataRows(clsorderHeader.PrevRow, dtorderHeader.Rows(0))
        ' clsorderHeader.CurrRow.Item("P_payment") = p_payment
        clsorderHeader.CurrRow.Item("paymentflag") = "U"
        Dim aclsobj() As Object = {clsorderHeader}
        cfc.SaveIntodb(aclsobj)
        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtchargingHeader As New DataTable
        dtchargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus=0 and orderheader =" & lorderHeader, "", "")
        For k = 0 To dtchargingHeader.Rows.Count - 1
            Dim clschargingHeader1 As New ChargingHeader.ChargingHeader.ChargingHeader
            clschargingHeader1.PrevRow = df1.UpdateDataRows(clschargingHeader1.PrevRow, dtchargingHeader.Rows(k))
            '   clschargingHeader1.CurrRow.Item("P_payment") = p_payment
            clschargingHeader1.CurrRow.Item("paymentflag") = "U"
            Dim aClsObject1() As Object = {clschargingHeader1}
            cfc.SaveIntodb(aClsObject1)
        Next
        'Deleting billpayflag mapping
        For k = 0 To dtchargingHeader.Rows.Count - 1
            Dim headerNo As Integer = df1.GetCellValue(dtchargingHeader.Rows(k), "headerno")
            Dim kSql As String = "delete from billpayflag where headerno = " & headerNo
            df1.SqlExecuteNonQuery(clschargingHeader.ServerDatabase, kSql)
        Next
        '   dtpayment = df1.GetDataFromSql(clsPayMent.ServerDatabase, clsPayMent.TableName, "", "", "rowstatus = 0 and p_payment = " & p_payment, "", "")
    End Sub
    Public Sub UpdateFinalColsInCustomersForCaseNew(ByVal customersDt As DataTable, ByVal OrderDt As DataTable, Optional ByVal orderheader As Integer = -1)
        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtchargingHeader As New DataTable
        Dim clscustomers As New Customers.Customers.Customers
        If orderheader = -1 Then
            For k = 0 To customersDt.Rows.Count - 1
                Dim custorder() As DataRow = OrderDt.Select("P_customers =" & df1.GetCellValue(customersDt.Rows(k), "p_customers") & "and servicecode = 2401")
                If custorder.Count > 0 Then
                    Dim mdateStr As String = ""
                    Dim mdt As DateTime = libcustomerfeature.GetBilledUpToDate(df1.GetCellValue(customersDt.Rows(k), "p_customers"), "main") 'df1.GetCellValue(custorder(0), "chargingdate")
                    Dim dd As String = mdt.Day.ToString.PadLeft(2, "0")
                    Dim mm As String = mdt.Month.ToString.PadLeft(2, "0")
                    mdateStr = "M" & dd & mm
                    Dim strsql As String = "update customers set billedupto = '" & mdt.ToString("yyyy-MM-dd H:mm:ss") & "' , amcmonth = '" & mdateStr & "'" & " , customerstatus = 'A' , activeF = 'Y' where p_customers = " & df1.GetCellValue(customersDt.Rows(k), "p_customers") & " and rowstatus = 0"
                    Dim l As Integer = df1.SqlExecuteNonQuery(clscustomers.ServerDatabase, strsql)
                End If
            Next
        Else
            'dtchargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus=0 and orderheader =" & orderheader, "", "")
            'Dim chrgingItem As New ChargingItems.ChargingItems.ChargingItems
            'For y = 0 To dtchargingHeader.Rows.Count - 1
            '    Dim dtchrgingItem As DataTable = df1.GetDataFromSql(chrgingItem.ServerDatabase, chrgingItem.TableName, "*", "", "headerno =" & dtchargingHeader.Rows(y).Item("headerno") & " and rowstatus = 0", "", "")
            '    Dim custorder() As DataRow = dtchrgingItem.Select("P_customers =" & df1.GetCellValue(dtchargingHeader.Rows(y), "p_customers") & "and servicecode = 2401")
            '    If custorder.Count > 0 Then
            '        Dim mdateStr As String = ""
            '        Dim mdt As DateTime = df1.GetCellValue(custorder(0), "chargingto")
            '        Dim dd As String = mdt.Day.ToString.PadLeft(2, "0")
            '        Dim mm As String = mdt.Month.ToString.PadLeft(2, "0")
            '        mdateStr = "M" & dd & mm
            '        Dim strsql As String = "update customers set billedupto = '" & df1.GetCellValue(custorder(0), "chargingto") & "' , amcmonth = '" & mdateStr & "'" & " , customerstatus = 'A' , activeF = 'Y' where p_customers = " & df1.GetCellValue(dtchargingHeader.Rows(y), "p_customers") & " and rowstatus = 0"
            '        Dim l As Integer = df1.SqlExecuteNonQuery(clscustomers.ServerDatabase, strsql)

            '    End If
            'Next
        End If
    End Sub
    Public Sub removeRowfromRegBlock(ByVal orderheader As Integer)
        Dim clschargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim dtchargingHeader As New DataTable
        dtchargingHeader = df1.GetDataFromSql(clschargingHeader.ServerDatabase, clschargingHeader.TableName, "*", "", "rowstatus=0 and orderheader =" & orderheader, "", "")
        For k = 0 To dtchargingHeader.Rows.Count - 1
            Dim headerno As Integer = df1.GetCellValue(dtchargingHeader.Rows(k), "headerno")
            Dim p_customers As Integer = df1.GetCellValue(dtchargingHeader.Rows(k), "p_customers")
            Dim sqlstr1 As String = "select * from chargingitems where  rowstatus = 0 and headerno = " & headerno
            Dim dt1 As DataTable = df1.SqlExecuteDataTable(clschargingHeader.ServerDatabase, sqlstr1)
            For o = 0 To dt1.Rows.Count - 1
                Dim servcCode As Integer = df1.GetCellValue(dt1.Rows(o), "servicecode")
                If servcCode = 2412 Or servcCode = 3019 Or servcCode = 2401 Then
                    Dim sqlstr As String = "delete from regblock where p_customers = " & p_customers
                    Dim clsregblock As New Regblock.Regblock.Regblock
                    Dim l As Integer = df1.SqlExecuteNonQuery(clsregblock.ServerDatabase, sqlstr)
                    Exit For
                End If
            Next
        Next
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dtunpaidorderList"></param>
    ''' <param name="sessionrow"></param>
    ''' <param name="totalamt">Total amount paid by payment gateway</param>
    ''' <param name="TotalAmtLessPaymentGateway">Total amount as sum total of chargingheaders</param>
    ''' <param name="amtwithoutPGCharges">Total amount offset by dealer/emp ledger</param>
    ''' <returns></returns>
    Public Function ProcessOrderFromUnpaidGrid(ByVal dtunpaidorderList As DataTable, ByVal sessionrow As DataRow, ByVal totalamt As Decimal, ByVal TotalAmtLessPaymentGateway As Decimal, ByVal amtwithoutPGCharges As Decimal) As Integer
        Dim unpaidList As New List(Of String)
        Dim P_customerStr As String = ""
        For k = 0 To dtunpaidorderList.Rows.Count - 1
            Dim p_orderheader As Integer = df1.GetCellValue(dtunpaidorderList.Rows(k), "P_orderheader")
            P_customerStr = P_customerStr & "," & df1.GetCellValue(dtunpaidorderList(k), "p_customers")
            If GF1.FindIndexListOfString(unpaidList, CStr(p_orderheader)) < 0 Then
                unpaidList.Add(CStr(p_orderheader))
            End If
        Next
        If P_customerStr.First = "," Then P_customerStr = P_customerStr.Substring(1)
        Dim clsOrderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim currdate As New DateTime
        currdate = df1.getDateTimeISTNow
        Dim logintype As String = sessionrow("linktype")
        Dim paymentby As String = logintype
        Dim logincode As Integer = sessionrow("linkcode")
        Dim lorderheader As Integer = -1
        Dim P_acccode As Integer = getAccCodefromLogincodetype(logintype, logincode)
        Dim p_customers As Integer = -1
        Dim webSessions_key As Integer = sessionrow("webSessions_key")
        clsOrderHeader.CurrRow.Item("p_customers") = P_customerStr
        clsOrderHeader.CurrRow.Item("orderdate") = df1.getDateTimeISTNow
        clsOrderHeader.CurrRow.Item("logintype") = logintype
        clsOrderHeader.CurrRow.Item("LoginCode") = logincode
        clsOrderHeader.CurrRow.Item("P_acccode") = P_acccode
        clsOrderHeader.CurrRow.Item("OrderSeries") = logintype & logincode
        clsOrderHeader.CurrRow.Item("paymentflag") = "U"
        clsOrderHeader.CurrRow.Item("totalamount") = totalamt + amtwithoutPGCharges  'IIf(OrderFinal.Rows.Count <= 0, 0, OrderFinal.Rows(0).Item("grandtotal"))
        clsOrderHeader.CurrRow.Item("calledFrom") = logintype & paymentby
        clsOrderHeader.CurrRow.Item("mtimestamp") = currdate
        clsOrderHeader.CurrRow.Item("websessions_key") = webSessions_key
        Dim aClsObject() As Object = {clsOrderHeader}
        Dim Success As Boolean = True
        Dim mserverdb As String = df1.GetServerMDFForTransanction(aClsObject)
        Dim mytrans As SqlTransaction = df1.BeginTransaction(mserverdb)
        Dim aLastKeysValues As New Hashtable
        aClsObject = df1.SetKeyValueIfNewInsert(mytrans, aClsObject)
        Dim HashPublicValues As New Hashtable
        Dim sqlexec As Boolean = df1.CheckTableClassUpdations(aClsObject)
        aClsObject = df1.LastKeysPlus(mytrans, aClsObject, aLastKeysValues)
        aClsObject = df1.SetFinalFieldsValues(aClsObject, HashPublicValues)
        Dim aam As Integer = -1
        Try
            If sqlexec = True Then
                aam = df1.InsertUpdateDeleteSqlTables(mytrans, aClsObject, aam)
                Dim OrderHeaderhash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValues, "OrderHeader")
                lorderheader = GF1.GetValueFromHashTable(OrderHeaderhash, "p_orderheader")
                mytrans.Commit()
                Success = True
            End If
        Catch ex As Exception
            mytrans.Rollback()
            Success = False
        End Try
        mytrans.Dispose()
        Return lorderheader
    End Function
    ''' <summary> 
    ''' </summary>
    ''' <param name="lorderheader"></param>
    ''' <param name="sessionRow"></param>
    ''' <param name="dtunpaidorderList"></param>
    ''' <param name="totalamt">Total amount paid by payment gateway</param>
    ''' <param name="TotalAmtLessPaymentGateway">Total amount as sum total of chargingheaders</param>
    ''' <param name="amtwithoutPGCharges">Total available amount offset by dealer/emp ledger</param>
    ''' <returns></returns>
    Public Function ProcessOrderPostPaymentUnpaidGrid(ByVal lorderheader As Integer, ByVal sessionRow As DataRow, ByVal dtunpaidorderList As DataTable, ByVal totalamt As Decimal, ByVal TotalAmtLessPaymentGateway As Decimal, ByVal amtwithoutPGCharges As Decimal) As DataTable
        Dim p_payment As Integer = -1
        Dim clsPayMent As New Payment.Payment.Payment
        Dim dtPayMent As New DataTable

        Dim currdate As New DateTime
        currdate = df1.getDateTimeISTNow
        Dim logintype As String = sessionRow("linktype")
        Dim paymentby As String = logintype
        Dim logincode As Integer = sessionRow("linkcode")
        Dim P_acccode As Integer = getAccCodefromLogincodetype(logintype, logincode)
        '   Dim p_acccode As Integer =
        If totalamt > 0 Then
            clsPayMent.CurrRow.Item("paymentmode") = 1
            clsPayMent.CurrRow.Item("PaymentDate") = df1.getDateTimeISTNow
            clsPayMent.CurrRow.Item("amount") = totalamt + amtwithoutPGCharges
            clsPayMent.CurrRow.Item("VerifyCode") = "V"
            clsPayMent.CurrRow.Item("VerifyEmp") = -1
            clsPayMent.CurrRow.Item("VerifyDate") = currdate 'df1.getDateTimeISTNow
            clsPayMent.CurrRow.Item("P_acccode") = P_acccode
            clsPayMent.CurrRow.Item("VerifyDate") = currdate 'df1.getDateTimeISTNow
            clsPayMent.CurrRow.Item("handleby") = "D"
            clsPayMent.CurrRow.Item("benaccount") = 2751
            clsPayMent.CurrRow.Item("status") = "U"
            clsPayMent.CurrRow.Item("websessions_key") = sessionRow("websessions_key")
            Dim aClsObject() As Object = {clsPayMent}
            Dim mserverdb As String = df1.GetServerMDFForTransanction(aClsObject)
            Dim mytrans As SqlTransaction = df1.BeginTransaction(mserverdb)
            Dim aLastKeysValues As New Hashtable
            aClsObject = df1.SetKeyValueIfNewInsert(mytrans, aClsObject)
            Dim HashPublicValues As New Hashtable
            Dim sqlexec As Boolean = df1.CheckTableClassUpdations(aClsObject)
            aClsObject = df1.LastKeysPlus(mytrans, aClsObject, aLastKeysValues)
            aClsObject = df1.SetFinalFieldsValues(aClsObject, HashPublicValues)
            Try
                If sqlexec = True Then
                    Dim aam As Boolean = df1.InsertUpdateDeleteSqlTables(mytrans, aClsObject)
                    Dim paymentHash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValues, "payment")
                    p_payment = GF1.GetValueFromHashTable(paymentHash, "p_payment")
                    mytrans.Commit()
                End If
            Catch ex As Exception
                mytrans.Rollback()
            End Try
            mytrans.Dispose()
        End If
        Dim clsorderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim dtorderHeader As New DataTable
        dtorderHeader = df1.GetDataFromSql(clsorderHeader.ServerDatabase, clsorderHeader.TableName, "*", "", "rowstatus=0 and p_orderheader =" & lorderheader, "", "")
        If dtorderHeader.Rows.Count < 0 Then
            Return dtPayMent
            Exit Function
        End If
        clsorderHeader.PrevRow = df1.UpdateDataRows(clsorderHeader.PrevRow, dtorderHeader.Rows(0))
        clsorderHeader.CurrRow.Item("P_payment") = p_payment
        clsorderHeader.CurrRow.Item("paymentflag") = "P"
        Dim aclsobj() As Object = {clsorderHeader}
        cfc.SaveIntodb(aclsobj)
        For j = 0 To dtunpaidorderList.Rows.Count - 1
            If dtunpaidorderList.Rows(j).Item("selectflag") = "false" Then Continue For
            If totalamt = 0 Then
                Dim clschargingheader As New ChargingHeader.ChargingHeader.ChargingHeader
                clschargingheader.PrevRow = df1.UpdateDataRows(clschargingheader.PrevRow, dtunpaidorderList.Rows(j))
                clschargingheader.CurrRow.Item("orderheader") = lorderheader
                clschargingheader.CurrRow.Item("logintype") = logintype
                clschargingheader.CurrRow.Item("logincode") = logincode
                clschargingheader.CurrRow.Item("p_acccode") = P_acccode
                clschargingheader.CurrRow.Item("mtimestamp") = currdate
                clschargingheader.CurrRow.Item("p_payment") = p_payment
                clschargingheader.CurrRow.Item("paymentflag") = "P"
                Dim clsob1() As Object = {clschargingheader}
                cfc.SaveIntodb(clsob1)
            Else
                Dim clschargingheader As New ChargingHeader.ChargingHeader.ChargingHeader
                clschargingheader.PrevRow = df1.UpdateDataRows(clschargingheader.PrevRow, dtunpaidorderList.Rows(j))
                Dim grandtotalprev As Decimal = df1.GetCellValue(clschargingheader.PrevRow, "grandtotal")
                Dim proptotal As Decimal = grandtotalprev / TotalAmtLessPaymentGateway
                clschargingheader.CurrRow.Item("orderheader") = lorderheader
                clschargingheader.CurrRow.Item("logintype") = logintype
                clschargingheader.CurrRow.Item("logincode") = logincode
                clschargingheader.CurrRow.Item("p_acccode") = P_acccode
                clschargingheader.CurrRow.Item("mtimestamp") = currdate
                clschargingheader.CurrRow.Item("grandtotal") = grandtotalprev + proptotal * (totalamt + amtwithoutPGCharges - TotalAmtLessPaymentGateway)
                Dim clsob1() As Object = {clschargingheader}
                cfc.SaveIntodb(clsob1)
                ' Dim clschargingItem1 As New ChargingItems.ChargingItems.ChargingItems
                Dim clschargingItem2 As New ChargingItems.ChargingItems.ChargingItems
                clschargingItem2.CurrRow.Item("headerno") = dtunpaidorderList.Rows(j).Item("headerno")
                '  am.Item("ItemSno") = dtchrging.Rows.Count + 1
                clschargingItem2.CurrRow.Item("orderheader") = lorderheader
                clschargingItem2.CurrRow.Item("chargingdate") = currdate
                clschargingItem2.CurrRow.Item("logintype") = logintype
                clschargingItem2.CurrRow.Item("logincode") = logincode
                clschargingItem2.CurrRow.Item("p_acccode") = P_acccode
                clschargingItem2.CurrRow.Item("p_customers") = dtunpaidorderList.Rows(j).Item("p_customers")
                clschargingItem2.CurrRow.Item("servicecode") = 3024
                clschargingItem2.CurrRow.Item("baseamount") = proptotal * (totalamt + amtwithoutPGCharges - TotalAmtLessPaymentGateway)
                clschargingItem2.CurrRow.Item("taxableamount") = proptotal * (totalamt + amtwithoutPGCharges - TotalAmtLessPaymentGateway)
                clschargingItem2.CurrRow.Item("mtimestamp") = currdate
                clschargingItem2.CurrRow.Item("WebSessions_Key") = sessionRow("websessions_key")
                clschargingItem2.CurrRow.Item("chargingitems_key") = -1
                clschargingItem2.CurrRow.Item("P_payment") = p_payment
                'dt1.Rows.Add(am)
                'clschargingItem1.CurrDt = dt1
                Dim clsob2() As Object = {clschargingItem2} ' New ChargingItems.ChargingItems.ChargingItems
                cfc.SaveIntodb(clsob2)
            End If
        Next
        If Not p_payment = -1 Then dtPayMent = df1.GetDataFromSql(clsPayMent.ServerDatabase, clsPayMent.TableName, "", "", "rowstatus = 0 and p_payment = " & p_payment, "", "")
        Return dtPayMent
    End Function
    Public Function cancelorderheader(ByVal dtunpaidorderList As DataTable, ByVal sessionrow As DataRow) As Integer
        Dim listoforderheader As New List(Of String)
        For k = 0 To dtunpaidorderList.Rows.Count - 1
            If dtunpaidorderList.Rows(k).Item("selectflag") = "false" Then Continue For
            Dim p_orderheader As Integer = df1.GetCellValue(dtunpaidorderList.Rows(k), "P_orderheader")

            If GF1.FindIndexListOfString(listoforderheader, CStr(p_orderheader)) < 0 Then
                listoforderheader.Add(CStr(p_orderheader))
            End If
        Next
        For l = 0 To listoforderheader.Count - 1
            Dim clsorderHeader As New OrderHeader.OrderHeader.OrderHeader
            Dim dt As DataTable = df1.GetDataFromSql(clsorderHeader.ServerDatabase, clsorderHeader.TableName, "*", "", "rowstatus = 0 and p_orderheader = " & listoforderheader(l), "", "")
            If dt.Rows.Count > 0 Then
                clsorderHeader.PrevRow = df1.UpdateDataRows(clsorderHeader.PrevRow, dt.Rows(0))
                clsorderHeader.CurrRow.Item("paymentflag") = "Q"
                clsorderHeader.CurrRow.Item("websessions_key") = sessionrow("websessions_key")
                Dim clsobj() As Object = {clsorderHeader}
                cfc.SaveIntodb(clsobj)
            End If
        Next
    End Function
    Public Function JoinChargingHeaderOrderHeaderForUnpaidOrders(ByVal orderheaderdt As DataTable, ByVal chargingheaderdt As DataTable)
        Dim FinalDt As New DataTable
        Dim ListOfOrderHeader As New List(Of String)
        FinalDt = chargingheaderdt.Clone
        FinalDt = df1.AddColumnsInDataTable(FinalDt, "rowtype,OrderHeader_key, P_OrderHeader,OrderDate,P_acccode,LoginCode,LoginType,Totalamount,FrmtOrderDate,WT", "System.String,System.int32,System.int32,system.datetime,System.int32,System.int32,System.String,system.decimal,System.String,System.String", "", "")
        For i = 0 To chargingheaderdt.Rows.Count - 1
            Dim orderheader As String = chargingheaderdt.Rows(i).Item("orderheader").ToString
            If GF1.FindIndexListOfString(ListOfOrderHeader, orderheader) < 0 Then
                ListOfOrderHeader.Add(orderheader)
            End If
        Next
        For i = 0 To ListOfOrderHeader.Count - 1
            Dim Rowchargingheader As DataRow() = chargingheaderdt.Select("orderheader = '" & ListOfOrderHeader(i) & "'")
            Dim Roworderheader As DataRow() = orderheaderdt.Select("P_OrderHeader = '" & ListOfOrderHeader(i) & "'")
            For j = 0 To Rowchargingheader.Length - 1
                Dim FinalRow As DataRow = FinalDt.NewRow
                If Rowchargingheader.Length = 1 Then
                    FinalRow.Item("RowType") = "single"
                Else
                    If j = 0 Then
                        FinalRow.Item("RowType") = "main"
                    Else
                        FinalRow.Item("RowType") = "sub"
                    End If
                End If
                FinalRow("ChargingHeader_Key") = Rowchargingheader(j).Item("ChargingHeader_Key")
                FinalRow("HeaderNo") = Rowchargingheader(j).Item("HeaderNo")
                FinalRow("P_Customers") = Rowchargingheader(j).Item("P_Customers")
                FinalRow("GrandTotal") = Rowchargingheader(j).Item("GrandTotal")
                FinalRow("orderheader") = Rowchargingheader(j).Item("orderheader")
                FinalRow("CustName") = Rowchargingheader(j).Item("CustName")
                FinalRow("CustCode") = Rowchargingheader(j).Item("CustCode")
                FinalRow("WT") = Rowchargingheader(j).Item("WT")
                FinalRow("OrderHeader_key") = Roworderheader(0).Item("OrderHeader_key")
                FinalRow("P_OrderHeader") = Roworderheader(0).Item("P_OrderHeader")
                FinalRow("OrderDate") = Roworderheader(0).Item("OrderDate")
                FinalRow("P_acccode") = Roworderheader(0).Item("P_acccode")
                FinalRow("LoginCode") = Roworderheader(0).Item("LoginCode")
                FinalRow("LoginType") = Roworderheader(0).Item("LoginType")
                FinalRow("Totalamount") = Roworderheader(0).Item("Totalamount")
                Dim TempDate As Date = Roworderheader(0).Item("OrderDate")
                FinalRow("FrmtOrderDate") = TempDate.ToString("dd-MM-yyyy hh:mm tt")
                FinalDt.Rows.Add(FinalRow)
            Next
        Next
        Return FinalDt
    End Function
    ''' <summary>
    ''' Function to Cancel whole Order according to P_OrderHeader provided.
    ''' </summary>
    ''' <param name="P_orderHeader">P_OrderHeader from OrderHeader table</param>
    Public Sub CancelOrderHeader(ByVal P_orderHeader As Integer, ByVal sessionRow As DataRow, ByVal p_acccode As Integer)
        Dim Success As Boolean = False
        Dim clsOrderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim clsChargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim clschargingitems As New ChargingItems.ChargingItems.ChargingItems
        Dim predtOrderHeader As New DataTable
        Dim predtChargingHeader As New DataTable
        Dim curdtOrderHeader As New DataTable
        Dim curdtChargingHeader As New DataTable
        Dim predtChargingItems As New DataTable
        Dim logintype As String = sessionRow("linktype")
        Dim logincode As Integer = sessionRow("linkcode")
        Dim currDate As DateTime = df1.getDateTimeISTNow
        predtOrderHeader = df1.GetDataFromSql(clsOrderHeader.ServerDatabase, clsOrderHeader.TableName, "*", "", "rowstatus = 0 and P_OrderHeader = " & P_orderHeader & "", "", "")
        predtChargingHeader = df1.GetDataFromSql(clsChargingHeader.ServerDatabase, clsChargingHeader.TableName, "*", "", "rowstatus = 0 and orderheader = " & P_orderHeader & "", "", "")
        ' predtChargingitems = df1.GetDataFromSql(clsChargingitems.ServerDatabase, clsChargingitems.TableName, "*", "", "orderheader = " & P_orderHeader & "", "", "")
        If predtOrderHeader.Rows.Count > 0 Then
            Dim prevorderRow As DataRow = predtOrderHeader.Rows(0)
            clsOrderHeader.PrevRow = df1.UpdateDataRows(clsOrderHeader.PrevRow, prevorderRow)
            clsOrderHeader.CurrRow.Item("logincode") = logincode
            clsOrderHeader.CurrRow.Item("logintype") = logintype
            clsOrderHeader.CurrRow.Item("p_acccode") = p_acccode
            clsOrderHeader.CurrRow.Item("mtimestamp") = currDate
            clsOrderHeader.CurrRow.Item("paymentflag") = "C"
            clsOrderHeader.CurrRow.Item("websessions_key") = sessionRow("websessions_key")
            Dim aclsobj() As Object = {clsOrderHeader}
            cfc.SaveIntodb(aclsobj)
        End If
        Dim clschrgngHeaderStr As String = ""
        For i = 0 To predtChargingHeader.Rows.Count - 1
            Dim clschrgingheader As New ChargingHeader.ChargingHeader.ChargingHeader
            clschrgingheader.PrevRow = df1.UpdateDataRows(clschrgingheader.PrevRow, predtChargingHeader.Rows(i))
            clschrgingheader.CurrRow.Item("logincode") = logincode 'sessionRow("")
            clschrgingheader.CurrRow.Item("logintype") = logintype ' sessionRow("")
            clschrgingheader.CurrRow.Item("p_acccode") = p_acccode
            clschrgingheader.CurrRow.Item("mtimestamp") = currDate
            clschrgingheader.CurrRow.Item("PaymentFlag") = "C"
            clschrgingheader.CurrRow.Item("websessions_key") = sessionRow("websessions_key")
            Dim aclsobj() As Object = {clschrgingheader}
            cfc.SaveIntodb(aclsobj)
            clschrgngHeaderStr = clschrgngHeaderStr & "," & predtChargingHeader.Rows(i).Item("headerno")
        Next
        'If clschrgngHeaderStr.First = "," Then clschrgngHeaderStr = clschrgngHeaderStr.Substring(1)

        'predtChargingItems = df1.GetDataFromSql(clschargingitems.ServerDatabase, clschargingitems.TableName, "*", "", "rowstatus=0 and headerno in ( " & clschrgngHeaderStr & ") and  orderheader = " & P_orderHeader & "", "", "")

        'For i = 0 To predtChargingItems.Rows.Count - 1
        '    Dim clschrgingitems As New ChargingItems.ChargingItems.ChargingItems
        '    clschrgingitems.PrevRow = df1.UpdateDataRows(clschrgingitems.PrevRow, predtChargingHeader.Rows(i))
        '    clschrgingitems.CurrRow.Item("PaymentFlag") = "C"
        '    clschrgingitems.CurrRow.Item("logincode") = logincode ' sessionRow("")
        '    clschrgingitems.CurrRow.Item("logintype") = logintype 'sessionRow("")
        '    clschrgingitems.CurrRow.Item("p_acccode") = p_acccode
        '    clschrgingitems.CurrRow.Item("mtimestamp") = currDate
        '    clschrgingitems.CurrRow.Item("websessions_key") = sessionRow("websessions_key")
        '    Dim aclsobj() As Object = {clschrgingitems}
        '    mdlRegistation.SaveIntodb(aclsobj, "")
        'Next
    End Sub
    ''' <summary>
    ''' function to create orderHeader row for payment from paymentdetails  page.
    ''' </summary>
    ''' <param name="TotalAmt">Amount to be paid</param>
    ''' <param name="sessionrow">session row contain details of login user</param>
    ''' <param name="p_acccode">p_acccode of login user</param>
    ''' <returns></returns>
    Public Function CreateOrderHeaderforPayment(ByVal TotalAmt As Decimal, ByVal sessionrow As DataRow, ByVal p_acccode As Integer) As Integer
        Dim clsOrderHeader As New OrderHeader.OrderHeader.OrderHeader
        Dim currdate As New DateTime
        currdate = df1.getDateTimeISTNow
        Dim logintype As String = sessionrow("linktype")
        Dim logincode As Integer = sessionrow("linkcode")
        Dim lorderheader As Integer = -1
        Dim webSessions_key As Integer = sessionrow("webSessions_key")
        clsOrderHeader.CurrRow.Item("p_customers") = -1
        clsOrderHeader.CurrRow.Item("orderdate") = df1.getDateTimeISTNow
        clsOrderHeader.CurrRow.Item("logintype") = logintype
        clsOrderHeader.CurrRow.Item("LoginCode") = logincode
        clsOrderHeader.CurrRow.Item("P_acccode") = p_acccode
        clsOrderHeader.CurrRow.Item("OrderSeries") = logintype & logincode
        clsOrderHeader.CurrRow.Item("paymentflag") = "U"
        clsOrderHeader.CurrRow.Item("totalamount") = TotalAmt
        clsOrderHeader.CurrRow.Item("calledFrom") = logintype & "D"
        clsOrderHeader.CurrRow.Item("mtimestamp") = currdate
        clsOrderHeader.CurrRow.Item("websessions_key") = webSessions_key
        Dim aClsObjectOH() As Object = {clsOrderHeader}
        Dim SuccessOH As Boolean = True
        Dim mserverdbOH As String = df1.GetServerMDFForTransanction(aClsObjectOH)
        Dim mytransOH As SqlTransaction = df1.BeginTransaction(mserverdbOH)
        Dim aLastKeysValuesOH As New Hashtable
        aClsObjectOH = df1.SetKeyValueIfNewInsert(mytransOH, aClsObjectOH)
        Dim HashPublicValuesOH As New Hashtable
        Dim sqlexecOH As Boolean = df1.CheckTableClassUpdations(aClsObjectOH)
        aClsObjectOH = df1.LastKeysPlus(mytransOH, aClsObjectOH, aLastKeysValuesOH)
        aClsObjectOH = df1.SetFinalFieldsValues(aClsObjectOH, HashPublicValuesOH)
        Dim aamOH As Integer = -1
        Try
            If sqlexecOH = True Then
                aamOH = df1.InsertUpdateDeleteSqlTables(mytransOH, aClsObjectOH, aamOH)
                Dim OrderHeaderhash As Hashtable = GF1.GetValueFromHashTable(aLastKeysValuesOH, "OrderHeader")
                lorderheader = GF1.GetValueFromHashTable(OrderHeaderhash, "p_orderheader")
                mytransOH.Commit()
                SuccessOH = True
            End If
        Catch ex As Exception
            mytransOH.Rollback()
            SuccessOH = False
        End Try
        mytransOH.Dispose()
        Return lorderheader
    End Function

#End Region
#Region "CSVM"
    ''' <summary>
    ''' This function prepares datatable with Relevant Columns which are to be shown in Customers Excel.
    ''' </summary>
    ''' <param name="dt">Datatable from Customers table according to conditions.</param>
    ''' <returns>datatable containing columns to be write in excel.</returns>
    Public Function GetDtForCustomerExcel(dt As DataTable) As DataTable
        Dim datatableMain As New DataTable
        datatableMain = df1.AddColumnsInDataTable(datatableMain, "S.no,CustCode,CustomerName,Address1,Address2,Address3,Address4,Pincode,MobileNo,PhoneNo,Email,HomeTown,BusinessType,Product,Lan,Nodes,FirstInstallDate,BilledUpto,OpenedUpto,LastRegDate,AmcMonthDate")
        If dt.Rows.Count > 0 Then
            'Dim Csvmt As New registrationTest
            For i = 0 To dt.Rows.Count - 1
                Dim aa As DataRow = datatableMain.NewRow
                aa.Item("S.no") = i + 1
                aa.Item("CustCode") = dt.Rows(i).Item("CustCode")
                aa.Item("CustomerName") = dt.Rows(i).Item("CustName")
                aa.Item("Address1") = dt.Rows(i).Item("PostalAddress1")
                aa.Item("Address2") = dt.Rows(i).Item("PostalAddress2")
                aa.Item("Address3") = dt.Rows(i).Item("PostalAddress3")
                aa.Item("Address4") = dt.Rows(i).Item("PostalAddress4")
                aa.Item("Pincode") = dt.Rows(i).Item("Pincode")
                aa.Item("MobileNo") = dt.Rows(i).Item("MobNo")
                aa.Item("PhoneNo") = dt.Rows(i).Item("Phone")
                aa.Item("Email") = dt.Rows(i).Item("Email")
                aa.Item("HomeTown") = dt.Rows(i).Item("TextHomeTown")
                aa.Item("BusinessType") = dt.Rows(i).Item("TextMainBussCode")
                aa.Item("Product") = dt.Rows(i).Item("TextProductCode")
                If IsDBNull(dt.Rows(i).Item("amcmonth")) = False Then
                    Dim amcmnth As String = dt.Rows(i).Item("amcmonth")
                    If String.IsNullOrEmpty(amcmnth) = False Then
                        Dim amcmnthwithoutM As String = amcmnth.Remove(0, 1)
                        Dim mnthdate As String = amcmnthwithoutM.Substring(2, 2) & amcmnthwithoutM.Substring(0, 2)
                        aa.Item("AmcMonthDate") = mnthdate
                    Else
                        aa.Item("AmcMonthDate") = amcmnth
                    End If
                Else
                    aa.Item("AmcMonthDate") = ""
                End If

                Dim Nodes As Integer = GetLanandNodeFromChargingItems(dt.Rows(i).Item("P_customers"))
                If Nodes = 0 Then
                    aa.Item("Lan") = "N"
                Else
                    aa.Item("Lan") = "Y"
                End If
                aa.Item("Nodes") = Nodes
                If IsDBNull(dt.Rows(i).Item("FirstInstallDate")) = False Then
                    Dim FirstInstallDate As Date = dt.Rows(i).Item("FirstInstallDate")
                    Dim StrFirstInstallDate As String = FirstInstallDate.ToString("dd-MM-yyyy")
                    aa.Item("FirstInstallDate") = StrFirstInstallDate
                Else
                    aa.Item("FirstInstallDate") = ""
                End If
                Dim ClsDealer As New Dealers.Dealers.Dealers
                Dim RegTran As DataTable = df1.GetDataFromSql(ClsDealer.ServerDatabase, "RegistrationTran", "regtype", "", "rowstatus = 0 and P_Customers =" & dt.Rows(i).Item("P_Customers"), "", "")
                Dim BilledUpto As Date
                If RegTran.Rows.Count > 0 Then
                    BilledUpto = libcustomerfeature.GetBilledUpToDate(dt.Rows(i).Item("P_Customers"), RegTran.Rows(0).Item("regtype"))
                End If
                ' Dim BilledUpto As Date = Csvmt.GetBilledUpToDate(dt.Rows(i).Item("P_Customers"),)
                Dim StrBilledUpto As String = BilledUpto.ToString("dd-MM-yyyy")
                aa.Item("BilledUpto") = StrBilledUpto
                'If IsDBNull(dt.Rows(i).Item("AllowUpto")) = False Then
                '    Dim AllowUpto As Date = dt.Rows(i).Item("AllowUpto")
                '    Dim StrAllowUpto As String = AllowUpto.ToString("dd-MM-yyyy")
                '    aa.Item("AllowUpto") = StrAllowUpto
                'Else
                '    aa.Item("AllowUpto") = ""
                'End If
                Dim OpenedUpto As Date = libcustomerfeature.GetOpenedUptoDate(dt.Rows(i).Item("P_Customers"))
                Dim StrOpenedUpto As String = OpenedUpto.ToString("dd-MM-yyyy")
                aa.Item("openedupto") = StrOpenedUpto
                If IsDBNull(dt.Rows(i).Item("CurrRegDate")) = False Then
                    Dim lastRegDate As Date = dt.Rows(i).Item("CurrRegDate")
                    Dim StrlastRegDate As String = lastRegDate.ToString("dd-MM-yyyy")
                    aa.Item("LastRegDate") = StrlastRegDate
                Else
                    aa.Item("LastRegDate") = ""
                End If
                'If IsDBNull(dt.Rows(i).Item("CeilingDate")) = False Then
                '    Dim CeilingDate As Date = dt.Rows(i).Item("CeilingDate")
                '    Dim StrCeilingDate As String = CeilingDate.ToString("dd-MM-yyyy")
                '    aa.Item("CeilingDate") = StrCeilingDate
                'Else
                '    aa.Item("CeilingDate") = ""
                'End If
                datatableMain.Rows.Add(aa)
            Next
        End If
        Return datatableMain
    End Function
    ''' <summary>
    ''' Function creates joined datatable of Chargingheader table and Customer table.  
    ''' </summary>
    ''' <param name="chargingheader">datatable containing data from chargingheader.</param>
    ''' <returns>Joined datatable with data from chargingheader and customer tables</returns>
    Public Function CreateJoinDT2(ByVal chargingheader As DataTable) As DataTable
        'customers = df1.GetDataFromSql("1_srv_1.0_mdf_0", "Customers")
        Dim CreateJoinDT As New DataTable
        CreateJoinDT = chargingheader.Clone
        CreateJoinDT = df1.AddColumnsInDataTable(CreateJoinDT, "customers_key,P_Customers,CustName")
        Dim k As New Customers.Customers.Customers
        For i = 0 To chargingheader.Rows.Count - 1
            Dim PcusValue As Integer = chargingheader.Rows(i).Item("P_Customers")
            Dim Customers As New DataTable
            Customers = df1.GetDataFromSql(k.ServerDatabase, k.TableName, "*", "", "P_Customers=" & PcusValue & " And rowstatus = 0", "", "")
            If Customers.Rows.Count > 0 Then
                Dim a As DataRow = CreateJoinDT.NewRow
                a("ChargingHeader_Key") = chargingheader.Rows(i).Item("ChargingHeader_Key")
                a("HeaderNo") = chargingheader.Rows(i).Item("HeaderNo")
                a("P_Customers") = chargingheader.Rows(i).Item("P_Customers")
                a("RowStatus") = chargingheader.Rows(i).Item("RowStatus")
                a("BillDate") = chargingheader.Rows(i).Item("BillDate")
                a("BillType") = chargingheader.Rows(i).Item("BillType")
                a("BillSeries") = chargingheader.Rows(i).Item("BillSeries")
                a("BillNo") = chargingheader.Rows(i).Item("BillNo")
                a("ProfarmaNo") = chargingheader.Rows(i).Item("ProfarmaNo")
                a("ProfarmaDate") = chargingheader.Rows(i).Item("ProfarmaDate")
                a("DebitedCode") = chargingheader.Rows(i).Item("DebitedCode")
                a("EmployeeCode") = chargingheader.Rows(i).Item("EmployeeCode")
                a("PaymentFlag") = chargingheader.Rows(i).Item("PaymentFlag")
                a("TextPaymentFlag") = chargingheader.Rows(i).Item("TextPaymentFlag")
                a("HandleBy") = chargingheader.Rows(i).Item("HandleBy")
                a("GrandTotal") = chargingheader.Rows(i).Item("GrandTotal")
                a("BillinSession") = chargingheader.Rows(i).Item("BillinSession")
                a("WebSessions_key") = chargingheader.Rows(i).Item("WebSessions_key")
                a("RoundOffAmt") = chargingheader.Rows(i).Item("RoundOffAmt")
                a("custname") = Customers.Rows(0).Item("custname")
                a("customers_key") = Customers.Rows(0).Item("customers_key")
                CreateJoinDT.Rows.Add(a)
            End If
        Next
        Return CreateJoinDT
    End Function
    ''' <summary>
    ''' Prepares the CustomerRegistration details datatable for each dealer to be sent trrough email or excel.
    ''' </summary>
    ''' <param name="fromDate"></param>
    ''' <param name="Todate"></param>
    ''' <param name="Fullfilepath"></param>
    ''' <param name="dealerRow"></param>
    ''' <returns></returns>
    Public Function RegistrationsCreatedReport(fromDate As Date, Todate As Date, Fullfilepath As String, dealerRow As DataRow) As DataTable
        If fromDate = Nothing Then fromDate = df1.getDateTimeISTNow
        If Todate = Nothing Then Todate = df1.getDateTimeISTNow
        Dim FromdateStr As String = fromDate.ToString("MM-dd-yyyy 00:00:00.00")
        Dim TodateStr As String = Todate.ToString("MM-dd-yyyy 23:59:59.999")
        Dim filename As String = Path.GetFileName(Fullfilepath)
        Dim filepath As String = Path.GetDirectoryName(Fullfilepath)
        If Not Directory.Exists(filepath) Then
            Directory.CreateDirectory(filepath)
        Else
            Dim di As System.IO.DirectoryInfo = New DirectoryInfo(filepath)
            For Each file As FileInfo In di.GetFiles()
                file.Delete()
            Next
        End If
        Dim emailId As String = dealerRow.Item("Email")
        Dim Query As String = String.Format("Select Customers.CustCode,Customers.CustName, InfoTable.NameOfInfo as HomeTown,RegistrationTran.RegSendDate,RegistrationTran.RegType,RegistrationTran.RegType2,RegistrationTran.OpenedUpto,RegistrationTran.Lan,RegistrationTran.Node,RegistrationTran.P_customers
    from
        RegistrationTran RegistrationTran
            inner join
        Customers Customers
            on Customers.P_Customers  = RegistrationTran.P_customers 
      inner join 
        InfoTable InfoTable
            on Customers.HomeTown = InfoTable.P_InfoTable      
    where
    RegistrationTran.RowStatus=0 and Customers.RowStatus=0 and (RegistrationTran.RegsendDate between '" & FromdateStr & "' and '" & TodateStr & "') and Customers.ServicingAgentCode=" & dealerRow.Item("P_acccode") & "
    Order By
    RegistrationTran.RegsendDate asc")
        Dim ClsChargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim RegistrationDetails As DataTable = df1.SqlExecuteDataTable(ClsChargingHeader.ServerDatabase, Query)
        RegistrationDetails = df1.AddColumnsInDataTable(RegistrationDetails, "S.no",,, "CustCode")
        RegistrationDetails = df1.AddColumnsInDataTable(RegistrationDetails, "PaymentApplicable")
        If RegistrationDetails.Rows.Count > 0 Then
            For i = 0 To RegistrationDetails.Rows.Count - 1
                RegistrationDetails.Rows(i).Item("S.no") = i + 1
                Dim chdt As DataTable = df1.GetDataFromSql(ClsChargingHeader.ServerDatabase, ClsChargingHeader.TableName, "*", "", "RowStatus=0 and P_Customers=" & RegistrationDetails.Rows(i).Item("P_customers") & " and (BillDate between '" & FromdateStr & "' and '" & TodateStr & "')", "", "")
                If chdt.Rows.Count > 0 Then
                    RegistrationDetails.Rows(i).Item("PaymentApplicable") = "Yes"
                Else
                    RegistrationDetails.Rows(i).Item("PaymentApplicable") = "No"
                End If
            Next
            RegistrationDetails.Columns.Remove("P_customers")
        End If
        Return RegistrationDetails
    End Function
    ''' <summary>
    ''' Prepares the CustomerRegistration details datatable for each dealer to be sent trrough email or excel.
    ''' </summary>
    ''' <param name="fromDate"></param>
    ''' <param name="Todate"></param>
    ''' <param name="Fullfilepath"></param>
    ''' <param name="dealerRow"></param>
    ''' <returns></returns>
    Public Function OrderStatementReport(fromDate As Date, Todate As Date, Fullfilepath As String, dealerRow As DataRow, DtInfoTable As DataTable) As DataTable
        If fromDate = Nothing Then
            fromDate = df1.getDateTimeISTNow
        End If
        If Todate = Nothing Then
            Todate = df1.getDateTimeISTNow
        End If
        Dim FromdateStr As String = fromDate.ToString("yyyy-MM-dd 00:00:00.00")
        Dim TodateStr As String = Todate.ToString("yyyy-MM-dd 23:59:59.999")
        Dim filename As String = Path.GetFileName(Fullfilepath)
        Dim filepath As String = Path.GetDirectoryName(Fullfilepath)
        If Not Directory.Exists(filepath) Then
            Directory.CreateDirectory(filepath)
        Else
            Dim di As System.IO.DirectoryInfo = New DirectoryInfo(filepath)
            For Each file As FileInfo In di.GetFiles()
                file.Delete()
            Next
        End If
        Dim emailId As String = df1.GetCellValue(dealerRow, "Email")
        '  Dim P_Dealer As Integer = dealerRow.Item("P_Dealers")
        Dim p_acccode As Integer = dealerRow.Item("P_acccode")
        Dim Combineddt As New DataTable
        Dim ClsChargingHeader As New ChargingHeader.ChargingHeader.ChargingHeader
        Dim ChargingHeaderdt As DataTable = df1.GetDataFromSql(ClsChargingHeader.ServerDatabase, ClsChargingHeader.TableName, "*", "", "rowstatus = 0 and PaymentFlag='P' and p_acccode = " & p_acccode & " and (mtimestamp between '" & FromdateStr & "' and '" & TodateStr & "')", "", "billdate")
        'Combineddt = ChargingHeaderdt.Copy
        Dim P_Customers As New List(Of Integer)
        For i = 0 To ChargingHeaderdt.Rows.Count - 1
            Dim tempP_Customers As Integer = ChargingHeaderdt.Rows(i).Item("P_Customers")
            If P_Customers.Contains(tempP_Customers) Then
            Else
                P_Customers.Add(ChargingHeaderdt.Rows(i).Item("P_Customers"))
            End If
        Next
        Dim ClsCustomers As New Customers.Customers.Customers
        df1.AddColumnsInDataTable(Combineddt, "S.NO,BillDate,CustCode,CustName,HomeTown,AllowUpto,BilledUpto,Amount")
        If P_Customers IsNot Nothing Then
            For i = 0 To P_Customers.Count - 1
                Dim Customersdt As DataTable = df1.GetDataFromSql(ClsCustomers.ServerDatabase, ClsCustomers.TableName, "*", "", "rowstatus=0 and P_Customers=" & P_Customers(i) & "", "", "")
                Customersdt = df1.AddingNameForCodesPrimamryCols(Customersdt, "HomeTown", "TextHomeTown", DtInfoTable, "NameOfInfo")
                For j = 0 To ChargingHeaderdt.Rows.Count - 1
                    If P_Customers(i) = ChargingHeaderdt.Rows(j).Item("P_Customers") Then
                        Dim a As DataRow = Combineddt.NewRow
                        a.Item("S.no") = i + 1
                        a.Item("CustCode") = df1.GetCellValue(Customersdt(0), "CustCode", "string")
                        a.Item("CustName") = df1.GetCellValue(Customersdt(0), "Custname", "string") 'Customersdt(0).Item("CustName")
                        a.Item("AllowUpto") = df1.GetCellValue(Customersdt(0), "Allowupto", "datetime") 'Customersdt(0).Item("AllowUpto")
                        a.Item("HomeTown") = df1.GetCellValue(Customersdt(0), "TextHomeTown", "string")
                        a.Item("BilledUpto") = ChargingHeaderdt(j).Item("BilledUpto")
                        a.Item("Amount") = ChargingHeaderdt(j).Item("GrandTotal")
                        ' a.Item("WithTax") = ChargingHeaderdt(j).Item("WT")
                        a.Item("BillDate") = ChargingHeaderdt(j).Item("billdate")
                        Combineddt.Rows.Add(a)
                        ' Exit For
                    End If
                Next
            Next
        End If
        Return Combineddt
    End Function
    ''' <summary>
    ''' Function creates the Email structure i.e,Subject,Body title etc. from datatable and sends the email to the desired email id.
    ''' </summary>
    ''' <param name="RegistrationDetails">Registrationdetails datatable from which email is to be prepared</param>
    ''' <param name="fromDate"></param>
    ''' <param name="emailId">Email id on which email is to be sent.</param>
    ''' <param name="attachmentfilepath">Full path of excel file to be attached.</param>
    ''' <returns></returns>
    Public Function SendEmailOfRegistrationDetails(RegistrationDetails As DataTable, fromDate As Date, emailId As String, attachmentfilepath As String) As Boolean
        Dim b As String = ""
        b = "<html><head></head><body><h2 align='center'>Details of Registation Opened On " & fromDate.ToString("dd-MM-yyyy") & "</h2>"
        If RegistrationDetails.Rows.Count <= 0 Then
            b += "<p>No Registration opened today</p>"
        Else
            b += "<table style='width:100%;'><thead><tr><th align='left'>S no</th><th align='left'>CustomerName</th><th align='left'>CustomerCode</th><th align='left'>RegSendDate</th><th align='left'>RegType</th><th align='left'>RegType2</th><th align='left'>OpenedUpto</th><th align='left'>Lan</th><th align='left'>Node</th><th align='left'>PaymentApplicable</th>"
            b += "</tr></thead><tbody>"
            For j = 0 To RegistrationDetails.Rows.Count - 1
                b += "<tr style='border-bottom: 1px solid black;'>"
                b += "<td style='width:5%;'>" & j + 1 & "</td>"
                b += "<td style='width:25%;'>" & RegistrationDetails.Rows(j).Item("CustName").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("CustCode").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("RegSendDate").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("RegType").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("RegType2").ToString & "</td>"
                Dim OpenedUpto As Date = RegistrationDetails.Rows(j).Item("OpenedUpto")
                Dim OpenedUptostr As String = OpenedUpto.ToString("dd-MM-yyyy")
                b += "<td style='width:10%;'>" & OpenedUptostr & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("Lan").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("Node").ToString & "</td>"
                b += "<td style='width:10%;'>" & RegistrationDetails.Rows(j).Item("PaymentApplicable").ToString & "</td>"
                b += "</tr>"
            Next
            b += "</tbody></table></body></html>"
        End If
        Dim subject As String = "Registrations Opened: " & fromDate.ToString("dd-MM-yyyy")
        Dim email As String = emailId & ", nehagupta@saralerp.com, hcgupta@saralerp.com,93nishayadav@gmail.com"
        'Dim email As String = "93nishayadav@gmail.com"
        If String.IsNullOrEmpty(email.ToString.Trim) Then
        Else
            GF2.SendingEmail(email.ToString.Trim, subject, b, attachmentfilepath)
        End If
    End Function
    Public Function GetDealerStatement_new(ByVal p_acccode As Integer, ByVal FromDate As DateTime, ByVal ToDate As DateTime, ByVal WT1 As String) As DataTable
        Dim clsRegistrations As New Registrations.Registrations.Registrations
        Dim strPayment As String = ""
        If WT1 = "Y" Then
            ' strPayment = " Select *, m2.nameofinfo as textbenaccount, m3.nameofinfo as textpaymentmode from payment m1  inner join infotable m2 on m1.benaccount = m2.p_infotable  inner join infotable m3 on m1.paymentmode = m3.p_infotable     where benaccount in (2751,3033) and  m1.p_acccode = " & p_acccode & " And m1.rowstatus = 0 And m1.verifycode = 'V' and m1.verifydate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.verifydate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "' and paymentdate >= '2019-07-01 00:00:00.000' order by verifydate asc"
            strPayment = " Select * from payment m1   where benaccount in (2751,3033) and  m1.p_acccode = " & p_acccode & " And m1.rowstatus = 0 And m1.verifycode = 'V' and m1.verifydate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.verifydate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.paymentdate >= '2019-07-01 00:00:00.000' order by verifydate asc"
        ElseIf WT1 = "N" Then
            '   strPayment = " Select *, m2.nameofinfo as textbenaccount, m3.nameofinfo as textpaymentmode from payment m1  inner join infotable m2 on m1.benaccount = m2.p_infotable  inner join infotable m3 on m1.paymentmode = m3.p_infotable     where benaccount <> 2751 and benaccount <> 3033 and  m1.p_acccode = " & p_acccode & " And m1.rowstatus = 0 And m1.verifycode = 'V' and m1.verifydate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.verifydate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "'  and paymentdate >= '2019-07-01 00:00:00.000' order by verifydate asc"
            strPayment = " Select * from payment m1   where benaccount <> 2751 and benaccount <> 3033 and  m1.p_acccode = " & p_acccode & " And m1.rowstatus = 0 And m1.verifycode = 'V' and m1.verifydate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.verifydate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "'  and m1.paymentdate >= '2019-07-01 00:00:00.000' order by verifydate asc"
        End If
        Dim strCharging As String = ""
        If WT1 = "Y" Then
            ' strCharging = "select *,m2.custname as custname ,m2.custcode as custcode from chargingheader m1 inner join customers m2 on m1.p_customers = m2.p_customers where m1.paymentflag = 'P' and WT='Y' and m1.p_acccode = " & p_acccode & " and m1.rowstatus = 0  and m1.billdate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.billdate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m2.rowstatus = 0 order by m1.billdate asc"
            strCharging = "select * from chargingheader m1  where m1.paymentflag = 'P' and WT='Y' and m1.p_acccode = " & p_acccode & " and m1.rowstatus = 0  and m1.billdate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.billdate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "'  order by m1.billdate asc"
        ElseIf WT1 = "N" Then
            '  strCharging = "select *,m2.custname as custname ,m2.custcode as custcode from chargingheader m1 inner join customers m2 on m1.p_customers = m2.p_customers where m1.paymentflag = 'P' and WT='N' and m1.p_acccode = " & p_acccode & " and m1.rowstatus = 0  and m1.billdate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.billdate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m2.rowstatus = 0 order by m1.billdate asc"
            strCharging = "select * from chargingheader m1 where m1.paymentflag = 'P' and WT='N' and m1.p_acccode = " & p_acccode & " and m1.rowstatus = 0  and m1.billdate >= '" & FromDate.ToString("yyyy-MM-dd H:mm:ss") & "' and m1.billdate <='" & ToDate.ToString("yyyy-MM-dd H:mm:ss") & "' order by m1.billdate asc"
        End If
        Dim dtpayment As DataTable = df1.SqlExecuteDataTable(clsRegistrations.ServerDatabase, strPayment)
        dtpayment = df1.AddColumnsInDataTable(dtpayment, "textbenaccount,textpaymentmode")
        Dim dtcharging As DataTable = df1.SqlExecuteDataTable(clsRegistrations.ServerDatabase, strCharging)

        dtcharging = df1.AddColumnsInDataTable(dtcharging, "custcode,custname")
        Dim DtStatement As New DataTable
        '   DtStatement = df1.AddColumnsInDataTable(DtStatement, "sno,postingdate,entrytype,entryid,creditamt,debitamt,narra1,narra2,narra3", "system.int32,system.datetime,system.string,system.int32,system.decimal,system.decimal,system.string,system.string,system.string")

        DtStatement = df1.AddColumnsInDataTable(DtStatement, "Sno,PostingDate,Particulars,CreditAmt,DebitAmt,Balance", "system.int32,system.datetime,system.string,system.decimal,system.decimal,system.decimal")

        Dim openBal As Decimal = GetLedgerBalanceOnDate(p_acccode, FromDate.ToString("yyyy-MM-dd H:mm:ss"), WT1)  'New Date(2019, 7, 1), "Y")
        DtStatement.Rows.Add()
        ' DtStatement.Rows(DtStatement.Rows.Count - 1).Item("entrytype") = "OpeningBalance"
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("creditamt") = openBal
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("Particulars") = "Balance Brought Forward " 'as on " & FromDate.ToString("yyyy-MM-dd H:mm:ss") ' New Date(2019, 7, 1).ToString("yyyy-MM-dd")
        For l = 0 To dtpayment.Rows.Count - 1
            Dim dtr As DataRow = DtStatement.NewRow
            dtr.Item("sno") = l + 2
            dtr.Item("postingdate") = df1.GetCellValue(dtpayment.Rows(l), "verifydate")
            Dim benaccount As Int32 = df1.GetCellValue(dtpayment.Rows(l), "benaccount")
            Dim paymentmode As Int32 = df1.GetCellValue(dtpayment.Rows(l), "paymentmode")
            Dim dtinfo1 As DataTable = df1.SqlExecuteDataTable(clsRegistrations.ServerDatabase, "select nameofinfo from infotable where p_infotable =" & benaccount)

            If dtinfo1.Rows.Count > 0 Then dtpayment.Rows(l).Item("textbenaccount") = df1.GetCellValue(dtinfo1.Rows(0), "nameofinfo", "string") ', "string"
            Dim dtinfo2 As DataTable = df1.SqlExecuteDataTable(clsRegistrations.ServerDatabase, "select nameofinfo from infotable where p_infotable =" & paymentmode)
            If dtinfo2.Rows.Count > 0 Then dtpayment.Rows(l).Item("textpaymentmode") = dtinfo2.Rows(0).Item("nameofinfo") ', "string"

            '  dtr.Item("entryid") = df1.GetCellValue(dtpayment.Rows(l), "p_payment")
            dtr.Item("creditamt") = df1.GetCellValue(dtpayment.Rows(l), "amount")
            Dim custcode As String = df1.GetCellValue(dtpayment.Rows(l), "custcode", "string").ToString.Trim
            ' dtr.Item("narra2") = df1.GetCellValue(dtpayment.Rows(l), "")
            Dim textBenAccount As String = df1.GetCellValue(dtpayment.Rows(l), "textbenaccount", "string").ToString.Trim
            Dim textpaymentmode As String = df1.GetCellValue(dtpayment.Rows(l), "textpaymentmode", "string").ToString.Trim
            If LCase(textpaymentmode) = "maingroup" Then textpaymentmode = ""

            Dim chqDDno As String = df1.GetCellValue(dtpayment.Rows(l), "chqDDNo", "string")
            If chqDDno.Trim = "" Then chqDDno = "" Else chqDDno = "ChqDDno=" & chqDDno
            Dim proceedings As String = df1.GetCellValue(dtpayment.Rows(l), "proceedings", "string").ToString.Trim
            textpaymentmode = Trim(textpaymentmode)
            dtr.Item("Particulars") = textBenAccount & IIf(textpaymentmode = "", "", " ; " & textpaymentmode) & IIf(chqDDno = "", "", ";" & chqDDno) & IIf(custcode = "", "", ";" & custcode) & IIf(proceedings = "", "", ";" & proceedings)
            DtStatement.Rows.Add(dtr)
            If df1.GetCellValue(dtpayment.Rows(l), "discount") <> 0 Then
                Dim dtr1 As DataRow = DtStatement.NewRow
                dtr1.Item("sno") = l + 1
                dtr1.Item("postingdate") = df1.GetCellValue(dtpayment.Rows(l), "verifydate")
                dtr1.Item("creditamt") = df1.GetCellValue(dtpayment.Rows(l), "discount")
                dtr1.Item("Particulars") = "Discount " & df1.GetCellValue(dtpayment.Rows(l), "p_payment")
                DtStatement.Rows.Add(dtr1)
            End If
        Next
        For u = 0 To dtcharging.Rows.Count - 1
            Dim dtr As DataRow = DtStatement.NewRow
            dtr.Item("sno") = u + 1 + dtpayment.Rows.Count
            dtr.Item("postingdate") = df1.GetCellValue(dtcharging.Rows(u), "billdate")
            Dim p_customers As Integer = df1.GetCellValue(dtcharging.Rows(u), "p_customers")
            Dim dtcust As DataTable = df1.SqlExecuteDataTable(clsRegistrations.ServerDatabase, "select custcode,custname from customers where rowstatus = 0 and p_customers=" & p_customers)
            If dtcust.Rows.Count > 0 Then
                dtcharging.Rows(u).Item("custcode") = df1.GetCellValue(dtcust.Rows(0), "custcode")
                dtcharging.Rows(u).Item("custname") = df1.GetCellValue(dtcust.Rows(0), "custname")
            End If
            dtr.Item("debitamt") = df1.GetCellValue(dtcharging.Rows(u), "roundoffamt")
            Dim dtbillpayflg As DataTable = df1.SqlExecuteDataTable(clsRegistrations.ServerDatabase, "select p_payment from billpayflag where headerno=" & df1.GetCellValue(dtcharging.Rows(u), "headerno"))
            Dim p_paymentstr As String = ""
            For t = 0 To dtbillpayflg.Rows.Count - 1
                p_paymentstr = p_paymentstr & "," & df1.GetCellValue(dtbillpayflg.Rows(t), "p_payment")
            Next
            If p_paymentstr = "" Then
                p_paymentstr = "null"
            Else
                If p_paymentstr.First = "," Then p_paymentstr = p_paymentstr.Substring(1)
            End If
            dtr.Item("Particulars") = df1.GetCellValue(dtcharging.Rows(u), "custname", "string").ToString.Trim & " (" & df1.GetCellValue(dtcharging.Rows(u), "custcode", "string") & ")" '&   "Adjusted by payment entry id:" & p_paymentstr
            DtStatement.Rows.Add(dtr)
        Next
        Dim dtview As New DataView(DtStatement)
        dtview.Sort = "postingdate asc"
        DtStatement = dtview.ToTable
        Dim crdAmt As Decimal = 0.0
        Dim dbtAmt As Decimal = 0.0
        For k = 0 To DtStatement.Rows.Count - 1
            DtStatement.Rows(k).Item("sno") = k + 1
            Dim j As Integer = 0
            If k = 0 Then j = 0 Else j = k - 1
            If k = 0 Then
                DtStatement.Rows(k).Item("Balance") = df1.GetCellValue(DtStatement.Rows(k), "creditamt", "decimal") - df1.GetCellValue(DtStatement.Rows(k), "debitAmt", "decimal")
            Else
                DtStatement.Rows(k).Item("Balance") = df1.GetCellValue(DtStatement.Rows(k - 1), "balance", "decimal") + df1.GetCellValue(DtStatement.Rows(k), "creditamt", "decimal") - df1.GetCellValue(DtStatement.Rows(k), "debitAmt", "decimal")
            End If
            crdAmt = crdAmt + df1.GetCellValue(DtStatement.Rows(k), "creditamt")
            dbtAmt = dbtAmt + df1.GetCellValue(DtStatement.Rows(k), "debitAmt")
        Next

        DtStatement.Rows.Add()
        '   DtStatement.Rows(DtStatement.Rows.Count - 1).Item("entrytype") = "ToTal Amount"
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("creditamt") = crdAmt 'GetLedgerBalanceOnDate(3, New Date(2019, 7, 1), "Y")
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("debitamt") = dbtAmt
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("Particulars") = "Total" '& New Date(2019, 7, 1).ToString("yyyy-MM-dd")

        DtStatement.Rows.Add()
        ' DtStatement.Rows(DtStatement.Rows.Count - 1).Item("entrytype") = "Net"
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("creditamt") = crdAmt - dbtAmt
        DtStatement.Rows(DtStatement.Rows.Count - 1).Item("Particulars") = "NetBalance as on " & ToDate.ToString("yyyy-MM-dd H:mm:ss")

        Return DtStatement
    End Function
    Public Function GetSearchStringForVerifyCustomer(search As String) As String
        Dim lcondition As String = ""
        If search <> "" And search <> "null" Then
            Dim search2() As String = search.Split(":")
            If search2(1) = "date" Then
                Dim search1 = search2(0).Split(",")
                Dim searchColumn As String = search1(2)
                If search1(0) <> "" And search1(1) <> "" Then
                    Dim min As Date = CDate(search1(0))
                    Dim max As Date = CDate(search1(1))
                    lcondition = lcondition & " and m2." & searchColumn & " >= '" & min.ToString("MM-dd-yyyy") & "' and m2." & searchColumn & "<='" & max.ToString("MM-dd-yyyy") & "'"
                ElseIf search1(0) = "" And search1(1) <> "" Then
                    Dim max As Date = CDate(search1(1))
                    lcondition = lcondition & " and m2." & searchColumn & "<='" & max.ToString("MM-dd-yyyy") & "'"
                ElseIf search1(0) <> "" And search1(1) = "" Then
                    Dim min As Date = CDate(search1(0))
                    lcondition = lcondition & " and m2." & searchColumn & " >= '" & min.ToString("MM-dd-yyyy") & "'"
                ElseIf search1(0) = "" And search1(1) = "" Then
                End If
            ElseIf search2(1) = "string" Then
                Dim search1 = search2(0).Split(",")
                Dim searchValue As String = LCase(search1(0))
                Dim searchColumn As String = search1(1)
                lcondition = lcondition & " and m2." & searchColumn & " Like '" & searchValue & "%'"
            ElseIf search2(1) = "integer" Then
                Dim search1 = search2(0).Split(",")
                Dim searchValue As String = search1(0)
                Dim searchColumn As String = search1(1)
                lcondition = lcondition & " and m2." & searchColumn & " = " & searchValue
            End If
        End If
        Return lcondition
    End Function
    ''' <summary>
    ''' This function get data for CustomerVerificationgrid  from db
    ''' </summary>
    ''' <param name="lcondition">lcondition</param>
    ''' <returns>dt</returns>
    Public Function GetCustomerVerificationDataCount(lcondition As String) As Integer
        Dim clsaccmaster As New Accmaster.Accmaster.Accmaster
        Dim Query As String = String.Format("select count(*) as RCount
from customerverification m1
inner join customers m2 on m1.p_customers = m2.p_customers
where m2.rowstatus = 0 and m1.status = 'P' and m1.rowstatus = 0 " & lcondition)
        Dim dt As DataTable = df1.SqlExecuteDataTable(clsaccmaster.ServerDatabase, Query)
        Return dt.Rows(0).Item("RCount")
    End Function
    ''' <summary>
    ''' This function get data for CustomerVerificationgrid  from db
    ''' </summary>
    ''' <param name="start">start no of row</param>
    ''' <param name="DtInfoTable">dt containing data from Infotable</param>
    ''' <param name="lcondition">lcondition</param>
    ''' <param name="pSize">pagesize</param>
    ''' <returns>dt</returns>
    Public Function GetCustomerVerificationData(start As Integer?, DtInfoTable As DataTable, lcondition As String, Optional pSize As Integer = 20) As DataTable
        Dim clsaccmaster As New Accmaster.Accmaster.Accmaster
        lcondition = " m2.rowstatus = 0 and m1.status = 'P' and m1.rowstatus = 0 " & lcondition
        Dim ljoin As String = "inner join customers m2 on m1.p_customers = m2.p_customers"
        Dim dt As DataTable = df1.GetDataFromSqlFixedRows(clsaccmaster.ServerDatabase, "CustomerVerification", "m1.customerverification_key as customerverification_key,m1.p_customerverification as p_customerverification , m1.regtran_key as regtran_key, m1.logincode as logincode ,m1.logintype as logintype ,m2.customers_key as customers_key,m2.p_customers as p_customers,m2.custname as CustName,m2.custcode as CustCode,m2.mobno as mobno ,m2.hometown as hometown,m2.mainbusscode as mainbusscode,m2.servicingagentcode as servicingagentcode", ljoin, lcondition, "", "customerverification_key desc", start, pSize, -1)
        dt = df1.AddColumnsInDataTable(dt, "LoginName,regtype,regtype2,dealername,regsenddate,allowuptodate", "system.string,system.string,system.string,system.string,system.datetime,system.datetime")
        Dim P_Customers As String = ""
        For i = 0 To dt.Rows.Count - 1
            Dim regtrandt1 As DataTable = df1.SqlExecuteDataTable(clsaccmaster.ServerDatabase, "select regtype,regtype2,regsenddate,allowuptodate from registrationtran where registrationTran_key =" & df1.GetCellValue(dt.Rows(i), "regtran_key"))
            If regtrandt1.Rows.Count > 0 Then
                dt.Rows(i).Item("regtype") = regtrandt1.Rows(0).Item("regtype")
                dt.Rows(i).Item("regtype2") = regtrandt1.Rows(0).Item("regtype2")
                dt.Rows(i).Item("regsenddate") = regtrandt1.Rows(0).Item("regsenddate")
                dt.Rows(i).Item("allowuptodate") = regtrandt1.Rows(0).Item("allowuptodate")
                Dim libSaralAuth As New SaralAuthLib.LoginFunctions
                Dim mDealer As DataRow = libSaralAuth.getAccMasterRowForp_acccode(df1.GetCellValue(dt.Rows(i), "ServicingAgentCode", "integer")) 'df1.SeekRecord(clsaccmaster, )
                If Not mDealer Is Nothing Then
                    dt.Rows(i).Item("DealerName") = mDealer("AccName").ToString.Trim
                End If
                'If dt.Rows(i).Item("logintype") = "D" Then
                'Dim DealerRow As DataRow = df1.SeekRecord(ClsDealer, df1.GetCellValue(dt.Rows(i), "logincode"))
                'dt.Rows(i).Item("LoginName") = DealerRow("DealerName").ToString.Trim
                'Else
                Dim empRow As DataRow = libSaralAuth.getUserLoginRowFromLinkcodeLinktype(df1.GetCellValue(dt.Rows(i), "logincode"), df1.GetCellValue(dt.Rows(i), "logintype"))
                '  Dim EmpRow As DataRow = df1.SeekRecord(clsaccmaster, df1.GetCellValue(dt.Rows(i), "logincode"))
                If Not empRow Is Nothing Then
                    dt.Rows(i).Item("LoginName") = empRow("Name").ToString.Trim
                End If
            End If
        Next
        dt = df1.AddingNameForCodesPrimamryCols(dt, "HomeTown,MainBussCode", "TextHomeTown,TextMainBussCode", DtInfoTable, "NameOfInfo")
        Return dt
    End Function

#End Region
End Class
