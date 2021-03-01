﻿
Imports Microsoft.VisualBasic
Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Forms.Screen
Imports System.Drawing

Public Class Variables
    Private Shared _FormsHeader As New DataTable
    Private Shared _RegistryFolder As String = "SOFTWARE\\SaralERP"
    Private Shared _AppControlHashTable As Hashtable = InitializeAppControl()
    Private Shared _MDFFiles As Hashtable = InitializeMdffiles()
    Private Shared _AllServers As Hashtable = InitializeAllServers()
    Private Shared _EventLogger As Boolean = False
    Private Shared _Language As DataRow = Nothing
    Private Shared _ColorScheme As New DataTable
    Private Shared _FontScheme As New DataTable
    Private Shared _UserRow As DataRow = Nothing
    Private Shared _BusinessFirmRow As DataRow = Nothing
    Private Shared _MDIHeight As Integer = 0
    Private Shared _MDIWidth As Integer = 0
    Private Shared _xLocalResolution As Integer = My.Computer.Screen.PrimaryScreen.Bounds.Width
    Private Shared _yLocalResolution As Integer = My.Computer.Screen.PrimaryScreen.Bounds.Height
    Public Shared xBaseResolution As Integer = 1360
    Public Shared yBaseResolution As Integer = 768
    Public Shared mmTopixels As Single = 3.77952755905511
    Public Shared _SqlVersion As Integer = 2012
    Private Shared _xFactor As Decimal = CalcXFactor()  'To get x-pixel multiplying ratio of x-local/x-base
    Private Shared _yFactor As Decimal = CalcYFactor()  'To get y-pixel multiplying ratio of y-local/y-base
    Private Shared _ActiveForms As New Hashtable
    Private Shared _AuthenticationChecked As Integer = 0
    Private Shared _ControlsList As New DataTable
    Private Shared _MasterOptions As New DataTable
    Private Shared _MenusList As New DataTable
    Private Shared _SaralType As String = ""

    Private Shared _LocalHostNo As String = ""
    Private Shared _EmailId As String = "saraluser2012@gmail.com"
    Private Shared _EmailPassword As String = "saral12345"
    Private Shared _LocalSMTPServerHost As String = "smtp.gmail.com"
    Private Shared _LocalSMTPServerPort As Integer = 587
    Private Shared _WebSMTPServerPort As Integer = 25
    Private Shared _LocalMTPServerEnableSsl As Boolean = True
    Private Shared _RunningAtWeb As Boolean = False
    Private Shared _WebSMTPServerEnableSsl As Boolean = False
    Private Shared _WebEmailId As String = "info@saralerp.com"
    Private Shared _WebEmailPwd As String = "123456"
    Private Shared _WebSMTPServerHost As String = "relay-hosting.secureserver.net"
    Private Shared _WebHostingUserName As String = "SaralWeb"
    Private Shared _WebHostingUserPassword As String = "Vg12345678#"
    Private Shared _DataFolderServerPhysicalPath As String = ""
    Private Shared _WebHostingServer As String = "182.50.133.109"
    Private Shared _DelayImage As Image = Nothing 'GetDelayImage()
    Private Shared _ErrorString As String = ""
    Public Shared StructureColumnsString As String = "TableName,PrimaryKey,FieldName,FieldType,Nullable,DefaultValue"
    Private Shared _TablesExcelControl As String = ""
    Private Shared _FieldsExcelControl As String = ""
    Public Shared xControl As Integer = 280196
    Private Shared _AllowDate As Date = #1/1/1000#
    Private Shared _DemoDate As Date = #1/1/1000#
    Private Shared _CustomKeyBoardValues As New Hashtable
    Private Shared _ObjectsHashTable As Hashtable = InitializeObjectsHashTable()
    Private Shared _TypeHashTable As Hashtable = InitializeTypesHashTable()
    Private Shared _ConversionTypeHashTable As Hashtable = InitializeConversionHashTable()
    Private Shared _entFormStru As DataTable = Nothing
    Private Shared _LastProcessingData As Object = Nothing
    ''' <summary>
    '''  Gets or Sets HashTable of values of custom keyboard. where key is e.keycode and value is string corressponding to this keycode ,Values are set or get in ApplicationControlData.
    ''' </summary>   
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property CustomKeyBoardValues() As Hashtable
        Get
            Return _CustomKeyBoardValues
        End Get
        Set(ByVal value As Hashtable)
            _CustomKeyBoardValues = value
        End Set
    End Property
    ''' <summary>
    ''' A hashtable having values are different objects and keys are object names.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ObjectsHashTable() As Hashtable
        Get
            Return _ObjectsHashTable
        End Get
        Set(ByVal value As Hashtable)
            _ObjectsHashTable = value
        End Set
    End Property
    ''' <summary>
    ''' Containing processing row or  data  of exception thrown
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property LastProcessingData() As Object
        Get
            Return _LastProcessingData
        End Get
        Set(ByVal value As Object)
            _LastProcessingData = value
        End Set
    End Property
    ''' <summary>
    ''' A hashtable having values are different System.Types and keys are type names.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property TypesHashTable() As Hashtable
        Get
            Return _TypeHashTable
        End Get
        Set(ByVal value As Hashtable)
            _TypeHashTable = value
        End Set
    End Property
    ''' <summary>
    ''' A hashtable having values are conversion codes and keys are  fullname of types.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ConversionTypeHashTable() As Hashtable
        Get
            Return _ConversionTypeHashTable
        End Get
        Set(ByVal value As Hashtable)
            _ConversionTypeHashTable = value
        End Set
    End Property

    Private Shared _NavigationHashTable As New Hashtable
    ''' <summary>
    ''' A hashtable having values are conversion codes and keys are  fullname of types.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property NavigationHashTable() As Hashtable
        Get
            Return _NavigationHashTable
        End Get
        Set(ByVal value As Hashtable)
            _NavigationHashTable = value
        End Set
    End Property


    ''' <summary>
    '''  Hosting webserver Name default is 182.50.133.109
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebHostingServer() As String
        Get
            Return _WebHostingServer
        End Get
        Set(ByVal value As String)
            _WebHostingServer = value
        End Set
    End Property


    Public Shared Property DemoDate() As Date
        Get
            Return _DemoDate
        End Get
        Set(ByVal value As Date)
            _DemoDate = value
        End Set
    End Property
    Public Shared Property AllowDate() As Date
        Get
            Return _AllowDate
        End Get
        Set(ByVal value As Date)
            _AllowDate = value
        End Set
    End Property

    ''' <summary>
    '''  Hosting webserver user password.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebHostingUserPassword() As String
        Get
            Return _WebHostingUserPassword
        End Get
        Set(ByVal value As String)
            _WebHostingUserPassword = value
        End Set
    End Property

    ''' <summary>
    '''  Returns Server Physical path
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property DataFolderServerPhysicalPath() As String
        Get
            Return _DataFolderServerPhysicalPath
        End Get
        Set(ByVal value As String)
            _DataFolderServerPhysicalPath = value
        End Set
    End Property

    ''' <summary>
    '''  SaralType acceptable values are WebLocal,WebAzure,WebGodaddy,WebCloud,ErpLAN,ErpSNG
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property SaralType() As String
        Get
            Return _SaralType
        End Get
        Set(ByVal value As String)
            _SaralType = value
        End Set
    End Property



    ''' <summary>
    '''  Hosting webserver user name default is "SaralWeb".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebHostingUserName() As String
        Get
            Return _WebHostingUserName
        End Get
        Set(ByVal value As String)
            _WebHostingUserName = value
        End Set
    End Property
    ''' <summary>
    ''' Option is true if website is running on local computer.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property LocalMTPServerEnableSsl() As Boolean
        Get
            Return _LocalMTPServerEnableSsl
        End Get
        Set(ByVal value As Boolean)
            _LocalMTPServerEnableSsl = value
        End Set
    End Property
    Private Shared _procnamefile As DataTable
    ''' <summary>
    ''' This datatable contains controls of all the forms 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ProcNAmeFile() As DataTable
        Get
            Return _ProcNAmeFile
        End Get
        Set(ByVal value As DataTable)
            _ProcNAmeFile = value
        End Set
    End Property

    ''' <summary>
    ''' This datatable contains controls of all the forms 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ENTformstru() As DataTable
        Get
            Return _entFormStru
        End Get
        Set(ByVal value As DataTable)
            _entformstru = value
        End Set
    End Property


    Private Shared _entcontrolProperties As DataTable
    ''' <summary>
    ''' This datatable contains rowwise properties of all the controls of the forms 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property EntControlProperties() As DataTable
        Get
            Return _entcontrolProperties
        End Get
        Set(ByVal value As DataTable)
            _entcontrolProperties = value
        End Set
    End Property

    '  EntControlProperties

    Private Shared _FormsProjectFiles As DataTable
    ''' <summary>
    ''' This datatable contains rows of all forms  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property FormsProjectFiles() As DataTable
        Get
            Return _FormsProjectFiles
        End Get
        Set(ByVal value As DataTable)
            _FormsProjectFiles = value
        End Set
    End Property




    Private Shared _GridCodeMain As DataTable
    ''' <summary>
    ''' This datatable contains rows for all grids  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property GridCodeMain() As DataTable
        Get
            Return _GridCodeMain
        End Get
        Set(ByVal value As DataTable)
            _GridCodeMain = value
        End Set
    End Property

    Private Shared _GridColumns As DataTable
    ''' <summary>
    ''' This datatable contains rows for all grids  
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property GridColumns() As DataTable
        Get
            Return _GridColumns
        End Get
        Set(ByVal value As DataTable)
            _GridColumns = value
        End Set
    End Property

    ''' <summary>
    ''' Option is false if website is running on web computer.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebSMTPServerEnableSsl() As Boolean
        Get
            Return _WebSMTPServerEnableSsl
        End Get
        Set(ByVal value As Boolean)
            _WebSMTPServerEnableSsl = value
        End Set
    End Property
    ''' <summary>
    ''' True if RunningAtWeb.Txt found on web server.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property RunningAtWeb() As Boolean
        Get
            Return _RunningAtWeb
        End Get
        Set(ByVal value As Boolean)
            _RunningAtWeb = value
        End Set
    End Property


    ''' <summary>
    '''  WebSMTPServerHost used as sending email id in WebServer website application default "relay-hosting.secureserver.net".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebSMTPServerHost() As String
        Get
            Return _WebSMTPServerHost
        End Get
        Set(ByVal value As String)
            _WebSMTPServerHost = value
        End Set
    End Property
    ''' <summary>
    ''' Email id  Password used as sending email id in WebServer website application.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebEmailPwd() As String
        Get
            Return _WebEmailPwd
        End Get
        Set(ByVal value As String)
            _WebEmailPwd = value
        End Set
    End Property
    ''' <summary>
    ''' Email id  used as sending email id in WebServer website application default is "info@saralerp.com".
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebEmailId() As String
        Get
            Return _WebEmailId
        End Get
        Set(ByVal value As String)
            _WebEmailId = value
        End Set
    End Property
    ''' <summary>
    ''' Gmail emailid SMTPSERVERport used in sending email id in Local website application default is 587.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property LocalSMTPServerPort() As Integer
        Get
            Return _LocalSMTPServerPort
        End Get
        Set(ByVal value As Integer)
            _LocalSMTPServerPort = value
        End Set
    End Property
    ''' <summary>
    ''' Gmail emailid SMTPSERVERport used in sending email id in Local website application default is 587.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property WebSMTPServerPort() As Integer
        Get
            Return _WebSMTPServerPort
        End Get
        Set(ByVal value As Integer)
            _WebSMTPServerPort = value
        End Set
    End Property


    ''' <summary>
    ''' Gmail emailid SMTP SERVER HOST NAME used as sending email id in Local website application default is smtp.gmail.com. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property LocalSMTPServerHost() As String
        Get
            Return _LocalSMTPServerHost
        End Get
        Set(ByVal value As String)
            _LocalSMTPServerHost = value
        End Set
    End Property
    ''' <summary>
    ''' Gmail emailid password used as sending email id in local website application.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property EmailPassword() As String
        Get
            Return _EmailPassword
        End Get
        Set(ByVal value As String)
            _EmailPassword = value
        End Set
    End Property
    ''' <summary>
    ''' Gmail emailid  used as sending email id in local website application. default is saraluser2012@gmail.com
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property EmailId() As String
        Get
            Return _EmailId
        End Get
        Set(ByVal value As String)
            _EmailId = value
        End Set
    End Property
    ''' <summary>
    ''' Local Host no. used in local website application.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property LocalHostNo() As String
        Get
            Return _LocalHostNo
        End Get
        Set(ByVal value As String)
            _LocalHostNo = value
        End Set
    End Property
    ''' <summary>
    '''Gif image control used in delay loop 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property DelayImage() As Image
        Get
            Return _DelayImage
        End Get
        Set(ByVal value As Image)
            _DelayImage = value
        End Set
    End Property
    ''' <summary>
    ''' A datatable having the rows of ~ separated fixed values, columns are FixedValues_Key,ValuesSet
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property MasterOptions() As DataTable
        Get
            Return _MasterOptions
        End Get
        Set(ByVal value As DataTable)
            _MasterOptions = value
        End Set
    End Property
    ''' <summary>
    ''' A datatable having the rows of  ContextMenu  details of  all controls
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property MenusList() As DataTable
        Get
            Return _MenusList
        End Get
        Set(ByVal value As DataTable)
            _MenusList = value
        End Set
    End Property
    Public Shared Property AuthenticationChecked() As Integer
        Get
            Return _AuthenticationChecked
        End Get
        Set(ByVal value As Integer)
            _AuthenticationChecked = value
        End Set
    End Property
    ''' <summary>
    ''' Containing error string of exception thrown
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ErrorString() As String
        Get
            Return _ErrorString
        End Get
        Set(ByVal value As String)
            _ErrorString = value
        End Set
    End Property




    Public Shared Property TablesExcelControl() As String
        Get
            Return _TablesExcelControl
        End Get
        Set(ByVal value As String)
            _TablesExcelControl = value
        End Set
    End Property
    Public Shared Property FieldsExcelControl() As String
        Get
            Return _FieldsExcelControl
        End Get
        Set(ByVal value As String)
            _FieldsExcelControl = value
        End Set
    End Property


    Public Shared Property Language() As DataRow
        Get
            Return _Language
        End Get
        Set(ByVal value As DataRow)
            _Language = value
        End Set
    End Property
    ''' <summary>
    ''' A datatable of active color scheme used in application
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ColorFontScheme() As DataTable
        Get
            Return _ColorScheme
        End Get
        Set(ByVal value As DataTable)
            _ColorScheme = value
        End Set
    End Property
    ''' <summary>
    ''' A datatable having the informations of all controls used in ERPapplication
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ERPControlsList() As DataTable
        Get
            Return _ControlsList
        End Get
        Set(ByVal value As DataTable)
            _ControlsList = value
        End Set
    End Property
    Private Shared _ERPControlPropertiesList As DataTable
    ''' <summary>
    ''' A datatable having the informations of all controls used in ERPapplication
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property ERPControlPropertiesList() As DataTable
        Get
            Return _ERPControlPropertiesList
        End Get
        Set(ByVal value As DataTable)
            _ERPControlPropertiesList = value
        End Set
    End Property
    Public Shared Property EventLogger() As Boolean
        Get
            Return _EventLogger
        End Get
        Set(ByVal value As Boolean)
            _EventLogger = value
        End Set
    End Property
    Public Shared Property MDIHeight() As Integer
        Get
            Return _MDIHeight
        End Get
        Set(ByVal value As Integer)
            _MDIHeight = value
        End Set
    End Property

    'Property containing the all databases used in Application
    '0_mdf_0--Database  mainconfig (Containing all tables with rows fixed by the developer. 
    '1_mdf_1--Database name of Business unit of financial year (Consist in a user defined folder \SaralErp\Data, Databasename will be first four letters
    'of business unit name + startingdate of financial year_ending date of financial year)
    '2_mdf_2--Database name adding [Temp] on above database created on local server
    '(not used now)3_mdf_3--Database  userconfig (Containing all tables with rows fixed by the developer and added rows by user. 
    Public Shared Property MDFFiles() As Hashtable
        Get
            Return _MDFFiles
        End Get
        Set(ByVal value As Hashtable)
            _MDFFiles = value
        End Set
    End Property

    'Property containing the all servers used in Application
    '0_srv_0--Local server name 
    '1_srv_1--Main server name
    Public Shared Property AllServers() As Hashtable
        Get
            Return _AllServers
        End Get
        Set(ByVal value As Hashtable)
            _AllServers = value
        End Set
    End Property
    Public Shared Property MDIWidth() As Integer
        Get
            Return _MDIWidth
        End Get
        Set(ByVal value As Integer)
            _MDIWidth = value
        End Set
    End Property
    Public Shared Property SqlVersion() As Integer
        Get
            Return _SqlVersion
        End Get
        Set(ByVal value As Integer)
            _SqlVersion = value
        End Set
    End Property

    Public Shared Property xLocalResolution() As Integer
        Get
            Return _xLocalResolution
        End Get
        Set(ByVal value As Integer)
            _xLocalResolution = value
        End Set
    End Property
    Public Shared Property yLocalResolution() As Integer
        Get
            Return _yLocalResolution
        End Get
        Set(ByVal value As Integer)
            _yLocalResolution = value
        End Set
    End Property
    Public Shared Property xFactor() As Decimal
        Get
            Return _xFactor
        End Get
        Set(ByVal value As Decimal)
            _xFactor = value
        End Set
    End Property
    Public Shared Property yFactor() As Decimal
        Get
            Return _yFactor
        End Get
        Set(ByVal value As Decimal)
            _yFactor = value
        End Set
    End Property
    Public Shared Property UserRow() As DataRow
        Get
            Return _UserRow
        End Get
        Set(ByVal value As DataRow)
            _UserRow = value
        End Set
    End Property
    Public Shared Property BusinessFirmRow() As DataRow
        Get
            Return _BusinessFirmRow
        End Get
        Set(ByVal value As DataRow)
            _BusinessFirmRow = value
        End Set
    End Property
    Public Shared Property ActiveForms() As Hashtable
        Get
            Return _ActiveForms
        End Get
        Set(ByVal value As Hashtable)
            _ActiveForms = value
        End Set
    End Property
    Public Enum FormLocation
        TL
        TR
        TC
        MC
        BL
        BR
        BC
        MANUAL
    End Enum
    Public Shared Property FormsHeader() As DataTable
        Get
            Return _FormsHeader
        End Get
        Set(ByVal value As DataTable)
            _FormsHeader = value
        End Set
    End Property

    Public Shared Property RegistryFolder() As String
        Get
            Return _RegistryFolder
        End Get
        Set(ByVal value As String)
            _RegistryFolder = value
        End Set
    End Property
    Public Shared Property AppControlHashTable() As Hashtable
        Get
            Return _AppControlHashTable
        End Get
        Set(ByVal value As Hashtable)
            _AppControlHashTable = value
        End Set
    End Property
    Public Shared Function InitializeConversionHashTable() As Hashtable
        Dim LHashTable As New Hashtable
        LHashTable.Add(LCase("System.Windows.Forms.AnchorStyles"), "ANCH")
        LHashTable.Add(LCase("System.Windows.Forms.Appearance"), "APP0")
        LHashTable.Add(LCase("System.Windows.Forms.DataGridViewAutoSizeColumnsMode"), "ASC0")

        LHashTable.Add(LCase("System.Boolean"), "BLN0")
        LHashTable.Add(LCase("System.Windows.Forms.BorderStyle"), "BST0")
        LHashTable.Add(LCase("System.Drawing.ContentAlignment"), "CAL0")
        LHashTable.Add(LCase("System.Windows.Forms.DataGridViewCellBorderStyle"), "CBS0")
        LHashTable.Add(LCase("CButtonLib.cBlendItems"), "CFB0")
        LHashTable.Add(LCase("System.Windows.Forms.DataGridViewHeaderBorderStyle"), "CHB0")
        LHashTable.Add(LCase("System.Windows.Forms.CheckState"), "CHS0")
        LHashTable.Add(LCase("System.Drawing.Color"), "CLR0")


        LHashTable.Add(LCase("System.Windows.Forms.ContextMenuStrip"), "CMS0")
        LHashTable.Add(LCase("System.Windows.Forms.Cursor"), "CRS0")
        LHashTable.Add(LCase("System.Windows.Forms.Control"), "CTR0")
        LHashTable.Add(LCase("System.Decimal"), "DEC0")
        LHashTable.Add(LCase("System.Windows.Forms.DockStyle"), "DOC0")
        LHashTable.Add(LCase("System.Data.DataRow"), "DRW0")
        LHashTable.Add(LCase("System.Data.DataTable"), "DTB0")
        LHashTable.Add(LCase("System.DateTime"), "DTP0")
        LHashTable.Add(LCase("System.Windows.Forms.DateTimePickerFormat"), "DTPF")
        LHashTable.Add(LCase("CButtonLib.CButton.eFillType"), "EFT0")
        LHashTable.Add(LCase("CButtonLib.CButton.eShape"), "ESHA")
        LHashTable.Add(LCase("System.Windows.Forms.FormBorderStyle"), "FBS0")
        LHashTable.Add(LCase("System.Windows.Forms.FlatStyle"), "FLS0")
        LHashTable.Add(LCase("System.Drawing.Font"), "FNT0")
        LHashTable.Add(LCase("System.Windows.Forms"), "FRM0")
        LHashTable.Add(LCase("System.Windows.Forms.FormStartPosition"), "FSP0")
        LHashTable.Add(LCase("System.Drawing.Drawing2D.LinearGradientMode"), "FTL0")
        LHashTable.Add(LCase("System.Windows.Forms.DataGridViewCellStyle"), "GCS0")
        LHashTable.Add(LCase("System.Windows.Forms.DataGridViewAutoSizeColumnMode"), "GSM0")
        LHashTable.Add(LCase("System.Windows.Forms.HorizontalAlignment"), "HAL0")
        LHashTable.Add(LCase("System.Collections.Hashtable"), "HAS0")
        LHashTable.Add(LCase("System.Windows.Forms.DataGridViewContentAlignment"), "CAL1")
        LHashTable.Add(LCase("System.Drawing.Icon"), "ICO0")
        LHashTable.Add(LCase("System.Windows.Forms.ImeMode"), "IME0")
        LHashTable.Add(LCase("System.Drawing.Image"), "IMG0")
        LHashTable.Add(LCase("System.Windows.Forms.ImageLayout"), "IML0")
        LHashTable.Add(LCase("System.Integer"), "INT0")
        LHashTable.Add(LCase("System.Int16"), "INT0")
        LHashTable.Add(LCase("System.Int32"), "INT0")
        LHashTable.Add(LCase("System.Int64"), "INT0")
        LHashTable.Add(LCase("System.Drawing.Point"), "LOC0")
        LHashTable.Add(LCase("System.Windows.Forms.ToolStripLayoutStyle"), "LOS0")
        LHashTable.Add(LCase("System.Windows.Forms.LeftRightAlignment"), "LRA0")
        LHashTable.Add(LCase("System.Windows.Forms.ComboBox"), "MFV0")
        LHashTable.Add(LCase("System.Windows.Forms.TreeNode"), "NOD0")

        LHashTable.Add(LCase("System.Object"), "OBJ0")
        ' LHashTable.Add(LCase("OTB0()
        LHashTable.Add(LCase("System.Windows.Forms.Padding"), "PAD0")
        LHashTable.Add(LCase("System.Drawing.Printing.PrintAction"), "PRA0")
        LHashTable.Add(LCase("System.Windows.Forms.ProgressBarStyle"), "PRO0")
        LHashTable.Add(LCase("System.Windows.Forms.RightToLeft"), "RTL0")
        LHashTable.Add(LCase("System.Drawing.Size"), "SIZ0")
        LHashTable.Add(LCase("System.Single"), "SNG0")
        LHashTable.Add(LCase("System.Environment.SpecialFolder"), "SPF0")
        LHashTable.Add(LCase("System.String"), "STR0")
        LHashTable.Add(LCase("System.TimeSpan"), "TIM0")
        LHashTable.Add(LCase("System.Windows.Forms.TextImageRelation"), "TIR0")
        LHashTable.Add(LCase("System.Windows.Forms.ToolTip"), "TLP0")
        LHashTable.Add(LCase("System.Windows.Forms.TreeView"), "TNC0")
        LHashTable.Add(LCase("System.Windows.Forms.ToolStripGripStyle"), "TSGS")
        LHashTable.Add(LCase("System.Windows.Forms.ToolTipIcon"), "TTI0")
        LHashTable.Add(LCase("System.Windows.Forms.ToolStripTextDirection"), "TXD0")
        LHashTable.Add(LCase("System.Windows.Forms.FormWindowState"), "WST0")
        Return LHashTable
    End Function
    Public Shared Function InitializeObjectsHashTable() As Hashtable
        Dim LHashTable As New Hashtable
        Dim mButton As New System.Windows.Forms.Button
        LHashTable.Add(LCase("Button"), mButton)
        LHashTable.Add(LCase("Button1"), mButton)
        Dim mCheckBox As New System.Windows.Forms.CheckBox
        LHashTable.Add(LCase("CheckBox"), mCheckBox)
        Dim mColorDialog As New System.Windows.Forms.ColorDialog
        LHashTable.Add(LCase("ColorDialog"), mColorDialog)
        Dim mComboBox As New System.Windows.Forms.ComboBox
        LHashTable.Add(LCase("ComboBox"), mComboBox)
        Dim mContextMenuStrip As New System.Windows.Forms.ContextMenuStrip
        LHashTable.Add(LCase("ContextMenuStrip"), mContextMenuStrip)
        Dim mDataGridView As New System.Windows.Forms.DataGridView
        LHashTable.Add(LCase("DataGridView"), mDataGridView)
        Dim mDateTimePicker As New System.Windows.Forms.DateTimePicker
        LHashTable.Add(LCase("DateTimePicker"), mDateTimePicker)
        Dim mFolderBrowserDialog As New System.Windows.Forms.FolderBrowserDialog
        LHashTable.Add(LCase("FolderBrowserDialog"), mFolderBrowserDialog)
        Dim mFontDialog As New System.Windows.Forms.FontDialog
        LHashTable.Add(LCase("FontDialog"), mFontDialog)
        Dim mForm As New System.Windows.Forms.Form
        LHashTable.Add(LCase("Form"), mForm)
        Dim mLabel As New System.Windows.Forms.Label
        LHashTable.Add(LCase("Label"), mLabel)
        Dim mMenuStrip As New System.Windows.Forms.MenuStrip
        LHashTable.Add(LCase("MenuStrip"), mMenuStrip)
        Dim mNumericUpDown As New System.Windows.Forms.NumericUpDown
        LHashTable.Add(LCase("NumericUpDown"), mNumericUpDown)
        Dim mOpenFileDialog As New System.Windows.Forms.OpenFileDialog
        LHashTable.Add(LCase("OpenFileDialog"), mOpenFileDialog)
        Dim mPanel As New System.Windows.Forms.Panel
        LHashTable.Add(LCase("Panel"), mPanel)
        Dim mPictureBox As New System.Windows.Forms.PictureBox
        LHashTable.Add(LCase("PictureBox"), mPictureBox)
        Dim mPrintDialog As New System.Windows.Forms.PrintDialog
        LHashTable.Add(LCase("PrintDialog"), mPrintDialog)
        Dim mProgressBar As New System.Windows.Forms.ProgressBar
        LHashTable.Add(LCase("ProgressBar"), mProgressBar)
        Dim mRadioButton As New System.Windows.Forms.RadioButton
        LHashTable.Add(LCase("RadioButton"), mRadioButton)
        Dim mStatusStrip As New System.Windows.Forms.StatusStrip
        LHashTable.Add(LCase("StatusStrip"), mStatusStrip)
        Dim mTextBox As New System.Windows.Forms.TextBox
        LHashTable.Add(LCase("TextBox"), mTextBox)
        Dim mTimer As New System.Windows.Forms.Timer
        LHashTable.Add(LCase("Timer"), mTimer)
        Dim mToolStripContainer As New System.Windows.Forms.ToolStripContainer
        LHashTable.Add(LCase("ToolStripContainer"), mToolStripContainer)
        Dim mToolStripMenuItem As New System.Windows.Forms.ToolStripMenuItem
        LHashTable.Add(LCase("ToolStripMenuItem"), mToolStripMenuItem)
        Dim mToolStripSeparator As New System.Windows.Forms.ToolStripSeparator
        LHashTable.Add(LCase("ToolStripSeparator"), mToolStripSeparator)
        Dim mToolTip As New System.Windows.Forms.ToolTip
        LHashTable.Add(LCase("ToolTip"), mToolTip)
        Dim mTreeView As New System.Windows.Forms.TreeView
        LHashTable.Add(LCase("TreeView"), mTreeView)

        Dim mLinkLabel As New System.Windows.Forms.LinkLabel
        LHashTable.Add(LCase("LinkLabel"), mLinkLabel)
        LHashTable.Add(LCase("MaskedTextBox"), New System.Windows.Forms.MaskedTextBox)
        LHashTable.Add(LCase("MonthCalender"), New System.Windows.Forms.MonthCalendar)
        LHashTable.Add(LCase("NotifyIcon"), New System.Windows.Forms.NotifyIcon)
        LHashTable.Add(LCase("RichTextBox"), New System.Windows.Forms.RichTextBox)
        LHashTable.Add(LCase("WebBrowser"), New System.Windows.Forms.WebBrowser)
        LHashTable.Add(LCase("FlowLayoutPanel"), New System.Windows.Forms.FlowLayoutPanel)
        LHashTable.Add(LCase("GroupBox"), New System.Windows.Forms.GroupBox)
        LHashTable.Add(LCase("SplitContainer"), New System.Windows.Forms.SplitContainer)
        LHashTable.Add(LCase("TabControl"), New System.Windows.Forms.TabControl)
        LHashTable.Add(LCase("TabPage"), New System.Windows.Forms.TabPage)
        LHashTable.Add(LCase("TabelLayoutPanel"), New System.Windows.Forms.TableLayoutPanel)
        Return LHashTable

    End Function
    Public Shared Function InitializeTypesHashTable() As Hashtable
        Dim LHashTable As New Hashtable


        LHashTable.Add(LCase("Button"), GetType(System.Windows.Forms.Button))
        LHashTable.Add(LCase("Button1"), GetType(System.Windows.Forms.Button))
        LHashTable.Add(LCase("CheckBox"), GetType(System.Windows.Forms.CheckBox))
        LHashTable.Add(LCase("ColorDialog"), GetType(System.Windows.Forms.ColorDialog))
        LHashTable.Add(LCase("ComboBox"), GetType(System.Windows.Forms.ComboBox))
        LHashTable.Add(LCase("ContextMenuStrip"), GetType(System.Windows.Forms.ContextMenuStrip))
        LHashTable.Add(LCase("DataGridView"), GetType(System.Windows.Forms.DataGridView))
        LHashTable.Add(LCase("DateTimePicker"), GetType(System.Windows.Forms.DateTimePicker))
        LHashTable.Add(LCase("FolderBrowserDialog"), GetType(System.Windows.Forms.FolderBrowserDialog))
        LHashTable.Add(LCase("FontDialog"), GetType(System.Windows.Forms.FontDialog))
        LHashTable.Add(LCase("Form"), GetType(System.Windows.Forms.Form))
        LHashTable.Add(LCase("Label"), GetType(System.Windows.Forms.Label))
        LHashTable.Add(LCase("MenuStrip"), GetType(System.Windows.Forms.MenuStrip))
        LHashTable.Add(LCase("NumericUpDown"), GetType(System.Windows.Forms.NumericUpDown))
        LHashTable.Add(LCase("OpenFileDialog"), GetType(System.Windows.Forms.OpenFileDialog))
        LHashTable.Add(LCase("Panel"), GetType(System.Windows.Forms.Panel))
        LHashTable.Add(LCase("PictureBox"), GetType(System.Windows.Forms.PictureBox))
        LHashTable.Add(LCase("PrintDialog"), GetType(System.Windows.Forms.PrintDialog))
        LHashTable.Add(LCase("ProgressBar"), GetType(System.Windows.Forms.ProgressBar))
        LHashTable.Add(LCase("RadioButton"), GetType(System.Windows.Forms.RadioButton))
        LHashTable.Add(LCase("StatusStrip"), GetType(System.Windows.Forms.StatusStrip))
        LHashTable.Add(LCase("TextBox"), GetType(System.Windows.Forms.TextBox))
        LHashTable.Add(LCase("Timer"), GetType(System.Windows.Forms.Timer))
        LHashTable.Add(LCase("ToolStripContainer"), GetType(System.Windows.Forms.ToolStripContainer))
        LHashTable.Add(LCase("ToolStripMenuItem"), GetType(System.Windows.Forms.ToolStripMenuItem))
        LHashTable.Add(LCase("ToolStripSeparator"), GetType(System.Windows.Forms.ToolStripSeparator))
        LHashTable.Add(LCase("ToolTip"), GetType(System.Windows.Forms.ToolTip))
        LHashTable.Add(LCase("TreeView"), GetType(System.Windows.Forms.TreeView))

        LHashTable.Add(LCase("LinkLabel"), GetType(System.Windows.Forms.LinkLabel))
        LHashTable.Add(LCase("MaskedTextBox"), GetType(System.Windows.Forms.MaskedTextBox))
        LHashTable.Add(LCase("MonthCalender"), GetType(System.Windows.Forms.MonthCalendar))
        LHashTable.Add(LCase("NotifyIcon"), GetType(System.Windows.Forms.NotifyIcon))
        LHashTable.Add(LCase("RichTextBox"), GetType(System.Windows.Forms.RichTextBox))
        LHashTable.Add(LCase("WebBrowser"), GetType(System.Windows.Forms.WebBrowser))
        LHashTable.Add(LCase("FlowLayoutPanel"), GetType(System.Windows.Forms.FlowLayoutPanel))
        LHashTable.Add(LCase("GroupBox"), GetType(System.Windows.Forms.GroupBox))
        LHashTable.Add(LCase("SplitContainer"), GetType(System.Windows.Forms.SplitContainer))
        LHashTable.Add(LCase("TabControl"), GetType(System.Windows.Forms.TabControl))
        LHashTable.Add(LCase("TabPage"), GetType(System.Windows.Forms.TabPage))
        LHashTable.Add(LCase("TabelLayoutPanel"), GetType(System.Windows.Forms.TableLayoutPanel))



        Return LHashTable

    End Function



    Public Shared Function InitializeAppControl() As Hashtable
        Dim LArrayToHashTable As New Hashtable
        LArrayToHashTable.Add(LCase("Client"), "")
        LArrayToHashTable.Add(LCase("CurrDt"), "")
        LArrayToHashTable.Add(LCase("CPUId"), "")
        LArrayToHashTable.Add(LCase("BaseId)"), "")
        LArrayToHashTable.Add(LCase("BiosId"), "")
        LArrayToHashTable.Add(LCase("LastDt"), "")
        LArrayToHashTable.Add(LCase("NoDays"), 0)
        LArrayToHashTable.Add(LCase("SaralKeyExists"), "")
        LArrayToHashTable.Add(LCase("SaralProduct"), "")
        LArrayToHashTable.Add(LCase("SaralVersion"), "")
        LArrayToHashTable.Add(LCase("NoOfVouch"), 0)
        LArrayToHashTable.Add(LCase("TypePC"), "")
        LArrayToHashTable.Add(LCase("SaralType"), "")   '*Permissible types are Local,LAN,WebLocal,WebServer,Cloud

        LArrayToHashTable.Add(LCase("AddlPCNo"), 0)
        LArrayToHashTable.Add(LCase("HomePCNo"), 0)
        LArrayToHashTable.Add(LCase("RemoteNo"), 0)
        LArrayToHashTable.Add(LCase("NodeNo"), 0)
        LArrayToHashTable.Add(LCase("MainClientCode"), "")
        LArrayToHashTable.Add(LCase("AllowBusinessType"), "")
        LArrayToHashTable.Add(LCase("ServicePhone"), "")
        LArrayToHashTable.Add(LCase("AppFolder"), "")
        LArrayToHashTable.Add(LCase("DataFolder"), "")
        LArrayToHashTable.Add(LCase("ImageFolder"), "")
        LArrayToHashTable.Add(LCase("ResourceFile"), "")
        LArrayToHashTable.Add(LCase("ComputerType"), "")   'Server/Node
        LArrayToHashTable.Add(LCase("LANSqlServer"), "")
        LArrayToHashTable.Add(LCase("WebSqlServer"), "")
        LArrayToHashTable.Add(LCase("CloudSqlServer"), "")
        LArrayToHashTable.Add(LCase("LocalSqlServer"), "")
        LArrayToHashTable.Add(LCase("SqlUserName"), "")
        LArrayToHashTable.Add(LCase("SqlUserPassword"), "")
        LArrayToHashTable.Add(LCase("LANSqlUserName"), "")
        LArrayToHashTable.Add(LCase("LANSqlUserPassword"), "")
        LArrayToHashTable.Add(LCase("WebSqlUserName"), "")
        LArrayToHashTable.Add(LCase("WebSqlUserPassword"), "")
        LArrayToHashTable.Add(LCase("CloudSqlUserName"), "")
        LArrayToHashTable.Add(LCase("CloudSqlUserPassword"), "")
        LArrayToHashTable.Add(LCase("RemoteServerName"), "")
        LArrayToHashTable.Add(LCase("RemoteServerUserName"), "")
        LArrayToHashTable.Add(LCase("RemoteServerPassword"), "")
        LArrayToHashTable.Add(LCase("ConnectionTimeOut"), 30)

        LArrayToHashTable.Add(LCase("DemoType"), "")

        Return LArrayToHashTable
    End Function
    Public Shared Function InitializeAllServers() As Hashtable
        Dim LArrayToHashTable As New Hashtable
        LArrayToHashTable.Add(LCase("0_srv_0"), "")
        LArrayToHashTable.Add(LCase("1_srv_1"), "")

        Return LArrayToHashTable
    End Function
    Public Shared Function InitializeMdffiles() As Hashtable
        Dim LArrayToHashTable As New Hashtable
        LArrayToHashTable.Add(LCase("0_mdf_0"), "")
        LArrayToHashTable.Add(LCase("1_mdf_1"), "")
        LArrayToHashTable.Add(LCase("2_mdf_2"), "")
        LArrayToHashTable.Add(LCase("3_mdf_3"), "")
        LArrayToHashTable.Add(LCase("4_mdf_4"), "")

        Return LArrayToHashTable
    End Function


    Public Shared Function GetDelayImage()
        Dim DelayImageFile As String = InitializeAppControl.Item("AppFolder") & "\ResxFolder\bluespinner.gif"
        Return System.Drawing.Image.FromFile(DelayImageFile)
    End Function



    Public Shared Function CalcXFactor() As Decimal
        Return xLocalResolution / xBaseResolution
    End Function
    Public Shared Function CalcYFactor() As Decimal
        Return xLocalResolution / yBaseResolution
    End Function
End Class


