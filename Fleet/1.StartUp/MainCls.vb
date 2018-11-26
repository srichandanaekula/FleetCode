Option Strict Off
Option Explicit On

Imports System.Timers

Public Class MainCls

#Region "Declaration"

    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public objUtilities As Utilities
    Public objDatabaseCreation As DatabaseCreation
    Public IsSAPHANA As Boolean = True

    Public WithEvents Timer As New Timer()
    Public GlobalFormUID As String = ""
    Public FormCloseBoolean As Boolean = False
    Public UpdateBaseDocument As Boolean = False

    'GeneralService
    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public oChildren As SAPbobsCOM.GeneralDataCollection
    Public oChild As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public oGeneralParams As SAPbobsCOM.GeneralDataParams

    Public Shared ohtLookUpForm As Hashtable = New Hashtable

    'Addon Files
    Public objAccidentHistory As clsAccidentHistory
    Public objTripSheet As clsTripSheet
    Public objCallibration As clsCallibration
    Public objPreventiveReminder As clsPreventiveReminder
    Public objFuelUpload As clsFuelUpload
    Public objDocumentType As clsDocumentType
    Public objInsurancePayments As clsInsurancePayments
    Public objImportData As clsImportData
    Public objDeliveryConfirmation As clsDeliveryConfirmation

    'Setup
    Public objConfigurationScreen As clsConfigurationScreen
    Public objDropDownConfig As clsDropDwnCofigScrn
    Public objSelectDistributionRules As clsSelectDistributionRules

    'Masters
    Public objCallibrationMaster As clsCallibrationMaster
    Public objDriverMaster As clsDriverMaster
    Public objRouteMaster As clsRouteMaster
    Public objTaskMaster As clsTaskMaster
    Public objVehicleMaster As clsVehicleMaster
    Public objVehicleStatus As clsVehicleStatus
    'Tyre Management
    Public objTyreMaster As clsTyreMaster
    Public objTyrePositionMaster As clsTyrePositionMaster
    Public objTyreMapping As clsTyreMapping
    Public objTyreMaintenance As clsTyreMaintenance

    'Tank Management
    Public objTankMaster As clsTankMaster
    Public objTankMapping As clsTankMapping
    Public objTankMaintenance As clsTankMaintenance

    'Standard Masters
    Public objGeneralSettings As clsGeneralSettings
    Public objItemGroups As clsItemGroup

    'Standard Transactions

    'Inventory
    Public objGoodsIssue As clsGoodIssue
    Public objInventoryTransfer As clsInventoryTransfer

    'Purchase
    Public objPurchaseOrder As clsPurchaseOrder
    Public objGrpo As clsGRPO
    Public objApInvoice As clsApInvoice
    Public objOutgoingPayment As clsOutgoingPayment

    'Sales
    Public objSalesOrder As clsSalesOrder
    Public objDelivery As clsDelivery
    Public objArInvoice As clsArInvoice
    Public objIncomingPayment As clsIncomingPayment
    Public objSalesQuotation As clsSalesQuotation
    Public objSerialBatchFiles As clsSerialBatchFiles
    Public objDeliveryBatchFile As clsDeliveryBatchFile
    Public objUser As clsUser

#End Region

    Public Sub New()
        objUtilities = New Utilities
        objDatabaseCreation = New DatabaseCreation
    End Sub

#Region "Initialilse"
    Public Function Initialise() As Boolean

        objApplication = objUtilities.GetApplication()
        If objApplication Is Nothing Then Return False
        objCompany = objUtilities.GetCompany(objApplication)
        If objCompany Is Nothing Then : Return False : Exit Function : End If
        If Not objDatabaseCreation.CreateTables() Then Return False
        CreateObjects()

        If objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
            IsSAPHANA = True
        ElseIf objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005 Or _
            objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008 Or _
            objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012 Or _
            objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014 Or _
            objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016 Or _
            objMain.objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL Then
            IsSAPHANA = False
        End If

        Dim CheckLicense As String = ""

        If objMain.IsSAPHANA = True Then
            CheckLicense = "Select CURRENT_TIMESTAMP From ""DUMMY"" Where CURRENT_TIMESTAMP Between '2018-10-01' And '2018-12-01'"
        Else
            CheckLicense = "Select GetDate()  Where GetDate() Between '2018-10-01' And '2018-12-01'"
        End If


        Dim oRsCheckLicense As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsCheckLicense.DoQuery(CheckLicense)
        If oRsCheckLicense.RecordCount = 0 Then
            objCompany.Disconnect()
            System.Windows.Forms.Application.Exit()
            objApplication.StatusBar.SetText("VConnection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        Me.LoadFromXML("Menu.xml")
        Try
            If objMain.objApplication.Menus.Exists("VSP_FLT") = True Then
                objMain.objApplication.Menus.Item("VSP_FLT").Image = System.Windows.Forms.Application.StartupPath & "/Lorry.png"
            End If
        Catch ex As Exception
        End Try
        objApplication.StatusBar.SetText("Vestrics FLeet Add-on is connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

#Region "Create Object"
    Private Sub CreateObjects()
        'Addon Files
        objAccidentHistory = New clsAccidentHistory
        objTripSheet = New clsTripSheet
        objCallibration = New clsCallibration
        objPreventiveReminder = New clsPreventiveReminder
        objFuelUpload = New clsFuelUpload
        objDocumentType = New clsDocumentType
        objInsurancePayments = New clsInsurancePayments
        objImportData = New clsImportData
        objDeliveryConfirmation = New clsDeliveryConfirmation()

        'Setup
        objConfigurationScreen = New clsConfigurationScreen
        objDropDownConfig = New clsDropDwnCofigScrn
        objSelectDistributionRules = New clsSelectDistributionRules

        'Masters
        objCallibrationMaster = New clsCallibrationMaster
        objDriverMaster = New clsDriverMaster
        objRouteMaster = New clsRouteMaster
        objTaskMaster = New clsTaskMaster
        objVehicleMaster = New clsVehicleMaster
        objVehicleStatus = New clsVehicleStatus

        'Tyre Management
        objTyreMaster = New clsTyreMaster
        objTyrePositionMaster = New clsTyrePositionMaster
        objTyreMapping = New clsTyreMapping
        objTyreMaintenance = New clsTyreMaintenance

        'Tank Management
        objTankMaster = New clsTankMaster
        objTankMapping = New clsTankMapping
        objTankMaintenance = New clsTankMaintenance

        'Standard Masters
        objGeneralSettings = New clsGeneralSettings
        objItemGroups = New clsItemGroup

        'Standard Transactions

        'Inventory
        objGoodsIssue = New clsGoodIssue
        objInventoryTransfer = New clsInventoryTransfer

        'Purchase
        objPurchaseOrder = New clsPurchaseOrder
        objGrpo = New clsGRPO
        objApInvoice = New clsApInvoice
        objOutgoingPayment = New clsOutgoingPayment

        'Sales
        objSalesOrder = New clsSalesOrder
        objDelivery = New clsDelivery
        objArInvoice = New clsArInvoice
        objIncomingPayment = New clsIncomingPayment
        objSalesQuotation = New clsSalesQuotation
        objSerialBatchFiles = New clsSerialBatchFiles
        objDeliveryBatchFile = New clsDeliveryBatchFile
        objUser = New clsUser
    End Sub
#End Region

#Region "    ~Create UDOs for the UDTs defined in DB Creation~     "
    Public Sub CreateVehicleMasterUDO()
        If Not Me.UDOExists("VSP_FLT_OVMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPVNO", "U_VSPVNO"}, {"U_VSPVNM", "U_VSPVNM"}, {"U_VSPPNO", "U_VSPPNO"}, {"U_VSPMODEL", "U_VSPMODEL"}, {"U_VSPCNTNM", "U_VSPCNTNM"}}
            Me.registerUDO("VSP_FLT_OVMSTR", "Vehicle Master UDO", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_VMSTR", _
                           "VSP_FLT_VMSTR_C0", "VSP_FLT_VMSTR_C1", "VSP_FLT_VMSTR_C2", "VSP_FLT_VMSTR_C5", "VSP_FLT_VMSTR_C6", "VSP_FLT_VMSTR_C7")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateDriverMasterUDO()
        If Not Me.UDOExists("VSP_FLT_ODRVRMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPFNAME", "U_VSPFNAME"}, {"U_VSPLNAME", "U_VSPLNAME"}, {"U_VSPCNAM", "U_VSPCNAM"}, {"U_VSPMOBNO", "U_VSPMOBNO"}, {"U_VSPEML", "U_VSPEML"}}
            Me.registerUDO("VSP_FLT_ODRVRMSTR", "Driver Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_DRVRMSTR", _
                           "VSP_FLT_DRMSTR_C0", "VSP_FLT_DRMSTR_C2", "VSP_FLT_DRMSTR_C3", "VSP_FLT_DRMSTR_C4")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreatDropDownConfigUDO()
        If Not Me.UDOExists("VSP_FLT_ODDCS") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_ODDCS", "Drop Down Config Scrn", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_DDCS", "VSP_FLT_DDCS_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTripSheetUDO()
        If Not Me.UDOExists("VSP_FLT_OTRSHT") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OTRSHT", "Trip Sheet", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_TRSHT", "VSP_FLT_TRSHT_C1", _
                           "VSP_FLT_TRSHT_C2", "VSP_FLT_TRSHT_C3", "VSP_FLT_TRSHT_C4", "VSP_FLT_TRSHT_C5", "VSP_FLT_TRSHT_C6", "VSP_FLT_TRSHT_C7", "VSP_FLT_TRSHT_C8", "VSP_FLT_TRSHT_C9")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateRouteMasterUDO()
        If Not Me.UDOExists("VSP_FLT_ORTMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPRNM", "U_VSPRNM"}, {"U_VSPSRCE", "U_VSPSRCE"}, {"U_VSPDEST", "U_VSPDEST"}, {"U_VSPTTLKM", "U_VSPTTLKM"}}
            Me.registerUDO("VSP_FLT_ORTMSTR", "Route Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_RTMSTR", "VSP_FLT_RTMSTR_C0", "VSP_FLT_RTMSTR_C1")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTaskMasterUDO()
        If Not Me.UDOExists("VSP_FLT_OTSKMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPTNUM", "U_VSPTNUM"}, {"U_VSPTDSC", "U_VSPTDSC"}, {"U_VSPTTYP", "U_VSPTTYP"}}
            Me.registerUDO("VSP_FLT_OTSKMSTR", "Task Master ", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_TSKMSTR")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTyreMasterUDO()
        If Not Me.UDOExists("VSP_FLT_OTYRMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPTRNUM", "U_VSPTRNUM"}, {"U_VSPWHL", "U_VSPWHL"}, {"U_VSPTBTP", "U_VSPTBTP"}, {"U_VSPINTTP", "U_VSPINTTP"}, {"U_VSPPCHFM", "U_VSPPCHFM"}, {"U_VSPMFR", "U_VSPMFR"}, {"U_VSPWNTY", "U_VSPWNTY"}}
            Me.registerUDO("VSP_FLT_OTYRMSTR", "Tyre Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_TYRMSTR")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTyreMappingUDO()
        If Not Me.UDOExists("VSP_FLT_OTYRMPG") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OTYRMPG", "Tyre Mapping", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_TYRMPG", "VSP_FLT_TYRMPG_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTyreMaintenanceUDO()
        If Not Me.UDOExists("VSP_FLT_OTYRMTNC") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OTYRMTNC", "Tyre Maintenance", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_TYRMTNC", "VSP_FLT_TYRMTNC_C1", "VSP_FLT_TYRMTNC_C2")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateBreakDownEntry()
        If Not Me.UDOExists("VSP_FLT_OBRKDETRY") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OBRKDETRY", "BreakDown Entry ", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_BRKDETRY", "VSP_FLT_BRKDETRY_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub Callibration()
        If Not Me.UDOExists("VSP_FLT_OCALBRTN") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OCALBRTN", "Callibration", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_CALBRTN", "VSP_FLT_CALBRTN_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub PreventiveReminder()
        If Not Me.UDOExists("VSP_FLT_OPRTVRMDR") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OPRTVRMDR", "Preventive Reminder ", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_PRTVRMDR")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CallibrationMstrUDO()
        If Not Me.UDOExists("VSP_FLT_OCALMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPTYP", "U_VSPTYP"}, {"U_VSPNAM", "U_VSPNAM"}}
            Me.registerUDO("VSP_FLT_OCALMSTR", "Callibration Master ", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_CALMSTR")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub TyrePositionMstrUDO()
        If Not Me.UDOExists("VSP_FLT_OTPMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPPSTN", "U_VSPPSTN"}, {"U_VSPNAM", "U_VSPNAM"}}
            Me.registerUDO("VSP_FLT_OTPMSTR", "Tyre Postion Master ", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_TPMSTR")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub ConfigurationScreenUDO()
        If Not Me.UDOExists("VSP_FLT_OCNFGSRN") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OCNFGSRN", "Configuration Screen ", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_CNFGSRN")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub CreateFuelUploadUDO()
        If Not Me.UDOExists("VSP_FLT_OFLUPLD") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OFLUPLD", "Fuel Upload ", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_FLUPLD", "VSP_FLT_FLUPLD_C0")
            findAliasNDescription = Nothing
        End If

    End Sub

    Public Sub CreateTankMasterUDO()
        If Not Me.UDOExists("VSP_FLT_OTANKMSTR") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPTNKNO", "U_VSPTNKNO"}, {"U_VSPTNKNM", "U_VSPTNKNM"}, {"U_VSPTNKMD", "U_VSPTNKMD"}, {"U_VSPTTYPE", "U_VSPTTYPE"}}
            Me.registerUDO("VSP_FLT_OTANKMSTR", "Tank Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSP_FLT_TANKMSTR", "VSP_FLT_TANKMSTR_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTankMappingUDO()
        If Not Me.UDOExists("VSP_FLT_OTANKMPG") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPVCHN0", "U_VSPVCHN0"}, {"U_VSPVCHNM", "U_VSPVCHNM"}, {"U_VSPVMDL", "U_VSPVMDL"}}
            Me.registerUDO("VSP_FLT_OTANKMPG", "Tank Mapping", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_TANKMPG", "VSP_FLT_TANKMPG_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateTankMaintenanceUDO()
        If Not Me.UDOExists("VSP_FLT_OTNKMTNC") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPVMD", "U_VSPVMD"}, {"U_VSPVNM", "U_VSPVNM"}}
            Me.registerUDO("VSP_FLT_OTNKMTNC", "Tank Maintenance", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_TNKMTNC", "VSP_FLT_TNKMTNC_C1", "VSP_FLT_TNKMTNC_C2")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateAccidentHistoryUDO()
        If Not Me.UDOExists("VSP_FLT_OACCHIST") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPVHCD", "U_VSPVHCD"}, {"U_VSPVHNM", "U_VSPVHNM"}, {"U_VSPPLTNO", "U_VSPPLTNO"}}
            Me.registerUDO("VSP_FLT_OACCHIST", "Accident History UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_ACCHIST", "VSP_FLT_ACCHIST_C0", "VSP_FLT_ACCHIST_C1", "VSP_FLT_ACCHIST_C2")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateInsurancePaymentsUDO()
        If Not Me.UDOExists("VSP_FLT_OINSPAY") Then
            Dim findAliasNDescription = New String(,) {{"U_VSPVNO", "U_VSPVNO"}, {"U_VSPVNM", "U_VSPVNM"}}
            Me.registerUDO("VSP_FLT_OINSPAY", "Insurance Payments UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_FLT_INSPAY", "VSP_FLT_INSPAY_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub CreateVehicleStatusUDO()
        If Not Me.UDOExists("VSP_FLT_OVEHST") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_FLT_OVEHST", "Vehicle Status UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_VECHSTS", "VSP_VECHSTS_C0")
            findAliasNDescription = Nothing
        End If
    End Sub


    Public Sub CreateDeliveryConfirmationUDO()
        If Not Me.UDOExists("VSPODELCONF") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSPODELCONF", "Delivery Confirmation UDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSPDELCONF")
            findAliasNDescription = Nothing
        End If
    End Sub

#End Region

#Region "UDO Exists"
    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function
#End Region

#Region "Register UDO"

    Function registerUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal childTable6 As String = "", Optional ByVal childTable7 As String = "", Optional ByVal childTable8 As String = "", Optional ByVal childTable9 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "L" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If
            If childTable6 <> "" Then
                v_udoMD.ChildTables.TableName = childTable6
                v_udoMD.ChildTables.Add()
            End If
            If childTable7 <> "" Then
                v_udoMD.ChildTables.TableName = childTable7
                v_udoMD.ChildTables.Add()
            End If
            If childTable8 <> "" Then
                v_udoMD.ChildTables.TableName = childTable8
                v_udoMD.ChildTables.Add()
            End If
            If childTable9 <> "" Then
                v_udoMD.ChildTables.TableName = childTable9
                v_udoMD.ChildTables.Add()
            End If
            If v_udoMD.Add() = 0 Then
                registerUDO = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function

    Function registerUDONoLog(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDONoLog = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "A" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDONoLog = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDONoLog = False
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function
#End Region

#Region "Add Menu's With XML"

    Private Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objApplication.GetLastBatchResults()

    End Sub

#End Region

#Region "Item Event"
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try

            '------------------------------------------------------------------------
            Try
                If Fleet.MainCls.ohtLookUpForm.ContainsValue(FormUID) = True Then
                    Dim keys As ICollection = Fleet.MainCls.ohtLookUpForm.Keys
                    Dim keysArray(Fleet.MainCls.ohtLookUpForm.Count - 1) As String
                    keys.CopyTo(keysArray, 0)
                    For Each key As String In keysArray
                        If FormUID = Fleet.MainCls.ohtLookUpForm(key) Then
                            While Fleet.MainCls.ohtLookUpForm.ContainsValue(key) = True
                                For Each dKey As String In keysArray
                                    If key = Fleet.MainCls.ohtLookUpForm(dKey) Then
                                        key = dKey
                                        Exit For
                                    End If
                                Next
                            End While
                            objMain.objApplication.Forms.Item(key).Select()
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try

            Select Case pVal.FormTypeEx
                'Addon Files
                Case "VSP_FLT_ACCHIST_Form"
                    objAccidentHistory.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_CALBRTN_Form"
                    objCallibration.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TRSHT_Form"
                    objTripSheet.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_PRTVRMDR_Form"
                    objPreventiveReminder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_DOCTYPE_Form"
                    objDocumentType.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_FLUPLD_Form"
                    objFuelUpload.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_INSPAY_Form"
                    objInsurancePayments.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_IMPDATA_Form"
                    objImportData.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "VSP_DELVCONFR_Form"
                    objDeliveryConfirmation.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Setup
                Case "VSP_FLT_CNFGSRN_Form"
                    objConfigurationScreen.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_DDCS_Form"
                    objDropDownConfig.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_DISTRL_Form"
                    objSelectDistributionRules.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Masters
                Case "VSP_FLT_CALMSTR_Form"
                    objCallibrationMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_DRVRMSTR_Form"
                    objDriverMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_RTMSTR_Form"
                    objRouteMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TSKMSTR_Form"
                    objTaskMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_VMSTR_Form"
                    objVehicleMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_VCHSTS_Form"
                    objVehicleStatus.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Tyre Management
                Case "VSP_FLT_TYRMSTR_Form"
                    objTyreMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TPMSTR_Form"
                    objTyrePositionMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TYRMPG_Form"
                    objTyreMapping.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TYRMTNC_Form"
                    objTyreMaintenance.ItemEvent(FormUID, pVal, BubbleEvent)
                
                    'Tank Management
                Case "VSP_FLT_TANKMSTR_Form"
                    objTankMaster.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TANKMPG_Form"
                    objTankMapping.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "VSP_FLT_TNKMTNC_Form"
                    objTankMaintenance.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Standard Masters
                Case "138"
                    objGeneralSettings.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "63"
                    objItemGroups.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Standard Transactions
                    'Inventory Transfer
                Case "940"
                    objInventoryTransfer.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Purchase               
                Case "142"
                    objPurchaseOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "143"
                    objGrpo.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "141"
                    objApInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "426"
                    objOutgoingPayment.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Sales
                Case "139"
                    objSalesOrder.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "140"
                    objDelivery.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "133"
                    objArInvoice.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "170"
                    objIncomingPayment.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "149"
                    objSalesQuotation.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "41"
                    objSerialBatchFiles.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "42"
                    objDeliveryBatchFile.ItemEvent(FormUID, pVal, BubbleEvent)
                Case "20700"
                    objUser.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Menu Events"
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        Dim objform As SAPbouiCOM.Form
        Try
            objform = objMain.objApplication.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case "VSP_FLT_VMSTR"
                    objVehicleMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_DRVRMSTR"
                    objDriverMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_DDCS"
                    objDropDownConfig.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TRSHT"
                    objTripSheet.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_RTMSTR"
                    objRouteMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TSKMSTR"
                    objTaskMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TYRMTNC"
                    objTyreMaintenance.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TYRMPG"
                    objTyreMapping.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TYRMSTR"
                    objTyreMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_CALBRTN"
                    objCallibration.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_PRTVRMDR"
                    objPreventiveReminder.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_CALMSTR"
                    objCallibrationMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TPMSTR"
                    objTyrePositionMaster.MenuEvent(pVal, BubbleEvent)             
                Case "VSP_FLT_CNFGSRN"
                    objConfigurationScreen.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_FLUPLD"
                    objFuelUpload.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TANKMSTR"
                    objTankMaster.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TNKMTNC"
                    objTankMaintenance.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_TANKMPG"
                    objTankMapping.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_ACCHIST"
                    objAccidentHistory.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_INSPAY"
                    objInsurancePayments.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_IMPDATA"
                    objImportData.MenuEvent(pVal, BubbleEvent)
                Case "VSP_FLT_VECHSTS"
                    objVehicleStatus.MenuEvent(pVal, BubbleEvent)
                    'Case "VSP_DEL_CONF"
                    '    objDeliveryConfirmation.MenuEvent(pVal, BubbleEvent)


                Case "1282"
                    If objform.TypeEx = "VSP_FLT_VMSTR_Form" Then
                        objVehicleMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_DRVRMSTR_Form" Then
                        objDriverMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_DDCS_Form" Then
                        objDropDownConfig.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TRSHT_Form" Then
                        objTripSheet.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_RTMSTR_Form" Then
                        objRouteMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TSKMSTR_Form" Then
                        objTaskMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TYRMTNC_Form" Then
                        objTyreMaintenance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TYRMPG_Form" Then
                        objTyreMapping.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TYRMSTR_Form" Then
                        objTyreMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_CALBRTN_Form" Then
                        objCallibration.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_PRTVRMDR_Form" Then
                        objPreventiveReminder.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_CALMSTR_Form" Then
                        objCallibrationMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TPMSTR_Form" Then
                        objTyrePositionMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_FLUPLD_Form" Then
                        objFuelUpload.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TANKMSTR_Form" Then
                        objTankMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TNKMTNC_Form" Then
                        objTankMaintenance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TANKMPG_Form" Then
                        objTankMapping.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_ACCHIST_Form" Then
                        objAccidentHistory.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_INSPAY_Form" Then
                        objInsurancePayments.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_VCHSTS_Form" Then
                        objVehicleStatus.MenuEvent(pVal, BubbleEvent)
                        'ElseIf objform.TypeEx = "VSP_DELVCONFR_Form" Then
                        '    objDeliveryConfirmation.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "View"
                    If objform.TypeEx = "VSP_FLT_VMSTR_Form" Then
                        objVehicleMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_DRVRMSTR_Form" Then
                        objDriverMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TYRMTNC_Form" Then
                        objTyreMaintenance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TNKMTNC_Form" Then
                        objTankMaintenance.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "Add Row"
                    If objform.TypeEx = "VSP_FLT_RTMSTR_Form" Then
                        objRouteMaster.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "Delete Row"
                    If objform.TypeEx = "VSP_FLT_DDCS_Form" Then
                        objDropDownConfig.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TYRMTNC_Form" Then
                        objTyreMaintenance.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_VMSTR_Form" Then
                        objVehicleMaster.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_FLUPLD_Form" Then
                        objFuelUpload.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TRSHT_Form" Then
                        objTripSheet.MenuEvent(pVal, BubbleEvent)
                    ElseIf objform.TypeEx = "VSP_FLT_TNKMTNC_Form" Then
                        objTankMaintenance.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "Generate"
                    If objform.TypeEx = "VSP_FLT_TRSHT_Form" Then
                        objTripSheet.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "Generate Delivery"
                    If objform.TypeEx = "VSP_FLT_TRSHT_Form" Then
                        objTripSheet.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "Generate Invoice"
                    If objform.TypeEx = "VSP_FLT_TRSHT_Form" Then
                        objTripSheet.MenuEvent(pVal, BubbleEvent)
                    End If
                Case "Create Service Type PO"
                    objTyreMaintenance.MenuEvent(pVal, BubbleEvent)
                    objTankMaintenance.MenuEvent(pVal, BubbleEvent)
                Case "Issue Component"
                    objCallibration.MenuEvent(pVal, BubbleEvent)
                Case "CreateGoodsReceipt"
                    objTyreMapping.MenuEvent(pVal, BubbleEvent)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Form Data Event"
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            'Addon Files
            Case "VSP_FLT_TRSHT_Form"
                objTripSheet.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "VSP_FLT_INSPAY_Form"
                objInsurancePayments.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "VSP_DELVCONFR_Form"
                objDeliveryConfirmation.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                'Masters
            Case "VSP_FLT_VMSTR_Form"
                objVehicleMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "VSP_FLT_DRVRMSTR_Form"
                objDriverMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "VSP_FLT_VCHSTS_Form"
                objVehicleStatus.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                'Tyre Management
            Case "VSP_FLT_TYRMSTR_Form"
                objTyreMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "VSP_FLT_TYRMPG_Form"
                objTyreMapping.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "VSP_FLT_TPMSTR_Form"
                objTyrePositionMaster.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                'Tank Management
            Case "VSP_FLT_TANKMPG_Form"
                objTankMapping.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                'Inventory
            Case "720"
                objGoodsIssue.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "940"
                objInventoryTransfer.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                'Purchase
            Case "142"
                objPurchaseOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "143"
                objGrpo.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "141"
                objApInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "426"
                objOutgoingPayment.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                'Sales
            Case "139"
                objSalesOrder.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "140"
                objDelivery.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "133"
                objArInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case "170"
                objIncomingPayment.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        End Select
    End Sub
#End Region

#Region "Right Click Event"
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim objForm As SAPbouiCOM.Form
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        If objForm.TypeEx = "VSP_FLT_VMSTR_Form" Then
            objVehicleMaster.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_DDCS_Form" Then
            objDropDownConfig.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_DRVRMSTR_Form" Then
            objDriverMaster.RightClickEvent(eventInfo, BubbleEvent)
       ElseIf objForm.TypeEx = "VSP_FLT_RTMSTR_Form" Then
            objRouteMaster.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_TYRMTNC_Form" Then
            objTyreMaintenance.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_FLUPLD_Form" Then
            objFuelUpload.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_CALBRTN_Form" Then
            objCallibration.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_TNKMTNC_Form" Then
            objTankMaintenance.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_TRSHT_Form" Then
            objTripSheet.RightClickEvent(eventInfo, BubbleEvent)
        ElseIf objForm.TypeEx = "VSP_FLT_TYRMPG_Form" Then
            objTyreMapping.RightClickEvent(eventInfo, BubbleEvent)
        End If
    End Sub
#End Region

#Region "Application Event"
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                objCompany.Disconnect()
                End
        End Select
    End Sub
#End Region

End Class