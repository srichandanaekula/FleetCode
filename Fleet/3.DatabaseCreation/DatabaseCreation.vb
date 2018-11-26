Public Class DatabaseCreation

#Region "Declaration"
    Private objUtilities As Utilities
    Dim DBCode As String = "v4.4"
    Dim DBName As String = "v4.4"
    Dim Version As String = "v4.4"
#End Region

#Region "DB Creation Main"
    Public Sub New()
        objUtilities = New Utilities
    End Sub

    Public Function CreateTables() As Boolean

        Try
            objUtilities.CreateTable("VSP_FLT_DBCONF", "VSPL DBCONFIG(FLEET) TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUtilities.AddAlphaField("@VSP_FLT_DBCONF", "VERSION", "VERSION", 100)
            Dim oRs As SAPbobsCOM.Recordset
            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objMain.IsSAPHANA = True Then
                oRs.DoQuery("SELECT * FROM ""@VSP_FLT_DBCONF"" where ""U_VERSION"" = '" & Version & "'")
            Else
                oRs.DoQuery("SELECT * FROM [@VSP_FLT_DBCONF] where U_VERSION = '" & Version & "'")
            End If


            Dim iDBConfigRecordCount As Integer = oRs.RecordCount
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            If iDBConfigRecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Your Database will now be upgraded to " + Version + ". Please Wait... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '------------------------------------------------------------------------------
                'Standard

                objMain.objUtilities.AddAlphaField("OUSR", "VSPRST", "Restriction", 1)
                objMain.objUtilities.AddFloatField("OBTN", "CenVat", "CenVat", SAPbobsCOM.BoFldSubTypes.st_Price)

                'OHEM
                objUtilities.AddAlphaField("OHEM", "VSPCNTR", "Contractor", 30)
                objUtilities.AddAlphaField("OHEM", "VSPDRCOD", "Driver Code", 30)

                'OADM
                objUtilities.AddAlphaMemoField("OADM", "VSPDRATC", "Driver Attachment", 64000)
                objUtilities.AddAlphaMemoField("OADM", "VSPVHATC", "Vehicle Attachment", 64000)
                objUtilities.AddAlphaField("OADM", "VSPCITM", "Commission Item", 30)
                objUtilities.AddAlphaField("OADM", "VSPTITM", "Transport Item", 30)

                'Newly Added on 10-09-2018 for OPOR table
                'OPOR
                objUtilities.AddAlphaField("OPOR", "TCCode", "Tripsheet Customer Code", 30) 'For Revenue Details Tab
                objUtilities.AddAlphaField("OPOR", "TCRef", "Tripsheet Customer Ref", 30)


                'POR1
                objUtilities.AddAlphaField("POR1", "VSPCODE", "TyreMaster Code", 30)

                'ORDR
                objUtilities.AddAlphaField("ORDR", "VSPFLSTS", "Fleet Status", 30)
                objMain.objUtilities.addField("ORDR", "VSPORTYP", "Order Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, "TR,MT,Other", "Transport,Material,Other", "Other", "Yes")
                objUtilities.AddFloatField("ORDR", "VSPVLQTY", "Vendor Loaded Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)  'Abinas 0n 08-11-2018



                objUtilities.AddAlphaField("ORDR", "VSPPONUM", "Po DocNumber", 30) ''Abinas 0n 26-09-2018
                objUtilities.AddAlphaField("ORDR", "VSPPOENT", "Po DocEntry", 30)    ''Abinas 0n 26-09-2018

                'RDR1
                objUtilities.AddFloatField("RDR1", "VSPUNPRC", "Unit Price", SAPbobsCOM.BoFldSubTypes.st_Price)

                'Tax Calculation Feilds
                'objUtilities.AddFloatField("RDR1", "VSPCCVDB", "CCVD BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPCCVDT", "CCVD TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPACVDB", "ACVD BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPACVDT", "ACVD TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPVATB", "VAT BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPVATT", "VAT TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPCSTB", "CST BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPCSTT", "CST TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPBCDB", "BCD BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPBCDT", "BCD TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPC_CVDB", "C_CVD BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPC_CVDT", "C_CVD TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPC_CB", "C_Cess BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPC_CT", "C_Cess TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPA_CVDB", "A_CVD BaseAmt", SAPbobsCOM.BoFldSubTypes.st_Price)
                'objUtilities.AddFloatField("RDR1", "VSPA_CVDT", "A_CVD TaxAmt", SAPbobsCOM.BoFldSubTypes.st_Price)

                'OPDN()
                objUtilities.AddAlphaField("OPDN", "VSPID", "ID", 30)   ''not executing
                objUtilities.AddAlphaField("OPDN", "VSPDNM", "Dealer Name ", 30)
                objUtilities.AddAlphaField("OPDN", "VSPCSTID", "Customer ID", 30)
                objUtilities.AddAlphaField("OPDN", "VSPLCTN", "Location", 30)
                objUtilities.AddAlphaField("OPDN", "VSPVCHNO", "Vehicle No.", 30)
                objUtilities.AddAlphaField("OPDN", "VSPCUR", "Currency", 30)
                objUtilities.AddAlphaField("OPDN", "VSPBAL", "Balance", 30)
                objUtilities.AddAlphaField("OPDN", "VSPEXPT", "Extra Points ", 30)
                objUtilities.AddAlphaField("OPDN", "MODVAT", "MODVAT", 10)
                objUtilities.AddFloatField("PDN1", "RATE", "Rate", SAPbobsCOM.BoFldSubTypes.st_Price)

                'OJDT
                objMain.objUtilities.AddAlphaField("OJDT", "VSPTRPNO", "Trip Sheet No.", 30)
                objMain.objUtilities.AddAlphaField("OJDT", "VSPDCTYP", "Document Type", 30)

                'OIGE
                objUtilities.AddAlphaField("OIGE", "VSPVHNO", "Clibration Vehicle No.", 30)
                objUtilities.AddAlphaField("OIGE", "VSPDCTYP", "DocType", 30)
                objUtilities.AddAlphaField("OIGE", "VSPDCNO", "DocNum", 30)
                objUtilities.AddAlphaField("OIGE", "VSPLNUM", "LineNum", 30)
                objUtilities.AddAlphaField("OIGE", "VSPTRSHP", "Trip Sheet No.", 30)
                objUtilities.AddAlphaField("OIGE", "VSPDCETY", "DocEntry", 30)
                objUtilities.AddAlphaField("OIGE", "VSPTYPS", "TyrePosition", 30)

                'OVPM
                objUtilities.AddAlphaField("OVPM", "VSPDCTYP", "DocType", 30)
                objUtilities.AddAlphaField("OVPM", "VSPDCNO", "DocNum", 30)

                'OITB
                objUtilities.AddAlphaField("OITB", "VSPCSA", "CSA", 1)

                'OITM 
                objUtilities.AddFloatField("OITM", "VSPTLPRC", "Tolerance Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage)    ''Abinas on 17-09-2018

                'OCRD
                objUtilities.AddFloatField("OCRD", "VSPCOMSN", "Commission", SAPbobsCOM.BoFldSubTypes.st_Percentage)   ''Added by Abinas on 13-11-2018

                'ODLN
                objUtilities.AddAlphaField("ODLN", "VSPDRV1", "Driver 1", 30)   ''Abinas on 17-09-2018
                objUtilities.AddAlphaField("ODLN", "VSPDRV2", "Driver 2", 30)   ''Abinas on 17-09-2018
                'INV1
                objUtilities.AddFloatField("INV1", "VSPADQTY", "Actual Delivered Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)    ''by Abinas on 19-09-2018
                objUtilities.AddFloatField("INV1", "VSPDFQTY", "Difference Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)          ''by Abinas on 19-09-2018
                objUtilities.AddFloatField("INV1", "VSPTLQTY", "Tolerance Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)           ''by Abinas on 19-09-2018
                objUtilities.AddFloatField("INV1", "VSPSTQTY", "Shortage Qty", SAPbobsCOM.BoFldSubTypes.st_Quantity)            ''by Abinas on 19-09-2018




                'UDT's
                Me.CreateVehicleMaster()
                Me.CreateDriverMaster()
                Me.CreatDropDownConfigScrn()
                Me.CreateTripSheet()
                Me.CreateRouteMaster()
                Me.CreatTaskMaster()
                Me.CreateTyreMaster()
                Me.CreateTyreMaintenance()
                Me.CreateTyreMapping()
                Me.CreateBreakDownEntry()
                Me.Callibration()
                Me.PreventiveReminder()
                Me.CreateTyrePostionMstr()
                Me.CreateCalibrationMaster()
                Me.CreateConfigurationScreen()
                Me.CreateFuelUpload()
                Me.CreateTankMaster()
                Me.CreateTankMaintenance()
                Me.CreateTankMapping()
                Me.CreateAccidentHistory()
                Me.CreateInsurancePayments()
                Me.CreateVehicleStatus()
                Me.CreateDeliveryConfirmation()

                'UDO's
                objMain.CreateVehicleMasterUDO()
                objMain.CreateDriverMasterUDO()
                objMain.CreatDropDownConfigUDO()
                objMain.CreateTripSheetUDO()
                objMain.CreateRouteMasterUDO()
                objMain.CreateTaskMasterUDO()
                objMain.CreateTyreMaintenanceUDO()
                objMain.CreateTyreMappingUDO()
                objMain.CreateTyreMasterUDO()
                objMain.CreateBreakDownEntry()
                objMain.Callibration()
                objMain.PreventiveReminder()
                objMain.CallibrationMstrUDO()
                objMain.TyrePositionMstrUDO()
                objMain.ConfigurationScreenUDO()
                objMain.CreateFuelUploadUDO()
                objMain.CreateTankMaintenanceUDO()
                objMain.CreateTankMappingUDO()
                objMain.CreateTankMasterUDO()
                objMain.CreateAccidentHistoryUDO()
                objMain.CreateInsurancePaymentsUDO()
                objMain.CreateVehicleStatusUDO()
                objMain.CreateDeliveryConfirmationUDO()
                '--------------------------------------------------------------------------------------
                'Close DB Script
                objUtilities.AddDataToNoObjectTable("VSP_FLT_DBCONF", DBCode, DBName, "U_Version", Version)

                objMain.objApplication.StatusBar.SetText("Your Database has now been upgraded to Version " + Version + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Create Tables"

    Sub CreateVehicleMaster()
        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR", "Vehicle Master Table", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPVNO", "Vehicle/Registration No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPVNM", "Vehicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPPNO", "Plate No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPOWSHP", "Ownership", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCHSNO", "Chasis No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCAT", "Category", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPLOC", "Location", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCNTR", "Contracter", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCNTNM", "Contracter Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPENGNO", "Engine No.", 30)

        'objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPODRDG", "Odometer Reading", 30)
        'objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPMWL", "Mileage(With Load)", 30)
        'objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPMWOL", "Mileage(Without Load)", 30)

        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR", "VSPODRDG", "Odometer Reading", SAPbobsCOM.BoFldSubTypes.st_Measurement)   ''changed AddAlphafield to Float
        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR", "VSPMWL", "Mileage(With Load)", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR", "VSPMWOL", "Mileage(Without Load)", SAPbobsCOM.BoFldSubTypes.st_Measurement)

        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPYEAR", "Year", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPMAKE", "Make", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPMODEL", "Model", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPSRNO", "Serial No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCOLOR", "Color", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPIDN", "Identification", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPGWGT", "Gross Weight", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPLEN", "Length", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPWIDTH", "Width", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPHGT", "Height", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPWBASE", "Wheel Base", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPENGSZ", "Engine Size", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPNOCYL", "No. Of Cylinders", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPTTYPE", "Transmission Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPFTYPE", "Fuel Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPSPPLG", "Spark Plug Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPBTRY", "Battery Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPHDLMP", "HeadLamp Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPMSC1", "Misc 1", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPMSC2", "Misc 2", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPNOAXL", "No. Of Axles", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPSZFT", "Size(Front Tyres)", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPPRFT", "Pressure(Front Tyres)", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPSZRT", "Size(Rear Tyres)", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPPRRT", "Pressure(Rear Tyres)", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPDTFT", "Dual Tyres(Front Tyres)", 1)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPDTRT", "Dual Tyres(Rear Tyres)", 1)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPAVLB", "Availability", 1)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCALB", "Calibrate", 1)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPTAC", "Tank Attached Capacity", 30)
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT1", "Photo 1")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT2", "Photo 2")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT3", "Photo 3")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT4", "Photo 4")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT5", "Photo 5")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT6", "Photo 6")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT7", "Photo 7")
        objMain.objUtilities.AddImageField("@VSP_FLT_VMSTR", "VSPPHT8", "Photo 8")
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPVAL", "Value", 30)

        'objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPTYRES", "No. Of Tyres", 30)
        objMain.objUtilities.AddInteger("@VSP_FLT_VMSTR", "VSPTYRES", "No. Of Tyres", SAPbobsCOM.BoFldSubTypes.st_None, 10) ''changed to integer feild 21-11-18

        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPDIITE", "Diesel Item Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPVEHCC", "Vehicle CC", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCDWGT", "Cabin Weight", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPTWGT", "Tank Weight", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCHWGT", "Chasis Weight", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPOTHFT", "Other Fittings", 30)

        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPCHK", "Check", 1)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR", "VSPUNPCK", "Under Periodic Check", 1)

        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR_C0", "Vehicle Master Child 1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C0", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C0", "VSPNAME", "Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C0", "VSPNUM", "Number", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C0", "VSPISSDT", "Issue Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C0", "VSPEXPDT", "Expiration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_VMSTR_C0", "VSPATTCH", "Attachment", 64000)
        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR_C0", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR_C1", "Vehicle Master Child 2", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C1", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C1", "VSPCMPY", "Company", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C1", "VSPLNNO", "Loan No.", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C1", "VSPSTRDT", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C1", "VSPENDDT", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddInteger("@VSP_FLT_VMSTR_C1", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_None, 11)
    
        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR_C2", "Vehicle Master Child 3", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C2", "VSPPART", "Part", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C2", "VSPDEALR", "Dealer", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C2", "VSPDTPUR", "Date Purchased", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR_C2", "VSPPRC", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C2", "VSPWEDT", "Warranty Expiration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C2", "VSPINSRV", "In Service", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C2", "VSPOTSRV", "Out of Service", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C2", "VSPTNSDT", "Transfer Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C2", "VSPDTSLD", "Date Sold", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C2", "VSPSLDTO", "Sold To", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C2", "VSPCMTS", "Comments", 250)
        objMain.objUtilities.AddInteger("@VSP_FLT_VMSTR_C2", "VSPNFKM", "No. of KM's", SAPbobsCOM.BoFldSubTypes.st_None, 11)
        objMain.objUtilities.AddInteger("@VSP_FLT_VMSTR_C2", "VSPNFDYS", "No. of Days", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR_C5", "Vehicle Master Child 6", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C5", "VSPFRMDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C5", "VSPTODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C5", "VSPCNTR", "Contractor", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C5", "VSPCNTNM", "Contractor Name", 100)

        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR_C6", "Vehicle Master Child 7", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C6", "VSPCMPNY", "Company", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C6", "VSPINSNO", "Insurance No.", 100)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C6", "VSPSTDT", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C6", "VSPENDDT", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR_C6", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@VSP_FLT_VMSTR_C6", "VSPVAMT", "Insurance Valid Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
      
        objMain.objUtilities.CreateTable("VSP_FLT_VMSTR_C7", "Vehicle Master Child 8", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C7", "VSPFRMDT", "From Date ", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_VMSTR_C7", "VSPTODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C7", "VSPRDNG", "Reading", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C7", "VSPCLSRD", "Odometer Closed Reading", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_VMSTR_C7", "VSPSRLNO", "Odometer Serial Number", 30)

    End Sub

    Sub CreateDriverMaster()
        objMain.objUtilities.CreateTable("VSP_FLT_DRVRMSTR", "Driver Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPFNAME", "First Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPEMPNO", "Employee No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPMNAME", "Middle Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPLNAME", "Last Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPSTS", "Status", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPLCTN", "Location", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPCTGRY", "Category", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRVRMSTR", "VSPHRDT", "Hire Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRVRMSTR", "VSPDTLVE", "Date of Leave", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRVRMSTR", "VSPDOB", "DOB", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPCNCD", "Contractor Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPCNAM", "Contractor Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPNUM", "Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPNAME", "Name", 100)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRVRMSTR", "VSPISUDT", "Issue Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRVRMSTR", "VSPEXPDT", "Expiration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPATTCH", "ReferredBy", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPADRS1", "Address 1", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPADRS2", "Address 2", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPCTY", "City", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPSTATE", "State", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPPCD", "Postal Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPHMPNO", "Home Phone No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPMOBNO", "Mobile No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPALTPN", "Alternate Phone No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPEML", "Mobile No", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPHL", "Harzardous License", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRVRMSTR", "VSPEXDT", "Expiry Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRVRMSTR", "VSPAVLBL", "Available", 1)
        objMain.objUtilities.AddImageField("@VSP_FLT_DRVRMSTR", "VSPIMG", "Photo 1")

        objMain.objUtilities.CreateTable("VSP_FLT_DRMSTR_C0", "Driver Master Child 1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C0", "VSPINTYP", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C0", "VSPINCMP", "Company", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C0", "VSPINLN", "Loan No.", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRMSTR_C0", "VSPINSDT", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRMSTR_C0", "VSPINEDT", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C0", "VSPINAMT", "Amount", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_DRMSTR_C2", "Driver Master Child 3", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C2", "VSPOANUM", "Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C2", "VSPOANME", "Name", 100)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRMSTR_C2", "VSPOAIDT", "Issue Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRMSTR_C2", "VSPOAEDT", "Expiration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_DRMSTR_C2", "VSPOAATC", "Attachment", 64000)

        objMain.objUtilities.CreateTable("VSP_FLT_DRMSTR_C3", "Driver Master Child 4", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C3", "VSPCHCLN", "CheckList Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C3", "VSPCHCHB", "CheckBox", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_DRMSTR_C4", "Driver Master Child 5", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRMSTR_C4", "VSPFDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_DRMSTR_C4", "VSPTODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C4", "VSPCNTCD", "Contractor Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DRMSTR_C4", "VSPCNTNM", "Contractor Name", 254)

    End Sub

    Sub CreatDropDownConfigScrn()

        objMain.objUtilities.CreateTable("VSP_FLT_DDCS", "DropDownCofigScrn", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS", "VSPFRMID", "Form ID", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS", "VSPFRMNM", "Form Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS", "VSPITCL", "ItemID or ColumnUID", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS", "VSPMATID", "MatrixUID", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS", "VSPACTV", "Active CheckBox", 1)

        objMain.objUtilities.CreateTable("VSP_FLT_DDCS_C0", "DropDownCofigScrn Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS_C0", "VSPVALUS", "Value", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_DDCS_C0", "VSPDESC", "Description", 100)

    End Sub

    Sub CreateTripSheet()

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT", "Trip Sheet", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPVHCL", "Vechicle", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPROUTE", "Route", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPSOURC", "Source", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPDEST", "Destination", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPCL1", "Cleaner 1", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPCL2", "Cleaner 2", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPCL3", "Cleaner 3", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPTANK", "Tank", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPDISIT", "Diesel Item", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPCNTR", "Contractor Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPCNTNM", "Contractor Name", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT", "VSPLCDT", "Last Calibration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        ' objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPGISU", "Goods Issue", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPEXPNO", "Expense JE", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPSTS", "Status", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT", "VSPDOCDT", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT", "VSPSRTDT", "Start Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT", "VSPENDDT", "End Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT", "VSPTOTKM", "Total KM's", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPTOTDY", "Total Days", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPMLGE", "Mileage", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPRTDET", "Route Details", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT", "VSPTYUP", "Tyre Master Updated", 1)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT", "VSPACMLG", "Auctual Mileage", SAPbobsCOM.BoFldSubTypes.st_Price)
        objUtilities.AddFloatField("@VSP_FLT_TRSHT", "VSPCOMSN", "Commission", SAPbobsCOM.BoFldSubTypes.st_Percentage)  ''Added by Abinas on 13-11-2018

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C1", "Trip Sheet Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C1", "VSPOPKM", "Open KM", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C1", "VSPCLKM", "Close KM", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C1", "VSPSOUR", "Source", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C1", "VSPDEST", "Destination", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C1", "VSPFRDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C1", "VSPFRTM", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C1", "VSPTODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C1", "VSPTOTM", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C1", "VSPLOAD", "Load", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C1", "VSPDICON", "Diesel Consumption", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C1", "TOTKM", "Total KM", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C1", "VSPLID", "LineId", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C2", "Trip Sheet Child 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C2", "VSPJOUDT", "Journey Details", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C2", "VSPAMGOO", "Amount Of Goods", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C2", "VSPADAMT", "Advance Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VSPFRACT", "From Account", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VAPTOACT", "To Account", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VSPDRVCC", "Driver CC", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VAPCAS", "Cashier", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VSPCOM", "Comments", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VSPOPNO", "Payment No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C2", "VSPJENO", "JE No.", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C3", "Trip Sheet Child 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPDOCTY", "DocType", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPGENTY", "Generation Type", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C3", "VSPDATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPDCNUM", "Document No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPSONUM", "Sales Order No.", 30)  ''added on 14-11-2018 by Abinas
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPBPCOD", "Customer/Vendor", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C3", "VSPQUANT", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPREF", "Reference No.", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C3", "VSPDCTOT", "Document Total", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C3", "VSPREM", "Remarks", 254)


        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C4", "Trip Sheet Child 4", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPFRRMA", "From Route Master", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPEXACC", "Expense Account Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPEXACN", "Expense Account Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPADACC", "Advance Account Code", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C4", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPBUD", "Budget", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C4", "VSPMTYP", "Material Type", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C5", "Trip Sheet Child 5", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C5", "VSPDATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C5", "VSPVENCO", "Vendor Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C5", "VSPVENNM", "Vendor Name", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C5", "VSPQUAN", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C5", "VSPRATE", "Rate", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C5", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C5", "VSPDRCC1", "Driver CC", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C5", "VSPDCNUM", "DocNumber", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C5", "VSPGISU", "Goods Issue", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C6", "Trip Sheet Child 6", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C6", "VSPDAT", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C6", "VSPSOU", "Source", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C6", "VSPSOUR", "Destination", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C6", "VSPFRDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C6", "VSPFRTM", "From Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C6", "VSPTODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C6", "VSPTOTM", "To Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C6", "VSPCHCOD", "Chemical Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C6", "VSPCHNAM", "Chemical Name", 100)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C6", "VSPWEIGH", "Weight", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C6", "VSPUOM", "UOM", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C7", "Trip Sheet Child 7", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C7", "VSPDRCOD", "Driver Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C7", "VSPDRFNM", "Driver First Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C7", "VSPDRMNM", "Driver Middle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C7", "VSPDRLNM", "Driver Last Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C7", "VSPMBNUM", "Mobile Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C7", "VSPLNNUM", "License Number", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C7", "VSPEXPDT", "Expration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C7", "VSPFRMDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_TRSHT_C7", "VSPTODT", "To Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C8", "Trip Sheet Child 8", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C8", "VSPTYRNO", "Tyre No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C8", "VSPTYRNM", "Tyre Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C8", "VSPTYMOD", "Tyre Model", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C8", "VSPTYPOS", "Tyre Position", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TRSHT_C8", "VSPKMS", "Kilometers", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C8", "VSPSTS", "Status", 30)


        objMain.objUtilities.CreateTable("VSP_FLT_TRSHT_C9", "Trip Sheet Child 9", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TRSHT_C9", "VSPATNM", "Attchment Name", 30)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_TRSHT_C9", "VSPAPTH", "Attchment Path", 64000)

    End Sub

    Sub CreateRouteMaster()
        objMain.objUtilities.CreateTable("VSP_FLT_RTMSTR", "Route Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR", "VSPRCD", "Route Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR", "VSPRNM", "Route Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR", "VSPSRCE", "Source", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR", "VSPDEST", "Destinatin", 100)
        objMain.objUtilities.AddFloatField("@VSP_FLT_RTMSTR", "VSPADAMT", "Advance Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@VSP_FLT_RTMSTR", "VSPTTLKM", "Total Kilometers", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_RTMSTR", "VSPATCH", "Attachment", 64000)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR", "VSPSHT", "SheetName", 100)
        objMain.objUtilities.AddFloatField("@VSP_FLT_RTMSTR", "VSPRTCMS", "Comission Percentage", SAPbobsCOM.BoFldSubTypes.st_Percentage) 'Abinas 17-09-2018


        objMain.objUtilities.CreateTable("VSP_FLT_RTMSTR_C0", "Route Master Child 1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR_C0", "VSPEACD", "Expenses Account Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR_C0", "VSPEANM", "Expenses Name", 100)
        objMain.objUtilities.AddFloatField("@VSP_FLT_RTMSTR_C0", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)

        objMain.objUtilities.CreateTable("VSP_FLT_RTMSTR_C1", "Route Master Child 2", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR_C1", "VSPFPLC", "From Place", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR_C1", "VSPTPLC", "To Place", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_RTMSTR_C1", "VSPDISTN", "Distance", 100)

    End Sub

    Sub CreatTaskMaster()

        objMain.objUtilities.CreateTable("VSP_FLT_TSKMSTR", "Task Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TSKMSTR", "VSPTNUM", "Task Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TSKMSTR", "VSPTDSC", "Task Description", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TSKMSTR", "VSPTTYP", "Task Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TSKMSTR", "VSPRMRK", "Remarks", 254)

    End Sub

    Sub CreateTyreMaster()

        objMain.objUtilities.CreateTable("VSP_FLT_TYRMSTR", "Tyre Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTRNUM", "Tyre No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPWHL", "Wheel", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPCPCTY", "Capacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM1", "UOM1", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTARC", "TotalAirCapacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPMINAC", "MinAirCapacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPMAXAC", "MaxAirCapacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTBTP", "Tube Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPINTTP", "Inner Tube Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM7", "UOM7", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPINTC", "InnerTubeCapcity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM8", "UOM8", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTRNM", "TyreName", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTRMDL", "TyreModel", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTRSZE", "Tyre Size", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM2", "UOM2", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPRPK", "RevolutionPerSec", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPSOASN", "SizeofAirSysNozl", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM5", "UOM5", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPPCHFM", "PurchaseFrom", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TYRMSTR", "VSPPCHON", "PurchaseOn", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPMOP", "Mode Of Purchase", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTRTYP", "Tyre Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPSLOC", "Storage Location", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPCAS", "CapacityOfAirStem", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM6", "UOM6", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPMNLC", "MinLoadofCapacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM3", "UOM3", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPMAXLC", "MaxLoadOfCapacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPUOM4", "UOM4", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPPNO", "Plot No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPWNTY", "Warranty", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPMFR", "Manufacturur", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPREMK", "Remarks", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPITCD", "GR ItemCode", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TYRMSTR", "VSPPRIC", "Price", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPPSTN", "Position", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPVNO", "Vechical No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPSRLNO", "Serial No", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPTXCD", "TaxCode", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPKMRUN", "KM's Run", 30)

        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMSTR", "VSPGIIC", "GI ItemCode", 30)

    End Sub

    Sub CreateTyreMaintenance()

        objMain.objUtilities.CreateTable("VSP_FLT_TYRMTNC", "Tyre Maintenance", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC", "VSPVNO", "Vechicle Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC", "VSPVNM", "Vechicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC", "VSPVMD", "Vechicle Model", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TYRMTNC", "VSPDCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC", "VSPODMTR", "Odo Meter", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC", "VSPRMK", "Remarks", 254)

        objMain.objUtilities.CreateTable("VSP_FLT_TYRMTNC_C1", "Tyre Maintenance Child 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC_C1", "VSPDOCTY", "DocType", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC_C1", "VSPDOCNM", "DocNum ", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TYRMTNC_C1", "VSPDATE", "Date ", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TYRMTNC_C1", "VSPDCTOT", "DocTotal ", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC_C1", "VSPCOMM", "Comments", 254)

        objMain.objUtilities.CreateTable("VSP_FLT_TYRMTNC_C2", "Tyre Maintenance Child 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_TYRMTNC_C2", "VSPPTH", "Path", 64000)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMTNC_C2", "VSPFLNM", "File Name", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TYRMTNC_C2", "VSPDT", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub

    Sub CreateTyreMapping()

        objMain.objUtilities.CreateTable("VSP_FLT_TYRMPG", "Tyre Mapping", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG", "VSPVCHN0", "Vechicle Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG", "VSPVCHNM", "Vechicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG", "VSPVMDL", "Vechicle Model", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG", "VSPREM", "Remarks", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TYRMPG_C0", "Tyre Mapping Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPTRNUM", "Tyre Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPTRNM", "Tyre Name ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPWLTYP", "Wheel Type ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPTRSIZ", "Tyre Size ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPCPCTY", "Capacity ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPUOM", "UOM", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPPSTN", "Position", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPSTS", "Status", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPUOM1", "UOM1", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPITCD", "GR Item Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPGDRPT", "Goods Receipt", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPGDIS", "Goods Issue", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TYRMPG_C0", "VSPRMKM", "Removal Km", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TYRMPG_C0", "VSPISKM", "Issues Km", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TYRMPG_C0", "VSPGIIC", " GI Item Code", 30)

    End Sub

    Sub CreateBreakDownEntry()

        objMain.objUtilities.CreateTable("VSP_FLT_BRKDETRY", "BreakDown Entry", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY", "VSPVCD", "Vehicle Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY", "VSPVNM", "Vehicle Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY", "VSPVHN0", "Vehicle No.", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_BRKDETRY", "VSPDDT", "Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("VSP_FLT_BRKDETRY_C0", "BreakDownEntry Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPTCD", "Task Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPTDSC", "Task Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPLOC", "Location", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_BRKDETRY_C0", "VSPDT", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_BRKDETRY_C0", "VSPTIM", "Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPSTS", "Status", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPOWNR", "Owner", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPTYP", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPDRVR", "Driver", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPRQBY", "Requested By", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_BRKDETRY_C0", "VSPPODNM", "Service Order No.", 30)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_BRKDETRY_C0", "VSPATMNT", "Attachments", 64000)

    End Sub

    Sub Callibration()

        objMain.objUtilities.CreateTable("VSP_FLT_CALBRTN", "Callibration", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPVCHID", "Vehicle Id", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPVHNME", "Vehicle Name", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_CALBRTN", "VSPLCDT", "Last Callibration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPLRM1", "Remarks 1", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPDBY", "Did By", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPVINCG", "Vehicle Incharge", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPVINM", "Vehicle Incharge Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPQCI", "Quality Control Incharge", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPPDBY", "Planned By", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_CALBRTN", "VSPNCD", "Next Callibration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPRMK2", "Remarks 2", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_CALBRTN", "VSPDUDT", "DueDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_CALBRTN", "VSPCCDT", "Cur Callibration Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPAPBY", "Approved By", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPCBY", "Callibrated By", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN", "VSPRMK", "Remarks", 30)


        objMain.objUtilities.CreateTable("VSP_FLT_CALBRTN_C0", "Callibration Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN_C0", "VSPTYP", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN_C0", "VSPNAM", "Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN_C0", "VSPVAL", "Value", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN_C0", "VSPATTCH", "Attachments", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALBRTN_C0", "VSPGI", "GoodsIssue", 30)

    End Sub

    Sub PreventiveReminder()

        objMain.objUtilities.CreateTable("VSP_FLT_PRTVRMDR", "PreventiveReminder", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRMDCD", "Reminder Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPREMNM", "Reminder Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPTYPE", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRMFR", "Reminder For", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRMVCD", "Reminder Vehicle Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRVNM", "Reminder Vehicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPTKNUM", "Task Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPTTYP", "Task Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPTDESC", "Task Description", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_PRTVRMDR", "VSPTHD", "Tresh Hold Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_PRTVRMDR", "VSPRMDT", "Reminder Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPREMTO", "Reminder To ", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_PRTVRMDR", "VSPCDT", "Created Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPSATS", "Status", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_PRTVRMDR", "VSPCLSDT", "Closed Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_PRTVRMDR", "VSPCLSTM", "Close Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRBON", "Reminder Based On", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPTHR", "Thresh Hold Reading", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRMRDG", "Reminder Reading", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRMDBY", "Remind By", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_PRTVRMDR", "VSPRSN", "Reason", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_PRTVRMDR", "VSPKMS", "Kilometers", SAPbobsCOM.BoFldSubTypes.st_Measurement)
    End Sub

    Sub CreateTyrePostionMstr()

        objMain.objUtilities.CreateTable("VSP_FLT_TPMSTR", "Tyre Position Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TPMSTR", "VSPPSTN", "Vechical Postion", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TPMSTR", "VSPNAM", "Name", 100)

    End Sub

    Sub CreateCalibrationMaster()

        objMain.objUtilities.CreateTable("VSP_FLT_CALMSTR", "Calibration Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALMSTR", "VSPTYP", "Type ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CALMSTR", "VSPNAM", "Name", 100)

    End Sub

    Sub CreateConfigurationScreen()
        objMain.objUtilities.CreateTable("VSP_FLT_CNFGSRN", "Configuration Screen", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPDEWHS", "Warehouse Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPCUSCD", "Customer Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPTXCD", "TaxCode", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPPSWD", "Password", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPOACT", "Account Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPLOCCD", "Location Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPFLVCD", "Fuel Vendor Code", 30)  ''Added on 26-09-2018 by Abinas

        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPDD", "Drivers Department", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPCNT", "Contractor", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPVHGRP", "Vehicle Group", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPTYRGR", "Tyre Group", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_CNFGSRN", "VSPDSLGR", "Diesel Group", 30)
    End Sub

    Sub CreateFuelUpload()

        objMain.objUtilities.CreateTable("VSP_FLT_FLUPLD", "Fuel Upload", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD", "VSPVCD", "Vendor Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD", "VSPVNAM", "Vendor Name", 30)
        objMain.objUtilities.AddAlphaMemoField("@VSP_FLT_FLUPLD", "VSPATCH", "Attachment Path", 64000)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD", "VSPUPLDR", "Uploader Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD", "VSPSHTNM", "Sheet Name ", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_FLUPLD", "VSPDT", "Document Date", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("VSP_FLT_FLUPLD_C0", "Fuel Upload Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPID", "ID", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_FLUPLD_C0", "VSPTDT", "Transaction Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPDLN", "Dealer Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPLCT", "Location", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCID", "Customer ID", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_FLUPLD_C0", "VSPQTY", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPVNO", "Vehical Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCR", "Curency", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_FLUPLD_C0", "VSPAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPBAL", " Balance", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPEXPT", "Extra Point", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPGRPO", "GRPO Document Num", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPTRNO", "Trip Sheet Num", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPISSNO", "Issue Num", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCC1", "Cost Center 1", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCC2", "Cost Center 2", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCC3", "Cost Center 3", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCC4", "Cost Center 4", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_FLUPLD_C0", "VSPCC5", "Cost Center 5", 30)

    End Sub

    Sub CreateTankMaster()
        objMain.objUtilities.CreateTable("VSP_FLT_TANKMSTR", "Tank Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPTNKNO", "Tank No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPTNKNM", "Tank Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPTNKMD", "Tank Model", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPTTYPE", "Tank Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPFTYPE", "Fuel Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPCPCTY", "Capacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPUOM1", "Capacity UoM", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPMAXLC", "Max Load Capacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPUOM4", "UoM 4", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPREMK", "Remarks", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPVNO", "Vehicle No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPWNTY", "Waranty", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPMFR", "Manufacturer", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPMNLC", "Min Load Capacity", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPUOM3", "UoM 3", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR", "VSPTNKCC", "Tank CC", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_TANKMSTR_C0", "Tank Master Child 1", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR_C0", "VSPITMCD", "Item Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR_C0", "VSPITMNM", "Item Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMSTR_C0", "VSPCPTY", "Capacity", 30)
    End Sub

    Sub CreateTankMapping()

        objMain.objUtilities.CreateTable("VSP_FLT_TANKMPG", "Tank Mapping", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG", "VSPVCHN0", "Vechicle Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG", "VSPVCHNM", "Vechicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG", "VSPVMDL", "Vechicle Model", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG", "VSPREM", "Remarks", 254)

        objMain.objUtilities.CreateTable("VSP_FLT_TANKMPG_C0", "Tank Mapping Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG_C0", "VSPTNUM", "Tank Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG_C0", "VSPTNM", "Tank Name ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG_C0", "VSPTYP", "Tank Type ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG_C0", "VSPCPCTY", "Capacity ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG_C0", "VSPUOM", "UOM", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TANKMPG_C0", "VSPSTS", "Status", 30)

    End Sub

    Sub CreateTankMaintenance()

        objMain.objUtilities.CreateTable("VSP_FLT_TNKMTNC", "Tank Maintenance", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC", "VSPVNO", "Vechicle Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC", "VSPVNM", "Vechicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC", "VSPVMD", "Vechicle Model", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TNKMTNC", "VSPDCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC", "VSPODMTR", "Odo Meter", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC", "VSPRMK", "Remarks", 254)

        objMain.objUtilities.CreateTable("VSP_FLT_TNKMTNC_C1", "Tank Maintenance Child 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC_C1", "VSPDOCTY", "DocType", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC_C1", "VSPDOCNM", "DocNum ", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TNKMTNC_C1", "VSPDATE", "Date ", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_TNKMTNC_C1", "VSPDCTOT", "DocTotal ", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC_C1", "VSPCOMM", "Comments", 254)

        objMain.objUtilities.CreateTable("VSP_FLT_TNKMTNC_C2", "Tank Maintenance Child 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC_C2", "VSPPTH", "Path", 254)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_TNKMTNC_C2", "VSPFLNM", "File Name", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_TNKMTNC_C2", "VSPDT", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
    End Sub

    Sub CreateAccidentHistory()

        objMain.objUtilities.CreateTable("VSP_FLT_ACCHIST", "Accident History", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPVHCD", "Vechicle Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPVHNM", "Vechicle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPPLTNO", "Plate No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPTRPNO", "Trip No.", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPLOC", "Location", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPCNTR", "Contractor", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPSRVTY", "Severity", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_ACCHIST", "VSPDATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_ACCHIST", "VSPTIME", "Time", SAPbobsCOM.BoFldSubTypes.st_Time)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPSTS", "Status", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPTCLMD", "Total Claimed", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPTEXP", "Total Expenses", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPODOMT", "OdoMeter", 30)
        objMain.objUtilities.AddImageField("@VSP_FLT_ACCHIST", "VSPIMG1", "Image 1")
        objMain.objUtilities.AddImageField("@VSP_FLT_ACCHIST", "VSPIMG2", "Image 2")
        objMain.objUtilities.AddImageField("@VSP_FLT_ACCHIST", "VSPIMG3", "Image 3")
        objMain.objUtilities.AddImageField("@VSP_FLT_ACCHIST", "VSPIMG4", "Image 4")
        objMain.objUtilities.AddImageField("@VSP_FLT_ACCHIST", "VSPIMG5", "Image 5")
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST", "VSPTYPE", "Type", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_ACCHIST_C0", "Accident History Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C0", "VSPDOCTY", "DocType", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C0", "VSPDOCNM", "DocNum ", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_ACCHIST_C0", "VSPDATE", "Date ", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_ACCHIST_C0", "VSPDCTOT", "DocTotal ", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C0", "VSPCOMM", "Comments", 254)

        objMain.objUtilities.CreateTable("VSP_FLT_ACCHIST_C1", "Accident History Child 2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C1", "VSPANM", "Attachment Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C1", "VSPAPATH", "Attachment Path ", 30)

        objMain.objUtilities.CreateTable("VSP_FLT_ACCHIST_C2", "Accident History Child 3", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C2", "VSPDRCD", "Driver Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C2", "VSPDRFNM", "Driver First Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C2", "VSPDRMNM", "Driver Middle Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C2", "VSPDRLNM", "Driver Last Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C2", "VSPMBNO", "Mobile Number", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_ACCHIST_C2", "VSPLCNO", "License Number", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_ACCHIST_C2", "VSPEXPDT", "Expration Date", SAPbobsCOM.BoFldSubTypes.st_None)

    End Sub

    Sub CreateInsurancePayments()

        objMain.objUtilities.CreateTable("VSP_FLT_INSPAY", "Insurance Payments", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_INSPAY", "VSPVNO", "Vechicle Code", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_INSPAY", "VSPVNM", "Vechicle Name", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_INSPAY", "VSPINAMT", "Insurance Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddInteger("@VSP_FLT_INSPAY", "VSPINST", "No. of Installments", SAPbobsCOM.BoFldSubTypes.st_None, 11)

        objMain.objUtilities.CreateTable("VSP_FLT_INSPAY_C0", "Insurance Payments Child 1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@VSP_FLT_INSPAY_C0", "VSPFRMDT", "From Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddDateField("@VSP_FLT_INSPAY_C0", "VSPTODT", "To Date ", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSP_FLT_INSPAY_C0", "VSPAMT", "Actual Amount ", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddFloatField("@VSP_FLT_INSPAY_C0", "VSPTOTPY", "Total Payments ", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_INSPAY_C0", "VSPJENO", "JE No. ", 30)
        objMain.objUtilities.AddDateField("@VSP_FLT_INSPAY_C0", "VSPDATE", "Date ", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_INSPAY_C0", "VSPCSHAC", "Cash Account ", 30)
        objMain.objUtilities.AddAlphaField("@VSP_FLT_INSPAY_C0", "VSPINSAC", "Insurance Account", 30)
        objMain.objUtilities.AddFloatField("@VSP_FLT_INSPAY_C0", "VSPTBPD", "Amount to be Paid ", SAPbobsCOM.BoFldSubTypes.st_Price)
    End Sub

    Sub CreateVehicleStatus()
        objMain.objUtilities.CreateTable("VSP_VECHSTS", "Vehicle Status", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddDateField("@VSP_VECHSTS", "VSPDT", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPVNO", "VehicleNo", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPCNTR", "Contarctor", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPCNM", "Contracor Name", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPDRIV", "Driver", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPCLN", "Cleaner", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPROUTE", "Route", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPTRSHT", "Trip SheetNo", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS", "VSPSTS", "Trip Sheet Status", 30)

        objMain.objUtilities.CreateTable("VSP_VECHSTS_C0", "Vehicle Status Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPTY", "Type", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPSTA", "Status", 30)
        objMain.objUtilities.AddFloatField("@VSP_VECHSTS_C0", "VSPOPKM", "Open KM", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddFloatField("@VSP_VECHSTS_C0", "VSPCLKM", "Close KMS", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPSOURC", "Source", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPDST", "Destination", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPLOC", "Location", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPSLNO", "Sales No", 30)    ''Added by Abinas on 13-11-2018
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPCHML", "Chemical", 30)
        objMain.objUtilities.AddFloatField("@VSP_VECHSTS_C0", "VSPQUA", "Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        'objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPCSVE", "Cusomer/Vendor", 30)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPREM", "Remarks", 254)
        objMain.objUtilities.AddFloatField("@VSP_VECHSTS_C0", "VSPTOTKM", "Total KM", SAPbobsCOM.BoFldSubTypes.st_Measurement)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPCHK", "Check", 1)
        objMain.objUtilities.AddAlphaField("@VSP_VECHSTS_C0", "VSPCHK1", "Check1", 1)

    End Sub



    Sub CreateDeliveryConfirmation()
        objMain.objUtilities.CreateTable("VSPDELCONF", "Delivery Confirmation", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSPDELCONF", "VSPDLENT", "Delevery Entry", 30)
        objMain.objUtilities.AddAlphaField("@VSPDELCONF", "VSPDLNO", "Delivery No", 30)
        objMain.objUtilities.AddDateField("@VSPDELCONF", "VSPDLDT", "Delivery Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSPDELCONF", "VSPITCOD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@VSPDELCONF", "VSPITNM", "Item Name", 50)
        objMain.objUtilities.AddFloatField("@VSPDELCONF", "VSPDLQTY", "Delivery Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddDateField("@VSPDELCONF", "VSPDOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@VSPDELCONF", "VSPDOCST", "Doc Status", 254)
        objMain.objUtilities.AddDateField("@VSPDELCONF", "VSPADLDT", "Actual Delvery Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddFloatField("@VSPDELCONF", "VSPADLQT", "Actual Delivery Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@VSPDELCONF", "VSPDFQTY", "Difference Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@VSPDELCONF", "VSPTLQTY", "Tollerence Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
        objMain.objUtilities.AddFloatField("@VSPDELCONF", "VSPSHQTY", "SHortage Quantity", SAPbobsCOM.BoFldSubTypes.st_Quantity)
    End Sub


#End Region

End Class





