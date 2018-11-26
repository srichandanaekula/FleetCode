Imports System.Threading
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class clsImportData

#Region " Declaration         "
    Dim objForm As SAPbouiCOM.Form
    Dim odt As SAPbouiCOM.DataTable
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objComboBox As SAPbouiCOM.ComboBox
#End Region

    Sub ImportDataScreen()
        Try

            objMain.objUtilities.LoadForm("ImportData.xml", "VSP_IMPDATA_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_IMPDATA_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.DataSources.UserDataSources.Add("U_SERVER", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_DBNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_USRNME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_PWD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_PATH", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SHTNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_DOCTYPE", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oEditText = objForm.Items.Item("4").Specific
            oEditText.DataBind.SetBound(True, "", "U_SERVER")
            objForm.Items.Item("4").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_SERVER").Value = objMain.objCompany.Server

            oEditText = objForm.Items.Item("6").Specific
            oEditText.DataBind.SetBound(True, "", "U_DBNAME")
            objForm.Items.Item("6").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_DBNAME").Value = objMain.objCompany.CompanyDB

            oEditText = objForm.Items.Item("8").Specific
            oEditText.DataBind.SetBound(True, "", "U_USRNME")
            objForm.Items.Item("8").AffectsFormMode = False
            objForm.DataSources.UserDataSources.Item("U_USRNME").Value = "sa"

            oEditText = objForm.Items.Item("10").Specific
            oEditText.DataBind.SetBound(True, "", "U_PWD")
            objForm.Items.Item("10").AffectsFormMode = False
            oEditText.IsPassword = True

            oEditText = objForm.Items.Item("12").Specific
            oEditText.DataBind.SetBound(True, "", "U_PATH")
            objForm.Items.Item("12").AffectsFormMode = False
            objForm.Items.Item("12").Enabled = True

            oEditText = objForm.Items.Item("14").Specific
            oEditText.DataBind.SetBound(True, "", "U_SHTNAME")
            objForm.Items.Item("14").AffectsFormMode = False

            objComboBox = objForm.Items.Item("16").Specific
            objComboBox.DataBind.SetBound(True, "", "U_DOCTYPE")
            objForm.Items.Item("16").AffectsFormMode = False
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Vehicle Master", "Vehicle Master")
            objComboBox.ValidValues.Add("Tyre Master", "Tyre Master")
            objComboBox.Select("Vehicle Master", SAPbouiCOM.BoSearchKey.psk_ByValue)

            odt = objForm.DataSources.DataTables.Add("dt")
            odt = objForm.DataSources.DataTables.Item("dt")

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_IMPDATA" And pVal.BeforeAction = False Then
                Me.ImportDataScreen()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Select Case pVal.EventType
            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.ItemUID = "15" And pVal.BeforeAction = False Then
                    If Me.ImportData(objForm.UniqueID, objForm.Items.Item("4").Specific.Value, _
                               objForm.Items.Item("6").Specific.Value, objForm.Items.Item("8").Specific.Value, _
                               objForm.Items.Item("10").Specific.Value, objForm.Items.Item("12").Specific.Value, _
                               objForm.Items.Item("14").Specific.Value) = True Then
                        If objForm.Items.Item("16").Specific.Selected.Value = "Vehicle Master" Then
                            Me.GenerateVehicleMaster(objForm.UniqueID)
                        ElseIf objForm.Items.Item("16").Specific.Selected.Value = "Tyre Master" Then
                            Me.GenerateTyreMaster(objForm.UniqueID)
                        End If

                    End If
                End If
            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                objForm = objMain.objApplication.Forms.Item(FormUID)
                If pVal.ItemUID = "12" And pVal.Before_Action = False Then
                    Me.BrowseFileDialog()
                End If
        End Select
    End Sub

    Function ImportData(ByVal FormUID As String, ByVal ServerName As String, _
         ByVal DBName As String, ByVal UserName As String, _
         ByVal Password As String, ByVal ExcelPath As String, ByVal SheetName As String)
        Try
            Dim ExceCon As String = _
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                ExcelPath & "; Extended Properties=Excel 12.0"
            Dim connectionString As String = "Server=" & ServerName & " ;" & _
                                             "DataBase=" & DBName & ";" & _
                                             "Uid=" & UserName & ";Pwd=" & Password & ";"

            Dim strSQL As String = "USE [" & DBName & "]" & vbCrLf & _
                                   "IF EXISTS (" & _
                                   "SELECT * " & _
                                   "FROM [" & DBName & "].dbo.sysobjects " & _
                                   "WHERE Name = '" & SheetName & "')" & vbCrLf & _
                                   "BEGIN" & vbCrLf & _
                                   "DROP TABLE [" & DBName & "].dbo." & SheetName & vbCrLf & _
                                   "END "

            Dim dbConnection As New SqlConnection(connectionString)
            ' A SqlCommand object is used to execute the SQL commands.
            Dim cmd As New SqlCommand(strSQL, dbConnection)
            dbConnection.Open()
            cmd.ExecuteNonQuery()
            dbConnection.Close()
            Dim excelConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(ExceCon)
            excelConnection.Open()
            Dim OleStr As String = "SELECT * INTO [ODBC; Driver={SQL Server};Server=" _
                                   & ServerName & ";Database=" & DBName & ";Uid=" & _
                                   UserName & ";Pwd=" & Password & "; ].[" & _
                                   SheetName & "]   FROM [" & SheetName & "$];"

            Dim excelCommand As New System.Data.OleDb.OleDbCommand(OleStr, excelConnection)
            excelCommand.ExecuteNonQuery()
            excelConnection.Close()
            cmd = New SqlCommand("DELETE TOP (1) FROM" & " " & SheetName & " Where Cast(Code as VarChar) = 'Code'", dbConnection)
            dbConnection.Open()
            cmd.ExecuteNonQuery()
            dbConnection.Close()
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

#Region "    Attachment    "

    Sub BrowseFileDialog()
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA)
                ShowFolderBrowserThread.Start()

            ElseIf ShowFolderBrowserThread.ThreadState = ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.StackTrace)
        End Try

    End Sub

    Sub ShowFolderBrowser()
        Dim MyTest1 As New OpenFileDialog
        Dim MyProcs() As Process
        MyProcs = Process.GetProcessesByName("SAP Business One")
        If MyProcs.Length <> 0 Then
            For i As Integer = 0 To MyProcs.Length - 1

                Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                MyTest1.FileName = "Select the Reference Document"
                'Windows XP
                If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                    Try
                        objForm.DataSources.UserDataSources.Item("U_PATH").Value = MyTest1.FileName
                    Catch ex As IO.IOException
                        objMain.objApplication.MessageBox(ex.Message)
                        Exit Sub
                    End Try
                    'Windows 7
                    'AttachMentPath +
                ElseIf MyTest1.ShowDialog() = DialogResult.OK Then
                    Try
                        objForm.DataSources.UserDataSources.Item("U_PATH").Value = MyTest1.FileName
                    Catch ex As IO.IOException
                        objMain.objApplication.MessageBox(ex.Message)
                        Exit Sub
                    End Try
                    System.Windows.Forms.Application.ExitThread()
                End If
            Next
        Else
            Console.WriteLine("No SBO instances found.")
        End If
    End Sub
#End Region

    Sub GenerateVehicleMaster(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            Dim GetDocNum As String = ""

            If objMain.IsSAPHANA = True Then
                GetDocNum = "Select ""Code"" From " & objForm.Items.Item("14").Specific.Value & " Where ""Code"" IS NOT NULL And ""Code"" <> '' Group By ""Code"""
            Else
                GetDocNum = "Select Code From " & objForm.Items.Item("14").Specific.Value & " Where Code IS NOT NULL And Code <> '' Group By Code"
            End If
            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocNum.DoQuery(GetDocNum)

            If oRsGetDocNum.RecordCount > 0 Then
                oRsGetDocNum.MoveFirst()
                For i As Integer = 1 To oRsGetDocNum.RecordCount
                    Dim GetDetails As String = ""

                    If objMain.IsSAPHANA = True Then
                        GetDetails = "Select * From " & objForm.Items.Item("14").Specific.Value & " Where ""Code"" = '" & oRsGetDocNum.Fields.Item(0).Value & "'"
                    Else
                        GetDetails = "Select * From " & objForm.Items.Item("14").Specific.Value & " Where Code = '" & oRsGetDocNum.Fields.Item(0).Value & "'"
                    End If
                    Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetDetails.DoQuery(GetDetails)

                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
                    objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    objMain.oGeneralData.SetProperty("Code", objMain.objUtilities.getMaxCode("@VSP_FLT_VMSTR"))
                    objMain.oGeneralData.SetProperty("U_VSPVNO", oRsGetDetails.Fields.Item("VehicleNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPVNM", oRsGetDetails.Fields.Item("VehicleName").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPNO", oRsGetDetails.Fields.Item("PlateNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTYPE", oRsGetDetails.Fields.Item("Type").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPOWSHP", oRsGetDetails.Fields.Item("Ownership").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPODRDG", oRsGetDetails.Fields.Item("OdometerReading").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCAT", oRsGetDetails.Fields.Item("Category").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPLOC", oRsGetDetails.Fields.Item("Location").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMWL", oRsGetDetails.Fields.Item("Mileage(WithLoad)").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMWOL", oRsGetDetails.Fields.Item("Mileage(WithoutLoad)").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPYEAR", oRsGetDetails.Fields.Item("Year").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMAKE", oRsGetDetails.Fields.Item("Make").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMODEL", oRsGetDetails.Fields.Item("Model").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSRNO", oRsGetDetails.Fields.Item("SerialNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCOLOR", oRsGetDetails.Fields.Item("Color").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPIDN", oRsGetDetails.Fields.Item("Identification").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPLEN", oRsGetDetails.Fields.Item("Length").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPWIDTH", oRsGetDetails.Fields.Item("Width").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPHGT", oRsGetDetails.Fields.Item("Height").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPWBASE", oRsGetDetails.Fields.Item("WheelBase").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPENGSZ", oRsGetDetails.Fields.Item("EngineSize").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPNOCYL", oRsGetDetails.Fields.Item("NoOfCylinders").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTTYPE", oRsGetDetails.Fields.Item("TransmissionType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPFTYPE", oRsGetDetails.Fields.Item("FuelType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSPPLG", oRsGetDetails.Fields.Item("SparkPlugType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPBTRY", oRsGetDetails.Fields.Item("BatteryType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPHDLMP", oRsGetDetails.Fields.Item("HeadLampType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMSC1", oRsGetDetails.Fields.Item("Misc1").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMSC2", oRsGetDetails.Fields.Item("Misc2").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPNOAXL", oRsGetDetails.Fields.Item("NoOfAxles").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSZFT", oRsGetDetails.Fields.Item("Size(FrontTyres)").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPRFT", oRsGetDetails.Fields.Item("Pressure(FrontTyres)").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSZRT", oRsGetDetails.Fields.Item("Size(RearTyres)").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPRRT", oRsGetDetails.Fields.Item("Pressure(RearTyres)").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPDTFT", "N")
                    objMain.oGeneralData.SetProperty("U_VSPDTRT", "N")
                    objMain.oGeneralData.SetProperty("U_VSPAVLB", "Y")
                    objMain.oGeneralData.SetProperty("U_VSPTAC", oRsGetDetails.Fields.Item("TankAttachedCapacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPVAL", oRsGetDetails.Fields.Item("Value").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPENGNO", oRsGetDetails.Fields.Item("EngineNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCHSNO", oRsGetDetails.Fields.Item("ChasisNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTYRES", oRsGetDetails.Fields.Item("NoOfTyres").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPDIITE", oRsGetDetails.Fields.Item("DieselItemCode").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCDWGT", oRsGetDetails.Fields.Item("CabinWeight").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTWGT", oRsGetDetails.Fields.Item("TankWeight").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCHWGT", oRsGetDetails.Fields.Item("ChasisWeight").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPOTHFT", oRsGetDetails.Fields.Item("OtherFittings").Value.ToString)
                    Dim TareWeight As String = oRsGetDetails.Fields.Item("CabinWeight").Value + oRsGetDetails.Fields.Item("TankWeight").Value _
                                                + oRsGetDetails.Fields.Item("ChasisWeight").Value + oRsGetDetails.Fields.Item("OtherFittings").Value
                    objMain.oGeneralData.SetProperty("U_VSPGWGT", TareWeight)

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C0")
                    objMain.oChild = objMain.oChildren.Add()

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C1")
                    objMain.oChild = objMain.oChildren.Add()

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C2")
                    objMain.oChild = objMain.oChildren.Add()

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C5")
                    objMain.oChild = objMain.oChildren.Add()

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C6")
                    objMain.oChild = objMain.oChildren.Add()

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C7")
                    objMain.oChild = objMain.oChildren.Add()

                    objMain.oGeneralService.Add(objMain.oGeneralData)

                    oRsGetDocNum.MoveNext()
                Next
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub GenerateTyreMaster(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            Dim GetDocNum As String = ""
            If objMain.IsSAPHANA = True Then
                GetDocNum = "Select ""Code"" From " & objForm.Items.Item("14").Specific.Value & " Where ""Code"" IS NOT NULL And ""Code"" <> '' Group By ""Code"""
            Else
                GetDocNum = "Select Code From " & objForm.Items.Item("14").Specific.Value & " Where Code IS NOT NULL And Code <> '' Group By Code"
            End If

            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocNum.DoQuery(GetDocNum)

            If oRsGetDocNum.RecordCount > 0 Then
                oRsGetDocNum.MoveFirst()
                For i As Integer = 1 To oRsGetDocNum.RecordCount
                    Dim GetDetails As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetDetails = "Select * From " & objForm.Items.Item("14").Specific.Value & " Where ""Code"" = '" & oRsGetDocNum.Fields.Item(0).Value & "'"
                    Else
                        GetDetails = "Select * From " & objForm.Items.Item("14").Specific.Value & " Where Code = '" & oRsGetDocNum.Fields.Item(0).Value & "'"
                    End If
                    Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetDetails.DoQuery(GetDetails)

                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMSTR")
                    objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    objMain.oGeneralData.SetProperty("Code", objMain.objUtilities.getMaxCode("@VSP_FLT_TYRMSTR"))
                    objMain.oGeneralData.SetProperty("U_VSPTRNUM", oRsGetDetails.Fields.Item("TyreNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPWHL", oRsGetDetails.Fields.Item("Wheel").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCPCTY", oRsGetDetails.Fields.Item("Capacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM1", oRsGetDetails.Fields.Item("UOM1").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTARC", oRsGetDetails.Fields.Item("TotalAirCapacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMINAC", oRsGetDetails.Fields.Item("MinAirCapacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMAXAC", oRsGetDetails.Fields.Item("MaxAirCapacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTBTP", oRsGetDetails.Fields.Item("TubeType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPINTTP", oRsGetDetails.Fields.Item("InnerTubeType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM7", oRsGetDetails.Fields.Item("UOM7").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPINTC", oRsGetDetails.Fields.Item("InnerTubeCapcity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM8", oRsGetDetails.Fields.Item("UOM8").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTRNM", oRsGetDetails.Fields.Item("TyreName").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTRMDL", oRsGetDetails.Fields.Item("TyreModel").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTRSZE", oRsGetDetails.Fields.Item("TyreSize").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM2", oRsGetDetails.Fields.Item("UOM2").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPRPK", oRsGetDetails.Fields.Item("RevolutionPerSec").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSOASN", oRsGetDetails.Fields.Item("SizeofAirSysNozl").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM5", oRsGetDetails.Fields.Item("UOM5").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPCHFM", oRsGetDetails.Fields.Item("PurchaseFrom").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPCHON", oRsGetDetails.Fields.Item("PurchaseOn").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMOP", oRsGetDetails.Fields.Item("ModeOfPurchase").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTRTYP", oRsGetDetails.Fields.Item("TyreType").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSLOC", oRsGetDetails.Fields.Item("StorageLocation").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPCAS", oRsGetDetails.Fields.Item("CapacityOfAirStem").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM6", oRsGetDetails.Fields.Item("UOM6").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMNLC", oRsGetDetails.Fields.Item("MinLoadofCapacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM3", oRsGetDetails.Fields.Item("UOM3").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMAXLC", oRsGetDetails.Fields.Item("MaxLoadOfCapacity").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPUOM4", oRsGetDetails.Fields.Item("UOM4").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPWNTY", oRsGetDetails.Fields.Item("Warranty").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPMFR", oRsGetDetails.Fields.Item("Manufacturur").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPREMK", oRsGetDetails.Fields.Item("Remarks").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPITCD", oRsGetDetails.Fields.Item("ItemCode").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPRIC", oRsGetDetails.Fields.Item("Price").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPSTN", oRsGetDetails.Fields.Item("Position").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSRLNO", oRsGetDetails.Fields.Item("SerialNo").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTXCD", oRsGetDetails.Fields.Item("TaxCode").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPKMRUN", oRsGetDetails.Fields.Item("KMRun").Value.ToString)

                    objMain.oGeneralService.Add(objMain.oGeneralData)

                    oRsGetDocNum.MoveNext()
                Next
            End If


            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
