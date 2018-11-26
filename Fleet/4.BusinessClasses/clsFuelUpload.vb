Imports System.Threading
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class clsFuelUpload

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Fuel Upload.xml", "VSP_FLT_FLUPLD_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_FLUPLD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD_C0")

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.CFLFilterVendors(objForm.UniqueID, "CFL_VCD")

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            objMatrix = objForm.Items.Item("11").Specific
            objMatrix.Columns.Item("V_13").Visible = False
            objMatrix.Columns.Item("V_14").Visible = False
            objMatrix.Columns.Item("V_15").Visible = False
            objMatrix.Columns.Item("V_16").Visible = False
            objMatrix.Columns.Item("V_17").Visible = False

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OFLUPLD"))
            oDBs_Head.SetValue("U_VSPUPLDR", oDBs_Head.Offset, objMain.objCompany.UserName)
           
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_FLUPLD" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_FLUPLD_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("11").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount
                    If objMatrix.IsRowSelected(i) = True And objMatrix.Columns.Item("V_11").Cells.Item(i).Specific.Value = "" Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "19" And pVal.BeforeAction = False Then
                        Me.BrowseFileDialog()
                    End If

                    If pVal.ItemUID = "14" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("18").Specific.Value <> "" Then
                            Me.UpLoadExcel(objForm.UniqueID)
                        Else
                            objMain.objApplication.StatusBar.SetText("Excel Sheet Name should not be blank")
                        End If
                    End If

                    If pVal.ItemUID = "20" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.PostGRPO(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "22" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.ConsumeDiesel(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_VCD" Or oCFL.UniqueID = "CFL_VNM" Then
                            oDBs_Head.SetValue("U_VSPVCD", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_VSPVNAM", oDBs_Head.Offset, oDT.GetValue("CardName", 0))
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

    Sub UpLoadExcel(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD_C0")
         
            objMatrix = objForm.Items.Item("11").Specific

            Dim GetPswd As String = ""

            If objMain.IsSAPHANA = True Then
                GetPswd = "Select ""U_VSPPSWD"" From ""@VSP_FLT_CNFGSRN"" "
            Else
                GetPswd = "Select ""U_VSPPSWD"" From [@VSP_FLT_CNFGSRN] "
            End If



            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(GetPswd)

            If Me.ImportData(objForm.UniqueID, objMain.objCompany.Server, objMain.objCompany.CompanyDB, objMain.objCompany.DbUserName, _
                          oRs.Fields.Item(0).Value, objForm.Items.Item("16").Specific.Value.Trim, _
                          objForm.Items.Item("18").Specific.Value.Trim) = True Then
                objMain.objApplication.StatusBar.SetText("Wait.... Form is Uploading", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim GetExcelDtails As String = ""

                If objMain.IsSAPHANA = True Then
                    GetExcelDtails = "Select *  From " & """" & objForm.Items.Item("18").Specific.Value & """" & " Where ""ID"" <> ''"
                Else
                    GetExcelDtails = "Select *  From " & "[" & objForm.Items.Item("18").Specific.Value & "]" & " Where ID <> ''"
                End If


                Dim oRsGetExcelDtails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetExcelDtails.DoQuery(GetExcelDtails)

                If objMatrix.VisualRowCount = 1 Then
                    If objMatrix.Columns.Item("V_0").Cells.Item(1).Specific.Value = "" Then
                        objMatrix.Clear()
                        oDBs_Details.Clear()
                        objMatrix.FlushToDataSource()
                    End If
                End If

                For i As Integer = 1 To oRsGetExcelDtails.RecordCount
                    objMatrix.AddRow()

                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_VSPID", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(0).Value)
                    oDBs_Details.SetValue("U_VSPTDT", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(1).Value)
                    oDBs_Details.SetValue("U_VSPDLN", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(2).Value)
                    oDBs_Details.SetValue("U_VSPLCT", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(3).Value)
                    oDBs_Details.SetValue("U_VSPCID", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(4).Value)
                    oDBs_Details.SetValue("U_VSPQTY", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(5).Value)
                    oDBs_Details.SetValue("U_VSPVNO", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(6).Value)
                    oDBs_Details.SetValue("U_VSPCR", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(7).Value)
                    oDBs_Details.SetValue("U_VSPAMT", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(8).Value)
                    oDBs_Details.SetValue("U_VSPBAL", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(9).Value)
                    oDBs_Details.SetValue("U_VSPEXPT", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(10).Value)
                    oDBs_Details.SetValue("U_VSPGRPO", oDBs_Details.Offset, "")

                    Dim TrpShtNo As String = ""
                    If objMain.IsSAPHANA = True Then
                        TrpShtNo = "Select ""DocEntry"" From ""@VSP_FLT_TRSHT"" Where ""U_VSPSTS"" = 'Open' And ""U_VSPVHCL"" = '" & oRsGetExcelDtails.Fields.Item(6).Value & "'"
                    Else
                        TrpShtNo = "Select DocEntry From [@VSP_FLT_TRSHT] Where U_VSPSTS = 'Open' And U_VSPVHCL = '" & oRsGetExcelDtails.Fields.Item(6).Value & "'"
                    End If
                    Dim oRsTrpShtNo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsTrpShtNo.DoQuery(TrpShtNo)

                    oDBs_Details.SetValue("U_VSPTRNO", oDBs_Details.Offset, oRsTrpShtNo.Fields.Item(0).Value)
                    oDBs_Details.SetValue("U_VSPISSNO", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPCC1", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(11).Value)
                    oDBs_Details.SetValue("U_VSPCC2", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(12).Value)
                    oDBs_Details.SetValue("U_VSPCC3", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(13).Value)
                    oDBs_Details.SetValue("U_VSPCC4", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(14).Value)
                    oDBs_Details.SetValue("U_VSPCC5", oDBs_Details.Offset, oRsGetExcelDtails.Fields.Item(15).Value)

                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    oRsGetExcelDtails.MoveNext()
                Next
            End If
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
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

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Sub PostGRPO(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("11").Specific

            'If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            Dim FuelUploadDtls As String = ""

            If objMain.IsSAPHANA = True Then
                FuelUploadDtls = "Select ""LineId"" , ""U_VSPID"" , ""U_VSPTDT"" , ""U_VSPDLN"" , ""U_VSPLCT"" , ""U_VSPCID"" , ""U_VSPQTY"" , " & _
           """U_VSPVNO"", ""U_VSPCR"", ""U_VSPAMT"", ""U_VSPBAL"", ""U_VSPEXPT"", ""U_VSPCC1"", ""U_VSPCC2"", ""U_VSPCC3"", ""U_VSPCC4"", ""U_VSPCC5"" , ""U_VSPGRPO"" , ""U_VSPTRNO"" " & _
           "From ""@VSP_FLT_FLUPLD_C0"" Where ""DocEntry""  = '" & objForm.Items.Item("8").Specific.Value & "' And ""U_VSPGRPO"" = '' "
            Else
                FuelUploadDtls = "Select LineId , U_VSPID , U_VSPTDT , U_VSPDLN , U_VSPLCT , U_VSPCID , U_VSPQTY , " & _
           "U_VSPVNO, U_VSPCR, U_VSPAMT, U_VSPBAL, U_VSPEXPT, U_VSPCC1, U_VSPCC2, U_VSPCC3, U_VSPCC4, U_VSPCC5 , U_VSPGRPO , U_VSPTRNO " & _
           "From [@VSP_FLT_FLUPLD_C0] Where DocEntry  = '" & objForm.Items.Item("8").Specific.Value & "' And U_VSPGRPO = '' "
            End If


            Dim oRsFuelUploadDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsFuelUploadDtls.DoQuery(FuelUploadDtls)



            Dim GetCongigDtls As String = ""
            If objMain.IsSAPHANA = True Then
                GetCongigDtls = "Select ""U_VSPDEWHS"" , ""U_VSPTXCD"" From ""@VSP_FLT_CNFGSRN"" "
            Else
                GetCongigDtls = "Select U_VSPDEWHS , U_VSPTXCD From [@VSP_FLT_CNFGSRN] "
            End If

            Dim oRsGetCongigDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCongigDtls.DoQuery(GetCongigDtls)

            If oRsFuelUploadDtls.RecordCount > 0 Then

                For i As Integer = 1 To oRsFuelUploadDtls.RecordCount

                    Try
                        Dim oGRPO As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

                        oGRPO.CardCode = objForm.Items.Item("4").Specific.Value
                        oGRPO.DocDate = oRsFuelUploadDtls.Fields.Item("U_VSPTDT").Value
                        oGRPO.DocDueDate = oRsFuelUploadDtls.Fields.Item("U_VSPTDT").Value
                        oGRPO.TaxDate = oRsFuelUploadDtls.Fields.Item("U_VSPTDT").Value

                        oGRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items

                        Dim GetItemCode As String = ""

                        If objMain.IsSAPHANA = True Then
                            GetItemCode = "Select ""U_VSPDIITE"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & oRsFuelUploadDtls.Fields.Item("U_VSPVNO").Value & "'"
                        Else
                            GetItemCode = "Select U_VSPDIITE From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & oRsFuelUploadDtls.Fields.Item("U_VSPVNO").Value & "'"
                        End If
                        Dim oRsGetItemCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetItemCode.DoQuery(GetItemCode)

                        oGRPO.Lines.ItemCode = oRsGetItemCode.Fields.Item("U_VSPDIITE").Value

                        Dim GetInStock As String = ""

                        If objMain.IsSAPHANA = True Then
                            GetInStock = "Select ""OnHand"" From ""OITW"" Where ""WhsCode"" = '" & oRsGetCongigDtls.Fields.Item("U_VSPDEWHS").Value & "' " & _
                                                   "And ""ItemCode"" = '" & oRsGetItemCode.Fields.Item("U_VSPDIITE").Value & "' "

                        Else
                            GetInStock = "Select OnHand From OITW Where WhsCode = '" & oRsGetCongigDtls.Fields.Item("U_VSPDEWHS").Value & "' " & _
                                                   "And ItemCode = '" & oRsGetItemCode.Fields.Item("U_VSPDIITE").Value & "' "

                        End If
                        Dim oRsGetInStock As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetInStock.DoQuery(GetInStock)

                        Dim GetMaxCapacity As String = ""

                        If objMain.IsSAPHANA = True Then
                            GetMaxCapacity = "Select ""MaxLevel"" From ""OITM"" Where ""ItemCode"" = '" & oRsGetItemCode.Fields.Item("U_VSPDIITE").Value & "'"
                        Else
                            GetMaxCapacity = "Select MaxLevel From OITM Where ItemCode = '" & oRsGetItemCode.Fields.Item("U_VSPDIITE").Value & "'"
                        End If
                        Dim oRsGetMaxCapacity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetMaxCapacity.DoQuery(GetMaxCapacity)

                        Dim CalculateQty As Double = oRsGetInStock.Fields.Item("OnHand").Value + oRsFuelUploadDtls.Fields.Item("U_VSPQTY").Value

                        If oRsGetMaxCapacity.Fields.Item("MaxLevel").Value > CalculateQty Then
                            oGRPO.Lines.Quantity = oRsFuelUploadDtls.Fields.Item("U_VSPQTY").Value
                        Else
                            objMain.objApplication.StatusBar.SetText("Sum of Quantity & InStockQty of the Item exceeds Maximum Diesel Tank Capacity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Try
                        End If

                        oGRPO.Lines.LineTotal = oRsFuelUploadDtls.Fields.Item("U_VSPAMT").Value
                        Dim UnitPrice As Double = oRsFuelUploadDtls.Fields.Item("U_VSPAMT").Value / oRsFuelUploadDtls.Fields.Item("U_VSPQTY").Value
                        oGRPO.Lines.UnitPrice = UnitPrice
                        oGRPO.Lines.TaxCode = oRsGetCongigDtls.Fields.Item("U_VSPTXCD").Value
                        oGRPO.Lines.WarehouseCode = oRsGetCongigDtls.Fields.Item("U_VSPDEWHS").Value

                        If oRsFuelUploadDtls.Fields.Item("U_VSPCC1").Value <> "" Then
                            oGRPO.Lines.CostingCode = oRsFuelUploadDtls.Fields.Item("U_VSPCC1").Value
                        End If
                        If oRsFuelUploadDtls.Fields.Item("U_VSPCC2").Value <> "" Then
                            oGRPO.Lines.CostingCode2 = oRsFuelUploadDtls.Fields.Item("U_VSPCC2").Value
                        End If
                        If oRsFuelUploadDtls.Fields.Item("U_VSPCC3").Value <> "" Then
                            oGRPO.Lines.CostingCode3 = oRsFuelUploadDtls.Fields.Item("U_VSPCC3").Value
                        End If
                        If oRsFuelUploadDtls.Fields.Item("U_VSPCC4").Value <> "" Then
                            oGRPO.Lines.CostingCode4 = oRsFuelUploadDtls.Fields.Item("U_VSPCC4").Value
                        End If
                        If oRsFuelUploadDtls.Fields.Item("U_VSPCC5").Value <> "" Then
                            oGRPO.Lines.CostingCode5 = oRsFuelUploadDtls.Fields.Item("U_VSPCC5").Value
                        End If

                        oGRPO.UserFields.Fields.Item("U_VSPID").Value = oRsFuelUploadDtls.Fields.Item("U_VSPID").Value
                        oGRPO.UserFields.Fields.Item("U_VSPDNM").Value = oRsFuelUploadDtls.Fields.Item("U_VSPDLN").Value
                        oGRPO.UserFields.Fields.Item("U_VSPCSTID").Value = oRsFuelUploadDtls.Fields.Item("U_VSPCID").Value
                        oGRPO.UserFields.Fields.Item("U_VSPLCTN").Value = oRsFuelUploadDtls.Fields.Item("U_VSPLCT").Value
                        oGRPO.UserFields.Fields.Item("U_VSPVCHNO").Value = oRsFuelUploadDtls.Fields.Item("U_VSPVNO").Value
                        oGRPO.UserFields.Fields.Item("U_VSPCUR").Value = oRsFuelUploadDtls.Fields.Item("U_VSPCR").Value
                        oGRPO.UserFields.Fields.Item("U_VSPBAL").Value = oRsFuelUploadDtls.Fields.Item("U_VSPBAL").Value
                        oGRPO.UserFields.Fields.Item("U_VSPEXPT").Value = oRsFuelUploadDtls.Fields.Item("U_VSPEXPT").Value

                        If oGRPO.Add = 0 Then

                            Dim GetDocEntry As String = ""
                            If objMain.IsSAPHANA = True Then
                                GetDocEntry = "Select Max(""DocEntry"") From ""OPDN"""
                            Else
                                GetDocEntry = "Select Max(DocEntry) From OPDN"
                            End If
                            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetDocEntry.DoQuery(GetDocEntry)
                            Dim DocNum As Integer = oRsGetDocEntry.Fields.Item(0).Value
                            Dim Line As Integer = oRsFuelUploadDtls.Fields.Item("LineId").Value
                            objMain.sCmp = objMain.objCompany.GetCompanyService
                            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OFLUPLD")
                            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                            objMain.oGeneralParams.SetProperty("DocEntry", objForm.Items.Item("8").Specific.Value)
                            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                            objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_FLUPLD_C0")
                            objMain.oChildren.Item(Line - 1).SetProperty("U_VSPGRPO", DocNum.ToString(""))
                            objMain.oGeneralService.Update(objMain.oGeneralData)


                            If oRsFuelUploadDtls.Fields.Item("U_VSPTRNO").Value <> "" Then
                                Me.UpdateGRPOInOpenTrpSht(objForm.UniqueID, oRsFuelUploadDtls.Fields.Item("U_VSPTRNO").Value, _
                                                          oRsGetItemCode.Fields.Item("U_VSPDIITE").Value, oRsGetCongigDtls.Fields.Item("U_VSPTXCD").Value, _
                                                          oRsFuelUploadDtls.Fields.Item("U_VSPTDT").Value, oRsFuelUploadDtls.Fields.Item("U_VSPAMT").Value, _
                                                          oRsFuelUploadDtls.Fields.Item("U_VSPQTY").Value, DocNum.ToString(""), oRsFuelUploadDtls.Fields.Item("U_VSPCC1").Value, _
                                                          oRsFuelUploadDtls.Fields.Item("U_VSPCC2").Value, oRsFuelUploadDtls.Fields.Item("U_VSPCC3").Value, _
                                                          oRsFuelUploadDtls.Fields.Item("U_VSPCC4").Value, oRsFuelUploadDtls.Fields.Item("U_VSPCC5").Value)
                            End If

                        Else
                            objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oGRPO)
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
                    oRsFuelUploadDtls.MoveNext()
                Next

                objForm.Freeze(True)
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD_C0")

                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))
                objMatrix = objForm.Items.Item("11").Specific
                objMatrix.LoadFromDataSource()
                objMatrix.AutoResizeColumns()
                objForm.Refresh()

                objForm.Freeze(False)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpdateGRPOInOpenTrpSht(ByVal FormUID As String, ByVal TripShtDocNum As Integer, ByVal ItemCode As String, _
                                ByVal TaxCode As String, ByVal DocDate As String, ByVal LineTotal As Double, _
                                ByVal Quantity As Double, ByVal GRPONUM As String, ByVal CC1 As String, ByVal CC2 As String, _
                                ByVal CC3 As String, ByVal CC4 As String, ByVal CC5 As String)
        Try
            Dim GetCount As String = ""
            If objMain.IsSAPHANA = True Then
                GetCount = "Select Max(""LineId"") From  ""@VSP_FLT_TRSHT_C5"" Where ""DocEntry"" = '" & TripShtDocNum & "' "
            Else
                GetCount = "Select Max(LineId) From  [@VSP_FLT_TRSHT_C5] Where DocEntry = '" & TripShtDocNum & "' "
            End If
            Dim oRsGetCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCount.DoQuery(GetCount)

            Dim oRsItemName As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If objMain.IsSAPHANA = True Then
                oRsItemName.DoQuery("Select ""ItemName"" From OITM Where ""ItemCode"" = '" & ItemCode & "' ")
            Else
                oRsItemName.DoQuery("Select ItemName From OITM Where ItemCode = '" & ItemCode & "' ")
            End If


            Dim Rate As Double = (LineTotal / Quantity)
            Dim LineId As Integer = CInt(oRsGetCount.Fields.Item(0).Value)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", TripShtDocNum.ToString(""))
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

            objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C5")
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPVENCO", objForm.Items.Item("4").Specific.Value)
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPDATE", DocDate)
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPRATE", Rate.ToString)
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPQUAN", Quantity)
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPAMT", LineTotal)
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPDCNUM", GRPONUM)
            objMain.oChildren.Add()

            objMain.oGeneralService.Update(objMain.oGeneralData)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLFilterVendors(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "CardType"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "S"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

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
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")
        If MyProcs.Length <> 0 Then
            For i As Integer = 0 To MyProcs.Length - 1

                Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                MyTest1.FileName = "Select the Reference Document"
                'Windows XP
                If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                    Try

                        oDBs_Head.SetValue("U_VSPATCH", oDBs_Head.Offset, MyTest1.FileName)
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    Catch ex As IO.IOException
                        objMain.objApplication.MessageBox(ex.Message)
                        Exit Sub
                    End Try
                    'Windows 7
                    'AttachMentPath +
                ElseIf MyTest1.ShowDialog() = DialogResult.OK Then
                    Try

                        oDBs_Head.SetValue("U_VSPATCH", oDBs_Head.Offset, MyTest1.FileName)
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
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

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        objMatrix = objForm.Items.Item("11").Specific

                        If eventInfo.ItemUID = "11" And eventInfo.ColUID = "V_-1" And objMatrix.RowCount >= 1 Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "11" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "11" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                Else
                    Try
                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ConsumeDiesel(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")

            Dim GetToBePosted As String = ""

            If objMain.IsSAPHANA = True Then
                GetToBePosted = "Select ""LineId"" -1 , ""U_VSPTDT"" , ""U_VSPGRPO"" , ""U_VSPTRNO"" From ""@VSP_FLT_FLUPLD_C0"" B Where B.""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And " & _
           """U_VSPGRPO"" <> '' And (""U_VSPISSNO"" IS NULL Or ""U_VSPISSNO"" = '')"
            Else
                GetToBePosted = "Select LineId -1 , U_VSPTDT , U_VSPGRPO , U_VSPTRNO From [@VSP_FLT_FLUPLD_C0] B Where B.DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And " & _
           "U_VSPGRPO <> '' And (U_VSPISSNO IS NULL Or U_VSPISSNO = '')"
            End If
            Dim oRsGetToBePosted As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetToBePosted.DoQuery(GetToBePosted)

            If oRsGetToBePosted.RecordCount > 0 Then
                oRsGetToBePosted.MoveFirst()

                For i As Integer = 1 To oRsGetToBePosted.RecordCount

                    Dim GetGRPODetails As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetGRPODetails = "Select ""ItemCode"" , ""Quantity"" , ""WhsCode"" , ""OcrCode"" , ""OcrCode2"" , ""OcrCode3"" , ""OcrCode4"" , ""OcrCode5"" From ""PDN1"" Where " & _
                   """DocEntry"" = '" & oRsGetToBePosted.Fields.Item("U_VSPGRPO").Value & "'"
                    Else
                        GetGRPODetails = "Select ItemCode , Quantity , WhsCode , OcrCode , OcrCode2 , OcrCode3 , OcrCode4 , OcrCode5 From PDN1 Where " & _
                   "DocEntry = '" & oRsGetToBePosted.Fields.Item("U_VSPGRPO").Value & "'"
                    End If
                    Dim oRsGetGRPODetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetGRPODetails.DoQuery(GetGRPODetails)

                    Dim oIssue As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    oIssue.DocDate = oRsGetToBePosted.Fields.Item("U_VSPTDT").Value
                    oIssue.UserFields.Fields.Item("U_VSPTRSHP").Value = oRsGetToBePosted.Fields.Item("U_VSPTRNO").Value.ToString
                    oIssue.Lines.ItemCode = oRsGetGRPODetails.Fields.Item("ItemCode").Value
                    oIssue.Lines.Quantity = oRsGetGRPODetails.Fields.Item("Quantity").Value
                    oIssue.Lines.WarehouseCode = oRsGetGRPODetails.Fields.Item("WhsCode").Value

                    If oRsGetGRPODetails.Fields.Item("OcrCode").Value <> "" Then
                        oIssue.Lines.CostingCode = oRsGetGRPODetails.Fields.Item("OcrCode").Value
                    End If

                    If oRsGetGRPODetails.Fields.Item("OcrCode2").Value <> "" Then
                        oIssue.Lines.CostingCode2 = oRsGetGRPODetails.Fields.Item("OcrCode2").Value
                    End If

                    If oRsGetGRPODetails.Fields.Item("OcrCode3").Value <> "" Then
                        oIssue.Lines.CostingCode3 = oRsGetGRPODetails.Fields.Item("OcrCode3").Value
                    End If

                    If oRsGetGRPODetails.Fields.Item("OcrCode4").Value <> "" Then
                        oIssue.Lines.CostingCode4 = oRsGetGRPODetails.Fields.Item("OcrCode4").Value
                    End If

                    If oRsGetGRPODetails.Fields.Item("OcrCode5").Value <> "" Then
                        oIssue.Lines.CostingCode5 = oRsGetGRPODetails.Fields.Item("OcrCode5").Value
                    End If


                    If oIssue.Add = 0 Then
                        Dim GetDocEntry As String = ""
                        If objMain.IsSAPHANA = True Then
                            GetDocEntry = "Select Max(""DocEntry"") From ""OIGE"""
                        Else
                            GetDocEntry = "Select Max(DocEntry) From OIGE"
                        End If
                        Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDocEntry.DoQuery(GetDocEntry)

                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OFLUPLD")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_FLUPLD_C0")
                        objMain.oChildren.Item(CInt(oRsGetToBePosted.Fields.Item(0).Value)).SetProperty("U_VSPISSNO", oRsGetDocEntry.Fields.Item(0).Value.ToString)
                        objMain.oGeneralService.Update(objMain.oGeneralData)
                    Else
                        objMain.objApplication.StatusBar.SetText("LineId : " & oRsGetToBePosted.Fields.Item(0).Value & ", " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If

                    oRsGetToBePosted.MoveNext()
                Next
            End If

            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD_C0")

            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_FLUPLD_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))
            objMatrix = objForm.Items.Item("11").Specific
            objMatrix.LoadFromDataSource()
            objMatrix.AutoResizeColumns()
            objForm.Refresh()

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
