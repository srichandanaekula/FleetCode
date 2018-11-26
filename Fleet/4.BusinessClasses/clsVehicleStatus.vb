Public Class clsVehicleStatus

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1 As SAPbouiCOM.Matrix
    Dim oGrid As SAPbouiCOM.Grid
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDT, oDT1 As SAPbouiCOM.DataTable
    Dim Path As String
    Dim GotDateValue, LostDateValue As String
    Dim count As Integer = 0
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("VehicleStatus.xml", "VSP_FLT_VCHSTS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VCHSTS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            ' objForm.Freeze(True)

            objMatrix = objForm.Items.Item("6").Specific

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000005").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000005").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000007").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000007").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000003").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000003").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("1000003").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000006").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000006").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
        
            objMain.objApplication.StatusBar.SetText("Please wait while vehicle status is loading", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If objMain.IsSAPHANA = True Then
                objMain.objUtilities.MatrixComboBoxValues(objMatrix.Columns.Item("V_12"), "Select ""U_VSPFPLC"" , ""U_VSPFPLC"" From ""@VSP_FLT_RTMSTR_C1"" T0 where ""U_VSPFPLC"" <> '' Group By ""U_VSPFPLC"" Order By T0.""U_VSPFPLC""")
                objMain.objUtilities.MatrixComboBoxValues(objMatrix.Columns.Item("V_11"), "Select ""U_VSPFPLC"" , ""U_VSPFPLC"" From ""@VSP_FLT_RTMSTR_C1"" T0 where ""U_VSPFPLC"" <> '' Group By ""U_VSPFPLC"" Order By T0.""U_VSPFPLC""")
            Else
                objMain.objUtilities.MatrixComboBoxValues(objMatrix.Columns.Item("V_12"), "Select U_VSPFPLC , U_VSPFPLC From [@VSP_FLT_RTMSTR_C1] T0 where U_VSPFPLC <> '' Group By U_VSPFPLC Order By T0.U_VSPFPLC")
                objMain.objUtilities.MatrixComboBoxValues(objMatrix.Columns.Item("V_11"), "Select U_VSPFPLC , U_VSPFPLC From [@VSP_FLT_RTMSTR_C1] T0 where U_VSPFPLC <> '' Group By U_VSPFPLC Order By T0.U_VSPFPLC")
            End If
           

            objMatrix.Columns.Item("V_4").ValidValues.Add("", "")
            objMatrix.Columns.Item("V_4").ValidValues.Add("Load", "Load")
            objMatrix.Columns.Item("V_4").ValidValues.Add("UnLoad", "UnLoad")
            objMatrix.Columns.Item("V_4").ValidValues.Add("None", "None")
            objMatrix.Columns.Item("V_4").ValidValues.Add("Empty", "Empty")

            objComboBox = objForm.Items.Item("1000006").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")

            'objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OVEHST"))
            oDBs_Head.SetValue("U_VSPDT", oDBs_Head.Offset, Today.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_VSPSTS", oDBs_Head.Offset, "Open")

            objMatrix = objForm.Items.Item("6").Specific

            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "6")
            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String, Optional ByVal OpenKM As String = "0", Optional ByVal Source As String = "")

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

            objMatrix = objForm.Items.Item("6").Specific
            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, OpenKM)
            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVAL As SAPbouiCOM.MenuEvent, ByRef BubleEvent As Boolean)
        Try
            If pVAL.MenuUID = "VSP_FLT_VECHSTS" And pVAL.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVAL.MenuUID = "1282" And pVAL.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VCHSTS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                Me.SetDefault(objForm.UniqueID)
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
                    objMatrix = objForm.Items.Item("6").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                                          pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1000007" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If objForm.Items.Item("23").Specific.Value = "" Then

                            Me.UpadateclosekmtoVMwithoutTripSht(objForm.UniqueID)
                        End If
                    End If

                    If pVal.ItemUID = "1000005" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.UpadateOdmtrRdngsFromVMWithTripSht(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "21" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.UpdateVehStatusInTrpSht(objForm.UniqueID, objForm.Items.Item("9").Specific.Value, _
                                                                                  objForm.Items.Item("1000003").Specific.Value, objForm.Items.Item("20").Specific.Value, _
                                                                                  objForm.Items.Item("10").Specific.Value, objForm.Items.Item("13").Specific.Value, _
                                                                                  objForm.Items.Item("8").Specific.Value, objForm.Items.Item("23").Specific.Value)
                    End If

                    If pVal.ItemUID = "15" And pVal.BeforeAction = False Then
                        If objMain.objApplication.MessageBox("Do you want to Export to PDF", 2, "Ok", "Cancel") = 1 Then
                            Dim VehicleStatus As New VehicleStatus
                            VehicleStatus.VSPDT1 = objForm.Items.Item("9").Specific.Value
                            VehicleStatus.PrintPDF = "Yes"
                            VehicleStatus.ShowDialog()
                        Else
                            Dim RptVehicleStatus As New VehicleStatus
                            RptVehicleStatus.VSPDT1 = objForm.Items.Item("9").Specific.Value
                            RptVehicleStatus.ShowDialog()
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

                    objMatrix = objForm.Items.Item("6").Specific


                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "CFL_TS" Then
                            If objForm.Items.Item("1000003").Specific.Value <> "" Then
                                Me.CFLFilterTripSheet(objForm.UniqueID, "CFL_TS")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_SONO" Then

                            Me.CFLFilterSO(objForm.UniqueID, oCFL.UniqueID)

                        End If

                        If oCFL.UniqueID = "CFL_VNO" Then
                            If objForm.Items.Item("9").Specific.Value = "" Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Please Enter Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                Me.CFLFilterForVehicles(objForm.UniqueID, oCFL.UniqueID)
                            End If
                        End If
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_SONO" Then
                            oDBs_Details.SetValue("U_VSPSLNO", oDBs_Details.Offset, oDT.GetValue("DocEntry", 0))

                            Dim GetSalesOrderDetals As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetSalesOrderDetals = "Select T1.""ItemCode"",T1.""Quantity"" From ""RDR1"" T1  where T1.""DocEntry""= '" & oDT.GetValue("DocEntry", 0) & "'"
                            Else
                                GetSalesOrderDetals = "Select T1.""ItemCode"",T1.""Quantity"" From ""RDR1"" T1  where T1.""DocEntry""= '" & oDT.GetValue("DocEntry", 0) & "'"
                            End If
                            Dim oRsGetSalesOrderDetals As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetSalesOrderDetals.DoQuery(GetSalesOrderDetals)



                            'If oRsGetSalesOrderDetals.RecordCount > 0 Then
                            '    oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, oRsGetSalesOrderDetals.Fields.Item(0).Value.ToString)
                            '    oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, CDbl(oRsGetSalesOrderDetals.Fields.Item(1).Value.ToString))
                            'End If

                            ' objMatrix.AddRow()
                            'oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                            'oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            ''oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                            ''oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                            ''oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                            'oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, objMatrix.Columns.Item("").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                            'oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                            'objMatrix.SetLineData(objMatrix.VisualRowCount)

                            'Else
                            If oRsGetSalesOrderDetals.RecordCount > 0 Then
                                oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, oRsGetSalesOrderDetals.Fields.Item(0).Value.ToString)
                                oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, CDbl(oRsGetSalesOrderDetals.Fields.Item(1).Value.ToString))
                            End If

                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                            If objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Checked = True Then
                                oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "Y")
                            Else
                                oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                            End If

                            If objMatrix.Columns.Item("V_13").Cells.Item(pVal.Row).Specific.Checked = True Then
                                oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "Y")
                            Else
                                oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                            End If
                            objMatrix.SetLineData(pVal.Row)


                            'End If



                        End If

                        If oCFL.UniqueID = "CFL_VNO" Then
                            Dim B As Double = oDT.GetValue("U_VSPODRDG", 0)

                            Me.LoadVehicleDetails(objForm.UniqueID, oDT.GetValue("U_VSPVNO", 0), oDT.GetValue("U_VSPCNTR", 0), _
                                                  oDT.GetValue("U_VSPCNTNM", 0), oDT.GetValue("U_VSPODRDG", 0), pVal.Row)

                            Dim GetDocentry As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetDocentry = "Select MAX(""DocEntry"")  From ""@VSP_VECHSTS"" where ""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                            Else
                                GetDocentry = "Select MAX(DocEntry)  From [@VSP_VECHSTS] where U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                            End If
                            Dim oRsGetDocentry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetDocentry.DoQuery(GetDocentry)

                            Dim GetLoad As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetLoad = "Select T0.""LineId"",T0.""U_VSPQUA"" ,T0.""U_VSPSOURC"" ,""U_VSPDST"" ,""U_VSPCHK""  from ""@VSP_VECHSTS_C0"" T0 inner join ""@VSP_VECHSTS"" T1 on T0.""DocEntry"" = T1.""DocEntry""  " & _
                                                   "where  T0.""DocEntry"" = '" & oRsGetDocentry.Fields.Item(0).Value & "' and  T0.""U_VSPTY"" = 'Load' and T0.""LineId"" = (Select MAX(T0.""LineId"" ) from ""@VSP_VECHSTS_C0"" T0 inner join ""@VSP_VECHSTS"" T1 on " & _
                                                   "T0.""DocEntry"" = T1.""DocEntry""  where  T0.""DocEntry"" = '" & oRsGetDocentry.Fields.Item(0).Value & "' and  T0.""U_VSPTY"" = 'Load')"
                            Else
                                GetLoad = "Select T0.LineId,T0.U_VSPQUA ,T0.U_VSPSOURC ,U_VSPDST ,U_VSPCHK  from [@VSP_VECHSTS_C0] T0 inner join [@VSP_VECHSTS] T1 on T0.DocEntry = T1.DocEntry  " & _
                                                   "where  T0.DocEntry = '" & oRsGetDocentry.Fields.Item(0).Value & "' and  T0.U_VSPTY = 'Load' and T0.LineId = (Select MAX(T0.LineId ) from [@VSP_VECHSTS_C0] T0 inner join [@VSP_VECHSTS] T1 on " & _
                                                   "T0.DocEntry = T1.DocEntry  where  T0.DocEntry = '" & oRsGetDocentry.Fields.Item(0).Value & "' and  T0.U_VSPTY = 'Load')"
                            End If
                            Dim oRsGetLoad As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetLoad.DoQuery(GetLoad)

                            Dim GetUnload As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetUnload = "Select Sum(""U_VSPQUA"") from ""@VSP_VECHSTS_C0"" where ""DocEntry"" = '" & oRsGetDocentry.Fields.Item(0).Value & "' and ""LineId"" > '" & oRsGetLoad.Fields.Item(0).Value & "' and ""U_VSPTY"" = 'UnLoad'"
                            Else
                                GetUnload = "Select Sum(U_VSPQUA) from [@VSP_VECHSTS_C0] where DocEntry = '" & oRsGetDocentry.Fields.Item(0).Value & "' and LineId > '" & oRsGetLoad.Fields.Item(0).Value & "' and U_VSPTY = 'UnLoad'"
                            End If
                            Dim oRsGetUnload As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetUnload.DoQuery(GetUnload)

                            Dim Load As Double = CDbl(oRsGetLoad.Fields.Item(1).Value)
                            Dim Unload As Double = CDbl(oRsGetUnload.Fields.Item(0).Value)
                            Dim Check As Double = Load - Unload

                            Dim GetVehicleSts As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetVehicleSts = "Select ""U_VSPSTS"" From ""@VSP_VECHSTS"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "' And  ""U_VSPSTS"" ='Close' " & _
                            "And ""DocEntry"" = (Select MAX (""DocEntry"")From ""@VSP_VECHSTS"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "' And  ""U_VSPSTS"" ='Close')"
                            Else
                                GetVehicleSts = "Select U_VSPSTS From [@VSP_VECHSTS] Where U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "' And  U_VSPSTS ='Close' " & _
                            "And DocEntry = (Select MAX (DocEntry)From [@VSP_VECHSTS] Where U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "' And  U_VSPSTS ='Close')"
                            End If
                            Dim oRsGetVehicleSts As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetVehicleSts.DoQuery(GetVehicleSts)

                            If oRsGetVehicleSts.RecordCount = 0 Then
                                If Check > 0 Then
                                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                                    oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "Load")
                                    oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, oRsGetLoad.Fields.Item("U_VSPSOURC").Value)
                                    oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, oRsGetLoad.Fields.Item("U_VSPDST").Value)
                                    oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, CDbl(Check))
                                    'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                                    oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, oRsGetLoad.Fields.Item("U_VSPCHK").Value)
                                    oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                                    objMatrix.SetLineData(objMatrix.VisualRowCount)

                                    Me.SetNewLine(objForm.UniqueID, "6")
                                    Me.SetCellsEditable(objForm.UniqueID)
                                End If

                            ElseIf Check > 0 Then
                                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                                oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "Load")
                                oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, CDbl(Check))
                                'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, oRsGetLoad.Fields.Item("U_VSPCHK").Value)
                                oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                                objMatrix.SetLineData(objMatrix.VisualRowCount)

                                Me.SetNewLine(objForm.UniqueID, "6", oDT.GetValue("U_VSPODRDG", 0))
                                Me.SetCellsEditable(objForm.UniqueID)
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_CHITEM" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, objMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                            objMatrix.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_VEN" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                            objMatrix.SetLineData(pVal.Row)

                            'If pVal.Row = objMatrix.VisualRowCount Then
                            '    Me.SetNewLine(objForm.UniqueID, "6")
                            'End If
                        End If

                        If oCFL.UniqueID = "CFL_TS" Then

                            Dim Getclskm As String = ""

                            If objMain.IsSAPHANA = True Then
                                Getclskm = "Select IFNULL(T1.""U_VSPCLKM"",0) From ""@VSP_FLT_TRSHT"" T0 Inner Join ""@VSP_FLT_TRSHT_C1"" T1 " & _
                                                      "On T1.""DocEntry""= T0.""DocEntry"" Where T0.""DocEntry"" = '" & oDT.GetValue(0, 0) & "' and " & _
                                                      """LineId"" = (Select MAX(""LineId"" )-1 From ""@VSP_FLT_TRSHT"" T0 Inner Join ""@VSP_FLT_TRSHT_C1"" T1 On T1.""DocEntry""= T0.""DocEntry""  Where t0.""DocEntry""  = '" & oDT.GetValue(0, 0) & "')"

                            Else
                                Getclskm = "Select ISNULL(T1.U_VSPCLKM,0) From [@VSP_FLT_TRSHT] T0 Inner Join [@VSP_FLT_TRSHT_C1] T1 " & _
                                                      "On T1.DocEntry= T0.DocEntry Where T0.DocEntry = '" & oDT.GetValue(0, 0) & "' and " & _
                                                      "LineId = (Select MAX(LineId )-1 From [@VSP_FLT_TRSHT] T0 Inner Join [@VSP_FLT_TRSHT_C1] T1 On T1.DocEntry= T0.DocEntry  Where t0.DocEntry  = '" & oDT.GetValue(0, 0) & "')"

                            End If
                            Dim oRsGetclskm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetclskm.DoQuery(Getclskm)

                            'Dim Getsource As String = "Select  min(T1.U_VSPDEST)  From [@VSP_FLT_TRSHT] T0 Inner Join [@VSP_FLT_TRSHT_C1] T1 " & _
                            '                          "On T1.DocEntry= T0.DocEntry Where DocNum = '" & oDT.GetValue(0, 0) & "'"
                            'Dim oRsGetsource As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRsGetsource.DoQuery(Getsource)
                            Dim getde As String = oDT.GetValue("DocNum", 0)
                            'Dim getdn As String = oDT.GetValue("", 0)
                            oDBs_Head.SetValue("U_VSPTRSHT", oDBs_Head.Offset, oDT.GetValue(0, 0))
                            oDBs_Head.SetValue("U_VSPROUTE", oDBs_Head.Offset, oDT.GetValue("U_VSPROUTE", 0))
                            oDBs_Head.SetValue("U_VSPSTS", oDBs_Head.Offset, oDT.GetValue("U_VSPSTS", 0))

                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")

                            If oRsGetclskm.Fields.Item(0).Value > 0 Then
                                oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, oRsGetclskm.Fields.Item(0).Value)
                            Else
                                If objMatrix.VisualRowCount = 0 Then
                                    Dim GetOdmtrRdng As String = ""

                                    If objMain.IsSAPHANA = True Then
                                        GetOdmtrRdng = "Select ""U_VSPODRDG""  From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                                    Else
                                        GetOdmtrRdng = "Select U_VSPODRDG  From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                                    End If
                                    Dim oRsGetOdmtrRdng As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRsGetOdmtrRdng.DoQuery(GetOdmtrRdng)
                                    oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, oRsGetOdmtrRdng.Fields.Item(0).Value)
                                End If
                            End If

                            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")

                            If oRsGetclskm.Fields.Item(0).Value > 0 Then
                                oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPDST", oDBs_Head.Offset, "")
                            Else
                                oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPDST", oDBs_Head.Offset, "")
                            End If

                            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")

                            objMatrix.SetLineData(objMatrix.VisualRowCount)
                        End If

                        If oCFL.UniqueID = "CFL_RTCD" Then
                            '  Me.LoadRouteDetails(objForm.UniqueID, oDT.GetValue("U_VSPRCD", 0), oDT.GetValue("U_VSPSRCE", 0), oDT.GetValue("U_VSPDEST", 0), pVal.Row)
                        End If

                        If pVal.ItemUID = "6" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ColUID = "V_11" Then
                                If objMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                    If pVal.Row = objMatrix.VisualRowCount Then
                                        Me.SetNewLine(objForm.UniqueID, "6", objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value, _
                                                      objMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                                    Else
                                        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row + 1)
                                        oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "0")
                                        oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                                        'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                                        oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                                        oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                                        objMatrix.SetLineData(pVal.Row + 1)
                                    End If
                                End If
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("6").Specific

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

                    objMatrix = objForm.Items.Item("6").Specific

                    If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "1000003" Then

                            Dim ChkContarcorExist As String = ""

                            If objMain.IsSAPHANA = True Then
                                ChkContarcorExist = "Select ""DocNum"" From ""@VSP_VECHSTS"" Where ""U_VSPVNO"" ='" & objForm.Items.Item("1000003").Specific.Value & "' and ""U_VSPDT"" = '" & objForm.Items.Item("9").Specific.Value.Trim & "'"
                            Else
                                ChkContarcorExist = "Select DocNum From [@VSP_VECHSTS] Where [U_VSPVNO] ='" & objForm.Items.Item("1000003").Specific.Value & "' and [U_VSPDT] = '" & objForm.Items.Item("9").Specific.Value.Trim & "'"
                            End If
                            Dim oRsChkcontarcorExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkcontarcorExist.DoQuery(ChkContarcorExist)

                            If oRsChkcontarcorExist.RecordCount > 0 Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("8").Specific.value = oRsChkcontarcorExist.Fields.Item(0).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                                objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                                ' objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("9").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                                ' objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                objForm.Items.Item("23").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                objMain.objApplication.StatusBar.SetText("VehicleNo and TripSheetNo Already Existed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "6" And pVal.ColUID = "V_6" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim OpenKM As Double = 0
                        Dim CloseKM As Double = 0

                        If objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            OpenKM = CDbl(objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                        End If
                        If objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            CloseKM = CDbl(objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                        End If

                        If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If CloseKM < OpenKM Or CloseKM <= 0 Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Close KM's Cannot Be Less than Open KM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            ElseIf CloseKM - OpenKM > 800 Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Close KM's cannot be more than 800Km", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                                oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, objMatrix.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, CloseKM - OpenKM)
                                oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                                oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                                objMatrix.SetLineData(pVal.Row)
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "6" And pVal.ColUID = "V_8" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                        If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Selected.Value = "None" And _
                        CDbl(objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value) <> 0 Then

                            objMain.objApplication.StatusBar.SetText("Qunatity for None Cannot be greater or Less than Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objForm.Freeze(True)
                            objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = 0
                            objForm.Freeze(False)
                            BubbleEvent = False
                            Exit Try

                        ElseIf objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Selected.Value = "UnLoad" And _
                        CDbl(objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value) = 0 Then

                            objMain.objApplication.StatusBar.SetText("Qunatity for Unload Cannot be Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objForm.Freeze(True)
                            ' objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = 0
                            objForm.Freeze(False)
                            BubbleEvent = False
                            Exit Try
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")
                    objMatrix = objForm.Items.Item("6").Specific


                    If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "UnLoad" Or objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "None" And pVal.BeforeAction = False Then
                        objMatrix.CommonSetting.SetRowEditable(pVal.Row, True)
                        objMatrix.Columns.Item("V_2").Editable = False
                        objMatrix.Columns.Item("V_12").Editable = True
                        objMatrix.Columns.Item("V_11").Editable = True
                        objMatrix.Columns.Item("V_10").Editable = False
                        objMatrix.Columns.Item("V_13").Editable = False
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "6" And pVal.ColUID = "V_4" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        BubbleEvent = False

                        Dim Quantity As Double = 0
                        Dim value As Double = 0

                        If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "Load" Then
                            For i As Integer = objMatrix.VisualRowCount - 1 To 1 Step -1

                                If objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value = "Load" Then
                                    value = objMatrix.Columns.Item("V_8").Cells.Item(i).Specific.Value
                                End If

                                If objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value <> "" Then
                                    If objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value = "UnLoad" Then
                                        If objMatrix.Columns.Item("V_8").Cells.Item(i).Specific.Value <> 0 Then
                                            Quantity = Quantity + objMatrix.Columns.Item("V_8").Cells.Item(i).Specific.Value
                                        End If

                                        'ElseIf objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value <> "" Then
                                        '    If objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value = "None" Then
                                        '        'Quantity = Quantity + objMatrix.Columns.Item("V_8").Cells.Item(i).Specific.Value
                                        '    End If

                                    End If

                                    If objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value <> "None" Then

                                        If value = Quantity Then
                                            objForm.Freeze(False)
                                            Exit Sub
                                        Else
                                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                                            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                                            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPSLNO", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                                            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                                            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                                            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")

                                            BubbleEvent = False
                                            objMatrix.SetLineData(pVal.Row)
                                            Me.SetCellsEditable(objForm.UniqueID)
                                        End If
                                    End If
                                End If
                                objForm.Freeze(False)
                            Next

                            If objMatrix.VisualRowCount = 1 Then
                                If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "Load" Then
                                    BubbleEvent = False

                                    Exit Try
                                End If
                            End If
                            objMain.objApplication.StatusBar.SetText("Please UnLoad the Total Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)


                            If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "Load" And objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = 0 And pVal.BeforeAction = False Then
                                objMatrix.CommonSetting.SetRowEditable(pVal.Row, False)
                                objMatrix.Columns.Item("V_4").Editable = True

                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                End If

                                objForm.Freeze(False)
                            End If
                            objForm.Freeze(False)
                        End If
                        objForm.Freeze(False)
                    End If

                    ''Added on 29-10-18

                    'If pVal.ItemUID = "6" And pVal.ColUID = "V_4" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                    '    If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "Load" Then

                    '        Try
                    '            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
                    '            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

                    '            objMatrix = objForm.Items.Item("6").Specific

                    '            If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "Load" Then


                    '                Dim GettripSheetNo As String = ""
                    '                If objMain.IsSAPHANA = True Then
                    '                    GettripSheetNo = "Select ""DocNum""  From  ""@VSP_FLT_TRSHT""  Where ""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "' "
                    '                Else
                    '                    GettripSheetNo = "Select ""DocNum""  From  ""@VSP_FLT_TRSHT""  Where ""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "' "
                    '                End If

                    '                Dim oRsGettripSheetNo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                oRsGettripSheetNo.DoQuery(GettripSheetNo)


                    '                Dim GetSalesOrderno As String = ""
                    '                If objMain.IsSAPHANA = True Then
                    '                    GetSalesOrderno = "Select T1.""U_VSPDCNUM""  From  ""@VSP_FLT_TRSHT_C3"" T1  Where T1.""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "'"
                    '                Else
                    '                    GetSalesOrderno = "Select T1.""U_VSPDCNUM""  From  ""@VSP_FLT_TRSHT_C3"" T1  Where T1.""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "' "
                    '                End If

                    '                Dim oRsGetSalesOrderno As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                oRsGetSalesOrderno.DoQuery(GetSalesOrderno)
                    '                If oRsGetSalesOrderno.RecordCount > 0 Then
                    '                    Dim DocNum As String = oRsGetSalesOrderno.Fields.Item(0).Value
                    '                    If DocNum = "" Then
                    '                        Dim tripsheetno As String = oRsGettripSheetNo.Fields.Item(0).Value
                    '                        objMain.objApplication.StatusBar.SetText("Please Post Sales Order For Trip Sheet No " & tripsheetno, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                        count = 1
                    '                        Exit Try
                    '                    End If
                    '                    Dim GetSalesOrderDetals As String = ""

                    '                    If objMain.IsSAPHANA = True Then
                    '                        GetSalesOrderDetals = "Select T1.""ItemCode"",T1.""Quantity"" From ""RDR1"" T1  where T1.""DocEntry""= '" & oRsGetSalesOrderno.Fields.Item(0).Value & "'"
                    '                    Else
                    '                        GetSalesOrderDetals = "Select T1.""ItemCode"",T1.""Quantity"" From ""RDR1"" T1  where T1.""DocEntry""= '" & oRsGetSalesOrderno.Fields.Item(0).Value & "'"
                    '                    End If
                    '                    Dim oRsGetSalesOrderDetals As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                    oRsGetSalesOrderDetals.DoQuery(GetSalesOrderDetals)

                    '                    If oRsGetSalesOrderDetals.RecordCount > 0 Then
                    '                        oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, oRsGetSalesOrderDetals.Fields.Item(0).Value.ToString)
                    '                        oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, CDbl(oRsGetSalesOrderDetals.Fields.Item(1).Value.ToString))
                    '                    End If
                    '                    If objMatrix.VisualRowCount = pVal.Row Then
                    '                        ' objMatrix.AddRow()
                    '                        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    '                        oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                    '                        oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                    '                        'oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                    '                        'oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                    '                        'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                    '                        oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                    '                        oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                    '                        objMatrix.SetLineData(objMatrix.VisualRowCount)
                    '                    End If
                    '                End If
                    '            End If
                    '        Catch ex As Exception

                    '        End Try
                    '    End If
                    'End If

                Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "9" And pVal.BeforeAction = False Then
                        GotDateValue = objForm.Items.Item("9").Specific.Value
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "9" And pVal.BeforeAction = False Then
                        LostDateValue = objForm.Items.Item("9").Specific.Value
                        If GotDateValue <> LostDateValue Then
                            objForm.Items.Item("1000003").Specific.Value = ""
                        End If
                    End If


                    If pVal.ItemUID = "6" And pVal.ColUID = "V_6" And pVal.BeforeAction = False Then

                        'Me.SetNewLine(objForm.UniqueID, "61", objMatrix.Columns.Item("V_1").Cells.Item(objMatrix1.VisualRowCount).Specific.Value, _
                        '                                                         objMatrix1.Columns.Item("V_3").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
                        oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

                        objMatrix = objForm.Items.Item("6").Specific

                        If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value <> "" Then

                            If objMatrix.VisualRowCount = pVal.Row Then
                                objMatrix.AddRow()
                                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                                oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                                'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                                oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                                oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                                objMatrix.SetLineData(objMatrix.VisualRowCount)

                            End If
                        End If
                    End If

                    ''Added on 26-10-18 by Abinas
                    'If pVal.ItemUID = "6" And pVal.ColUID = "V_4" And pVal.BeforeAction = False Then

                    '    Try
                    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
                    '        oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

                    '        objMatrix = objForm.Items.Item("6").Specific

                    '        If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = "Load" Then


                    '            Dim GettripSheetNo As String = ""
                    '            If objMain.IsSAPHANA = True Then
                    '                GettripSheetNo = "Select ""DocNum""  From  ""@VSP_FLT_TRSHT""  Where ""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "' "
                    '            Else
                    '                GettripSheetNo = "Select ""DocNum""  From  ""@VSP_FLT_TRSHT""  Where ""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "' "
                    '            End If

                    '            Dim oRsGettripSheetNo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGettripSheetNo.DoQuery(GettripSheetNo)


                    '            Dim GetSalesOrderno As String = ""
                    '            If objMain.IsSAPHANA = True Then
                    '                GetSalesOrderno = "Select T1.""U_VSPDCNUM""  From  ""@VSP_FLT_TRSHT_C3"" T1  Where T1.""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "'"
                    '            Else
                    '                GetSalesOrderno = "Select T1.""U_VSPDCNUM""  From  ""@VSP_FLT_TRSHT_C3"" T1  Where T1.""DocEntry"" ='" & objForm.Items.Item("23").Specific.Value & "' "
                    '            End If

                    '            Dim oRsGetSalesOrderno As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGetSalesOrderno.DoQuery(GetSalesOrderno)
                    '            If oRsGetSalesOrderno.RecordCount > 0 Then
                    '                oRsGetSalesOrderno.MoveFirst()
                    '                Dim DocNum As String = oRsGetSalesOrderno.Fields.Item(0).Value
                    '                If DocNum = "" Then
                    '                    Dim tripsheetno As String = oRsGettripSheetNo.Fields.Item(0).Value
                    '                    objMain.objApplication.StatusBar.SetText("Please Post Sales Order For Trip Sheet No " & tripsheetno, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                    count = 1
                    '                    Exit Try
                    '                End If
                    '                Dim GetSalesOrderDetals As String = ""

                    '                If objMain.IsSAPHANA = True Then
                    '                    GetSalesOrderDetals = "Select T1.""ItemCode"",T1.""Quantity"" From ""RDR1"" T1  where T1.""DocEntry""= '" & oRsGetSalesOrderno.Fields.Item(0).Value & "'"
                    '                Else
                    '                    GetSalesOrderDetals = "Select T1.""ItemCode"",T1.""Quantity"" From ""RDR1"" T1  where T1.""DocEntry""= '" & oRsGetSalesOrderno.Fields.Item(0).Value & "'"
                    '                End If
                    '                Dim oRsGetSalesOrderDetals As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '                oRsGetSalesOrderDetals.DoQuery(GetSalesOrderDetals)

                    '                If oRsGetSalesOrderDetals.RecordCount > 0 Then
                    '                    oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, oRsGetSalesOrderDetals.Fields.Item(0).Value.ToString)
                    '                    oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, CDbl(oRsGetSalesOrderDetals.Fields.Item(1).Value.ToString))
                    '                End If
                    '                If objMatrix.VisualRowCount = pVal.Row Then
                    '                    ' objMatrix.AddRow()
                    '                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    '                    oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, "Load")
                    '                    oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                    '                    oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, "")
                    '                    'oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, "")
                    '                    'oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, "")
                    '                    'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, "")
                    '                    oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
                    '                    oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
                    '                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                    '                End If
                    '            End If
                    '        End If
                    '    Catch ex As Exception

                    '    End Try
                    'End If
            End Select
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetCellsEditable(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("6").Specific

            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("V_13").Cells.Item(i).Specific.Checked = True Then
                    objMatrix.CommonSetting.SetRowEditable(i, False)
                Else
                    objMatrix.CommonSetting.SetRowEditable(i, True)
                End If
            Next
            objMatrix.Columns.Item("V_-1").Editable = False
            objMatrix.Columns.Item("V_2").Editable = False
            objMatrix.Columns.Item("V_12").Editable = True
            objMatrix.Columns.Item("V_10").Editable = False
            objMatrix.Columns.Item("V_13").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("6").Specific

            If objForm.Items.Item("9").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Date Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            ElseIf objForm.Items.Item("1000003").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("23").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Trip Sheet No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("20").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Route Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VCHSTS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")
            objMatrix = objForm.Items.Item("6").Specific

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        Me.SetCellsEditable(objForm.UniqueID)
                    End If

                    'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    '    If BusinessObjectInfo.BeforeAction = True Then

                    '        Try
                    '            Dim GetVehiclefromVS As String = "Select  ISNULL(U_VSPCLKM,0) From [@VSP_VECHSTS_C0] Where  DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'" & _
                    '            " and LineId = (Select  Max(LineId)-1 From [@VSP_VECHSTS_C0] Where  DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "') "
                    '            Dim oRsGetVehiclefromVS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGetVehiclefromVS.DoQuery(GetVehiclefromVS)

                    '            Dim GetVehicleCode As String = "Select Code , ISNULL(U_VSPODRDG,0) From [@VSP_FLT_VMSTR] T0 Where T0.U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                    '            Dim oRsGetVehicleCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRsGetVehicleCode.DoQuery(GetVehicleCode)

                    '            objMain.sCmp = objMain.objCompany.GetCompanyService
                    '            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
                    '            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    '            objMain.oGeneralParams.SetProperty("Code", oRsGetVehicleCode.Fields.Item(0).Value)
                    '            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    '            objMain.oGeneralData.SetProperty("U_VSPODRDG", CDbl(oRsGetVehiclefromVS.Fields.Item(0).Value).ToString)

                    '            objMain.oGeneralService.Update(objMain.oGeneralData)
                    '        Catch ex As Exception
                    '            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '        End Try
                    '    End If


                    'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    '    If BusinessObjectInfo.BeforeAction = True Then
                    '        objMatrix = objForm.Items.Item("6").Specific

                    '        If count = 1 Then
                    '            objMain.objApplication.StatusBar.SetText("Please Post Sales Order in trip Sheet", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            BubbleEvent = False
                    '        End If
                    '        For i As Integer = 1 To objMatrix.VisualRowCount
                    '            If objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value = "" Then
                    '                objMain.objApplication.StatusBar.SetText("Please Post Sales Order in trip Sheet", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '            End If
                    '        Next



                    '    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadVehicleDetails(ByVal FormUID As String, ByVal VehicleNum As String, ByVal ConCode As String, _
                           ByVal ConName As String, ByVal ODOReading As String, ByVal row As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

            objMatrix = objForm.Items.Item("6").Specific

            objForm.Freeze(True)

            oDBs_Head.SetValue("U_VSPVNO", oDBs_Head.Offset, VehicleNum)
            oDBs_Head.SetValue("U_VSPCNTR", oDBs_Head.Offset, ConCode)
            oDBs_Head.SetValue("U_VSPCNM", oDBs_Head.Offset, ConName)
            oDBs_Head.SetValue("U_VSPTRSHT", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_VSPROUTE", oDBs_Head.Offset, "")

            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "6")

            'Set OdometerReading
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, 1)
            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, ODOReading)
            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, objMatrix.Columns.Item("V_12").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(1).Specific.Value)
            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, objMatrix.Columns.Item("V_9").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")
            objMatrix.SetLineData(1)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadRouteDetails(ByVal FormUID As String, ByVal Route As String, ByVal Source As String, ByVal Desintation As String, ByVal row As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

            objMatrix = objForm.Items.Item("6").Specific
            objForm.Freeze(True)

            oDBs_Head.SetValue("U_VSPROUTE", oDBs_Head.Offset, Route)
            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Head.Offset, Source)
            oDBs_Details.SetValue("U_VSPDST", oDBs_Head.Offset, Desintation)

            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, 1)
            oDBs_Details.SetValue("U_VSPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPSTA", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPOPKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPCLKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPSOURC", oDBs_Details.Offset, Source)
            oDBs_Details.SetValue("U_VSPDST", oDBs_Details.Offset, objMatrix.Columns.Item("V_11").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPLOC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPCHML", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPQUA", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(1).Specific.Value)
            'oDBs_Details.SetValue("U_VSPCSVE", oDBs_Details.Offset, objMatrix.Columns.Item("V_9").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPREM", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPTOTKM", oDBs_Details.Offset, objMatrix.Columns.Item("V_10").Cells.Item(1).Specific.Value)
            oDBs_Details.SetValue("U_VSPCHK", oDBs_Details.Offset, "N")
            oDBs_Details.SetValue("U_VSPCHK1", oDBs_Details.Offset, "N")

            objMatrix.SetLineData(1)
            objMatrix.AutoResizeColumns()
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpadateclosekmtoVMwithoutTripSht(ByVal FormUID As String)

        Try
            Dim GetVehiclefromVS As String = ""
            If objMain.IsSAPHANA = True Then
                GetVehiclefromVS = "Select  IFNULL(""U_VSPCLKM"",0) From ""@VSP_VECHSTS_C0"" Where  ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'" & _
                                            "and ""LineId"" = (Select  Max(""LineId"")-1 From ""@VSP_VECHSTS_C0"" Where  ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "') "
            Else
                GetVehiclefromVS = "Select  ISNULL(U_VSPCLKM,0) From [@VSP_VECHSTS_C0] Where  DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'" & _
                                                            "and LineId = (Select  Max(LineId)-1 From [@VSP_VECHSTS_C0] Where  DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "') "
            End If
            Dim oRsGetVehiclefromVS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehiclefromVS.DoQuery(GetVehiclefromVS)

            Dim GetVehicleCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetVehicleCode = "Select ""Code"" , IFNULL(""U_VSPODRDG"",0) From ""@VSP_FLT_VMSTR"" T0 Where T0.""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "'"
            Else
                GetVehicleCode = "Select Code , ISNULL(U_VSPODRDG,0) From [@VSP_FLT_VMSTR] T0 Where T0.U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "'"
            End If
            Dim oRsGetVehicleCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehicleCode.DoQuery(GetVehicleCode)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", oRsGetVehicleCode.Fields.Item(0).Value)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            Dim s As String = oRsGetVehiclefromVS.Fields.Item(0).Value.ToString
            objMain.oGeneralData.SetProperty("U_VSPODRDG", CDbl(oRsGetVehiclefromVS.Fields.Item(0).Value).ToString)

            objMain.oGeneralService.Update(objMain.oGeneralData)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

        'Me.RefreshData(objForm.UniqueID)

        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        End If

    End Sub

    Sub UpadateOdmtrRdngsFromVMWithTripSht(ByVal FormUID As String)

        Try
            Dim GetVehiclefromVM As String = ""
            If objMain.IsSAPHANA = True Then
                GetVehiclefromVM = "Select  IFNULL(""U_VSPODRDG"",0) from ""@VSP_FLT_VMSTR"" T1 where " & _
                                            "T1.""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "' and ""U_VSPCHK"" = 'Y' "
            Else
                GetVehiclefromVM = "Select  ISNULL(U_VSPODRDG,0) from [@VSP_FLT_VMSTR] T1 where " & _
                                            "T1.U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "' and U_VSPCHK = 'Y' "
            End If
            Dim oRsGetVehiclefromVM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehiclefromVM.DoQuery(GetVehiclefromVM)

            Dim GetVehicleCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetVehicleCode = "Select max(""LineId"") from ""@VSP_VECHSTS"" T0 Inner Join ""@VSP_VECHSTS_C0"" T1 ON T0.""DocEntry"" =  T1.""DocEntry"" " & _
                                        "Where T0.""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "' and T0.""DocEntry"" = '" & objForm.Items.Item("8").Specific.Value & "' "
            Else
                GetVehicleCode = "Select max(LineId) from [@VSP_VECHSTS] T0 Inner Join [@VSP_VECHSTS_C0] T1 ON T0.DocEntry =  T1.DocEntry " & _
                                        "Where T0.U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "' and T0.DocEntry = '" & objForm.Items.Item("8").Specific.Value & "' "
            End If
            Dim oRsGetVehicleCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehicleCode.DoQuery(GetVehicleCode)

            Dim LineId As Integer = CInt(oRsGetVehicleCode.Fields.Item(0).Value)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVEHST")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", objForm.Items.Item("8").Specific.Value)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oChildren = objMain.oGeneralData.Child("VSP_VECHSTS_C0")
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPOPKM", oRsGetVehiclefromVM.Fields.Item(0).Value)
            objMain.oGeneralService.Update(objMain.oGeneralData)

            objMain.objApplication.StatusBar.SetText("Odometer Readings Sucessfully Updated ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            Me.RefreshData(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub UpdateVehStatusInTrpSht(ByVal FormUID As String, ByVal DocDate As String, ByVal Vehicleno As String, ByVal Route As String _
                                , ByVal Contcode As String, ByVal Contname As String, _
                                ByVal DocNum As String, ByVal TripShtDocNum As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("6").Specific

            If TripShtDocNum = "" Then
                objMain.objApplication.StatusBar.SetText("Please select Tripsheet No.")
                Exit Try
            End If

            Dim GetCount As String = ""

            If objMain.IsSAPHANA = TripShtDocNum Then
                GetCount = "Select  Max(T1.""LineId"") ""C1LineId"", MAX (T3.""LineId"") ""C6LineId"", MAX (T4.""LineId"") ""C3LineId"" From  ""@VSP_FLT_TRSHT"" T0 " & _
                                     "Inner Join ""@VSP_FLT_TRSHT_C1"" T1 On T0.""DocEntry"" = T1.""DocEntry"" " & _
                                     "Inner Join ""@VSP_FLT_TRSHT_C6"" T3 On T3.""DocEntry"" = T1.""DocEntry""  " & _
                                     "Inner Join ""@VSP_FLT_TRSHT_C3"" T4 On T4.""DocEntry"" = T1.""DocEntry"" " & _
                                     "Where T0.""DocEntry""  = '" & TripShtDocNum & "'"
            Else
                GetCount = "Select  Max(T1.LineId) 'C1LineId', MAX (T3.LineId) 'C6LineId', MAX (T4.LineId) 'C3LineId' From  [@VSP_FLT_TRSHT] T0 " & _
                                     "Inner Join [@VSP_FLT_TRSHT_C1] T1 On T0.DocEntry = T1.DocEntry " & _
                                     "Inner Join [@VSP_FLT_TRSHT_C6] T3 On T3.DocEntry = T1.DocEntry  " & _
                                     "Inner Join [@VSP_FLT_TRSHT_C3] T4 On T4.DocEntry = T1.DocEntry " & _
                                     "Where T0.DocEntry  = '" & TripShtDocNum & "'"
            End If
            Dim oRsGetCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCount.DoQuery(GetCount)





            'Dim GetCount As String = "Select  Max(T1.LineId) 'C1LineId', MAX (T3.LineId) 'C6LineId', MAX (T4.LineId) 'C3LineId' From  [@VSP_FLT_TRSHT] T0 " & _
            '                        "Inner Join [@VSP_FLT_TRSHT_C1] T1 On T0.DocEntry = T1.DocEntry " & _
            '                        "Inner Join [@VSP_FLT_TRSHT_C6] T3 On T3.DocEntry = T1.DocEntry  " & _
            '                        "Inner Join [@VSP_FLT_TRSHT_C3] T4 On T4.DocEntry = T1.DocEntry " & _
            '                        "Where T0.DocEntry  = (Select DocEntry From [@VSP_FLT_TRSHT] T0 where DocNum ='" & TripShtDocNum & "')"
            'Dim oRsGetCount As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRsGetCount.DoQuery(GetCount)

            'Dim GetDetails As String = "Select T1.LineId,T1.U_VSPTY, T1.U_VSPSTA, T1.U_VSPOPKM, T1.U_VSPCLKM, T1.U_VSPLOC, " & _
            '"T1.U_VSPCHML, T1.U_VSPQUA, T1.U_VSPCSVE, T1.U_VSPREM, T1.U_VSPTOTKM, T1.U_VSPSOURC, T1.U_VSPDST From [@VSP_VECHSTS_C0] T1 " & _
            '"where docentry =(Select DocEntry From [@VSP_VECHSTS] T0 where DocNum ='" & DocNum & "')  and T1.U_VSPTY <> ''  and T1.U_VSPCHK = 'Y' and T1.U_VSPCHK1 = 'N'"

            Dim GetDetails As String = ""

            If objMain.IsSAPHANA = True Then
                GetDetails = "Select T1.""LineId"",T1.""U_VSPTY"", T1.""U_VSPSTA"", T1.""U_VSPOPKM"", T1.""U_VSPCLKM"", T1.""U_VSPLOC"", " & _
          "T1.""U_VSPCHML"", T1.""U_VSPQUA"", T1.""U_VSPREM"", T1.""U_VSPTOTKM"", T1.""U_VSPSOURC"", T1.""U_VSPDST"" From ""@VSP_VECHSTS_C0"" T1 " & _
          "where ""DocEntry"" =(Select ""DocEntry"" From ""@VSP_VECHSTS"" T0 where ""DocNum"" ='" & DocNum & "')  and T1.""U_VSPTY"" <> ''  and T1.""U_VSPCHK"" = 'Y' and T1.""U_VSPCHK1"" = 'N'"
            Else
                GetDetails = "Select T1.LineId,T1.U_VSPTY, T1.U_VSPSTA, T1.U_VSPOPKM, T1.U_VSPCLKM, T1.U_VSPLOC, " & _
          "T1.U_VSPCHML, T1.U_VSPQUA, T1.U_VSPREM, T1.U_VSPTOTKM, T1.U_VSPSOURC, T1.U_VSPDST From [@VSP_VECHSTS_C0] T1 " & _
          "where docentry =(Select DocEntry From [@VSP_VECHSTS] T0 where DocNum ='" & DocNum & "')  and T1.U_VSPTY <> ''  and T1.U_VSPCHK = 'Y' and T1.U_VSPCHK1 = 'N'"
            End If

            Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDetails.DoQuery(GetDetails)

            oRsGetDetails.MoveFirst()
            'Dim getentry As String = "Select DocEntry From [@VSP_FLT_TRSHT] T0 where DocNum ='" & TripShtDocNum & "'"
            'Dim orsgetentry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'orsgetentry.DoQuery(getentry)
            Dim Getdate As String = ""
            If objMain.IsSAPHANA = True Then
                Getdate = "Select ""U_VSPSRTDT"" From ""@VSP_FLT_TRSHT"" where ""DocEntry""  = '" & TripShtDocNum & "'"
            Else
                Getdate = "Select U_VSPSRTDT From [@VSP_FLT_TRSHT] where DocEntry  = '" & TripShtDocNum & "'"
            End If

            Dim oRsGetdate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetdate.DoQuery(Getdate)

            Dim C1LineId As Integer = CInt(oRsGetCount.Fields.Item(0).Value)
            Dim C6LineId As Integer = CInt(oRsGetCount.Fields.Item(1).Value)
            Dim C3LineId As Integer = CInt(oRsGetCount.Fields.Item(2).Value)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", TripShtDocNum)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

            DocDate = objForm.Items.Item("9").Specific.Value
            DocDate = DocDate.Insert("4", "-")
            DocDate = DocDate.Insert("7", "-")

            If oRsGetDetails.RecordCount > 0 Then
                For i As Integer = 1 To oRsGetDetails.RecordCount
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C1")

                    Dim Getchmname As String = ""

                    If objMain.IsSAPHANA = True Then
                        Getchmname = "select ""ItemName""  from OITM  where ""ItemCode"" = '" & oRsGetDetails.Fields.Item("U_VSPCHML").Value & "'"

                    Else
                        Getchmname = "select ItemName  from OITM  where ItemCode = '" & oRsGetDetails.Fields.Item("U_VSPCHML").Value & "'"
                    End If
                    Dim oRsGGetchmname As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGGetchmname.DoQuery(Getchmname)

                    If oRsGetDetails.Fields.Item("U_VSPTY").Value = "Load" Then
                        objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPLOAD", "Yes")
                    ElseIf oRsGetDetails.Fields.Item("U_VSPTY").Value = "UnLoad" Then
                        objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPLOAD", "No")
                    End If

                    Dim GetMileageDetails As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetMileageDetails = "Select ""U_VSPMWL"" , ""U_VSPMWOL"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                    Else
                        GetMileageDetails = "Select U_VSPMWL , U_VSPMWOL From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("1000003").Specific.Value & "'"
                    End If
                    Dim oRsGetMileageDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetMileageDetails.DoQuery(GetMileageDetails)

                    Dim DieselConsumption As Double = oRsGetDetails.Fields.Item("U_VSPTOTKM").Value
                    If oRsGetDetails.Fields.Item("U_VSPTY").Value = "Load" Then
                        DieselConsumption = DieselConsumption / oRsGetMileageDetails.Fields.Item(0).Value
                    ElseIf oRsGetDetails.Fields.Item("U_VSPTY").Value = "UnLoad" Then
                        DieselConsumption = DieselConsumption / oRsGetMileageDetails.Fields.Item(1).Value
                    End If
                    Dim l As Integer = C1LineId
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPOPKM", CDbl(oRsGetDetails.Fields.Item("U_VSPOPKM").Value))
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPCLKM", CDbl(oRsGetDetails.Fields.Item("U_VSPCLKM").Value))
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPSOUR", oRsGetDetails.Fields.Item("U_VSPSOURC").Value)
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPDEST", oRsGetDetails.Fields.Item("U_VSPDST").Value)
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPFRDT", oRsGetdate.Fields.Item("U_VSPSRTDT").Value)
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPTODT", DocDate)
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_VSPDICON", DieselConsumption.ToString)
                    objMain.oChildren.Item(C1LineId - 1).SetProperty("U_TOTKM", CDbl(oRsGetDetails.Fields.Item("U_VSPTOTKM").Value))
                    objMain.oChildren.Add()

                    If oRsGetDetails.Fields.Item("U_VSPCHML").Value <> "" Then
                        objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C6")
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPSOU", oRsGetDetails.Fields.Item("U_VSPSOURC").Value)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPSOUR", oRsGetDetails.Fields.Item("U_VSPDST").Value)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPCHCOD", oRsGetDetails.Fields.Item("U_VSPCHML").Value)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPCHNAM", oRsGGetchmname.Fields.Item("ItemName").Value.ToString.Trim)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPWEIGH", oRsGetDetails.Fields.Item("U_VSPQUA").Value)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPDAT", DocDate)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPFRDT", oRsGetdate.Fields.Item("U_VSPSRTDT").Value)
                        objMain.oChildren.Item(C6LineId - 1).SetProperty("U_VSPTODT", DocDate)
                        objMain.oChildren.Add()
                    End If

                    'If oRsGetDetails.Fields.Item("U_VSPCSVE").Value <> "" Then
                    '    Dim CardType As String = "Select CardType From OCRD Where CardCode = '" & oRsGetDetails.Fields.Item("U_VSPCSVE").Value & "'"
                    '    Dim oRsCardType As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '    oRsCardType.DoQuery(CardType)

                    '    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C3")
                    '    If oRsCardType.Fields.Item(0).Value = "S" Then
                    '        objMain.oChildren.Item(C3LineId - 1).SetProperty("U_VSPTYPE", "Purchase")
                    '    ElseIf oRsCardType.Fields.Item(0).Value = "C" Then
                    '        objMain.oChildren.Item(C3LineId - 1).SetProperty("U_VSPTYPE", "Sales")
                    '    End If
                    '    objMain.oChildren.Item(C3LineId - 1).SetProperty("U_VSPBPCOD", oRsGetDetails.Fields.Item("U_VSPCSVE").Value)
                    '    objMain.oChildren.Item(C3LineId - 1).SetProperty("U_VSPDATE", DocDate)
                    '    objMain.oChildren.Add()
                    'End If

                    C6LineId += 1
                    C1LineId += 1
                    C3LineId += 1

                    oRsGetDetails.MoveNext()
                Next
                objMain.oGeneralService.Update(objMain.oGeneralData)
            End If

            If oRsGetDetails.RecordCount > 0 Then
                oRsGetDetails.MoveFirst()

                Dim GetDocEntry As String = ""
                If objMain.IsSAPHANA = True Then
                    GetDocEntry = "Select ""DocEntry"" From  ""@VSP_VECHSTS"" Where ""DocNum"" = '" & objForm.Items.Item("8").Specific.Value & "'"
                Else
                    GetDocEntry = "Select DocEntry From  [@VSP_VECHSTS] Where DocNum = '" & objForm.Items.Item("8").Specific.Value & "'"
                End If
                Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDocEntry.DoQuery(GetDocEntry)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVEHST")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oRsGetDocEntry.Fields.Item(0).Value)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                For i As Integer = 1 To oRsGetDetails.RecordCount

                    Dim LineId As Integer = oRsGetDetails.Fields.Item("LineId").Value
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_VECHSTS_C0")
                    objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPCHK1", "Y")
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                    oRsGetDetails.MoveNext()
                Next
            End If

            objMain.objApplication.StatusBar.SetText("Vehicle Status Sucessfully Updated in TripSheet", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Me.RefreshData(objForm.UniqueID)
            Me.SetCellsEditable(objForm.UniqueID)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLFilterTripSheet(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "U_VSPVHCL"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = objForm.Items.Item("1000003").Specific.Value
            oChooseFromList.SetConditions(oConditions)
            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPSTS"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Open"
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RefreshData(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0")

            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_VECHSTS_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))

            objMatrix = objForm.Items.Item("6").Specific

            objMatrix.LoadFromDataSource()
            objMatrix.AutoResizeColumns()
            objForm.Refresh()
            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterForVehicles(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetTankItems As String = ""
            If objMain.IsSAPHANA = True Then
                GetTankItems = "Select ""U_VSPVNO"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" Not In (Select ""U_VSPVNO"" From ""@VSP_VECHSTS"" " & _
           "Where ""U_VSPDT"" = '" & objForm.Items.Item("9").Specific.Value & "') "
            Else
                GetTankItems = "Select U_VSPVNO From [@VSP_FLT_VMSTR] Where U_VSPVNO Not In (Select U_VSPVNO From [@VSP_VECHSTS] " & _
           "Where Convert(Date,U_VSPDT,103) = '" & objForm.Items.Item("9").Specific.Value & "') "
            End If
            Dim oRsGetTankItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTankItems.DoQuery(GetTankItems)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetTankItems.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "U_VSPVNO"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetTankItems.Fields.Item("U_VSPVNO").Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetTankItems.MoveNext()
                For i As Integer = 1 To oRsGetTankItems.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "U_VSPVNO"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetTankItems.Fields.Item("U_VSPVNO").Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetTankItems.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "U_VSPVNO"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterSO(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetDocEntry As String = ""
            If objMain.IsSAPHANA = True Then

                GetDocEntry = "Select A.""U_VSPDCNUM"" from ""@VSP_FLT_TRSHT_C3"" A Inner Join ""@VSP_FLT_TRSHT"" B On A.""DocEntry""=B.""DocEntry""  Inner Join ORDR C On C.""DocEntry""=A.""U_VSPDCNUM"" Where B.""DocEntry""='" & objForm.Items.Item("23").Specific.Value & "' And A.""U_VSPDCNUM""<>'' And C.""DocStatus""='O'"
            Else
                GetDocEntry = "Select A.""U_VSPDCNUM"" from ""@VSP_FLT_TRSHT_C3"" A Inner Join ""@VSP_FLT_TRSHT"" B On A.""DocEntry""=B.""DocEntry""  Inner Join ORDR C On C.""DocEntry""=A.""U_VSPDCNUM"" Where B.""DocEntry""='" & objForm.Items.Item("23").Specific.Value & "' And A.""U_VSPDCNUM""<>'' And C.""DocStatus""='O'"
            End If
            Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocEntry.DoQuery(GetDocEntry)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()


            If oRsGetDocEntry.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "DocEntry"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetDocEntry.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetDocEntry.MoveNext()
                For i As Integer = 1 To oRsGetDocEntry.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "DocEntry"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetDocEntry.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetDocEntry.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "DocEntry"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If















            'Dim oConditions As SAPbouiCOM.Conditions
            'Dim oCondition As SAPbouiCOM.Condition
            'Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            'Dim emptyCon As New SAPbouiCOM.Conditions
            'oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            'oChooseFromList.SetConditions(emptyCon)
            'oConditions = oChooseFromList.GetConditions()
            'If oRsGetDocNum.RecordCount > 0 Then
            '    oCondition = oConditions.Add()
            '    oCondition.Alias = "U_VSPDCNUM"
            '    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCondition.CondVal =
            '    oChooseFromList.SetConditions(oConditions)
            '    oRsGetDocNum.MoveNext()
            '    For i As Integer = 1 To oRsGetDocNum.RecordCount - 1
            '        oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            '        oCondition = oConditions.Add()
            '        oCondition.Alias = "U_VSPVNO"
            '        oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '        oCondition.CondVal = oRsGetDocNum.Fields.Item("U_VSPVNO").Value
            '        oChooseFromList.SetConditions(oConditions)
            '        oRsGetDocNum.MoveNext()
            '    Next
            'Else
            '    oCondition = oConditions.Add()
            '    oCondition.Alias = "U_VSPVNO"
            '    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            '    oCondition.CondVal = ""
            '    oChooseFromList.SetConditions(oConditions)
            'End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class

