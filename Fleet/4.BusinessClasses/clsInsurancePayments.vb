Public Class clsInsurancePayments
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Insurance Payments.xml", "VSP_FLT_INSPAY_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_INSPAY_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0")

            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.CFLAccounts(objForm.UniqueID, "CFL_CSHACC")
            Me.CFLAccounts(objForm.UniqueID, "CFL_INSACC")
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0")

            objMatrix = objForm.Items.Item("11").Specific
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OINSPAY"))

            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            objMatrix.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_INSPAY" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
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
                    objMatrix = objForm.Items.Item("11").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "14" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.PostJE(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("11").Specific
                    If pVal.Row = objMatrix.VisualRowCount And pVal.BeforeAction = True Then
                        If (objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value) > 0 Then
                            If CDbl(objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value) <> CDbl(objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value) Then
                                objMain.objApplication.StatusBar.SetText("Total Payment must be equal to Amount to be Paid", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value = 0
                                BubbleEvent = False
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0")
                    objMatrix = objForm.Items.Item("11").Specific
                    If pVal.ItemUID = "10" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("10").Specific.Value <> "" Then

                            objMatrix.Clear()
                            oDBs_Details.Clear()
                            objMatrix.FlushToDataSource()
                            objMatrix.AutoResizeColumns()

                            Dim NoofInstallments As Integer = objForm.Items.Item("10").Specific.Value
                            Dim InsuranceAmount As Double = objForm.Items.Item("13").Specific.Value
                            Dim TotalAmount As Double = InsuranceAmount / NoofInstallments
                            For i As Integer = 1 To NoofInstallments

                                objMatrix.AddRow()
                                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, i)
                                oDBs_Details.SetValue("U_VSPFRMDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(i).Specific.Value)
                                oDBs_Details.SetValue("U_VSPTODT", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(i).Specific.Value)
                                oDBs_Details.SetValue("U_VSPAMT", oDBs_Details.Offset, TotalAmount)
                                oDBs_Details.SetValue("U_VSPTBPD", oDBs_Details.Offset, TotalAmount)
                                oDBs_Details.SetValue("U_VSPTOTPY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value)
                                oDBs_Details.SetValue("U_VSPJENO", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value)
                                oDBs_Details.SetValue("U_VSPDATE", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(i).Specific.Value)
                                oDBs_Details.SetValue("U_VSPCSHAC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value)
                                oDBs_Details.SetValue("U_VSPINSAC", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value)
                                objMatrix.SetLineData(i)

                            Next
                        End If
                    End If
                    If pVal.ItemUID = "11" And pVal.ColUID = "V_4" And pVal.BeforeAction = False Then
                        If objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value > 0 Then
                            Dim Amount As Double = objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value
                            Dim InsuranceAmount As Double = objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value
                            Dim AmounttobePaid As Double = objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value
                            Dim DeducedAmount As Double = AmounttobePaid - Amount
                            If pVal.Row < objMatrix.VisualRowCount Then
                                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row + 1)
                                oDBs_Details.SetValue("U_VSPFRMDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPTODT", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPAMT", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPTBPD", oDBs_Details.Offset, DeducedAmount + InsuranceAmount)
                                oDBs_Details.SetValue("U_VSPTOTPY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPJENO", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPDATE", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPCSHAC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row + 1).Specific.Value)
                                oDBs_Details.SetValue("U_VSPINSAC", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row + 1).Specific.Value)
                                objMatrix.SetLineData(pVal.Row + 1)
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0")

                    objMatrix = objForm.Items.Item("11").Specific

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

                        If oCFL.UniqueID = "CFL_VHCD" Or oCFL.UniqueID = "CFL_VHNM" Then
                            oDBs_Head.SetValue("U_VSPVNO", oDBs_Head.Offset, oDT.GetValue("U_VSPVNO", 0))
                            oDBs_Head.SetValue("U_VSPVNM", oDBs_Head.Offset, oDT.GetValue("U_VSPVNM", 0))
                        End If

                        If oCFL.UniqueID = "CFL_CSHACC" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPFRMDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTODT", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPAMT", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTBPD", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTOTPY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPJENO", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPDATE", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCSHAC", oDBs_Details.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Details.SetValue("U_VSPINSAC", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_INSACC" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPFRMDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTODT", oDBs_Details.Offset, objMatrix.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPAMT", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTBPD", oDBs_Details.Offset, objMatrix.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTOTPY", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPJENO", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPDATE", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPCSHAC", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPINSAC", oDBs_Details.Offset, oDT.GetValue("AcctCode", 0))
                            objMatrix.SetLineData(pVal.Row)
                        End If
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub PostJE(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0")

            objMatrix = objForm.Items.Item("11").Specific

            Dim i As Integer = 0
            Dim LineSelect As Boolean = False
            For i = 1 To objMatrix.VisualRowCount
                If objMatrix.IsRowSelected(i) = True Then
                    LineSelect = True
                    Exit For
                End If
            Next
            If LineSelect = False Then Exit Try

            If objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value <> "" Then
                objMain.objApplication.StatusBar.SetText("Journey Entry already Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Cash Account Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Insurance Account Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim JourneyDate As String = objMatrix.Columns.Item("V_2").Cells.Item(i).Specific.Value
            JourneyDate = JourneyDate.Insert("4", "-")
            JourneyDate = JourneyDate.Insert("7", "-")

            oJE.ReferenceDate = JourneyDate
            oJE.TaxDate = JourneyDate
            oJE.DueDate = JourneyDate

            oJE.Lines.AccountCode = objMatrix.Columns.Item("V_1").Cells.Item(i).Specific.Value
            oJE.Lines.Debit = objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value

            oJE.Lines.Add()

            oJE.Lines.AccountCode = objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value
            oJE.Lines.Credit = objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value

            If oJE.Add = 0 Then
                objMain.objApplication.StatusBar.SetText("Journey Entry Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim GetTransID As String = ""
                If objMain.IsSAPHANA = True Then
                    GetTransID = "Select Max(""TransId"") From OJDT"
                Else
                    GetTransID = "Select Max(TransId) From OJDT"
                End If
                Dim oRsGetTransID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetTransID.DoQuery(GetTransID)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OINSPAY")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_INSPAY_C0")
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPJENO", oRsGetTransID.Fields.Item(0).Value.ToString)      
                objMain.oGeneralService.Update(objMain.oGeneralData)

                oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY")
                oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0")

                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_INSPAY_C0"), "DocEntry", oDBs_Details.GetValue("DocEntry", 0))

                objMatrix = objForm.Items.Item("11").Specific
                objMatrix.LoadFromDataSource()
                objMatrix.AutoResizeColumns()
                objForm.Refresh()
                Me.SetCellsEditable(objForm.UniqueID)

            Else
                objMain.objApplication.StatusBar.SetText("Failed to post JE, Error : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetCellsEditable(ByVal FormUID As String)
        Try
            objMatrix = objForm.Items.Item("11").Specific

            For i As Integer = 1 To objMatrix.VisualRowCount
                If objMatrix.Columns.Item("V_3").Cells.Item(i).Specific.Value <> "" Then
                    objMatrix.CommonSetting.SetRowEditable(i, False)
                Else
                    objMatrix.CommonSetting.SetRowEditable(i, True)
                End If
            Next
            objMatrix.Columns.Item("V_-1").Editable = False
            objMatrix.Columns.Item("V_5").Editable = False
            objMatrix.Columns.Item("V_8").Editable = False
            objMatrix.Columns.Item("V_3").Editable = False
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLAccounts(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_INSPAY_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        Me.SetCellsEditable(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class