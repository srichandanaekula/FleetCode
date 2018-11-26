Public Class clsTankMapping
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oColumn As SAPbouiCOM.Column
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Tank Mapping.xml", "VSP_FLT_TANKMPG_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TANKMPG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG_C0")

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMatrix1 = objForm.Items.Item("7").Specific

            objMatrix1.Columns.Item("V_8").ValidValues.Add("", "")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Attached", "Attached")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Removed", "Removed")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Repair", "Repair")

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG_C0")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OTANKMPG"))

            objMatrix1 = objForm.Items.Item("7").Specific
            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()

            Me.SetNewLine(objForm.UniqueID)

            For i As Integer = 1 To objMatrix1.VisualRowCount
                objMatrix1.CommonSetting.SetRowEditable(i, True)
            Next

            objMatrix1.Columns.Item("V_-1").Editable = False
            objMatrix1.Columns.Item("V_1").Editable = False
            objMatrix1.Columns.Item("V_2").Editable = False
            objMatrix1.Columns.Item("V_5").Editable = False
            objMatrix1.Columns.Item("V_6").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG_C0")

            objMatrix1 = objForm.Items.Item("7").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_VSPTNUM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPTNM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPTYP", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, "")
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
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

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                                          pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMPG_C0")
                    objMatrix1 = objForm.Items.Item("7").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "CFL_TYRNO" Then
                            Me.CFLFilter(objForm.UniqueID, oCFL.UniqueID)
                        End If
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_VCHNO" Then
                            oDBs_Head.SetValue("U_VSPVCHN0", oDBs_Head.Offset, oDT.GetValue("U_VSPVNO", 0))
                            oDBs_Head.SetValue("U_VSPVCHNM", oDBs_Head.Offset, oDT.GetValue("U_VSPVNM", 0))
                            oDBs_Head.SetValue("U_VSPVMDL", oDBs_Head.Offset, oDT.GetValue("U_VSPMODEL", 0))
                        End If

                        If oCFL.UniqueID = "CFL_TYRNO" Then
                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_VSPTNUM", oDBs_Details1.Offset, oDT.GetValue("U_VSPTNKNO", 0))
                            oDBs_Details1.SetValue("U_VSPTNM", oDBs_Details1.Offset, oDT.GetValue("U_VSPTNKNM", 0))
                            oDBs_Details1.SetValue("U_VSPTYP", oDBs_Details1.Offset, oDT.GetValue("U_VSPTTYPE", 0))
                            oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, oDT.GetValue("U_VSPCPCTY", 0))
                            oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, oDT.GetValue("U_VSPUOM1", 0))
                            oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix1.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "200" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim ChkItemExist As String = ""
                        If objMain.IsSAPHANA = True Then
                            ChkItemExist = "Select ""DocNum"" From ""@VSP_FLT_TANKMPG"" Where ""U_VSPVCHN0"" ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"
                        Else
                            ChkItemExist = "Select DocNum From [@VSP_FLT_TANKMPG] Where [U_VSPVCHN0] ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"
                        End If

                        Dim oRsChkItemExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkItemExist.DoQuery(ChkItemExist)
                        If oRsChkItemExist.RecordCount > 0 Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("12").Specific.value = oRsChkItemExist.Fields.Item(0).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TANKMPG" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TANKMPG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMatrix1 = objForm.Items.Item("7").Specific

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Me.UpdateVehicleNo(objForm.UniqueID)

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Me.UpdateVehicleNo(objForm.UniqueID)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        For i As Integer = 1 To objMatrix1.VisualRowCount
                            If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" Then
                                objMatrix1.CommonSetting.SetRowEditable(i, False)
                            ElseIf (objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Or _
                               objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Repair" Or _
                               objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "") Then
                                objMatrix1.CommonSetting.SetRowEditable(i, True)
                            End If
                        Next

                        objMatrix1.Columns.Item("V_-1").Editable = False
                        objMatrix1.Columns.Item("V_1").Editable = False
                        objMatrix1.Columns.Item("V_2").Editable = False
                        objMatrix1.Columns.Item("V_5").Editable = False
                        objMatrix1.Columns.Item("V_6").Editable = False
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Sub UpdateVehicleNo(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("7").Specific

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                    Dim GetCode As String = ""

                    If objMain.IsSAPHANA = True Then
                        GetCode = "Select ""Code"" From ""@VSP_FLT_TANKMSTR"" Where ""U_VSPTNKNO"" = '" & objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value & "'"
                    Else
                        GetCode = "Select Code From [@VSP_FLT_TANKMSTR] Where U_VSPTNKNO = '" & objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value & "'"
                    End If
                    Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetCode.DoQuery(GetCode)

                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTANKMSTR")
                    objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    objMain.oGeneralParams.SetProperty("Code", oRsGetCode.Fields.Item(0).Value)
                    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Then
                        objMain.oGeneralData.SetProperty("U_VSPVNO", objForm.Items.Item("200").Specific.Value.ToString)
                    ElseIf (objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" Or _
                            objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Repair") Then
                        objMain.oGeneralData.SetProperty("U_VSPVNO", "")
                    End If
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If
            Next

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" Then
                    objMatrix1.CommonSetting.SetRowEditable(i, False)
                ElseIf (objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Or _
                              objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "") Then
                    objMatrix1.CommonSetting.SetRowEditable(i, True)
                End If
            Next

            objMatrix1.Columns.Item("V_2").Editable = False
            objMatrix1.Columns.Item("V_5").Editable = False
            objMatrix1.Columns.Item("V_6").Editable = False
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "U_VSPVNO"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
            oChooseFromList.SetConditions(oConditions)

            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPVNO"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = ""
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("7").Specific

            Dim AttachedCount As Integer = 0

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Then
                    AttachedCount = AttachedCount + 1

                End If
            Next

            If AttachedCount > 1 Then
                objMain.objApplication.StatusBar.SetText("Tank can only be 1 for a Vehicle")
                Return False
            End If


            If objForm.Items.Item("200").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value = "" Then
                    objMain.objApplication.StatusBar.SetText("Tank No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Function
End Class
