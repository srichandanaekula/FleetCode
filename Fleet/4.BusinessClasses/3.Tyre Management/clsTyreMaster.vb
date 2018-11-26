Public Class clsTyreMaster

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Tyre Master.xml", "VSP_FLT_TYRMSTR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TYRMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMSTR")

            Me.CFLFilterVendors(objForm.UniqueID, "CFL_VCD")

            Me.CellsMasking(objForm.UniqueID)

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMSTR")
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSP_FLT_TYRMSTR"))

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("500").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("500").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("81").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("81").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("79").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("79").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000006").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000006").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("83").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("83").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("71").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("71").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("71").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("60").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("60").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("60").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("75").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("75").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("75").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("77").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("77").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("77").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("86").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("86").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("86").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("66").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("66").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("66").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
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
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "200" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim ChkItemExist As String = ""

                        If objMain.IsSAPHANA = True Then
                            ChkItemExist = "Select ""Code"" From ""@VSP_FLT_TYRMSTR"" Where ""U_VSPTRNUM"" ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"
                        Else
                            ChkItemExist = "Select Code From [@VSP_FLT_TYRMSTR] Where [U_VSPTRNUM] ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"
                        End If

                        Dim oRsChkItemExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkItemExist.DoQuery(ChkItemExist)
                        If oRsChkItemExist.RecordCount > 0 Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("500").Specific.value = oRsChkItemExist.Fields.Item(0).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMSTR")

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

                        If oCFL.UniqueID = "CFL_WHS" Then
                            oDBs_Head.SetValue("U_VSPSLOC", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_VCD" Then
                            oDBs_Head.SetValue("U_VSPPCHFM", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_ITMCD" Then
                            oDBs_Head.SetValue("U_VSPITCD", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_GIIC" Then
                            oDBs_Head.SetValue("U_VSPGIIC", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_TXCD" Then
                            oDBs_Head.SetValue("U_VSPTXCD", oDBs_Head.Offset, oDT.GetValue("Code", 0))
                        End If

                        If oCFL.UniqueID = "CFL_PSTN" Then
                            oDBs_Head.SetValue("U_VSPPSTN", oDBs_Head.Offset, oDT.GetValue("U_VSPPSTN", 0))
                        End If

                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TYRMSTR" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
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

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TYRMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        If objForm.Items.Item("200").Specific.Value = "" Then
                            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        Else
                            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region



    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objForm.Items.Item("200").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tyre No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            ElseIf objForm.Items.Item("8").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tyre Name Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            ElseIf objForm.Items.Item("75").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("GR ItemCode Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            ElseIf objForm.Items.Item("77").Specific.Value = 0.0 Then
                objMain.objApplication.StatusBar.SetText("Price Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("86").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("TaxCode Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("66").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Storage Location Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("71").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Purchase From Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("60").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Purchased On Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Function


    'Function Validation(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        If objForm.Items.Item("75").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("ItemCode Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("200").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("Tyre No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("8").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("Tyre Name Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("77").Specific.Value = 0.0 Then
    '            objMain.objApplication.StatusBar.SetText("Price Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("86").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("TaxCode Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("66").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("Storage Location Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("71").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("Purchase From Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        ElseIf objForm.Items.Item("60").Specific.Value = "" Then
    '            objMain.objApplication.StatusBar.SetText("Purchased On Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '            Return False
    '        End If
    '        Return True
    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try

    'End Function

End Class
