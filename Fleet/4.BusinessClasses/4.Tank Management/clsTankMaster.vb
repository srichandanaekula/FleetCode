Public Class clsTankMaster

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Tank Master.xml", "VSP_FLT_TANKMSTR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TANKMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR_C0")

            Me.CellsMasking(objForm.UniqueID)

            objForm.Items.Item("36").TextStyle = 4

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            Me.CFLFilterItems(objForm.UniqueID, "CFL_MITMCD")
            Me.CFLFilterItems(objForm.UniqueID, "CFL_MITMNM")
            Me.CFLFilterCostCenter(objForm.UniqueID, "CFL_CC", "2")
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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR_C0")

            objMatrix = objForm.Items.Item("37").Specific
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSP_FLT_TANKMSTR"))

            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "37")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR_C0")

            objMatrix = objForm.Items.Item("37").Specific

            Select Case MatrixUID

                Case "37"
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_VSPITMCD", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPITMNM", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPCPTY", oDBs_Details.Offset, "")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("500").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("500").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("79").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("79").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

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

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "200" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim ChkItemExist As String = ""

                        If objMain.IsSAPHANA = True Then
                            ChkItemExist = "Select ""Code"" From ""@VSP_FLT_TANKMSTR"" Where ""U_VSPTNKNO"" ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"
                        Else
                            ChkItemExist = "Select Code From [@VSP_FLT_TANKMSTR] Where [U_VSPTNKNO] ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"
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

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TANKMSTR_C0")

                    objMatrix = objForm.Items.Item("37").Specific

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

                        If oCFL.UniqueID = "CFL_MITMCD" Or oCFL.UniqueID = "CFL_MITMNM" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPITMCD", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_VSPITMNM", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details.SetValue("U_VSPCPTY", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "37")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_CC" Then
                            oDBs_Head.SetValue("U_VSPTNKCC", oDBs_Head.Offset, oDT.GetValue("PrcCode", 0))
                        End If
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TANKMSTR" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLFilterItems(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            Dim GetItmGrp As String = ""
            If objMain.IsSAPHANA = True Then
                GetItmGrp = "Select ""ItmsGrpCod"" From OITM  Where ""ItmsGrpCod"" In (Select ""ItmsGrpCod"" From OITB Where ""U_VSPCSA"" ='Y') " & _
                                                                       "Group By  ""ItmsGrpCod"""
            Else
                GetItmGrp = "Select ItmsGrpCod From OITM  Where ItmsGrpCod In (Select ItmsGrpCod From OITB Where U_VSPCSA ='Y') " & _
                                                                       "Group By  ItmsGrpCod"
            End If
            Dim oRsGetItmGrp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetItmGrp.DoQuery(GetItmGrp)
            If oRsGetItmGrp.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "ItmsGrpCod"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetItmGrp.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetItmGrp.MoveNext()
                For i As Integer = 1 To oRsGetItmGrp.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "ItmsGrpCod"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetItmGrp.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetItmGrp.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "ItmsGrpCod"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterCostCenter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal DimCode As String)
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
            oCondition.Alias = "DimCode"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = DimCode
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objForm.Items.Item("200").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tank No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("8").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tank Name Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("10").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tank Model Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("64").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tank Type Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("6").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Capacity Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            ElseIf objForm.Items.Item("70").Specific.Selected.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Capacity U_oM Field Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("39").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Tank C.C Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Function
End Class
