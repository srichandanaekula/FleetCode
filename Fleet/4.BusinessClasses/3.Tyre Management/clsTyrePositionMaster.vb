Public Class clsTyrePositionMaster

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Tyre Position Master.xml", "VSP_FLT_TPMSTR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TPMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TPMSTR")

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TPMSTR")
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSP_FLT_TPMSTR"))

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
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
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TPMSTR" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
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
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TPMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TPMSTR")

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = True Then
                        Dim tyreposition As String = objForm.Items.Item("6").Specific.Value.ToString.Trim
                        Try
                            Dim CheckConfig As String = "Select ""Code"",""U_VSPPSTN"" From ""@VSP_FLT_TPMSTR"" Where ""U_VSPPSTN"" = '" & tyreposition & "' "
                            Dim oRsCheckConfig As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckConfig.DoQuery(CheckConfig)
                            If oRsCheckConfig.RecordCount > 0 Then
                                objMain.objApplication.StatusBar.SetText("Tyre Position already exist ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oDBs_Head.SetValue("U_VSPPSTN", oDBs_Head.Offset, "")
                                oDBs_Head.SetValue("U_VSPNAM", oDBs_Head.Offset, "")
                                BubbleEvent = False
                                Exit Try
                            End If

                        Catch ex As Exception

                        End Try

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

            If objForm.Items.Item("6").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Position Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("8").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Name Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Function

End Class
