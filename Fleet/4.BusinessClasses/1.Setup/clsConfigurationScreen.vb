Public Class clsConfigurationScreen

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("ConfigurationSreen.xml", "VSP_FLT_CNFGSRN_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_CNFGSRN_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CNFGSRN")

            Dim Check As String = ""

            If objMain.IsSAPHANA = True Then
                Check = "Select ""DocNum"" From ""@VSP_FLT_CNFGSRN"""
            Else
                Check = "Select DocNum From [@VSP_FLT_CNFGSRN]"
            End If
            Dim oRsChec As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsChec.DoQuery(Check)
            If oRsChec.RecordCount = 0 Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
                oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, 1)
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("7").Specific.Value = 1
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Else
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objForm.Items.Item("7").Specific.Value = 1
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("13").Specific, "Select ""Code"" ,""Name"" From  OUDP ")
            objForm.Items.Item("13").DisplayDesc = True

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000005").Specific, "Select ""Code"" ,""Name"" From  OUDP ")
            objForm.Items.Item("1000005").DisplayDesc = True

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("18").Specific, "Select ""ItmsGrpCod"" , ""ItmsGrpNam"" From OITB ")
            objForm.Items.Item("18").DisplayDesc = True

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("20").Specific, "Select ""ItmsGrpCod"" , ""ItmsGrpNam"" From OITB ")
            objForm.Items.Item("20").DisplayDesc = True

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("22").Specific, "Select ""ItmsGrpCod"" , ""ItmsGrpNam"" From OITB ")
            objForm.Items.Item("22").DisplayDesc = True

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("26").Specific, "Select ""Code"" , ""Location"" From OLCT ")
            objForm.Items.Item("26").DisplayDesc = True

            Me.CFLCustomerFilter(objForm.UniqueID, "CFL_CUS")
            Me.CFLAccountFilter(objForm.UniqueID, "CFL_OACT")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_CNFGSRN" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CNFGSRN")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If oCFL.UniqueID = "CFL_VEN" And pVal.BeforeAction = True Then
                        Me.CFLFilterForPO(objForm.UniqueID, oCFL.UniqueID, "S")
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_WHS" Then
                            oDBs_Head.SetValue("U_VSPDEWHS", oDBs_Head.Offset, oDT.GetValue("WhsCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_TAXCD" Then
                            oDBs_Head.SetValue("U_VSPTXCD", oDBs_Head.Offset, oDT.GetValue("Code", 0))
                        End If

                        If oCFL.UniqueID = "CFL_CUS" Then
                            oDBs_Head.SetValue("U_VSPCUSCD", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_OACT" Then
                            oDBs_Head.SetValue("U_VSPOACT", oDBs_Head.Offset, oDT.GetValue("AcctCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_VEN" Then
                            oDBs_Head.SetValue("U_VSPFLVCD", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLCustomerFilter(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.CondVal = "C"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLAccountFilter(ByVal FormUID As String, ByVal CFL_ID As String)
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

    Sub CFLFilterForPO(ByVal FormUID As String, ByVal CFL_ID As String, ByVal CardType As String)
        Try

            'Dim getvendoe As String = CardCode
            'If getvendoe = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Configure Fuel Vendor Code in Configuration Screen")

            '    Exit Try
            'End If

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
            oCondition.CondVal = CardType
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
