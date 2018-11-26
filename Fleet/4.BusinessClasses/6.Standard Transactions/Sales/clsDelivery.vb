Public Class clsDelivery

    Public objform As SAPbouiCOM.Form
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText
    Dim oDBs_Head, oDBs_Detail1 As SAPbouiCOM.DBDataSource

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                        Me.CflAdding(objform.UniqueID)  ''Adding CFL to TextBox
                        Me.CflAdding1(objform.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objform.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If oCFL.UniqueID = "CFL_VHID" And pVal.BeforeAction = True Then
                        Me.CFLFilterForDriver(objform.UniqueID, oCFL.UniqueID)
                    End If
                    If oCFL.UniqueID = "CFL_VHID1" And pVal.BeforeAction = True Then
                        Me.CFLFilterForDriver(objform.UniqueID, oCFL.UniqueID)
                    End If



                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)



                        If oCFL.UniqueID = "CFL_VHID" Then
                            Try
                                objform.Items.Item("txt_Drv1").Specific.Value = oDT.GetValue("U_VSPFNAME", 0)
                            Catch ex As Exception
                            End Try

                        End If

                        If oCFL.UniqueID = "CFL_VHID1" Then
                            Try
                                objform.Items.Item("txt_Drv2").Specific.Value = oDT.GetValue("U_VSPFNAME", 0)
                            Catch ex As Exception
                            End Try
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    objMatrix = objform.Items.Item("38").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                        Dim ChkRestriction As String = ""

                        If objMain.IsSAPHANA = True Then
                            ChkRestriction = "Select ""U_VSPRST"" From OUSR Where  ""U_VSPRST"" = 'Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
                        Else
                            ChkRestriction = "Select U_VSPRST From OUSR Where  U_VSPRST = 'Y' And USER_CODE = '" & objMain.objCompany.UserName & "'"
                        End If
                        Dim oRsChkRestriction As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkRestriction.DoQuery(ChkRestriction)

                        If objform.Items.Item("txt_MODVAT").Specific.Value <> "" Then
                            If oRsChkRestriction.RecordCount > 0 And objform.Items.Item("txt_MODVAT").Specific.Value.Trim = "02" Then
                                If objMatrix.Columns.Item("U_RATE").Cells.Item(1).Specific.Value > 0.0 Then
                                    Dim Qty As Double = 0.0
                                    Dim UnitPrice As Double = 0.0
                                    Dim Rate As Double = 0.0

                                    If objMatrix.Columns.Item("11").Cells.Item(1).Specific.Value <> "" Then
                                        Qty = CDbl(objMatrix.Columns.Item("11").Cells.Item(1).Specific.Value)
                                    End If
                                    If objMatrix.Columns.Item("14").Cells.Item(1).Specific.Value <> "" Then
                                        Dim Price As String = objMatrix.Columns.Item("14").Cells.Item(1).Specific.Value()
                                        Price = Price.Replace("INR", "")
                                        UnitPrice = CDbl(Price)
                                    End If

                                    Rate = CDbl(objMatrix.Columns.Item("U_RATE").Cells.Item(1).Specific.Value)


                                    objMatrix.Columns.Item("21").Cells.Item(1).Specific.Value = (Qty * UnitPrice) + (Qty * Rate)
                                    Me.PostJE(objform.UniqueID, (Qty * UnitPrice) + (Qty * Rate))
                                End If
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And _
                                                                    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And objMain.FormCloseBoolean = True Then
                        objform.Close()
                        objMain.FormCloseBoolean = False
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Sub CFLFilterForDriver(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPSTS"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Active"
            oChooseFromList.SetConditions(oConditions)
           
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub PostJE(ByVal FormUID As String, ByVal Amount As String)
        Try






        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Type", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "ODLN", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Number", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "ODLN", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("lbl_DocNum").Left, _
            '                objform.Items.Item("lbl_DocNum").Width, "Fleet Status", "lbl_DocNum")
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("14").Left, _
            '                               objform.Items.Item("14").Width, "ODLN", "U_VSPFLSTS", "lbl_FltSts")

            'objComboBox = objform.Items.Item("txt_FltSts").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open to Fleet", "Open to Fleet")
            'objComboBox.ValidValues.Add("Linked to Fleet", "Linked to Fleet")


            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_MODVAT", objform.Items.Item("230").Top + 15, objform.Items.Item("230").Left, _
                                         objform.Items.Item("230").Width, "MODVAT", "230")
            objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_MODVAT", objform.Items.Item("222").Top + 15, objform.Items.Item("222").Left, _
                                            objform.Items.Item("222").Width, "ODLN", "U_MODVAT", "lbl_MODVAT")




            ''Add Driver1 and Driver2
            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_Drv1", objform.Items.Item("86").Top + 15, objform.Items.Item("86").Left, _
                                         objform.Items.Item("86").Width, "Driver 1", "86")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_Drv1", objform.Items.Item("86").Top + 15, objform.Items.Item("46").Left, _
                                            objform.Items.Item("46").Width, "ODLN", "U_VSPDRV1", "lbl_Drv1")

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_Drv2", objform.Items.Item("lbl_Drv1").Top + 15, objform.Items.Item("lbl_Drv1").Left, _
                                        objform.Items.Item("lbl_Drv1").Width, "Driver 2", "lbl_Drv1")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_Drv2", objform.Items.Item("lbl_Drv1").Top + 15, objform.Items.Item("txt_Drv1").Left, _
                                            objform.Items.Item("txt_Drv1").Width, "ODLN", "U_VSPDRV2", "lbl_Drv2")

           




        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objform = objMain.objApplication.Forms.GetForm("140", objMain.objApplication.Forms.ActiveForm.TypeCount)

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMain.objDocumentType.UpdateDocument("ODLN", "Delivery Order")
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

    Sub CflAdding(ByVal FormUID As String)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        oCFLs = objForm.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList

        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFLCreationParams.MultiSelection = False
        oCFLCreationParams.ObjectType = "VSP_FLT_ODRVRMSTR"
        oCFLCreationParams.UniqueID = "CFL_VHID"
        Try
            oCFL = oCFLs.Add(oCFLCreationParams)
            oEditText = objform.Items.Item("txt_Drv1").Specific
            oEditText.ChooseFromListUID = "CFL_VHID"
            oEditText.ChooseFromListAlias = "U_VSPFNAME"
        Catch ex As Exception
        End Try
    End Sub
    Sub CflAdding1(ByVal FormUID As String)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        oCFLs = objform.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList

        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFLCreationParams.MultiSelection = False
        oCFLCreationParams.ObjectType = "VSP_FLT_ODRVRMSTR"
        oCFLCreationParams.UniqueID = "CFL_VHID1"
        Try
            oCFL = oCFLs.Add(oCFLCreationParams)
            oEditText = objform.Items.Item("txt_Drv2").Specific
            oEditText.ChooseFromListUID = "CFL_VHID1"
            oEditText.ChooseFromListAlias = "U_VSPFNAME"
        Catch ex As Exception
        End Try
    End Sub

End Class
