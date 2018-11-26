Public Class clsDeliveryConfirmation

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
#End Region
    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("DeliverConfirmation.xml", "VSP_DELVCONFR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_DELVCONFR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPDELCONF")

            'objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)



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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPDELCONF")
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSPODELCONF"))
            oDBs_Head.SetValue("U_VSPDOCDT", oDBs_Head.Offset, DateAndTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_VSPDOCST", oDBs_Head.Offset, "Open")
            Me.CellsMasking(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000002").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000002").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000006").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000006").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("28").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("30").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("30").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

           

          
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
   
    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPDELCONF")

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPDELCONF")


                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If oCFL.UniqueID = "CFL_DENT" And pVal.BeforeAction = True Then
                        Me.CFLFilterDelivery(objForm.UniqueID, oCFL.UniqueID)
                    End If
                    If oCFL.UniqueID = "CFL_DNO" And pVal.BeforeAction = True Then
                        Me.CFLFilterDelivery(objForm.UniqueID, oCFL.UniqueID)
                    End If


                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_DENT" Then
                            Try


                                objForm.Items.Item("102").Specific.Value = ""
                                objForm.Items.Item("4").Specific.Value = ""
                                Dim docnum As String = oDT.GetValue("DocNum", 0)
                                Dim docdt As Date = oDT.GetValue("DocDate", 0)

                                oDBs_Head.SetValue("U_VSPDLENT", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))
                                oDBs_Head.SetValue("U_VSPDLNO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                                oDBs_Head.SetValue("U_VSPDLDT", oDBs_Head.Offset, docdt.ToString("yyyyMMdd"))

                                Dim Getitemcode As String = ""

                                If objMain.IsSAPHANA = True Then

                                    Getitemcode = "Select T1.""ItemCode"",T1.""Dscription"" as ""ItemName"" from ODLN T2 inner join DLN1 T1 on T1.""DocEntry""=T2.""DocEntry"" where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' "
                                Else
                                    Getitemcode = "Select T1.""ItemCode"",T1.""Dscription"" as ""ItemName"" form ODLN T2 inner join DLN1 T1 on T1.""DocEntry""=T2.""DocEntry"" where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' "
                                End If
                                Dim oRsGetitemcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetitemcode.DoQuery(Getitemcode)

                                oDBs_Head.SetValue("U_VSPITCOD", oDBs_Head.Offset, oRsGetitemcode.Fields.Item("ItemCode").Value)
                                oDBs_Head.SetValue("U_VSPITNM", oDBs_Head.Offset, oRsGetitemcode.Fields.Item("ItemName").Value)

                                ''Quantity
                                Dim GetQuantity As String = ""

                                If objMain.IsSAPHANA = True Then

                                    GetQuantity = "SELECT Sum(T0.""Quantity"") as""Quantity"" FROM DLN1 T0  INNER JOIN ODLN T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""DocEntry"" ='562'"
                                Else
                                    GetQuantity = "SELECT Sum(T0.""Quantity"") as""Quantity"" FROM DLN1 T0  INNER JOIN ODLN T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""DocEntry"" ='562'"
                                End If
                                Dim oRsGetQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetQuantity.DoQuery(GetQuantity)


                                oDBs_Head.SetValue("U_VSPDLQTY", oDBs_Head.Offset, oRsGetQuantity.Fields.Item("Quantity").Value)

                                Dim gettolerance As String = ""
                                If objMain.IsSAPHANA = True Then

                                    gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & oRsGetitemcode.Fields.Item("ItemCode").Value & "' "
                                Else
                                    gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & oRsGetitemcode.Fields.Item("ItemCode").Value & "' "
                                End If
                                Dim oRsgettolerance As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsgettolerance.DoQuery(gettolerance)

                                oDBs_Head.SetValue("U_VSPTLQTY", oDBs_Head.Offset, oRsgettolerance.Fields.Item("U_VSPTLPRC").Value)

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try

                        End If

                        If oCFL.UniqueID = "CFL_DNO" Then

                            objForm.Items.Item("102").Specific.Value = ""
                            objForm.Items.Item("4").Specific.Value = ""

                            Dim docdt As Date = oDT.GetValue("DocDate", 0)

                            oDBs_Head.SetValue("U_VSPDLNO", oDBs_Head.Offset, oDT.GetValue("DocNum", 0))
                            oDBs_Head.SetValue("U_VSPDLENT", oDBs_Head.Offset, oDT.GetValue("DocEntry", 0))
                            oDBs_Head.SetValue("U_VSPDLDT", oDBs_Head.Offset, docdt.ToString("yyyyMMdd"))

                            Dim Getitemcode As String = ""

                            If objMain.IsSAPHANA = True Then

                                Getitemcode = "Select T1.""ItemCode"",T1.""Dscription"" as ""ItemName"" from ODLN T2 inner join DLN1 T1 on T1.""DocEntry""=T2.""DocEntry"" where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' "
                            Else
                                Getitemcode = "Select T1.""ItemCode"",T1.""Dscription"" as ""ItemName"" form ODLN T2 inner join DLN1 T1 on T1.""DocEntry""=T2.""DocEntry"" where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' "
                            End If
                            Dim oRsGetitemcode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetitemcode.DoQuery(Getitemcode)

                            oDBs_Head.SetValue("U_VSPITCOD", oDBs_Head.Offset, oRsGetitemcode.Fields.Item("ItemCode").Value)
                            oDBs_Head.SetValue("U_VSPITNM", oDBs_Head.Offset, oRsGetitemcode.Fields.Item("ItemName").Value)

                            ''Quantity
                            Dim GetQuantity As String = ""

                            If objMain.IsSAPHANA = True Then

                                GetQuantity = "SELECT Sum(T0.""Quantity"") as""Quantity"" FROM DLN1 T0  INNER JOIN ODLN T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""DocEntry"" ='562'"
                            Else
                                GetQuantity = "SELECT Sum(T0.""Quantity"") as""Quantity"" FROM DLN1 T0  INNER JOIN ODLN T1 ON T0.""DocEntry"" = T1.""DocEntry"" WHERE T0.""DocEntry"" ='562'"
                            End If
                            Dim oRsGetQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetQuantity.DoQuery(GetQuantity)


                            oDBs_Head.SetValue("U_VSPDLQTY", oDBs_Head.Offset, oRsGetQuantity.Fields.Item("Quantity").Value)

                            Dim gettolerance As String = ""
                            If objMain.IsSAPHANA = True Then

                                gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & oRsGetitemcode.Fields.Item("ItemCode").Value & "' "
                            Else
                                gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & oRsGetitemcode.Fields.Item("ItemCode").Value & "' "
                            End If
                            Dim oRsgettolerance As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsgettolerance.DoQuery(gettolerance)

                            oDBs_Head.SetValue("U_VSPTLQTY", oDBs_Head.Offset, oRsgettolerance.Fields.Item("U_VSPTLPRC").Value)


                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPDELCONF")

                    If pVal.ItemUID = "22" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("22").Specific.Value <> 0.0 Then
                            Dim delvQty As Double = 0.0
                            Dim actQty As Double = 0.0
                            Dim DifferenceQty As Double = 0.0
                            Dim ShortageQty As Double = 0.0
                            Dim ToleranceQty As Double = 0.0
                            Dim ActualToleranceQty As Double = 0.0

                            delvQty = objForm.Items.Item("18").Specific.Value
                            actQty = objForm.Items.Item("22").Specific.Value
                            ToleranceQty = objForm.Items.Item("28").Specific.Value
                            DifferenceQty = delvQty - actQty
                            oDBs_Head.SetValue("U_VSPDFQTY", oDBs_Head.Offset, CDbl(DifferenceQty))

                            ActualToleranceQty = delvQty * ToleranceQty / 100

                            If DifferenceQty > ActualToleranceQty Then
                                ShortageQty = DifferenceQty - ActualToleranceQty

                                oDBs_Head.SetValue("U_VSPSHQTY", oDBs_Head.Offset, CDbl(ShortageQty))
                            Else
                                oDBs_Head.SetValue("U_VSPSHQTY", oDBs_Head.Offset, 0.0)
                            End If
                        End If

                    End If



            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLFilterDelivery(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "DocStatus"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "O"
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_DEL_CONF" And pVal.BeforeAction = False Then

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
        objForm = objMain.objApplication.Forms.GetForm("VSP_DELVCONFR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = True Then
                        Dim cheDocumentExeist As String = ""
                        Dim docentry As String = objForm.Items.Item("102").Specific.Value
                        Dim docnum As String = objForm.Items.Item("4").Specific.Value


                        If objMain.IsSAPHANA = True Then
                            cheDocumentExeist = "Select DocNum,""U_VSPDLENT"",""U_VSPDLNO"" from ""@VSPDELCONF"" where ""U_VSPDLENT""='" & docentry & "' And ""U_VSPDLNO""='" & docnum & "' "
                        Else
                            cheDocumentExeist = "Select DocNum,""U_VSPDLENT"",""U_VSPDLNO"" from ""@VSPDELCONF"" where ""U_VSPDLENT""='" & docentry & "' And ""U_VSPDLNO""='" & docnum & "' "
                        End If
                        Dim oRsGetcheckContract As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetcheckContract.DoQuery(cheDocumentExeist)

                        If oRsGetcheckContract.RecordCount > 0 Then
                            BubbleEvent = False
                            objMain.objApplication.StatusBar.SetText("Delivery Number Alreday Exist", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Try
                        End If
                    End If

                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

End Class
