Public Class clsArInvoice

    Public objform As SAPbouiCOM.Form
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If Fleet.MainCls.ohtLookUpForm.ContainsKey(objform.UniqueID) = True And pVal.BeforeAction = False Then
                        Fleet.MainCls.ohtLookUpForm.Remove(objform.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objform.Items.Item("38").Specific
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                                                            pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then

                        If Me.Validation(objform.UniqueID) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And _
                                                                    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And objMain.FormCloseBoolean = True Then
                        objform.Close()
                        objMain.FormCloseBoolean = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objform.Items.Item("38").Specific
                    Try
                        objform.Freeze(True)
                        If pVal.ItemUID = "38" And pVal.ColUID = "U_VSPADQTY" And pVal.BeforeAction = False Then
                            If objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(pVal.Row).Specific.Value <> 0.0 Then

                                Dim ActualDelvQty As Double = objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(pVal.Row).Specific.Value
                                Dim Quantity As Double = objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value

                                Dim checkNonInventoryServiceItem As String = ""
                                If objMain.IsSAPHANA = True Then

                                    checkNonInventoryServiceItem = "SELECT T0.""InvntItem"",T0.""ItemClass"" FROM OITM T0 WHERE T0.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value) & "' "
                                Else
                                    checkNonInventoryServiceItem = "SELECT T0.""InvntItem"",T0.""ItemClass"" FROM OITM T0 WHERE T0.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value) & "' "
                                End If
                                Dim oRscheckNonInventoryServiceItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRscheckNonInventoryServiceItem.DoQuery(checkNonInventoryServiceItem)

                                If oRscheckNonInventoryServiceItem.Fields.Item("InvntItem").Value = "N" Then
                                    If oRscheckNonInventoryServiceItem.Fields.Item("ItemClass").Value = "2" Then

                                        If Quantity = 0.0 Then
                                            objMain.objApplication.StatusBar.SetText("Please Enter Quantity For The ItemCode : " + CStr(objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value) + " in Row Level " + CStr(pVal.Row), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            objform.Freeze(False)
                                            Exit Try
                                        End If

                                If CInt(Quantity) < CInt(ActualDelvQty) Then

                                            objMain.objApplication.StatusBar.SetText("Actual Deleivered Quantity Should Not Be Gratter Than From Quantity In Row Level " + CStr(pVal.Row), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                    objMatrix.Columns.Item("U_VSPSTQTY").Editable = True
                                    objMatrix.Columns.Item("U_VSPDFQTY").Editable = True
                                    objMatrix.Columns.Item("U_VSPTLQTY").Editable = True
                                    objMatrix.Columns.Item("U_VSPTLQTY").Cells.Item(pVal.Row).Specific.Value = 0.0
                                    objMatrix.Columns.Item("U_VSPDFQTY").Cells.Item(pVal.Row).Specific.Value = 0.0
                                    objMatrix.Columns.Item("U_VSPSTQTY").Cells.Item(pVal.Row).Specific.Value = 0.0

                                    objMatrix.Columns.Item("U_VSPTLQTY").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                    objMatrix.Columns.Item("U_VSPSTQTY").Editable = False
                                    objMatrix.Columns.Item("U_VSPDFQTY").Editable = False
                                    objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objMatrix.Columns.Item("U_VSPTLQTY").Editable = False
                                    objform.Freeze(False)
                                    Exit Try
                                End If

                                objMatrix.Columns.Item("U_VSPSTQTY").Editable = True
                                objMatrix.Columns.Item("U_VSPDFQTY").Editable = True
                                objMatrix.Columns.Item("U_VSPTLQTY").Editable = True

                                Dim delvQty As Double = 0.0
                                Dim actDelvQty As Double = 0.0
                                Dim DifferenceQty As Double = 0.0
                                Dim ShortageQty As Double = 0.0
                                Dim TolerancePer As Double = 0.0
                                Dim ActualToleranceQty As Double = 0.0

                                Dim itemcode As String = objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value

                                objMatrix.Columns.Item("U_VSPTLQTY").Cells.Item(pVal.Row).Specific.Value = 0.0
                                objMatrix.Columns.Item("U_VSPDFQTY").Cells.Item(pVal.Row).Specific.Value = 0.0
                                objMatrix.Columns.Item("U_VSPSTQTY").Cells.Item(pVal.Row).Specific.Value = 0.0

                                Dim gettolerance As String = ""
                                If itemcode = "" Then
                                    Exit Try
                                End If

                                'SELECT T0."ItemClass" FROM OITM T0 WHERE T0."ItemCode" ='Tets2'
                                If objMain.IsSAPHANA = True Then

                                            gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & itemcode & "' "
                                Else
                                    gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & itemcode & "' "
                                End If
                                Dim oRsgettolerance As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsgettolerance.DoQuery(gettolerance)
                                Dim tolerance As String = oRsgettolerance.Fields.Item("U_VSPTLPRC").Value

                                If CStr(oRsgettolerance.Fields.Item("U_VSPTLPRC").Value) = 0 Then
                                    objMain.objApplication.StatusBar.SetText("Please Enter Tolerance Percentage in Item Master Data for the Item Code : " + itemcode + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    objform.Freeze(False)
                                    objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objMatrix.Columns.Item("U_VSPSTQTY").Editable = False
                                    objMatrix.Columns.Item("U_VSPDFQTY").Editable = False
                                    objMatrix.Columns.Item("U_VSPTLQTY").Editable = False
                                    Exit Try
                                End If

                                objMatrix.Columns.Item("U_VSPSTQTY").Editable = True
                                objMatrix.Columns.Item("U_VSPDFQTY").Editable = True
                                objMatrix.Columns.Item("U_VSPTLQTY").Editable = True

                                delvQty = objMatrix.Columns.Item("11").Cells.Item(pVal.Row).Specific.Value
                                actDelvQty = objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(pVal.Row).Specific.Value
                                TolerancePer = CDbl(oRsgettolerance.Fields.Item("U_VSPTLPRC").Value)
                                DifferenceQty = delvQty - actDelvQty
                                objMatrix.Columns.Item("U_VSPDFQTY").Cells.Item(pVal.Row).Specific.Value = CDbl(DifferenceQty)

                                ActualToleranceQty = delvQty * TolerancePer / 100

                                objMatrix.Columns.Item("U_VSPTLQTY").Cells.Item(pVal.Row).Specific.Value = ActualToleranceQty

                                If DifferenceQty > ActualToleranceQty Then
                                    ShortageQty = DifferenceQty - ActualToleranceQty
                                    objMatrix.Columns.Item("U_VSPSTQTY").Cells.Item(pVal.Row).Specific.Value = CDbl(ShortageQty)

                                Else
                                    objMatrix.Columns.Item("U_VSPSTQTY").Cells.Item(pVal.Row).Specific.Value = 0.0

                                End If

                                Try
                                    objMatrix.Columns.Item("U_VSPTLQTY").Editable = False

                                Catch ex As Exception
                                    objform.Freeze(False)
                                End Try
                                Try
                                    objMatrix.Columns.Item("U_VSPDFQTY").Editable = False
                                    objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Catch ex As Exception
                                    objform.Freeze(False)
                                End Try
                                Try
                                    objMatrix.Columns.Item("U_VSPSTQTY").Editable = False
                                    objMatrix.Columns.Item("1").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objform.Freeze(False)
                                Catch ex As Exception
                                    objform.Freeze(False)
                                End Try
                                objform.Freeze(False)
                            End If

                                End If

                            End If
                        End If


                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                        objform.Freeze(False)
                    End Try

                    objform.Freeze(False)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            objform.Freeze(False)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objform.Items.Item("38").Specific

            For i As Integer = 1 To objMatrix.VisualRowCount


                Dim checkNonInventoryServiceItem As String = ""
                If objMain.IsSAPHANA = True Then

                    checkNonInventoryServiceItem = "SELECT T0.""InvntItem"",T0.""ItemClass"" FROM OITM T0 WHERE T0.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value) & "' "
                Else
                    checkNonInventoryServiceItem = "SELECT T0.""InvntItem"",T0.""ItemClass"" FROM OITM T0 WHERE T0.""ItemCode"" ='" & CStr(objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value) & "' "
                End If
                Dim oRscheckNonInventoryServiceItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRscheckNonInventoryServiceItem.DoQuery(checkNonInventoryServiceItem)

                If oRscheckNonInventoryServiceItem.Fields.Item("InvntItem").Value = "N" Then
                    If oRscheckNonInventoryServiceItem.Fields.Item("ItemClass").Value = "2" Then

                        If objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(i).Specific.Value = 0.0 And objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
                            objMain.objApplication.StatusBar.SetText("Actual Deliverd Quantity is Empty in Line : " & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False

                        ElseIf objMatrix.Columns.Item("U_VSPTLQTY").Cells.Item(i).Specific.Value = 0.0 And objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value <> "" Then
                            objMain.objApplication.StatusBar.SetText("Tolerance Quantity is Empty in Line : " & i & "", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If

                    End If
                End If

            Next

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Type", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "OINV", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Number", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "OINV", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("lbl_DocNum").Left, _
            '                objform.Items.Item("lbl_DocNum").Width, "Fleet Status", "lbl_DocNum")
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("14").Left, _
            '                               objform.Items.Item("14").Width, "OINV", "U_VSPFLSTS", "lbl_FltSts")

            'objComboBox = objform.Items.Item("txt_FltSts").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open to Fleet", "Open to Fleet")
            'objComboBox.ValidValues.Add("Linked to Fleet", "Linked to Fleet")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objform = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)
        objMatrix = objform.Items.Item("38").Specific

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    'If BusinessObjectInfo.BeforeAction = True Then
                    '    For i As Integer = 1 To objMatrix.VisualRowCount
                    '        Dim Quantity As Double = 0.0
                    '        Dim Quantity1 As Double = 0.0
                    '        Dim itemcode As String = objMatrix.Columns.Item("1").Cells.Item(i).Specific.Value

                    '        If itemcode <> "" Then
                    '            Quantity = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                    '            Dim ActualDeliverQuantity As Double = objMatrix.Columns.Item("U_VSPADQTY").Cells.Item(i).Specific.Value
                    '            objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value = CDbl(ActualDeliverQuantity)
                    '            Quantity1 = objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                    '            If Quantity1 <> ActualDeliverQuantity Then
                    '                objMain.objApplication.StatusBar.SetText("Quantity Is Not Equal to Actual Deliver Quantity")
                    '                BubbleEvent = False
                    '                Exit Try
                    '            End If

                    '        End If


                    '    Next

                    'End If

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMain.objDocumentType.UpdateDocument("OINV", "A/R Invoice")
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

End Class
