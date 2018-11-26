Public Class clsApInvoice

    Public objform As SAPbouiCOM.Form
    Dim objComboBox As SAPbouiCOM.ComboBox

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)

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

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Type", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "OPCH", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Number", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "OPCH", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("lbl_DocNum").Left, _
            '                objform.Items.Item("lbl_DocNum").Width, "Fleet Status", "lbl_DocNum")
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("14").Left, _
            '                               objform.Items.Item("14").Width, "OPCH", "U_VSPFLSTS", "lbl_FltSts")

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
        objform = objMain.objApplication.Forms.GetForm("141", objMain.objApplication.Forms.ActiveForm.TypeCount)

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMain.objDocumentType.UpdateDocument("OPCH", "A/P Invoice")
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

End Class




