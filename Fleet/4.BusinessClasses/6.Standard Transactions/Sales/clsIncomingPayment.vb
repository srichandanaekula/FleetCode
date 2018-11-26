Public Class clsIncomingPayment

    Public objform As SAPbouiCOM.Form
    Public OpenedFromDocumentType As Boolean = False
    Public DocumentType As String = String.Empty
    Public DocNum As String = String.Empty

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

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("151").Top + 15, objform.Items.Item("151").Left, _
                                          objform.Items.Item("151").Width, "Document Type", "151")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocTyp", objform.Items.Item("151").Top + 15, objform.Items.Item("152").Left, _
                                            objform.Items.Item("152").Width, "ORCT", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("151").Left, _
                                          objform.Items.Item("151").Width, "Document Number", "lbl_DocTyp")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("152").Left, _
                                            objform.Items.Item("152").Width, "ORCT", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objform = objMain.objApplication.Forms.GetForm("170", objMain.objApplication.Forms.ActiveForm.TypeCount)

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = True Then
                        If OpenedFromDocumentType = True Then
                            OpenedFromDocumentType = False
                            objform.Items.Item("txt_DocTyp").Specific.Value = DocumentType
                            objform.Items.Item("txt_DocNum").Specific.Value = DocNum
                        End If
                    End If

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMain.objDocumentType.UpdateDocument("ORCT", "Incoming Payment")
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

End Class


