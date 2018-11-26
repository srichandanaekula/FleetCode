Public Class clsSalesQuotation
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

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_FltSts", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
            '                            objform.Items.Item("70").Width, "Fleet Status", "70")
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_FltSts", objform.Items.Item("70").Top + 15, objform.Items.Item("14").Left, _
            '                               objform.Items.Item("14").Width, "OQUT", "U_VSPFLSTS", "lbl_FltSts")

            'objComboBox = objform.Items.Item("txt_FltSts").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open to Fleet", "Open to Fleet")
            'objComboBox.ValidValues.Add("Linked to Fleet", "Linked to Fleet")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
