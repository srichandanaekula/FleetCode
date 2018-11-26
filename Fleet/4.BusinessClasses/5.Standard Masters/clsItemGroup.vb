Public Class clsItemGroup
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim ButtonID As String
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objForm.UniqueID)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddCheckBox(objForm.UniqueID, "chkCSA", objForm.Items.Item("6").Top, objForm.Items.Item("6").Left + 200, _
                                               objForm.Items.Item("10002023").Width, "OITB", "U_VSPCSA", "C/S/A", 0, 0)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
