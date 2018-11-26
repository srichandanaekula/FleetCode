Public Class clsUser
    Dim objForm As SAPbouiCOM.Form

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

            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbl_Chk", objForm.Items.Item("200000118").Top, objForm.Items.Item("18").Left + 20, 50, _
                                           "Restriction", "12")
            objMain.objUtilities.AddCheckBox(objForm.UniqueID, "Rst_Chk", objForm.Items.Item("200000118").Top, objForm.Items.Item("18").Left, _
                                              objForm.Items.Item("lbl_Chk").Width, "OUSR", "U_VSPRST", "Restriction")

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
