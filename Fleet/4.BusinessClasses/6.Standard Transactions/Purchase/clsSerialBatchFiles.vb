Public Class clsSerialBatchFiles

    Dim objForm, objBatchFile As SAPbouiCOM.Form
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim objBatchFileMatrix, objMatrix As SAPbouiCOM.Matrix

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim ChkRestriction As String = ""
                        If objMain.IsSAPHANA = True Then
                            ChkRestriction = "Select ""U_VSPRST"" From OUSR Where  ""U_VSPRST"" = 'Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
                        Else
                            ChkRestriction = "Select U_VSPRST From OUSR Where  U_VSPRST = 'Y' And USER_CODE = '" & objMain.objCompany.UserName & "'"
                        End If
                        Dim oRsChkRestriction As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkRestriction.DoQuery(ChkRestriction)
                        If oRsChkRestriction.RecordCount > 0 Then
                            objBatchFile = objMain.objApplication.Forms.GetForm("41", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            objBatchFileMatrix = objBatchFile.Items.Item("3").Specific
                            objForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            objMatrix = objForm.Items.Item("38").Specific
                            objMatrix.Columns.Item("U_RATE").Cells.Item(1).Specific.Value = objBatchFileMatrix.Columns.Item("U_CenVat").Cells.Item(1).Specific.Value
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

End Class
