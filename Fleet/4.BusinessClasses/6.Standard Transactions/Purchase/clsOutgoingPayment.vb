Public Class clsOutgoingPayment

    Public objform As SAPbouiCOM.Form
    Public OpenedFromDocumentType As Boolean = False
    Public DocumentType As String = String.Empty
    Public DocNum As String = String.Empty
    Dim oDBs_Head As SAPbouiCOM.DBDataSource

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
                                            objform.Items.Item("152").Width, "OVPM", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("151").Left, _
                                          objform.Items.Item("151").Width, "Document Number", "lbl_DocTyp")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("152").Left, _
                                            objform.Items.Item("152").Width, "OVPM", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PostAdvanceDetails(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objform.Freeze(True)
            oDBs_Head = objform.DataSources.DBDataSources.Item("OVPM")

            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0)

            Dim GetPaymentDetails As String = ""

            If objMain.IsSAPHANA = True Then
                GetPaymentDetails = "Select T0.""DocDate"" ,T1.""SumApplied"" as ""Amount"" ,CASE When T0.""CashSum"" <> 0 Then T0.""CashAcct"" When " & _
           "T0.""TrsfrSum"" <> 0 Then T0.""TrsfrAcct"" When T0.""CreditSum"" <> 0 Then (Select ""CreditAcct"" From VPM3 Where ""DocNum"" = T0.""DocEntry"" ) When " & _
           "T0.""CheckSum"" <> 0 Then (Select ""CheckAct"" From VPM1 Where ""DocNum"" = T0.""DocEntry"") End as ""From Account"" ,T1.""AcctCode"" as ""To Account"" , " & _
           "T1.""OcrCode3"" as ""Driver CC"" ,T0.""DocEntry"" as ""Payment No"" ,T0.""TransId"" as ""JE No."" ,T1.""OcrCode2"" as ""Vehicle CC"" ,T0.""Comments"" From OVPM T0 Inner Join " & _
           "VPM4 T1 On T0.""DocEntry"" = T1.""DocNum"" Where T0.""DocType"" = 'A' And T0.""DocEntry"" = '" & DocEntry & "'"
            Else
                GetPaymentDetails = "Select T0.DocDate ,T1.SumApplied as 'Amount' ,CASE When T0.CashSum <> 0 Then T0.CashAcct When " & _
           "T0.TrsfrSum <> 0 Then T0.TrsfrAcct When T0.CreditSum <> 0 Then (Select CreditAcct From VPM3 Where DocNum = T0.DocEntry ) When " & _
           "T0.CheckSum <> 0 Then (Select CheckAct From VPM1 Where DocNum = T0.DocEntry) End as 'From Account' ,T1.AcctCode as 'To Account' , " & _
           "T1.OcrCode3 as 'Driver CC' ,T0.DocEntry as 'Payment No' ,T0.TransId as 'JE No.' ,T1.OcrCode2 as 'Vehicle CC' ,T0.Comments From OVPM T0 Inner Join " & _
           "VPM4 T1 On T0.DocEntry = T1.DocNum Where T0.DocType = 'A' And T0.DocEntry = '" & DocEntry & "'"
            End If
            Dim oRsGetPaymentDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetPaymentDetails.DoQuery(GetPaymentDetails)

            If oRsGetPaymentDetails.RecordCount > 0 Then
                For i As Integer = 1 To oRsGetPaymentDetails.RecordCount

                    Dim GetTripEntry As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetTripEntry = "Select Max(""DocEntry"") From ""@VSP_FLT_TRSHT"" Where ""U_VSPVHCL"" = (Select ""OcrName"" From OOCR Where " & _
                   """OcrCode"" = '" & oRsGetPaymentDetails.Fields.Item("Vehicle CC").Value & "') And ""U_VSPSTS"" = 'Open'"
                    Else
                        GetTripEntry = "Select Max(DocEntry) From [@VSP_FLT_TRSHT] Where U_VSPVHCL = (Select OcrName FRom OOCR Where " & _
                   "OcrCode = '" & oRsGetPaymentDetails.Fields.Item("Vehicle CC").Value & "') And U_VSPSTS = 'Open'"
                    End If
                    Dim oRsGetTripEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetTripEntry.DoQuery(GetTripEntry)

                    If oRsGetTripEntry.RecordCount > 0 Then
                        Dim GetAdvancesLine As String = ""
                        If objMain.IsSAPHANA = True Then
                            GetAdvancesLine = "Select Max(""LineId"")-1 ,""U_VSPADAMT"" From ""@VSP_FLT_TRSHT_C2"" Where " & _
                        """DocEntry"" = '" & oRsGetTripEntry.Fields.Item(0).Value & "' Group By ""U_VSPADAMT"""
                        Else
                            GetAdvancesLine = "Select Max(LineId)-1 ,U_VSPADAMT From [@VSP_FLT_TRSHT_C2] Where " & _
                        "DocEntry = '" & oRsGetTripEntry.Fields.Item(0).Value & "' Group By U_VSPADAMT"
                        End If
                        Dim oRsGetAdvancesLine As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetAdvancesLine.DoQuery(GetAdvancesLine)

                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oRsGetTripEntry.Fields.Item(0).Value.ToString)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C2")

                        If CInt(oRsGetAdvancesLine.Fields.Item(0).Value) = 0 And CDbl(oRsGetAdvancesLine.Fields.Item(1).Value) = 0 Then
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPJOUDT", oRsGetPaymentDetails.Fields.Item("DocDate").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPAMGOO", 0)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPADAMT", oRsGetPaymentDetails.Fields.Item("Amount").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPFRACT", oRsGetPaymentDetails.Fields.Item("From Account").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VAPTOACT", oRsGetPaymentDetails.Fields.Item("To Account").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPDRVCC", oRsGetPaymentDetails.Fields.Item("Driver CC").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VAPCAS", "")
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPCOM", oRsGetPaymentDetails.Fields.Item("Comments").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPOPNO", oRsGetPaymentDetails.Fields.Item("Payment No").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPJENO", oRsGetPaymentDetails.Fields.Item("JE No.").Value.ToString)
                            'objMain.oGeneralService.Update(objMain.oGeneralData)
                        Else
                            objMain.oChild = objMain.oChildren.Add
                            objMain.oChild.SetProperty("U_VSPJOUDT", oRsGetPaymentDetails.Fields.Item("DocDate").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPAMGOO", 0)
                            objMain.oChild.SetProperty("U_VSPADAMT", oRsGetPaymentDetails.Fields.Item("Amount").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPFRACT", oRsGetPaymentDetails.Fields.Item("From Account").Value.ToString)
                            objMain.oChild.SetProperty("U_VAPTOACT", oRsGetPaymentDetails.Fields.Item("To Account").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPDRVCC", oRsGetPaymentDetails.Fields.Item("Driver CC").Value.ToString)
                            objMain.oChild.SetProperty("U_VAPCAS", "")
                            objMain.oChild.SetProperty("U_VSPCOM", oRsGetPaymentDetails.Fields.Item("Comments").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPOPNO", oRsGetPaymentDetails.Fields.Item("Payment No").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPJENO", oRsGetPaymentDetails.Fields.Item("JE No.").Value.ToString)
                            'objMain.oGeneralService.Update(objMain.oGeneralData)
                        End If
                        'objMain.oChildren.Add()
                        objMain.oGeneralService.Update(objMain.oGeneralData)
                    End If
                        oRsGetPaymentDetails.MoveNext()
                Next
            End If

            objform.Freeze(False)
        Catch ex As Exception
            objform.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objform = objMain.objApplication.Forms.GetForm("426", objMain.objApplication.Forms.ActiveForm.TypeCount)

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
                        objMain.objDocumentType.UpdateDocument("OVPM", "Outgoing Payment")
                        Me.PostAdvanceDetails(objform.UniqueID)
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

End Class

