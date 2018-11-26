Public Class clsPurchaseOrder

    Public objform As SAPbouiCOM.Form
    Dim objComboBox As SAPbouiCOM.ComboBox
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

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Type", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "OPOR", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Number", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "OPOR", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("lbl_DocNum").Left, _
            '                                objform.Items.Item("lbl_DocNum").Width, "Fleet Status", "lbl_DocNum")
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("14").Left, _
            '                               objform.Items.Item("14").Width, "OPOR", "U_VSPFLSTS", "lbl_FltSts")

            'objComboBox = objform.Items.Item("txt_FltSts").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open to Fleet", "Open to Fleet")
            'objComboBox.ValidValues.Add("Linked to Fleet", "Linked to Fleet")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PostRevenueDetails(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objform.Freeze(True)
            oDBs_Head = objform.DataSources.DBDataSources.Item("OPOR")

            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0)

            Dim GetCCDetails As String = ""
            If objMain.IsSAPHANA = True Then
                GetCCDetails = "Select T0.""DocDate"",T0.""DocEntry"",T0.""U_TCCode"" ,T0.""U_TCRef"" ,Sum(T1.""Quantity"" * T1.""U_VSPUNPRC"") as ""Total"",T0.""Comments"" , " & _
          "T1.""OcrCode2"" as ""Vehicle CC"" From OPOR T0 Inner Join POR1 T1 On T0.""DocEntry"" = T1.""DocEntry"" Where T0.""DocEntry"" = '" & DocEntry & "' And " & _
          "((T1.""ItemCode"" In (Select ""U_VSPCITM"" From OADM)) Or (T1.""ItemCode"" In (Select ""U_VSPTITM""  From OADM))) Group By T0.""DocDate"",T0.""DocEntry"", " & _
          "T0.""U_TCCode"" ,T0.""U_TCRef"" ,T0.""Comments"",T1.""OcrCode2"""
            Else
                GetCCDetails = "Select T0.DocDate,T0.DocEntry,T0.U_TCCode ,T0.U_TCRef ,Sum(T1.Quantity * T1.U_VSPUNPRC) as Total,T0.Comments , " & _
          "T1.OcrCode2 as 'Vehicle CC' From OPOR T0 Inner Join POR1 T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = '" & DocEntry & "' And " & _
          "((T1.ItemCode In (Select U_VSPCITM From OADM)) Or (T1.ItemCode In (Select U_VSPTITM  From OADM))) Group By T0.DocDate,T0.DocEntry, " & _
          "T0.U_TCCode ,T0.U_TCRef ,T0.Comments,T1.OcrCode2"
            End If
            Dim oRsGetCCDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCCDetails.DoQuery(GetCCDetails)

            If oRsGetCCDetails.RecordCount > 0 Then
                For i As Integer = 1 To oRsGetCCDetails.RecordCount

                    Dim GetTripEntry As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetTripEntry = "Select Max(""DocEntry"") From ""@VSP_FLT_TRSHT"" Where ""U_VSPVHCL"" = (Select ""OcrName"" FRom OOCR Where " & _
                   """OcrCode"" = '" & oRsGetCCDetails.Fields.Item("Vehicle CC").Value & "') And ""U_VSPSTS"" = 'Open'"
                    Else
                        GetTripEntry = "Select Max(DocEntry) From [@VSP_FLT_TRSHT] Where U_VSPVHCL = (Select OcrName FRom OOCR Where " & _
                   "OcrCode = '" & oRsGetCCDetails.Fields.Item("Vehicle CC").Value & "') And U_VSPSTS = 'Open'"
                    End If
                    Dim oRsGetTripEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetTripEntry.DoQuery(GetTripEntry)

                    If oRsGetTripEntry.RecordCount > 0 Then
                        Dim GetRevenueLine As String = ""
                        If objMain.IsSAPHANA = True Then
                            GetRevenueLine = "Select Max(""LineId"")-1 From ""@VSP_FLT_TRSHT_C3"" Where " & _
                                                    """DocEntry"" = '" & oRsGetTripEntry.Fields.Item(0).Value & "'"
                        Else
                            GetRevenueLine = "Select Max(LineId)-1 From [@VSP_FLT_TRSHT_C3] Where " & _
                        "DocEntry = '" & oRsGetTripEntry.Fields.Item(0).Value & "'"
                        End If
                        Dim oRsGetRevenueLine As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetRevenueLine.DoQuery(GetRevenueLine)

                        Dim GetLineType As String = ""

                        If objMain.IsSAPHANA = True Then
                            GetLineType = "Select ""U_VSPTYPE"" From ""@VSP_FLT_TRSHT_C3"" Where " & _
                       """DocEntry"" = '" & oRsGetTripEntry.Fields.Item(0).Value & "' And ""LineId"" = '" & oRsGetRevenueLine.Fields.Item(0).Value + 1 & "'"
                        Else
                            GetLineType = "Select U_VSPTYPE From [@VSP_FLT_TRSHT_C3] Where " & _
                       "DocEntry = '" & oRsGetTripEntry.Fields.Item(0).Value & "' And LineId = '" & oRsGetRevenueLine.Fields.Item(0).Value + 1 & "'"
                        End If
                        Dim oRsGetLineType As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetLineType.DoQuery(GetLineType)

                        Dim doc As String = oRsGetTripEntry.Fields.Item(0).Value.ToString

                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oRsGetTripEntry.Fields.Item(0).Value.ToString)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C3")

                        If oRsGetLineType.Fields.Item(0).Value = "" Then
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPTYPE", "Purchase")
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPDOCTY", "Purchase Order")
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPGENTY", "Existing")
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPDATE", oRsGetCCDetails.Fields.Item("DocDate").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPDCNUM", oRsGetCCDetails.Fields.Item("DocEntry").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPBPCOD", oRsGetCCDetails.Fields.Item("U_TCCode").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPREF", oRsGetCCDetails.Fields.Item("U_TCRef").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPDCTOT", oRsGetCCDetails.Fields.Item("Total").Value.ToString)
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPREM", oRsGetCCDetails.Fields.Item("Comments").Value.ToString)
                            'objMain.oChildren.Item(CInt(oRsGetAdvancesLine.Fields.Item(0).Value)).SetProperty("U_VSPJENO", oRsGetPaymentDetails.Fields.Item("JE No.").Value.ToString)
                            'objMain.oGeneralService.Update(objMain.oGeneralData)
                        Else
                            objMain.oChild = objMain.oChildren.Add
                            objMain.oChild.SetProperty("U_VSPTYPE", "Purchase")
                            objMain.oChild.SetProperty("U_VSPDOCTY", "Purchase Order")
                            objMain.oChild.SetProperty("U_VSPGENTY", "Existing")
                            objMain.oChild.SetProperty("U_VSPDATE", oRsGetCCDetails.Fields.Item("DocDate").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPDCNUM", oRsGetCCDetails.Fields.Item("DocEntry").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPBPCOD", oRsGetCCDetails.Fields.Item("U_TCCode").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPREF", oRsGetCCDetails.Fields.Item("U_TCRef").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPDCTOT", oRsGetCCDetails.Fields.Item("Total").Value.ToString)
                            objMain.oChild.SetProperty("U_VSPREM", oRsGetCCDetails.Fields.Item("Comments").Value.ToString)
                            'objMain.oChild.SetProperty("U_VSPJENO", oRsGetPaymentDetails.Fields.Item("JE No.").Value.ToString)
                            'objMain.oGeneralService.Update(objMain.oGeneralData)
                        End If
                        'objMain.oChildren.Add()
                        objMain.oGeneralService.Update(objMain.oGeneralData)
                    End If
                    oRsGetCCDetails.MoveNext()
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
        objform = objMain.objApplication.Forms.GetForm("142", objMain.objApplication.Forms.ActiveForm.TypeCount)

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        objMain.objDocumentType.UpdateDocument("OPOR", "Purchase Order")
                        Me.PostRevenueDetails(objform.UniqueID)
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

End Class


