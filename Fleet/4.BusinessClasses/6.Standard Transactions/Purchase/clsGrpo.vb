Public Class clsGRPO
    Dim objForm, objBatchFile As SAPbouiCOM.Form
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim objBatchFileMatrix, objMatrix As SAPbouiCOM.Matrix

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                    End If
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("38").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim ChkRestriction As String = ""
                        If objMain.IsSAPHANA = True Then
                            ChkRestriction = "Select ""U_VSPRST"" From OUSR Where  ""U_VSPRST"" = 'Y' And ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
                        Else
                            ChkRestriction = "Select U_VSPRST From OUSR Where  U_VSPRST = 'Y' And USER_CODE = '" & objMain.objCompany.UserName & "'"
                        End If
                        Dim oRsChkRestriction As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkRestriction.DoQuery(ChkRestriction)
                        If objForm.Items.Item("txt_MODVAT").Specific.Value <> "" Then
                            If oRsChkRestriction.RecordCount > 0 And objForm.Items.Item("txt_MODVAT").Specific.Value.Trim = "02" Then
                                If objMatrix.Columns.Item("U_RATE").Cells.Item(1).Specific.Value > 0.0 Then
                                    Dim Qty As Double = 0.0
                                    Dim UnitPrice As Double = 0.0
                                    Dim Rate As Double = 0.0
                                    If objMatrix.Columns.Item("11").Cells.Item(1).Specific.Value <> "" Then
                                        Qty = CDbl(objMatrix.Columns.Item("11").Cells.Item(1).Specific.Value)
                                    End If
                                    If objMatrix.Columns.Item("14").Cells.Item(1).Specific.Value <> "" Then
                                        Dim Price As String = objMatrix.Columns.Item("14").Cells.Item(1).Specific.Value()
                                        Price = Price.Replace("INR", "")
                                        UnitPrice = CDbl(Price)
                                    End If
                                    Rate = CDbl(objMatrix.Columns.Item("U_RATE").Cells.Item(1).Specific.Value)
                                    objMatrix.Columns.Item("21").Cells.Item(1).Specific.Value = (Qty * UnitPrice) + (Qty * Rate)

                                    Me.PostJE(objForm.UniqueID, (Qty * UnitPrice) + (Qty * Rate))

                                End If
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And _
                                                                    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And objMain.FormCloseBoolean = True Then

                        objForm.Close()
                        objMain.FormCloseBoolean = False
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)

        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        Me.CreateVehicleMaster(objForm.UniqueID)
                        objMain.objDocumentType.UpdateDocument("OPDN", "GRPO")
                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

    Private Sub PostJE(ByVal FormUID As String, ByVal Amount As String)
        Try





        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub


    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Type", "70")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_DocTyp", objForm.Items.Item("70").Top + 15, objForm.Items.Item("14").Left, _
                                            objForm.Items.Item("14").Width, "OPDN", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Number", "70")
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_DocNum", objForm.Items.Item("lbl_DocTyp").Top + 15, objForm.Items.Item("14").Left, _
                                            objForm.Items.Item("14").Width, "OPDN", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objForm.UniqueID, "lbl_FltSts", objForm.Items.Item("lbl_DocNum").Top + 15, objForm.Items.Item("lbl_DocNum").Left, _
            '                objForm.Items.Item("lbl_DocNum").Width, "Fleet Status", "lbl_DocNum")
            'objMain.objUtilities.AddComboBox(objForm.UniqueID, "txt_FltSts", objForm.Items.Item("lbl_DocNum").Top + 15, objForm.Items.Item("14").Left, _
            '                               objForm.Items.Item("14").Width, "OPDN", "U_VSPFLSTS", "lbl_FltSts")

            'objComboBox = objForm.Items.Item("txt_FltSts").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open to Fleet", "Open to Fleet")
            'objComboBox.ValidValues.Add("Linked to Fleet", "Linked to Fleet")





            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbl_MODVAT", objForm.Items.Item("230").Top + 15, objForm.Items.Item("230").Left, _
                                         objForm.Items.Item("230").Width, "MODVAT", "230")
            objMain.objUtilities.AddComboBox(objForm.UniqueID, "txt_MODVAT", objForm.Items.Item("222").Top + 15, objForm.Items.Item("222").Left, _
                                            objForm.Items.Item("222").Width, "OPDN", "U_MODVAT", "lbl_MODVAT")

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub CreateVehicleMaster(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            'Creating Vehicle Master
            Dim GetGrpCod As String = ""
            GetGrpCod = "Select ""U_VSPVHGRP"" From ""@VSP_FLT_CNFGSRN"" "
            Dim oRsGetGrpCod As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetGrpCod.DoQuery(GetGrpCod)

            Dim GetGrpoDtls As String = ""

            If objMain.IsSAPHANA = True Then
                GetGrpoDtls = "Select T0.""ItemCode"" ,T0.""Dscription"" , T0.""Price"" , T1.""IntrSerial""   From PDN1 T0 Inner Join OSRI T1 On " & _
           "T1.""BaseEntry""  = T0.""DocEntry"" Inner Join OITM T2 On T0.""ItemCode"" = T2.""ItemCode"" And T0.""LineNum"" = T1.""BaseLinNum"" And T1.""BaseType"" = '20' And " & _
           "T2.""ItmsGrpCod"" = '" & oRsGetGrpCod.Fields.Item("U_VSPVHGRP").Value & "' And T1.""BaseEntry""  = (Select Max(""DocEntry"") From OPDN) "
            Else
                GetGrpoDtls = "Select T0.ItemCode ,T0.Dscription , T0.Price , T1.IntrSerial   From PDN1 T0 Inner Join OSRI T1 On " & _
           "T1.BaseEntry  = T0.DocEntry Inner Join OITM T2 On T0.ItemCode = T2.ItemCode And T0.LineNum = T1.BaseLinNum And T1.BaseType = '20' And " & _
           "T2.ItmsGrpCod = '" & oRsGetGrpCod.Fields.Item("U_VSPVHGRP").Value & "' And T1.BaseEntry  = (Select Max(DocEntry) From OPDN) "
            End If
            Dim oRsGetGrpoDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetGrpoDtls.DoQuery(GetGrpoDtls)

            If oRsGetGrpCod.RecordCount > 0 Then
                oRsGetGrpoDtls.MoveFirst()

                For i As Integer = 1 To oRsGetGrpoDtls.RecordCount
                    Dim Price As Double = oRsGetGrpoDtls.Fields.Item("Price").Value
                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
                    objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    objMain.oGeneralData.SetProperty("Code", objMain.objUtilities.getMaxCode("@VSP_FLT_VMSTR"))
                    objMain.oGeneralData.SetProperty("U_VSPVNM", oRsGetGrpoDtls.Fields.Item("Dscription").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPNO", oRsGetGrpoDtls.Fields.Item("IntrSerial").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPVAL", Price.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPODRDG", "0")
                    objMain.oGeneralData.SetProperty("U_VSPCALB", "N")
                    objMain.oGeneralData.SetProperty("U_VSPAVLB", "Y")

                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C0")
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C1")
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C2")
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C5")
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChild.SetProperty("U_VSPTODT", "9999-12-31")
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_VMSTR_C6")
                    objMain.oChild = objMain.oChildren.Add

                    objMain.oGeneralService.Add(objMain.oGeneralData)
                    oRsGetGrpoDtls.MoveNext()
                Next
            End If

            'Updating Serial No. in Tyre Master
            GetGrpCod = "Select ""U_VSPTYRGR"" From ""@VSP_FLT_CNFGSRN"" "
            oRsGetGrpCod.DoQuery(GetGrpCod)

            Dim GetSerialNo As String = ""

            If objMain.IsSAPHANA = True Then

                GetSerialNo = "Select T0.""ItemCode"" ,T0.""Dscription"" , T0.""Price"" , T1.""IntrSerial"" , T0.""TaxCode"" , T0.""WhsCode""  From PDN1 T0 Inner Join OSRI T1 On " & _
            "T1.""BaseEntry""  = T0.""DocEntry"" Inner Join OITM T2 On T0.""ItemCode"" = T2.""ItemCode"" And T0.""LineNum"" = T1.""BaseLinNum"" And T1.""BaseType"" = '20' And " & _
            "T2.""ItmsGrpCod"" = '" & oRsGetGrpCod.Fields.Item("U_VSPTYRGR").Value & "' And T1.""BaseEntry""  = (Select Max(""DocEntry"") From OPDN) "
            Else
                GetSerialNo = "Select T0.ItemCode ,T0.Dscription , T0.Price , T1.IntrSerial , T0.TaxCode , T0.WhsCode  From PDN1 T0 Inner Join OSRI T1 On " & _
            "T1.BaseEntry  = T0.DocEntry Inner Join OITM T2 On T0.ItemCode = T2.ItemCode And T0.LineNum = T1.BaseLinNum And T1.BaseType = '20' And " & _
            "T2.ItmsGrpCod = '" & oRsGetGrpCod.Fields.Item("U_VSPTYRGR").Value & "' And T1.BaseEntry  = (Select Max(DocEntry) From OPDN) "
            End If
            Dim oRsGetSerialNo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetSerialNo.DoQuery(GetSerialNo)

            If oRsGetSerialNo.RecordCount > 0 Then
                oRsGetSerialNo.MoveFirst()

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMSTR")
                objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                For i As Integer = 1 To oRsGetSerialNo.RecordCount
                    objMain.oGeneralData.SetProperty("Code", objMain.objUtilities.getMaxCode("@VSP_FLT_TYRMSTR"))
                    objMain.oGeneralData.SetProperty("U_VSPTRNUM", oRsGetSerialNo.Fields.Item("IntrSerial").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSRLNO", oRsGetSerialNo.Fields.Item("IntrSerial").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPNO", objForm.Items.Item("8").Specific.Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPCHFM", objForm.Items.Item("4").Specific.Value.ToString)

                    Dim PurDate As String = objForm.Items.Item("10").Specific.Value
                    PurDate = PurDate.Insert("4", "-")
                    PurDate = PurDate.Insert("7", "-")

                    objMain.oGeneralData.SetProperty("U_VSPPCHON", PurDate)
                    objMain.oGeneralData.SetProperty("U_VSPITCD", oRsGetSerialNo.Fields.Item("ItemCode").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPPRIC", oRsGetSerialNo.Fields.Item("Price").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPTXCD", oRsGetSerialNo.Fields.Item("TaxCode").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPSLOC", oRsGetSerialNo.Fields.Item("WhsCode").Value.ToString)
                    objMain.oGeneralData.SetProperty("U_VSPKMRUN", "0")
                    objMain.oGeneralService.Add(objMain.oGeneralData)

                    oRsGetSerialNo.MoveNext()
                Next
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
