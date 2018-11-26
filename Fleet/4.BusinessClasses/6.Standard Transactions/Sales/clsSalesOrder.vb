Public Class clsSalesOrder

    Public objform As SAPbouiCOM.Form
    Dim objComboBox, objComboBox1 As SAPbouiCOM.ComboBox
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim oEditText As SAPbouiCOM.EditText

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If Fleet.MainCls.ohtLookUpForm.ContainsKey(objform.UniqueID) = True And pVal.BeforeAction = False Then
                        Fleet.MainCls.ohtLookUpForm.Remove(objform.UniqueID)
                    End If
                    '-----------------------------------------------------

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objform.UniqueID)
                        Me.CflAdding(objform.UniqueID)
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objform.Items.Item("txt_PoNo").Enabled = True
                        objform.Items.Item("txt_PoEnt").Enabled = True
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Try
                            objform.Items.Item("txt_PoEnt").Enabled = False
                        Catch ex As Exception

                        End Try

                        Dim status As String = objform.Items.Item("81").Specific.Value
                        If status = "3" Then
                            Try
                                objform.Items.Item("txt_PoNo").Enabled = False
                                objform.Items.Item("txt_PoEnt").Enabled = False
                                objform.Items.Item("txt_VLQTY").Enabled = False
                            Catch ex As Exception

                            End Try

                        ElseIf status = "1" Then
                            Try
                                objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform.Items.Item("txt_PoNo").Enabled = True
                                objform.Items.Item("txt_PoEnt").Enabled = True
                                objform.Items.Item("txt_VLQTY").Enabled = True
                                objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Try
                                    objform.Items.Item("txt_PoEnt").Enabled = False
                                Catch ex As Exception

                                End Try

                            Catch ex As Exception

                            End Try

                            'objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    Try

                        If pVal.ItemUID = "txt_OrTyp" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            Dim ortyp As String = objform.Items.Item("txt_OrTyp").Specific.Selected.Value.ToString.Trim
                            If objform.Items.Item("txt_OrTyp").Specific.Selected.Value.ToString.Trim = "TR" Or objform.Items.Item("txt_OrTyp").Specific.Selected.Value.ToString.Trim = "Other" Then
                                objform.Items.Item("txt_PoNo").Specific.Value = ""
                                objform.Items.Item("txt_PoEnt").Enabled = True
                                objform.Items.Item("txt_PoEnt").Specific.Value = ""
                                objform.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform.Items.Item("txt_PoEnt").Enabled = False
                            End If
                        End If


                    Catch ex As Exception

                    End Try


                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                    objform = objMain.objApplication.Forms.Item(FormUID)
                    Try
                        If pVal.ItemUID = "txt_PoNo" And pVal.BeforeAction = False And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            Dim PoNumber As String = CStr(objform.Items.Item("txt_PoNo").Specific.Value).Trim
                            If PoNumber = "" Then
                                objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform.Items.Item("txt_PoEnt").Enabled = True
                                objform.Items.Item("txt_PoEnt").Specific.Value = ""
                                objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform.Items.Item("txt_PoEnt").Enabled = False
                            End If
                        End If
                    Catch ex As Exception

                    End Try


                    'objform = objMain.objApplication.Forms.Item(FormUID)
                    'Try
                    '    Dim PoNumber As String = CStr(objform.Items.Item("txt_PoNo").Specific.Value).Trim

                    '    If pVal.ItemUID = "txt_PoNo" And pVal.BeforeAction = False And objform.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '        If PoNumber <> "" And CStr(objform.Items.Item("txt_OrTyp").Specific.Selected.Value).Trim = "MT" Then
                    '            Dim checkPOINSaleOrderExist As String = ""
                    '            If objMain.IsSAPHANA = True Then
                    '                checkPOINSaleOrderExist = "Select ""U_VSPPONUM"",""U_VSPPOENT"" from ORDR where ""U_VSPPONUM""='" & PoNumber & "'  "
                    '            Else
                    '                checkPOINSaleOrderExist = "Select ""U_VSPPONUM"",""U_VSPPOENT"" from ORDR where ""U_VSPPONUM""='" & PoNumber & "'  "
                    '            End If
                    '            Dim oRscheckPOINSaleOrderExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '            oRscheckPOINSaleOrderExist.DoQuery(checkPOINSaleOrderExist)

                    '            If oRscheckPOINSaleOrderExist.RecordCount > 0 Then
                    '                objMain.objApplication.StatusBar.SetText("This PO Number Is Alreday Assign to Some Other Sales Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    '                Exit Try
                    '            End If
                    '        End If

                    '        If PoNumber = "" Then
                    '            objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '            objform.Items.Item("txt_PoEnt").Enabled = True
                    '            objform.Items.Item("txt_PoEnt").Specific.Value = ""
                    '            objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '            objform.Items.Item("txt_PoEnt").Enabled = False
                    '        End If
                    '    End If


                    'Catch ex As Exception

                    'End Try

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objform = objMain.objApplication.Forms.Item(FormUID)

                    Try

                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim CFL_Id As String
                        CFL_Id = CFLEvent.ChooseFromListUID
                        oCFL = objform.ChooseFromLists.Item(CFL_Id)
                        Dim oDT As SAPbouiCOM.DataTable
                        oDT = CFLEvent.SelectedObjects
                        objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        Dim getFuelVendorCode As String = ""
                        If objMain.IsSAPHANA = True Then
                            getFuelVendorCode = "Select ""U_VSPFLVCD"" from ""@VSP_FLT_CNFGSRN"""
                        Else
                            getFuelVendorCode = "Select ""U_VSPFLVCD"" from ""@VSP_FLT_CNFGSRN"""
                        End If

                        Dim orsgetVendorCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsgetVendorCode.DoQuery(getFuelVendorCode)

                        Dim getvendoe As String = orsgetVendorCode.Fields.Item(0).Value


                        If oCFL.UniqueID = "CFL_PONO" And pVal.BeforeAction = True Then
                            If getvendoe = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Configure Fuel Vendor Code in Configuration Screen")
                                BubbleEvent = False
                                Exit Try
                            End If
                            Me.CFLFilterForPO(objform.UniqueID, oCFL.UniqueID, CStr(orsgetVendorCode.Fields.Item(0).Value))
                        End If
                        'If oCFL.UniqueID = "CFL_POENT" And pVal.BeforeAction = True Then
                        '    Me.CFLFilterForPO(objform.UniqueID, oCFL.UniqueID, CStr(orsgetVendorCode.Fields.Item(0).Value))
                        'End If



                        If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                            If objform.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)



                            If oCFL.UniqueID = "CFL_PONO" Then

                                Try
                                    objform.Items.Item("txt_PoNo").Specific.Value = oDT.GetValue("DocNum", 0)
                                Catch ex As Exception

                                End Try
                                Try
                                    objform.Items.Item("txt_PoEnt").Enabled = True
                                    objform.Items.Item("txt_PoEnt").Specific.Value = oDT.GetValue("DocEntry", 0)
                                    objform.Items.Item("txt_PoNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    objform.Items.Item("txt_PoEnt").Enabled = False
                                Catch ex As Exception

                                End Try


                            End If

                        End If

                    Catch ex As Exception

                    End Try

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objform.Items.Item("38").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And _
                                                                    pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And objMain.FormCloseBoolean = True Then
                        objform.Close()
                        objMain.FormCloseBoolean = False
                    End If

                    'If pVal.ItemUID = "38" And pVal.ColUID = "160" And pVal.BeforeAction = False Then

                    'End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub TaxCalculation(ByVal FormUID As String, ByVal TaxCode As String, ByVal Row As Integer)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objform.Items.Item("38").Specific

            Dim CCVDTaxAmt, ACVDTaxAmt, VATTaxAmt, CSTTaxAmt, BCDTaxAmt, C_CVDTaxAmt, C_CessTaxAmt, A_CVDTaxAmt As Double
            Dim CCVDBaseAmt, ACVDBaseAmt, VATBaseAmt, CSTBaseAmt, BCDBaseAmt, C_CVDBaseAmt, C_CessBaseAmt, A_CVDBaseAmt As Double
            Dim Weight As Double = CDbl(objMatrix.Columns.Item("58").Cells.Item(Row).Specific.Value)
            Dim Qty As Double = CDbl(objMatrix.Columns.Item("11").Cells.Item(Row).Specific.Value)
            Dim LineTotal As Double = CDbl(objMatrix.Columns.Item("21").Cells.Item(Row).Specific.Value)

            Dim GetTaxRate As String = ""
            If objMain.IsSAPHANA = True Then
                GetTaxRate = "Select ""Rate"" From OSTC Where ""Code"" = '" & TaxCode & "'"
            Else
                GetTaxRate = "Select Rate From OSTC Where Code = '" & TaxCode & "'"
            End If
            Dim oRsGetTaxRate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTaxRate.DoQuery(GetTaxRate)

            Dim TaxRate As Double = oRsGetTaxRate.Fields.Item(0).Value

            Select Case TaxCode

                Case "CCVD"
                    'if (weight >0) {CCVD_BaseAmt = (weight*0.001*qty)} else {CCVD_BaseAmt = Total} CCVD_TaxAmt = CCVD_BaseAmt * CCVD_Rate

                    If Weight > 0 Then
                        CCVDBaseAmt = Weight * 0.001 * Qty
                    Else
                        CCVDBaseAmt = LineTotal
                    End If
                    'CCVDTaxAmt = CCVDBaseAmt * CCVDRate

                Case "ACVD@5"
                    'ACVD_BaseAmt = CCVD_BaseAmt+CCVD_TaxAmt+(((CCVD_BaseAmt*5)/105)+CCVD_TaxAmt)*0.03 ACVD_TaxAmt = ACVD_BaseAmt * ACVD_Rate

                    If Weight > 0 Then
                        CCVDBaseAmt = Weight * 0.001 * Qty
                    Else
                        CCVDBaseAmt = LineTotal
                    End If
                    'CCVDTaxAmt = CCVDBaseAmt * CCVDRate

                    ACVDBaseAmt = CCVDBaseAmt + CCVDTaxAmt + (((CCVDBaseAmt * 5) / 105) + CCVDTaxAmt) * 0.03
                    'ACVDTaxAmt = ACVDBaseAmt * ACVDRate

                Case "VAT@CVD"
                    'VAT_BaseAmt = Total+CCVD_TaxAmt+ACVD_TaxAmt VAT_TaxAmt = VAT_BaseAmt * VAT_Rate
                    If Weight > 0 Then
                        CCVDBaseAmt = Weight * 0.001 * Qty
                    Else
                        CCVDBaseAmt = LineTotal
                    End If
                    'CCVDTaxAmt = CCVDBaseAmt * CCVDRate

                    VATBaseAmt = LineTotal + CCVDTaxAmt + ACVDTaxAmt
                    'VATTaxAmt = VATBaseAmt * VATRate

                Case "CST@CVD"
                    'CST_BaseAmt = Total+CCVD_TaxAmt+ACVD_TaxAmt CST_TaxAmt = CST_BaseAmt * CST_Rate

                    If Weight > 0 Then
                        CCVDBaseAmt = Weight * 0.001 * Qty
                    Else
                        CCVDBaseAmt = LineTotal
                    End If
                    'CCVDTaxAmt = CCVDBaseAmt * CCVDRate

                    CSTBaseAmt = LineTotal + CCVDTaxAmt + ACVDTaxAmt
                    'CSTTaxAmt = CSTBaseAmt * CSTRate

                Case "BCD"
                    'if (weight > 0) { BCD_BaseAmt = weight*0.001*qty } else { BCD_BaseAmt = Total } BCD_TaxAmt=BCD_BaseAmt*BCD_Rate
                    If (Weight > 0) Then
                        BCDBaseAmt = Weight * 0.001 * Qty
                    Else
                        BCDBaseAmt = LineTotal
                    End If
                    'BCDTaxAmt = BCDBaseAmt * BCDRate

                Case "C_CVD@BCD"
                    'C_CVD_BaseAmt = BCD_BaseAmt+BCD_TaxAmt C_CVD_TaxAmt = C_CVD_BaseAmt * C_CVD_Rate
                    If (Weight > 0) Then
                        BCDBaseAmt = Weight * 0.001 * Qty
                    Else
                        BCDBaseAmt = LineTotal
                    End If
                    'BCDTaxAmt = BCDBaseAmt * BCDRate

                    C_CVDBaseAmt = BCDBaseAmt + BCDTaxAmt
                    'C_CVDTaxAmt = C_CVDBaseAmt * C_CVDRate

                Case "C_Cess@BCD"
                    'C_Cess_BaseAmt = BCD_TaxAmt+C_CVD_TaxAmt C_Cess_TaxAmt = C_Cess_BaseAmt * C_Cess_Rate
                    If (Weight > 0) Then
                        BCDBaseAmt = Weight * 0.001 * Qty
                    Else
                        BCDBaseAmt = LineTotal
                    End If
                    'BCDTaxAmt = BCDBaseAmt * BCDRate

                    C_CessBaseAmt = BCDTaxAmt + C_CVDTaxAmt
                    'C_CessTaxAmt = C_CessBaseAmt * C_CessRate

                Case "A_CVD@BCD"
                    'A_CVD_BaseAmt = C_Cess_BaseAmt + C_Cess_TaxAmt +BCD_BaseAmt A_CVD_TaxAmt = A_CVD_BaseAmt * A_CVD_Rate
                    If (Weight > 0) Then
                        BCDBaseAmt = Weight * 0.001 * Qty
                    Else
                        BCDBaseAmt = LineTotal
                    End If
                    'BCDTaxAmt = BCDBaseAmt * BCDRate

                    A_CVDBaseAmt = C_CessBaseAmt + C_CessTaxAmt + BCDBaseAmt
                    'A_CVDTaxAmt = A_CVDBaseAmt * A_CVDRate

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub AddItems(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Type", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocTyp", objform.Items.Item("70").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "ORDR", "U_VSPDCTYP", "lbl_DocTyp")
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocTyp").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("70").Left, _
                                          objform.Items.Item("70").Width, "Document Number", "70")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_DocNum", objform.Items.Item("lbl_DocTyp").Top + 15, objform.Items.Item("14").Left, _
                                            objform.Items.Item("14").Width, "ORDR", "U_VSPDCNO", "lbl_DocNum")
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objform.Items.Item("txt_DocNum").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("lbl_DocNum").Left, _
            '                            objform.Items.Item("lbl_DocNum").Width, "Fleet Status", "lbl_DocNum")
            'objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_FltSts", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("14").Left, _
            '                               objform.Items.Item("14").Width, "ORDR", "U_VSPFLSTS", "lbl_FltSts")


            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_OrTyp", objform.Items.Item("lbl_DocNum").Top + 15, objform.Items.Item("lbl_DocNum").Left, _
                                       objform.Items.Item("lbl_DocNum").Width, "Order Type", "lbl_DocNum")
            objMain.objUtilities.AddComboBox(objform.UniqueID, "txt_OrTyp", objform.Items.Item("txt_DocNum").Top + 15, objform.Items.Item("14").Left, _
                                           objform.Items.Item("14").Width, "ORDR", "U_VSPORTYP", "lbl_OrTyp")

            ''Added on 26-09-2018

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_PoNo", objform.Items.Item("86").Top + 15, objform.Items.Item("86").Left, _
                                       objform.Items.Item("86").Width, "PO Number", "86")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_PoNo", objform.Items.Item("86").Top + 15, objform.Items.Item("46").Left, _
                                            objform.Items.Item("46").Width, "ORDR", "U_VSPPONUM", "lbl_PoNo")

            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_PoEnt", objform.Items.Item("lbl_PoNo").Top + 15, objform.Items.Item("lbl_PoNo").Left, _
                                        objform.Items.Item("lbl_PoNo").Width, "PO Entry", "lbl_PoNo")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_PoEnt", objform.Items.Item("lbl_PoNo").Top + 15, objform.Items.Item("txt_PoNo").Left, _
                                            objform.Items.Item("txt_PoNo").Width, "ORDR", "U_VSPPOENT", "lbl_PoEnt")

            ''Added on 08-11-02018
            objMain.objUtilities.AddLabel(objform.UniqueID, "lbl_VLQTY", objform.Items.Item("lbl_PoEnt").Top + 15, objform.Items.Item("lbl_PoEnt").Left, _
                                       objform.Items.Item("lbl_PoEnt").Width, "Vendor Loaded Qty", "lbl_PoEnt")
            objMain.objUtilities.AddEditBox(objform.UniqueID, "txt_VLQTY", objform.Items.Item("lbl_VLQTY").Top, objform.Items.Item("txt_PoEnt").Left, _
                                            objform.Items.Item("txt_PoEnt").Width, "ORDR", "U_VSPVLQTY", "lbl_VLQTY")


            'objComboBox = objform.Items.Item("txt_FltSts").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Open to Fleet", "Open to Fleet")
            'objComboBox.ValidValues.Add("Linked to Fleet", "Linked to Fleet")

            'objComboBox1 = objform.Items.Item("txt_OrTyp").Specific
            'objComboBox1.ValidValues.Add("", "")
            'objComboBox1.ValidValues.Add("TR", "Transport")
            'objComboBox1.ValidValues.Add("MT", "Material")
            'objComboBox1.ValidValues.Add("Other", "Other")
            'If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            '    objComboBox1.Select("Other", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PostRevenueDetails(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)
            objform.Freeze(True)
            oDBs_Head = objform.DataSources.DBDataSources.Item("ORDR")

            Dim DocEntry As String = oDBs_Head.GetValue("DocEntry", 0)

            Dim GetCCDetails As String = ""

            If objMain.IsSAPHANA = True Then
                GetCCDetails = "Select T0.""DocDate"",T0.""DocEntry"",T0.""U_TCCode"" ,T0.""U_TCRef"" ,Sum(T1.""Quantity"" * T1.""U_VSPUNPRC"") as ""Total"",T0.""Comments"" , " & _
           "T1.""OcrCode2"" as ""Vehicle CC"" From ORDR T0 Inner Join RDR1 T1 On T0.""DocEntry"" = T1.""DocEntry"" Where T0.""DocEntry"" = '" & DocEntry & "' And " & _
           "((T1.""ItemCode"" In (Select ""U_VSPCITM"" From OADM)) Or (T1.""ItemCode"" In (Select ""U_VSPTITM""  From OADM))) Group By T0.""DocDate"",T0.""DocEntry"", " & _
           "T0.""U_TCCode"" ,T0.""U_TCRef"" ,T0.""Comments"",T1.""OcrCode2"""

            Else
                GetCCDetails = "Select T0.DocDate,T0.DocEntry,T0.U_TCCode ,T0.U_TCRef ,Sum(T1.Quantity * T1.U_VSPUNPRC) as Total,T0.Comments , " & _
           "T1.OcrCode2 as 'Vehicle CC' From ORDR T0 Inner Join RDR1 T1 On T0.DocEntry = T1.DocEntry Where T0.DocEntry = '" & DocEntry & "' And " & _
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

                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("DocEntry", oRsGetTripEntry.Fields.Item(0).Value.ToString)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C3")

                        If oRsGetLineType.Fields.Item(0).Value = "" Then
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPTYPE", "Sales")
                            objMain.oChildren.Item(CInt(oRsGetRevenueLine.Fields.Item(0).Value)).SetProperty("U_VSPDOCTY", "Sales Order")
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
                            objMain.oChild.SetProperty("U_VSPTYPE", "Sales")
                            objMain.oChild.SetProperty("U_VSPDOCTY", "Sales Order")
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
        objform = objMain.objApplication.Forms.GetForm("139", objMain.objApplication.Forms.ActiveForm.TypeCount)
        objMatrix = objform.Items.Item("38").Specific

        Select Case BusinessObjectInfo.EventType

            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                Dim status As String = objform.Items.Item("81").Specific.Value
                If status = "3" Then
                    Try
                        objform.Items.Item("txt_PoNo").Enabled = False
                        objform.Items.Item("txt_PoEnt").Enabled = False
                        objform.Items.Item("txt_VLQTY").Enabled = False
                    Catch ex As Exception

                    End Try
                    
                ElseIf status = "1" Then
                    Try
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objform.Items.Item("txt_PoNo").Enabled = True
                        objform.Items.Item("txt_PoEnt").Enabled = True
                        objform.Items.Item("txt_VLQTY").Enabled = True
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Try
                            objform.Items.Item("txt_PoEnt").Enabled = False
                        Catch ex As Exception

                        End Try


                        ' objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Catch ex As Exception

                    End Try

                End If

            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                Try
                    If BusinessObjectInfo.BeforeAction = False Then
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objform.Items.Item("txt_PoNo").Enabled = True
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        objform.Items.Item("txt_PoEnt").Enabled = False
                        objform.Items.Item("46").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                    End If

                Catch ex As Exception

                End Try


            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try

                    If BusinessObjectInfo.BeforeAction = True Then

                        Dim ordertype As String = CStr(objform.Items.Item("txt_OrTyp").Specific.Selected.Value).Trim
                        If ordertype = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Select order Type", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            BubbleEvent = False
                            Exit Try
                        End If

                        If CStr(objform.Items.Item("txt_OrTyp").Specific.Selected.Value).Trim = "TR" Then
                            ''Check Contract is there with this Customer or not

                            Dim checkContract As String = ""
                            Dim BpCode As String = objform.Items.Item("4").Specific.Value
                            Dim docdate As String = objform.Items.Item("10").Specific.Value
                            Dim BlanketQuantity As Double = 0.0
                            Dim SalesOrderQuantity As Double = 0.0
                            Dim quantity As Double = 0.0
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                quantity = quantity + objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                            Next

                            docdate = docdate.Insert(4, "/")
                            docdate = docdate.Insert(7, "/")
                            If objMain.IsSAPHANA = True Then
                                checkContract = "Select T0.""Number"",T0.""BpCode"",T0.""BpName"",T1.""PlanQty"",T0.""StartDate"",T0.""EndDate"" from OOAT T0  INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" where T0.""BpCode""='" & BpCode & "' And   '" & CDate(docdate).ToString("yyyyMMdd") & "' between    T0.""StartDate"" And T0.""EndDate"" "
                            Else
                                checkContract = checkContract = "Select T0.""Number"",T0.""BpCode"",T0.""BpName"",T1.""PlanQty"",T0.""StartDate"",T0.""EndDate"" from OOAT T0  INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" where T0.""BpCode""='" & BpCode & "' And   '" & CDate(docdate).ToString("yyyyMMdd") & "' between    T0.""StartDate"" And T0.""EndDate"" "
                            End If
                            Dim oRsGetcheckContract As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetcheckContract.DoQuery(checkContract)
                            If oRsGetcheckContract.RecordCount = 0 Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("There Is No Contract With This Customer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objform.Freeze(False)
                                Exit Try

                            End If

                            Dim startdate As String = oRsGetcheckContract.Fields.Item("StartDate").Value
                            Dim enddate As String = oRsGetcheckContract.Fields.Item("EndDate").Value
                            Dim checkSoQuantity As String = ""
                            If objMain.IsSAPHANA = True Then
                                checkSoQuantity = "SELECT Sum(T1.""Quantity"") as ""Qty"" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.""DocEntry"" = T1.""DocEntry""" & _
                                                   " WHERE T0.""CardCode"" ='" & BpCode & "' And  T0.""DocDate""  between  '" & CDate(startdate).ToString("yyyyMMdd") & "'" & _
                                                    " And  '" & CDate(enddate).ToString("yyyyMMdd") & "' And T0.""U_VSPORTYP""='MT'"
                            Else
                                checkSoQuantity = "SELECT Sum(T1.""Quantity"") as ""Qty"" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.""DocEntry"" = T1.""DocEntry""" & _
                                                    " WHERE T0.""CardCode"" ='" & BpCode & "' And  T0.""DocDate""  between  '" & CDate(startdate).ToString("yyyyMMdd") & "'" & _
                                                     " And  '" & CDate(enddate).ToString("yyyyMMdd") & "' And T0.""U_VSPORTYP""='MT'"
                            End If
                            Dim oRsGetcheckSoQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetcheckSoQuantity.DoQuery(checkSoQuantity)

                            BlanketQuantity = CDbl(oRsGetcheckContract.Fields.Item("PlanQty").Value)
                            SalesOrderQuantity = CDbl(oRsGetcheckSoQuantity.Fields.Item("Qty").Value)
                            SalesOrderQuantity = SalesOrderQuantity + quantity

                            If SalesOrderQuantity > BlanketQuantity Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Sales Order Quantity Can Not Be Greater  Than From Contract  Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objform.Freeze(False)
                                Exit Try
                            End If
                        ElseIf CStr(objform.Items.Item("txt_OrTyp").Specific.Selected.Value).Trim = "MT" Then
                            ''Po Number is Mandatory
                            Dim PoNumber As String = CStr(objform.Items.Item("txt_PoNo").Specific.Value).Trim
                            If PoNumber = "" Then
                                objMain.objApplication.StatusBar.SetText("Please Select Po Number", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objform.Items.Item("txt_PoEnt").Enabled = True
                                objform.Items.Item("txt_PoEnt").Specific.Value = ""
                                objform.Items.Item("txt_PoNo").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform.Items.Item("txt_PoEnt").Enabled = False
                                BubbleEvent = False

                                Exit Try
                            End If

                            ''Check Purchase Order No Is Exist In Sales Order

                            'Dim checkPOINSaleOrderExist As String = ""
                            'If objMain.IsSAPHANA = True Then
                            '    checkPOINSaleOrderExist = "Select ""U_VSPPONUM"",""U_VSPPOENT"" from ORDR where ""U_VSPPONUM""='" & PoNumber & "'  "
                            'Else
                            '    checkPOINSaleOrderExist = "Select ""U_VSPPONUM"",""U_VSPPOENT"" from ORDR where ""U_VSPPONUM""='" & PoNumber & "'  "
                            'End If
                            'Dim oRscheckPOINSaleOrderExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            'oRscheckPOINSaleOrderExist.DoQuery(checkPOINSaleOrderExist)

                            'If oRscheckPOINSaleOrderExist.RecordCount > 0 Then
                            '    objMain.objApplication.StatusBar.SetText("This PO Number Is Alreday Assign to Some Other Sales Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    BubbleEvent = False
                            '    Exit Try
                            'End If

                            ''End code Po 

                            ''Check Contract is there with this Customer or not

                            Dim checkContract As String = ""
                            Dim BpCode As String = objform.Items.Item("4").Specific.Value
                            Dim docdate As String = objform.Items.Item("10").Specific.Value
                            Dim BlanketQuantity As Double = 0.0
                            Dim SalesOrderQuantity As Double = 0.0
                            Dim quantity As Double = 0.0
                            For i As Integer = 1 To objMatrix.VisualRowCount
                                quantity = quantity + objMatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                            Next


                            docdate = docdate.Insert(4, "/")
                            docdate = docdate.Insert(7, "/")

                            If objMain.IsSAPHANA = True Then
                                checkContract = "Select T0.""Number"",T0.""BpCode"",T0.""BpName"",T1.""PlanQty"",T1.""ItemCode"",T0.""StartDate"",T0.""EndDate"" from OOAT T0  INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" where T0.""BpCode""='" & BpCode & "' And   '" & CDate(docdate).ToString("yyyyMMdd") & "' between    T0.""StartDate"" And T0.""EndDate"" "
                            Else
                                checkContract = "Select T0.""Number"",T0.""BpCode"",T0.""BpName"",T1.""PlanQty"",T1.""ItemCode"",T0.""StartDate"",T0.""EndDate"" from OOAT T0  INNER JOIN OAT1 T1 ON T0.""AbsID"" = T1.""AgrNo"" where T0.""BpCode""='" & BpCode & "' And   '" & CDate(docdate).ToString("yyyyMMdd") & "' between    T0.""StartDate"" And T0.""EndDate"" "
                            End If
                            Dim oRsGetcheckContract As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetcheckContract.DoQuery(checkContract)
                            

                            If oRsGetcheckContract.RecordCount = 0 Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("There Is No Contract With This Customer", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Try
                            End If

                            Dim startdate As String = oRsGetcheckContract.Fields.Item("StartDate").Value
                            Dim enddate As String = oRsGetcheckContract.Fields.Item("EndDate").Value
                            Dim checkSoQuantity As String = ""
                            If objMain.IsSAPHANA = True Then
                                checkSoQuantity = "SELECT Sum(T1.""Quantity"") as ""Qty"" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.""DocEntry"" = T1.""DocEntry""" & _
                                                   " WHERE T0.""CardCode"" ='" & BpCode & "' And  T0.""DocDate""  between  '" & CDate(startdate).ToString("yyyyMMdd") & "'" & _
                                                    " And  '" & CDate(enddate).ToString("yyyyMMdd") & "' And T0.""U_VSPORTYP""='MT'"
                            Else
                                checkSoQuantity = "SELECT Sum(T1.""Quantity"") as ""Qty"" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.""DocEntry"" = T1.""DocEntry""" & _
                                                    " WHERE T0.""CardCode"" ='" & BpCode & "' And  T0.""DocDate""  between  '" & CDate(startdate).ToString("yyyyMMdd") & "'" & _
                                                     " And  '" & CDate(enddate).ToString("yyyyMMdd") & "' And T0.""U_VSPORTYP""='MT'"
                            End If
                            Dim oRsGetcheckSoQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetcheckSoQuantity.DoQuery(checkSoQuantity)

                            BlanketQuantity = CDbl(oRsGetcheckContract.Fields.Item("PlanQty").Value)
                            SalesOrderQuantity = CDbl(oRsGetcheckSoQuantity.Fields.Item("Qty").Value)
                            SalesOrderQuantity = SalesOrderQuantity + quantity

                            If SalesOrderQuantity > BlanketQuantity Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Sales Order Quantity Can Not Be Greater Than From Contract Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                objform.Freeze(False)
                                Exit Try
                            End If

                        End If
                    End If

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                        objMain.objDocumentType.UpdateDocument("ORDR", "Sales Order")
                        Me.PostRevenueDetails(objform.UniqueID)

                    End If


                    'If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                    '    objMain.objDocumentType.UpdateDocument("ORDR", "Sales Order")
                    '    Me.PostRevenueDetails(objform.UniqueID)
                    'End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

    Sub CFLFilterForPO(ByVal FormUID As String, ByVal CFL_ID As String, ByVal CardCode As String)
        Try

            'Dim getvendoe As String = CardCode
            'If getvendoe = "" Then
            '    objMain.objApplication.StatusBar.SetText("Please Configure Fuel Vendor Code in Configuration Screen")

            '    Exit Try
            'End If

            objform = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "CardCode"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = CardCode
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CflAdding(ByVal FormUID As String)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        oCFLs = objForm.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList

        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFLCreationParams.MultiSelection = False
        oCFLCreationParams.ObjectType = "22"
        oCFLCreationParams.UniqueID = "CFL_PONO"
        Try
            oCFL = oCFLs.Add(oCFLCreationParams)
            oEditText = objform.Items.Item("txt_PoNo").Specific
            oEditText.ChooseFromListUID = "CFL_PONO"
            oEditText.ChooseFromListAlias = "DocNum"
        Catch ex As Exception
        End Try
    End Sub
    Sub CflAdding1(ByVal FormUID As String)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        oCFLs = objform.ChooseFromLists
        Dim oCFL As SAPbouiCOM.ChooseFromList

        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

        oCFLCreationParams.MultiSelection = False
        oCFLCreationParams.ObjectType = "22"
        oCFLCreationParams.UniqueID = "CFL_POENT"
        Try
            oCFL = oCFLs.Add(oCFLCreationParams)
            oEditText = objform.Items.Item("txt_PoEnt").Specific
            oEditText.ChooseFromListUID = "CFL_POENT"
            oEditText.ChooseFromListAlias = "DocEntry"
        Catch ex As Exception
        End Try
    End Sub

End Class
