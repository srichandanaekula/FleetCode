Public Class clsDocumentType

#Region " Declaration         "
    Dim objForm As SAPbouiCOM.Form
    Dim odt As SAPbouiCOM.DataTable
    Dim oOptionBtn As SAPbouiCOM.OptionBtn
    Dim oDBs_Head As SAPbouiCOM.DBDataSource
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim OtherForm As SAPbouiCOM.Form
    Dim BaseDocNum, BaseDocType
#End Region

    Sub CreateForm(ByVal BaseForm As String, ByVal BaseNum As String, ByVal BaseType As String)
        Try
            objMain.objUtilities.LoadForm("Document Type.xml", "VSP_FLT_DOCTYPE_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_DOCTYPE_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            If Fleet.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = False Then
                Fleet.MainCls.ohtLookUpForm.Add(objForm.UniqueID, BaseForm)
            End If

            Me.BaseDocNum = BaseNum
            Me.BaseDocType = BaseType

            objForm.DataSources.UserDataSources.Add("U_PURORD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_GRPO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_APINV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_OUTPAY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_SALORD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_DELVRY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_ARINV", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_INPAY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            objForm.DataSources.UserDataSources.Add("U_INVTRF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

            oOptionBtn = objForm.Items.Item("9").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_PURORD")
            objForm.Items.Item("9").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("10").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_GRPO")
            objForm.Items.Item("10").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("3").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_APINV")
            objForm.Items.Item("3").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("4").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_OUTPAY")
            objForm.Items.Item("4").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("5").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_SALORD")
            objForm.Items.Item("5").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("6").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_DELVRY")
            objForm.Items.Item("6").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("7").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_ARINV")
            objForm.Items.Item("7").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("8").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_INPAY")
            objForm.Items.Item("8").AffectsFormMode = False

            oOptionBtn = objForm.Items.Item("11").Specific
            oOptionBtn.DataBind.SetBound(True, "", "U_INVTRF")
            objForm.Items.Item("11").AffectsFormMode = False

            objForm.Items.Item("9").Specific.GroupWith("10")
            objForm.Items.Item("9").Specific.GroupWith("3")
            objForm.Items.Item("9").Specific.GroupWith("4")
            objForm.Items.Item("9").Specific.GroupWith("5")
            objForm.Items.Item("9").Specific.GroupWith("6")
            objForm.Items.Item("9").Specific.GroupWith("7")
            objForm.Items.Item("9").Specific.GroupWith("8")
            objForm.Items.Item("9").Specific.GroupWith("11")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If Fleet.MainCls.ohtLookUpForm.ContainsKey(objForm.UniqueID) = True And pVal.BeforeAction = False Then
                        Fleet.MainCls.ohtLookUpForm.Remove(objForm.UniqueID)
                    End If
                    '-----------------------------------------------------

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "OK" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("9").Specific.Selected = "True" Then 'Purchase Order
                            objMain.objApplication.ActivateMenuItem("2305")
                            OtherForm = objMain.objApplication.Forms.GetForm("142", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("10").Specific.Selected = "True" Then 'GRPO
                            objMain.objApplication.ActivateMenuItem("2306")
                            OtherForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("3").Specific.Selected = "True" Then 'A/P Invoice
                            objMain.objApplication.ActivateMenuItem("2308")
                            OtherForm = objMain.objApplication.Forms.GetForm("141", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("4").Specific.Selected = "True" Then 'Outgoing Payments
                            objMain.objApplication.ActivateMenuItem("2818")
                            objMain.objOutgoingPayment.OpenedFromDocumentType = True
                            objMain.objOutgoingPayment.DocumentType = BaseDocType
                            objMain.objOutgoingPayment.DocNum = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("5").Specific.Selected = "True" Then 'Sales Order
                            objMain.objApplication.ActivateMenuItem("2050")
                            OtherForm = objMain.objApplication.Forms.GetForm("139", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("6").Specific.Selected = "True" Then 'Delivery
                            objMain.objApplication.ActivateMenuItem("2051")
                            OtherForm = objMain.objApplication.Forms.GetForm("140", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("7").Specific.Selected = "True" Then 'A/R Invoice
                            objMain.objApplication.ActivateMenuItem("2053")
                            OtherForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("8").Specific.Selected = "True" Then 'Incoming Payments
                            objMain.objApplication.ActivateMenuItem("2817")
                            objMain.objIncomingPayment.OpenedFromDocumentType = True
                            objMain.objIncomingPayment.DocumentType = BaseDocType
                            objMain.objIncomingPayment.DocNum = BaseDocNum
                            objForm.Close()
                        ElseIf objForm.Items.Item("11").Specific.Selected = "True" Then 'Inventory Transfer
                            objMain.objApplication.ActivateMenuItem("3080")
                            OtherForm = objMain.objApplication.Forms.GetForm("940", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            OtherForm.Items.Item("txt_DocTyp").Specific.Value = BaseDocType
                            OtherForm.Items.Item("txt_DocNum").Specific.Value = BaseDocNum
                            objForm.Close()
                        Else
                            objMain.objApplication.StatusBar.SetText("Atlease One Document Type Has to Be Selected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub UpdateDocument(ByVal TableName As String, ByVal DocType As String)
        Try
            'If objMain.UpdateBaseDocument = False Then Exit Try

            'Dim GetGIDtls As String = "Select DocEntry , DocDate , U_VSPDCTYP , U_VSPDCNO , DocTotal , CardCode , NumAtCard , Comments From " & TableName & " Where DocEntry = (Select MAX(DocEntry) From " & TableName & ")"
            Dim GetGIDtls As String = ""
            If objMain.IsSAPHANA = True Then
                GetGIDtls = "Select ""DocEntry"" , TO_NVARCHAR(""DocDate"",'yyyy/MM/dd') as ""DCDate"" ,""DocNum"", ""U_VSPDCTYP"" , ""U_VSPDCNO"" , ""DocTotal"" , ""CardCode"" ,""NumAtCard"" , ""Comments"" From """ & TableName & """ Where ""DocEntry"" = (Select MAX(""DocEntry"") From """ & TableName & """)"
            Else
                GetGIDtls = "Select DocEntry ,""DocNum"", DocDate , U_VSPDCTYP , U_VSPDCNO , DocTotal , CardCode , NumAtCard , Comments From " & TableName & " Where DocEntry = (Select MAX(DocEntry) From " & TableName & ")"
            End If

            Dim oRsGetGIDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetGIDtls.DoQuery(GetGIDtls)

            If oRsGetGIDtls.Fields.Item("U_VSPDCTYP").Value = "Accident History" Then

                Dim getDocentry As String = ""
                If objMain.IsSAPHANA = True Then
                    getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_ACCHIST"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                Else
                    getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_ACCHIST"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                End If

                Dim oRsgetDocentry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsgetDocentry.DoQuery(getDocentry)



                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OACCHIST")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oRsgetDocentry.Fields.Item("DocEntry").Value.ToString)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_ACCHIST_C0")

                'Dim CheckFirstLine As String = "Select U_VSPDOCTY From [@VSP_FLT_ACCHIST_C0] Where DocEntry = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And LineId = '1'"
                Dim CheckFirstLine As String = ""
                If objMain.IsSAPHANA = True Then
                    CheckFirstLine = "Select ""U_VSPDOCTY"" From ""@VSP_FLT_ACCHIST_C0"" Where ""DocEntry"" = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And ""LineId"" = '1'"
                Else
                    CheckFirstLine = "Select U_VSPDOCTY From [@VSP_FLT_ACCHIST_C0] Where DocEntry = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And LineId = '1'"
                End If

                Dim oRsCheckFirstLine As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsCheckFirstLine.DoQuery(CheckFirstLine)

                If oRsCheckFirstLine.Fields.Item(0).Value = "" Then
                    objMain.oChildren.Item(0).SetProperty("U_VSPDOCTY", DocType)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDOCNM", oRsGetGIDtls.Fields.Item(0).Value.ToString)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                Else
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChild.SetProperty("U_VSPDOCTY", DocType)
                    objMain.oChild.SetProperty("U_VSPDOCNM", oRsGetGIDtls.Fields.Item(0).Value.ToString)
                    objMain.oChild.SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                    objMain.oChild.SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If
                objMain.FormCloseBoolean = True
            ElseIf oRsGetGIDtls.Fields.Item("U_VSPDCTYP").Value = "Tyre Maintenance" Then
                Dim getDocentry As String = ""
                If objMain.IsSAPHANA = True Then
                    getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_TYRMTNC"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                Else
                    getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_TYRMTNC"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                End If

                Dim oRsgetDocentry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsgetDocentry.DoQuery(getDocentry)
                Dim doc As String = oRsgetDocentry.Fields.Item("DocEntry").Value.ToString

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMTNC")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oRsgetDocentry.Fields.Item("DocEntry").Value.ToString)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TYRMTNC_C1")



                'Dim CheckFirstLine As String = "Select U_VSPDOCTY From [@VSP_FLT_TYRMTNC_C1] Where DocEntry = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And LineId = '1'"

                Dim CheckFirstLine As String = ""
                If objMain.IsSAPHANA = True Then
                    CheckFirstLine = "Select ""U_VSPDOCTY"" From ""@VSP_FLT_TYRMTNC_C1"" Where ""DocEntry"" = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And ""LineId"" = '1'"
                Else
                    CheckFirstLine = "Select U_VSPDOCTY From [@VSP_FLT_TYRMTNC_C1] Where DocEntry = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And LineId = '1'"
                End If

                Dim oRsCheckFirstLine As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsCheckFirstLine.DoQuery(CheckFirstLine)

                If oRsCheckFirstLine.Fields.Item(0).Value = "" Then
                    objMain.oChildren.Item(0).SetProperty("U_VSPDOCTY", DocType)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDOCNM", oRsGetGIDtls.Fields.Item(0).Value.ToString)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                Else
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChild.SetProperty("U_VSPDOCTY", DocType)
                    objMain.oChild.SetProperty("U_VSPDOCNM", oRsGetGIDtls.Fields.Item(0).Value.ToString)
                    objMain.oChild.SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                    objMain.oChild.SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If
                objMain.FormCloseBoolean = True
            ElseIf oRsGetGIDtls.Fields.Item("U_VSPDCTYP").Value = "Tank Maintenance" Then

                Dim getDocentry As String = ""
                If objMain.IsSAPHANA = True Then
                    getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_TNKMTNC"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                Else
                    getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_TNKMTNC"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                End If

                Dim oRsgetDocentry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsgetDocentry.DoQuery(getDocentry)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTNKMTNC")

                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oRsgetDocentry.Fields.Item("DocEntry").Value.ToString)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TNKMTNC_C1")

                Dim CheckFirstLine As String = ""
                If objMain.IsSAPHANA = True Then
                    CheckFirstLine = "Select ""U_VSPDOCTY"" From ""@VSP_FLT_TNKMTNC_C1"" Where ""DocEntry"" = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And ""LineId"" = '1'"
                Else
                    CheckFirstLine = "Select U_VSPDOCTY From [@VSP_FLT_TNKMTNC_C1] Where DocEntry = '" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value & "' And LineId = '1'"
                End If
                Dim oRsCheckFirstLine As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsCheckFirstLine.DoQuery(CheckFirstLine)


                If oRsCheckFirstLine.Fields.Item(0).Value = "" Then
                    objMain.oChildren.Item(0).SetProperty("U_VSPDOCTY", DocType)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDOCNM", oRsGetGIDtls.Fields.Item(0).Value.ToString)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                    objMain.oChildren.Item(0).SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                Else
                    objMain.oChild = objMain.oChildren.Add
                    objMain.oChild.SetProperty("U_VSPDOCTY", DocType)
                    objMain.oChild.SetProperty("U_VSPDOCNM", oRsGetGIDtls.Fields.Item(0).Value.ToString)
                    objMain.oChild.SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                    objMain.oChild.SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                    Dim doc1 As String = oRsGetGIDtls.Fields.Item(0).Value.ToString
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If
                objMain.FormCloseBoolean = True

            ElseIf oRsGetGIDtls.Fields.Item("U_VSPDCTYP").Value = "Trip Sheet" Then
               
                Dim otherForm As SAPbouiCOM.Form = objMain.objApplication.Forms.GetForm("VSP_FLT_TRSHT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                Dim objMatrix As SAPbouiCOM.Matrix = otherForm.Items.Item("1000023").Specific
                Dim line As Integer = objMatrix.VisualRowCount

                Dim GetRef As String = oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value
                Dim DocEntry As String = GetRef.Substring(0, GetRef.LastIndexOf("-"))
                Dim i As Integer = GetRef.Remove(0, GetRef.LastIndexOf("-") + 1)


                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", DocEntry)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C3")

                Select Case TableName

                    Case "ORDR"
                        
                        Dim getCommission As String = ""
                        If objMain.IsSAPHANA = True Then
                            getCommission = "Select ""U_VSPCOMSN"" from OCRD where ""CardCode""='" & oRsGetGIDtls.Fields.Item("CardCode").Value.ToString & "'"
                        Else
                            getCommission = "Select ""U_VSPCOMSN"" from OCRD where ""CardCode""='" & oRsGetGIDtls.Fields.Item("CardCode").Value.ToString & "'"
                        End If
                        Dim oRsgetCommission As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsgetCommission.DoQuery(getCommission)

                        If oRsgetCommission.RecordCount > 0 Then
                            Dim commission As Double = 0
                            commission = CDbl(oRsgetCommission.Fields.Item(0).Value)
                            objMain.oGeneralData.SetProperty("U_VSPCOMSN", commission)
                        End If


                        Dim DocDate As String = oRsGetGIDtls.Fields.Item("DCDate").Value

                        'Dim DocDate As Date = oRsGetGIDtls.Fields.Item("DocDate").Value

                        'If Not DocDate.ToString("yyyyMMdd") = "18991230" Then
                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPDATE", DocDate.ToString("yyyyMMdd"))

                        'End If
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPDATE", DocDate.ToString)

                        ' objMain.oChildren.Item(i - 1).SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPDCNUM", oRsGetGIDtls.Fields.Item("DocEntry").Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPSONUM", oRsGetGIDtls.Fields.Item("DocNum").Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPBPCOD", oRsGetGIDtls.Fields.Item("CardCode").Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPREF", oRsGetGIDtls.Fields.Item("NumAtCard").Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                        objMain.oChildren.Item(i - 1).SetProperty("U_VSPREM", oRsGetGIDtls.Fields.Item("Comments").Value.ToString)
                        objMain.oGeneralService.Update(objMain.oGeneralData)

                        objMain.oChild = objMain.oChildren.Add
                        objMain.oChild.SetProperty("U_VSPDOCTY", "")
                        objMain.oChild.SetProperty("U_VSPDATE", "")
                        objMain.oChild.SetProperty("U_VSPDCNUM", "")
                        objMain.oChild.SetProperty("U_VSPSONUM", "")
                        objMain.oChild.SetProperty("U_VSPBPCOD", "")
                        ' objMain.oChild.SetProperty("U_VSPQUANT", "")
                        objMain.oChild.SetProperty("U_VSPREF", "")
                        objMain.oChild.SetProperty("U_VSPDCTOT", "")
                        objMain.oChild.SetProperty("U_VSPREM", "")
                        objMain.oGeneralService.Update(objMain.oGeneralData)

                    Case "ODLN"

                        Dim getDeliveryQuantity As String = ""
                        If objMain.IsSAPHANA = True Then
                            getDeliveryQuantity = "Select T1.""Quantity"",T1.""DocEntry"" from ODLN T0 Inner Join DLN1 T1 on T0.""DocEntry""=T1.""DocEntry""  where T1.""DocEntry""=(Select max(""DocEntry"") from ""ODLN"") "
                        Else
                            getDeliveryQuantity = "Select T1.""Quantity"",T1.""DocEntry"" from ODLN T0 Inner Join DLN1 T1 on T0.""DocEntry""=T1.""DocEntry""  where T1.""DocEntry""=(Select max(""DocEntry"") from ""ODLN"") "
                        End If
                        Dim oRsgetDeliveryQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsgetDeliveryQuantity.DoQuery(getDeliveryQuantity)

                        Dim getLineId As String = ""

                        If objMain.IsSAPHANA = True Then
                            getLineId = "Select max(""LineId"") as ""LineNum"" From ""@VSP_FLT_TRSHT_C3"" Where ""DocEntry"" = '" & DocEntry & "' "
                        Else
                            getLineId = "Select max(""LineId"") as ""LineNum"" From ""@VSP_FLT_TRSHT_C3"" Where ""DocEntry"" = '" & DocEntry & "' "
                        End If

                        Dim oRsgetLineId As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsgetLineId.DoQuery(getLineId)

                        Dim line1 As String = oRsgetLineId.Fields.Item("LineNum").Value

                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPDOCTY", "Delivery")

                        'Dim DocDate As Date = oRsGetGIDtls.Fields.Item("DocDate").Value

                        'If Not DocDate.ToString("yyyyMMdd") = "18991230" Then
                        '    objMain.oChildren.Item(line - 1).SetProperty("U_VSPDATE", DocDate.ToString("yyyyMMdd"))

                        'End If


                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DCDate").Value.ToString)
                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPDCNUM", oRsGetGIDtls.Fields.Item("DocEntry").Value.ToString)
                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPSONUM", oRsGetGIDtls.Fields.Item("DocNum").Value.ToString)
                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPBPCOD", oRsGetGIDtls.Fields.Item("CardCode").Value.ToString)
                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPREF", oRsGetGIDtls.Fields.Item("NumAtCard").Value.ToString)
                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                        objMain.oChildren.Item(line - 1).SetProperty("U_VSPREM", oRsGetGIDtls.Fields.Item("Comments").Value.ToString)
                        objMain.oGeneralService.Update(objMain.oGeneralData)


                        objMain.oChild = objMain.oChildren.Add
                        objMain.oChild.SetProperty("U_VSPDOCTY", "")
                        objMain.oChild.SetProperty("U_VSPDATE", "")
                        objMain.oChild.SetProperty("U_VSPDCNUM", "")
                        objMain.oChild.SetProperty("U_VSPSONUM", "")
                        objMain.oChild.SetProperty("U_VSPBPCOD", "")
                        objMain.oChild.SetProperty("U_VSPREF", "")
                        objMain.oChild.SetProperty("U_VSPDCTOT", "")
                        objMain.oChild.SetProperty("U_VSPREM", "")
                        objMain.oGeneralService.Update(objMain.oGeneralData)

                        objMain.objUtilities.RefreshDatasourceFromDB(otherForm.UniqueID, otherForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT"), "DocEntry", DocEntry.Trim)
                        objMain.objUtilities.RefreshDatasourceFromDB(otherForm.UniqueID, otherForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3"), "DocEntry", DocEntry.Trim)
                        otherForm.Refresh()
                        Try
                            objMain.objApplication.ActivateMenuItem("1304") ''For refresh DB
                        Catch ex As Exception

                        End Try


                        'Case "OINV"

                        '    Dim getINVQuantity As String = ""
                        '    If objMain.IsSAPHANA = True Then
                        '        getINVQuantity = "Select T1.""Quantity"",T1.""DocEntry"" from OINV T0 Inner Join INV1 T1 on T0.""DocEntry""=T1.""DocEntry""  where T1.""DocEntry""=(Select max(""DocEntry"") from ""OINV"") "
                        '    Else
                        '        getINVQuantity = "Select T1.""Quantity"",T1.""DocEntry"" from OINV T0 Inner Join INV1 T1 on T0.""DocEntry""=T1.""DocEntry""  where T1.""DocEntry""=(Select max(""DocEntry"") from ""OINV"") "
                        '    End If
                        '    Dim oRsgetINVQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRsgetINVQuantity.DoQuery(getINVQuantity)

                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPDATE", oRsGetGIDtls.Fields.Item("DocDate").Value.ToString)
                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPDCNUM", oRsGetGIDtls.Fields.Item("DocEntry").Value.ToString)
                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPBPCOD", oRsGetGIDtls.Fields.Item("CardCode").Value.ToString)
                        '    'objMain.oChildren.Item(i - 1).SetProperty("U_VSPQUANT", oRsgetINVQuantity.Fields.Item("Quantity").Value.ToString)
                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPREF", oRsGetGIDtls.Fields.Item("NumAtCard").Value.ToString)
                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPDCTOT", oRsGetGIDtls.Fields.Item("DocTotal").Value.ToString)
                        '    objMain.oChildren.Item(i - 1).SetProperty("U_VSPREM", oRsGetGIDtls.Fields.Item("Comments").Value.ToString)
                        '    objMain.oGeneralService.Update(objMain.oGeneralData)

                        '    objMain.objUtilities.RefreshDatasourceFromDB(otherForm.UniqueID, otherForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT"), "DocEntry", DocEntry.Trim)
                        '    objMain.objUtilities.RefreshDatasourceFromDB(otherForm.UniqueID, otherForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3"), "DocEntry", DocEntry.Trim)
                        '    otherForm.Refresh()

                        '    Try
                        '        objMain.objApplication.ActivateMenuItem("1304")
                        '    Catch ex As Exception

                        '    End Try
                End Select

                objMain.FormCloseBoolean = True
            End If

            objMain.UpdateBaseDocument = False
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
