Public Class clsGoodIssue

#Region "        Declaration        "
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Public objform As SAPbouiCOM.Form

#End Region

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        objform = objMain.objApplication.Forms.GetForm("720", objMain.objApplication.Forms.ActiveForm.TypeCount)
        Select Case BusinessObjectInfo.EventType
            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                Try
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Dim GetGIDtls As String = "Select ""DocEntry"" , ""U_VSPDCTYP"" , ""U_VSPLNUM"" , ""U_VSPDCNO"" From ""OIGE""  Where ""DocNum"" = (Select MAX(""DocNum"") From ""OIGE"" )"
                        Dim oRsGetGIDtls As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetGIDtls.DoQuery(GetGIDtls)

                        Dim getDocentry As String = ""
                        If objMain.IsSAPHANA = True Then
                            getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_CALBRTN"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                        Else
                            getDocentry = "Select ""DocEntry"" from ""@VSP_FLT_CALBRTN"" where ""DocNum""='" & oRsGetGIDtls.Fields.Item("U_VSPDCNO").Value.ToString & "' "
                        End If

                        Dim oRsgetDocentry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsgetDocentry.DoQuery(getDocentry)



                        If oRsGetGIDtls.Fields.Item("U_VSPDCTYP").Value = "Calibration" Then
                            Dim LN As Integer = oRsGetGIDtls.Fields.Item("U_VSPLNUM").Value
                            Dim DocNum As Integer = oRsGetGIDtls.Fields.Item("DocEntry").Value

                            objMain.sCmp = objMain.objCompany.GetCompanyService
                            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OCALBRTN")
                            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                            objMain.oGeneralParams.SetProperty("DocEntry", oRsgetDocentry.Fields.Item("DocEntry").Value)
                            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                            objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_CALBRTN_C0")

                            'objMain.oChildren.Item(LN - 1).SetProperty("U_VSPGI", DocNum.ToString(""))
                            objMain.oChildren.Item(LN - 1).SetProperty("U_VSPGI", DocNum.ToString())
                            objMain.oGeneralService.Update(objMain.oGeneralData)

                        End If

                        If objform.Items.Item("txt_DcTyp").Specific.Value = "TyreMapping" And _
                      objform.Items.Item("txt_DcEtr").Specific.Value <> "" _
                      And objform.Items.Item("txt_LnNo").Specific.Value <> "" _
                       And objform.Items.Item("txt_TyPS").Specific.Value <> "" Then

                            Me.UpdateTyreMapping(objform.UniqueID)

                        End If

                    End If
                Catch ex As Exception
                    objMain.objApplication.StatusBar.SetText(ex.Message)
                End Try
        End Select
    End Sub
#End Region

    Sub UpdateTyreMapping(ByVal FormUID As String)
        Try
            objform = objMain.objApplication.Forms.Item(FormUID)

            If objform.Items.Item("txt_DcTyp").Specific.Value = "TyreMapping" And _
                       objform.Items.Item("txt_DcEtr").Specific.Value <> "" _
                       And objform.Items.Item("txt_LnNo").Specific.Value <> "" _
                        And objform.Items.Item("txt_TyPS").Specific.Value <> "" Then

                Dim DocEntry As Integer = objform.Items.Item("txt_DcEtr").Specific.Value
                Dim LineNum As Integer = objform.Items.Item("txt_LnNo").Specific.Value
                Dim TyrPstn As String = objform.Items.Item("txt_TyPS").Specific.Value
                Dim objMat As SAPbouiCOM.Matrix
                objMat = objform.Items.Item("13").Specific

                Dim GetDocEntry As String = ""
                If objMain.IsSAPHANA = True Then
                    GetDocEntry = "Select ""DocEntry"" From OIGE Where ""DocNum"" = '" & objform.Items.Item("7").Specific.Value & "'"
                Else
                    GetDocEntry = "Select DocEntry From OIGE Where DocNum = '" & objform.Items.Item("7").Specific.Value & "'"
                End If
                Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRs.DoQuery(GetDocEntry)
                If oRs.RecordCount > 0 Then
                    Dim DEtry As String = oRs.Fields.Item(0).Value
                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMPG")
                    objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                    objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    objMain.oGeneralParams.SetProperty("DocEntry", DocEntry)
                    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    'objMain.oGeneralData.SetProperty("U_VSPITCD", "")
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TYRMPG_C0")
                    objMain.oChildren.Item(LineNum - 1).SetProperty("U_VSPGDIS", DEtry)

                    Dim GetTyreNo As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetTyreNo = "Select T3.""DistNumber"", T0.""ItemCode"" From IGE1  T0 Inner Join OITL T1 " & _
                    "On T0.""DocEntry""  = T1.""ApplyEntry"" And T1.""ApplyType"" = '60' Inner join ITL1 T2 on T1.""LogEntry"" = T2.""LogEntry"" " & _
                    "Inner join OSRN T3 on T2.""ItemCode"" = '" & objMat.Columns.Item("1").Cells.Item(1).Specific.Value & "' And T2.""MdAbsEntry"" = T3.""AbsEntry"" And T1.""ApplyEntry"" = '" & oRs.Fields.Item(0).Value & "'"

                    Else
                        GetTyreNo = "Select T3.DistNumber, T0.ItemCode From [dbo].IGE1  T0 Inner Join [dbo].[OITL] T1 " & _
                                            "On T0.DocEntry  = T1.ApplyEntry And T1.ApplyType = '60' Inner join [dbo].[ITL1] T2 on T1.LogEntry = T2.LogEntry " & _
                                            "Inner join [dbo].[OSRN] T3 on T2.Itemcode = '" & objMat.Columns.Item("1").Cells.Item(1).Specific.Value & "' And T2.MdAbsEntry = T3.AbsEntry And T1.ApplyEntry = '" & oRs.Fields.Item(0).Value & "'"

                    End If
                    Dim oRSGetTyreNo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRSGetTyreNo.DoQuery(GetTyreNo)

                    Dim TyrNo As String = oRSGetTyreNo.Fields.Item(0).Value
                    Dim GIItem As String = oRSGetTyreNo.Fields.Item(1).Value

                    Dim GetTyreDet As String = ""

                    If objMain.IsSAPHANA = True Then
                        GetTyreDet = "Select * From ""@VSP_FLT_TYRMSTR"" Where ""U_VSPTRNUM""='" & TyrNo & "'"
                    Else
                        GetTyreDet = "Select * From [@VSP_FLT_TYRMSTR] Where U_VSPTRNUM='" & TyrNo & "'"
                    End If
                    Dim oRsGetTyreDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetTyreDet.DoQuery(GetTyreDet)

                    Dim Gettotrows As String = ""
                    If objMain.IsSAPHANA = True Then
                        Gettotrows = "Select COUNT(""U_VSPTRNUM"") From ""@VSP_FLT_TYRMPG_C0"" Where ""DocEntry""='" & DocEntry & "' and ""U_VSPTRNUM"" <>''"
                    Else
                        Gettotrows = "Select COUNT(U_VSPTRNUM) From [@VSP_FLT_TYRMPG_C0] Where DocEntry='" & DocEntry & "' and U_VSPTRNUM <>''"
                    End If
                    Dim oRsGettotrows As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGettotrows.DoQuery(Gettotrows)

                    Dim LNo As Integer = oRsGettotrows.Fields.Item(0).Value

                    'objMain.oChildren.Item(LNo).SetProperty("LineId", LNo + 1)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPTRNUM", TyrNo)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPTRNM", oRsGetTyreDet.Fields.Item("U_VSPTRNM").Value)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPWLTYP", oRsGetTyreDet.Fields.Item("U_VSPWHL").Value)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPTRSIZ", oRsGetTyreDet.Fields.Item("U_VSPTRSZE").Value)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPUOM1", oRsGetTyreDet.Fields.Item("U_VSPUOM2").Value)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPCPCTY", oRsGetTyreDet.Fields.Item("U_VSPCPCTY").Value)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPUOM", oRsGetTyreDet.Fields.Item("U_VSPUOM1").Value)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPPSTN", TyrPstn)
                    objMain.oChildren.Item(LNo).SetProperty("U_VSPSTS", "Attached")

                    objMain.oGeneralService.Update(objMain.oGeneralData)

                    Dim GetTrMstrDEntry As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetTrMstrDEntry = "Select ""Code"" From ""@VSP_FLT_TYRMSTR"" Where ""U_VSPTRNUM"" = '" & TyrNo & "'"
                    Else
                        GetTrMstrDEntry = "Select Code From [@VSP_FLT_TYRMSTR] Where U_VSPTRNUM = '" & TyrNo & "'"
                    End If
                    Dim oRsGetTrMstrDEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetTrMstrDEntry.DoQuery(GetTrMstrDEntry)

                    If oRsGetTrMstrDEntry.RecordCount > 0 Then

                        Dim TYMstrDCNO As String = oRsGetTrMstrDEntry.Fields.Item(0).Value
                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMSTR")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("Code", TYMstrDCNO)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        'objMain.oGeneralData.SetProperty("U_VSPITCD", GIItem)
                        objMain.oGeneralData.SetProperty("U_VSPGIIC", GIItem)
                        objMain.oGeneralService.Update(objMain.oGeneralData)

                    End If

                    Dim objTympForm As SAPbouiCOM.Form
                    objTympForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TYRMPG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    objMain.objTyreMapping.RefreshData(objTympForm.UniqueID)
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
