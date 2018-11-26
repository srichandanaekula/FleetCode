Public Class clsTyreMapping

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oColumn As SAPbouiCOM.Column
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Tyre Mapping.xml", "VSP_FLT_TYRMPG_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TYRMPG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
            
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("200").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMatrix1 = objForm.Items.Item("7").Specific

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            objMatrix1.Columns.Item("V_8").ValidValues.Add("", "")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Attached", "Attached")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Removed", "Removed")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Repair", "Repair")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Stepney", "Stepney")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OTYRMPG"))

            objMatrix1 = objForm.Items.Item("7").Specific
            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()

            Me.SetNewLine(objForm.UniqueID)

            For i As Integer = 1 To objMatrix1.VisualRowCount
                objMatrix1.CommonSetting.SetRowEditable(i, True)
            Next

            objMatrix1.Columns.Item("V_-1").Editable = False
            objMatrix1.Columns.Item("V_1").Editable = False
            objMatrix1.Columns.Item("V_2").Editable = False
            objMatrix1.Columns.Item("V_3").Editable = False
            objMatrix1.Columns.Item("V_4").Editable = False
            objMatrix1.Columns.Item("V_5").Editable = False
            objMatrix1.Columns.Item("V_6").Editable = False
            objMatrix1.Columns.Item("V_10").Editable = False
            objMatrix1.Columns.Item("V_11").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")

            objMatrix1 = objForm.Items.Item("7").Specific

            objMatrix1.AddRow()
            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
            oDBs_Details1.SetValue("U_VSPTRNUM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPTRNM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPWLTYP", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPTRSIZ", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPPSTN", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPUOM1", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPITCD", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPGIIC", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPGDRPT", oDBs_Details1.Offset, "")
            oDBs_Details1.SetValue("U_VSPGDIS", oDBs_Details1.Offset, "")

            objMatrix1.SetLineData(objMatrix1.VisualRowCount)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                                           pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then
                            BubbleEvent = False
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
                    objMatrix1 = objForm.Items.Item("7").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "CFL_TYRNO" Then
                            Me.CFLFilter(objForm.UniqueID, oCFL.UniqueID)
                            Me.SetCellsEditable(objForm.UniqueID)
                        End If
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_VCHNO" Then
                            oDBs_Head.SetValue("U_VSPVCHN0", oDBs_Head.Offset, oDT.GetValue("U_VSPVNO", 0))
                            oDBs_Head.SetValue("U_VSPVCHNM", oDBs_Head.Offset, oDT.GetValue("U_VSPVNM", 0))
                            oDBs_Head.SetValue("U_VSPVMDL", oDBs_Head.Offset, oDT.GetValue("U_VSPMODEL", 0))
                        End If

                        If oCFL.UniqueID = "CFL_TYRNO" Then
                            For i As Integer = 1 To objMatrix1.VisualRowCount
                                If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value = oDT.GetValue("U_VSPTRNUM", 0) And _
                                objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value <> "Removed" Then
                                    If pVal.Row > 1 Then
                                        objMain.objApplication.StatusBar.SetText("Selected Tyre already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                End If
                            Next

                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_VSPTRNUM", oDBs_Details1.Offset, oDT.GetValue("U_VSPTRNUM", 0))
                            oDBs_Details1.SetValue("U_VSPTRNM", oDBs_Details1.Offset, oDT.GetValue("U_VSPTRNM", 0))
                            oDBs_Details1.SetValue("U_VSPWLTYP", oDBs_Details1.Offset, oDT.GetValue("U_VSPWHL", 0))
                            oDBs_Details1.SetValue("U_VSPTRSIZ", oDBs_Details1.Offset, oDT.GetValue("U_VSPTRSZE", 0))
                            oDBs_Details1.SetValue("U_VSPUOM1", oDBs_Details1.Offset, oDT.GetValue("U_VSPUOM2", 0))
                            oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, oDT.GetValue("U_VSPCPCTY", 0))
                            oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, oDT.GetValue("U_VSPUOM1", 0))
                            oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPPSTN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPITCD", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_VSPGIIC", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_VSPGDRPT", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_VSPGDIS", oDBs_Details1.Offset, "")
                            objMatrix1.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_PST" Then

                            For i As Integer = 1 To objMatrix1.VisualRowCount
                                If objMatrix1.Columns.Item("V_7").Cells.Item(i).Specific.Value = oDT.GetValue("U_VSPPSTN", 0) And objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value <> "Removed" Then
                                    If pVal.Row > 1 Then
                                        objMain.objApplication.StatusBar.SetText("Selected Position already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    End If
                                End If
                            Next

                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_VSPTRNUM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPTRNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPWLTYP", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPTRSIZ", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPUOM1", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details1.SetValue("U_VSPPSTN", oDBs_Details1.Offset, oDT.GetValue("U_VSPPSTN", 0))
                            oDBs_Details1.SetValue("U_VSPITCD", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_VSPGIIC", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_VSPGDRPT", oDBs_Details1.Offset, "")
                            oDBs_Details1.SetValue("U_VSPGDIS", oDBs_Details1.Offset, "")
                            objMatrix1.SetLineData(pVal.Row)

                            If pVal.Row = objMatrix1.VisualRowCount() Then
                                SetNewLine(objForm.UniqueID)
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_GIIC" And pVal.BeforeAction = False Then

                            If objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value = "" Then
                                If objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = "Removed" Then

                                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                                    oDBs_Details1.SetValue("U_VSPTRNUM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPTRNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPWLTYP", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPTRSIZ", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPUOM1", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPPSTN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPITCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPGIIC", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                                    oDBs_Details1.SetValue("U_VSPGDRPT", oDBs_Details1.Offset, "")
                                    oDBs_Details1.SetValue("U_VSPGDIS", oDBs_Details1.Offset, "")
                                    objMatrix1.SetLineData(pVal.Row)

                                    If objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value = "" Then
                                        If objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = "Removed" Then
                                            If pVal.Row = objMatrix1.VisualRowCount Then
                                                SetNewLine(objForm.UniqueID)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If oCFL.UniqueID = "CFL_ITCD" And pVal.BeforeAction = False Then

                            If objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value = "" Then
                                If objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = "Removed" Then

                                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                                    oDBs_Details1.SetValue("U_VSPTRNUM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPTRNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPWLTYP", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPTRSIZ", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPUOM1", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPCPCTY", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPUOM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPSTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPPSTN", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Details1.SetValue("U_VSPITCD", oDBs_Details1.Offset, oDT.GetValue("ItemCode", 0))
                                    oDBs_Details1.SetValue("U_VSPGIIC", oDBs_Details1.Offset, "")
                                    oDBs_Details1.SetValue("U_VSPGDRPT", oDBs_Details1.Offset, "")
                                    oDBs_Details1.SetValue("U_VSPGDIS", oDBs_Details1.Offset, "")
                                    objMatrix1.SetLineData(pVal.Row)

                                    If objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value = "" Then
                                        If objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = "Removed" Then
                                            If pVal.Row = objMatrix1.VisualRowCount Then
                                                SetNewLine(objForm.UniqueID)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    'Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
                    '    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
                    '    objMatrix1 = objForm.Items.Item("7").Specific

                    '    If pVal.ItemUID = "7" And pVal.ColUID = "V_10" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '        If objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value = "" Then

                    '        End If
                    '    End If

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
                    objMatrix1 = objForm.Items.Item("7").Specific

                    If pVal.ItemUID = "7" And pVal.ColUID = "V_10" And pVal.BeforeAction = False Then
                        If objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value <> "" Then

                            Dim GDcNum As String = ""

                            If objMain.IsSAPHANA = True Then
                                GDcNum = "Select ""DocNum"" From OIGN Where ""DocEntry""='" & objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value & "'"
                            Else
                                GDcNum = "Select DocNum From OIGN Where DocEntry='" & objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value & "'"

                            End If

                            Dim oRSGDcNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSGDcNum.DoQuery(GDcNum)

                            Dim GdRCPTForm As SAPbouiCOM.Form
                            objMain.objApplication.ActivateMenuItem("3078")
                            GdRCPTForm = objMain.objApplication.Forms.GetForm("721", objMain.objApplication.Forms.ActiveForm.TypeCount)

                            GdRCPTForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            GdRCPTForm.Freeze(True)
                            GdRCPTForm.Items.Item("7").Specific.Value = oRSGDcNum.Fields.Item(0).Value
                            GdRCPTForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            GdRCPTForm.Freeze(False)
                        End If
                    End If

                    If pVal.ItemUID = "7" And pVal.ColUID = "V_11" And pVal.BeforeAction = False Then
                        If objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value <> "" Then

                            Dim GDcNum As String = ""

                            If objMain.IsSAPHANA = True Then
                                GDcNum = "Select ""DocNum"" From OIGE Where ""DocEntry""='" & objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value & "'"
                            Else
                                GDcNum = "Select DocNum From OIGE Where DocEntry='" & objMatrix1.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value & "'"
                            End If

                            Dim oRSGDcNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRSGDcNum.DoQuery(GDcNum)

                            Dim GdIssForm As SAPbouiCOM.Form
                            objMain.objApplication.ActivateMenuItem("3079")
                            GdIssForm = objMain.objApplication.Forms.GetForm("720", objMain.objApplication.Forms.ActiveForm.TypeCount)

                            GdIssForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            GdIssForm.Freeze(True)
                            GdIssForm.Items.Item("7").Specific.Value = oRSGDcNum.Fields.Item(0).Value
                            GdIssForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            GdIssForm.Freeze(False)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "200" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim ChkItemExist As String = ""
                        If objMain.IsSAPHANA = True Then
                            ChkItemExist = "Select ""DocNum"" From ""@VSP_FLT_TYRMPG"" Where ""U_VSPVCHN0"" ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"

                        Else
                            ChkItemExist = "Select DocNum From [@VSP_FLT_TYRMPG] Where [U_VSPVCHN0] ='" & objForm.Items.Item("200").Specific.Value.Trim & "'"

                        End If
                        Dim oRsChkItemExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkItemExist.DoQuery(ChkItemExist)
                        If oRsChkItemExist.RecordCount > 0 Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("12").Specific.value = oRsChkItemExist.Fields.Item(0).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub SetCellsEditable(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objMatrix1 = objForm.Items.Item("7").Specific

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value = "" Then
                    objMatrix1.CommonSetting.SetRowEditable(i, True)
                Else
                    objMatrix1.CommonSetting.SetRowEditable(i, False)
                End If
            Next

            objMatrix1.Columns.Item("V_-1").Editable = False
            'objMatrix1.Columns.Item("V_0").Editable = False
            objMatrix1.Columns.Item("V_1").Editable = False
            objMatrix1.Columns.Item("V_2").Editable = False
            objMatrix1.Columns.Item("V_3").Editable = False
            objMatrix1.Columns.Item("V_4").Editable = False
            objMatrix1.Columns.Item("V_5").Editable = False
            objMatrix1.Columns.Item("V_6").Editable = False
            objMatrix1.Columns.Item("V_7").Editable = True
            objMatrix1.Columns.Item("V_8").Editable = True
            objMatrix1.Columns.Item("V_9").Editable = True
            objMatrix1.Columns.Item("V_14").Editable = True
            objMatrix1.Columns.Item("V_10").Editable = False
            objMatrix1.Columns.Item("V_11").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TYRMPG" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "CreateGoodsReceipt" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TYRMPG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
                objMatrix1 = objForm.Items.Item("7").Specific

                Dim i As Integer
                For i = 1 To objMatrix1.VisualRowCount
                    If objMatrix1.IsRowSelected(i) = True Then Exit For
                Next

                If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" And objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value <> "" Then
                    If objMatrix1.Columns.Item("V_10").Cells.Item(i).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(i).Specific.Value = "" Then

                        Me.CreateGdRcptAndGdIssue(objForm.UniqueID, objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value.Trim, _
                                                  objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value.Trim, oDBs_Head.GetValue("DocEntry", 0), i, _
                                                  objMatrix1.Columns.Item("V_7").Cells.Item(i).Specific.Value.Trim, objMatrix1.Columns.Item("V_14").Cells.Item(i).Specific.Value.Trim)

                    End If
                End If

                End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TYRMPG_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objMatrix1 = objForm.Items.Item("7").Specific

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Me.UpdateVehicleNo(objForm.UniqueID)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            Me.UpdateVehicleNo(objForm.UniqueID)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        For i As Integer = 1 To objMatrix1.VisualRowCount
                            If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" Then
                                objMatrix1.CommonSetting.SetRowEditable(i, False)
                            ElseIf (objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Or _
                              objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = " Stepney" Or _
                               objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Repair" Or _
                               objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "") Then
                                objMatrix1.CommonSetting.SetRowEditable(i, True)
                            End If
                        Next

                        objMatrix1.Columns.Item("V_-1").Editable = False
                        objMatrix1.Columns.Item("V_1").Editable = False
                        objMatrix1.Columns.Item("V_2").Editable = False
                        objMatrix1.Columns.Item("V_3").Editable = False
                        objMatrix1.Columns.Item("V_4").Editable = False
                        objMatrix1.Columns.Item("V_5").Editable = False
                        objMatrix1.Columns.Item("V_6").Editable = False
                        objMatrix1.Columns.Item("V_9").Editable = True
                        objMatrix1.Columns.Item("V_14").Editable = True

                        'For i As Integer = 1 To objMatrix1.VisualRowCount
                        '    If objMatrix1.Columns.Item("V_9").Cells.Item(i).Specific.Value = "" Then
                        '        objMatrix1.CommonSetting.SetRowEditable(i, True)
                        '    Else
                        '        objMatrix1.CommonSetting.SetRowEditable(i, False)
                        '    End If
                        'Next

                        'objMatrix1.Columns.Item("V_-1").Editable = False
                        'objMatrix1.Columns.Item("V_1").Editable = False
                        'objMatrix1.Columns.Item("V_2").Editable = False
                        'objMatrix1.Columns.Item("V_3").Editable = False
                        'objMatrix1.Columns.Item("V_4").Editable = False
                        'objMatrix1.Columns.Item("V_5").Editable = False
                        'objMatrix1.Columns.Item("V_6").Editable = False

                        If objMatrix1.Columns.Item("V_9").Cells.Item(objMatrix1.VisualRowCount).Specific.Value = "" Then
                            objMatrix1.Columns.Item("V_9").Editable = True
                        Else
                            objMatrix1.Columns.Item("V_9").Editable = False
                        End If
                        If objMatrix1.Columns.Item("V_14").Cells.Item(objMatrix1.VisualRowCount).Specific.Value = "" Then
                            objMatrix1.Columns.Item("V_14").Editable = True
                        Else
                            objMatrix1.Columns.Item("V_14").Editable = False
                        End If
                        Dim EV_2 As Integer
                        For i As Integer = 0 To objMatrix1.Columns.Count - 1
                            If objMatrix1.Columns.Item(i).UniqueID = "V_0" Then
                                EV_2 = i
                                Exit For
                            End If
                        Next
                        For i As Integer = 1 To objMatrix1.VisualRowCount
                            If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                                objMatrix1.CommonSetting.SetCellEditable(i, EV_2, False)
                            Else
                                objMatrix1.CommonSetting.SetCellEditable(i, EV_2, True)
                            End If
                        Next
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Sub UpdateVehicleNo(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("7").Specific

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then

                    Dim GetCode As String = ""

                    If objMain.IsSAPHANA = True Then
                        GetCode = "Select ""Code"" From ""@VSP_FLT_TYRMSTR"" Where ""U_VSPTRNUM"" = '" & objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value & "'"
                    Else
                        GetCode = "Select Code From [@VSP_FLT_TYRMSTR] Where U_VSPTRNUM = '" & objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value & "'"

                    End If

                    Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetCode.DoQuery(GetCode)

                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMSTR")
                    objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    objMain.oGeneralParams.SetProperty("Code", oRsGetCode.Fields.Item(0).Value)
                    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Or _
                        objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Stepney" Then
                        objMain.oGeneralData.SetProperty("U_VSPVNO", objForm.Items.Item("200").Specific.Value.ToString)
                        objMain.oGeneralData.SetProperty("U_VSPPSTN", objMatrix1.Columns.Item("V_7").Cells.Item(i).Specific.Value.ToString)
                    ElseIf (objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" Or _
                            objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Repair") Then
                        objMain.oGeneralData.SetProperty("U_VSPVNO", "")
                        objMain.oGeneralData.SetProperty("U_VSPPSTN", "")
                    End If
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If
            Next

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Removed" Then
                    objMatrix1.CommonSetting.SetRowEditable(i, False)
                ElseIf (objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Or _
                              objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Stepney" Or _
                              objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "") Then
                    objMatrix1.CommonSetting.SetRowEditable(i, True)
                End If
            Next
            objMatrix1.Columns.Item("V_1").Editable = False
            objMatrix1.Columns.Item("V_2").Editable = False
            objMatrix1.Columns.Item("V_3").Editable = False
            objMatrix1.Columns.Item("V_4").Editable = False
            objMatrix1.Columns.Item("V_5").Editable = False
            objMatrix1.Columns.Item("V_6").Editable = False
            objMatrix1.Columns.Item("V_9").Editable = True
            objMatrix1.Columns.Item("V_14").Editable = True

            Dim EV_2 As Integer
            For i As Integer = 0 To objMatrix1.Columns.Count - 1
                If objMatrix1.Columns.Item(i).UniqueID = "V_0" Then
                    EV_2 = i
                    Exit For
                End If
            Next
            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                    objMatrix1.CommonSetting.SetCellEditable(i, EV_2, False)
                Else
                    objMatrix1.CommonSetting.SetCellEditable(i, EV_2, True)
                End If
            Next


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPVNO"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
            oChooseFromList.SetConditions(oConditions)

            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR

            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPVNO"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = ""
            oChooseFromList.SetConditions(oConditions)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("7").Specific

            Dim GetNoofTyres As String = ""
            If objMain.IsSAPHANA = True Then
                GetNoofTyres = "Select Cast(""U_VSPTYRES"" As Integer)  As ""U_VSPTYRES""  From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO""='" & objForm.Items.Item("200").Specific.Value & "' "
            Else
                GetNoofTyres = "Select Cast(U_VSPTYRES As Integer)  As U_VSPTYRES  From [@VSP_FLT_VMSTR] Where U_VSPVNO='" & objForm.Items.Item("200").Specific.Value & "' "

            End If

            Dim oRsGetNoofTyres As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetNoofTyres.DoQuery(GetNoofTyres)

            Dim AttachedCount As Integer = 0

            For i As Integer = 1 To objMatrix1.VisualRowCount
                If objMatrix1.Columns.Item("V_8").Cells.Item(i).Specific.Value = "Attached" Then
                    AttachedCount = AttachedCount + 1
                End If
            Next

            If AttachedCount > oRsGetNoofTyres.Fields.Item("U_VSPTYRES").Value Then
                objMain.objApplication.StatusBar.SetText("No. of Tyres must not be greater than Tyres in Vehicle Master")
                Return False
            End If
            If objForm.Items.Item("200").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Function

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING

        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        objMatrix1 = objForm.Items.Item("7").Specific

        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If eventInfo.ItemUID = "7" And eventInfo.ColUID = "V_-1" And objMatrix1.RowCount > 0 Then

                            If objMatrix1.Columns.Item("V_8").Cells.Item(eventInfo.Row).Specific.Value = "Removed" And objMatrix1.Columns.Item("V_9").Cells.Item(eventInfo.Row).Specific.Value <> "" And _
                            objMatrix1.Columns.Item("V_14").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If objMatrix1.Columns.Item("V_10").Cells.Item(eventInfo.Row).Specific.Value = "" And objMatrix1.Columns.Item("V_11").Cells.Item(eventInfo.Row).Specific.Value = "" Then
                                    Try
                                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                        oMenus = oMenuItem.SubMenus
                                        If oMenus.Exists("CreateGoodsReceipt") = False Then
                                            oCreationPackage.UniqueID = "CreateGoodsReceipt"
                                            oCreationPackage.String = "CreateGoodsReceipt"
                                            oCreationPackage.Enabled = True
                                            oMenus.AddEx(oCreationPackage)
                                        End If
                                    Catch ex As Exception
                                        objMain.objApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                Else
                                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                    oMenus = oMenuItem.SubMenus
                                    Try
                                        If oMenus.Exists("CreateGoodsReceipt") = True Then
                                            objMain.objApplication.Menus.RemoveEx("CreateGoodsReceipt")
                                        End If
                                    Catch ex As Exception
                                        objMain.objApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CreateGdRcptAndGdIssue(ByVal FormUID As String, ByVal GRItemCode As String, ByVal TyreNo As String, ByVal DocEntry As Integer, ByVal Row As Integer, _
                               ByVal TyrePos As String, ByVal GIItemCode As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
            objMatrix1 = objForm.Items.Item("7").Specific

            Dim oGoodsReceipt As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

            Dim DocDate As String = DateTime.Now.ToString("yyyyMMdd")
            DocDate = DocDate.Insert(4, "-")
            DocDate = DocDate.Insert(7, "-")
            oGoodsReceipt.DocDate = DocDate

            Dim TaxDate As String = DateTime.Now.ToString("yyyyMMdd")
            TaxDate = TaxDate.Insert(4, "-")
            TaxDate = TaxDate.Insert(7, "-")
            oGoodsReceipt.TaxDate = TaxDate

            oGoodsReceipt.Lines.ItemCode = GRItemCode

            Dim GetCost As String = ""

            If objMain.IsSAPHANA = True Then
                GetCost = "SELECT T0.""CostTotal"" FROM OSRN T0 WHERE T0.""DistNumber"" = '" & TyreNo & "'"
            Else
                GetCost = "SELECT T0.[CostTotal] FROM OSRN T0 WHERE T0.[DistNumber] = '" & TyreNo & "'"


            End If
            Dim oRsGetCost As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCost.DoQuery(GetCost)

            If oRsGetCost.RecordCount > 0 Then
                oGoodsReceipt.Lines.UnitPrice = oRsGetCost.Fields.Item(0).Value
            Else
                oGoodsReceipt.Lines.UnitPrice = 0
            End If

            oGoodsReceipt.Lines.SerialNumbers.SetCurrentLine(0)
            'oGoodsReceipt.Lines.SerialNum = TyreNo
            oGoodsReceipt.Lines.SerialNumbers.InternalSerialNumber = TyreNo
            oGoodsReceipt.Lines.SerialNumbers.Quantity = 1
            oGoodsReceipt.Lines.SerialNumbers.Add()

            Dim GetCrCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetCrCode = "Select ""OcrCode"" From OOCR where ""OcrName""='" & objForm.Items.Item("200").Specific.Value & "'"

            Else
                GetCrCode = "Select ""OcrCode"" From OOCR where ""OcrName""='" & objForm.Items.Item("200").Specific.Value & "'"

            End If
            Dim oRsGetCrCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCrCode.DoQuery(GetCrCode)
            If oRsGetCrCode.RecordCount > 0 Then
                oGoodsReceipt.Lines.CostingCode2 = oRsGetCrCode.Fields.Item(0).Value
            End If

            If oGoodsReceipt.Add = 0 Then

                objMain.objApplication.StatusBar.SetText("Goods Receipt Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                objMatrix1.CommonSetting.SetRowEditable(Row, False)

                'Updating Goods Receipt no to Tyre mapping
                Dim GetGdRcptDEntry As String = ""

                If objMain.IsSAPHANA = True Then
                    GetGdRcptDEntry = "Select Max(""DocEntry"") From OIGN"
                Else
                    GetGdRcptDEntry = "Select Max(DocEntry) From OIGN"
                End If
                Dim oRsGetGdRcptDEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetGdRcptDEntry.DoQuery(GetGdRcptDEntry)
                Dim GDrcpVal As String = oRsGetGdRcptDEntry.Fields.Item(0).Value
                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMPG")
                objMain.oGeneralData = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TYRMPG_C0")
                objMain.oChildren.Item(Row - 1).SetProperty("U_VSPGDRPT", GDrcpVal)
                objMain.oGeneralService.Update(objMain.oGeneralData)

                'To Refresh Form
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
                oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")
                objMatrix1 = objForm.Items.Item("7").Specific
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                Me.SetNewLine(objForm.UniqueID)
                objMatrix1.LoadFromDataSource()
                objMatrix1.AutoResizeColumns()
                objForm.Refresh()
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                End If

                'Updating Tyre Master 
                Dim GetTrMstrDEntry As String = ""
                If objMain.IsSAPHANA = True Then
                    GetTrMstrDEntry = "Select ""Code"" From ""@VSP_FLT_TYRMSTR"" Where ""U_VSPTRNUM"" = '" & TyreNo & "'"

                Else
                    GetTrMstrDEntry = "Select Code From [@VSP_FLT_TYRMSTR] Where U_VSPTRNUM = '" & TyreNo & "'"

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
                    objMain.oGeneralData.SetProperty("U_VSPITCD", GRItemCode)
                    'objMain.oGeneralData.SetProperty("U_VSPGIIC", GIItemCode)
                    objMain.oGeneralService.Update(objMain.oGeneralData)

                End If

                'Opening Goods Issue
                Dim GdIssueForm As SAPbouiCOM.Form
                Dim objMat As SAPbouiCOM.Matrix
                objMain.objApplication.ActivateMenuItem("3079")
                GdIssueForm = objMain.objApplication.Forms.GetForm("720", objMain.objApplication.Forms.ActiveForm.TypeCount)
                GdIssueForm.Freeze(True)
                objMat = GdIssueForm.Items.Item("13").Specific

                objMain.objUtilities.AddLabel(GdIssueForm.UniqueID, "lbl_DcEtr", GdIssueForm.Items.Item("20").Top + 20, GdIssueForm.Items.Item("20").Left, GdIssueForm.Items.Item("20").Width, _
                                           "Ref DocEntry", "20")
                objMain.objUtilities.AddEditBox(GdIssueForm.UniqueID, "txt_DcEtr", GdIssueForm.Items.Item("20").Top + 20, GdIssueForm.Items.Item("3").Left, GdIssueForm.Items.Item("20").Width - 10, _
                                                "OIGE", "U_VSPDCETY", "lbl_DcEtr")

                objMain.objUtilities.AddLabel(GdIssueForm.UniqueID, "lbl_LnNo", GdIssueForm.Items.Item("20").Top + 20, GdIssueForm.Items.Item("txt_DcEtr").Left + GdIssueForm.Items.Item("txt_DcEtr").Width + 25, GdIssueForm.Items.Item("20").Width, _
                                           "Ref LineNo", "3")
                objMain.objUtilities.AddEditBox(GdIssueForm.UniqueID, "txt_LnNo", GdIssueForm.Items.Item("20").Top + 20, GdIssueForm.Items.Item("lbl_LnNo").Left + GdIssueForm.Items.Item("lbl_LnNo").Width + 4, GdIssueForm.Items.Item("20").Width - 10, _
                                                "OIGE", "U_VSPLNUM", "lbl_LnNo")

                objMain.objUtilities.AddLabel(GdIssueForm.UniqueID, "lbl_DcTyp", GdIssueForm.Items.Item("20").Top + 35, GdIssueForm.Items.Item("20").Left, GdIssueForm.Items.Item("20").Width, _
                                           "Ref DocType", "20")
                objMain.objUtilities.AddEditBox(GdIssueForm.UniqueID, "txt_DcTyp", GdIssueForm.Items.Item("20").Top + 35, GdIssueForm.Items.Item("3").Left, GdIssueForm.Items.Item("20").Width - 10, _
                                                "OIGE", "U_VSPDCTYP", "lbl_DcTyp")

                objMain.objUtilities.AddLabel(GdIssueForm.UniqueID, "lbl_TyPS", GdIssueForm.Items.Item("20").Top + 35, GdIssueForm.Items.Item("txt_DcTyp").Left + GdIssueForm.Items.Item("txt_DcTyp").Width + 25, GdIssueForm.Items.Item("20").Width, _
                                          "TyrePosition", "lbl_LnNo")
                objMain.objUtilities.AddEditBox(GdIssueForm.UniqueID, "txt_TyPS", GdIssueForm.Items.Item("20").Top + 35, GdIssueForm.Items.Item("lbl_TyPS").Left + GdIssueForm.Items.Item("lbl_TyPS").Width + 4, GdIssueForm.Items.Item("20").Width + 30, _
                                                "OIGE", "U_VSPTYPS", "lbl_TyPS")

                GdIssueForm.Items.Item("txt_DcEtr").Specific.Value = DocEntry
                GdIssueForm.Items.Item("txt_LnNo").Specific.Value = Row
                GdIssueForm.Items.Item("txt_DcTyp").Specific.Value = "TyreMapping"
                GdIssueForm.Items.Item("txt_TyPS").Specific.Value = TyrePos

                GdIssueForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                GdIssueForm.Items.Item("txt_DcEtr").Enabled = False
                GdIssueForm.Items.Item("txt_LnNo").Enabled = False
                GdIssueForm.Items.Item("txt_DcTyp").Enabled = False
                GdIssueForm.Items.Item("txt_TyPS").Enabled = False

                objMat.Columns.Item("1").Cells.Item(1).Specific.Value = GIItemCode
                If oRsGetCrCode.RecordCount > 0 Then
                    Dim Ccode As String = ";" & oRsGetCrCode.Fields.Item(0).Value
                    objMat.Columns.Item("10001018").Cells.Item(1).Specific.Value = Ccode
                End If
                GdIssueForm.Freeze(False)

            Else
                objMain.objApplication.StatusBar.SetText("GoodsReceipt Error : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RefreshData(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0")

            objMatrix1 = objForm.Items.Item("7").Specific

            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TYRMPG_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
            objMatrix1.LoadFromDataSource()
            objMatrix1.AutoResizeColumns()
            objForm.Refresh()
            objForm.Freeze(False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
