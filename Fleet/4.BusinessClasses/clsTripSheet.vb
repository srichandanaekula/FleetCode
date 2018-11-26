Imports System.Threading
Imports System.IO
Public Class clsTripSheet

#Region "Declaration"
    Dim objForm, objVehicleMasterForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Detail1, oDBs_Detail2, oDBs_Detail3, oDBs_Detail4, oDBs_Detail5, oDBs_Detail6, oDBs_Detail7, oDBs_Detail8, oDBs_Detail9 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3, objMatrix4, objMatrix5, objMatrix6, objMatrix7, objMatrix8, objMatrix9 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oLink As SAPbouiCOM.LinkedButton
    Dim Path As String
    Public Row As Integer
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("TripSheet.xml", "VSP_FLT_TRSHT_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TRSHT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            objForm.Freeze(True)

            objMain.objApplication.StatusBar.SetText("Please Wait While Screen Loads.........", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix3 = objForm.Items.Item("1000023").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix5 = objForm.Items.Item("67").Specific
            objMatrix6 = objForm.Items.Item("1000028").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific
            objMatrix9 = objForm.Items.Item("1000030").Specific

            objMatrix1.Columns.Item("V_9").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix1.Columns.Item("V_10").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix2.Columns.Item("V_2").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix3.Columns.Item("V_8").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix4.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix5.Columns.Item("V_5").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
            objMatrix5.Columns.Item("V_3").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto

            Me.CellsMasking(objForm.UniqueID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
            oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
            oDBs_Detail3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3")
            oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")
            oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")
            oDBs_Detail6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C6")
            oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")
            oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")
            oDBs_Detail9 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9")

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            objComboBox = objForm.Items.Item("1000002").Specific
            objComboBox.ValidValues.Add("", "")
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")

            objMatrix1.Columns.Item("V_8").ValidValues.Add("", "")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("Yes", "Yes")
            objMatrix1.Columns.Item("V_8").ValidValues.Add("No", "No")

            'objMatrix3.Columns.Item("V_0").ValidValues.Add("", "")
            'objMatrix3.Columns.Item("V_0").ValidValues.Add("Sales", "Sales")
            'objMatrix3.Columns.Item("V_0").ValidValues.Add("Purchase", "Purchase")
            'objMatrix3.Columns.Item("V_0").ValidValues.Add("Delivery", "Delivery")
            'objMatrix3.Columns.Item("V_0").ValidValues.Add("A/R Invoice", "A/R Invoice")

            objMatrix3.Columns.Item("V_1").ValidValues.Add("", "")
            objMatrix3.Columns.Item("V_1").ValidValues.Add("Sales Order", "Sales Order")
            objMatrix3.Columns.Item("V_1").ValidValues.Add("Delivery", "Delivery")
            objMatrix3.Columns.Item("V_1").ValidValues.Add("A/R Invoice", "A/R Invoice")
            'objMatrix3.Columns.Item("V_1").ValidValues.Add("G.Receipt PO", "G.Receipt PO")
            'objMatrix3.Columns.Item("V_1").ValidValues.Add("A/P Invoice", "A/P Invoice")

            'objMatrix3.Columns.Item("V_2").ValidValues.Add("", "")
            'objMatrix3.Columns.Item("V_2").ValidValues.Add("New", "New")
            'objMatrix3.Columns.Item("V_2").ValidValues.Add("Existing", "Existing")

            
            Me.CFLAccounts(objForm.UniqueID, "CFL_FRACT", "N", "", "N")
            Me.CFLAccounts(objForm.UniqueID, "CFL_TOACT", "N", "", "N")
            '  Me.CFLExpensesAccounts(objForm.UniqueID, "CFL_EXACT", "Y", "5", "1", "N")
            Me.CFLVendorFilter(objForm.UniqueID, "CFL_VEN")
            ' Me.CFLFilter(objForm.UniqueID, "CFL_SO")
            'Me.CFLFilter(objForm.UniqueID, "CFL_INV")
            'Me.CFLFilter(objForm.UniqueID, "CFL_DLV")
            'Me.CFLFilter(objForm.UniqueID, "CFL_PO")
            'Me.CFLFilter(objForm.UniqueID, "CFL_APINV")
            ' Me.CFLFilter(objForm.UniqueID, "CFL_PO1")

            Me.CFLFilterForVehicles(objForm.UniqueID, "CFL_VHCL")

            Me.CFLFilterCostCentres(objForm.UniqueID, "CFL_DRC", "3")
            Me.CFLFilterCostCentres1(objForm.UniqueID, "CFL_DRC1", "3")



            If objMain.IsSAPHANA = True Then
                objMain.objUtilities.MatrixComboBoxValues(objMatrix1.Columns.Item("V_3"), "Select ""U_VSPFPLC"" , ""U_VSPFPLC"" From ""@VSP_FLT_RTMSTR_C1"" Group By ""U_VSPFPLC""")
            Else
                objMain.objUtilities.MatrixComboBoxValues(objMatrix1.Columns.Item("V_3"), "Select U_VSPFPLC , U_VSPFPLC From [@VSP_FLT_RTMSTR_C1] Group By U_VSPFPLC")
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
            oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
            oDBs_Detail3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3")
            oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")
            oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")
            oDBs_Detail6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C6")
            oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")
            oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")
            oDBs_Detail9 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OTRSHT"))
            oDBs_Head.SetValue("U_VSPDOCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_VSPSRTDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Head.SetValue("U_VSPSTS", oDBs_Head.Offset, "Open")

            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix3 = objForm.Items.Item("1000023").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix5 = objForm.Items.Item("67").Specific
            objMatrix6 = objForm.Items.Item("1000028").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific
            objMatrix9 = objForm.Items.Item("1000030").Specific

            objMatrix1.Clear()
            oDBs_Detail1.Clear()
            objMatrix1.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "61")

            objMatrix3.Clear()
            oDBs_Detail3.Clear()
            objMatrix3.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "1000023")

            objMatrix2.Clear()
            oDBs_Detail2.Clear()
            objMatrix2.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "1000022")

            objMatrix5.Clear()
            oDBs_Detail5.Clear()
            objMatrix5.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "67")

            objMatrix6.Clear()
            oDBs_Detail6.Clear()
            objMatrix6.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "1000028")

            objMatrix9.Clear()
            oDBs_Detail9.Clear()
            objMatrix9.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "1000030")

            objForm.Items.Item("1000013").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Me.SetDefaultCellsEditable(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub PressedOnLinked(ByVal TarGetItm As String)
        Try
            If TarGetItm <> "" Then
                objMain.objApplication.ActivateMenuItem("VSP_FLT_VMSTR")
                objVehicleMasterForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VMSTR_Form", _
                                                             objMain.objApplication.Forms.ActiveForm.TypeCount)
                objVehicleMasterForm.Freeze(True)
                objVehicleMasterForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                objVehicleMasterForm.Items.Item("5").Specific.Value = TarGetItm
                objVehicleMasterForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                objVehicleMasterForm.Freeze(False)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)

        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    objMatrix1 = objForm.Items.Item("61").Specific
                    objMatrix2 = objForm.Items.Item("1000022").Specific
                    objMatrix3 = objForm.Items.Item("1000023").Specific
                    objMatrix4 = objForm.Items.Item("1000024").Specific
                    objMatrix5 = objForm.Items.Item("67").Specific
                    objMatrix6 = objForm.Items.Item("1000028").Specific
                    objMatrix7 = objForm.Items.Item("69").Specific
                    objMatrix8 = objForm.Items.Item("1000029").Specific
                    objMatrix9 = objForm.Items.Item("1000030").Specific

                    If (pVal.ItemUID = "1000013" Or pVal.ItemUID = "1000014" Or pVal.ItemUID = "1000015" Or pVal.ItemUID = "1000016" Or pVal.ItemUID = "57" Or _
                    pVal.ItemUID = "1000019" Or pVal.ItemUID = "1000021" Or pVal.ItemUID = "60") And pVal.BeforeAction = False Then
                        objForm.Freeze(True)
                        Select Case pVal.ItemUID
                            Case "1000013"
                                objForm.PaneLevel = 1
                                objMatrix1.AutoResizeColumns()
                            Case "1000014"
                                objForm.PaneLevel = 2
                                objMatrix2.AutoResizeColumns()
                            Case "1000015"
                                objForm.PaneLevel = 3
                                objMatrix3.AutoResizeColumns()
                                objForm.Settings.MatrixUID = "1000023"
                            Case "1000016"
                                objForm.PaneLevel = 4
                                objMatrix4.AutoResizeColumns()
                                objMatrix5.AutoResizeColumns()
                            Case "57"
                                objForm.PaneLevel = 5
                                objMatrix6.AutoResizeColumns()
                            Case "1000019"
                                objForm.PaneLevel = 6
                                objMatrix7.AutoResizeColumns()
                            Case "1000021"
                                objForm.PaneLevel = 7
                                objMatrix8.AutoResizeColumns()
                            Case "60"
                                objForm.PaneLevel = 8
                                objMatrix9.AutoResizeColumns()
                        End Select
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                                           pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                    'If pVal.ItemUID = "1000020" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    '    Me.UpadateOdmtrRdngsFromVM(objForm.UniqueID)
                    'End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1000031" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If objMatrix2.VisualRowCount > 0 Then
                            If objMatrix2.Columns.Item("V_7").Cells.Item(objMatrix2.VisualRowCount).Specific.Value = "" Then
                                objMain.objApplication.StatusBar.SetText("Previous Advance Should Be Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                Me.SetNewLine(objForm.UniqueID, "1000022")
                                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "85" And pVal.BeforeAction = False Then
                        Me.PressedOnLinked(objForm.Items.Item("4").Specific.Value)
                    End If

                    If pVal.ItemUID = "1000032" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.PostJE(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1000033" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.PostExpense(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1000034" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.PostDiesel(objForm.UniqueID)

                    End If

                    If pVal.ItemUID = "1000035" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.CloseTripSheet(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "74" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.UpdateTyreDetails(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "86" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim oChk As SAPbouiCOM.CheckBox = objForm.Items.Item("79").Specific
                        If oChk.Checked = False Then
                            Me.LoadTyreDetails(objForm.UniqueID)
                        End If
                    End If

                    ''Code for button  Add Row(ID=80).
                    If pVal.ItemUID = "80" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If objMatrix3.Columns.Item("V_5").Cells.Item(objMatrix3.VisualRowCount).Specific.Value <> "" Then
                            Me.SetNewLine(objForm.UniqueID, "1000023")
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            ' Me.SetCellsEditable(objForm.UniqueID)
                        Else
                            objMain.objApplication.StatusBar.SetText("Previous Document Has to Be Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
                    oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
                    oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
                    oDBs_Detail3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3")
                    oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")
                    oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")
                    oDBs_Detail6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C6")
                    oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")
                    oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")
                    oDBs_Detail9 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9")

                    objMatrix1 = objForm.Items.Item("61").Specific
                    objMatrix2 = objForm.Items.Item("1000022").Specific
                    objMatrix3 = objForm.Items.Item("1000023").Specific
                    objMatrix4 = objForm.Items.Item("1000024").Specific
                    objMatrix5 = objForm.Items.Item("67").Specific
                    objMatrix6 = objForm.Items.Item("1000028").Specific
                    objMatrix7 = objForm.Items.Item("69").Specific
                    objMatrix8 = objForm.Items.Item("1000029").Specific
                    objMatrix9 = objForm.Items.Item("1000030").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "CFL_RTCD" Then
                            If objForm.Items.Item("4").Specific.Value = "" Then
                                BubbleEvent = False
                                objMain.objApplication.StatusBar.SetText("Please Select Vehicle No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_CHITEM" And pVal.BeforeAction = True Then
                            Me.CFLFilterForChemicalItem(objForm.UniqueID, "CFL_CHITEM")
                        End If

                        If oCFL.UniqueID = "CFL_VHCL" Then
                            Me.CFLFilterForVehicles(objForm.UniqueID, oCFL.UniqueID)
                        End If


                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_VHCL" Then
                            Dim DIESE As String = oDT.GetValue("U_VSPDIITE", 0)
                            Dim odom As Double = oDT.GetValue("U_VSPODRDG", 0)
                            Me.LoadVehicleDetails(objForm.UniqueID, oDT.GetValue("U_VSPVNO", 0), oDT.GetValue("U_VSPDIITE", 0), oDT.GetValue("U_VSPCNTR", 0), _
                                                  oDT.GetValue("U_VSPCNTNM", 0), oDT.GetValue("U_VSPODRDG", 0))
                        End If

                        If oCFL.UniqueID = "CFL_RTCD" Then
                            Me.LoadRouteDetails(objForm.UniqueID, oDT.GetValue("U_VSPRCD", 0), oDT.GetValue("U_VSPSRCE", 0), oDT.GetValue("U_VSPDEST", 0), _
                                                oDT.GetValue("U_VSPADAMT", 0))
                        End If

                        If oCFL.UniqueID = "CFL_FRACT" Or oCFL.UniqueID = "CFL_TOACT" Then
                            oDBs_Detail2.SetValue("LineId", oDBs_Detail2.Offset, pVal.Row)
                            oDBs_Detail2.SetValue("U_VSPJOUDT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPAMGOO", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPADAMT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            If oCFL.UniqueID = "CFL_FRACT" Then
                                oDBs_Detail2.SetValue("U_VSPFRACT", oDBs_Detail2.Offset, oDT.GetValue("AcctCode", 0))
                                oDBs_Detail2.SetValue("U_VAPTOACT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            Else
                                oDBs_Detail2.SetValue("U_VSPFRACT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail2.SetValue("U_VAPTOACT", oDBs_Detail2.Offset, oDT.GetValue("AcctCode", 0))
                            End If
                            oDBs_Detail2.SetValue("U_VAPCAS", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPDRVCC", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPCOM", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPJENO", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPOPNO", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix2.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_EXACT" Then
                            oDBs_Detail4.SetValue("LineId", oDBs_Detail4.Offset, pVal.Row)
                            oDBs_Detail4.SetValue("U_VSPTYPE", oDBs_Detail4.Offset, objMatrix4.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail4.SetValue("U_VSPFRRMA", oDBs_Detail4.Offset, "N")
                            oDBs_Detail4.SetValue("U_VSPEXACC", oDBs_Detail4.Offset, oDT.GetValue("AcctCode", 0))
                            oDBs_Detail4.SetValue("U_VSPEXACN", oDBs_Detail4.Offset, oDT.GetValue("AcctName", 0))
                            If pVal.Row > 1 Then
                                oDBs_Detail4.SetValue("U_VSPADACC", oDBs_Detail4.Offset, objMatrix4.Columns.Item("V_4").Cells.Item(pVal.Row - 1).Specific.Value)
                            Else
                                If objMatrix2.VisualRowCount > 0 Then
                                    If objMatrix2.Columns.Item("V_4").Cells.Item(objMatrix2.VisualRowCount - 1).Specific.Value.Trim <> "" Then
                                        oDBs_Detail4.SetValue("U_VSPADACC", oDBs_Detail4.Offset, objMatrix2.Columns.Item("V_4").Cells.Item(objMatrix2.VisualRowCount - 1).Specific.Value.Trim)
                                    Else
                                        oDBs_Detail4.SetValue("U_VSPADACC", oDBs_Detail4.Offset, objMatrix4.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                    End If
                                Else
                                    oDBs_Detail4.SetValue("U_VSPADACC", oDBs_Detail4.Offset, objMatrix4.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                End If

                            End If

                            oDBs_Detail4.SetValue("U_VSPAMT", oDBs_Detail4.Offset, 0)
                            oDBs_Detail4.SetValue("U_VSPBUD", oDBs_Detail4.Offset, 0)
                            oDBs_Detail4.SetValue("U_VSPMTYP", oDBs_Detail4.Offset, objMatrix4.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix4.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix4.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000024")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_CHITEM" Then
                            oDBs_Detail6.SetValue("LineId", oDBs_Detail6.Offset, pVal.Row)
                            oDBs_Detail6.SetValue("U_VSPDAT", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPSOU", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPSOUR", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPFRDT", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPFRTM", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPTODT", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPTOTM", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPCHCOD", oDBs_Detail6.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Detail6.SetValue("U_VSPCHNAM", oDBs_Detail6.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Detail6.SetValue("U_VSPWEIGH", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail6.SetValue("U_VSPUOM", oDBs_Detail6.Offset, objMatrix6.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix6.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix6.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000028")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_DRC" Then
                            oDBs_Detail2.SetValue("LineId", oDBs_Detail2.Offset, pVal.Row)
                            oDBs_Detail2.SetValue("U_VSPJOUDT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPAMGOO", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPADAMT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPFRACT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VAPTOACT", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPDRVCC", oDBs_Detail2.Offset, oDT.GetValue("PrcCode", 0))
                            oDBs_Detail2.SetValue("U_VAPCAS", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPCOM", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPJENO", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail2.SetValue("U_VSPOPNO", oDBs_Detail2.Offset, objMatrix2.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix2.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_DRC1" Then
                            oDBs_Detail5.SetValue("LineId", oDBs_Detail5.Offset, pVal.Row)
                            oDBs_Detail5.SetValue("U_VSPDATE", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPVENCO", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPVENNM", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPQUAN", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPRATE", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPAMT", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPDRCC1", oDBs_Detail5.Offset, oDT.GetValue("PrcCode", 0))
                            oDBs_Detail5.SetValue("U_VSPDCNUM", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPGISU", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            objMatrix5.SetLineData(pVal.Row)
                        End If

                        If oCFL.UniqueID = "CFL_VEN" Then
                            oDBs_Detail5.SetValue("LineId", oDBs_Detail5.Offset, pVal.Row)
                            oDBs_Detail5.SetValue("U_VSPDATE", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail5.SetValue("U_VSPVENCO", oDBs_Detail5.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Detail5.SetValue("U_VSPVENNM", oDBs_Detail5.Offset, oDT.GetValue("CardName", 0))
                            oDBs_Detail5.SetValue("U_VSPQUAN", oDBs_Detail5.Offset, "")
                            oDBs_Detail5.SetValue("U_VSPRATE", oDBs_Detail5.Offset, 0)
                            oDBs_Detail5.SetValue("U_VSPAMT", oDBs_Detail5.Offset, 0)
                            oDBs_Detail5.SetValue("U_VSPDCNUM", oDBs_Detail5.Offset, "")
                            objMatrix5.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix5.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "67")
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_SO" Then
                            oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, Row)
                            oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, CDate(oDT.GetValue("DocDate", 0)).ToString("yyyyMMdd"))
                            oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, oDT.GetValue("DocEntry", 0))
                            oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, oDT.GetValue("NumAtCard", 0))
                            'Dim CardCode As String = oDT.GetValue("CardCode", 0)
                            'CardCode = CardCode.Substring(0, 3)
                            'If CardCode = "DTP" Then
                            '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, oDT.GetValue("DocTotal", 0))
                            'ElseIf CardCode = "DRM" Then

                            '    Dim GetFTot As String = "Select T0.Quantity,T1.TotalSumSy From RDR1 T0 inner join RDR3 T1 on T0.DocEntry=T1.DocEntry Where T1.DocEntry='" & oDT.GetValue("DocEntry", 0) & "' And T1.ExpnsCode='3'"
                            '    Dim oRsGetFTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRsGetFTot.DoQuery(GetFTot)
                            '    Dim Qty As Double = oRsGetFTot.Fields.Item(0).Value
                            '    Dim DCTot As Double = Qty * oRsGetFTot.Fields.Item(1).Value
                            '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, DCTot)
                            'End If
                            Dim GetTot As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetTot = "Select Sum(T0.""Quantity""*T0.""U_VSPUNPRC"") as ""Total"" From RDR1 T0 Where T0.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "'"
                            Else
                                GetTot = "Select Sum(T0.Quantity*T0.U_VSPUNPRC) as Total From RDR1 T0 Where T0.DocEntry='" & oDT.GetValue("DocEntry", 0) & "'"
                            End If
                            Dim oRsGetTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetTot.DoQuery(GetTot)

                            Dim Tot As Double = oRsGetTot.Fields.Item(0).Value

                            oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, Tot)


                            oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, oDT.GetValue("Comments", 0))
                            objMatrix3.SetLineData(Row)
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim oOrders As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                            If oOrders.GetByKey(oDT.GetValue("DocEntry", 0)) Then
                                oOrders.UserFields.Fields.Item("U_VSPFLSTS").Value = "Linked to Fleet"
                                If oOrders.Update() <> 0 Then
                                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                                End If
                            End If
                            ' Me.SetCellsEditable(objForm.UniqueID)
                        End If

                        If oCFL.UniqueID = "CFL_DLV" Then
                            oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, Row)
                            oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, CDate(oDT.GetValue("DocDate", 0)).ToString("yyyyMMdd"))
                            oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, oDT.GetValue("DocEntry", 0))
                            oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, oDT.GetValue("NumAtCard", 0))
                            'Dim CardCode As String = oDT.GetValue("CardCode", 0)
                            'CardCode = CardCode.Substring(0, 3)
                            'If CardCode = "DTP" Then
                            '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, oDT.GetValue("DocTotal", 0))
                            'ElseIf CardCode = "DRM" Then

                            '    Dim GetFTot As String = "Select T0.Quantity,T1.TotalSumSy From RDR1 T0 inner join RDR3 T1 on T0.DocEntry=T1.DocEntry Where T1.DocEntry='" & oDT.GetValue("DocEntry", 0) & "' And T1.ExpnsCode='3'"
                            '    Dim oRsGetFTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRsGetFTot.DoQuery(GetFTot)
                            '    Dim Qty As Double = oRsGetFTot.Fields.Item(0).Value
                            '    Dim DCTot As Double = Qty * oRsGetFTot.Fields.Item(1).Value
                            '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, DCTot)
                            'End If
                            Dim GetTot As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetTot = "Select Sum(T0.""Quantity""*T0.""U_VSPUNPRC"") as ""Total"" From DLN1 T0 Where T0.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "'"
                            Else
                                GetTot = "Select Sum(T0.Quantity*T0.U_VSPUNPRC) as Total From DLN1 T0 Where T0.DocEntry='" & oDT.GetValue("DocEntry", 0) & "'"
                            End If
                            Dim oRsGetTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetTot.DoQuery(GetTot)

                            Dim Tot As Double = oRsGetTot.Fields.Item(0).Value

                            oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, Tot)


                            oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, oDT.GetValue("Comments", 0))
                            objMatrix3.SetLineData(Row)
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim oDelivery As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                            If oDelivery.GetByKey(oDT.GetValue("DocEntry", 0)) Then
                                oDelivery.UserFields.Fields.Item("U_VSPFLSTS").Value = "Linked to Fleet"
                                If oDelivery.Update() <> 0 Then
                                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                                End If
                            End If
                            ' Me.SetCellsEditable(objForm.UniqueID)
                        End If

                        If oCFL.UniqueID = "CFL_INV" Then
                            oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, Row)
                            oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(Row).Specific.Value)
                            oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, CDate(oDT.GetValue("DocDate", 0)).ToString("yyyyMMdd"))
                            oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, oDT.GetValue("DocEntry", 0))
                            oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, oDT.GetValue("NumAtCard", 0))
                            Dim CardCode As String = oDT.GetValue("CardCode", 0)
                            CardCode = CardCode.Substring(0, 3)
                            'If CardCode = "DTP" Then
                            '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, oDT.GetValue("DocTotal", 0))
                            'ElseIf CardCode = "DRM" Then

                            '    Dim GetFTot As String = ""
                            '    If objMain.IsSAPHANA = True Then
                            '        GetFTot = "Select T0.""Quantity"",T1.""TotalSumSy"" From INV1 T0 inner join INV3 T1 on T0.""DocEntry""=T1.""DocEntry"" Where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' And T1.""ExpnsCode""='3'"
                            '    Else
                            '        GetFTot = "Select T0.Quantity,T1.TotalSumSy From INV1 T0 inner join INV3 T1 on T0.DocEntry=T1.DocEntry Where T1.DocEntry='" & oDT.GetValue("DocEntry", 0) & "' And T1.ExpnsCode='3'"
                            '    End If
                            '    Dim oRsGetFTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            '    oRsGetFTot.DoQuery(GetFTot)
                            '    Dim Qty As Double = oRsGetFTot.Fields.Item(0).Value
                            '    Dim DCTot As Double = Qty * oRsGetFTot.Fields.Item(1).Value
                            '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, DCTot)
                            'End If
                            Dim GetTot As String = ""
                            If objMain.IsSAPHANA = True Then
                                GetTot = "Select Sum(T0.""Quantity""*T0.""U_VSPUNPRC"") as ""Total"" From INV1 T0 Where T0.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "'"
                            Else
                                GetTot = "Select Sum(T0.Quantity*T0.U_VSPUNPRC) as Total From INV1 T0 Where T0.DocEntry='" & oDT.GetValue("DocEntry", 0) & "'"
                            End If
                            Dim oRsGetTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetTot.DoQuery(GetTot)

                            Dim Tot As Double = oRsGetTot.Fields.Item(0).Value

                            oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, Tot)

                            oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, oDT.GetValue("Comments", 0))
                            objMatrix3.SetLineData(Row)
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim oInvoice As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                            If oInvoice.GetByKey(oDT.GetValue("DocEntry", 0)) Then
                                oInvoice.UserFields.Fields.Item("U_VSPFLSTS").Value = "Linked to Fleet"
                                If oInvoice.Update() <> 0 Then
                                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                                End If
                            End If
                            Me.SetCellsEditable(objForm.UniqueID)
                        End If

                            If oCFL.UniqueID = "CFL_PO" Then
                                oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, Row)
                                oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, CDate(oDT.GetValue("DocDate", 0)).ToString("yyyyMMdd"))
                                oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, oDT.GetValue("DocEntry", 0))
                                oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, oDT.GetValue("CardCode", 0))
                                oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, oDT.GetValue("NumAtCard", 0))

                                Dim GetDcTot As String = ""
                                If objMain.IsSAPHANA = True Then
                                    GetDcTot = "Select T0.""Quantity"",T1.""TotalSumSy"" From ""PDN1"" T0 inner join ""PDN3"" T1 on T0.""DocEntry""=T1.""DocEntry"" Where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' And T1.""ExpnsCode""='3'"
                                Else
                                    GetDcTot = "Select T0.Quantity,T1.TotalSumSy From PDN1 T0 inner join PDN3 T1 on T0.DocEntry=T1.DocEntry Where T1.DocEntry='" & oDT.GetValue("DocEntry", 0) & "' And T1.ExpnsCode='3'"
                                End If


                                Dim oRsGetDcTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetDcTot.DoQuery(GetDcTot)
                                If oRsGetDcTot.RecordCount > 0 Then
                                    Dim Qty As Double = oRsGetDcTot.Fields.Item(0).Value
                                    Dim DCTot As Double = Qty * oRsGetDcTot.Fields.Item(1).Value
                                    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, DCTot)
                                End If
                                oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, oDT.GetValue("Comments", 0))
                                objMatrix3.SetLineData(Row)
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Dim oPurchaseOrder As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                If oPurchaseOrder.GetByKey(oDT.GetValue("DocEntry", 0)) Then
                                    oPurchaseOrder.UserFields.Fields.Item("U_VSPFLSTS").Value = "Linked to Fleet"
                                    If oPurchaseOrder.Update() <> 0 Then
                                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                                    End If
                                End If
                                Me.SetCellsEditable(objForm.UniqueID)
                            End If

                            If oCFL.UniqueID = "CFL_APINV" Then
                                oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, Row)
                                oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, CDate(oDT.GetValue("DocDate", 0)).ToString("yyyyMMdd"))
                                oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, oDT.GetValue("DocEntry", 0))
                                oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, oDT.GetValue("CardCode", 0))
                                oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, oDT.GetValue("NumAtCard", 0))
                                Dim GetDcTot As String = ""
                                If objMain.IsSAPHANA = True Then
                                    GetDcTot = "Select T0.""Quantity"",T1.""TotalSumSy"" From ""PCH1"" T0 inner join ""PCH3"" T1 on T0.""DocEntry""=T1.""DocEntry"" Where T1.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "' And T1.""ExpnsCode""='3'"
                                Else
                                    GetDcTot = "Select T0.Quantity,T1.TotalSumSy From PCH1 T0 inner join PCH3 T1 on T0.DocEntry=T1.DocEntry Where T1.DocEntry='" & oDT.GetValue("DocEntry", 0) & "' And T1.ExpnsCode='3'"
                                End If

                                Dim oRsGetDcTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetDcTot.DoQuery(GetDcTot)
                                If oRsGetDcTot.RecordCount > 0 Then
                                    Dim Qty As Double = oRsGetDcTot.Fields.Item(0).Value
                                    Dim DCTot As Double = Qty * oRsGetDcTot.Fields.Item(1).Value
                                    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, DCTot)
                                End If
                                oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, oDT.GetValue("Comments", 0))
                                objMatrix3.SetLineData(Row)

                                Dim oPurchaseInvoice As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                If oPurchaseInvoice.GetByKey(oDT.GetValue("DocEntry", 0)) Then
                                    oPurchaseInvoice.UserFields.Fields.Item("U_VSPFLSTS").Value = "Linked to Fleet"
                                    If oPurchaseInvoice.Update() <> 0 Then
                                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                                    End If
                                End If
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Me.SetCellsEditable(objForm.UniqueID)
                            End If

                            If oCFL.UniqueID = "CFL_PO1" Then
                                oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, Row)
                                oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(Row).Specific.Value)
                                oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, CDate(oDT.GetValue("DocDate", 0)).ToString("yyyyMMdd"))
                                oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, oDT.GetValue("DocEntry", 0))
                                oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, oDT.GetValue("CardCode", 0))
                                oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, oDT.GetValue("NumAtCard", 0))
                                'Dim GetDcTot As String = "Select T0.Quantity,T1.TotalSumSy From POR1 T0 inner join POR3 T1 on T0.DocEntry=T1.DocEntry Where T1.DocEntry='" & oDT.GetValue("DocEntry", 0) & "' And T1.ExpnsCode='3'"
                                'Dim oRsGetDcTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'oRsGetDcTot.DoQuery(GetDcTot)
                                'If oRsGetDcTot.RecordCount > 0 Then
                                '    Dim Qty As Double = oRsGetDcTot.Fields.Item(0).Value
                                '    Dim DCTot As Double = Qty * oRsGetDcTot.Fields.Item(1).Value
                                '    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, DCTot)
                                'End If

                                Dim GetTot As String = ""

                                If objMain.IsSAPHANA = True Then
                                    GetTot = "Select Sum(T0.""Quantity""*T0.""U_VSPUNPRC"") as ""Total"" From POR1 T0 Where T0.""DocEntry""='" & oDT.GetValue("DocEntry", 0) & "'"
                                Else
                                    GetTot = "Select Sum(T0.Quantity*T0.U_VSPUNPRC) as Total From POR1 T0 Where T0.DocEntry='" & oDT.GetValue("DocEntry", 0) & "'"
                                End If
                                Dim oRsGetTot As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetTot.DoQuery(GetTot)

                                Dim Tot As Double = oRsGetTot.Fields.Item(0).Value

                                oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, Tot)

                                oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, oDT.GetValue("Comments", 0))
                                objMatrix3.SetLineData(Row)
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                Dim oPurchaseOrder As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                If oPurchaseOrder.GetByKey(oDT.GetValue("DocEntry", 0)) Then
                                    oPurchaseOrder.UserFields.Fields.Item("U_VSPFLSTS").Value = "Linked to Fleet"
                                    If oPurchaseOrder.Update() <> 0 Then
                                        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                                    End If
                                End If
                                Me.SetCellsEditable(objForm.UniqueID)
                            End If
                        End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("61").Specific
                    objMatrix5 = objForm.Items.Item("67").Specific
                    objMatrix7 = objForm.Items.Item("69").Specific

                    oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
                    oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")
                    oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")

                    If pVal.ItemUID = "61" And pVal.ColUID = "V_1" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim OpenKM As Double = 0
                        Dim CloseKM As Double = 0

                        If objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            OpenKM = objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value
                        End If
                        If objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            CloseKM = objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value
                        End If

                        If CloseKM < OpenKM And CloseKM > 0 Then
                            BubbleEvent = False
                            objMain.objApplication.StatusBar.SetText("Close KM's Cannot Be Less than Open KM", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, pVal.Row)
                            oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, CloseKM - OpenKM)
                            objMatrix1.SetLineData(pVal.Row)
                        End If
                    End If

                    If pVal.ItemUID = "67" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_4") And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim Quantity As Double = 0
                        Dim Rate As Double = 0
                        If objMatrix5.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Quantity = objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value
                        End If
                        If objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Rate = objMatrix5.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value
                        End If

                        oDBs_Detail5.SetValue("LineId", oDBs_Detail5.Offset, pVal.Row)
                        oDBs_Detail5.SetValue("U_VSPDATE", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Detail5.SetValue("U_VSPVENCO", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Detail5.SetValue("U_VSPVENNM", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                        oDBs_Detail5.SetValue("U_VSPQUAN", oDBs_Detail5.Offset, Quantity)
                        oDBs_Detail5.SetValue("U_VSPRATE", oDBs_Detail5.Offset, Rate)
                        oDBs_Detail5.SetValue("U_VSPAMT", oDBs_Detail5.Offset, Quantity * Rate)
                        oDBs_Detail5.SetValue("U_VSPDCNUM", oDBs_Detail5.Offset, objMatrix5.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                        objMatrix5.SetLineData(pVal.Row)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("61").Specific
                    objMatrix3 = objForm.Items.Item("1000023").Specific

                    oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")

                    If pVal.ItemUID = "1000023" And pVal.ColUID = "V_2" And pVal.BeforeAction = True Then
                        Row = pVal.Row
                    End If

                   

                    If pVal.ItemUID = "1000036" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        If objForm.Items.Item("1000036").Specific.Value.ToString.Trim <> "" Then
                            Dim s As String = objForm.Items.Item("1000036").Specific.Value
                            oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, objMatrix1.VisualRowCount)
                            oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, objForm.Items.Item("1000036").Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                            objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                            Me.SetNewLine(objForm.UniqueID, "61", objMatrix1.Columns.Item("V_1").Cells.Item(objMatrix1.VisualRowCount).Specific.Value, _
                                                                                    objMatrix1.Columns.Item("V_3").Cells.Item(objMatrix1.VisualRowCount).Specific.Value)
                        End If
                    End If

                    If pVal.ItemUID = "61" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                        If pVal.ColUID = "V_3" Then
                            If objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                If pVal.Row = objMatrix1.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "61", objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value, _
                                                                                            objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                Else
                                    oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, pVal.Row + 1)
                                    oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, "0")
                                    oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                    oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, "")
                                    oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, "")
                                    objMatrix1.SetLineData(pVal.Row + 1)
                                End If
                            End If
                        End If

                        If pVal.ColUID = "V_8" Then
                            If objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value <> "" Then

                                Dim GetMileageDetails As String = ""
                                If objMain.IsSAPHANA = True Then
                                    GetMileageDetails = "Select ""U_VSPMWL"" , ""U_VSPMWOL"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("4").Specific.Value & "'"
                                Else
                                    GetMileageDetails = "Select U_VSPMWL , U_VSPMWOL From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("4").Specific.Value & "'"
                                End If
                                Dim oRsGetMileageDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetMileageDetails.DoQuery(GetMileageDetails)

                                Dim DieselConsumption As Double = objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value
                                If objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value = "Yes" Then
                                    DieselConsumption = DieselConsumption / oRsGetMileageDetails.Fields.Item(0).Value
                                Else
                                    DieselConsumption = DieselConsumption / oRsGetMileageDetails.Fields.Item(1).Value
                                End If

                                oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, pVal.Row)
                                oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, DieselConsumption)
                                oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                                objMatrix1.SetLineData(pVal.Row)
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "69" And pVal.ColUID = "V_0" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then

                        If pVal.BeforeAction = False Then
                            For i As Integer = 1 To objMatrix7.VisualRowCount
                                If i <> pVal.Row Then
                                    If objMatrix7.Columns.Item("V_0").Cells.Item(i).Specific.Value = objMatrix7.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value Then
                                        If pVal.Row > 1 Then
                                            objMain.objApplication.StatusBar.SetText("Selected Driver already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                            oDBs_Detail7.SetValue("LineId", oDBs_Detail7.Offset, pVal.Row)
                                            oDBs_Detail7.SetValue("U_VSPDRCOD", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPDRFNM", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPDRMNM", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPDRLNM", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPMBNUM", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPLNNUM", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPEXPDT", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPFRMDT", oDBs_Detail7.Offset, "")
                                            oDBs_Detail7.SetValue("U_VSPTODT", oDBs_Detail7.Offset, "")

                                            objMatrix7.SetLineData(pVal.Row)
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                            Next
                        End If

                        If pVal.BeforeAction = False Then
                            If objMatrix7.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                Dim GetDriverDetails As String = ""

                                If objMain.IsSAPHANA = True Then
                                    GetDriverDetails = "Select ""U_VSPFNAME"" , ""U_VSPMNAME"" , ""U_VSPLNAME"" , ""U_VSPMOBNO"" , ""U_VSPNUM"" , ""U_VSPEXPDT"" " & _
                               "From ""@VSP_FLT_DRVRMSTR"" Where ""Code"" = '" & objMatrix7.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value & "'"
                                Else
                                    GetDriverDetails = "Select U_VSPFNAME , U_VSPMNAME , U_VSPLNAME , U_VSPMOBNO , U_VSPNUM , U_VSPEXPDT " & _
                               "From [@VSP_FLT_DRVRMSTR] Where Code = '" & objMatrix7.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value & "'"
                                End If
                                Dim oRsGetDriverDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetDriverDetails.DoQuery(GetDriverDetails)

                                oDBs_Detail7.SetValue("LineId", oDBs_Detail7.Offset, pVal.Row)
                                oDBs_Detail7.SetValue("U_VSPDRCOD", oDBs_Detail7.Offset, objMatrix7.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail7.SetValue("U_VSPDRFNM", oDBs_Detail7.Offset, oRsGetDriverDetails.Fields.Item("U_VSPFNAME").Value)
                                oDBs_Detail7.SetValue("U_VSPDRMNM", oDBs_Detail7.Offset, oRsGetDriverDetails.Fields.Item("U_VSPMNAME").Value)
                                oDBs_Detail7.SetValue("U_VSPDRLNM", oDBs_Detail7.Offset, oRsGetDriverDetails.Fields.Item("U_VSPLNAME").Value)
                                oDBs_Detail7.SetValue("U_VSPMBNUM", oDBs_Detail7.Offset, oRsGetDriverDetails.Fields.Item("U_VSPMOBNO").Value)
                                oDBs_Detail7.SetValue("U_VSPLNNUM", oDBs_Detail7.Offset, oRsGetDriverDetails.Fields.Item("U_VSPNUM").Value)
                                oDBs_Detail7.SetValue("U_VSPEXPDT", oDBs_Detail7.Offset, CDate(oRsGetDriverDetails.Fields.Item("U_VSPEXPDT").Value).ToString("yyyyMMdd"))
                                oDBs_Detail7.SetValue("U_VSPFRMDT", oDBs_Detail7.Offset, objMatrix7.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Detail7.SetValue("U_VSPTODT", oDBs_Detail7.Offset, objMatrix7.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)

                                objMatrix7.SetLineData(pVal.Row)
                                If pVal.Row = objMatrix7.VisualRowCount Then
                                    Me.SetNewLine(objForm.UniqueID, "69")
                                End If
                            End If
                        End If

                    End If

                    If pVal.ItemUID = "1000023" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        If pVal.ColUID = "V_2" And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If objMatrix3.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value = "Existing" Then
                                If objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Sales Order" Then
                                    objForm.Items.Item("CFLSO").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "A/R Invoice" Then
                                    objForm.Items.Item("CFLINV").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Delivery" Then
                                    objForm.Items.Item("CFLDLV").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "G.Receipt PO" Then
                                    '    objForm.Items.Item("CFLPO").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "A/P Invoice" Then
                                    '    objForm.Items.Item("CFLAPINV").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Purchase Order" Then
                                    objForm.Items.Item("CFLPO1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                End If

                            End If
                        End If
                    End If

                    If pVal.ItemUID = "1000023" And pVal.ColUID = "V_0" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        If objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value.Trim = "Sales" Then
                            If (objMatrix3.Columns.Item("V_1").ValidValues.Count <> 0) Then
                                For R As Integer = objMatrix3.Columns.Item("V_1").ValidValues.Count - 1 To 0 Step -1
                                    Try
                                        objMatrix3.Columns.Item("V_1").ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Catch ex As Exception
                                    End Try
                                Next
                            End If

                            objMatrix3.Columns.Item("V_1").ValidValues.Add("", "")
                            objMatrix3.Columns.Item("V_1").ValidValues.Add("Sales Order", "Sales Order")
                            objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        ElseIf objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value.Trim = "Purchase" Then
                            If (objMatrix3.Columns.Item("V_1").ValidValues.Count <> 0) Then
                                For R As Integer = objMatrix3.Columns.Item("V_1").ValidValues.Count - 1 To 0 Step -1
                                    Try
                                        objMatrix3.Columns.Item("V_1").ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Catch ex As Exception
                                    End Try
                                Next
                            End If

                            objMatrix3.Columns.Item("V_1").ValidValues.Add("", "")
                            objMatrix3.Columns.Item("V_1").ValidValues.Add("Purchase Order", "Purchase Order")
                            objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        ElseIf objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value.Trim = "Delivery" Then
                            If (objMatrix3.Columns.Item("V_1").ValidValues.Count <> 0) Then
                                For R As Integer = objMatrix3.Columns.Item("V_1").ValidValues.Count - 1 To 0 Step -1
                                    Try
                                        objMatrix3.Columns.Item("V_1").ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Catch ex As Exception
                                    End Try
                                Next
                            End If

                            objMatrix3.Columns.Item("V_1").ValidValues.Add("", "")
                            objMatrix3.Columns.Item("V_1").ValidValues.Add("Delivery", "Delivery")
                            objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        ElseIf objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value.Trim = "A/R Invoice" Then
                            If (objMatrix3.Columns.Item("V_1").ValidValues.Count <> 0) Then
                                For R As Integer = objMatrix3.Columns.Item("V_1").ValidValues.Count - 1 To 0 Step -1
                                    Try
                                        objMatrix3.Columns.Item("V_1").ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Catch ex As Exception
                                    End Try
                                Next
                            End If

                            objMatrix3.Columns.Item("V_1").ValidValues.Add("", "")
                            objMatrix3.Columns.Item("V_1").ValidValues.Add("A/R Invoice", "A/R Invoice")
                            objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                        Else
                            If (objMatrix3.Columns.Item("V_1").ValidValues.Count <> 0) Then
                                For R As Integer = objMatrix3.Columns.Item("V_1").ValidValues.Count - 1 To 0 Step -1
                                    Try
                                        objMatrix3.Columns.Item("V_1").ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Catch ex As Exception
                                    End Try
                                Next
                            End If
                            objMatrix3.Columns.Item("V_1").ValidValues.Add("", "")
                            objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Select("", SAPbouiCOM.BoSearchKey.psk_Index)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix3 = objForm.Items.Item("1000023").Specific

                    If pVal.ItemUID = "1000023" And pVal.ColUID = "V_5" And pVal.BeforeAction = True Then
                        If objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Sales Order" Then
                            oLink = objMatrix3.Columns.Item("V_5").ExtendedObject
                            oLink.LinkedObjectType = 17
                           

                        ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "A/R Invoice" Then
                            oLink = objMatrix3.Columns.Item("V_5").ExtendedObject
                            oLink.LinkedObjectType = 13
                        ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Delivery" Then
                            oLink = objMatrix3.Columns.Item("V_5").ExtendedObject
                            oLink.LinkedObjectType = 15
                            'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "G.Receipt PO" Then
                            '    oLink = objMatrix3.Columns.Item("V_5").ExtendedObject
                            '    oLink.LinkedObjectType = 20
                            'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "A/P Invoice" Then
                            '    oLink = objMatrix3.Columns.Item("V_5").ExtendedObject
                            '    oLink.LinkedObjectType = 18
                        ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Purchase Order" Then
                            oLink = objMatrix3.Columns.Item("V_5").ExtendedObject
                            oLink.LinkedObjectType = 22
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Generate") = True Then
                            objMain.objApplication.Menus.RemoveEx("Generate")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                    Try
                        If oMenus.Exists("Generate Delivery") = True Then
                            objMain.objApplication.Menus.RemoveEx("Generate Delivery")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                    Try
                        If oMenus.Exists("Generate Invoice") = True Then
                            objMain.objApplication.Menus.RemoveEx("Generate Invoice")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Generate") = True Then
                            objMain.objApplication.Menus.RemoveEx("Generate")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                    Try
                        If oMenus.Exists("Generate Delivery") = True Then
                            objMain.objApplication.Menus.RemoveEx("Generate Delivery")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                    Try
                        If oMenus.Exists("Generate Invoice") = True Then
                            objMain.objApplication.Menus.RemoveEx("Generate Invoice")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Me.RefreshData(objForm.UniqueID)
                        Me.SetCellsEditable(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)


                    If pVal.ItemUID = "1000030" And pVal.ColUID = "V_1" And pVal.BeforeAction = False Then
                        objMatrix9 = objForm.Items.Item("1000030").Specific
                        Path = pVal.Row
                        Me.BrowseFileDialog()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix4 = objForm.Items.Item("1000024").Specific
                    objMatrix2 = objForm.Items.Item("1000022").Specific
                    objMatrix3 = objForm.Items.Item("1000023").Specific
                    If pVal.ItemUID = "1000024" And pVal.ColUID = "V_5" And pVal.BeforeAction = False Then
                        If objMatrix4.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value <> 0 And objMatrix4.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix4.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000024")
                            End If
                        End If
                    End If

                    ''Added on 16-10-2018 by Abinas

                    If pVal.ItemUID = "1000022" And pVal.ColUID = "V_2" And pVal.BeforeAction = False Then
                        If objMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value <> 0.0 And objMatrix2.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value <> 0.0 And objMatrix2.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix2.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000022")
                            End If
                        End If
                    End If


                    If pVal.ItemUID = "1000023" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        If objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Sales Order" Then
                                Try

                                    Dim GetQuantity As String = ""
                                    If objMain.IsSAPHANA = True Then
                                        GetQuantity = "Select ""Quantity"" from RDR1 where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value & "'"
                                    Else
                                        GetQuantity = "Select ""Quantity"" from RDR1 where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value & "'"
                                    End If

                                    Dim orsGetQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orsGetQuantity.DoQuery(GetQuantity)
                                    Dim Quantity As Double = CDbl(orsGetQuantity.Fields.Item(0).Value)
                                    Dim EnterQuantity As Double = CDbl(objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                    If Quantity <> EnterQuantity Then
                                        objMain.objApplication.StatusBar.SetText("Quantity Cannot Be Not Change", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = Quantity

                                    End If
                                Catch ex As Exception

                                End Try

                            ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "Delivery" Then
                                Try
                                    Dim GetQuantity As String = ""
                                    If objMain.IsSAPHANA = True Then
                                        GetQuantity = "Select ""Quantity"" from DLN1 where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value & "'"
                                    Else
                                        GetQuantity = "Select ""Quantity"" from DLN1 where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value & "'"
                                    End If
                                    Dim orsGetQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orsGetQuantity.DoQuery(GetQuantity)
                                    Dim Quantity As Double = CDbl(orsGetQuantity.Fields.Item(0).Value)
                                    Dim EnterQuantity As Double = CDbl(objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                    If Quantity <> EnterQuantity Then
                                        objMain.objApplication.StatusBar.SetText("Quantity Cannot Be Not Change", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = Quantity

                                    End If
                                Catch ex As Exception

                                End Try

                            ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value = "A/R Invoice" Then
                                Try
                                    Dim GetQuantity As String = ""
                                    If objMain.IsSAPHANA = True Then
                                        GetQuantity = "Select ""Quantity"" from INV1 where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value & "'"
                                    Else
                                        GetQuantity = "Select ""Quantity"" from INV1 where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value & "'"
                                    End If
                                    Dim orsGetQuantity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    orsGetQuantity.DoQuery(GetQuantity)
                                    Dim Quantity As Double = CDbl(orsGetQuantity.Fields.Item(0).Value)
                                    Dim EnterQuantity As Double = CDbl(objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                    If Quantity <> EnterQuantity Then
                                        objMain.objApplication.StatusBar.SetText("Quantity Cannot Be Not Change", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objMatrix3.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = Quantity

                                    End If
                                Catch ex As Exception

                                End Try
                            End If
                        End If

                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TRSHT" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Generate" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TRSHT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
                objMatrix3 = objForm.Items.Item("1000023").Specific

                Dim i As Integer
                For i = 1 To objMatrix3.VisualRowCount
                    If objMatrix3.IsRowSelected(i) = True Then Exit For
                Next

                ' If objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value = "Sales" Then
                'Me.GenerateInvoice(objForm.UniqueID, i)
                Dim otherForm As SAPbouiCOM.Form
                If objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "Sales Order" Then
                    objMain.objApplication.ActivateMenuItem("2050")
                    otherForm = objMain.objApplication.Forms.GetForm("139", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    otherForm.Items.Item("txt_DocTyp").Specific.Value = "Trip Sheet"
                    otherForm.Items.Item("txt_DocNum").Specific.Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                    If Fleet.MainCls.ohtLookUpForm.ContainsKey(otherForm.UniqueID) = False Then
                        Fleet.MainCls.ohtLookUpForm.Add(otherForm.UniqueID, objForm.UniqueID)
                    End If
                    'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "A/R Invoice" Then
                    '    objMain.objApplication.ActivateMenuItem("2053")
                    '    otherForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    '    otherForm.Items.Item("txt_DocTyp").Specific.Value = "Trip Sheet"
                    '    otherForm.Items.Item("txt_DocNum").Specific.Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                    '    If Fleet.MainCls.ohtLookUpForm.ContainsKey(otherForm.UniqueID) = False Then
                    '        Fleet.MainCls.ohtLookUpForm.Add(otherForm.UniqueID, objForm.UniqueID)
                    '    End If
                    'End If
                    'ElseIf objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value = "Purchase" Then
                    'Dim otherForm As SAPbouiCOM.Form
                ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "Purchase Order" Then
                    objMain.objApplication.ActivateMenuItem("2305")
                    otherForm = objMain.objApplication.Forms.GetForm("142", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    otherForm.Items.Item("txt_DocTyp").Specific.Value = "Trip Sheet"
                    otherForm.Items.Item("txt_DocNum").Specific.Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                    If Fleet.MainCls.ohtLookUpForm.ContainsKey(otherForm.UniqueID) = False Then
                        Fleet.MainCls.ohtLookUpForm.Add(otherForm.UniqueID, objForm.UniqueID)
                    End If
                    'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "G.Receipt PO" Then
                    '    objMain.objApplication.ActivateMenuItem("2306")
                    '    otherForm = objMain.objApplication.Forms.GetForm("143", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    '    otherForm.Items.Item("txt_DocTyp").Specific.Value = "Trip Sheet"
                    '    otherForm.Items.Item("txt_DocNum").Specific.Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                    '    If Fleet.MainCls.ohtLookUpForm.ContainsKey(otherForm.UniqueID) = False Then
                    '        Fleet.MainCls.ohtLookUpForm.Add(otherForm.UniqueID, objForm.UniqueID)
                    '    End If
                    'ElseIf objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "A/P Invoice" Then
                    '    objMain.objApplication.ActivateMenuItem("2308")
                    '    otherForm = objMain.objApplication.Forms.GetForm("141", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    '    otherForm.Items.Item("txt_DocTyp").Specific.Value = "Trip Sheet"
                    '    otherForm.Items.Item("txt_DocNum").Specific.Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                    '    If Fleet.MainCls.ohtLookUpForm.ContainsKey(otherForm.UniqueID) = False Then
                    '        Fleet.MainCls.ohtLookUpForm.Add(otherForm.UniqueID, objForm.UniqueID)
                    '    End If
                End If
                ' End If



            ElseIf pVal.MenuUID = "Generate Delivery" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TRSHT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
                objMatrix3 = objForm.Items.Item("1000023").Specific
                Dim i As Integer
                For i = 1 To objMatrix3.VisualRowCount
                    If objMatrix3.IsRowSelected(i) = True Then Exit For
                Next
                'If objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "Delivery" Then
                Try
                    Dim SoEntry As String = objMatrix3.Columns.Item("V_5").Cells.Item(i).Specific.Value

                    Dim getSOSetalis As String = "Select * from ORDR T1 Inner Join RDR1 T2 on T1.""DocEntry""=T2.""DocEntry"" where T2.""DocEntry""='" & SoEntry & "'"
                    Dim orsgetSOSetalis As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    orsgetSOSetalis.DoQuery(getSOSetalis)
                    If orsgetSOSetalis.RecordCount > 0 Then

                        orsgetSOSetalis.MoveFirst()

                        Dim oDeliveryOrder As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                        oDeliveryOrder.CardCode = orsgetSOSetalis.Fields.Item("CardCode").Value
                        oDeliveryOrder.DocDate = orsgetSOSetalis.Fields.Item("DocDate").Value
                        oDeliveryOrder.DocDueDate = orsgetSOSetalis.Fields.Item("DocDueDate").Value
                        oDeliveryOrder.TaxDate = orsgetSOSetalis.Fields.Item("TaxDate").Value
                        oDeliveryOrder.UserFields.Fields.Item("U_VSPDCTYP").Value = "Trip Sheet"
                        oDeliveryOrder.UserFields.Fields.Item("U_VSPDCNO").Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                        Dim LineCount As Integer = 0
                        For j As Integer = 1 To orsgetSOSetalis.RecordCount
                            If LineCount > 0 Then
                                oDeliveryOrder.Lines.Add()
                            End If

                            'oDeliveryOrder.Lines.ItemCode = orsgetSOSetalis.Fields.Item("DocEntry").Value
                            'oDeliveryOrder.Lines.ItemDescription = orsgetSOSetalis.Fields.Item("Dscription").Value
                            'oDeliveryOrder.Lines.Quantity = objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value
                            'oDeliveryOrder.Lines.Price = orsgetSOSetalis.Fields.Item("Price").Value
                            'oDeliveryOrder.Lines.WarehouseCode = orsgetSOSetalis.Fields.Item("WhsCode").Value
                            'oDeliveryOrder.Lines.TaxCode = orsgetSOSetalis.Fields.Item("TaxCode").Value

                            'If CDbl(objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value) > CDbl(orsgetSOSetalis.Fields.Item("Quantity").Value) Then
                            '    objMain.objApplication.StatusBar.SetText("Delivery Quantity is Not Gratter Than From Sales Order Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            '    Exit Try
                            'End If

                            oDeliveryOrder.Lines.Quantity = CDbl(orsgetSOSetalis.Fields.Item("Quantity").Value)
                            oDeliveryOrder.Lines.BaseEntry = CInt(orsgetSOSetalis.Fields.Item("DocEntry").Value)
                            oDeliveryOrder.Lines.BaseLine = LineCount
                            oDeliveryOrder.Lines.BaseType = "17"

                            LineCount = LineCount + 1
                            orsgetSOSetalis.MoveNext()
                        Next

                        If oDeliveryOrder.Add = 0 Then
                            objMain.objApplication.StatusBar.SetText("Delivery Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            objMain.objDocumentType.UpdateDocument("ODLN", "Delivery")
                        Else
                            Dim err As String = objMain.objCompany.GetLastErrorDescription
                            objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)

                        End If

                    End If
                Catch ex As Exception

                End Try
                ' End If

                'ElseIf pVal.MenuUID = "Generate Invoice" And pVal.BeforeAction = False Then
                '    objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TRSHT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                '    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
                '    objMatrix3 = objForm.Items.Item("1000023").Specific

                '    Dim i As Integer
                '    For i = 1 To objMatrix3.VisualRowCount
                '        If objMatrix3.IsRowSelected(i) = True Then Exit For
                '    Next
                '    'If objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value = "Delivery" Then
                '    'Dim otherForm As SAPbouiCOM.Form
                '    If objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "A/R Invoice" Then

                '        Try
                '            Dim DeliveryEntry As String = objMatrix3.Columns.Item("V_5").Cells.Item(i - 1).Specific.Value

                '            Dim getSOSetalis As String = "Select * from ODLN T1 Inner Join DLN1 T2 on T1.""DocEntry""=T2.""DocEntry"" where T2.""DocEntry""='" & DeliveryEntry & "'"
                '            Dim orsgetSOSetalis As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '            orsgetSOSetalis.DoQuery(getSOSetalis)
                '            If orsgetSOSetalis.RecordCount > 0 Then
                '                Dim oInvoice As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                '                oInvoice.CardCode = orsgetSOSetalis.Fields.Item("CardCode").Value
                '                oInvoice.CardName = orsgetSOSetalis.Fields.Item("CardName").Value
                '                oInvoice.DocDate = orsgetSOSetalis.Fields.Item("DocDate").Value
                '                oInvoice.TaxDate = orsgetSOSetalis.Fields.Item("TaxDate").Value
                '                oInvoice.DocDueDate = orsgetSOSetalis.Fields.Item("DocDueDate").Value
                '                oInvoice.UserFields.Fields.Item("U_VSPDCTYP").Value = "Trip Sheet"
                '                oInvoice.UserFields.Fields.Item("U_VSPDCNO").Value = oDBs_Head.GetValue("DocEntry", 0) & "-" & i

                '                Dim LineCount As Integer = 0
                '                For j As Integer = 1 To orsgetSOSetalis.RecordCount
                '                    If LineCount > 0 Then
                '                        oInvoice.Lines.Add()
                '                    End If

                '                    'Dim ActualDelvQty As Double = objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value
                '                    'Dim Quantity As Double = orsgetSOSetalis.Fields.Item("Quantity").Value

                '                    'If CInt(Quantity) < CInt(ActualDelvQty) Then
                '                    '    objMain.objApplication.StatusBar.SetText("Actual Deleivered Quantity Should Not Be Gratter Than From Quantity In Row Level ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '                    '    Exit Try
                '                    'End If

                '                    oInvoice.Lines.ItemCode = orsgetSOSetalis.Fields.Item("ItemCode").Value
                '                    oInvoice.Lines.ItemDescription = orsgetSOSetalis.Fields.Item("Dscription").Value
                '                    oInvoice.Lines.Price = orsgetSOSetalis.Fields.Item("Price").Value
                '                    oInvoice.Lines.TaxCode = orsgetSOSetalis.Fields.Item("TaxCode").Value
                '                    oInvoice.Lines.WarehouseCode = orsgetSOSetalis.Fields.Item("WhsCode").Value


                '                    'Dim DifferenceQty As Double = 0.0
                '                    'Dim ShortageQty As Double = 0.0
                '                    'Dim TolerancePer As Double = 0.0
                '                    'Dim ActualToleranceQty As Double = 0.0

                '                    'Dim itemcode As String = orsgetSOSetalis.Fields.Item("ItemCode").Value

                '                    'Dim gettolerance As String = ""
                '                    'If objMain.IsSAPHANA = True Then

                '                    '    gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & itemcode & "' "
                '                    'Else
                '                    '    gettolerance = "Select ""U_VSPTLPRC"" from OITM where ""ItemCode""='" & itemcode & "' "
                '                    'End If
                '                    'Dim oRsgettolerance As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '                    'oRsgettolerance.DoQuery(gettolerance)

                '                    'If CStr(oRsgettolerance.Fields.Item("U_VSPTLPRC").Value) = 0 Then
                '                    '    objMain.objApplication.StatusBar.SetText("Please Enter Tolerance Percentage in Item Master Data for the Item Code : " + itemcode + " ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '                    '    Exit Try
                '                    'End If
                '                    'Dim tolerance As String = oRsgettolerance.Fields.Item("U_VSPTLPRC").Value


                '                    'TolerancePer = CDbl(oRsgettolerance.Fields.Item("U_VSPTLPRC").Value)
                '                    'DifferenceQty = Quantity - ActualDelvQty
                '                    'oInvoice.Lines.UserFields.Fields.Item("U_VSPADQTY").Value = CDbl(ActualDelvQty)
                '                    'oInvoice.Lines.UserFields.Fields.Item("U_VSPDFQTY").Value = CDbl(DifferenceQty)
                '                    'ActualToleranceQty = Quantity * TolerancePer / 100
                '                    'oInvoice.Lines.UserFields.Fields.Item("U_VSPTLQTY").Value = CDbl(ActualToleranceQty)

                '                    'If DifferenceQty > ActualToleranceQty Then
                '                    '    ShortageQty = DifferenceQty - ActualToleranceQty
                '                    '    oInvoice.Lines.UserFields.Fields.Item("U_VSPSTQTY").Value = CDbl(ShortageQty)
                '                    'Else
                '                    '    oInvoice.Lines.UserFields.Fields.Item("U_VSPSTQTY").Value = 0.0

                '                    'End If

                '                    'oInvoice.Lines.Quantity = ActualDelvQty

                '                    oInvoice.Lines.BaseEntry = CInt(orsgetSOSetalis.Fields.Item("DocEntry").Value)
                '                    oInvoice.Lines.BaseLine = LineCount
                '                    oInvoice.Lines.BaseType = "15"
                '                    LineCount = LineCount + 1
                '                    orsgetSOSetalis.MoveNext()
                '                Next
                '                If oInvoice.Add = 0 Then
                '                    'Dim GetDocEntry As String = ""
                '                    'If objMain.IsSAPHANA = True Then
                '                    '    GetDocEntry = "Select Max(""DocEntry"") From OINV"
                '                    'Else
                '                    '    GetDocEntry = "Select Max(DocEntry) From OINV"
                '                    'End If
                '                    'Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '                    'oRsGetDocEntry.DoQuery(GetDocEntry)
                '                    objMain.objApplication.StatusBar.SetText("A/R Invoice Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '                    objMain.objDocumentType.UpdateDocument("OINV", "A/R Invoice")
                '                Else
                '                    Dim err As String = objMain.objCompany.GetLastErrorDescription
                '                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '                End If
                '            End If
                '        Catch ex As Exception

                '        End Try
                '    End If

                ' End If

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TRSHT_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False Then
                        Me.SetCellsEditable(objForm.UniqueID)
                        Dim getname As String = ""
                        If objMain.IsSAPHANA = True Then
                            getname = "Select ""Code"" , ""U_VSPFNAME""  ||  ' '  || ""U_VSPLNAME"" as ""Names"" From ""@VSP_FLT_DRVRMSTR"" Where ""U_VSPCNCD"" = '" & objForm.Items.Item("22").Specific.Value & "'  order by ""Names"" asc"
                            objMain.objUtilities.MatrixComboBoxValues(objForm.Items.Item("69").Specific.Columns.Item("V_0"), getname)
                            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000036").Specific, "Select ""U_VSPFPLC"" , ""U_VSPFPLC"" From ""@VSP_FLT_RTMSTR"" R0 Inner Join " & _
                                            """@VSP_FLT_RTMSTR_C1"" R1 On R0.""Code"" = R1.""Code"" Where R0.""U_VSPRCD""  = '" & objForm.Items.Item("12").Specific.Value & "' Group By ""U_VSPFPLC""")
                        Else
                            getname = "Select Code , U_VSPFNAME + Space (1) + U_VSPLNAME as Names From [@VSP_FLT_DRVRMSTR] Where U_VSPCNCD = '" & objForm.Items.Item("22").Specific.Value & "'  order by Names asc"
                            objMain.objUtilities.MatrixComboBoxValues(objForm.Items.Item("69").Specific.Columns.Item("V_0"), getname)
                            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000036").Specific, "Select U_VSPFPLC , U_VSPFPLC From [@VSP_FLT_RTMSTR] R0 Inner Join " & _
                                            "[@VSP_FLT_RTMSTR_C1] R1 On R0.Code = R1.Code Where R0.U_VSPRCD  = '" & objForm.Items.Item("12").Specific.Value & "' Group By U_VSPFPLC")
                        End If


                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = True Then

                        If objForm.Items.Item("4").Specific.Value = String.Empty Then
                            Me.Validation(objForm.UniqueID)
                            BubbleEvent = False
                            Exit Try
                        End If

                        Dim GetCode As String = ""
                        If objMain.IsSAPHANA = True Then
                            GetCode = "Select ""Code"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("4").Specific.Value & "'"
                        Else
                            GetCode = "Select Code From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("4").Specific.Value & "'"
                        End If

                        Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetCode.DoQuery(GetCode)

                        objMain.sCmp = objMain.objCompany.GetCompanyService
                        objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
                        objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                        objMain.oGeneralParams.SetProperty("Code", oRsGetCode.Fields.Item(0).Value)
                        objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                        objMain.oGeneralData.SetProperty("U_VSPAVLB", "N")
                        ' objMain.oGeneralData.SetProperty("U_VSPCALB", "N")
                        objMain.oGeneralService.Update(objMain.oGeneralData)
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)

        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams

        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING

        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        objMatrix3 = objForm.Items.Item("1000023").Specific

        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If eventInfo.ItemUID = "1000023" And eventInfo.ColUID = "V_-1" And objMatrix3.RowCount > 0 Then

                            If objMatrix3.Columns.Item("V_1").Cells.Item(eventInfo.Row).Specific.Value = "Sales Order" Then
                                If objMatrix3.Columns.Item("V_5").Cells.Item(eventInfo.Row).Specific.Value = "" Then
                                    Try
                                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                        oMenus = oMenuItem.SubMenus
                                        If oMenus.Exists("Generate") = False Then
                                            oCreationPackage.UniqueID = "Generate"
                                            oCreationPackage.String = "Generate"
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
                                        If oMenus.Exists("Generate") = True Then
                                            objMain.objApplication.Menus.RemoveEx("Generate")
                                        End If
                                    Catch ex As Exception
                                        objMain.objApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                End If

                            Else
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                Try
                                    If oMenus.Exists("Generate") = True Then
                                        objMain.objApplication.Menus.RemoveEx("Generate")
                                    End If
                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            End If
                        End If

                        ''Generate Delivery Menu

                        If eventInfo.ItemUID = "1000023" And eventInfo.ColUID = "V_-1" And objMatrix3.RowCount > 0 Then

                            'If objMatrix3.Columns.Item("V_0").Cells.Item(eventInfo.Row).Specific.Value <> 0.0 Then
                            '    If objMatrix3.Columns.Item("V_5").Cells.Item(eventInfo.Row - 1).Specific.Value <> "" Then
                            objMatrix3 = objForm.Items.Item("1000023").Specific

                            Dim i As Integer
                            For i = 1 To objMatrix3.VisualRowCount
                                If objMatrix3.IsRowSelected(i) = True Then Exit For
                            Next
                            If objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Value = "Sales Order" Then
                                If objMatrix3.Columns.Item("V_5").Cells.Item(i).Specific.Value <> "" Then

                                    Try

                                        Dim CheckSaleStatus As String = ""

                                        If objMain.IsSAPHANA = True Then
                                            CheckSaleStatus = "Select ""DocStatus"",""DocNum"" From ORDR Where ""DocNum"" = (Select ""DocNum"" From ORDR where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(i).Specific.Value & "') "
                                        Else
                                            CheckSaleStatus = "Select ""DocStatus"",""DocNum"" From ORDR Where ""DocNum"" = (Select ""DocNum"" From ORDR where ""DocEntry""='" & objMatrix3.Columns.Item("V_5").Cells.Item(i).Specific.Value & "') "
                                        End If
                                        Dim oRsCheckSaleStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRsCheckSaleStatus.DoQuery(CheckSaleStatus)


                                        If oRsCheckSaleStatus.RecordCount > 0 Then
                                            Dim status As String = ""
                                            status = oRsCheckSaleStatus.Fields.Item(0).Value.ToString()
                                            Dim docnum As String = oRsCheckSaleStatus.Fields.Item(1).Value.ToString()
                                            If status = "O" Then
                                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                                oMenus = oMenuItem.SubMenus
                                                If oMenus.Exists("Generate Delivery") = False Then
                                                    oCreationPackage.UniqueID = "Generate Delivery"
                                                    oCreationPackage.String = "Generate Delivery"
                                                    oCreationPackage.Enabled = True
                                                    oMenus.AddEx(oCreationPackage)
                                                End If

                                            Else

                                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                                oMenus = oMenuItem.SubMenus
                                                Try
                                                    If oMenus.Exists("Generate Delivery") = True Then
                                                        objMain.objApplication.Menus.RemoveEx("Generate Delivery")
                                                    End If
                                                Catch ex As Exception
                                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                                End Try


                                            End If
                                        Else
                                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                            oMenus = oMenuItem.SubMenus
                                            Try
                                                If oMenus.Exists("Generate Delivery") = True Then
                                                    objMain.objApplication.Menus.RemoveEx("Generate Delivery")
                                                End If
                                            Catch ex As Exception
                                                objMain.objApplication.StatusBar.SetText(ex.Message)
                                            End Try
                                        End If
                                    Catch ex As Exception
                                        objMain.objApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                Else
                                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                    oMenus = oMenuItem.SubMenus
                                    Try
                                        If oMenus.Exists("Generate Delivery") = True Then
                                            objMain.objApplication.Menus.RemoveEx("Generate Delivery")
                                        End If
                                    Catch ex As Exception
                                        objMain.objApplication.StatusBar.SetText(ex.Message)
                                    End Try
                                End If


                            Else
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                Try
                                    If oMenus.Exists("Generate Delivery") = True Then
                                        objMain.objApplication.Menus.RemoveEx("Generate Delivery")
                                    End If
                                Catch ex As Exception
                                    objMain.objApplication.StatusBar.SetText(ex.Message)
                                End Try
                            End If

                        End If

                        ''Generate A/R Invoice

                        'If eventInfo.ItemUID = "1000023" And eventInfo.ColUID = "V_-1" And objMatrix3.RowCount > 0 Then
                        '    If objMatrix3.Columns.Item("V_5").Cells.Item(eventInfo.Row).Specific.Value = "" Then
                        '        If objMatrix3.Columns.Item("V_1").Cells.Item(eventInfo.Row).Specific.Value = "A/R Invoice" Then
                        '            If objMatrix3.Columns.Item("V_0").Cells.Item(eventInfo.Row).Specific.Value <> 0.0 Then
                        '                If objMatrix3.Columns.Item("V_5").Cells.Item(eventInfo.Row - 1).Specific.Value <> "" Then
                        '                    Try
                        '                        Dim count As Integer = 0
                        '                        'For i As Integer = 1 To objMatrix3.VisualRowCount
                        '                        '    If objMatrix3.Columns.Item("V_1").Cells.Item(i).Specific.Selected.Value = "A/R Invoice" Then
                        '                        '        count = 1
                        '                        '    End If
                        '                        'Next
                        '                        If count = 0 Then
                        '                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        '                            oMenus = oMenuItem.SubMenus
                        '                            If oMenus.Exists("Generate Invoice") = False Then
                        '                                oCreationPackage.UniqueID = "Generate Invoice"
                        '                                oCreationPackage.String = "Generate Invoice"
                        '                                oCreationPackage.Enabled = True
                        '                                oMenus.AddEx(oCreationPackage)

                        '                            End If
                        '                        End If

                        '                    Catch ex As Exception
                        '                        objMain.objApplication.StatusBar.SetText(ex.Message)
                        '                    End Try
                        '                Else
                        '                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        '                    oMenus = oMenuItem.SubMenus
                        '                    Try
                        '                        If oMenus.Exists("Generate Invoice") = True Then
                        '                            objMain.objApplication.Menus.RemoveEx("Generate Invoice")
                        '                        End If
                        '                    Catch ex As Exception
                        '                        objMain.objApplication.StatusBar.SetText(ex.Message)
                        '                    End Try
                        '                End If
                        '            End If
                        '        End If
                        '    End If
                        'End If
                    End If
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region "Methods"

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Sub CFLAccounts(ByVal FormUID As String, ByVal CFL_ID As String, ByVal GroupMask As String, ByVal GroupNumber As String, ByVal CashAccount As String)
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
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)

            If GroupMask = "Y" Then
                If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCondition = oConditions.Add()
                oCondition.Alias = "GroupMask"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = GroupNumber
                oChooseFromList.SetConditions(oConditions)
            End If

            If CashAccount = "Y" Then
                If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                oCondition = oConditions.Add()
                oCondition.Alias = "Finanse"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = "Y"
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLExpensesAccounts(ByVal FormUID As String, ByVal CFL_ID As String, ByVal GroupMask As String, ByVal GroupNumber As String, ByVal GroupNumber1 As String, ByVal CashAccount As String)
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
            oCondition.Alias = "Postable"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)

            If GroupMask = "Y" Then
                If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCondition = oConditions.Add()
                oCondition.Alias = "GroupMask"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = GroupNumber
                oChooseFromList.SetConditions(oConditions)

                If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                oCondition = oConditions.Add()
                oCondition.Alias = "GroupMask"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = GroupNumber1
                oChooseFromList.SetConditions(oConditions)
            End If

            If CashAccount = "Y" Then
                If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

                oCondition = oConditions.Add()
                oCondition.Alias = "Finanse"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = "Y"
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLVendorFilter(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "CardType"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "S"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterForChemicalItem(ByVal FormUID As String, ByVal CFL_ID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetTankItems As String = ""

            If objMain.IsSAPHANA = True Then
                GetTankItems = "Select ""U_VSPITMCD"" From ""@VSP_FLT_TANKMSTR_C0"" A Inner Join ""@VSP_FLT_TANKMSTR"" B On A.""Code"" = B.""Code"" Where " & _
                                              "B.""U_VSPTNKNO"" = '" & objForm.Items.Item("71").Specific.Value & "' And ""U_VSPITMCD"" <> '' "
            Else
                GetTankItems = "Select U_VSPITMCD From [@VSP_FLT_TANKMSTR_C0] A Inner Join [@VSP_FLT_TANKMSTR] B On A.Code = B.Code Where " & _
                                                              "B.U_VSPTNKNO = '" & objForm.Items.Item("71").Specific.Value & "' And U_VSPITMCD <> '' "
            End If
            Dim oRsGetTankItems As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTankItems.DoQuery(GetTankItems)

            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetTankItems.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetTankItems.Fields.Item("U_VSPITMCD").Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetTankItems.MoveNext()
                For i As Integer = 1 To oRsGetTankItems.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "ItemCode"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetTankItems.Fields.Item("U_VSPITMCD").Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetTankItems.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "ItemCode"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
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
            oCondition.Alias = "U_VSPFLSTS"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Open to Fleet"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterCostCentres(ByVal FormUID As String, ByVal CFL_ID As String, ByVal DimCode As String)
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
            oCondition.Alias = "DimCode"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = DimCode
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterCostCentres1(ByVal FormUID As String, ByVal CFL_ID As String, ByVal DimCode As String)
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
            oCondition.Alias = "DimCode"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = DimCode
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterForVehicles(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "U_VSPAVLB"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)

            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPCHK"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)

            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPCALB"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)

            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPUNPCK"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "N"
            oChooseFromList.SetConditions(oConditions)




        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String, Optional ByVal OpenKM As String = "0", Optional ByVal Source As String = "")
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
            oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
            oDBs_Detail3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3")
            oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")
            oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")
            oDBs_Detail6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C6")
            oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")
            oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")
            oDBs_Detail9 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9")

            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix3 = objForm.Items.Item("1000023").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix5 = objForm.Items.Item("67").Specific
            objMatrix6 = objForm.Items.Item("1000028").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific
            objMatrix9 = objForm.Items.Item("1000030").Specific

            Select Case MatrixUID
                Case "61"
                    objMatrix1.AddRow()
                    oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, OpenKM)
                    oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, Source)
                    oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, "")
                    oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                Case "1000022"
                    objMatrix2.AddRow()
                    oDBs_Detail2.SetValue("LineId", oDBs_Detail2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Detail2.SetValue("U_VSPJOUDT", oDBs_Detail2.Offset, DateTime.Now.ToString("yyyyMMdd"))
                    oDBs_Detail2.SetValue("U_VSPAMGOO", oDBs_Detail2.Offset, 0)
                    oDBs_Detail2.SetValue("U_VSPADAMT", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VSPFRACT", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VAPTOACT", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VSPDRVCC", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VAPCAS", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VSPCOM", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VSPJENO", oDBs_Detail2.Offset, "")
                    oDBs_Detail2.SetValue("U_VSPOPNO", oDBs_Detail2.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                Case "1000023"
                    objMatrix3.AddRow()
                    oDBs_Detail3.SetValue("LineId", oDBs_Detail3.Offset, objMatrix3.VisualRowCount)
                    oDBs_Detail3.SetValue("U_VSPTYPE", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPDOCTY", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPGENTY", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPDATE", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPDCNUM", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPBPCOD", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPQUANT", oDBs_Detail3.Offset, 0.0)
                    oDBs_Detail3.SetValue("U_VSPREF", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPDCTOT", oDBs_Detail3.Offset, "")
                    oDBs_Detail3.SetValue("U_VSPREM", oDBs_Detail3.Offset, "")
                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                Case "1000024"
                    objMatrix4.AddRow()
                    oDBs_Detail4.SetValue("LineId", oDBs_Detail4.Offset, objMatrix4.VisualRowCount)
                    oDBs_Detail4.SetValue("U_VSPTYPE", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPFRRMA", oDBs_Detail4.Offset, "N")
                    oDBs_Detail4.SetValue("U_VSPEXACC", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPEXACN", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPADACC", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPAMT", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPBUD", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPMTYP", oDBs_Detail4.Offset, "")
                    objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                Case "67"
                    objMatrix5.AddRow()
                    oDBs_Detail5.SetValue("LineId", oDBs_Detail5.Offset, objMatrix5.VisualRowCount)
                    oDBs_Detail5.SetValue("U_VSPDATE", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPVENCO", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPVENNM", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPQUAN", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPRATE", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPAMT", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPDRCC1", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPDCNUM", oDBs_Detail5.Offset, "")
                    oDBs_Detail5.SetValue("U_VSPGISU", oDBs_Detail5.Offset, "")
                    objMatrix5.SetLineData(objMatrix5.VisualRowCount)

                Case "1000028"
                    objMatrix6.AddRow()
                    oDBs_Detail6.SetValue("LineId", oDBs_Detail6.Offset, objMatrix6.VisualRowCount)
                    oDBs_Detail6.SetValue("U_VSPDAT", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPSOU", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPSOUR", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPFRDT", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPFRTM", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPTODT", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPTOTM", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPCHCOD", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPCHNAM", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPWEIGH", oDBs_Detail6.Offset, "")
                    oDBs_Detail6.SetValue("U_VSPUOM", oDBs_Detail6.Offset, "")
                    objMatrix6.SetLineData(objMatrix6.VisualRowCount)

                Case "69"
                    objMatrix7.AddRow()
                    oDBs_Detail7.SetValue("LineId", oDBs_Detail7.Offset, objMatrix7.VisualRowCount)
                    oDBs_Detail7.SetValue("U_VSPDRCOD", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPDRFNM", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPDRMNM", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPDRLNM", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPMBNUM", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPLNNUM", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPEXPDT", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPFRMDT", oDBs_Detail7.Offset, "")
                    oDBs_Detail7.SetValue("U_VSPTODT", oDBs_Detail7.Offset, "")
                    objMatrix7.SetLineData(objMatrix7.VisualRowCount)

                Case "1000030"
                    objMatrix9.AddRow()
                    oDBs_Detail9.SetValue("LineId", oDBs_Detail9.Offset, objMatrix9.VisualRowCount)
                    oDBs_Detail9.SetValue("U_VSPATNM", oDBs_Detail9.Offset, "")
                    oDBs_Detail9.SetValue("U_VSPAPTH", oDBs_Detail9.Offset, "")
                    objMatrix9.SetLineData(objMatrix9.VisualRowCount)
            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("79").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("79").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("80").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("80").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("54").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("54").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("56").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("56").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("71").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("71").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("49").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("49").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("73").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("73").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            'objForm.Items.Item("1000020").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'objForm.Items.Item("1000020").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000001").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000001").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("16").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000002").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000002").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("18").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000027").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000027").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000010").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000010").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000012").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000012").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000018").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000018").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            objForm.Items.Item("1000037").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000037").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


            'Buttons
            objForm.Items.Item("1000031").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000031").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000032").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000032").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000035").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000035").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("89").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("89").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadVehicleDetails(ByVal FormUID As String, ByVal VehicleNum As String, ByVal DieselItem As String, ByVal ConCode As String, _
                           ByVal ConName As String, ByVal ODOReading As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
            oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
            oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")
            oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")
            oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")

            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific

            Dim CheckIfTyreDetailsUpdated As String = ""
            If objMain.IsSAPHANA = True Then
                CheckIfTyreDetailsUpdated = "Select ""DocNum"" From ""@VSP_FLT_TRSHT"" Where ""U_VSPVHCL"" = '" & VehicleNum & "' And IFNULL(""U_VSPTYUP"",'N') = 'N'"
            Else
                CheckIfTyreDetailsUpdated = "Select DocNum From [@VSP_FLT_TRSHT] Where U_VSPVHCL = '" & VehicleNum & "' And ISNULL(U_VSPTYUP,'N') = 'N'"
            End If


            Dim oRsCheckIfTyreDetailsUpdated As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCheckIfTyreDetailsUpdated.DoQuery(CheckIfTyreDetailsUpdated)

            If oRsCheckIfTyreDetailsUpdated.RecordCount > 0 Then
                objMain.objApplication.StatusBar.SetText("Tyre Details Not Updated For Trip Sheet No. " & oRsCheckIfTyreDetailsUpdated.Fields.Item(0).Value, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            objForm.Freeze(True)

            objMatrix8.Clear()
            oDBs_Detail8.Clear()
            objMatrix8.FlushToDataSource()

            objMatrix2.Clear()
            oDBs_Detail2.Clear()
            objMatrix2.FlushToDataSource()

            objMatrix4.Clear()
            oDBs_Detail4.Clear()
            objMatrix4.FlushToDataSource()

            objMatrix7.Clear()
            oDBs_Detail7.Clear()
            objMatrix7.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID, "69")

            oDBs_Head.SetValue("U_VSPVHCL", oDBs_Head.Offset, VehicleNum)
            oDBs_Head.SetValue("U_VSPDISIT", oDBs_Head.Offset, DieselItem)
            oDBs_Head.SetValue("U_VSPCNTR", oDBs_Head.Offset, ConCode)
            oDBs_Head.SetValue("U_VSPCNTNM", oDBs_Head.Offset, ConName)

            'Clearing Route Variables
            oDBs_Head.SetValue("U_VSPROUTE", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_VSPSOURC", oDBs_Head.Offset, "")
            oDBs_Head.SetValue("U_VSPDEST", oDBs_Head.Offset, "")

            Dim str As String = ""
            If objMain.IsSAPHANA = True Then
                str = "Select ""Code"" , ""U_VSPFNAME""  ||  ' '  ||   ""U_VSPLNAME"" From ""@VSP_FLT_DRVRMSTR"" Where ""U_VSPCNCD"" = '" & ConCode & "'"
            Else
                str = "Select Code , U_VSPFNAME + Space (1) + U_VSPLNAME From [@VSP_FLT_DRVRMSTR] Where U_VSPCNCD = '" & ConCode & "'"
            End If

            objMain.objUtilities.MatrixComboBoxValues(objForm.Items.Item("69").Specific.Columns.Item("V_0"), str)
            'Get Tank No.
            Dim GetTankNo As String = ""

            If objMain.IsSAPHANA = True Then
                GetTankNo = "Select ""U_VSPTNUM"" From ""@VSP_FLT_TANKMPG"" T Inner Join ""@VSP_FLT_TANKMPG_C0"" T1 On T.""DocEntry"" = T1.""DocEntry"" " & _
                                        "Where ""U_VSPVCHN0"" = '" & VehicleNum & "' And ""U_VSPSTS"" = 'Attached'"
            Else
                GetTankNo = "Select U_VSPTNUM From [@VSP_FLT_TANKMPG] T Inner Join [@VSP_FLT_TANKMPG_C0] T1 On T.DocEntry = T1.DocEntry " & _
                                        "Where U_VSPVCHN0 = '" & VehicleNum & "' And U_VSPSTS = 'Attached'"
            End If

            Dim oRsGetTankNo As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTankNo.DoQuery(GetTankNo)
            oDBs_Head.SetValue("U_VSPTANK", oDBs_Head.Offset, oRsGetTankNo.Fields.Item(0).Value)

            'Get Last Calibrated Date
            Dim GetLastCalibratedDate As String = ""

            If objMain.IsSAPHANA = True Then
                GetLastCalibratedDate = "Select ""U_VSPCCDT"" From ""@VSP_FLT_CALBRTN"" Where ""U_VSPVCHID"" = '" & VehicleNum & "' And " & _
            """DocNum"" = (Select Max(""DocNum"") From ""@VSP_FLT_CALBRTN"") And ""U_VSPAPBY"" <> ''"
            Else
                GetLastCalibratedDate = "Select Convert(Varchar,U_VSPCCDT,112) From [@VSP_FLT_CALBRTN] Where U_VSPVCHID = '" & VehicleNum & "' And " & _
            "DocNum = (Select Max(DocNum) From [@VSP_FLT_CALBRTN]) And U_VSPAPBY <> ''"
            End If
            Dim oRsGetLastCalibratedDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetLastCalibratedDate.DoQuery(GetLastCalibratedDate)
            If oRsGetLastCalibratedDate.RecordCount > 0 Then

                oDBs_Head.SetValue("U_VSPLCDT", oDBs_Head.Offset, CDate(oRsGetLastCalibratedDate.Fields.Item(0).Value).ToString("yyyyMMdd"))
            End If
            Dim GetNoofTyres As String = "Select Cast(""U_VSPTYRES"" As Integer)  As ""U_VSPTYRES""  From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO""= '" & VehicleNum & "' "
            Dim oRsGetNoofTyres As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetNoofTyres.DoQuery(GetNoofTyres)

            Dim GetNoofTyresfromTYmp As String = ""
            If objMain.IsSAPHANA = True Then
                GetNoofTyresfromTYmp = "Select Count(T1.""LineId"")  from ""@VSP_FLT_TYRMPG"" T0 inner join ""@VSP_FLT_TYRMPG_C0"" T1 on T1.""DocEntry"" = T0.""DocEntry"" " & _
                                                 "where T0.""U_VSPVCHN0"" = '" & VehicleNum & "' and T1.""U_VSPSTS"" = 'Attached' "
            Else
                GetNoofTyresfromTYmp = "Select Count(T1.LineId)  from [@VSP_FLT_TYRMPG] T0 inner join [@VSP_FLT_TYRMPG_C0] T1 on T1.DocEntry = T0.DocEntry " & _
                                                 "where T0.U_VSPVCHN0 = '" & VehicleNum & "' and T1.U_VSPSTS = 'Attached' "
            End If
            Dim oRsGetNoofTyresfromTYmp As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetNoofTyresfromTYmp.DoQuery(GetNoofTyresfromTYmp)

            Dim tyrno As Integer = oRsGetNoofTyres.Fields.Item(0).Value
            Dim tyrno1 As Integer = oRsGetNoofTyresfromTYmp.Fields.Item(0).Value

            If oRsGetNoofTyres.Fields.Item(0).Value = oRsGetNoofTyresfromTYmp.Fields.Item(0).Value Then
                Me.LoadTyreDetails(objForm.UniqueID)
            End If

            'Get Tyre Details
            'Dim GetTyreDetails As String = "Select T3.U_VSPTRNUM , T3.U_VSPTRNM , T3.U_VSPPSTN , U_VSPTRMDL , U_VSPKMRUN From [@VSP_FLT_TYRMPG] T0 Inner Join " & _
            '"[@VSP_FLT_TYRMPG_C0] T1 On T0.DocEntry = T1.DocEntry Inner Join [@VSP_FLT_TYRMSTR] T3 On T3.U_VSPTRNUM = T1.U_VSPTRNUM Where " & _
            '"U_VSPVCHN0 = '" & VehicleNum & "' And (U_VSPSTS = 'Attached' Or U_VSPSTS = 'Stepney')"
            'Dim oRsGetTyreDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRsGetTyreDetails.DoQuery(GetTyreDetails)

            'If oRsGetTyreDetails.RecordCount > 0 Then
            '    oRsGetTyreDetails.MoveFirst()

            '    For i As Integer = 1 To oRsGetTyreDetails.RecordCount

            '        objMatrix8.AddRow()
            '        oDBs_Detail8.SetValue("LineId", oDBs_Detail8.Offset, objMatrix8.VisualRowCount)
            '        oDBs_Detail8.SetValue("U_VSPTYRNO", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPTRNUM").Value)
            '        oDBs_Detail8.SetValue("U_VSPTYRNM", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPTRNM").Value)
            '        oDBs_Detail8.SetValue("U_VSPTYMOD", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPTRMDL").Value)
            '        oDBs_Detail8.SetValue("U_VSPTYPOS", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPPSTN").Value)
            '        oDBs_Detail8.SetValue("U_VSPKMS", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPKMRUN").Value)
            '        objMatrix8.SetLineData(objMatrix8.VisualRowCount)

            '        oRsGetTyreDetails.MoveNext()
            '    Next

            'Else
            '    objMain.objApplication.StatusBar.SetText("Please Update Tyre Details In Tyre Mapping Screen", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            'End If

            'Set ODOMeter Reading
            oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, objMatrix1.VisualRowCount)
            oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, ODOReading)
            oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, "")
            objMatrix1.SetLineData(objMatrix1.VisualRowCount)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadTyreDetails(ByVal FormUID As String)

        Try
            Dim GetTyreDetails As String = ""

            If objMain.IsSAPHANA = True Then
                GetTyreDetails = "Select T3.""U_VSPTRNUM"" , T3.""U_VSPTRNM"" , T3.""U_VSPPSTN"" , ""U_VSPTRMDL"" , ""U_VSPKMRUN"", ""U_VSPSTS"" From ""@VSP_FLT_TYRMPG"" T0 Inner Join " & _
           """@VSP_FLT_TYRMPG_C0"" T1 On T0.""DocEntry"" = T1.""DocEntry"" Inner Join ""@VSP_FLT_TYRMSTR"" T3 On T3.""U_VSPTRNUM"" = T1.""U_VSPTRNUM"" Where " & _
           """U_VSPVCHN0"" = '" & objForm.Items.Item("4").Specific.Value & "' And (""U_VSPSTS"" = 'Attached' Or ""U_VSPSTS"" = 'Stepney')"


            Else
                GetTyreDetails = "Select T3.U_VSPTRNUM , T3.U_VSPTRNM , T3.U_VSPPSTN , U_VSPTRMDL , U_VSPKMRUN, U_VSPSTS From [@VSP_FLT_TYRMPG] T0 Inner Join " & _
           "[@VSP_FLT_TYRMPG_C0] T1 On T0.DocEntry = T1.DocEntry Inner Join [@VSP_FLT_TYRMSTR] T3 On T3.U_VSPTRNUM = T1.U_VSPTRNUM Where " & _
           "U_VSPVCHN0 = '" & objForm.Items.Item("4").Specific.Value & "' And (U_VSPSTS = 'Attached' Or U_VSPSTS = 'Stepney')"


            End If
            Dim oRsGetTyreDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTyreDetails.DoQuery(GetTyreDetails)

            objMatrix8.Clear()

            If oRsGetTyreDetails.RecordCount > 0 Then
                oRsGetTyreDetails.MoveFirst()

                For i As Integer = 1 To oRsGetTyreDetails.RecordCount

                    objMatrix8.AddRow()
                    oDBs_Detail8.SetValue("LineId", oDBs_Detail8.Offset, objMatrix8.VisualRowCount)
                    oDBs_Detail8.SetValue("U_VSPTYRNO", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPTRNUM").Value)
                    oDBs_Detail8.SetValue("U_VSPTYRNM", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPTRNM").Value)
                    oDBs_Detail8.SetValue("U_VSPTYMOD", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPTRMDL").Value)
                    oDBs_Detail8.SetValue("U_VSPTYPOS", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPPSTN").Value)
                    oDBs_Detail8.SetValue("U_VSPKMS", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPKMRUN").Value)
                    oDBs_Detail8.SetValue("U_VSPSTS", oDBs_Detail8.Offset, oRsGetTyreDetails.Fields.Item("U_VSPSTS").Value)

                    objMatrix8.SetLineData(objMatrix8.VisualRowCount)

                    oRsGetTyreDetails.MoveNext()
                Next
            Else
                objMain.objApplication.StatusBar.SetText("Please Update Tyre Details In Tyre Mapping Screen", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub LoadRouteDetails(ByVal FormUID As String, ByVal Route As String, ByVal Source As String, ByVal Desintation As String, _
                                                                                                            ByVal AdvanceAmount As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
            oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
            oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")

            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            Dim str As String = ""
            If objMain.IsSAPHANA = True Then
                str = "Select ""U_VSPFPLC"" , ""U_VSPFPLC"" From ""@VSP_FLT_RTMSTR"" R0 Inner Join " & _
                                        """@VSP_FLT_RTMSTR_C1"" R1 On R0.""Code"" = R1.""Code"" Where R0.""U_VSPRCD""  = '" & Route & "' Group By ""U_VSPFPLC"""
            Else
                str = "Select U_VSPFPLC , U_VSPFPLC From [@VSP_FLT_RTMSTR] R0 Inner Join " & _
                                        "[@VSP_FLT_RTMSTR_C1] R1 On R0.Code = R1.Code Where R0.U_VSPRCD  = '" & Route & "' Group By U_VSPFPLC"
            End If

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000036").Specific, str)

            oDBs_Head.SetValue("U_VSPROUTE", oDBs_Head.Offset, Route)
            oDBs_Head.SetValue("U_VSPSOURC", oDBs_Head.Offset, Source)
            oDBs_Head.SetValue("U_VSPDEST", oDBs_Head.Offset, Desintation)

            oDBs_Detail1.SetValue("LineId", oDBs_Detail1.Offset, 1)
            oDBs_Detail1.SetValue("U_VSPOPKM", oDBs_Detail1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(1).Specific.Value)
            oDBs_Detail1.SetValue("U_VSPCLKM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPSOUR", oDBs_Detail1.Offset, Source)
            oDBs_Detail1.SetValue("U_VSPDEST", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPFRDT", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPFRTM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPTODT", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPTOTM", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPLOAD", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_VSPDICON", oDBs_Detail1.Offset, "")
            oDBs_Detail1.SetValue("U_TOTKM", oDBs_Detail1.Offset, "")
            objMatrix1.SetLineData(1)
            objMatrix1.AutoResizeColumns()

            objMatrix2.Clear()
            oDBs_Detail2.Clear()
            objMatrix2.FlushToDataSource()

            objMatrix2.AddRow()
            oDBs_Detail2.SetValue("LineId", oDBs_Detail2.Offset, objMatrix2.VisualRowCount)
            oDBs_Detail2.SetValue("U_VSPJOUDT", oDBs_Detail2.Offset, DateTime.Now.ToString("yyyyMMdd"))
            oDBs_Detail2.SetValue("U_VSPAMGOO", oDBs_Detail2.Offset, 0)
            oDBs_Detail2.SetValue("U_VSPADAMT", oDBs_Detail2.Offset, AdvanceAmount)
            oDBs_Detail2.SetValue("U_VSPFRACT", oDBs_Detail2.Offset, "")
            oDBs_Detail2.SetValue("U_VAPTOACT", oDBs_Detail2.Offset, "")
            oDBs_Detail2.SetValue("U_VSPDRVCC", oDBs_Detail2.Offset, "")
            oDBs_Detail2.SetValue("U_VAPCAS", oDBs_Detail2.Offset, "")
            oDBs_Detail2.SetValue("U_VSPCOM", oDBs_Detail2.Offset, "")
            oDBs_Detail2.SetValue("U_VSPJENO", oDBs_Detail2.Offset, "")
            oDBs_Detail2.SetValue("U_VSPOPNO", oDBs_Detail2.Offset, "")
            objMatrix2.SetLineData(objMatrix2.VisualRowCount)

            Dim GetExpenseDetails As String = ""

            If objMain.IsSAPHANA = True Then
                GetExpenseDetails = "Select ""U_VSPEACD"" as ""AccountCode"" , ""U_VSPEANM"" as ""AccountName"" , ""U_VSPAMT"" as ""Amount"" From ""@VSP_FLT_RTMSTR"" R0 " & _
           "Inner Join ""@VSP_FLT_RTMSTR_C0"" R1 On R0.""Code"" = R1.""Code"" Where R0.""U_VSPRCD"" = '" & Route & "' And ""U_VSPEACD"" <> ''"

            Else
                GetExpenseDetails = "Select U_VSPEACD as 'AccountCode' , U_VSPEANM as 'AccountName' , U_VSPAMT as 'Amount' From [@VSP_FLT_RTMSTR] R0 " & _
           "Inner Join [@VSP_FLT_RTMSTR_C0] R1 On R0.Code = R1.Code Where R0.U_VSPRCD = '" & Route & "' And U_VSPEACD <> ''"

            End If
            Dim oRsGetExpenseDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetExpenseDetails.DoQuery(GetExpenseDetails)

            If oRsGetExpenseDetails.RecordCount > 0 Then
                oRsGetExpenseDetails.MoveFirst()
                For i As Integer = 1 To oRsGetExpenseDetails.RecordCount
                    objMatrix4.AddRow()
                    oDBs_Detail4.SetValue("LineId", oDBs_Detail4.Offset, objMatrix4.VisualRowCount)
                    oDBs_Detail4.SetValue("U_VSPTYPE", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPFRRMA", oDBs_Detail4.Offset, "Y")
                    oDBs_Detail4.SetValue("U_VSPEXACC", oDBs_Detail4.Offset, oRsGetExpenseDetails.Fields.Item(0).Value)
                    oDBs_Detail4.SetValue("U_VSPEXACN", oDBs_Detail4.Offset, oRsGetExpenseDetails.Fields.Item(1).Value)
                    oDBs_Detail4.SetValue("U_VSPADACC", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPAMT", oDBs_Detail4.Offset, "")
                    oDBs_Detail4.SetValue("U_VSPBUD", oDBs_Detail4.Offset, oRsGetExpenseDetails.Fields.Item(2).Value)
                    oDBs_Detail4.SetValue("U_VSPMTYP", oDBs_Detail4.Offset, "")

                    If i = oRsGetExpenseDetails.RecordCount Then
                        oDBs_Detail4.SetValue("U_VSPFRRMA", oDBs_Detail4.Offset, "N")
                    End If
                    objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                    oRsGetExpenseDetails.MoveNext()
                Next
            End If

            Me.SetNewLine(objForm.UniqueID, "1000024")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PostJE(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")

            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific

            Dim i As Integer = 0
            Dim LineSelect As Boolean = False
            For i = 1 To objMatrix2.VisualRowCount
                If objMatrix2.IsRowSelected(i) = True Then
                    LineSelect = True
                    Exit For
                End If
            Next
            If LineSelect = False Then Exit Try

            If objMatrix2.Columns.Item("V_7").Cells.Item(i).Specific.Value <> "" Then
                objMain.objApplication.StatusBar.SetText("Journey Entry Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix2.Columns.Item("V_0").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Journey Date Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix2.Columns.Item("V_2").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Advance Amount Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix2.Columns.Item("V_2").Cells.Item(i).Specific.Value = 0 Then
                objMain.objApplication.StatusBar.SetText("Advance Amount Should be Greater than Zero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix2.Columns.Item("V_3").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("From Account Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix2.Columns.Item("V_4").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("To Account Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim JourneyDate As String = objMatrix2.Columns.Item("V_0").Cells.Item(i).Specific.Value
            JourneyDate = JourneyDate.Insert("4", "-")
            JourneyDate = JourneyDate.Insert("7", "-")

            oJE.ReferenceDate = JourneyDate
            oJE.TaxDate = JourneyDate
            oJE.DueDate = JourneyDate

            oJE.UserFields.Fields.Item("U_VSPTRPNO").Value = objForm.Items.Item("16").Specific.Value
            oJE.UserFields.Fields.Item("U_VSPDCTYP").Value = "Trip Sheet"

            oJE.Lines.AccountCode = objMatrix2.Columns.Item("V_4").Cells.Item(i).Specific.Value
            oJE.Lines.Debit = objMatrix2.Columns.Item("V_2").Cells.Item(i).Specific.Value
            oJE.Lines.CostingCode3 = objMatrix2.Columns.Item("V_8").Cells.Item(1).Specific.Value
            oJE.Lines.Add()

            oJE.Lines.AccountCode = objMatrix2.Columns.Item("V_3").Cells.Item(i).Specific.Value
            oJE.Lines.Credit = objMatrix2.Columns.Item("V_2").Cells.Item(i).Specific.Value
            oJE.Lines.CostingCode3 = objMatrix2.Columns.Item("V_8").Cells.Item(1).Specific.Value

            If oJE.Add = 0 Then
                Dim GetTransID As String = ""
                If objMain.IsSAPHANA = True Then
                    GetTransID = "Select Max(""TransId"") From OJDT"
                Else
                    GetTransID = "Select Max(TransId) From OJDT"
                End If
                Dim oRsGetTransID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetTransID.DoQuery(GetTransID)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C2")
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPJENO", oRsGetTransID.Fields.Item(0).Value.ToString)

                If i = 1 Then
                    objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C4")
                    For j As Integer = 1 To objMatrix4.VisualRowCount
                        objMain.oChildren.Item(j - 1).SetProperty("U_VSPADACC", objMatrix2.Columns.Item("V_4").Cells.Item(i).Specific.Value)
                    Next
                End If
                objMain.oGeneralService.Update(objMain.oGeneralData)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to post JE, Error : " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            Me.RefreshData(objForm.UniqueID)
            Me.SetCellsEditable(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub PostExpense(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")

            objMatrix4 = objForm.Items.Item("1000024").Specific

            If objForm.Items.Item("1000001").Specific.Value <> "" Then
                objMain.objApplication.StatusBar.SetText("Expense Already Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim CheckIfAdvancesNotPosted As String = ""

            If objMain.IsSAPHANA = True Then
                CheckIfAdvancesNotPosted = "Select * From ""@VSP_FLT_TRSHT_C2"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' " & _
           "And (""U_VSPJENO"" IS NULL Or ""U_VSPJENO"" = '')"
            Else
                CheckIfAdvancesNotPosted = "Select * From [@VSP_FLT_TRSHT_C2] Where DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "' " & _
           "And (U_VSPJENO IS NULL Or U_VSPJENO = '')"
            End If
            Dim oRsCheckIfAdvancesNotPosted As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCheckIfAdvancesNotPosted.DoQuery(CheckIfAdvancesNotPosted)

            'If oRsCheckIfAdvancesNotPosted.RecordCount > 0 Then
            '    objMain.objApplication.StatusBar.SetText("Please Post Advances", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Try
            'End If

            Dim ExpenseAmount As Double = 0
            For i As Integer = 1 To objMatrix4.VisualRowCount
                If objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value <> "" Then
                    ExpenseAmount = ExpenseAmount + objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value
                End If
            Next
            If ExpenseAmount = 0 Then
                objMain.objApplication.StatusBar.SetText("Total Amount is Zero, Expense Cannot Be Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            If objForm.Items.Item("59").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please Provide End Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim oJE As SAPbobsCOM.JournalEntries = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)

            Dim EndDate As String = objForm.Items.Item("59").Specific.Value
            EndDate = EndDate.Insert("4", "-")
            EndDate = EndDate.Insert("7", "-")

            oJE.ReferenceDate = EndDate
            oJE.TaxDate = EndDate
            oJE.DueDate = EndDate

            oJE.UserFields.Fields.Item("U_VSPTRPNO").Value = objForm.Items.Item("16").Specific.Value
            oJE.UserFields.Fields.Item("U_VSPDCTYP").Value = "Trip Sheet Expense"

            Dim LineCount As Integer = 0

            Dim GetVehicleCC As String = ""
            If objMain.IsSAPHANA = True Then
                GetVehicleCC = "Select ""U_VSPVEHCC"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("4").Specific.Value & "'"
            Else
                GetVehicleCC = "Select U_VSPVEHCC From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("4").Specific.Value & "'"
            End If
            Dim oRsGetVehicleCC As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehicleCC.DoQuery(GetVehicleCC)

            Dim GetDriverCC As String = ""
            If objMain.IsSAPHANA = True Then
                GetDriverCC = "Select Distinct ""U_VSPDRVCC"" From ""@VSP_FLT_TRSHT_C2"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            Else
                GetDriverCC = "Select Distinct U_VSPDRVCC From [@VSP_FLT_TRSHT_C2] Where DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            End If
            Dim oRsGetDriverCC As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDriverCC.DoQuery(GetDriverCC)

            'Debit(Expenses)
            Dim GetExpenseAcc As String = ""
            If objMain.IsSAPHANA = True Then
                GetExpenseAcc = "Select ""U_VSPEXACC"" ,Sum(""U_VSPAMT"") as ""Amount"" From ""@VSP_FLT_TRSHT_C4"" Where " & _
                           """DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And ""U_VSPAMT"" > 0 Group By ""U_VSPEXACC"" "
            Else
                GetExpenseAcc = "Select U_VSPEXACC ,Sum(U_VSPAMT) as 'Amount' From [@VSP_FLT_TRSHT_C4] Where " & _
           "DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "' And U_VSPAMT > 0 Group By U_VSPEXACC "
            End If
            Dim oRsGetExpenseAcc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetExpenseAcc.DoQuery(GetExpenseAcc)

            For i As Integer = 1 To oRsGetExpenseAcc.RecordCount
                If LineCount > 0 Then
                    oJE.Lines.Add()
                End If

                oJE.Lines.AccountCode = oRsGetExpenseAcc.Fields.Item("U_VSPEXACC").Value
                oJE.Lines.Debit = oRsGetExpenseAcc.Fields.Item("Amount").Value
                oJE.Lines.CostingCode2 = oRsGetVehicleCC.Fields.Item(0).Value
                oJE.Lines.CostingCode3 = oRsGetDriverCC.Fields.Item(0).Value

                LineCount = 1
                oRsGetExpenseAcc.MoveNext()
            Next

            'Credit(Advances)
            Dim GetAdvanceAcc As String = ""

            If objMain.IsSAPHANA = True Then
                GetAdvanceAcc = "Select ""U_VSPADACC"" ,Sum(""U_VSPAMT"") as ""Amount"" From ""@VSP_FLT_TRSHT_C4"" Where " & _
            """DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'  And ""U_VSPAMT"" > 0 Group By ""U_VSPADACC"""
            Else
                GetAdvanceAcc = "Select U_VSPADACC ,Sum(U_VSPAMT) as 'Amount' From [@VSP_FLT_TRSHT_C4] Where " & _
            "DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'  And U_VSPAMT > 0 Group By U_VSPADACC"
            End If
            Dim oRsGetAdvanceAcc As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetAdvanceAcc.DoQuery(GetAdvanceAcc)

            For i As Integer = 1 To oRsGetAdvanceAcc.RecordCount

                oJE.Lines.Add()
                oJE.Lines.AccountCode = oRsGetAdvanceAcc.Fields.Item("U_VSPADACC").Value
                oJE.Lines.Credit = oRsGetAdvanceAcc.Fields.Item("Amount").Value
                oJE.Lines.CostingCode2 = oRsGetVehicleCC.Fields.Item(0).Value
                oJE.Lines.CostingCode3 = oRsGetDriverCC.Fields.Item(0).Value

                oRsGetAdvanceAcc.MoveNext()
            Next

            'For i As Integer = 1 To objMatrix4.VisualRowCount
            '    If objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value <> "" Then
            '        If objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value > 0 Then
            '            If LineCount > 0 Then
            '                oJE.Lines.Add()
            '            End If

            '            oJE.Lines.AccountCode = objMatrix4.Columns.Item("V_2").Cells.Item(i).Specific.Value
            '            oJE.Lines.Debit = objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value
            '            oJE.Lines.CostingCode2 = oRsGetVehicleCC.Fields.Item(0).Value

            '            oJE.Lines.Add()
            '            oJE.Lines.AccountCode = objMatrix4.Columns.Item("V_4").Cells.Item(i).Specific.Value
            '            oJE.Lines.Credit = objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value
            '            oJE.Lines.CostingCode2 = oRsGetVehicleCC.Fields.Item(0).Value

            '            LineCount = 1
            '        End If
            '    End If
            'Next

            If oJE.Add = 0 Then
                Dim GetTransID As String = ""
                If objMain.IsSAPHANA = True Then
                    GetTransID = "Select Max(""TransId"") From OJDT"
                Else
                    GetTransID = "Select Max(TransId) From OJDT"
                End If
                Dim oRsGetTransID As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetTransID.DoQuery(GetTransID)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_VSPEXPNO", oRsGetTransID.Fields.Item(0).Value.ToString)
                objMain.oGeneralService.Update(objMain.oGeneralData)

                Me.RefreshData(objForm.UniqueID)
                Me.SetCellsEditable(objForm.UniqueID)

            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CloseTripSheet(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")

            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific

            Dim CheckIfAdvancesNotPosted As String = ""

            If objMain.IsSAPHANA = True Then
                CheckIfAdvancesNotPosted = "Select * From ""@VSP_FLT_TRSHT_C2"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' " & _
           "And (""U_VSPJENO"" IS NULL Or ""U_VSPJENO"" = '')"
            Else
                CheckIfAdvancesNotPosted = "Select * From [@VSP_FLT_TRSHT_C2] Where DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "' " & _
           "And (U_VSPJENO IS NULL Or U_VSPJENO = '')"
            End If
            Dim oRsCheckIfAdvancesNotPosted As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCheckIfAdvancesNotPosted.DoQuery(CheckIfAdvancesNotPosted)

            'If oRsCheckIfAdvancesNotPosted.RecordCount > 0 Then
            '    objMain.objApplication.StatusBar.SetText("Please Post Advances", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Exit Try
            'End If


            ''Check Sales Order is Closed or not
            Dim CheckIfSaleNotPosted As String = ""

            If objMain.IsSAPHANA = True Then
                CheckIfSaleNotPosted = "Select * From ""@VSP_FLT_TRSHT_C3"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""U_VSPDOCTY""='Sales Order'  "
            Else
                CheckIfSaleNotPosted = "Select * From ""@VSP_FLT_TRSHT_C3"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "' and ""U_VSPDOCTY""='Sales Order'  "
            End If
            Dim oRsCheckIfSaleNotPosted As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsCheckIfSaleNotPosted.DoQuery(CheckIfSaleNotPosted)
            If oRsCheckIfSaleNotPosted.RecordCount > 0 Then
                oRsCheckIfSaleNotPosted.MoveFirst()
                Dim DocumentNo As String = ""
              
                For i As Integer = 1 To oRsCheckIfSaleNotPosted.RecordCount

                    DocumentNo = oRsCheckIfSaleNotPosted.Fields.Item("U_VSPDCNUM").Value
                    Dim doctype As String = oRsCheckIfSaleNotPosted.Fields.Item("U_VSPDOCTY").Value

                    If doctype = "Sales Order" Then
                        If DocumentNo = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Post Sales Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Try
                        End If
                        If DocumentNo <> "" Then
                            Dim CheckSaleStatus As String = ""

                            If objMain.IsSAPHANA = True Then
                                CheckSaleStatus = "Select ""DocStatus"",""DocNum"" From ORDR Where ""DocNum"" = (Select ""DocNum"" From ORDR where ""DocEntry""='" & DocumentNo & "') "
                            Else
                                CheckSaleStatus = "Select ""DocStatus"",""DocNum"" From ORDR Where ""DocNum"" = (Select ""DocNum"" From ORDR where ""DocEntry""='" & DocumentNo & "') "
                            End If
                            Dim oRsCheckSaleStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsCheckSaleStatus.DoQuery(CheckSaleStatus)


                            If oRsCheckSaleStatus.RecordCount > 0 Then
                                Dim status As String = ""
                                status = oRsCheckSaleStatus.Fields.Item(0).Value.ToString()
                                Dim docnum As String = oRsCheckSaleStatus.Fields.Item(1).Value.ToString()
                                If status = "O" Then
                                    objMain.objApplication.StatusBar.SetText("TripSheet Can Not Be Close ...Sales Order Is Open With Order No : " & docnum, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Try
                                End If

                            End If
                        End If


                    End If
                    oRsCheckIfSaleNotPosted.MoveNext()
                Next
            Else

                objMain.objApplication.StatusBar.SetText("TripSheet Can Not Be Close  Without Sales Order", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try

            End If



            Dim ExpenseAmount As Double = 0
            For i As Integer = 1 To objMatrix4.VisualRowCount
                If objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value <> "" Then
                    ExpenseAmount = ExpenseAmount + objMatrix4.Columns.Item("V_5").Cells.Item(i).Specific.Value
                End If
            Next
            If ExpenseAmount > 0 And objForm.Items.Item("1000001").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please Post Expenses", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim oChk As SAPbouiCOM.CheckBox = objForm.Items.Item("79").Specific
            If oChk.Checked = False Then
                objMain.objApplication.StatusBar.SetText("Please Update Tyre Details", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            If objMain.objApplication.MessageBox("Closing Trip Sheet is irreversable. Do you want to continue......", 2, "Ok", "Cancel") = 2 Then
                Exit Try
            End If

            objMain.objApplication.StatusBar.SetText("Please Wait While Trip Sheet Closes.........", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            'If Me.PostDieselConsumption(objForm.UniqueID) = False Then Exit Try

            Dim GetTotalKM As String = ""

            If objMain.IsSAPHANA = True Then
                GetTotalKM = "Select IFNULL(SUM(IFNULL(""U_VSPDICON"",0)),0) , IFNULL(SUM(IFNULL(""U_TOTKM"",0)),0) From ""@VSP_FLT_TRSHT_C1"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            Else
                GetTotalKM = "Select ISNULL(SUM(ISNULL(U_VSPDICON,0)),0) , ISNULL(SUM(ISNULL(U_TOTKM,0)),0) From [@VSP_FLT_TRSHT_C1] Where DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            End If
            Dim oRsGetTotalKM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTotalKM.DoQuery(GetTotalKM)
            If oRsGetTotalKM.Fields.Item(0).Value > 0 Then
                oDBs_Head.SetValue("U_VSPTOTKM", oDBs_Head.Offset, CDbl(oRsGetTotalKM.Fields.Item(1).Value))
                oDBs_Head.SetValue("U_VSPMLGE", oDBs_Head.Offset, Math.Round(CDbl(oRsGetTotalKM.Fields.Item(1).Value) / CDbl(oRsGetTotalKM.Fields.Item(0).Value), 3))

                Dim GetTotalDesielQty As String = ""

                If objMain.IsSAPHANA = True Then
                    GetTotalDesielQty = "Select IFNULL(SUM(IFNULL(""U_VSPQUAN"",0)),0)  From ""@VSP_FLT_TRSHT_C5"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
                Else
                    GetTotalDesielQty = "Select ISNULL(SUM(ISNULL(U_VSPQUAN,0)),0)  From [@VSP_FLT_TRSHT_C5] Where DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
                End If
                Dim oRsGetTotalDesielQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetTotalDesielQty.DoQuery(GetTotalDesielQty)
                If oRsGetTotalDesielQty.Fields.Item(0).Value > 0 Then
                    oDBs_Head.SetValue("U_VSPACMLG", oDBs_Head.Offset, Math.Round(CDbl(oRsGetTotalKM.Fields.Item(1).Value) / CDbl(oRsGetTotalDesielQty.Fields.Item(0).Value), 3))
                End If
            End If

            Dim GetVehicleCode As String = ""

            If objMain.IsSAPHANA = True Then
                GetVehicleCode = "Select ""Code"" , IFNULL(""U_VSPODRDG"",0) From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("4").Specific.Value & "'"
            Else
                GetVehicleCode = "Select Code , ISNULL(U_VSPODRDG,0) From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("4").Specific.Value & "'"
            End If
            Dim oRsGetVehicleCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehicleCode.DoQuery(GetVehicleCode)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", oRsGetVehicleCode.Fields.Item(0).Value)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            Dim X As Double = (CDbl(oRsGetVehicleCode.Fields.Item(1).Value) + CDbl(oRsGetTotalKM.Fields.Item(1).Value)).ToString
            objMain.oGeneralData.SetProperty("U_VSPODRDG", (CDbl(oRsGetVehicleCode.Fields.Item(1).Value) + CDbl(oRsGetTotalKM.Fields.Item(1).Value)).ToString)
            objMain.oGeneralData.SetProperty("U_VSPAVLB", "Y")
            objMain.oGeneralService.Update(objMain.oGeneralData)

            Dim GetDays As String = ""
            If objMain.IsSAPHANA = True Then
                GetDays = "Select DAYS_BETWEEN(TO_DATE(MIN(""U_VSPFRDT"")),TO_DATE(MAX(""U_VSPTODT""))) + 1 From ""@VSP_FLT_TRSHT_C1"" Where " & _
          """DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            Else
                GetDays = "Select DATEDIFF(D,MIN(U_VSPFRDT),MAX(U_VSPTODT)) + 1 From [@VSP_FLT_TRSHT_C1] Where " & _
            "DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"
            End If
            Dim oRsGetDays As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDays.DoQuery(GetDays)
            oDBs_Head.SetValue("U_VSPTOTDY", oDBs_Head.Offset, oRsGetDays.Fields.Item(0).Value)

            oDBs_Head.SetValue("U_VSPSTS", oDBs_Head.Offset, "Close")

            Me.UpdateVehicleStatus(objForm.UniqueID)

            For i As Integer = 1 To objMatrix8.VisualRowCount
                Dim GetKM As Double = 0
                If objMatrix8.Columns.Item("V_4").Cells.Item(i).Specific.Value <> "" Then
                    GetKM = objMatrix8.Columns.Item("V_4").Cells.Item(i).Specific.Value
                End If
                Dim L As Double = CDbl(oRsGetTotalKM.Fields.Item(1).Value)
                GetKM = GetKM + CDbl(oRsGetTotalKM.Fields.Item(1).Value)

                oDBs_Detail8.SetValue("LineId", oDBs_Detail8.Offset, i)
                oDBs_Detail8.SetValue("U_VSPTYRNO", oDBs_Detail8.Offset, objMatrix8.Columns.Item("V_0").Cells.Item(i).Specific.Value)
                oDBs_Detail8.SetValue("U_VSPTYRNM", oDBs_Detail8.Offset, objMatrix8.Columns.Item("V_1").Cells.Item(i).Specific.Value)
                oDBs_Detail8.SetValue("U_VSPTYMOD", oDBs_Detail8.Offset, objMatrix8.Columns.Item("V_2").Cells.Item(i).Specific.Value)
                oDBs_Detail8.SetValue("U_VSPTYPOS", oDBs_Detail8.Offset, objMatrix8.Columns.Item("V_3").Cells.Item(i).Specific.Value)
                oDBs_Detail8.SetValue("U_VSPKMS", oDBs_Detail8.Offset, GetKM)
                objMatrix8.SetLineData(i)
            Next

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Me.SetCellsEditable(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpdateVehicleStatus(ByVal FormUID As String)

        Try
            Dim GetVehicleStsDocEntry As String = ""

            If objMain.IsSAPHANA = True Then
                GetVehicleStsDocEntry = "Select MAX (""DocEntry"") From ""@VSP_VECHSTS"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("4").Specific.Value & "'   "
            Else
                GetVehicleStsDocEntry = "Select MAX (DocEntry) From [@VSP_VECHSTS] Where U_VSPVNO = '" & objForm.Items.Item("4").Specific.Value & "'   "
            End If
            Dim oRsGetVehicleStsDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehicleStsDocEntry.DoQuery(GetVehicleStsDocEntry)

            If oRsGetVehicleStsDocEntry.RecordCount > 0 Then

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVEHST")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oRsGetVehicleStsDocEntry.Fields.Item(0).Value)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_VSPSTS", objForm.Items.Item("1000002").Specific.Value)
                objMain.oGeneralService.Update(objMain.oGeneralData)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

    Sub RefreshData(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")
            oDBs_Detail1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1")
            oDBs_Detail2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2")
            oDBs_Detail3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3")
            oDBs_Detail4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4")
            oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")
            oDBs_Detail6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C6")
            oDBs_Detail7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7")
            oDBs_Detail8 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8")
            oDBs_Detail9 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9")

            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C1"), "DocEntry", oDBs_Detail1.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C2"), "DocEntry", oDBs_Detail2.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C3"), "DocEntry", oDBs_Detail3.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C4"), "DocEntry", oDBs_Detail4.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5"), "DocEntry", oDBs_Detail5.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C6"), "DocEntry", oDBs_Detail6.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C7"), "DocEntry", oDBs_Detail7.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C8"), "DocEntry", oDBs_Detail8.GetValue("DocEntry", 0))
            objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9"), "DocEntry", oDBs_Detail9.GetValue("DocEntry", 0))

            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix3 = objForm.Items.Item("1000023").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix5 = objForm.Items.Item("67").Specific
            objMatrix6 = objForm.Items.Item("1000028").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific
            objMatrix9 = objForm.Items.Item("1000030").Specific

            objMatrix1.LoadFromDataSource()
            objMatrix2.LoadFromDataSource()
            objMatrix3.LoadFromDataSource()
            objMatrix4.LoadFromDataSource()
            objMatrix5.LoadFromDataSource()
            objMatrix6.LoadFromDataSource()
            objMatrix7.LoadFromDataSource()
            objMatrix8.LoadFromDataSource()
            objMatrix9.LoadFromDataSource()

            objMatrix1.AutoResizeColumns()
            objMatrix2.AutoResizeColumns()
            objMatrix3.AutoResizeColumns()
            objMatrix4.AutoResizeColumns()
            objMatrix5.AutoResizeColumns()
            objMatrix6.AutoResizeColumns()
            objMatrix7.AutoResizeColumns()
            objMatrix8.AutoResizeColumns()
            objMatrix9.AutoResizeColumns()
            objForm.Refresh()

            objForm.Freeze(False)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetCellsEditable(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            'objMatrix3 = objForm.Items.Item("1000023").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix5 = objForm.Items.Item("67").Specific
            objMatrix6 = objForm.Items.Item("1000028").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific
            objMatrix9 = objForm.Items.Item("1000030").Specific

            'Advance Matrix
            For i As Integer = 1 To objMatrix2.VisualRowCount
                If objMatrix2.Columns.Item("V_7").Cells.Item(i).Specific.Value = "" Then
                    objMatrix2.CommonSetting.SetRowEditable(i, True)
                Else
                    objMatrix2.CommonSetting.SetRowEditable(i, False)
                End If
            Next
            objMatrix2.Columns.Item("V_-1").Editable = False
            objMatrix2.Columns.Item("V_7").Editable = False
            objMatrix2.Columns.Item("V_9").Editable = False

            'Expense Matrix
            For i As Integer = 1 To objMatrix4.VisualRowCount
                If objMatrix4.Columns.Item("V_1").Cells.Item(i).Specific.Checked = True Then
                    objMatrix4.CommonSetting.SetRowEditable(i, False)
                Else
                    objMatrix4.CommonSetting.SetRowEditable(i, True)
                End If
            Next
            objMatrix4.Columns.Item("V_-1").Editable = False
            objMatrix4.Columns.Item("V_1").Editable = False
            objMatrix4.Columns.Item("V_0").Editable = True
            objMatrix4.Columns.Item("V_2").Editable = True
            objMatrix4.Columns.Item("V_3").Editable = False
            objMatrix4.Columns.Item("V_4").Editable = True
            objMatrix4.Columns.Item("V_5").Editable = True
            objMatrix4.Columns.Item("V_6").Editable = False

            Dim EV_2 As Integer
            For i As Integer = 0 To objMatrix4.Columns.Count - 1
                If objMatrix4.Columns.Item(i).UniqueID = "V_2" Then
                    EV_2 = i
                    Exit For
                End If
            Next
            For i As Integer = 1 To objMatrix4.VisualRowCount
                If objMatrix4.Columns.Item("V_1").Cells.Item(i).Specific.Checked = True Then
                    objMatrix4.CommonSetting.SetCellEditable(i, EV_2, False)
                Else
                    objMatrix4.CommonSetting.SetCellEditable(i, EV_2, True)
                End If
            Next

            'Diesel Matrix
            For i As Integer = 1 To objMatrix5.VisualRowCount
                If objMatrix5.Columns.Item("V_7").Cells.Item(i).Specific.Value <> "" Then
                    objMatrix5.CommonSetting.SetRowEditable(i, False)
                Else
                    objMatrix5.CommonSetting.SetRowEditable(i, True)
                End If
            Next
            objMatrix5.Columns.Item("V_-1").Editable = False
            objMatrix5.Columns.Item("V_2").Editable = False
            objMatrix5.Columns.Item("V_5").Editable = False
            objMatrix5.Columns.Item("V_7").Editable = False
            objMatrix5.Columns.Item("V_6").Editable = False

            'Revenue Matrix
            'Dim RV_0, RV_1, RV_2, RV_3, RV_9, RV_10 As Integer
            'For i As Integer = 0 To objMatrix3.Columns.Count - 1
            '    If objMatrix3.Columns.Item(i).UniqueID = "V_0" Then
            '        RV_0 = i
            '        Exit For
            '    End If
            'Next
            'For i As Integer = 0 To objMatrix3.Columns.Count - 1
            '    If objMatrix3.Columns.Item(i).UniqueID = "V_1" Then
            '        RV_1 = i
            '        Exit For
            '    End If
            'Next
            'For i As Integer = 0 To objMatrix3.Columns.Count - 1
            '    If objMatrix3.Columns.Item(i).UniqueID = "V_2" Then
            '        RV_2 = i
            '        Exit For
            '    End If
            'Next
            'For i As Integer = 0 To objMatrix3.Columns.Count - 1
            '    If objMatrix3.Columns.Item(i).UniqueID = "V_9" Then
            '        RV_9 = i
            '        Exit For
            '    End If
            'Next
            'For i As Integer = 0 To objMatrix3.Columns.Count - 1
            '    If objMatrix3.Columns.Item(i).UniqueID = "V_10" Then
            '        RV_10 = i
            '        Exit For
            '    End If
            'Next
            'For i As Integer = 1 To objMatrix3.VisualRowCount
            '    If objMatrix3.Columns.Item("V_5").Cells.Item(i).Specific.Value <> "" Then
            '        objMatrix3.CommonSetting.SetCellEditable(i, RV_0, False)
            '        objMatrix3.CommonSetting.SetCellEditable(i, RV_1, False)
            '        objMatrix3.CommonSetting.SetCellEditable(i, RV_2, False)
            '        objMatrix3.CommonSetting.SetCellEditable(i, RV_9, False)
            '        objMatrix3.CommonSetting.SetCellEditable(i, RV_10, False)
            '    End If
            'Next

            If objForm.Items.Item("79").Specific.Checked = True Then
                objForm.Items.Item("74").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000029").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Else
                objForm.Items.Item("74").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("74").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000029").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If

            'Closing
            If objForm.Items.Item("1000002").Specific.Selected.Value = "Close" Then
                objForm.Items.Item("61").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("61").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000022").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000022").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000023").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000023").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("67").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("67").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("69").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("69").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000031").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("1000032").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("1000033").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("1000035").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("1000036").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("80").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                objForm.Items.Item("59").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            Else
                objForm.Items.Item("61").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000022").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000023").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("67").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("69").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000031").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000031").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000032").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000032").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000033").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000033").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000035").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000035").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("80").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("80").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000036").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("59").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            End If

            If objForm.Items.Item("1000001").Specific.Value = "" Then
                objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000033").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000033").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                objForm.Items.Item("1000033").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefaultCellsEditable(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)
            objMatrix1 = objForm.Items.Item("61").Specific
            objMatrix2 = objForm.Items.Item("1000022").Specific
            objMatrix3 = objForm.Items.Item("1000023").Specific
            objMatrix4 = objForm.Items.Item("1000024").Specific
            objMatrix5 = objForm.Items.Item("67").Specific
            objMatrix6 = objForm.Items.Item("1000028").Specific
            objMatrix7 = objForm.Items.Item("69").Specific
            objMatrix8 = objForm.Items.Item("1000029").Specific
            objMatrix9 = objForm.Items.Item("1000030").Specific

            For i As Integer = 1 To objMatrix2.VisualRowCount
                objMatrix2.CommonSetting.SetRowEditable(i, True)
            Next
            objMatrix2.Columns.Item("V_-1").Editable = False
            objMatrix2.Columns.Item("V_7").Editable = False
            objMatrix2.Columns.Item("V_9").Editable = False

            For i As Integer = 1 To objMatrix5.VisualRowCount
                objMatrix5.CommonSetting.SetRowEditable(i, True)
            Next
            objMatrix5.Columns.Item("V_-1").Editable = False
            objMatrix5.Columns.Item("V_2").Editable = False
            objMatrix5.Columns.Item("V_5").Editable = False
            objMatrix5.Columns.Item("V_7").Editable = False

            objForm.Items.Item("61").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("1000022").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("1000023").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("1000024").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("67").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("69").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("1000029").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("59").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("1000036").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function GoodsIssues(ByVal FormUID As String, ByVal i As Integer)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Detail5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C5")

            Dim oGoodsIssue As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

            Dim DocDate As String = objMatrix5.Columns.Item("V_0").Cells.Item(i).Specific.Value
            DocDate = DocDate.Insert("4", "-")
            DocDate = DocDate.Insert("7", "-")

            oGoodsIssue.DocDate = DocDate
            oGoodsIssue.TaxDate = DocDate

            Dim GetWhse As String = ""

            If objMain.IsSAPHANA = True Then
                GetWhse = "Select ""U_VSPDEWHS"" From ""@VSP_FLT_CNFGSRN"" Where ""DocEntry"" = '1'"
            Else
                GetWhse = "Select U_VSPDEWHS From [@VSP_FLT_CNFGSRN] Where DocEntry = '1'"
            End If
            Dim oRsGetWhse As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWhse.DoQuery(GetWhse)

            oGoodsIssue.Lines.ItemCode = objForm.Items.Item("1000027").Specific.Value
            oGoodsIssue.Lines.WarehouseCode = oRsGetWhse.Fields.Item(0).Value
            oGoodsIssue.Lines.Quantity = objMatrix5.Columns.Item("V_3").Cells.Item(i).Specific.Value   'Math.Round(oRsGetTotalQty.Fields.Item(0).Value, 3)
            oGoodsIssue.Lines.COGSCostingCode3 = objMatrix5.Columns.Item("V_8").Cells.Item(i).Specific.Value
            If oGoodsIssue.Add = 0 Then


                Dim GetDocEntry As String = ""

                If objMain.IsSAPHANA = True Then
                    GetDocEntry = "Select Max(""DocEntry"") From OIGE"
                Else
                    GetDocEntry = "Select Max(DocEntry) From OIGE"
                End If
                Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDocEntry.DoQuery(GetDocEntry)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C5")
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPGISU", oRsGetDocEntry.Fields.Item(0).Value.ToString)
                objMain.oGeneralService.Update(objMain.oGeneralData)
                objMain.oGeneralService.Update(objMain.oGeneralData)

                Me.RefreshData(objForm.UniqueID)
                Return True
            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            ' End If
            Me.RefreshData(objForm.UniqueID)
            Me.SetCellsEditable(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function PostDieselConsumption(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")

            Dim GetTotalQty As String = ""

            If objMain.IsSAPHANA = True Then
                GetTotalQty = "Select IfNULL(SUM(IFNULL(""U_VSPDICON"",0)),0) From ""@VSP_FLT_TRSHT_C1"" Where ""DocEntry"" = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"

            Else
                GetTotalQty = "Select ISNULL(SUM(ISNULL(U_VSPDICON,0)),0) From [@VSP_FLT_TRSHT_C1] Where DocEntry = '" & oDBs_Head.GetValue("DocEntry", 0) & "'"

            End If
            Dim oRsGetTotalQty As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTotalQty.DoQuery(GetTotalQty)

            If oRsGetTotalQty.Fields.Item(0).Value > 0 Then
                Dim oGoodsIssue As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

                Dim DocDate As String = objForm.Items.Item("59").Specific.Value
                DocDate = DocDate.Insert("4", "-")
                DocDate = DocDate.Insert("7", "-")

                oGoodsIssue.DocDate = DocDate
                oGoodsIssue.TaxDate = DocDate

                Dim GetWhse As String = ""
                If objMain.IsSAPHANA = True Then
                    GetWhse = "Select ""U_VSPDEWHS"" From ""@VSP_FLT_CNFGSRN"" Where ""DocEntry"" = '1'"

                Else
                    GetWhse = "Select U_VSPDEWHS From [@VSP_FLT_CNFGSRN] Where DocEntry = '1'"

                End If
                Dim oRsGetWhse As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetWhse.DoQuery(GetWhse)

                oGoodsIssue.Lines.ItemCode = objForm.Items.Item("1000027").Specific.Value
                oGoodsIssue.Lines.WarehouseCode = oRsGetWhse.Fields.Item(0).Value
                oGoodsIssue.Lines.Quantity = Math.Round(oRsGetTotalQty.Fields.Item(0).Value, 3)

                If oGoodsIssue.Add = 0 Then
                    Dim GetDocEntry As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetDocEntry = "Select Max(""DocEntry"") From OIGE"
                    Else
                        GetDocEntry = "Select Max(DocEntry) From OIGE"
                    End If
                    Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetDocEntry.DoQuery(GetDocEntry)

                    objMain.sCmp = objMain.objCompany.GetCompanyService
                    objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                    objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                    objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    objMain.oGeneralData.SetProperty("U_VSPGISU", oRsGetDocEntry.Fields.Item(0).Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)

                    Me.RefreshData(objForm.UniqueID)
                    Return True
                Else
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function UpdateTyreDetails(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")

            objMatrix8 = objForm.Items.Item("1000029").Specific

            For i As Integer = 1 To objMatrix8.VisualRowCount

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTYRMSTR")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)

                If objMatrix8.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                    Dim GetCode As String = ""
                    If objMain.IsSAPHANA = True Then
                        GetCode = "Select ""Code"" From ""@VSP_FLT_TYRMSTR"" Where ""U_VSPTRNUM"" = '" & objMatrix8.Columns.Item("V_0").Cells.Item(i).Specific.Value & "'"
                    Else
                        GetCode = "Select Code From [@VSP_FLT_TYRMSTR] Where U_VSPTRNUM = '" & objMatrix8.Columns.Item("V_0").Cells.Item(i).Specific.Value & "'"
                    End If
                    Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRsGetCode.DoQuery(GetCode)

                    objMain.oGeneralParams.SetProperty("Code", oRsGetCode.Fields.Item(0).Value)
                    objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                    objMain.oGeneralData.SetProperty("U_VSPKMRUN", objMatrix8.Columns.Item("V_4").Cells.Item(i).Specific.Value.ToString)
                    objMain.oGeneralService.Update(objMain.oGeneralData)
                End If
            Next

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_VSPTYUP", "Y")
            objMain.oGeneralService.Update(objMain.oGeneralData)

            Me.RefreshData(objForm.UniqueID)
            Me.SetCellsEditable(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objForm.Items.Item("4").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle Number Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("12").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Route Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("18").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Document Date Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("20").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Start Date Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function PostDiesel(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT")

            objMatrix5 = objForm.Items.Item("67").Specific

            Dim i As Integer = 0
            Dim LineSelect As Boolean = False
            For i = 1 To objMatrix5.VisualRowCount
                If objMatrix5.IsRowSelected(i) = True Then
                    LineSelect = True
                    Exit For
                End If
            Next

            If LineSelect = False Then Exit Try

            If objMatrix5.Columns.Item("V_7").Cells.Item(i).Specific.Value <> "" Then
                objMain.objApplication.StatusBar.SetText("Document Posted", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix5.Columns.Item("V_0").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please Provide Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix5.Columns.Item("V_1").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please Provide Vendor Code", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf objMatrix5.Columns.Item("V_3").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please Provide Quantity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim oGRPO As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

            oGRPO.CardCode = objMatrix5.Columns.Item("V_1").Cells.Item(i).Specific.Value

            Dim DocDate As String = objMatrix5.Columns.Item("V_0").Cells.Item(i).Specific.Value
            DocDate = DocDate.Insert("4", "-")
            DocDate = DocDate.Insert("7", "-")
            oGRPO.DocDate = DocDate
            oGRPO.DocDueDate = DocDate
            oGRPO.TaxDate = DocDate

            oGRPO.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
            oGRPO.Lines.ItemCode = objForm.Items.Item("1000027").Specific.Value

            Dim GetTaxCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetTaxCode = "Select ""U_VSPTXCD"" , ""U_VSPDEWHS"" From ""@VSP_FLT_CNFGSRN"" Where ""DocNum"" = '1'"
            Else
                GetTaxCode = "Select U_VSPTXCD , U_VSPDEWHS From [@VSP_FLT_CNFGSRN] Where DocNum = '1'"
            End If
            Dim oRsGetTaxCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetTaxCode.DoQuery(GetTaxCode)



            Dim GetInStock As String = ""

            If objMain.IsSAPHANA = True Then
                GetInStock = "Select ""OnHand""  From OITW Where ""WhsCode"" ='" & oRsGetTaxCode.Fields.Item(1).Value & "' " & _
                                        "And ""ItemCode"" = '" & objForm.Items.Item("1000027").Specific.Value & "' "
            Else
                GetInStock = "Select OnHand  From OITW Where WhsCode ='" & oRsGetTaxCode.Fields.Item(1).Value & "' " & _
                                        "And ItemCode = '" & objForm.Items.Item("1000027").Specific.Value & "' "
            End If
            Dim oRsGetInStock As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetInStock.DoQuery(GetInStock)

            Dim GetMaxCapacity As String = ""
            If objMain.IsSAPHANA = True Then
                GetMaxCapacity = "Select ""MaxLevel"" From OITM Where ""ItemCode""= '" & objForm.Items.Item("1000027").Specific.Value & "'"
            Else
                GetMaxCapacity = "Select MaxLevel From OITM Where ItemCode = '" & objForm.Items.Item("1000027").Specific.Value & "'"
            End If
            Dim oRsGetMaxCapacity As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetMaxCapacity.DoQuery(GetMaxCapacity)

            Dim CalculateQty As Double = oRsGetInStock.Fields.Item("OnHand").Value + objMatrix5.Columns.Item("V_3").Cells.Item(i).Specific.Value

            If oRsGetMaxCapacity.Fields.Item("MaxLevel").Value > CalculateQty Then
                oGRPO.Lines.Quantity = objMatrix5.Columns.Item("V_3").Cells.Item(i).Specific.Value
            Else
                objMain.objApplication.StatusBar.SetText("Sum of Quantity & InStockQty of the Item exceeds Maximum Diesel Tank Capacity", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            oGRPO.Lines.UnitPrice = objMatrix5.Columns.Item("V_4").Cells.Item(i).Specific.Value
            oGRPO.Lines.WarehouseCode = oRsGetTaxCode.Fields.Item(1).Value
            oGRPO.Lines.TaxCode = oRsGetTaxCode.Fields.Item(0).Value
            oGRPO.Lines.LineTotal = objMatrix5.Columns.Item("V_5").Cells.Item(i).Specific.Value

            If oGRPO.Add = 0 Then

                Dim GetDocEntry As String = ""
                If objMain.IsSAPHANA = True Then
                    GetDocEntry = "Select Max(""DocEntry"") From OPDN"
                Else
                    GetDocEntry = "Select Max(DocEntry) From OPDN"
                End If
                Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDocEntry.DoQuery(GetDocEntry)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C5")
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPDCNUM", oRsGetDocEntry.Fields.Item(0).Value.ToString)
                objMain.oGeneralService.Update(objMain.oGeneralData)

                Me.GoodsIssues(objForm.UniqueID, i)

            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription & "Line : " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            Me.RefreshData(objForm.UniqueID)
            Me.SetCellsEditable(objForm.UniqueID)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub GenerateInvoice(ByVal FormUID As String, ByVal i As Integer)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix3 = objForm.Items.Item("1000023").Specific
            Dim PricePerKM As Double = 0
            Dim KMS As Double = 0

            If objMatrix3.Columns.Item("V_3").Cells.Item(i).Specific.Value <> "" Then
                PricePerKM = objMatrix3.Columns.Item("V_3").Cells.Item(i).Specific.Value
            End If

            If objMatrix3.Columns.Item("V_4").Cells.Item(i).Specific.Value <> "" Then
                KMS = objMatrix3.Columns.Item("V_4").Cells.Item(i).Specific.Value
            End If

            If objMatrix3.Columns.Item("V_10").Cells.Item(i).Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Please Provide Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            ElseIf PricePerKM * KMS = 0 Then
                objMain.objApplication.StatusBar.SetText("Document Cannot Be Raised for Zero Amount", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Try
            End If

            Dim oInvoice As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            Dim GetCustCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetCustCode = "Select ""U_VSPCUSCD"" , ""U_VSPOACT"" , ""U_VSPTXCD"" , ""U_VSPLOCCD"" From ""@VSP_FLT_CNFGSRN"" Where ""DocNum"" = '1'"
            Else
                GetCustCode = "Select U_VSPCUSCD , U_VSPOACT , U_VSPTXCD , U_VSPLOCCD From [@VSP_FLT_CNFGSRN] Where DocNum = '1'"
            End If
            Dim oRsGetCustCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCustCode.DoQuery(GetCustCode)

            oInvoice.CardCode = oRsGetCustCode.Fields.Item(0).Value

            Dim DocDate As String = objMatrix3.Columns.Item("V_10").Cells.Item(i).Specific.Value
            DocDate = DocDate.Insert("4", "-")
            DocDate = DocDate.Insert("7", "-")
            oInvoice.DocDate = DocDate
            oInvoice.TaxDate = DocDate
            oInvoice.DocDueDate = DocDate

            oInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
            oInvoice.Lines.AccountCode = oRsGetCustCode.Fields.Item(1).Value
            oInvoice.Lines.TaxCode = oRsGetCustCode.Fields.Item(2).Value
            oInvoice.Lines.LocationCode = oRsGetCustCode.Fields.Item(3).Value
            oInvoice.Lines.LineTotal = PricePerKM * KMS

            If oInvoice.Add = 0 Then
                Dim GetDocEntry As String = ""
                If objMain.IsSAPHANA = True Then
                    GetDocEntry = "Select Max(""DocEntry"") From OINV"
                Else
                    GetDocEntry = "Select Max(DocEntry) From OINV"
                End If
                Dim oRsGetDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetDocEntry.DoQuery(GetDocEntry)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)

                objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C3")
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPDCNUM", oRsGetDocEntry.Fields.Item(0).Value.ToString)
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPBPCOD", oRsGetCustCode.Fields.Item(0).Value.ToString)
                'objMain.oChildren.Item(i - 1).SetProperty("U_VSPBPCOD", oRsGetCustCode.Fields.Item(0).Value.ToString)
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPDOCTY", "Invoice")
                objMain.oChildren.Item(i - 1).SetProperty("U_VSPGENTY", "New")
                objMain.oGeneralService.Update(objMain.oGeneralData)
            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            Me.RefreshData(objForm.UniqueID)
            Me.SetCellsEditable(objForm.UniqueID)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpadateOdmtrRdngsFromVM(ByVal FormUID As String)

        Try
            Dim GetVehiclefromVM As String = ""

            If objMain.IsSAPHANA = True Then
                GetVehiclefromVM = "Select  IFNULL(""U_VSPCLSRD"",0) From ""@VSP_FLT_VMSTR_C7"" T0 Inner Join ""@VSP_FLT_VMSTR"" T1 On T1.""Code""  = T0.""Code"" " & _
               "where  T1.""U_VSPVNO"" = '" & objForm.Items.Item("4").Specific.Value & "' and  ""LineId"" = (Select  Max(""LineId"")-1 From ""@VSP_FLT_VMSTR_C7"") And ""U_VSPCHK"" = 'Y' "
            Else
                GetVehiclefromVM = "Select  ISNULL(U_VSPCLSRD,0) From [@VSP_FLT_VMSTR_C7] T0 Inner Join [@VSP_FLT_VMSTR] T1 On T1.Code  = T0.Code " & _
                               "where  T1.U_VSPVNO = '" & objForm.Items.Item("4").Specific.Value & "' and  LineId = (Select  Max(LineId)-1 From [@VSP_FLT_VMSTR_C7]) And U_VSPCHK = 'Y' "
            End If
            Dim oRsGetVehiclefromVM As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehiclefromVM.DoQuery(GetVehiclefromVM)

            Dim GetVehicleCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetVehicleCode = "Select Max(""LineId"") from ""@VSP_FLT_TRSHT"" T0 Inner Join ""@VSP_FLT_TRSHT_C1"" T1 ON T0.""DocEntry"" =  T1.""DocEntry"" " & _
                                           "Where T0.""U_VSPVHCL"" = '" & objForm.Items.Item("4").Specific.Value & "' and T0.""DocEntry"" = '" & objForm.Items.Item("16").Specific.Value & "' "

            Else
                GetVehicleCode = "Select Max(LineId) from [@VSP_FLT_TRSHT] T0 Inner Join [@VSP_FLT_TRSHT_C1] T1 ON T0.DocEntry =  T1.DocEntry " & _
                                           "Where T0.U_VSPVHCL = '" & objForm.Items.Item("4").Specific.Value & "' and T0.DocEntry = '" & objForm.Items.Item("16").Specific.Value & "' "

            End If
            Dim oRsGetVehicleCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetVehicleCode.DoQuery(GetVehicleCode)

            Dim LineId As Integer = CInt(oRsGetVehicleCode.Fields.Item(0).Value)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OTRSHT")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("DocEntry", objForm.Items.Item("16").Specific.Value)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oChildren = objMain.oGeneralData.Child("VSP_FLT_TRSHT_C1")
            objMain.oChildren.Item(LineId - 1).SetProperty("U_VSPCLKM", oRsGetVehiclefromVM.Fields.Item(0).Value)
            objMain.oGeneralService.Update(objMain.oGeneralData)


            objMain.objApplication.StatusBar.SetText("Odometer Readings Sucessfully Updated ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Attachments"

    Sub BrowseFileDialog()
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA)
                ShowFolderBrowserThread.Start()

            ElseIf ShowFolderBrowserThread.ThreadState = ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.StackTrace)
        End Try
    End Sub

    Sub ShowFolderBrowser()

        Dim MyTest1 As New OpenFileDialog
        Dim MyProcs() As Process
        Try
            oDBs_Detail9 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TRSHT_C9")
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"
                    objMatrix9 = objForm.Items.Item("1000030").Specific

                    'Windows XP
                    If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                        Try
                            oDBs_Detail9.SetValue("LineId", oDBs_Detail9.Offset, Path)
                            oDBs_Detail9.SetValue("U_VSPATNM", oDBs_Detail9.Offset, objMatrix9.Columns.Item("V_0").Cells.Item(CInt(Path)).Specific.Value)
                            oDBs_Detail9.SetValue("U_VSPAPTH", oDBs_Detail9.Offset, MyTest1.FileName)
                            objMatrix9.SetLineData(Path)
                            If Path = objMatrix9.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000030")
                            End If
                        Catch ex As IO.IOException
                            objMain.objApplication.MessageBox(ex.Message)
                            Exit Sub
                        End Try

                        'Windows 7

                    ElseIf MyTest1.ShowDialog() = DialogResult.OK Then
                        Try
                            oDBs_Detail9.SetValue("LineId", oDBs_Detail9.Offset, Path)
                            oDBs_Detail9.SetValue("U_VSPATNM", oDBs_Detail9.Offset, objMatrix9.Columns.Item("V_0").Cells.Item(CInt(Path)).Specific.Value)
                            oDBs_Detail9.SetValue("U_VSPAPTH", oDBs_Detail9.Offset, MyTest1.FileName)
                            objMatrix9.SetLineData(Path)
                            If Path = objMatrix9.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000030")
                            End If
                        Catch ex As IO.IOException
                            objMain.objApplication.MessageBox(ex.Message)
                            Exit Sub
                        End Try

                        System.Windows.Forms.Application.ExitThread()
                    End If
                Next
            Else
                Console.WriteLine("No SBO instances found.")
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Sub

#End Region

#End Region

End Class













