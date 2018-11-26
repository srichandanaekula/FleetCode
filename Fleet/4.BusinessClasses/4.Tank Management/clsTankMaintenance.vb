Imports System.Threading
Imports System.IO
Public Class clsTankMaintenance
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
    Dim Path As String
    Dim oLink As SAPbouiCOM.LinkedButton
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Tank Maintenance.xml", "VSP_FLT_TNKMTNC_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TNKMTNC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2")

            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objMatrix2 = objForm.Items.Item("19").Specific
            
            objForm.Items.Item("16").AffectsFormMode = False
            objForm.Items.Item("17").AffectsFormMode = False

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OTNKMTNC"))
            oDBs_Head.SetValue("U_VSPDCDT", oDBs_Head.Offset, DateTime.Today.ToString("yyyyMMdd"))

            objMatrix2 = objForm.Items.Item("19").Specific
            objMatrix2.Clear()
            oDBs_Details2.Clear()
            objMatrix2.FlushToDataSource()
            objMatrix2.AutoResizeColumns()

            objMatrix3 = objForm.Items.Item("20").Specific
            objMatrix3.Clear()
            oDBs_Details3.Clear()
            objMatrix3.FlushToDataSource()
            objMatrix3.AutoResizeColumns()

            Me.SetNewLine(objForm.UniqueID, "19")
            Me.SetNewLine(objForm.UniqueID, "20")

            objForm.PaneLevel = 2
            objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2")

            objMatrix2 = objForm.Items.Item("19").Specific
            objMatrix3 = objForm.Items.Item("20").Specific

            Select Case MatrixUID
                Case "19"
                    objMatrix2.AddRow()
                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Details2.SetValue("U_VSPDOCTY", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDOCNM", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDATE", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDCTOT", oDBs_Details2.Offset, 0)
                    oDBs_Details2.SetValue("U_VSPCOMM", oDBs_Details2.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                Case "20"
                    objMatrix3.AddRow()
                    oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, objMatrix3.VisualRowCount)
                    oDBs_Details3.SetValue("U_VSPPTH", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPFLNM", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPDT", oDBs_Details3.Offset, "")
                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetRowEditable(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            objMatrix2 = objForm.Items.Item("19").Specific

            For i As Integer = 1 To objMatrix2.VisualRowCount
                If objMatrix2.Columns.Item("V_8").Cells.Item(i).Specific.Value = "" Then
                    objMatrix2.CommonSetting.SetRowEditable(i, True)
                Else
                    objMatrix2.CommonSetting.SetRowEditable(i, False)
                End If
            Next
            objMatrix2.Columns.Item("V_-1").Editable = False

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
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

                    If pVal.ItemUID = "16" And pVal.BeforeAction = False Then
                        objMatrix2 = objForm.Items.Item("19").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 2
                        objForm.Settings.MatrixUID = "19"
                        objMatrix2.AutoResizeColumns()
                        objForm.Freeze(False)
                    ElseIf pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                        objMatrix3 = objForm.Items.Item("20").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 3
                        objForm.Settings.MatrixUID = "20"
                        objMatrix3.AutoResizeColumns()
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "22" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objMain.objDocumentType.CreateForm(objForm.UniqueID, objForm.Items.Item("10").Specific.Value, "Tank Maintenance")
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "20" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        objMatrix3 = objForm.Items.Item("20").Specific
                        Path = pVal.Row
                        Me.BrowseFileDialog()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C1")
                    oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2")

                    objMatrix2 = objForm.Items.Item("19").Specific
                    objMatrix3 = objForm.Items.Item("20").Specific

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_VCHID" Then
                            oDBs_Head.SetValue("U_VSPVNO", oDBs_Head.Offset, oDT.GetValue("U_VSPVNO", 0))
                            oDBs_Head.SetValue("U_VSPVNM", oDBs_Head.Offset, oDT.GetValue("U_VSPVNM", 0))
                            oDBs_Head.SetValue("U_VSPODMTR", oDBs_Head.Offset, oDT.GetValue("U_VSPODRDG", 0))
                            oDBs_Head.SetValue("U_VSPVMD", oDBs_Head.Offset, oDT.GetValue("U_VSPMODEL", 0))
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("19").Specific

                    If pVal.ItemUID = "19" And pVal.ColUID = "V_1" And pVal.BeforeAction = True Then
                        If objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "Purchase Order" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 22
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "GRPO" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 20
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "A/P Invoice" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 18
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "Outgoing Payment" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 46
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "Sales Order" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 17
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "Delivery Order" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 15
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "A/R Invoice" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 13
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "Incoming Payment" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 24
                        ElseIf objMatrix1.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = "Inventory Transfer" Then
                            oLink = objMatrix1.Columns.Item("V_1").ExtendedObject
                            oLink.LinkedObjectType = 67
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Freeze(True)

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC")
                        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C1")
                        oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2")

                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C1"), "DocEntry", oDBs_Details2.GetValue("DocEntry", 0))
                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2"), "DocEntry", oDBs_Details2.GetValue("DocEntry", 0))

                        objMatrix1 = objForm.Items.Item("19").Specific
                        objMatrix2 = objForm.Items.Item("20").Specific

                        objMatrix1.LoadFromDataSource()
                        objMatrix2.LoadFromDataSource()

                        objMatrix1.AutoResizeColumns()
                        objMatrix2.AutoResizeColumns()

                        objForm.Refresh()

                        objForm.Freeze(False)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If

                        If oMenus.Exists("View") = True Then
                            objMain.objApplication.Menus.RemoveEx("View")
                        End If

                        If oMenus.Exists("Create Service Type PO") = True Then
                            objMain.objApplication.Menus.RemoveEx("Create Service Type PO")
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
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If

                        If oMenus.Exists("View") = True Then
                            objMain.objApplication.Menus.RemoveEx("View")
                        End If

                        If oMenus.Exists("Create Service Type PO") = True Then
                            objMain.objApplication.Menus.RemoveEx("Create Service Type PO")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_TNKMTNC" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TNKMTNC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix3 = objForm.Items.Item("20").Specific
                For i As Integer = 1 To objMatrix3.VisualRowCount - 1
                    If objMatrix3.IsRowSelected(i) = True Then
                        objMatrix3.DeleteRow(i)
                    End If
                Next

                For i As Integer = 1 To objMatrix3.VisualRowCount
                    objMatrix3.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            ElseIf pVal.MenuUID = "View" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_TNKMTNC_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)

                objMatrix3 = objForm.Items.Item("20").Specific

                For i As Integer = 1 To objMatrix3.VisualRowCount
                    If objMatrix3.IsRowSelected(i) = True Then
                        If objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                            System.Diagnostics.Process.Start(objMatrix3.Columns.Item("V_0").Cells.Item(i).Specific.Value)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Attachment Of TyreMaintanance"
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
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_TNKMTNC_C2")
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"
                    objMatrix3 = objForm.Items.Item("20").Specific

                    'Windows XP
                    If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                        Try
                            Dim GetPath As String = ""
                            If objMain.IsSAPHANA = True Then
                                GetPath = "Select ""U_VSPVHATC"" From OADM"
                            Else
                                GetPath = "Select U_VSPVHATC From OADM"
                            End If
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPVHATC").Value <> "" Then
                                Me.TyreMntenanceAtchmnt(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("10").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("There is no Specified Path for Vehicle Attachment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                            objForm.Refresh()

                        Catch ex As IO.IOException
                            objMain.objApplication.MessageBox(ex.Message)
                            Exit Sub
                        End Try
                        'Windows 7
                    ElseIf MyTest1.ShowDialog() = DialogResult.OK Then
                        Try
                            Dim GetPath As String = ""
                            If objMain.IsSAPHANA = True Then
                                GetPath = "Select ""U_VSPVHATC"" From OADM"
                            Else
                                GetPath = "Select U_VSPVHATC From OADM"
                            End If
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPVHATC").Value <> "" Then
                                Me.TyreMntenanceAtchmnt(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("10").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("There is no Specified Path for Driver Attachment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If

                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

                            objForm.Refresh()
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

    Sub TyreMntenanceAtchmnt(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)

                oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, Path)
                oDBs_Details3.SetValue("U_VSPPTH", oDBs_Details3.Offset, Destination)
                oDBs_Details3.SetValue("U_VSPFLNM", oDBs_Details3.Offset, FileName)
                oDBs_Details3.SetValue("U_VSPDT", oDBs_Details3.Offset, Today.ToString("yyyyMMdd"))
                objMatrix3.SetLineData(Path)

                objMatrix3.AutoResizeColumns()

                If objMatrix3.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "20")
                End If
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)
                oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, Path)
                oDBs_Details3.SetValue("U_VSPPTH", oDBs_Details3.Offset, Destination)
                oDBs_Details3.SetValue("U_VSPFLNM", oDBs_Details3.Offset, FileName)
                oDBs_Details3.SetValue("U_VSPDT", oDBs_Details3.Offset, Today.ToString("yyyyMMdd"))
                objMatrix3.SetLineData(Path)

                objMatrix3.AutoResizeColumns()

                If objMatrix3.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "20")
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        objMatrix3 = objForm.Items.Item("20").Specific
                        objMatrix2 = objForm.Items.Item("19").Specific
                        If eventInfo.ItemUID = "20" And eventInfo.ColUID = "V_-1" And objMatrix3.RowCount > 1 Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                                If oMenus.Exists("View") = False Then
                                    oCreationPackage.UniqueID = "View"
                                    oCreationPackage.String = "View"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try

                        ElseIf eventInfo.ItemUID = "20" And objMatrix3.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                If oMenus.Exists("View") = True Then
                                    objMain.objApplication.Menus.RemoveEx("View")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If

                        If eventInfo.ItemUID <> "20" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                If oMenus.Exists("View") = True Then
                                    objMain.objApplication.Menus.RemoveEx("View")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                Else
                    Try
                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                        If oMenus.Exists("View") = True Then
                            objMain.objApplication.Menus.RemoveEx("View")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)

        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objForm.Items.Item("4").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle No Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try

    End Function

End Class
