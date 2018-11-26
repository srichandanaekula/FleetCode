Imports System.Threading
Imports System.IO

Public Class clsCallibration

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
    Dim Path As String
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Callibration.xml", "VSP_FLT_CALBRTN_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_CALBRTN_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN_C0")

            Me.CellsMasking(objForm.UniqueID)

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

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN_C0")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OCALBRTN"))
            oDBs_Head.SetValue("U_VSPCCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            oDBs_Head.SetValue("U_VSPCBY", oDBs_Head.Offset, objMain.objCompany.UserName)

            objMatrix = objForm.Items.Item("35").Specific
            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("37").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("37").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("27").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("27").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("39").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("39").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        Catch ex As Exception
            objMain.objApplication.Forms.Item(FormUID)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN_C0")

            objMatrix = objForm.Items.Item("35").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_VSPTYP", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPNAM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPVAL", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPGI", oDBs_Details.Offset, "")
            objMatrix.SetLineData(objMatrix.VisualRowCount)

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

                    If pVal.ItemUID = "37" And pVal.BeforeAction = False Then

                        Try
                            Dim currentdate As String = DateAndTime.Now.ToString("yyyyMMdd")
                            Dim caldate As String = objForm.Items.Item("16").Specific.Value
                            If caldate = "" Then
                                objMain.objApplication.StatusBar.SetText("Next Calibration Date Should Not Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Try
                            End If
                            If caldate >= currentdate Then
                                If objMain.objApplication.MessageBox("Callibration is irreversible. No Changes can be done later ." & _
                                                                "Do you want to proceed?", 2, "Ok", "Cancel") = 1 Then
                                    Me.UpdatingCalibration(objForm.UniqueID)
                                End If

                            End If
                        Catch ex As Exception

                        End Try

                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.ItemUID = "35" And pVal.ColUID = "V_3" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("35").Specific
                        Path = pVal.Row
                        Me.BrowseFileDialog()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN_C0")
                    objMatrix = objForm.Items.Item("35").Specific
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If pVal.BeforeAction = True Then
                        If oCFL.UniqueID = "CFL_VNO" Then
                            Me.CFLFilterVechicleNo(objForm.UniqueID, oCFL.UniqueID)
                        End If
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_VNO" Then
                            oDBs_Head.SetValue("U_VSPVCHID", oDBs_Head.Offset, oDT.GetValue("U_VSPVNO", 0))
                            oDBs_Head.SetValue("U_VSPVHNME", oDBs_Head.Offset, oDT.GetValue("U_VSPVNM", 0))
                            oDBs_Head.SetValue("U_VSPVINM", oDBs_Head.Offset, oDT.GetValue("U_VSPCNTNM", 0))
                            oDBs_Head.SetValue("U_VSPVINCG", oDBs_Head.Offset, oDT.GetValue("U_VSPCNTR", 0))

                            Dim GetLastClbDt As String = ""
                            'If objMain.IsSAPHANA = True Then
                            '    GetLastClbDt = "Select ""U_VSPCCDT"" From ""@VSP_FLT_CALBRTN"" Where ""DocNum"" = (Select MAX (""DocEntry"") From ""@VSP_FLT_CALBRTN"") "
                            'Else
                            '    GetLastClbDt = "Select U_VSPCCDT From [@VSP_FLT_CALBRTN] Where DocNum = (Select MAX (DocEntry) From [@VSP_FLT_CALBRTN]) "
                            'End If

                            If objMain.IsSAPHANA = True Then
                                GetLastClbDt = "Select ""U_VSPCCDT"" From ""@VSP_FLT_CALBRTN"" Where ""DocNum"" = (Select MAX (""DocNum"") From ""@VSP_FLT_CALBRTN"") "
                            Else
                                GetLastClbDt = "Select U_VSPCCDT From [@VSP_FLT_CALBRTN] Where DocNum = (Select MAX (DocNum) From [@VSP_FLT_CALBRTN]) "
                            End If

                            Dim oRsGetLastClbDt As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetLastClbDt.DoQuery(GetLastClbDt)

                            If oRsGetLastClbDt.RecordCount > 0 Then
                                Dim sDate As Date = oRsGetLastClbDt.Fields.Item(0).Value
                                oDBs_Head.SetValue("U_VSPLCDT", oDBs_Head.Offset, sDate.ToString("yyyyMMdd"))
                            End If
                        End If

                        If oCFL.UniqueID = "CFL_TYP" Then
                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPTYP", oDBs_Details.Offset, oDT.GetValue("U_VSPTYP", 0))
                            oDBs_Details.SetValue("U_VSPNAM", oDBs_Details.Offset, oDT.GetValue("U_VSPNAM", 0))
                            oDBs_Details.SetValue("U_VSPVAL", oDBs_Details.Offset, "")
                            oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, "")
                            objMatrix.SetLineData(pVal.Row)

                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Issue Component") = True Then
                            objMain.objApplication.Menus.RemoveEx("Issue Component")
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
                        If oMenus.Exists("Issue Component") = True Then
                            objMain.objApplication.Menus.RemoveEx("Issue Component")
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
            If pVal.MenuUID = "VSP_FLT_CALBRTN" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Issue Component" And pVal.BeforeAction = False Then
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    For i As Integer = 1 To objMatrix.VisualRowCount
                        If objMatrix.IsRowSelected(i) = True And objMatrix.Columns.Item("V_4").Cells.Item(i).Specific.Value = "" Then
                            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_CALBRTN_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            objMain.objApplication.ActivateMenuItem("3079")
                            Dim UDOForm As SAPbouiCOM.Form
                            UDOForm = objMain.objApplication.Forms.GetForm("-720", objMain.objApplication.Forms.ActiveForm.TypeCount)
                            UDOForm.Items.Item("U_VSPVHNO").Specific.Value = objForm.Items.Item("101").Specific.Value
                            UDOForm.Items.Item("U_VSPLNUM").Specific.Value = i
                            UDOForm.Items.Item("U_VSPDCNO").Specific.Value = objForm.Items.Item("25").Specific.Value
                            UDOForm.Items.Item("U_VSPDCTYP").Specific.Value = "Calibration"
                        End If
                        Exit For
                    Next
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "CallibrationAttachments"

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
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN_C0")
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"
                    objMatrix = objForm.Items.Item("35").Specific

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
                                Me.CallibrationAtchmnt(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("25").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("There is no Specified Path for Driver Attachment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                                Me.CallibrationAtchmnt(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("25").Specific.Value)
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

    Sub CallibrationAtchmnt(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, Path)
                oDBs_Details.SetValue("U_VSPTYP", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(CInt(Path)).Specific.Value)
                oDBs_Details.SetValue("U_VSPNAM", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(CInt(Path)).Specific.Value)
                oDBs_Details.SetValue("U_VSPVAL", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(CInt(Path)).Specific.Value)
                oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, Destination)
                objMatrix.SetLineData(Path)

                objMatrix.AutoResizeColumns()

                If objMatrix.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID)
                End If
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, Path)
                oDBs_Details.SetValue("U_VSPTYP", oDBs_Details.Offset, objMatrix.Columns.Item("V_0").Cells.Item(CInt(Path)).Specific.Value)
                oDBs_Details.SetValue("U_VSPNAM", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(CInt(Path)).Specific.Value)
                oDBs_Details.SetValue("U_VSPVAL", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(CInt(Path)).Specific.Value)
                oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, Destination)

                objMatrix.SetLineData(Path)

                objMatrix.AutoResizeColumns()

                If objMatrix.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID)
                End If

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

    Sub UpdatingCalibration(ByVal FormUID As String)
        Try
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_CALBRTN")
            oDBs_Head.SetValue("U_VSPAPBY", oDBs_Head.Offset, objMain.objCompany.UserName)

            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            Dim GetCode As String = ""
            If objMain.IsSAPHANA = True Then
                GetCode = "Select ""Code"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" = '" & objForm.Items.Item("101").Specific.Value & "'"
            Else
                GetCode = "Select Code From [@VSP_FLT_VMSTR] Where U_VSPVNO = '" & objForm.Items.Item("101").Specific.Value & "'"
            End If

            Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCode.DoQuery(GetCode)

            objMain.sCmp = objMain.objCompany.GetCompanyService
            objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_OVMSTR")
            objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            objMain.oGeneralParams.SetProperty("Code", oRsGetCode.Fields.Item(0).Value)
            objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
            objMain.oGeneralData.SetProperty("U_VSPCALB", "Y")
            objMain.oGeneralService.Update(objMain.oGeneralData)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

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
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objMatrix = objForm.Items.Item("35").Specific
                        If eventInfo.ItemUID = "35" And eventInfo.ColUID = "V_-1" Then
                            Try
                                If objMatrix.Columns.Item("V_0").Cells.Item(1).Specific.Value <> "" Then
                                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                    oMenus = oMenuItem.SubMenus
                                    If oMenus.Exists("Issue Component") = False Then
                                        oCreationPackage.UniqueID = "Issue Component"
                                        oCreationPackage.String = "Issue Component"
                                        oCreationPackage.Enabled = True
                                        oMenus.AddEx(oCreationPackage)
                                    End If
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "35" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Issue Component") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Issue Component")
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "35" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Issue Component") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Issue Component")
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
                        If oMenus.Exists("Issue Component") = True Then
                            objMain.objApplication.Menus.RemoveEx("Issue Component")
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

    Sub CFLFilterVechicleNo(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "U_VSPCALB"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
            If oConditions.Count > 0 Then oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCondition = oConditions.Add()
            oCondition.Alias = "U_VSPAVLB"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = "Y"
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)


            If objForm.Items.Item("101").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle Id Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("16").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Next Calibration Date Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("29").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Due Date Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False

            End If



            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
End Class
