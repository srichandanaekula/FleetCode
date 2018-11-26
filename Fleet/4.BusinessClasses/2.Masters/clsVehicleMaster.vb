Imports System.Threading
Imports System.IO

Public Class clsVehicleMaster

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details, oDBs_Details1, oDBs_Details2, oDBs_Details5, oDBs_Details6, oDBs_Details7 As SAPbouiCOM.DBDataSource
    Dim objMatrix, objMatrix1, objMatrix2, objMatrix5, objMatrix6, objMatrix7 As SAPbouiCOM.Matrix
    Dim objPicture As SAPbouiCOM.PictureBox
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim Path As String
    Dim iPath As Integer
    Dim oButton, oButton1 As SAPbouiCOM.Button
    Dim oPVGrid, oACGrid As SAPbouiCOM.Grid
    Dim oDT, oDT1 As SAPbouiCOM.DataTable
    Dim oColumn As SAPbouiCOM.Column
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Vehicle Master.xml", "VSP_FLT_VMSTR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C2")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C5")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C6")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C7")

            'objForm.Items.Item("3").TextStyle = 4
            objForm.Items.Item("26").TextStyle = 4
            objForm.Items.Item("39").TextStyle = 4
            objForm.Items.Item("50").TextStyle = 4
            objForm.Items.Item("59").TextStyle = 4
            objForm.Items.Item("70").TextStyle = 4

            Me.CellsMasking(objForm.UniqueID)

            objForm.Items.Item("109").Visible = False
            objForm.Items.Item("111").Visible = False
            objForm.Items.Item("113").Visible = False
            objForm.Items.Item("115").Visible = False
            objForm.Items.Item("117").Visible = False
            objForm.Items.Item("119").Visible = False
            objForm.Items.Item("121").Visible = False
            objForm.Items.Item("123").Visible = False
            
            oButton = objForm.Items.Item("109").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("111").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("113").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("115").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("117").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("119").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("121").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("123").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            Me.CFLFilterContractor(objForm.UniqueID, "CFL_CNTR")
            Me.CFLFilterDieselItems(objForm.UniqueID, "CFL_DIITEM")
            Me.CFLFilterCostCenter(objForm.UniqueID, "CFL_VEHCC", "2")

            'objComboBox = objForm.Items.Item("154").Specific
            'objComboBox.ValidValues.Add("", "")
            'objComboBox.ValidValues.Add("Active", "Active")
            'objComboBox.ValidValues.Add("Inactive", "Inactive")
            'objComboBox.ValidValues.Add("Under Periodic Check", "Under Periodic Check")

            oDT = objForm.DataSources.DataTables.Add("dt1")
            oDT = objForm.DataSources.DataTables.Item("dt1")

            oDT1 = objForm.DataSources.DataTables.Add("dt2")
            oDT1 = objForm.DataSources.DataTables.Item("dt2")

            objForm.PaneLevel = 1

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMatrix = objForm.Items.Item("91").Specific
            objMatrix1 = objForm.Items.Item("92").Specific
            objMatrix2 = objForm.Items.Item("1000003").Specific
            objMatrix5 = objForm.Items.Item("125").Specific
            objMatrix6 = objForm.Items.Item("142").Specific
            objMatrix7 = objForm.Items.Item("150").Specific

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C2")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C5")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C6")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C7")

            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSP_FLT_VMSTR"))
            oDBs_Head.SetValue("U_VSPAVLB", oDBs_Head.Offset, "Y")
            oDBs_Head.SetValue("U_VSPCALB", oDBs_Head.Offset, "N")

            oDBs_Head.SetValue("U_VSPCHK", oDBs_Head.Offset, "Y")
            oDBs_Head.SetValue("U_VSPUNPCK", oDBs_Head.Offset, "N")

            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            SetNewLine(objForm.UniqueID, "91")

            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            SetNewLine(objForm.UniqueID, "92")

            objMatrix2.Clear()
            oDBs_Details2.Clear()
            objMatrix2.FlushToDataSource()
            SetNewLine(objForm.UniqueID, "1000003")

            objMatrix5.Clear()
            oDBs_Details5.Clear()
            objMatrix5.FlushToDataSource()
            SetNewLine(objForm.UniqueID, "125")

            objMatrix6.Clear()
            oDBs_Details6.Clear()
            objMatrix6.FlushToDataSource()
            SetNewLine(objForm.UniqueID, "142")

            objMatrix7.Clear()
            oDBs_Details7.Clear()
            objMatrix7.FlushToDataSource()
            SetNewLine(objForm.UniqueID, "150")
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("21").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("17").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("127").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("127").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("102").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("102").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("132").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("132").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("25").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("7").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("41").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("41").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("152").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            ' objForm.Items.Item("152").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("152").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("153").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("153").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("91").Specific
            objMatrix1 = objForm.Items.Item("92").Specific
            objMatrix2 = objForm.Items.Item("1000003").Specific
            objMatrix5 = objForm.Items.Item("125").Specific
            objMatrix6 = objForm.Items.Item("142").Specific
            objMatrix7 = objForm.Items.Item("150").Specific

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C0")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C1")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C2")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C5")
            oDBs_Details6 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C6")
            oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C7")

            Select Case MatrixUID
                Case "91"
                    objMatrix.AddRow()
                    oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
                    oDBs_Details.SetValue("U_VSPTYPE", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPNAME", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPNUM", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPISSDT", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPEXPDT", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPAMT", oDBs_Details.Offset, "")
                    oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, "")
                    objMatrix.SetLineData(objMatrix.VisualRowCount)
                Case "92"
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_VSPTYPE", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPCMPY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPLNNO", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPSTRDT", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPENDDT", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPAMT", oDBs_Details1.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                Case "1000003"
                    objMatrix2.AddRow()
                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Details2.SetValue("U_VSPPART", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDEALR", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDTPUR", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPPRC", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPWEDT", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPINSRV", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPOTSRV", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPTNSDT", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDTSLD", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPSLDTO", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPCMTS", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPNFKM", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPNFDYS", oDBs_Details2.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                Case "125"
                    objMatrix5.AddRow()
                    oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, objMatrix5.VisualRowCount)
                    oDBs_Details5.SetValue("U_VSPFRMDT", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, "99991231")
                    oDBs_Details5.SetValue("U_VSPCNTR", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, "")
                    objMatrix5.SetLineData(objMatrix5.VisualRowCount)
                Case "142"
                    objMatrix6.AddRow()
                    oDBs_Details6.SetValue("LineId", oDBs_Details6.Offset, objMatrix6.VisualRowCount)
                    oDBs_Details6.SetValue("U_VSPCMPNY", oDBs_Details6.Offset, "")
                    oDBs_Details6.SetValue("U_VSPINSNO", oDBs_Details6.Offset, "")
                    oDBs_Details6.SetValue("U_VSPSTDT", oDBs_Details6.Offset, "")
                    oDBs_Details6.SetValue("U_VSPENDDT", oDBs_Details6.Offset, "")
                    oDBs_Details6.SetValue("U_VSPAMT", oDBs_Details6.Offset, "")
                    oDBs_Details6.SetValue("U_VSPVAMT", oDBs_Details6.Offset, "")
                    objMatrix6.SetLineData(objMatrix6.VisualRowCount)
                Case "150"
                    objMatrix7.AddRow()
                    oDBs_Details7.SetValue("LineId", oDBs_Details7.Offset, objMatrix7.VisualRowCount)
                    oDBs_Details7.SetValue("U_VSPFRMDT", oDBs_Details7.Offset, "")
                    oDBs_Details7.SetValue("U_VSPTODT", oDBs_Details7.Offset, "")
                    oDBs_Details7.SetValue("U_VSPRDNG", oDBs_Details7.Offset, "")
                    oDBs_Details7.SetValue("U_VSPCLSRD", oDBs_Details7.Offset, "")
                    oDBs_Details7.SetValue("U_VSPSRLNO", oDBs_Details7.Offset, "")
                    objMatrix7.SetLineData(objMatrix7.VisualRowCount)
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_VMSTR" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("91").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next

                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            ElseIf pVal.MenuUID = "View" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("91").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        If objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value <> "" Then
                            System.Diagnostics.Process.Start(objMatrix.Columns.Item("V_0").Cells.Item(i).Specific.Value)
                        End If
                    End If
                Next
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
                    objMatrix = objForm.Items.Item("91").Specific
                    objMatrix1 = objForm.Items.Item("92").Specific
                    objMatrix2 = objForm.Items.Item("1000003").Specific
                    objMatrix5 = objForm.Items.Item("125").Specific
                    objMatrix6 = objForm.Items.Item("142").Specific
                    objMatrix7 = objForm.Items.Item("150").Specific
                    oPVGrid = objForm.Items.Item("1000005").Specific
                    oACGrid = objForm.Items.Item("1000007").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or _
                                                                           pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1000006" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 1
                        objForm.Settings.MatrixUID = "125"
                        objMatrix5.AutoResizeColumns()
                    ElseIf pVal.ItemUID = "85" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 2
                        objForm.Settings.MatrixUID = "91"
                        objMatrix.AutoResizeColumns()
                    ElseIf pVal.ItemUID = "86" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 3
                        objForm.Settings.MatrixUID = "92"
                        objMatrix1.AutoResizeColumns()
                    ElseIf pVal.ItemUID = "87" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 4
                        objForm.Settings.MatrixUID = "1000003"
                        objMatrix2.AutoResizeColumns()
                    ElseIf pVal.ItemUID = "88" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 5
                        objForm.Settings.MatrixUID = "1000005"
                    ElseIf pVal.ItemUID = "89" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 6
                        objForm.Settings.MatrixUID = "1000007"
                    ElseIf pVal.ItemUID = "90" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 7
                    ElseIf pVal.ItemUID = "141" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 8
                    ElseIf pVal.ItemUID = "149" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 9
                        objForm.Settings.MatrixUID = "150"
                        objMatrix7.AutoResizeColumns()
                    End If

                    If pVal.ItemUID = "90" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        objForm.Items.Item("109").Visible = True
                        objForm.Items.Item("111").Visible = True
                        objForm.Items.Item("113").Visible = True
                        objForm.Items.Item("115").Visible = True
                        objForm.Items.Item("117").Visible = True
                        objForm.Items.Item("119").Visible = True
                        objForm.Items.Item("121").Visible = True
                        objForm.Items.Item("123").Visible = True
                    End If

                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "110" Then
                            objPicture = objForm.Items.Item("94").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("110").Visible = False
                        ElseIf pVal.ItemUID = "112" Then
                            objPicture = objForm.Items.Item("95").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("112").Visible = False
                        ElseIf pVal.ItemUID = "114" Then
                            objPicture = objForm.Items.Item("96").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("114").Visible = False
                        ElseIf pVal.ItemUID = "116" Then
                            objPicture = objForm.Items.Item("101").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("116").Visible = False
                        ElseIf pVal.ItemUID = "118" Then
                            objPicture = objForm.Items.Item("98").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("118").Visible = False
                        ElseIf pVal.ItemUID = "120" Then
                            objPicture = objForm.Items.Item("97").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("120").Visible = False
                        ElseIf pVal.ItemUID = "122" Then
                            objPicture = objForm.Items.Item("100").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("122").Visible = False
                        ElseIf pVal.ItemUID = "124" Then
                            objPicture = objForm.Items.Item("99").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("124").Visible = False
                        End If
                    End If

                    If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "109" Then
                            objPicture = objForm.Items.Item("94").Specific
                            oButton1 = objForm.Items.Item("110").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "111" Then
                            objPicture = objForm.Items.Item("95").Specific
                            oButton1 = objForm.Items.Item("112").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "113" Then
                            objPicture = objForm.Items.Item("96").Specific
                            oButton1 = objForm.Items.Item("114").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "115" Then
                            objPicture = objForm.Items.Item("101").Specific
                            oButton1 = objForm.Items.Item("116").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "117" Then
                            objPicture = objForm.Items.Item("98").Specific
                            oButton1 = objForm.Items.Item("118").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "119" Then
                            objPicture = objForm.Items.Item("97").Specific
                            oButton1 = objForm.Items.Item("120").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "121" Then
                            objPicture = objForm.Items.Item("100").Specific
                            oButton1 = objForm.Items.Item("122").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "123" Then
                            objPicture = objForm.Items.Item("99").Specific
                            oButton1 = objForm.Items.Item("124").Specific
                            Me.PictureBrowseFileDialouge()
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix5 = objForm.Items.Item("125").Specific
                    If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "5" Then
                            Dim ChkVehicleNoExists As String = ""
                            If objMain.IsSAPHANA = True Then
                                ChkVehicleNoExists = "Select ""Code"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPVNO"" ='" & objForm.Items.Item("5").Specific.Value.Trim & "'"
                            Else
                                ChkVehicleNoExists = "Select Code From [@VSP_FLT_VMSTR] Where [U_VSPVNO] ='" & objForm.Items.Item("5").Specific.Value.Trim & "'"
                            End If

                            Dim oRsChkVehicleNoExists As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkVehicleNoExists.DoQuery(ChkVehicleNoExists)

                            If oRsChkVehicleNoExists.RecordCount > 0 Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("21").Specific.value = oRsChkVehicleNoExists.Fields.Item(0).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If
                        If pVal.ItemUID = "138" Then
                            Dim ChkChasisNoExists As String = ""

                            If objMain.IsSAPHANA = True Then
                                ChkChasisNoExists = "Select ""Code"" From ""@VSP_FLT_VMSTR"" Where ""U_VSPCHSNO"" ='" & objForm.Items.Item("138").Specific.Value.Trim & "'"
                            Else
                                ChkChasisNoExists = "Select Code From [@VSP_FLT_VMSTR] Where [U_VSPCHSNO] ='" & objForm.Items.Item("138").Specific.Value.Trim & "'"

                            End If
                            Dim oRsChkChasisNoExists As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsChkChasisNoExists.DoQuery(ChkChasisNoExists)
                            If oRsChkChasisNoExists.RecordCount > 0 Then
                                objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                objForm.Items.Item("21").Specific.value = oRsChkChasisNoExists.Fields.Item(0).Value
                                objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If
                        If pVal.ItemUID = "125" And pVal.ColUID = "V_3" Then
                            If objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value <> "" And pVal.Row <> "1" Then
                                Dim BeforeDate As String = objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row - 1).Specific.Value
                                Dim PresentDate As String = objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value

                                Dim GetDate As String = ""

                                If objMain.IsSAPHANA = True Then
                                    GetDate = "Select  DAYS_BETWEEN (TO_DATE('" & BeforeDate & "'),TO_DATE('" & PresentDate & "')), " & _
                                                                 "MONTHS_BETWEEN (To_DATE('" & BeforeDate & "'),TO_DATE('" & PresentDate & "')),YEARS_BETWEEN (TO_DATE('" & BeforeDate & "'),TO_DATE('" & PresentDate & "')) From DUMMY"
                                Else
                                    GetDate = "Select  DATEDIFF (D,'" & BeforeDate & "','" & PresentDate & "'), " & _
                               "DATEDIFF (M,'" & BeforeDate & "','" & PresentDate & "'),DATEDIFF (Y,'" & BeforeDate & "','" & PresentDate & "')"
                                End If
                                Dim oRsGetDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetDate.DoQuery(GetDate)

                                If pVal.Row > 1 Then
                                    If oRsGetDate.Fields.Item(0).Value > 0 And oRsGetDate.Fields.Item(1).Value >= 0 And oRsGetDate.Fields.Item(2).Value >= 0 Then
                                        oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, pVal.Row)
                                        oDBs_Details5.SetValue("U_VSPFRMDT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details5.SetValue("U_VSPCNTR", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                        objMatrix5.SetLineData(pVal.Row)
                                    Else
                                        objMain.objApplication.StatusBar.SetText("Date should not be less than Preceeding Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = ""
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                End If
                            End If
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "91" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        objMatrix = objForm.Items.Item("91").Specific
                        Path = pVal.Row
                        iPath = pVal.Row
                        Me.BrowseFileDialog()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    objMatrix5 = objForm.Items.Item("125").Specific
                    oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C5")

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

                        If oCFL.UniqueID = "CFL_CNTR" Then

                            oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, pVal.Row)
                            oDBs_Details5.SetValue("U_VSPFRMDT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPCNTR", oDBs_Details5.Offset, oDT.GetValue("empID", 0))

                            Dim GetContractors As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetContractors = "Select ""firstName"",""lastName"" From OHEM Where ""empID"" = '" & oDT.GetValue("empID", 0) & "'"

                            Else
                                GetContractors = "Select firstName,lastName From OHEM Where empID = '" & oDT.GetValue("empID", 0) & "'"

                            End If
                            Dim oRsGetContractors As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetContractors.DoQuery(GetContractors)

                            oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, oRsGetContractors.Fields.Item("firstName").Value & " " & oRsGetContractors.Fields.Item("lastName").Value)
                            objMatrix5.SetLineData(pVal.Row)

                            objMatrix5.AutoResizeColumns()
                            If objMatrix5.VisualRowCount = pVal.Row Then
                                Me.SetNewLine(objForm.UniqueID, "125")
                            End If

                        End If

                        If oCFL.UniqueID = "CFL_DIITEM" Then
                            oDBs_Head.SetValue("U_VSPDIITE", oDBs_Head.Offset, oDT.GetValue("ItemCode", 0))
                        End If

                        If oCFL.UniqueID = "CFL_VEHCC" Then
                            oDBs_Head.SetValue("U_VSPVEHCC", oDBs_Head.Offset, oDT.GetValue("PrcCode", 0))
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("92").Specific
                    objMatrix2 = objForm.Items.Item("1000003").Specific
                    objMatrix5 = objForm.Items.Item("125").Specific
                    objMatrix6 = objForm.Items.Item("142").Specific
                    objMatrix7 = objForm.Items.Item("150").Specific
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C2")
                    oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C5")
                    oDBs_Details7 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C7")

                    If pVal.ItemUID = "150" And pVal.ColUID = "V_3" And pVal.BeforeAction = False Then
                        If objMatrix7.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix7.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "150")
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "92" And pVal.ColUID = "V_4" And pVal.BeforeAction = False Then
                        If objMatrix1.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "92")
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "142" And pVal.ColUID = "V_4" And pVal.BeforeAction = False Then
                        If objMatrix6.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix6.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "142")
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "1000003" And pVal.ColUID = "V_9" And pVal.BeforeAction = False Then
                        If objMatrix2.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix2.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "1000003")
                            End If
                        End If
                    End If

                    If pVal.ItemUID = "125" And pVal.ColUID = "V_3" And pVal.BeforeAction = False Then

                        If pVal.Row > 1 And objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value <> "" Then

                            Dim GetDate As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetDate = "Select ADD_DAYS(" & _
                           "CAST(REPLACE('" & objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "', '-', '')as DATETIME ), -1) From DUMMY "
                            Else
                                GetDate = "Select  DATEADD(DAY, -1, " & _
                           "CAST(REPLACE('" & objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value & "', '-', '') AS DATETIME)) "
                            End If
                            Dim oRsGetDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetDate.DoQuery(GetDate)

                            Dim SDate As Date = oRsGetDate.Fields.Item(0).Value

                            oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, pVal.Row - 1)
                            oDBs_Details5.SetValue("U_VSPFRMDT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row - 1).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, SDate.ToString("yyyyMMdd"))
                            oDBs_Details5.SetValue("U_VSPCNTR", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row - 1).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row - 1).Specific.Value)
                            objMatrix5.SetLineData(pVal.Row - 1)
                        End If
                    End If

                    If pVal.ItemUID = "1000003" And pVal.BeforeAction = False Then
                        If pVal.ColUID = "V_11" Then
                            If objMatrix2.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value <> "" Then

                                Dim DatePurchased As String = objMatrix2.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value
                                DatePurchased = DatePurchased.Insert(4, "-")
                                DatePurchased = DatePurchased.Insert(7, "-")

                                Dim GetExpDate As String = ""

                                If objMain.IsSAPHANA = True Then
                                    GetExpDate = "Select ADD_DAYS('" & DatePurchased & "','" & objMatrix2.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value & "') as ""ExpiryDate"" From ""@VSP_FLT_VMSTR_C2"""
                                Else
                                    GetExpDate = "Select DATEADD(D," & objMatrix2.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value & ",'" & DatePurchased & "') as 'ExpiryDate' From [@VSP_FLT_VMSTR_C2]"
                                End If

                                Dim oRsGetExpDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetExpDate.DoQuery(GetExpDate)

                                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, pVal.Row)
                                oDBs_Details2.SetValue("U_VSPPART", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPDEALR", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPDTPUR", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPPRC", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                                Dim ExpDate As Date = oRsGetExpDate.Fields.Item("ExpiryDate").Value
                                oDBs_Details2.SetValue("U_VSPWEDT", oDBs_Details2.Offset, ExpDate.ToString("yyyyMMdd"))
                                oDBs_Details2.SetValue("U_VSPINSRV", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPOTSRV", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPTNSDT", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPDTSLD", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPSLDTO", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPCMTS", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPNFKM", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPNFDYS", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_11").Cells.Item(pVal.Row).Specific.Value)
                                objMatrix2.SetLineData(pVal.Row)
                            End If
                        End If

                        If pVal.ColUID = "V_6" Then
                            If objMatrix2.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value <> "" Then

                                Dim DatePurchased As String = objMatrix2.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value
                                DatePurchased = DatePurchased.Insert(4, "-")
                                DatePurchased = DatePurchased.Insert(7, "-")

                                Dim ExpDate As String = objMatrix2.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value
                                ExpDate = ExpDate.Insert(4, "-")
                                ExpDate = ExpDate.Insert(7, "-")

                                Dim GetNoofDays As String = ""

                                If objMain.IsSAPHANA = True Then

                                    GetNoofDays = "Select DAYS_BETWEEN (TO_DATE('" & DatePurchased & "'),TO_DATE('" & ExpDate & "')) as ""NoofDays"" From ""@VSP_FLT_VMSTR_C2"" "
                                Else
                                    GetNoofDays = "Select DATEDIFF (D,'" & DatePurchased & "','" & ExpDate & "') as 'NoofDays' From [@VSP_FLT_VMSTR_C2] "
                                End If
                                Dim oRsGetNoofDays As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetNoofDays.DoQuery(GetNoofDays)

                                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, pVal.Row)
                                oDBs_Details2.SetValue("U_VSPPART", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_10").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPDEALR", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_9").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPDTPUR", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_8").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPPRC", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_7").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPWEDT", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_6").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPINSRV", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPOTSRV", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPTNSDT", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPDTSLD", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPSLDTO", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPCMTS", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPNFKM", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_12").Cells.Item(pVal.Row).Specific.Value)
                                oDBs_Details2.SetValue("U_VSPNFDYS", oDBs_Details2.Offset, oRsGetNoofDays.Fields.Item("NoofDays").Value)
                                objMatrix2.SetLineData(pVal.Row)
                            End If
                        End If
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
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CalculateTareWeight(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim CabinWeight As Double = 0
            Dim ChasisWeight As Double = 0
            Dim TankWeight As Double = 0
            Dim OtherFittings As Double = 0
            If objForm.Items.Item("1000011").Specific.Value <> "" Then
                CabinWeight = objForm.Items.Item("1000011").Specific.Value
            End If
            If objForm.Items.Item("144").Specific.Value <> "" Then
                ChasisWeight = objForm.Items.Item("144").Specific.Value
            End If
            If objForm.Items.Item("146").Specific.Value <> "" Then
                TankWeight = objForm.Items.Item("146").Specific.Value
            End If
            If objForm.Items.Item("148").Specific.Value <> "" Then
                OtherFittings = objForm.Items.Item("148").Specific.Value
            End If

            Dim TareWeight As Double = CabinWeight + ChasisWeight + TankWeight + OtherFittings
            oDBs_Head.SetValue("U_VSPGWGT", oDBs_Head.Offset, TareWeight)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterContractor(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "dept"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            Dim GetCntr As String = ""

            If objMain.IsSAPHANA = True Then
                GetCntr = "Select ""U_VSPCNT""  From ""@VSP_FLT_CNFGSRN"""
            Else
                GetCntr = "Select U_VSPCNT  From [@VSP_FLT_CNFGSRN]"
            End If
            Dim oRsGetCntr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCntr.DoQuery(GetCntr)
            oCondition.CondVal = oRsGetCntr.Fields.Item("U_VSPCNT").Value
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterDieselItems(ByVal FormUID As String, ByVal CFL_ID As String)
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
            oCondition.Alias = "ItmsGrpCod"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            Dim GetCntr As String = ""
            If objMain.IsSAPHANA = True Then
                GetCntr = "Select ""U_VSPDSLGR""  From ""@VSP_FLT_CNFGSRN"""
            Else
                GetCntr = "Select U_VSPDSLGR  From [@VSP_FLT_CNFGSRN] "
            End If
            Dim oRsGetCntr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCntr.DoQuery(GetCntr)
            oCondition.CondVal = oRsGetCntr.Fields.Item("U_VSPDSLGR").Value
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilterCostCenter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal DimCode As String)
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

    Sub VehicleAttachment(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("91").Specific

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = DestinationPath & "\" & FileName
                File.Copy(AttchPath, Destination)
                objPicture.Picture = Destination
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)
                objPicture.Picture = Destination
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region "    Picture Attachment    "

    Sub PictureBrowseFileDialouge()
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowserPicture)
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

    Sub ShowFolderBrowserPicture()
        Dim MyTest1 As New OpenFileDialog
        Dim MyProcs() As Process
        Try
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C0")
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To 0 'MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"

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
                            Dim Attachment As String = MyTest1.FileName
                            If oRsGetPath.Fields.Item("U_VSPVHATC").Value <> "" Then
                                Me.VehicleAttachment(objForm.UniqueID, Path, Attachment, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("5").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("There is no Specified Path for Vehicle Attachment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If

                            objForm.Refresh()
                            oButton1.Image = Application.StartupPath & "\RemoveImage.jpg"
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
                            Dim Attachment As String = MyTest1.FileName
                            If oRsGetPath.Fields.Item("U_VSPVHATC").Value <> "" Then
                                Me.VehicleAttachment(objForm.UniqueID, Path, Attachment, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("5").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("There is no Specified Path for Vehicle Attachment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End If

                            objForm.Refresh()
                            oButton1.Image = Application.StartupPath & "\RemoveImage.jpg"
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

#Region " Vechical Attachment"

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

            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To 0 'MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"
                    objMatrix = objForm.Items.Item("91").Specific

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
                                Me.AttachmentVechicals(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("21").Specific.Value)
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
                                Me.AttachmentVechicals(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("21").Specific.Value)
                            Else
                                objMain.objApplication.StatusBar.SetText("There is no Specified Path for Vehicle Attachment", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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

    Sub AttachmentVechicals(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix = objForm.Items.Item("91").Specific
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR_C0")
            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)

                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, Path)
                oDBs_Details.SetValue("U_VSPTYPE", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPNAME", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPNUM", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPISSDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPEXPDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, Destination)
                objMatrix.SetLineData(Path)

                objMatrix.AutoResizeColumns()
                If objMatrix.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "91")
                End If
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)
                oDBs_Details.SetValue("LineId", oDBs_Details.Offset, Path)
                oDBs_Details.SetValue("U_VSPTYPE", oDBs_Details.Offset, objMatrix.Columns.Item("V_5").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPNAME", oDBs_Details.Offset, objMatrix.Columns.Item("V_4").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPNUM", oDBs_Details.Offset, objMatrix.Columns.Item("V_3").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPISSDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_2").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPEXPDT", oDBs_Details.Offset, objMatrix.Columns.Item("V_1").Cells.Item(iPath).Specific.Value)
                oDBs_Details.SetValue("U_VSPATTCH", oDBs_Details.Offset, Destination)
                objMatrix.SetLineData(Path)

                objMatrix.AutoResizeColumns()
                If objMatrix.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "91")
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
                        objMatrix = objForm.Items.Item("91").Specific

                        If eventInfo.ItemUID = "91" And eventInfo.ColUID = "V_-1" And objMatrix.RowCount > 1 Then
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
                        ElseIf eventInfo.ItemUID = "91" And objMatrix.RowCount <= 1 Then
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
                        If eventInfo.ItemUID <> "91" Then
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

#Region " FormDataEvent"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_VMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_VMSTR")
            objMatrix5 = objForm.Items.Item("125").Specific
            objMatrix7 = objForm.Items.Item("150").Specific ''Added on 12-09-2018 
            objMatrix = objForm.Items.Item("91").Specific
            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                    If BusinessObjectInfo.BeforeAction = True Then
                        If objMatrix5.VisualRowCount <> 1 Then
                            Try
                                oDBs_Head.SetValue("U_VSPCNTR", oDBs_Head.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)
                                oDBs_Head.SetValue("U_VSPCNTNM", oDBs_Head.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)

                                ''Added on 12-09-2018 by Abinas (for solve issue like invalid row index)
                                If objMatrix7.VisualRowCount > 1 Then
                                    oDBs_Head.SetValue("U_VSPODRDG", oDBs_Head.Offset, objMatrix7.Columns.Item("V_1").Cells.Item(objMatrix7.VisualRowCount - 1).Specific.Value)
                                End If
                                'oDBs_Head.SetValue("U_VSPODRDG", oDBs_Head.Offset, objMatrix7.Columns.Item("V_1").Cells.Item(objMatrix7.VisualRowCount - 1).Specific.Value)
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If

                        Try
                            Me.CalculateTareWeight(objForm.UniqueID)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If


                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    If BusinessObjectInfo.BeforeAction = True Then
                        Try
                            oDBs_Head.SetValue("U_VSPCNTR", oDBs_Head.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)
                            oDBs_Head.SetValue("U_VSPCNTNM", oDBs_Head.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)

                            ''Added on 12-09-2018 by Abinas(for solve issue like invalid row index)
                            If objMatrix7.VisualRowCount > 1 Then
                                oDBs_Head.SetValue("U_VSPODRDG", oDBs_Head.Offset, objMatrix7.Columns.Item("V_1").Cells.Item(objMatrix7.VisualRowCount - 1).Specific.Value)
                            End If

                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                        Try
                            Me.CalculateTareWeight(objForm.UniqueID)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.LoadGrid(objForm.UniqueID)

                        If objForm.Items.Item("5").Specific.Value = "" Then
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        Else
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If
                        objForm.Items.Item("152").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Private Sub LoadGrid(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oPVGrid = objForm.Items.Item("1000005").Specific
            oACGrid = objForm.Items.Item("1000007").Specific

            oPVGrid.DataTable = oDT
            Dim Query As String = ""
            If objMain.IsSAPHANA = True Then
                Query = "Select ""U_VSPREMNM"" as ""Reminder Name"" , ""U_VSPRMFR"" as ""For"" , ""U_VSPTDESC"" as ""Task Description"" , ""U_VSPCDT"" as ""Raised Date"" , " & _
            """U_VSPRBON"" as ""Based On"" , ""U_VSPKMS"" as ""Kms"" , ""U_VSPCLSDT"" as ""Closed Date"" , ""U_VSPCLSTM"" as ""Closed Time"" , ""U_VSPSATS"" as ""Status"" , " & _
            """U_VSPREMTO"" as ""Remind To"" , ""U_VSPRMDBY"" as ""Remind By"" From ""@VSP_FLT_PRTVRMDR"" Where ""U_VSPRMVCD"" = '" & objForm.Items.Item("5").Specific.Value & "'"
            Else
                Query = "Select U_VSPREMNM as 'Reminder Name' , U_VSPRMFR as 'For' , U_VSPTDESC as 'Task Description' , U_VSPCDT as 'Raised Date' , " & _
       "U_VSPRBON as 'Based On' , U_VSPKMS as 'Kms' , U_VSPCLSDT as 'Closed Date' , U_VSPCLSTM as 'Closed Time' , U_VSPSATS as 'Status' , " & _
       "U_VSPREMTO as 'Remind To' , U_VSPRMDBY as 'Remind By' From [@VSP_FLT_PRTVRMDR] Where U_VSPRMVCD = '" & objForm.Items.Item("5").Specific.Value & "'"

            End If


            oPVGrid.DataTable.ExecuteQuery(Query)

            For i As Integer = 0 To oPVGrid.DataTable.Columns.Count - 1
                oPVGrid.Columns.Item(i).Editable = False
            Next
            oPVGrid.AutoResizeColumns()

            oACGrid.DataTable = oDT1
            Dim ACQuery As String = ""
            If objMain.IsSAPHANA = True Then
                ACQuery = "Select A.""U_VSPVHCD"" as ""Vehicle Code"",A.""U_VSPVHNM"" as ""Vehicle Name"",A.""U_VSPPLTNO"" as ""Plate No"", A.""U_VSPTRPNO"" as ""Trip No"", " & _
                                   "A.""U_VSPCNTR"" as ""Manager"",C.""Location"" as ""Location"", A.""U_VSPTYPE"" as ""Type"",A.""U_VSPSRVTY"" as ""Severity""," & _
                                   "A.""U_VSPODOMT"" as ""OdoMeter"",A.""U_VSPDATE"" as ""Date"", A.""U_VSPTIME"" as ""Time"" From ""@VSP_FLT_ACCHIST"" A " & _
                                   "Inner Join OLCT C On A.""U_VSPLOC"" = C.""Code"" Where A.""U_VSPVHCD"" = '" & objForm.Items.Item("5").Specific.Value & "'"

            Else
                ACQuery = "Select A.U_VSPVHCD as 'Vehicle Code',A.U_VSPVHNM as 'Vehicle Name',A.U_VSPPLTNO as 'Plate No', A.U_VSPTRPNO as 'Trip No', " & _
                                   "A.U_VSPCNTR as 'Manager',C.Location as 'Location', A.U_VSPTYPE as 'Type',A.U_VSPSRVTY as 'Severity'," & _
                                   "A.U_VSPODOMT as 'OdoMeter',A.U_VSPDATE as 'Date', A.U_VSPTIME as 'Time' From [@VSP_FLT_ACCHIST] A " & _
                                   "Inner Join OLCT C On A.U_VSPLOC = C.Code Where A.U_VSPVHCD = '" & objForm.Items.Item("5").Specific.Value & "'"

            End If
            oACGrid.DataTable.ExecuteQuery(ACQuery)

            For i As Integer = 0 To oACGrid.DataTable.Columns.Count - 1
                oACGrid.Columns.Item(i).Editable = False
            Next
            oACGrid.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix5 = objForm.Items.Item("125").Specific

            If objForm.Items.Item("5").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle No./Registration No. Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("131").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle Name Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("7").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Plate No. Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("104").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Mileage With Load Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("106").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Mileage W/O Load Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("25").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Odometer Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("151").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Vehicle C.C  Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("140").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("No. of Tyres Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If objMatrix5.VisualRowCount <= 1 Then
                objMain.objApplication.StatusBar.SetText("Manager Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            'If objMatrix.VisualRowCount <= 3 Then
            '    objMain.objApplication.StatusBar.SetText("Attach atleast 3 Statutory Documents", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
End Class
