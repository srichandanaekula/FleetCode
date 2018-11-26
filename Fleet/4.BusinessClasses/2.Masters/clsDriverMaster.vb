Imports System.Threading
Imports System.IO
Public Class clsDriverMaster

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details3, oDBs_Details4, oDBs_Details5 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix3, objMatrix4, objMatrix5 As SAPbouiCOM.Matrix
    Dim objPicture As SAPbouiCOM.PictureBox
    Dim ImgBtn1, ImgBtn2 As SAPbouiCOM.Button
    Dim Path As String
    Dim iPath As Integer
    Dim oAHGrid As SAPbouiCOM.Grid
    Dim oDT As SAPbouiCOM.DataTable
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("DriverMaster.xml", "VSP_FLT_DRVRMSTR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_DRVRMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRVRMSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C0")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C2")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C3")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C4")

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            ImgBtn1 = objForm.Items.Item("74").Specific
            ImgBtn1.Image = Application.StartupPath & "\Image.jpg"

            ImgBtn1 = objForm.Items.Item("75").Specific
            ImgBtn1.Image = Application.StartupPath & "\RemoveImage.jpg"

            objForm.Items.Item("1000011").TextStyle = 4
            objForm.Items.Item("1000009").TextStyle = 4
            objForm.Items.Item("28").TextStyle = 4
            objForm.Items.Item("39").TextStyle = 4

            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000002").Specific, "Select ""Code"" , ""Name"" From OCST Where ""Country"" = 'IN'")
            objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000005").Specific, "Select ""Code"" , ""Location"" From OLCT ")

            objForm.Items.Item("65").AffectsFormMode = False
            objForm.Items.Item("66").AffectsFormMode = False
            objForm.Items.Item("67").AffectsFormMode = False
            objForm.Items.Item("1000015").AffectsFormMode = False

            Me.CFLFilter(objForm.UniqueID, "CFL_CNTCD")
            Me.CFLFilter(objForm.UniqueID, "CFL_CM")

            oDT = objForm.DataSources.DataTables.Add("dt1")
            oDT = objForm.DataSources.DataTables.Item("dt1")

            Me.CellsMasking(objForm.UniqueID)

            objForm.PaneLevel = 1
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("74").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("74").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("74").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("75").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("75").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("75").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000010").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000010").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("1000013").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("1000013").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("13").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("72").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("72").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("60").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("60").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRVRMSTR")
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSP_FLT_DRVRMSTR"))
            oDBs_Head.SetValue("U_VSPAVLBL", oDBs_Head.Offset, "N")

            objForm.Items.Item("76").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.PaneLevel = 1

            objMatrix1 = objForm.Items.Item("67").Specific
            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            objMatrix1.AutoResizeColumns()

            objMatrix3 = objForm.Items.Item("66").Specific
            objMatrix3.Clear()
            oDBs_Details3.Clear()
            objMatrix3.FlushToDataSource()
            objMatrix3.AutoResizeColumns()

            objMatrix4 = objForm.Items.Item("65").Specific
            objMatrix4.Clear()
            oDBs_Details4.Clear()
            objMatrix4.FlushToDataSource()
            objMatrix4.AutoResizeColumns()

            objMatrix5 = objForm.Items.Item("77").Specific
            objMatrix5.Clear()
            oDBs_Details5.Clear()
            objMatrix5.FlushToDataSource()
            objMatrix5.AutoResizeColumns()

            Me.SetNewLine(objForm.UniqueID, "67")
            Me.SetNewLine(objForm.UniqueID, "66")
            Me.SetNewLine(objForm.UniqueID, "65")
            Me.SetNewLine(objForm.UniqueID, "77")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRVRMSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C0")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C2")
            oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C3")
            oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C4")

            Select Case MatrixUID

                Case "67"
                    objMatrix1 = objForm.Items.Item("67").Specific
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_VSPINTYP", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPINCMP", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPINLN", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPINSDT", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPINEDT", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPINAMT", oDBs_Details1.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                Case "66"
                    objMatrix3 = objForm.Items.Item("66").Specific
                    objMatrix3.AddRow()
                    oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, objMatrix3.VisualRowCount)
                    oDBs_Details3.SetValue("U_VSPOANUM", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPOANME", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPOAIDT", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPOAEDT", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPOAATC", oDBs_Details3.Offset, "")
                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                Case "65"
                    objMatrix4 = objForm.Items.Item("65").Specific
                    objMatrix4.AddRow()
                    oDBs_Details4.SetValue("LineId", oDBs_Details4.Offset, objMatrix4.VisualRowCount)
                    oDBs_Details4.SetValue("U_VSPCHCLN", oDBs_Details4.Offset, "")
                    oDBs_Details4.SetValue("U_VSPCHCHB", oDBs_Details4.Offset, "")
                    objMatrix4.SetLineData(objMatrix4.VisualRowCount)

                Case "77"
                    objMatrix5 = objForm.Items.Item("77").Specific
                    objMatrix5.AddRow()
                    oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, objMatrix5.VisualRowCount)
                    oDBs_Details5.SetValue("U_VSPFDT", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, "99991231")
                    oDBs_Details5.SetValue("U_VSPCNTCD", oDBs_Details5.Offset, "")
                    oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, "")
                    objMatrix5.SetLineData(objMatrix5.VisualRowCount)
            End Select

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
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                    If pVal.ItemUID = "64" And pVal.BeforeAction = False Then
                        objMatrix1 = objForm.Items.Item("67").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 4
                        objForm.Settings.MatrixUID = "67"
                        objMatrix1.AutoResizeColumns()
                        objForm.Freeze(False)
                    ElseIf pVal.ItemUID = "63" And pVal.BeforeAction = False Then
                        oAHGrid = objForm.Items.Item("1000015").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 2
                        objForm.Settings.MatrixUID = "1000015"
                        objForm.Freeze(False)
                    ElseIf pVal.ItemUID = "61" And pVal.BeforeAction = False Then
                        objMatrix3 = objForm.Items.Item("66").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 3
                        objForm.Settings.MatrixUID = "66"
                        objMatrix3.AutoResizeColumns()
                        objForm.Freeze(False)
                    ElseIf pVal.ItemUID = "62" And pVal.BeforeAction = False Then
                        objMatrix4 = objForm.Items.Item("65").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 5
                        objForm.Settings.MatrixUID = "65"
                        objMatrix4.AutoResizeColumns()
                        objForm.Freeze(False)
                    ElseIf pVal.ItemUID = "76" And pVal.BeforeAction = False Then
                        objMatrix5 = objForm.Items.Item("77").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 1
                        objForm.Settings.MatrixUID = "77"
                        objMatrix5.AutoResizeColumns()
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "74" And pVal.BeforeAction = False Then
                        Me.PictureBrowseFileDialog()
                    End If

                    If pVal.ItemUID = "75" And pVal.BeforeAction = False Then
                        objPicture = objForm.Items.Item("1000003").Specific
                        IO.File.Delete(objPicture.Picture)
                        objPicture.Picture = String.Empty
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "66" And pVal.ColUID = "V_4" And pVal.BeforeAction = False Then
                        objMatrix3 = objForm.Items.Item("66").Specific
                        Path = pVal.Row
                        iPath = pVal.Row
                        Me.BrowseFileDialog()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix5 = objForm.Items.Item("77").Specific
                    If pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If pVal.ItemUID = "77" And pVal.ColUID = "V_0" And pVal.Row <> "1" Then
                            If objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                Dim BeforeDate As String = objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row - 1).Specific.Value
                                Dim PresentDate As String = objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value

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
                                        oDBs_Details5.SetValue("U_VSPFDT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details5.SetValue("U_VSPCNTCD", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                                        oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                                        objMatrix5.SetLineData(pVal.Row)
                                    Else
                                        objMain.objApplication.StatusBar.SetText("Date should not be less than Preceeding Date", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value = ""
                                        BubbleEvent = False
                                        Exit Try
                                    End If
                                End If
                            End If
                        End If
                    End If


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRVRMSTR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C0")
                    oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C2")
                    oDBs_Details4 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C3")
                    oDBs_Details5 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C4")
                    objMatrix5 = objForm.Items.Item("77").Specific
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
                        If oCFL.UniqueID = "CFL_CM" Then
                            oDBs_Head.SetValue("U_VSPCNCD", oDBs_Head.Offset, oDT.GetValue("U_VSPCNTR", 0))
                        End If

                        If oCFL.UniqueID = "CFL_CNTCD" Then

                            Dim GetFstNmLstNm As String = ""
                            If objMain.IsSAPHANA = True Then
                                GetFstNmLstNm = "Select ""firstName"" , ""lastName"" From OHEM Where ""empID"" = '" & oDT.GetValue(0, 0) & "' "
                            Else
                                GetFstNmLstNm = "Select firstName , lastName From OHEM Where empId = '" & oDT.GetValue(0, 0) & "' "
                            End If

                            Dim oRsGetFstNmLstNm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetFstNmLstNm.DoQuery(GetFstNmLstNm)
                            oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, pVal.Row)
                            oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPFDT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPCNTCD", oDBs_Details5.Offset, oDT.GetValue(0, 0))
                            oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, oRsGetFstNmLstNm.Fields.Item(0).Value & " " & oRsGetFstNmLstNm.Fields.Item(1).Value)
                            objMatrix5.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix5.VisualRowCount Then
                                SetNewLine(objForm.UniqueID, "77")
                            End If
                        End If

                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix4 = objForm.Items.Item("65").Specific
                    objMatrix5 = objForm.Items.Item("77").Specific
                    objMatrix1 = objForm.Items.Item("67").Specific
                    If pVal.ItemUID = "77" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        objMatrix5 = objForm.Items.Item("77").Specific
                        If pVal.Row > 1 And objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Dim GetDate As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetDate = "Select ADD_DAYS(" & _
                          "CAST(REPLACE('" & objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value & "', '-', '') as DATETIME), -1) From DUMMY "
                            Else
                                GetDate = "Select  DATEADD(DAY, -1, " & _
                          "CAST(REPLACE('" & objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value & "', '-', '') AS DATETIME)) "
                            End If
                            Dim oRsGetDate As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetDate.DoQuery(GetDate)
                            Dim SDate As Date = oRsGetDate.Fields.Item(0).Value

                            oDBs_Details5.SetValue("LineId", oDBs_Details5.Offset, pVal.Row - 1)
                            oDBs_Details5.SetValue("U_VSPTODT", oDBs_Details5.Offset, SDate.ToString("yyyyMMdd"))
                            oDBs_Details5.SetValue("U_VSPFDT", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_0").Cells.Item(pVal.Row - 1).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPCNTCD", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(pVal.Row - 1).Specific.Value)
                            oDBs_Details5.SetValue("U_VSPCNTNM", oDBs_Details5.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(pVal.Row - 1).Specific.Value)
                            objMatrix5.SetLineData(pVal.Row - 1)
                        End If
                    End If
                    If pVal.ItemUID = "65" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        If objMatrix4.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix4.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "65")
                            End If
                        End If
                    End If
                    If pVal.ItemUID = "67" And pVal.ColUID = "V_1" And pVal.BeforeAction = False Then
                        If objMatrix1.Columns.Item("V_1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                Me.SetNewLine(objForm.UniqueID, "67")
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                    If pVal.ItemUID = "" And pVal.BeforeAction = False And pVal.CharPressed = "9" Then
                        objMain.objSelectDistributionRules.CreateForm(objForm.UniqueID, "@VSP_FLT_DRVRMSTR", "Yes", objForm.TypeEx, "", 0)
                    ElseIf pVal.ItemUID = "" And pVal.ColUID = "" And pVal.BeforeAction = False And pVal.CharPressed = "9" Then
                        objMain.objSelectDistributionRules.CreateForm(objForm.UniqueID, "@VSP_FLT_DRVRMSTR", "No", objForm.TypeEx, "Matrixid", pVal.Row)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
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
                        If oMenus.Exists("View") = True Then
                            objMain.objApplication.Menus.RemoveEx("View")
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
            If pVal.MenuUID = "VSP_FLT_DRVRMSTR" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "View" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_DRVRMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix3 = objForm.Items.Item("66").Specific
                For i As Integer = 1 To objMatrix3.VisualRowCount
                    If objMatrix3.IsRowSelected(i) = True Then
                        If objMatrix3.Columns.Item("V_4").Cells.Item(i).Specific.Value <> "" Then
                            System.Diagnostics.Process.Start(objMatrix3.Columns.Item("V_4").Cells.Item(i).Specific.Value)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub LoadGrid(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oAHGrid = objForm.Items.Item("1000015").Specific
            oAHGrid.DataTable = oDT

            Dim AHQuery As String = ""

            If objMain.IsSAPHANA = True Then
                AHQuery = "Select A.""U_VSPVHCD"" as ""Vehicle Code"",A.""U_VSPVHNM"" as ""Vehicle Name"",A.""U_VSPPLTNO"" as ""Plate No"", A.""U_VSPTRPNO"" as ""Trip No"", " & _
                                   "A.""U_VSPCNTR"" as ""Manager"",C.""Location"" as ""Location"", A.""U_VSPTYPE"" as ""Type"",A.""U_VSPSRVTY"" as ""Severity""," & _
                                   "A.""U_VSPODOMT"" as ""OdoMeter"",A.""U_VSPDATE"" as ""Date"", A.""U_VSPTIME"" as ""Time"" From ""@VSP_FLT_ACCHIST"" A Inner Join " & _
                                   """@VSP_FLT_ACCHIST_C2"" B On A.""DocEntry""=B.""DocEntry"" Inner Join OLCT C On A.""U_VSPLOC"" = C.""Code"" " & _
                                   "Where B.""U_VSPDRCD"" = '" & objForm.Items.Item("72").Specific.Value & "'"

            Else
                AHQuery = "Select A.U_VSPVHCD as 'Vehicle Code',A.U_VSPVHNM as 'Vehicle Name',A.U_VSPPLTNO as 'Plate No', A.U_VSPTRPNO as 'Trip No', " & _
                                    "A.U_VSPCNTR as 'Manager',C.Location as 'Location', A.U_VSPTYPE as 'Type',A.U_VSPSRVTY as 'Severity'," & _
                                    "A.U_VSPODOMT as 'OdoMeter',A.U_VSPDATE as 'Date', A.U_VSPTIME as 'Time' From [@VSP_FLT_ACCHIST] A Inner Join " & _
                                    "[@VSP_FLT_ACCHIST_C2] B On A.DocEntry=B.DocEntry Inner Join OLCT C On A.U_VSPLOC = C.Code " & _
                                    "Where B.U_VSPDRCD = '" & objForm.Items.Item("72").Specific.Value & "'"

            End If




            oAHGrid.DataTable.ExecuteQuery(AHQuery)

            For i As Integer = 0 To oAHGrid.DataTable.Columns.Count - 1
                oAHGrid.Columns.Item(i).Editable = False
            Next
            oAHGrid.AutoResizeColumns()

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#Region "    Picture Attachment    "
    Sub PictureBrowseFileDialog()
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
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"

                    'Windows XP
                    If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                        Try
                            objPicture = objForm.Items.Item("1000003").Specific
                            Dim GetPath As String = ""
                            If objMain.IsSAPHANA = True Then
                                GetPath = "Select U_VSPDRATC From OADM"
                            Else
                                GetPath = "Select U_VSPDRATC From OADM"
                            End If

                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPDRATC").Value <> "" Then
                                Me.DriverAttachment(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPDRATC").Value & "\" & objForm.Items.Item("72").Specific.Value)
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
                            objPicture = objForm.Items.Item("1000003").Specific

                            Dim GetPath As String = ""

                            If objMain.IsSAPHANA = True Then
                                GetPath = "Select ""U_VSPDRATC"" From OADM"
                            Else
                                GetPath = "Select U_VSPDRATC From OADM"
                            End If
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPDRATC").Value <> "" Then
                                Me.DriverAttachment(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPDRATC").Value & "\" & objForm.Items.Item("72").Specific.Value)
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
#End Region

    Sub DriverAttachment(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
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

#Region "Attachment"
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
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRMSTR_C2")
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"
                    objMatrix3 = objForm.Items.Item("66").Specific

                    'Windows XP
                    If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                        Try
                            Dim GetPath As String = "Select U_VSPDRATC From OADM"
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPDRATC").Value <> "" Then
                                Me.DriverAttachmentVechicals(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPDRATC").Value & "\" & objForm.Items.Item("72").Specific.Value)
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

                            Dim GetPath As String = "Select U_VSPDRATC From OADM"
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPDRATC").Value <> "" Then
                                Me.DriverAttachmentVechicals(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPDRATC").Value & "\" & objForm.Items.Item("72").Specific.Value)
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

    Sub DriverAttachmentVechicals(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)

                oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, Path)
                oDBs_Details3.SetValue("U_VSPOANUM", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOANME", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOAIDT", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOAEDT", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_3").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOAATC", oDBs_Details3.Offset, Destination)
                objMatrix3.SetLineData(Path)

                objMatrix3.AutoResizeColumns()

                If objMatrix3.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "66")
                End If
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)
                oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, Path)
                oDBs_Details3.SetValue("U_VSPOANUM", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_0").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOANME", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_1").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOAIDT", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_2").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOAEDT", oDBs_Details3.Offset, objMatrix3.Columns.Item("V_3").Cells.Item(iPath).Specific.Value)
                oDBs_Details3.SetValue("U_VSPOAATC", oDBs_Details3.Offset, Destination)
                objMatrix3.SetLineData(Path)
                objMatrix3.AutoResizeColumns()

                If objMatrix3.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "66")
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
                        objMatrix3 = objForm.Items.Item("66").Specific
                        If eventInfo.ItemUID = "66" And eventInfo.ColUID = "V_-1" And objMatrix3.RowCount > 1 Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("View") = False Then
                                    oCreationPackage.UniqueID = "View"
                                    oCreationPackage.String = "View"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "66" And objMatrix3.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("View") = True Then
                                    objMain.objApplication.Menus.RemoveEx("View")
                                End If
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "66" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
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
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_DRVRMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_DRVRMSTR")
            objMatrix5 = objForm.Items.Item("77").Specific

            Select Case BusinessObjectInfo.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD

                    If BusinessObjectInfo.BeforeAction = True And objMatrix5.VisualRowCount <> 1 Then
                        Try
                            oDBs_Head.SetValue("U_VSPCNCD", oDBs_Head.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)
                            oDBs_Head.SetValue("U_VSPCNAM", oDBs_Head.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        If BusinessObjectInfo.FormTypeEx = "VSP_FLT_DRVRMSTR_Form" And objForm.Items.Item("14").Specific.Value <> "" Then
                            Me.AddingEmpRecord(objForm.UniqueID)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE

                    If BusinessObjectInfo.BeforeAction = True Then
                        Try
                            oDBs_Head.SetValue("U_VSPCNCD", oDBs_Head.Offset, objMatrix5.Columns.Item("V_2").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)
                            oDBs_Head.SetValue("U_VSPCNAM", oDBs_Head.Offset, objMatrix5.Columns.Item("V_3").Cells.Item(objMatrix5.VisualRowCount - 1).Specific.Value)
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then
                        Me.LoadGrid(objForm.UniqueID)

                        If objForm.Items.Item("14").Specific.Value = "" Then
                            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        Else
                            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                            objForm.Items.Item("14").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                        End If

                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Sub AddingEmpRecord(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oEmpMstr As SAPbobsCOM.EmployeesInfo = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo)

            oEmpMstr.FirstName = objForm.Items.Item("14").Specific.Value
            oEmpMstr.LastName = objForm.Items.Item("16").Specific.Value
            oEmpMstr.MobilePhone = objForm.Items.Item("53").Specific.Value
            oEmpMstr.eMail = objForm.Items.Item("57").Specific.Value
            oEmpMstr.UserFields.Fields.Item("U_VSPDRCOD").Value = objForm.Items.Item("72").Specific.Value

            Dim GetDriverDepartment As String = ""

            If objMain.IsSAPHANA = True Then
                GetDriverDepartment = "Select ""U_VSPDD"" From ""@VSP_FLT_CNFGSRN"""
            Else
                GetDriverDepartment = "Select U_VSPDD From [@VSP_FLT_CNFGSRN]"
            End If
            Dim oRsGetDriverDepartment As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDriverDepartment.DoQuery(GetDriverDepartment)
            oEmpMstr.Department = oRsGetDriverDepartment.Fields.Item(0).Value

            If oEmpMstr.Add() = 0 Then
                objMain.objApplication.StatusBar.SetText("Employee Record Succefully Added", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim GetEmployeeNo As String = ""
                If objMain.IsSAPHANA = True Then
                    GetEmployeeNo = "Select Max(""empID"") From OHEM"
                Else
                    GetEmployeeNo = "Select Max(""empID"") From OHEM"
                End If

                Dim oRsGetEmployeeNo As SAPbobsCOM.Recordset = _
                                objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetEmployeeNo.DoQuery(GetEmployeeNo)


                Dim GetMaxCode As String = ""
                If objMain.IsSAPHANA = True Then
                    GetMaxCode = "Select ""Code"" From ""@VSP_FLT_DRVRMSTR"" Where ""DocEntry"" = (Select Max(""DocEntry"") From ""@VSP_FLT_DRVRMSTR"")"
                Else
                    GetMaxCode = "Select Code From ""@VSP_FLT_DRVRMST"" Where DocEntry = (Select Max(""DocEntry"") From ""@VSP_FLT_DRVRMSTR"")"
                End If



                Dim oRsGetMaxCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetMaxCode.DoQuery(GetMaxCode)

                objMain.sCmp = objMain.objCompany.GetCompanyService
                objMain.oGeneralService = objMain.sCmp.GetGeneralService("VSP_FLT_ODRVRMSTR")
                objMain.oGeneralParams = objMain.oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                objMain.oGeneralParams.SetProperty("Code", oRsGetMaxCode.Fields.Item(0).Value)
                objMain.oGeneralData = objMain.oGeneralService.GetByParams(objMain.oGeneralParams)
                objMain.oGeneralData.SetProperty("U_VSPEMPNO", oRsGetEmployeeNo.Fields.Item(0).Value.ToString)
                
                objMain.oGeneralService.Update(objMain.oGeneralData)
            Else
                objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
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
            oCondition.Alias = "dept"
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL

            Dim GetCntr As String = ""

            If objMain.IsSAPHANA = True Then
                GetCntr = "Select ""U_VSPCNT"" From ""@VSP_FLT_CNFGSRN"""
            Else
                GetCntr = "Select ""U_VSPCNT"" From ""@VSP_FLT_CNFGSRN"""
            End If



            Dim oRsGetCntr As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCntr.DoQuery(GetCntr)

            oCondition.CondVal = oRsGetCntr.Fields.Item("U_VSPCNT").Value
            oChooseFromList.SetConditions(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix5 = objForm.Items.Item("77").Specific

            If objForm.Items.Item("14").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("First Name Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("16").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Last Name Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("23").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Date of Birth Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("53").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Mobile No. Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("21").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Hire Date Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("79").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("City Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
                'ElseIf objForm.Items.Item("1000002").Specific.Value = "" Then
                '    objMain.objApplication.StatusBar.SetText("State Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '    Return False
            End If

            If objMatrix5.VisualRowCount <= 1 Then
                objMain.objApplication.StatusBar.SetText("Manager Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

End Class

