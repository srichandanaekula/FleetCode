Imports System.Threading
Imports System.IO
Public Class clsAccidentHistory
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2, oDBs_Details3 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2, objMatrix3 As SAPbouiCOM.Matrix
    Dim objComboBox As SAPbouiCOM.ComboBox
    Dim oColumn As SAPbouiCOM.Column
    Dim oButton, oButton1 As SAPbouiCOM.Button
    Dim objPicture As SAPbouiCOM.PictureBox
    Dim iPath As Integer
    Dim Path As String
    Dim oLink As SAPbouiCOM.LinkedButton
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Accident History.xml", "VSP_FLT_ACCHIST_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_ACCHIST_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

            objMatrix1 = objForm.Items.Item("35").Specific
            objMatrix2 = objForm.Items.Item("54").Specific
            objMatrix3 = objForm.Items.Item("1000006").Specific

            objComboBox = objForm.Items.Item("53").Specific
            objComboBox.ValidValues.Add("Low", "Low")
            objComboBox.ValidValues.Add("Medium", "Medium")
            objComboBox.ValidValues.Add("High", "High")
            objComboBox.ValidValues.Add("Critical", "Critical")

            objComboBox = objForm.Items.Item("33").Specific
            objComboBox.ValidValues.Add("Open", "Open")
            objComboBox.ValidValues.Add("Close", "Close")
            If objMain.IsSAPHANA = True Then
                objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000002").Specific, "Select ""Code"",""Location"" From OLCT")
            Else
                objMain.objUtilities.ComboBoxLoadValues(objForm.Items.Item("1000002").Specific, "Select Code,Location From OLCT")
            End If



            objForm.Items.Item("43").Visible = False
            objForm.Items.Item("45").Visible = False
            objForm.Items.Item("47").Visible = False
            objForm.Items.Item("49").Visible = False
            objForm.Items.Item("51").Visible = False

            oButton = objForm.Items.Item("43").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("45").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("47").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("49").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            oButton = objForm.Items.Item("51").Specific
            oButton.Image = Application.StartupPath & "\Image.jpg"

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            Me.CellMasking(objForm.UniqueID)

            objForm.PaneLevel = 1
            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CellMasking(ByVal FormUID As String)

        objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        'objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        'objForm.Items.Item("20").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        'objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        'objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        'objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        'objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        objForm.Items.Item("26").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("26").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        objForm.Items.Item("56").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("56").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        objForm.Items.Item("57").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        objForm.Items.Item("57").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_FLT_OACCHIST"))

            objMatrix1 = objForm.Items.Item("35").Specific
            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()

            objMatrix2 = objForm.Items.Item("54").Specific
            objMatrix2.Clear()
            oDBs_Details2.Clear()
            objMatrix2.FlushToDataSource()

            objMatrix3 = objForm.Items.Item("1000006").Specific
            objMatrix3.Clear()
            oDBs_Details3.Clear()
            objMatrix3.FlushToDataSource()

            Me.SetNewLine(objForm.UniqueID, "35")
            Me.SetNewLine(objForm.UniqueID, "54")
            Me.SetNewLine(objForm.UniqueID, "1000006")

            objComboBox = objForm.Items.Item("33").Specific
            objComboBox.Select("Open", SAPbouiCOM.BoSearchKey.psk_ByValue)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

            objMatrix1 = objForm.Items.Item("35").Specific
            objMatrix2 = objForm.Items.Item("54").Specific
            objMatrix3 = objForm.Items.Item("1000006").Specific
            Select Case MatrixUID

                Case "35"
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_VSPDOCTY", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPDOCNM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPDATE", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPDCTOT", oDBs_Details1.Offset, 0)
                    oDBs_Details1.SetValue("U_VSPCOMM", oDBs_Details1.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)
                Case "54"
                    objMatrix2.AddRow()
                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Details2.SetValue("U_VSPANM", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPAPATH", oDBs_Details2.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)
                Case "1000006"
                    objMatrix3.AddRow()
                    oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, objMatrix3.VisualRowCount)
                    oDBs_Details3.SetValue("U_VSPDRCD", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPDRFNM", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPDRMNM", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPDRLNM", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPMBNO", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPLCNO", oDBs_Details3.Offset, "")
                    oDBs_Details3.SetValue("U_VSPEXPDT", oDBs_Details3.Offset, "")
                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)
            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_FLT_ACCHIST" And pVal.BeforeAction = False Then
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("35").Specific
                    objMatrix2 = objForm.Items.Item("54").Specific
                    objMatrix3 = objForm.Items.Item("1000006").Specific
                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If
                    If pVal.ItemUID = "1000001" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 1
                        objForm.Settings.MatrixUID = "35"
                        objMatrix1.AutoResizeColumns()
                    ElseIf pVal.ItemUID = "34" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 2
                        objForm.Settings.MatrixUID = "54"
                        objMatrix2.AutoResizeColumns()
                    ElseIf pVal.ItemUID = "1000003" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 3
                    ElseIf pVal.ItemUID = "1000005" And pVal.BeforeAction = False Then
                        objForm.PaneLevel = 4
                        objForm.Settings.MatrixUID = "1000006"
                        objMatrix3.AutoResizeColumns()
                    End If

                    If pVal.ItemUID = "1000003" And pVal.BeforeAction = False And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        objForm.Items.Item("43").Visible = True
                        objForm.Items.Item("45").Visible = True
                        objForm.Items.Item("47").Visible = True
                        objForm.Items.Item("49").Visible = True
                        objForm.Items.Item("51").Visible = True
                    End If

                    If pVal.BeforeAction = False Then
                        If pVal.ItemUID = "52" Then
                            objPicture = objForm.Items.Item("36").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("52").Visible = False
                        ElseIf pVal.ItemUID = "44" Then
                            objPicture = objForm.Items.Item("37").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("44").Visible = False
                        ElseIf pVal.ItemUID = "46" Then
                            objPicture = objForm.Items.Item("38").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("46").Visible = False
                        ElseIf pVal.ItemUID = "48" Then
                            objPicture = objForm.Items.Item("39").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("48").Visible = False
                        ElseIf pVal.ItemUID = "50" Then
                            objPicture = objForm.Items.Item("40").Specific
                            IO.File.Delete(objPicture.Picture)
                            objPicture.Picture = String.Empty
                            objForm.Items.Item("50").Visible = False
                        End If
                    End If

                    If pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "51" Then
                            objPicture = objForm.Items.Item("36").Specific
                            oButton1 = objForm.Items.Item("52").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "43" Then
                            objPicture = objForm.Items.Item("37").Specific
                            oButton1 = objForm.Items.Item("44").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "45" Then
                            objPicture = objForm.Items.Item("38").Specific
                            oButton1 = objForm.Items.Item("46").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "47" Then
                            objPicture = objForm.Items.Item("39").Specific
                            oButton1 = objForm.Items.Item("48").Specific
                            Me.PictureBrowseFileDialouge()
                        ElseIf pVal.ItemUID = "49" Then
                            objPicture = objForm.Items.Item("40").Specific
                            oButton1 = objForm.Items.Item("50").Specific
                            Me.PictureBrowseFileDialouge()
                        End If
                    End If

                    If pVal.ItemUID = "57" And pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objMain.objDocumentType.CreateForm(objForm.UniqueID, objForm.Items.Item("10").Specific.Value, "Accident History")
                    End If

                Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "54" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        objMatrix2 = objForm.Items.Item("54").Specific
                        Path = pVal.Row
                        iPath = pVal.Row
                        Me.BrowseFileDialog()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
                    oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

                    objMatrix1 = objForm.Items.Item("35").Specific
                    objMatrix2 = objForm.Items.Item("54").Specific
                    objMatrix3 = objForm.Items.Item("1000006").Specific

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

                        If oCFL.UniqueID = "CFL_VHCD" Or oCFL.UniqueID = "CFL_VHNM" Then
                            oDBs_Head.SetValue("U_VSPVHCD", oDBs_Head.Offset, oDT.GetValue("U_VSPVNO", 0))
                            oDBs_Head.SetValue("U_VSPVHNM", oDBs_Head.Offset, oDT.GetValue("U_VSPVNM", 0))
                            oDBs_Head.SetValue("U_VSPPLTNO", oDBs_Head.Offset, oDT.GetValue("U_VSPPNO", 0))
                        End If

                        If oCFL.UniqueID = "CFL_TRPNO" Then
                            Me.LoadTripDetails(objForm.UniqueID, oDT.GetValue("DocNum", 0), oDT.GetValue("U_VSPCNTNM", 0))
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix1 = objForm.Items.Item("35").Specific

                    If pVal.ItemUID = "35" And pVal.ColUID = "V_2" And pVal.BeforeAction = True Then
                        If objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "Purchase Order" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 22
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "GRPO" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 20
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "A/P Invoice" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 18
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "Outgoing Payment" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 46
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "Sales Order" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 17
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "Delivery Order" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 15
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "A/R Invoice" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 13
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "Incoming Payment" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 24
                        ElseIf objMatrix1.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value = "Inventory Transfer" Then
                            oLink = objMatrix1.Columns.Item("V_2").ExtendedObject
                            oLink.LinkedObjectType = 67
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    If pVal.BeforeAction = False And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        objForm.Freeze(True)

                        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
                        oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
                        oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
                        oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")


                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST"), "DocEntry", oDBs_Head.GetValue("DocEntry", 0))
                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0"), "DocEntry", oDBs_Details1.GetValue("DocEntry", 0))
                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1"), "DocEntry", oDBs_Details2.GetValue("DocEntry", 0))
                        objMain.objUtilities.RefreshDatasourceFromDB(FormUID, objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2"), "DocEntry", oDBs_Details3.GetValue("DocEntry", 0))

                        objMatrix1 = objForm.Items.Item("35").Specific
                        objMatrix2 = objForm.Items.Item("54").Specific
                        objMatrix3 = objForm.Items.Item("1000006").Specific

                        objMatrix1.LoadFromDataSource()
                        objMatrix2.LoadFromDataSource()
                        objMatrix3.LoadFromDataSource()

                        objMatrix1.AutoResizeColumns()
                        objMatrix2.AutoResizeColumns()
                        objMatrix3.AutoResizeColumns()

                        objForm.Refresh()

                        objForm.Freeze(False)
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub LoadTripDetails(ByVal FormUID As String, ByVal TripNo As String, ByVal Contractor As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

            objMatrix1 = objForm.Items.Item("35").Specific
            objMatrix2 = objForm.Items.Item("54").Specific
            objMatrix3 = objForm.Items.Item("1000006").Specific

            objMatrix3.Clear()
            oDBs_Details3.Clear()
            objMatrix3.FlushToDataSource()

            oDBs_Head.SetValue("U_VSPTRPNO", oDBs_Head.Offset, TripNo)
            oDBs_Head.SetValue("U_VSPCNTR", oDBs_Head.Offset, Contractor)

            Dim GetOdoMeterReading As String = ""
            If objMain.IsSAPHANA = True Then
                GetOdoMeterReading = "Select Max(""U_VSPOPKM"") as ""OpenKM"",Max(""U_VSPCLKM"") as ""CloseKM"" From ""@VSP_FLT_TRSHT_C1"" Where ""DocEntry"" = " & _
                                               "(Select ""DocNum"" From ""@VSP_FLT_TRSHT"" Where ""DocEntry"" = '" & TripNo & "')"
            Else
                GetOdoMeterReading = "Select Max(U_VSPOPKM) as 'OpenKM',Max(U_VSPCLKM) as 'CloseKM' From [@VSP_FLT_TRSHT_C1] Where DocEntry = " & _
                                               "(Select DocNum From [@VSP_FLT_TRSHT] Where DocEntry = '" & TripNo & "')"
            End If


            Dim oRsGetOdoMeterReading As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetOdoMeterReading.DoQuery(GetOdoMeterReading)

            If oRsGetOdoMeterReading.Fields.Item("CloseKM").Value = 0 Then
                oDBs_Head.SetValue("U_VSPODOMT", oDBs_Head.Offset, oRsGetOdoMeterReading.Fields.Item("OpenKM").Value)
            Else
                oDBs_Head.SetValue("U_VSPODOMT", oDBs_Head.Offset, oRsGetOdoMeterReading.Fields.Item("CloseKM").Value)
            End If

            Dim GetDriverDetails As String = ""
            If objMain.IsSAPHANA = True Then
                GetDriverDetails = "Select * From ""@VSP_FLT_TRSHT_C7"" Where ""DocEntry"" = '" & TripNo & "' And ""U_VSPDRCOD"" <> '' "
            Else
                GetDriverDetails = "Select * From [@VSP_FLT_TRSHT_C7] Where DocEntry = '" & TripNo & "' And U_VSPDRCOD <> '' "
            End If


            Dim oRsGetDriverDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDriverDetails.DoQuery(GetDriverDetails)

            If oRsGetDriverDetails.RecordCount > 0 Then
                oRsGetDriverDetails.MoveFirst()

                For i As Integer = 1 To oRsGetDriverDetails.RecordCount
                    objMatrix3.AddRow()
                    oDBs_Details3.SetValue("LineId", oDBs_Details3.Offset, objMatrix3.VisualRowCount)
                    oDBs_Details3.SetValue("U_VSPDRCD", oDBs_Details3.Offset, oRsGetDriverDetails.Fields.Item("U_VSPDRCOD").Value)
                    oDBs_Details3.SetValue("U_VSPDRFNM", oDBs_Details3.Offset, oRsGetDriverDetails.Fields.Item("U_VSPDRFNM").Value)
                    oDBs_Details3.SetValue("U_VSPDRMNM", oDBs_Details3.Offset, oRsGetDriverDetails.Fields.Item("U_VSPDRMNM").Value)
                    oDBs_Details3.SetValue("U_VSPDRLNM", oDBs_Details3.Offset, oRsGetDriverDetails.Fields.Item("U_VSPDRLNM").Value)
                    oDBs_Details3.SetValue("U_VSPMBNO", oDBs_Details3.Offset, oRsGetDriverDetails.Fields.Item("U_VSPMBNUM").Value)
                    oDBs_Details3.SetValue("U_VSPLCNO", oDBs_Details3.Offset, oRsGetDriverDetails.Fields.Item("U_VSPLNNUM").Value)
                    Dim ExpDate As Date = oRsGetDriverDetails.Fields.Item("U_VSPEXPDT").Value
                    oDBs_Details3.SetValue("U_VSPEXPDT", oDBs_Details3.Offset, ExpDate.ToString("yyyyMMdd"))
                    objMatrix3.SetLineData(objMatrix3.VisualRowCount)

                    oRsGetDriverDetails.MoveNext()
                Next
            End If

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub VehicleAttachment(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objMatrix1 = objForm.Items.Item("35").Specific
            objMatrix2 = objForm.Items.Item("54").Specific
            objMatrix3 = objForm.Items.Item("1000006").Specific

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = DestinationPath & "\" & FileName
                File.Copy(AttchPath, Destination, True)
                objPicture.Picture = Destination
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination, True)
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
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

                            'Dim GetPath As String = "Select U_VSPVHATC From OADM"
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)
                            Dim Attachment As String = MyTest1.FileName
                            If oRsGetPath.Fields.Item("U_VSPVHATC").Value <> "" Then
                                Me.VehicleAttachment(objForm.UniqueID, Path, Attachment, oRsGetPath.Fields.Item("U_VSPVHATC").Value & objForm.Items.Item("4").Specific.Value)
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

                            'Dim GetPath As String = "Select U_VSPVHATC From OADM"
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)
                            Dim Attachment As String = MyTest1.FileName
                            If oRsGetPath.Fields.Item("U_VSPVHATC").Value <> "" Then
                                Me.VehicleAttachment(objForm.UniqueID, Path, Attachment, oRsGetPath.Fields.Item("U_VSPVHATC").Value & objForm.Items.Item("4").Specific.Value)
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
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                    MyTest1.FileName = "Select the Reference Document"
                    objMatrix1 = objForm.Items.Item("35").Specific
                    objMatrix2 = objForm.Items.Item("54").Specific
                    objMatrix3 = objForm.Items.Item("1000006").Specific

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
                                Me.AttachmentVechicals(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPVHATC").Value & "\" & objForm.Items.Item("10").Specific.Value)
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
                                GetPath = "Select U_VSPDRATC From OADM"
                            Else
                                GetPath = "Select U_VSPDRATC From OADM"
                            End If
                            Dim oRsGetPath As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRsGetPath.DoQuery(GetPath)

                            If oRsGetPath.Fields.Item("U_VSPDRATC").Value <> "" Then
                                Me.AttachmentVechicals(objForm.UniqueID, Path, MyTest1.FileName, oRsGetPath.Fields.Item("U_VSPDRATC").Value & "\" & objForm.Items.Item("10").Specific.Value)
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

    Sub AttachmentVechicals(ByVal FormUID As String, ByVal Row As String, ByVal AttchPath As String, ByVal DestinationPath As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C1")
            oDBs_Details3 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_ACCHIST_C2")

            objMatrix1 = objForm.Items.Item("35").Specific
            objMatrix2 = objForm.Items.Item("54").Specific
            objMatrix3 = objForm.Items.Item("1000006").Specific

            Dim DI As DirectoryInfo = New DirectoryInfo(DestinationPath)
            If DI.Exists Then
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)

                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, Path)
                oDBs_Details2.SetValue("U_VSPANM", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_1").Cells.Item(iPath).Specific.Value)
                oDBs_Details2.SetValue("U_VSPAPATH", oDBs_Details2.Offset, Destination)
                objMatrix2.SetLineData(Path)

                objMatrix2.AutoResizeColumns()
                If objMatrix2.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "54")
                End If
            Else
                DI.Create()
                Dim FileName As String = System.IO.Path.GetFileName(AttchPath)
                Dim Destination As String = System.IO.Path.Combine(DestinationPath, FileName)
                File.Copy(AttchPath, Destination)
                oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, Path)
                oDBs_Details2.SetValue("U_VSPANM", oDBs_Details2.Offset, objMatrix2.Columns.Item("V_1").Cells.Item(iPath).Specific.Value)
                oDBs_Details2.SetValue("U_VSPAPATH", oDBs_Details2.Offset, Destination)
                objMatrix2.SetLineData(Path)
                objMatrix2.AutoResizeColumns()
                If objMatrix2.VisualRowCount = Path Then
                    Me.SetNewLine(objForm.UniqueID, "54")
                End If
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

End Class
