Imports System.Threading
Imports System.Data.SqlClient
Imports System.Windows.Forms

Public Class clsRouteMaster
#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details1, oDBs_Details2 As SAPbouiCOM.DBDataSource
    Dim objMatrix1, objMatrix2 As SAPbouiCOM.Matrix
    Dim objPicture As SAPbouiCOM.PictureBox
    Dim ImgBtn1, ImgBtn2 As SAPbouiCOM.Button
    Dim Path As String
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Route Master.xml", "VSP_FLT_RTMSTR_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_RTMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C1")

            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("12").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("22").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("18").AffectsFormMode = False
            objForm.Items.Item("17").AffectsFormMode = False

            objMain.objUtilities.AddValidValue(objForm.UniqueID, objForm.TypeEx)

            Me.CFLFilterExpenseAcc(objForm.UniqueID, "CFL_EACCT")

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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C1")
            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSP_FLT_RTMSTR"))

            objForm.Items.Item("17").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objForm.PaneLevel = 1

            objMatrix1 = objForm.Items.Item("19").Specific
            objMatrix1.Clear()
            oDBs_Details1.Clear()
            objMatrix1.FlushToDataSource()
            objMatrix1.AutoResizeColumns()

            objMatrix2 = objForm.Items.Item("20").Specific
            objMatrix2.Clear()
            oDBs_Details2.Clear()
            objMatrix2.FlushToDataSource()
            objMatrix2.AutoResizeColumns()

            Me.SetNewLine(objForm.UniqueID, "19")
            Me.SetNewLine(objForm.UniqueID, "20")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String, ByVal MatrixUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C1")

            objMatrix1 = objForm.Items.Item("19").Specific
            objMatrix2 = objForm.Items.Item("20").Specific

            Select Case MatrixUID
                Case "19"
                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_VSPEACD", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPEANM", oDBs_Details1.Offset, "")
                    oDBs_Details1.SetValue("U_VSPAMT", oDBs_Details1.Offset, "")
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                Case "20"
                    objMatrix2.AddRow()
                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Details2.SetValue("U_VSPFPLC", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPTPLC", oDBs_Details2.Offset, "")
                    oDBs_Details2.SetValue("U_VSPDISTN", oDBs_Details2.Offset, "")
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)
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

                    If pVal.ItemUID = "17" And pVal.BeforeAction = False Then
                        objMatrix1 = objForm.Items.Item("19").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 1
                        objForm.Settings.MatrixUID = "19"
                        objMatrix1.AutoResizeColumns()
                        objForm.Freeze(False)
                    ElseIf pVal.ItemUID = "18" And pVal.BeforeAction = False Then
                        objMatrix2 = objForm.Items.Item("20").Specific
                        objForm.Freeze(True)
                        objForm.PaneLevel = 2
                        objForm.Settings.MatrixUID = "20"
                        objMatrix2.AutoResizeColumns()
                        objForm.Freeze(False)
                    End If

                    If pVal.ItemUID = "25" And pVal.BeforeAction = False Then
                        Me.BrowseFileDialog()
                    End If

                    If pVal.ItemUID = "22" And pVal.BeforeAction = False Then
                        If objForm.Items.Item("24").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Select Excel File", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        ElseIf objForm.Items.Item("29").Specific.Value = "" Then
                            objMain.objApplication.StatusBar.SetText("Please Enter Sheet Name", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Else
                            Me.UpLoadExcel(objForm.UniqueID)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    If pVal.ItemUID = "4" And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        Dim ChkItemExist As String = ""
                        If objMain.IsSAPHANA = True Then
                            ChkItemExist = "Select ""Code"" From ""@VSP_FLT_RTMSTR"" Where ""U_VSPRCD"" ='" & objForm.Items.Item("4").Specific.Value.Trim & "'"
                        Else
                            ChkItemExist = "Select Code From [@VSP_FLT_RTMSTR] Where [U_VSPRCD] ='" & objForm.Items.Item("4").Specific.Value.Trim & "'"
                        End If

                        Dim oRsChkItemExist As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsChkItemExist.DoQuery(ChkItemExist)
                        If oRsChkItemExist.RecordCount > 0 Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("12").Specific.value = oRsChkItemExist.Fields.Item(0).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix2 = objForm.Items.Item("20").Specific
                    If pVal.ItemUID = "20" And pVal.ColUID = "V_2" And pVal.BeforeAction = False Then
                        If pVal.Row = objMatrix2.VisualRowCount And objMatrix2.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            SetNewLine(objForm.UniqueID, "20")
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR")
                    oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C0")
                    oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C1")
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

                        If oCFL.UniqueID = "CFL_EACCT" Then

                            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, pVal.Row)
                            oDBs_Details1.SetValue("U_VSPEACD", oDBs_Details1.Offset, oDT.GetValue(0, 0))
                            oDBs_Details1.SetValue("U_VSPEANM", oDBs_Details1.Offset, oDT.GetValue(1, 0))
                            oDBs_Details1.SetValue("U_VSPAMT", oDBs_Details1.Offset, "")
                            objMatrix1.SetLineData(pVal.Row)
                            If pVal.Row = objMatrix1.VisualRowCount Then
                                SetNewLine(objForm.UniqueID, "19")
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
                        If oMenus.Exists("Add Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Add Row")
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
                        If oMenus.Exists("Add Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Add Row")
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
            If pVal.MenuUID = "VSP_FLT_RTMSTR" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Add Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_RTMSTR_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix2 = objForm.Items.Item("20").Specific
                For i As Integer = 1 To objMatrix2.VisualRowCount - 1
                    If objMatrix2.IsRowSelected(i) = True Then
                        objMatrix2.AddRow(1, i)
                    End If
                Next
                For i As Integer = 1 To objMatrix2.VisualRowCount
                    objMatrix2.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub CFLFilterExpenseAcc(ByVal FormUID As String, ByVal CFL_ID As String)
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
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub UpLoadExcel(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR")
            oDBs_Details1 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C0")
            oDBs_Details2 = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR_C1")

            objMatrix1 = objForm.Items.Item("19").Specific
            objMatrix2 = objForm.Items.Item("20").Specific

            Dim GetPswd As String = ""
            If objMain.IsSAPHANA = True Then
                GetPswd = "Select ""U_VSPPSWD"" From ""@VSP_FLT_CNFGSRN"""
            Else
                GetPswd = "Select U_VSPPSWD From [@VSP_FLT_CNFGSRN]"
            End If
            Dim oRs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery(GetPswd)

            If Me.ImportData(objForm.UniqueID, objMain.objCompany.Server, objMain.objCompany.CompanyDB, objMain.objCompany.DbUserName, _
                          oRs.Fields.Item(0).Value, objForm.Items.Item("24").Specific.Value.Trim, _
                          objForm.Items.Item("29").Specific.Value.Trim) = True Then

                Dim GetExcelDtails As String = ""

                If objMain.IsSAPHANA = True Then
                    GetExcelDtails = "Select *  From " & """" & objForm.Items.Item("29").Specific.Value & """" & ""
                Else
                    GetExcelDtails = "Select *  From " & "[" & objForm.Items.Item("29").Specific.Value & "]" & ""
                End If
                Dim oRsGetExcelDtails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetExcelDtails.DoQuery(GetExcelDtails)

                objMatrix1.Clear()
                oDBs_Details1.Clear()
                objMatrix1.FlushToDataSource()

                objMatrix2.Clear()
                oDBs_Details2.Clear()
                objMatrix2.FlushToDataSource()

                oDBs_Head.SetValue("U_VSPRCD", oDBs_Head.Offset, oRsGetExcelDtails.Fields.Item(0).Value)
                oDBs_Head.SetValue("U_VSPRNM", oDBs_Head.Offset, oRsGetExcelDtails.Fields.Item(1).Value)
                oDBs_Head.SetValue("U_VSPSRCE", oDBs_Head.Offset, oRsGetExcelDtails.Fields.Item(2).Value)
                oDBs_Head.SetValue("U_VSPDEST", oDBs_Head.Offset, oRsGetExcelDtails.Fields.Item(3).Value)
                oDBs_Head.SetValue("U_VSPADAMT", oDBs_Head.Offset, oRsGetExcelDtails.Fields.Item(4).Value)
                oDBs_Head.SetValue("U_VSPTTLKM", oDBs_Head.Offset, oRsGetExcelDtails.Fields.Item(5).Value)

                For i As Integer = 1 To oRsGetExcelDtails.RecordCount

                    objMatrix1.AddRow()
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, objMatrix1.VisualRowCount)
                    oDBs_Details1.SetValue("U_VSPEACD", oDBs_Details1.Offset, oRsGetExcelDtails.Fields.Item(7).Value)
                    oDBs_Details1.SetValue("U_VSPEANM", oDBs_Details1.Offset, oRsGetExcelDtails.Fields.Item(8).Value)
                    oDBs_Details1.SetValue("U_VSPAMT", oDBs_Details1.Offset, oRsGetExcelDtails.Fields.Item(9).Value)
                    objMatrix1.SetLineData(objMatrix1.VisualRowCount)

                    objMatrix2.AddRow()
                    oDBs_Details2.SetValue("LineId", oDBs_Details2.Offset, objMatrix2.VisualRowCount)
                    oDBs_Details2.SetValue("U_VSPFPLC", oDBs_Details2.Offset, oRsGetExcelDtails.Fields.Item(10).Value)
                    oDBs_Details2.SetValue("U_VSPTPLC", oDBs_Details2.Offset, oRsGetExcelDtails.Fields.Item(11).Value)
                    oDBs_Details2.SetValue("U_VSPDISTN", oDBs_Details2.Offset, oRsGetExcelDtails.Fields.Item(12).Value)
                    objMatrix2.SetLineData(objMatrix2.VisualRowCount)

                    oRsGetExcelDtails.MoveNext()
                Next

                objMatrix1.AutoResizeColumns()
                objMatrix2.AutoResizeColumns()
            End If

            Me.SetNewLine(objForm.UniqueID, "19")
            Me.SetNewLine(objForm.UniqueID, "20")

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function ImportData(ByVal FormUID As String, ByVal ServerName As String, _
        ByVal DBName As String, ByVal UserName As String, _
        ByVal Password As String, ByVal ExcelPath As String, ByVal SheetName As String)
        Try
            Dim ExceCon As String = _
                "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                ExcelPath & "; Extended Properties=Excel 12.0"
            Dim connectionString As String = "Server=" & ServerName & " ;" & _
                                             "DataBase=" & DBName & ";" & _
                                             "Uid=" & UserName & ";Pwd=" & Password & ";"

            Dim strSQL As String = "USE [" & DBName & "]" & vbCrLf & _
                                   "IF EXISTS (" & _
                                   "SELECT * " & _
                                   "FROM [" & DBName & "].dbo.sysobjects " & _
                                   "WHERE Name = '" & SheetName & "')" & vbCrLf & _
                                   "BEGIN" & vbCrLf & _
                                   "DROP TABLE [" & DBName & "].dbo." & SheetName & vbCrLf & _
                                   "END "

            Dim dbConnection As New SqlConnection(connectionString)
            ' A SqlCommand object is used to execute the SQL commands.
            Dim cmd As New SqlCommand(strSQL, dbConnection)
            dbConnection.Open()
            cmd.ExecuteNonQuery()
            dbConnection.Close()
            Dim excelConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(ExceCon)
            excelConnection.Open()
            Dim OleStr As String = "SELECT * INTO [ODBC; Driver={SQL Server};Server=" _
                                   & ServerName & ";Database=" & DBName & ";Uid=" & _
                                   UserName & ";Pwd=" & Password & "; ].[" & _
                                   SheetName & "]   FROM [" & SheetName & "$];"

            Dim excelCommand As New System.Data.OleDb.OleDbCommand(OleStr, excelConnection)
            excelCommand.ExecuteNonQuery()
            excelConnection.Close()

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            Return False
        End Try
    End Function

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objForm.Items.Item("4").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Route Code Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("6").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Route Name Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("8").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Source Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("10").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Destination Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("14").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Advance Amount Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("16").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Total KM's Cannot Be Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

#Region " Right Click Event"
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
                        objMatrix2 = objForm.Items.Item("20").Specific
                        If eventInfo.ItemUID = "20" And eventInfo.ColUID = "V_-1" And objMatrix2.RowCount > 1 Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Add Row") = False Then
                                    oCreationPackage.UniqueID = "Add Row"
                                    oCreationPackage.String = "Add Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "20" And objMatrix2.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
                                End If

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "20" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
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
                        If oMenus.Exists("Add Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Add Row")
                        End If

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
                End If
            End If

            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "    Attachment    "

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
        MyProcs = Process.GetProcessesByName("SAP Business One")
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_FLT_RTMSTR")
        If MyProcs.Length <> 0 Then
            For i As Integer = 0 To MyProcs.Length - 1

                Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)
                MyTest1.FileName = "Select the Reference Document"
                'Windows XP
                If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                    Try

                        oDBs_Head.SetValue("U_VSPATCH", oDBs_Head.Offset, MyTest1.FileName)
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    Catch ex As IO.IOException
                        objMain.objApplication.MessageBox(ex.Message)
                        Exit Sub
                    End Try
                    'Windows 7
                    'AttachMentPath +
                ElseIf MyTest1.ShowDialog() = DialogResult.OK Then
                    Try

                        oDBs_Head.SetValue("U_VSPATCH", oDBs_Head.Offset, MyTest1.FileName)
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
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
    End Sub
#End Region

End Class

