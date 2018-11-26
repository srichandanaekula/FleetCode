Public Class clsSelectDistributionRules

#Region "        Declaration        "
    Dim objForm, objBaseForm As SAPbouiCOM.Form
    Dim oGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oDBs_Head, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim oEditTextCol As SAPbouiCOM.EditTextColumn
    Dim sHeader, sFormName, sMatrixID, sTableName, sBaseFormUID As String
    Dim objMatrix1 As SAPbouiCOM.Matrix
    Dim iRow As Integer
#End Region

    Sub CreateForm(ByVal BaseFormUID As String, ByVal TableName As String, ByVal Header As String, ByVal FormName As String, ByVal MatrixID As String, ByVal Row As Integer)
        Try
            objMain.objUtilities.LoadForm("Select Distribution Rule.xml", "VSP_FLT_DISTRL_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_FLT_DISTRL_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            sHeader = Header
            sFormName = FormName
            sMatrixID = MatrixID
            iRow = Row
            sTableName = TableName
            sBaseFormUID = BaseFormUID
            objBaseForm = objMain.objApplication.Forms.Item(BaseFormUID)

            If Header = "Yes" Then
                oDBs_Head = objBaseForm.DataSources.DBDataSources.Item(sTableName)
            End If
            oDBs_Details1 = objBaseForm.DataSources.DBDataSources.Item(sTableName)
            objMain.GlobalFormUID = objForm.UniqueID
            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")
            Me.AddCFL(objForm.UniqueID)

            Me.LoadGrid(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.ItemUID = "OK" And pVal.BeforeAction = False Then
                        Try
                            Dim DIM1 As String = oGrid.DataTable.Columns.Item(1).Cells.Item(0).Value
                            Dim DIM2 As String = oGrid.DataTable.Columns.Item(1).Cells.Item(1).Value
                            Dim DIM3 As String = oGrid.DataTable.Columns.Item(1).Cells.Item(2).Value
                            Dim DIM4 As String = oGrid.DataTable.Columns.Item(1).Cells.Item(3).Value
                            Dim DIM5 As String = oGrid.DataTable.Columns.Item(1).Cells.Item(4).Value

                            If sHeader = "Yes" Then
                                oDBs_Head.SetValue("U_VSPDISTR", oDBs_Head.Offset, DIM1 & ";" & DIM2 & ";" & DIM3 & ";" & DIM4 & ";" & DIM5)
                                oDBs_Head.SetValue("U_VSPCC1", oDBs_Head.Offset, DIM1)
                                oDBs_Head.SetValue("U_VSPCC2", oDBs_Head.Offset, DIM2)
                                oDBs_Head.SetValue("U_VSPCC3", oDBs_Head.Offset, DIM3)
                                oDBs_Head.SetValue("U_VSPCC4", oDBs_Head.Offset, DIM4)
                                oDBs_Head.SetValue("U_VSPCC5", oDBs_Head.Offset, DIM5)
                            Else
                                Select Case sFormName
                                    Case "VSP_FLT_DRVRMSTR_Form"
                                        Select Case sMatrixID
                                            Case ""
                                                Me.DrvrMstrMatrix(iRow, sTableName, sMatrixID, DIM1, DIM2, DIM3, DIM4, DIM5)
                                        End Select
                                    Case "VSP_FLT_VMSTR_Form"
                                        Select Case sMatrixID
                                            Case ""
                                                Me.VechicalMstrMatrix(iRow, sTableName, sMatrixID, DIM1, DIM2, DIM3, DIM4, DIM5)
                                        End Select
                                    Case "VSP_FLT_TRSHT_Form"
                                        Me.TripSheetMatrix(iRow, sTableName, sMatrixID, DIM1, DIM2, DIM3, DIM4, DIM5)
                                End Select
                            End If
                            objForm.Close()
                        Catch ex As Exception
                            objMain.objApplication.StatusBar.SetText(ex.Message)
                        End Try
                    End If
                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                    If pVal.BeforeAction = True Then
                        oGrid = objForm.Items.Item("3").Specific
                        Dim GetDimCode As String = ""
                        If objMain.IsSAPHANA = True Then
                            GetDimCode = "select ""DimCode"" From ODIM Where ""DimDesc"" = " & _
                                                   "'" & oGrid.DataTable.Columns.Item(0).Cells.Item(pVal.Row).Value & "' And ""DimActive"" = 'Y' "
                        Else
                            GetDimCode = "select DimCode From ODIM Where DimDesc = " & _
                                                   "'" & oGrid.DataTable.Columns.Item(0).Cells.Item(pVal.Row).Value & "' And DimActive = 'Y' "

                        End If
                        Dim oRsGetDimCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDimCode.DoQuery(GetDimCode)
                        Me.CFLFilter(objForm.UniqueID, "CFL_DSTR", oRsGetDimCode.Fields.Item(0).Value.ToString)
                    End If

                    If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                        If oCFL.UniqueID = "CFL_DSTR" Then
                            oGrid = objForm.Items.Item("3").Specific
                            oGrid.DataTable.Columns.Item(1).Cells.Item(pVal.Row).Value = oDT.GetValue(0, 0)
                            oGrid.DataTable.Columns.Item(2).Cells.Item(pVal.Row).Value = oDT.GetValue(1, 0)
                        End If
                    End If
            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub LoadGrid(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oGrid = objForm.Items.Item("3").Specific

            oGrid.DataTable = oDt
            Dim Query As String = ""

            If objMain.IsSAPHANA = True Then
                Query = "Select ""DimDesc"" ,'              ' As ""Dist Rule Code"" , '             ' As " & _
                                  """Dist Rule Name"" from ODIM  "
            Else
                Query = "Select DimDesc ,'              ' As 'Dist Rule Code' , '             ' As " & _
                                  "'Dist Rule Name' from ODIM  "
            End If
            oGrid.DataTable.ExecuteQuery(Query)

            oEditTextCol = oGrid.Columns.Item(1)
            oEditTextCol.LinkedObjectType = 61
            oEditTextCol.ChooseFromListUID = "CFL_DSTR"
            oEditTextCol.ChooseFromListAlias = "OcrCode"

            oEditTextCol = oGrid.Columns.Item(2)

            oGrid.Columns.Item(0).Editable = False
            oGrid.Columns.Item(2).Editable = False

            oGrid.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AddCFL(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams

            oCFLs = objForm.ChooseFromLists
            oCFLCreationParams = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "62"
            oCFLCreationParams.UniqueID = "CFL_DSTR"
            oCFLCreationParams.MultiSelection = False

            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CFLFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal DimCode As String)
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

    Sub DrvrMstrMatrix(ByVal RowIndex As Integer, _
                ByVal MatTableName As String, ByVal MatID As String, _
                ByVal Dim1 As String, ByVal Dim2 As String, ByVal Dim3 As String, ByVal Dim4 As String, _
                ByVal Dim5 As String)
        Try
            
            objBaseForm = objMain.objApplication.Forms.Item(sBaseFormUID)
            objMatrix1 = objBaseForm.Items.Item(MatID).Specific
            oDBs_Details1 = objBaseForm.DataSources.DBDataSources.Item(MatTableName)

            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, RowIndex)
            oDBs_Details1.SetValue("U_VSPFDT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(RowIndex).Specific.Value)
            oDBs_Details1.SetValue("U_VSPTODT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(RowIndex).Specific.Value)
            oDBs_Details1.SetValue("U_VSPCNTCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(RowIndex).Specific.Value)
            oDBs_Details1.SetValue("U_VSPCNTNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(RowIndex).Specific.Value)
            oDBs_Details1.SetValue("U_VSPCNTDR", oDBs_Details1.Offset, Dim1 & ";" & Dim2 & ";" & Dim3 & ";" & Dim4 & ";" & Dim5)
            oDBs_Details1.SetValue("U_VSPCC1", oDBs_Details1.Offset, Dim1)
            oDBs_Details1.SetValue("U_VSPCC2", oDBs_Details1.Offset, Dim2)
            oDBs_Details1.SetValue("U_VSPCC3", oDBs_Details1.Offset, Dim3)
            oDBs_Details1.SetValue("U_VSPCC4", oDBs_Details1.Offset, Dim4)
            oDBs_Details1.SetValue("U_VSPCC5", oDBs_Details1.Offset, Dim5)
            objMatrix1.SetLineData(RowIndex)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub VechicalMstrMatrix(ByVal RowIndex As Integer, _
               ByVal MatTableName As String, ByVal MatID As String, _
               ByVal Dim1 As String, ByVal Dim2 As String, ByVal Dim3 As String, ByVal Dim4 As String, _
               ByVal Dim5 As String)
        Try

            objBaseForm = objMain.objApplication.Forms.Item(sBaseFormUID)
            objMatrix1 = objBaseForm.Items.Item(MatID).Specific
            oDBs_Details1 = objBaseForm.DataSources.DBDataSources.Item(MatTableName)

            oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, RowIndex)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, objMatrix1.Columns.Item("").Cells.Item(RowIndex).Specific.Value)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, Dim1 & ";" & Dim2 & ";" & Dim3 & ";" & Dim4 & ";" & Dim5)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, Dim1)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, Dim2)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, Dim3)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, Dim4)
            oDBs_Details1.SetValue("", oDBs_Details1.Offset, Dim5)
            objMatrix1.SetLineData(RowIndex)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub TripSheetMatrix(ByVal RowIndex As Integer, _
              ByVal MatTableName As String, ByVal MatID As String, _
              ByVal Dim1 As String, ByVal Dim2 As String, ByVal Dim3 As String, ByVal Dim4 As String, _
              ByVal Dim5 As String)
        Try

            objBaseForm = objMain.objApplication.Forms.Item(sBaseFormUID)
            objMatrix1 = objBaseForm.Items.Item(MatID).Specific
            oDBs_Details1 = objBaseForm.DataSources.DBDataSources.Item(MatTableName)

            Select Case MatID

                Case "27"
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, RowIndex)
                    oDBs_Details1.SetValue("U_VSPJRNDT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPAMTGD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPFRPLC", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPADAMT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPTOPLC", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPADVFM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPADVTO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_5").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPCSHNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPCMTS", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPCSTCN", oDBs_Details1.Offset, Dim1 & ";" & Dim2 & ";" & Dim3 & ";" & Dim4 & ";" & Dim5)
                    oDBs_Details1.SetValue("U_VSPCC1", oDBs_Details1.Offset, Dim1)
                    oDBs_Details1.SetValue("U_VSPCC2", oDBs_Details1.Offset, Dim2)
                    oDBs_Details1.SetValue("U_VSPCC3", oDBs_Details1.Offset, Dim3)
                    oDBs_Details1.SetValue("U_VSPCC4", oDBs_Details1.Offset, Dim4)
                    oDBs_Details1.SetValue("U_VSPCC5", oDBs_Details1.Offset, Dim5)
                    oDBs_Details1.SetValue("U_VSPJENO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_11").Cells.Item(RowIndex).Specific.Value)
                    objMatrix1.SetLineData(RowIndex)

                Case "28"
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, RowIndex)
                    oDBs_Details1.SetValue("U_VSPEXPAC", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPACTNM", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_1").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPACTCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_2").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPAMT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPCSTCN", oDBs_Details1.Offset, Dim1 & ";" & Dim2 & ";" & Dim3 & ";" & Dim4 & ";" & Dim5)
                    oDBs_Details1.SetValue("U_VSPCC1", oDBs_Details1.Offset, Dim1)
                    oDBs_Details1.SetValue("U_VSPCC2", oDBs_Details1.Offset, Dim2)
                    oDBs_Details1.SetValue("U_VSPCC3", oDBs_Details1.Offset, Dim3)
                    oDBs_Details1.SetValue("U_VSPCC4", oDBs_Details1.Offset, Dim4)
                    oDBs_Details1.SetValue("U_VSPCC5", oDBs_Details1.Offset, Dim5)
                    objMatrix1.SetLineData(RowIndex)

                Case "48"
                    oDBs_Details1.SetValue("LineId", oDBs_Details1.Offset, RowIndex)
                    oDBs_Details1.SetValue("U_VSPVENCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_0").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPDT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_9").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPITMCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_10").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPDESC", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_11").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPQTYIL", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_6").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPTXCD", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_8").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPRATE", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_3").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPAMT", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_4").Cells.Item(RowIndex).Specific.Value)
                    oDBs_Details1.SetValue("U_VSPCSTCN", oDBs_Details1.Offset, Dim1 & ";" & Dim2 & ";" & Dim3 & ";" & Dim4 & ";" & Dim5)
                    oDBs_Details1.SetValue("U_VSPCC1", oDBs_Details1.Offset, Dim1)
                    oDBs_Details1.SetValue("U_VSPCC2", oDBs_Details1.Offset, Dim2)
                    oDBs_Details1.SetValue("U_VSPCC3", oDBs_Details1.Offset, Dim3)
                    oDBs_Details1.SetValue("U_VSPCC4", oDBs_Details1.Offset, Dim4)
                    oDBs_Details1.SetValue("U_VSPCC5", oDBs_Details1.Offset, Dim5)
                    oDBs_Details1.SetValue("U_VSPGRPO", oDBs_Details1.Offset, objMatrix1.Columns.Item("V_7").Cells.Item(RowIndex).Specific.Value)
                    objMatrix1.SetLineData(RowIndex)
            End Select

            objMatrix1.AutoResizeColumns()

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
