Public Class VehicleStatus

#Region "Declaration"
    Dim VechstDT As New DataTable
    'Public DNum As String
    Public VSPDT1 As String
    Public PrintPDF As String = "No"
#End Region

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Try
            VechstDT.Columns.Add("VSPVNO")
            VechstDT.Columns.Add("VSPCNTR")
            VechstDT.Columns.Add("VSPCNM")
            VechstDT.Columns.Add("VSPDRIV")
            VechstDT.Columns.Add("VSPCLN")
            VechstDT.Columns.Add("VSPROUTE")
            VechstDT.Columns.Add("VSPTY")
            VechstDT.Columns.Add("VSPSTA")
            VechstDT.Columns.Add("VSPOPKM")
            VechstDT.Columns.Add("VSPCLKM")
            VechstDT.Columns.Add("VSPSOURC")
            VechstDT.Columns.Add("VSPDST")
            VechstDT.Columns.Add("VSPLOC")
            VechstDT.Columns.Add("VSPCHML")
            VechstDT.Columns.Add("VSPQUA", System.Type.GetType("System.Double"))
            VechstDT.Columns.Add("VSPCSVE")
            VechstDT.Columns.Add("VSPREM")
            VechstDT.Columns.Add("VSPTOTKM")
            VechstDT.Columns.Add("VSPDT", System.Type.GetType("System.DateTime"))

            'Dim VehicleStatus As String = "Select T0.U_VSPVNO,T0.U_VSPDT ,T0.U_VSPCNM ,T0.U_VSPCNTR ,T0.U_VSPDT ,T0.U_VSPROUTE ,T1.U_VSPTY,T0.U_VSPTRSHT ,T1.U_VSPSTA ,T1.U_VSPOPKM,T1.U_VSPCLKM ,T1.U_VSPSOURC, " & _
            '                              "T1.U_VSPDST,T1.U_VSPLOC,T1.U_VSPCHML,T1.U_VSPQUA,T1.U_VSPCSVE,T1.U_VSPREM,T1.U_VSPTOTKM  From [@VSP_VECHSTS] T0 Inner Join [@VSP_VECHSTS_C0] T1 " & _
            '                              "on T0.DocEntry = T1.DocENtry Where T0.U_VSPDT = '" & VSPDT1 & "' and T1.U_VSPTY <> ' '"

            Dim VehicleStatus As String = ""

            If objMain.IsSAPHANA = True Then
                VehicleStatus = "Select T0.""U_VSPVNO"",T0.""U_VSPDT"" ,T0.""U_VSPCNM"" ,T0.""U_VSPCNTR"" ,T0.""U_VSPDT"" ,T0.""U_VSPROUTE"" ,T1.""U_VSPTY"",T0.""U_VSPTRSHT"" ,T1.""U_VSPSTA"" ,T1.""U_VSPOPKM"",T1.""U_VSPCLKM"" ,T1.""U_VSPSOURC"", " & _
                                         "T1.""U_VSPDST"",T1.""U_VSPLOC"",T1.""U_VSPCHML"",T1.""U_VSPQUA"",T1.""U_VSPREM"",T1.""U_VSPTOTKM""  From ""@VSP_VECHSTS"" T0 Inner Join ""@VSP_VECHSTS_C0"" T1 " & _
                                         "on T0.""DocEntry"" = T1.""DocEntry"" Where T0.""U_VSPDT"" = '" & VSPDT1 & "' and T1.""U_VSPTY"" <> ' '"
            Else
                VehicleStatus = "Select T0.U_VSPVNO,T0.U_VSPDT ,T0.U_VSPCNM ,T0.U_VSPCNTR ,T0.U_VSPDT ,T0.U_VSPROUTE ,T1.U_VSPTY,T0.U_VSPTRSHT ,T1.U_VSPSTA ,T1.U_VSPOPKM,T1.U_VSPCLKM ,T1.U_VSPSOURC, " & _
                                         "T1.U_VSPDST,T1.U_VSPLOC,T1.U_VSPCHML,T1.U_VSPQUA,T1.U_VSPREM,T1.U_VSPTOTKM  From [@VSP_VECHSTS] T0 Inner Join [@VSP_VECHSTS_C0] T1 " & _
                                         "on T0.DocEntry = T1.DocENtry Where T0.U_VSPDT = '" & VSPDT1 & "' and T1.U_VSPTY <> ' '"
            End If

           
            Dim oRsVehicleStatus As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsVehicleStatus.DoQuery(VehicleStatus)

            Dim Head As Integer = 0

            If oRsVehicleStatus.RecordCount > 0 Then
                oRsVehicleStatus.MoveFirst()

                Dim DocDate As String = oRsVehicleStatus.Fields.Item("U_VSPDT").Value

                While Not oRsVehicleStatus.EoF
                    VechstDT.Rows.Add(1)
                    ' VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("DocNum") = oRsVehicleStatus.Fields.Item("DocEntry").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPDT") = oRsVehicleStatus.Fields.Item("U_VSPDT").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPVNO") = oRsVehicleStatus.Fields.Item("U_VSPVNO").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPCNTR") = oRsVehicleStatus.Fields.Item("U_VSPCNTR").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPCNM") = oRsVehicleStatus.Fields.Item("U_VSPCNM").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPDRIV") = oRsVehicleStatus.Fields.Item("U_VSPTRSHT").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPROUTE") = oRsVehicleStatus.Fields.Item("U_VSPROUTE").Value

                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPTY") = oRsVehicleStatus.Fields.Item("U_VSPTY").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPSTA") = oRsVehicleStatus.Fields.Item("U_VSPSTA").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPOPKM") = oRsVehicleStatus.Fields.Item("U_VSPOPKM").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPCLKM") = oRsVehicleStatus.Fields.Item("U_VSPCLKM").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPSOURC") = oRsVehicleStatus.Fields.Item("U_VSPSOURC").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPDST") = oRsVehicleStatus.Fields.Item("U_VSPDST").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPLOC") = oRsVehicleStatus.Fields.Item("U_VSPLOC").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPCHML") = oRsVehicleStatus.Fields.Item("U_VSPCHML").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPQUA") = oRsVehicleStatus.Fields.Item("U_VSPQUA").Value
                    'VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPCSVE") = oRsVehicleStatus.Fields.Item("U_VSPCSVE").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPREM") = oRsVehicleStatus.Fields.Item("U_VSPREM").Value
                    VechstDT.Rows.Item(VechstDT.Rows.Count - 1).Item("VSPTOTKM") = oRsVehicleStatus.Fields.Item("U_VSPTOTKM").Value

                    oRsVehicleStatus.MoveNext()
                End While
            End If

            If VechstDT.Rows.Count > 0 Then
                Dim CR As New RptVehicleStatus
                CR.Database.Tables.Item("VehicleStatus").SetDataSource(VechstDT)
                CrystalReportViewer1.ReportSource = CR
                CrystalReportViewer1.Refresh()
            End If

            If VechstDT.Rows.Count > 0 Then
                If PrintPDF = "No" Then
                    Dim CR As New RptVehicleStatus
                    CR.Database.Tables.Item("VehicleStatus").SetDataSource(VechstDT)
                    CrystalReportViewer1.ReportSource = CR
                    CrystalReportViewer1.Refresh()
                Else
                    Dim Date1 As String = VSPDT1
                    Date1 = Date1.Insert("4", "-")
                    Date1 = Date1.Insert("7", "-")

                    Dim CR As New RptVehicleStatus
                    CR.Database.Tables.Item("VehicleStatus").SetDataSource(VechstDT)
                    Dim CrExportOptions As CrystalDecisions.Shared.ExportOptions
                    Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
                    Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
                    CrDiskFileDestinationOptions.DiskFileName = "D:\VH " & Date1 & ".PDF"
                    CrExportOptions = CR.ExportOptions
                    With CrExportOptions
                        .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .DestinationOptions = CrDiskFileDestinationOptions
                        CrFormatTypeOptions.FirstPageNumber = 1
                        CrFormatTypeOptions.LastPageNumber = 100
                        CrFormatTypeOptions.UsePageRange = True
                        .FormatOptions = CrFormatTypeOptions
                        CR.Export()
                        Me.Close()
                    End With
                    System.Diagnostics.Process.Start(CrDiskFileDestinationOptions.DiskFileName)
                End If
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRsVehicleStatus)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
End Class
