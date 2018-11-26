Imports System.Threading
Imports System.IO
Public Class clsGeneralSettings

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim ButtonID As String
#End Region

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = False Then
                        Me.AddItems(objForm.UniqueID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If (pVal.ItemUID = "btn_VA" Or pVal.ItemUID = "btn_DA") And pVal.BeforeAction = False Then
                        ButtonID = pVal.ItemUID
                        Me.BrowseFileDialog()
                    End If

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub AddItems(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbl_VA", objForm.Items.Item("10000268").Top + 22, objForm.Items.Item("10000266").Left, objForm.Items.Item("10000266").Width, _
                                            "Vehicle Attachment", "10000266", 7, 7)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_VA", objForm.Items.Item("10000268").Top + 22, objForm.Items.Item("10000267").Left, objForm.Items.Item("10000267").Width, _
                                            "OADM", "U_VSPVHATC", "lbl_VA", 7, 7)
            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_VA", objForm.Items.Item("10000268").Top + 22, objForm.Items.Item("10000268").Left, objForm.Items.Item("10000268").Width, _
                                             "txt_VA", "...", 7, 7)

            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbl_DA", objForm.Items.Item("10000268").Top + 44, objForm.Items.Item("10000266").Left, objForm.Items.Item("10000266").Width, _
                                            "Driver Attachment", "10000266", 7, 7)
            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_DA", objForm.Items.Item("10000268").Top + 44, objForm.Items.Item("10000267").Left, objForm.Items.Item("10000267").Width, _
                                            "OADM", "U_VSPDRATC", "lbl_DA", 7, 7)
            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_DA", objForm.Items.Item("10000268").Top + 44, objForm.Items.Item("10000268").Left, objForm.Items.Item("10000268").Width, _
                                             "txt_DA", "...", 7, 7)

            objForm.Items.Item("txt_VA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("txt_DA").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

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
        Dim MyTest1 As New FolderBrowserDialog
        Dim MyProcs() As Process
        Try
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length <> 0 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)

                    'Windows XP
                    If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                        Try
                            If ButtonID = "btn_VA" Then
                                objForm.Items.Item("txt_VA").Specific.Value = MyTest1.SelectedPath
                            Else
                                objForm.Items.Item("txt_DA").Specific.Value = MyTest1.SelectedPath
                            End If
                            If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        Catch ex As IO.IOException
                            objMain.objApplication.MessageBox(ex.Message)
                            Exit Sub
                        End Try

                        'Windows 7
                    ElseIf MyTest1.ShowDialog() = DialogResult.OK Then
                        Try
                            If ButtonID = "btn_VA" Then
                                objForm.Items.Item("txt_VA").Specific.Value = MyTest1.SelectedPath
                            Else
                                objForm.Items.Item("txt_DA").Specific.Value = MyTest1.SelectedPath
                            End If
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
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

End Class
