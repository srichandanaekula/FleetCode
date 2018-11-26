
Imports System.Reflection
Public Class Utilities
    Public strLastErrorCode As String
    Public strLastError As String
    Private objForm As SAPbouiCOM.Form
    Dim oItem As SAPbouiCOM.Item

#Region " Get Application "
    Public Function GetApplication() As SAPbouiCOM.Application
        Dim objApp As SAPbouiCOM.Application
        Try
            Dim objSboGuiApi As New SAPbouiCOM.SboGuiApi
            Dim strConnectionString As String = String.Empty
            If strConnectionString = "" Then
                If Environment.GetCommandLineArgs().Length = 1 Then
                    strConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
                Else
                    strConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                End If
            End If
            objSboGuiApi = New SAPbouiCOM.SboGuiApi
            objSboGuiApi.Connect(strConnectionString)
            objApp = objSboGuiApi.GetApplication()
            Return objApp
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            System.Windows.Forms.Application.Exit()
            Return Nothing
        End Try
    End Function
#End Region

#Region " Get Company "
    Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
        Dim objCompany As SAPbobsCOM.Company

        Dim strCookie As String
        Dim strCookieContext As String

        Try
            objCompany = New SAPbobsCOM.Company
            strCookie = objCompany.GetContextCookie
            strCookieContext = SBOApplication.Company.GetConnectionContext(strCookie)
            objCompany.SetSboLoginContext(strCookieContext)
            If objCompany.Connect <> 0 Then
                SBOApplication.StatusBar.SetText("Connection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If
            Return objCompany
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
            Return Nothing
        End Try
    End Function
#End Region

#Region " Load Form "
    Public Sub LoadForm(ByVal XMLFile As String, ByVal FormType As String, Optional ByVal FileType As ResourceType = ResourceType.Content)
        Try
            Dim AppAssemblty As Assembly = Assembly.GetExecutingAssembly()
            Dim sExecutingAssemblyNmae As String = AppAssemblty.GetName().Name.ToString()
            Dim xmldoc As New Xml.XmlDocument
            XMLFile = sExecutingAssemblyNmae + "." + XMLFile
            Dim Streaming As System.IO.Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(XMLFile)
            Dim StreamRead As New System.IO.StreamReader(Streaming, True)
            xmldoc.LoadXml(StreamRead.ReadToEnd)
            StreamRead.Close()
            If Not xmldoc.SelectSingleNode("//form") Is Nothing Then
                xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value & "_" & objMain.objApplication.Forms.Count
                objMain.objApplication.LoadBatchActions(xmldoc.InnerXml)
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Create Table"
    Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objMain.objApplication.StatusBar.SetText("Creating table... [@" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                Else
                    'Vj Added for testing///////////////
                    Dim lret As Integer
                    Dim sret As String = String.Empty
                    objMain.objCompany.GetLastError(lret, sret)
                    objMain.objApplication.MessageBox(lret & " : " & sret)
                    '//////////////////Done
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Fields Creation"
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType


                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()

                    Next
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String, ByVal Mandetory As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If Mandetory.Trim.ToUpper = "YES" Then
                    objUserFieldMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    objUserFieldMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If

                'Else
                '    objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                '    objUserFieldMD.TableName = TableName
                '    objUserFieldMD.Name = ColumnName
                '    objUserFieldMD.Description = ColDescription
                '    objUserFieldMD.Type = FieldType

                '    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                '        objUserFieldMD.Size = Size
                '    Else
                '        objUserFieldMD.EditSize = Size
                '    End If
                '    objUserFieldMD.SubType = SubType
                '    objUserFieldMD.DefaultValue = SetValidValue
                '    For intLoop = 0 To strValue.GetLength(0) - 1
                '        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                '        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                '        objUserFieldMD.ValidValues.Add()
                '    Next
                '    If (objUserFieldMD.Update() <> 0) Then
                '        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                '    End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub

    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            Dim str As String = ""
            If objMain.IsSAPHANA = True Then
                str = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & TableName & "' AND ""AliasID"" = '" & ColumnName & "'"
            Else
                str = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" & TableName & "' AND AliasID = '" & ColumnName & "'"
            End If

            objRecordSet.DoQuery(str)
            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try

    End Function

    Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)

        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub
    Public Sub AddInteger(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal Size As Integer)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SubType, "", "", "")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Sub UniqueIDField(ByVal TableName As String, ByVal FieldName As String, ByVal oCompany As SAPbobsCOM.Company)

        '//****************************************************************************
        '// The UserKeysMD represents a meta-data object that allows you
        '// to add\remove user defined keys.
        '//****************************************************************************

        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        '//flag
        Dim bFlagFirst As Boolean

        bFlagFirst = True

        '//****************************************************************************
        '// In any meta-data operation there should be no other object "alive"
        '// but the meta-data object, otherwise the operation will fail.
        '// This restriction is intended to prevent collisions.
        '//****************************************************************************

        '// The meta-data object must be initialized with a
        '// regular UserKeys object
        oUserKeysMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

        '// Set the table name and the key name
        'oUserKeysMD.TableName = "OCRD" '// BP table
        'oUserKeysMD.KeyName = "BE_MyKey1"

        oUserKeysMD.TableName = TableName '// BP table
        oUserKeysMD.KeyName = FieldName


        '//*******************************************
        '// Add a column to a key button:
        '//-------------------------------------------
        '// To add an additional column to
        '// the key, an additional element must be
        '// created in the Elements collection.
        '// The Add method of the Elements collection
        '// must be used only as of the second element.

        '// Do not use the Add method for the first element
        If bFlagFirst = True Then
            bFlagFirst = False
        Else
            '// Add an item to the Elements collection
            oUserKeysMD.Elements.Add()
            strLastErrorCode = oCompany.GetLastErrorCode()
            strLastError = oCompany.GetLastErrorDescription()
        End If

        '// Set the column's alias
        oUserKeysMD.Elements.ColumnAlias = FieldName

        '// Determine whether the key is unique or not
        'oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES
        oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tNO

        '// Add the key
        oUserKeysMD.Add()
        strLastErrorCode = oCompany.GetLastErrorCode()
        strLastError = oCompany.GetLastErrorDescription()
        'If (oUserKeysMD <> DBNull.Value) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)

        'End If
    End Sub
    Public Sub AddLinkField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SubType, "", "", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub AddImageField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String)
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_Image, "", "", "")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Add Data to No Object Table"
    Public Function AddDataToNoObjectTable(ByVal TableName As String, ByVal Code As String, ByVal Name As String, Optional ByVal UDFName1 As String = "", Optional ByVal UDFValue1 As String = "")
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim lReturn As Integer
        Dim ErrorString As String
        oUserTable = objMain.objCompany.UserTables.Item(TableName)

        If oUserTable.GetByKey(Code) = False Then
            'Set default, mandatory fields
            oUserTable.Code = Code
            oUserTable.Name = Name

            'Set user field
            If UDFName1 <> String.Empty Then oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1

            oUserTable.Add()
            If lReturn <> 0 Then
                objMain.objCompany.GetLastError(lReturn, ErrorString)
                Return (ErrorString)
            End If
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
        Return ("")
    End Function
#End Region

#Region "add menus with xml"
    Public Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objMain.objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objMain.objApplication.GetLastBatchResults()

    End Sub
#End Region

#Region " Check if Form Exists - ## Not Used "
    Public Function FormExist(ByVal FormUID As String) As Boolean
        Dim intLoop As Integer

        For intLoop = objMain.objApplication.Forms.Count - 1 To 0 Step -1
            If Trim(FormUID) = Trim(objMain.objApplication.Forms.Item(intLoop).UniqueID) Then
                Return True
            End If
        Next
        Return False
    End Function
#End Region

#Region " Get MaxCode "
    Public Function getMaxCode(ByVal sTable As String) As String
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try

            If objMain.IsSAPHANA = True Then
                strSQL = "SELECT MAX(CAST(""Code"" AS int)) AS ""code"" FROM """ & sTable & """"
            Else
                strSQL = "SELECT MAX(CAST(Code AS int)) AS code FROM [" & sTable & "]"
            End If

            oRS.DoQuery(strSQL)
            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If
            sCode = MaxCode
            Return sCode

        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region " UDO Document Numbering "
    Public Function GetNextDocNum(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, Optional ByVal SeriesName As String = "Primary") As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If objMain.IsSAPHANA = True Then
            Str = "select ""Series"" from NNM1 where ""ObjectCode"" = '" & UDOName & "' and ""SeriesName"" = '" & SeriesName & "'"
        Else
            Str = "select Series from NNM1 where objectCode = '" & UDOName & "' and SeriesName = '" & SeriesName & "'"
        End If

        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = Objform.BusinessObject.GetNextSerialNumber(oRs.Fields.Item(0).Value, UDOName)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("DN: " + ex.Message)
        End Try
    End Function
#End Region

#Region " Load DataSource from DB "
    Public Sub RefreshDatasourceFromDB(ByVal FormUID As String, ByRef oDBs_Head As SAPbouiCOM.DBDataSource, ByVal ConditionAlias As String, ByVal ConditionValue As String)
        Try
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions = New SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            oCondition = oConditions.Add()
            oCondition.Alias = ConditionAlias
            oCondition.ComparedAlias = ConditionAlias
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = ConditionValue
            oDBs_Head.Query(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "      ***Load Values to ComboBoxes***       "
    Public Sub ComboBoxLoadValues(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                objCombo.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub MatrixComboBoxValues(ByVal oColumn As SAPbouiCOM.Column, ByVal QueryAsValueAndDescription As String)
        Try
            If (oColumn.ValidValues.Count <> 0) Then
                For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
                    Try
                        oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If
            If oColumn.ValidValues.Count = 0 Then
                Dim objRecSet1 As SAPbobsCOM.Recordset
                objRecSet1 = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet1.DoQuery(QueryAsValueAndDescription)
                objRecSet1.MoveFirst()
                oColumn.ValidValues.Add("", "")
                While Not objRecSet1.EoF
                    Try
                        oColumn.ValidValues.Add(objRecSet1.Fields.Item(0).Value, objRecSet1.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet1.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region " Checking DataType"
    Public Function IsAlpha(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            For i = 0 To str.Length - 1
                If Not Char.IsLetter(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsNumeric(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                Return False
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsFloat(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Substring(0, 1) = "." Then
                Return False
            End If
            If str.Contains(".") = False Then
                Return False
            Else
                str = str.Remove(str.LastIndexOfAny("."), 1)
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Function IsPercentage(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                str = str.Remove(str.LastIndexOfAny("."), 1)
                If str.Contains(".") = True Then
                    Return False
                End If
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function
#End Region

#Region " Adding Items To Forms"
    Public Sub AddLabel(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                        ByVal iWidth As Integer, ByVal iCaption As String, ByVal iLink As String, _
                        Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = iLink
        oItem.Specific.Caption = iCaption
    End Sub

    Public Sub AddEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                          ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, _
                          ByVal LinkTo As String, Optional ByVal iFromPane As Integer = 0, _
                          Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = 14
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddExtendedEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, _
                                  ByVal iLeft As Integer, ByVal iWidth As Integer, ByVal TableName As String, _
                                  ByVal UdFName As String, ByVal LinkTo As String, Optional ByVal iFromPane As Integer = 0, _
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = 80
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddComboBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal LinkTo As String, _
                           Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddFolder(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer, _
                         ByVal UdFName As String, ByVal Caption As String, ByVal AliasName As String, _
                         ByVal GroupItem As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        Dim oFolder As SAPbouiCOM.Folder
        oFolder = oItem.Specific
        oFolder.Caption = Caption
        oFolder.GroupWith(GroupItem)
    End Sub

    Public Sub AddButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer, _
                         ByVal LinkTo As String, ByVal Caption As String, Optional ByVal iFromPane As Integer = 0, _
                         Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        Dim btn As SAPbouiCOM.Button = objForm.Items.Item(ItemUID).Specific
        btn.Caption = Caption
    End Sub

    Public Sub AddLinkButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, _
                             ByVal iLeft As Integer, ByVal iLinkTo As String, ByVal LinkedObject As Integer, Optional ByVal iFromPane As Integer = 0, _
                            Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        Dim LinkBtn As SAPbouiCOM.LinkedButton
        LinkBtn = oItem.Specific
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = iLinkTo
        If LinkedObject <> 0 Then
            LinkBtn.LinkedObject = LinkedObject
        End If
    End Sub

    Public Sub AddCheckBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal iCaption As String, _
                           Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
        oItem.Specific.Caption = iCaption
    End Sub
#End Region

#Region "Conversions"
    Function RupeesToWord(ByVal MyNumber As String) As String
        Dim Temp As String
        Dim Rupees As String = String.Empty
        Dim Paisa As String = String.Empty
        Dim DecimalPlace As String = String.Empty
        Dim iCount As String = String.Empty
        Dim Hundred As String = String.Empty
        Dim Words As String = String.Empty

        Dim ValidateNumber As String = MyNumber

        Dim place(9) As String
        place(0) = " Thousand "
        place(2) = " Lakh "
        place(4) = " Crore "
        place(6) = " Hundred "
        place(8) = " Kharab "
        If ValidateNumber.Length > 9 Then
            If ValidateNumber.Length = 10 And ValidateNumber.Substring(1, ValidateNumber.Length - 1) = "0" Then
                place(4) = " Crore "
                place(6) = " Hundred Crore "
            ElseIf ValidateNumber.Length = 11 And (ValidateNumber.Substring(2, ValidateNumber.Length - 2) = "0") Then
                place(4) = " Crore "
                place(6) = " Hundred "
            Else
                place(4) = " Crore "
                place(6) = " Hundred "
            End If
        End If

        On Error Resume Next
        ' Convert MyNumber to a string, trimming extra spaces.
        MyNumber = Trim(Str(MyNumber))

        ' Find decimal place.
        DecimalPlace = InStr(MyNumber, ".")

        ' If we find decimal place...
        If DecimalPlace > 0 Then
            ' Convert Paisa
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paisa = " and " & ConvertTens(Temp) & " Paisa"

            ' Strip off paisa from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        '===============================================================
        Dim TM As String  ' If MyNumber between Rs.1 To 99 Only.
        TM = Right(MyNumber, 2)

        If Len(MyNumber) > 0 And Len(MyNumber) <= 2 Then
            If Len(TM) = 1 Then
                Words = ConvertDigit(TM)
                RupeesToWord = "Rupees " & Words & Paisa & " Only"

                Exit Function

            Else
                If Len(TM) = 2 Then
                    Words = ConvertTens(TM)
                    RupeesToWord = "Rupees " & Words & Paisa & " Only"
                    Exit Function

                End If
            End If
        End If
        '===============================================================


        ' Convert last 3 digits of MyNumber to ruppees in word.
        Hundred = ConvertHundreds(Right(MyNumber, 3))
        ' Strip off last three digits
        MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        iCount = 0
        Do While MyNumber <> ""
            'Strip last two digits
            Temp = Right(MyNumber, 2)
            If Len(MyNumber) = 1 Then


                If Trim(Words) = "Thousand" Or _
                Trim(Words) = "Lakh  Thousand" Or _
                Trim(Words) = "Lakh" Or _
                Trim(Words) = "Crore" Or _
                Trim(Words) = "Crore  Lakh  Thousand" Or _
                Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or _
                Trim(Words) = "Hundred" Or _
                Trim(Words) = "Kharab  Hundred  Crore  Lakh  Thousand" Or _
                Trim(Words) = "Kharab" Then

                    Words = ConvertDigit(Temp) & place(iCount)
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                Else

                    Words = ConvertDigit(Temp) & place(iCount) & Words
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                End If
            Else

                If Trim(Words) = "Thousand" Or _
                   Trim(Words) = "Lakh  Thousand" Or _
                   Trim(Words) = "Lakh" Or _
                   Trim(Words) = "Crore" Or _
                   Trim(Words) = "Crore  Lakh  Thousand" Or _
                   Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or _
                   Trim(Words) = "Hundred" Then


                    Words = ConvertTens(Temp) & place(iCount)


                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else

                    '=================================================================
                    ' if only Lakh, Crore, Arab, Kharab

                    If Trim(ConvertTens(Temp) & place(iCount)) = "Lakh" Or _
                       Trim(ConvertTens(Temp) & place(iCount)) = "Crore" Or _
                       Trim(ConvertTens(Temp) & place(iCount)) = "Hundred" Then

                        Words = Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                        Words = ConvertTens(Temp) & place(iCount) & Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    End If

                End If
            End If

            iCount = iCount + 2
        Loop

        RupeesToWord = "Rupees " & Words & Hundred & Paisa & " Only"

    End Function

    Private Function ConvertHundreds(ByVal MyNumber As String) As String
        Dim Result As String = String.Empty
        'Return String.Empty
        ' Exit if there is nothing to convert.
        If Val(MyNumber) = 0 Then
            Return Nothing
            'Exit Function
        End If



        ' Append leading zeros to number.
        MyNumber = Right("000" & MyNumber, 3)

        ' Do we have a hundreds place digit to convert?
        If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
        End If

        ' Do we have a tens place digit to convert?
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
        Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
        End If

        ConvertHundreds = Trim(Result)
    End Function

    Private Function ConvertTens(ByVal MyTens As String) As String
        Dim Result As String = String.Empty

        ' Is value between 10 and 19?
        If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
        End If

        ConvertTens = Result
    End Function

    Private Function ConvertDigit(ByVal MyDigit As String) As String
        Select Case Val(MyDigit)
            Case 1 : ConvertDigit = "One"
            Case 2 : ConvertDigit = "Two"
            Case 3 : ConvertDigit = "Three"
            Case 4 : ConvertDigit = "Four"
            Case 5 : ConvertDigit = "Five"
            Case 6 : ConvertDigit = "Six"
            Case 7 : ConvertDigit = "Seven"
            Case 8 : ConvertDigit = "Eight"
            Case 9 : ConvertDigit = "Nine"
            Case Else : ConvertDigit = ""
        End Select
    End Function
#End Region

#Region " LoadValidValues "
    Public Sub AddValidValue(ByVal FormUID As String, ByVal FormType As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetDocNum As String = ""
            If objMain.IsSAPHANA = True Then
                GetDocNum = "Select ""DocNum"" , ""U_VSPMATID"" , ""U_VSPITCL"" From ""@VSP_FLT_DDCS"" Where ""U_VSPFRMID"" = '" & FormType & "'  And ""U_VSPACTV"" = 'Y'"
            Else
                GetDocNum = "Select DocNum , U_VSPMATID , U_VSPITCL From [@VSP_FLT_DDCS] Where U_VSPFRMID = '" & FormType & "'  And U_VSPACTV = 'Y'"
            End If

            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocNum.DoQuery(GetDocNum)

            If oRsGetDocNum.RecordCount > 0 Then
                oRsGetDocNum.MoveFirst()
                For i As Integer = 1 To oRsGetDocNum.RecordCount
                    Try

                        Dim GetDetails As String = ""
                        If objMain.IsSAPHANA = True Then

                            GetDetails = "Select T1.""U_VSPVALUS"" , T1.""U_VSPDESC"" From ""@VSP_FLT_DDCS"" T0 Inner Join ""@VSP_FLT_DDCS_C0"" T1 On T0.""DocEntry"" = T1.""DocEntry"" " & _
                            "Where ""DocNum"" = '" & oRsGetDocNum.Fields.Item(0).Value & "' And ""U_VSPVALUS"" <> '' And ""U_VSPACTV"" = 'Y' "
                        Else
                            GetDetails = "Select T1.U_VSPVALUS , T1.U_VSPDESC From [@VSP_FLT_DDCS] T0 Inner Join [@VSP_FLT_DDCS_C0] T1 On T0.DocEntry = T1.DocEntry " & _
                            "Where DocNum = '" & oRsGetDocNum.Fields.Item(0).Value & "' And U_VSPVALUS <> '' And U_VSPACTV = 'Y' "
                        End If


                        Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDetails.DoQuery(GetDetails)

                        If oRsGetDetails.RecordCount > 0 Then
                            If oRsGetDocNum.Fields.Item(1).Value <> "" Then
                                oRsGetDetails.MoveFirst()

                                Dim objMatrix As SAPbouiCOM.Matrix
                                Dim oColumn As SAPbouiCOM.Column
                                objMatrix = objForm.Items.Item(oRsGetDocNum.Fields.Item("U_VSPMATID").Value).Specific
                                oColumn = objMatrix.Columns.Item(oRsGetDocNum.Fields.Item("U_VSPITCL").Value)

                                If (oColumn.ValidValues.Count <> 0) Then
                                    For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If

                                oColumn.ValidValues.Add("", "")
                                While Not oRsGetDetails.EoF
                                    Try
                                        oColumn.ValidValues.Add(oRsGetDetails.Fields.Item("U_VSPVALUS").Value, oRsGetDetails.Fields.Item("U_VSPDESC").Value)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While

                            Else
                                oRsGetDetails.MoveFirst()

                                Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item(oRsGetDocNum.Fields.Item("U_VSPITCL").Value).Specific

                                If (objCombo.ValidValues.Count <> 0) Then
                                    For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If
                                objCombo.ValidValues.Add("", "")
                                While Not oRsGetDetails.EoF
                                    Try
                                        objCombo.ValidValues.Add(oRsGetDetails.Fields.Item("U_VSPVALUS").Value, oRsGetDetails.Fields.Item("U_VSPDESC").Value)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    oRsGetDocNum.MoveNext()
                Next
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Function GetDocEntry(ByVal DocNum As String, ByVal TableName As String)
        Try
            Dim ReturnDocEntry As String = "Select DocEntry From [" & TableName & "]"
            Dim oRsReturnDocEntry As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsReturnDocEntry.DoQuery(ReturnDocEntry)

            Return oRsReturnDocEntry.Fields.Item(0).Value
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

End Class

Public Enum ResourceType
    Embeded
    Content
End Enum



