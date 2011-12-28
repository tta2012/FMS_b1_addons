Option Explicit On

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices

Module modImportSeaLCL
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Private VocNo, VendorName, PayTo, PayType, BankName, CheqNo, Status, Currency, PaymentDate, GST, PrepBy, ExRate, Remark As String
    Private Total, GSTAmt, SubTotal, vocTotal, gstTotal As Double
    Private vendorCode As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sql As String = ""
    Dim strAPInvNo, strOutPayNo As String

    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public Function DoImportSeaLCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************Ite
        '   Function    :   DoImportSeaLCLFormDataEvent
        '   Purpose     :   This function will be providing to proceed validating for
        '                   Inventory [All] Form Data Event information
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.BusinessObjectInfo
        '                       pVal =  set the SAP UI Menu Event Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim FunctionName As String = "DoImportSeaLCLFormDataEvent"
        Dim sKeyValue As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSQLQuery As String = String.Empty
        Dim oDocument As SAPbobsCOM.Documents
        Dim oXmlReader As XmlTextReader
        Dim sDocNum As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oMatrix, oChMatrix As SAPbouiCOM.Matrix

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", FunctionName)
            Select Case BusinessObjectInfo.FormTypeEx
                'Voucher POP UP
                Case "LCLVOUCHER"
                    If BusinessObjectInfo.BeforeAction = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            Try
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                oPayForm = p_oSBOApplication.Forms.GetForm("LCLVOUCHER", 1)
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                If oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value = "" Then
                                    oChMatrix.DeleteRow(oChMatrix.RowCount)
                                End If
                                vocTotal = Convert.ToDouble(oPayForm.Items.Item("ed_Total").Specific.Value)
                                gstTotal = Convert.ToDouble(oPayForm.Items.Item("ed_GSTAmt").Specific.Value)
                                SaveToPurchaseVoucher(oPayForm, True)
                                SaveToDraftPurchaseVoucher(oPayForm)

                            Catch ex As Exception
                                BubbleEvent = False
                                MessageBox.Show(ex.Message)
                            End Try
                        End If
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE Then
                            Try
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                oPayForm = p_oSBOApplication.Forms.GetForm("LCLVOUCHER", 1)
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                If oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value = "" Then
                                    oChMatrix.DeleteRow(oChMatrix.RowCount)
                                End If
                                vocTotal = Convert.ToDouble(oPayForm.Items.Item("ed_Total").Specific.Value)
                                gstTotal = Convert.ToDouble(oPayForm.Items.Item("ed_GSTAmt").Specific.Value)
                                SaveToPurchaseVoucher(oPayForm, False)
                                SaveToDraftPurchaseVoucher(oPayForm)


                            Catch ex As Exception
                                BubbleEvent = False
                                MessageBox.Show(ex.Message)
                            End Try
                        End If
                    End If

                    If BusinessObjectInfo.ActionSuccess = True Then
                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                        oPayForm = p_oSBOApplication.Forms.GetForm("LCLVOUCHER", 1)
                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_LCL05_VOUCHER", True)
                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_LCL05_VOUCHER", False)
                        End If

                    End If

                Case "142"
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            Dim oImportSeaLCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            oDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            Dim sCode As String = String.Empty
                            Dim sName As String = String.Empty
                            Dim sAttention As String = String.Empty
                            Dim sPhone As String = String.Empty
                            Dim sFax As String = String.Empty
                            Dim sMail As String = String.Empty
                            sDocNum = BusinessObjectInfo.ObjectKey
                            oXmlReader = New XmlTextReader(New IO.StringReader(sDocNum))
                            While oXmlReader.Read()
                                If oXmlReader.NodeType = XmlNodeType.XmlDeclaration Then
                                    sDocNum = oXmlReader.ReadElementString
                                End If
                            End While
                            oXmlReader.Close()
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("SELECT CardCode,CardName FROM OPOR WHERE DocEntry = '" + sDocNum + "'")
                            If oRecordSet.RecordCount > 0 Then
                                sCode = oRecordSet.Fields.Item("CardCode").Value.ToString
                                sName = oRecordSet.Fields.Item("CardName").Value.ToString
                            End If
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select OCPR.Name,OCPR.Tel1,OCPR.Fax,OCPR.E_MailL from OCPR LEFT OUTER JOIN OCRD ON OCPR.Name = OCRD.CntctPrsn where OCRD.CardCode = '" + sCode + "'")
                            If oRecordSet.RecordCount > 0 Then
                                sAttention = oRecordSet.Fields.Item("Name").Value.ToString
                                sPhone = oRecordSet.Fields.Item("Tel1").Value.ToString
                                sFax = oRecordSet.Fields.Item("Fax").Value.ToString
                                sMail = oRecordSet.Fields.Item("E_MailL").Value.ToString
                            End If
                            oImportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = sDocNum
                            oEditText = oImportSeaLCLForm.Items.Item("ed_Trucker").Specific
                            oEditText.DataBind.SetBound(True, "", "TKREXTR")
                            oEditText.ChooseFromListUID = "CFLTKRV"
                            oEditText.ChooseFromListAlias = "CardName"
                            oImportSeaLCLForm.DataSources.UserDataSources.Item("TKREXTR").ValueEx = sName
                            oImportSeaLCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = sAttention
                            oImportSeaLCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = sPhone
                            oImportSeaLCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = sFax
                            oImportSeaLCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = sMail
                        End If
                    End If

                Case "IMPORTSEALCL"
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            LoadHolidayMarkUp(ImportSeaLCLForm)
                            ImportSeaLCLForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty 'MSW
                            If Not String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value) Then
                                If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            If Not String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_DspDate").Specific.Value) Then
                                If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoImportSeaLCLFormDataEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoImportSeaLCLFormDataEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()
        End Try
    End Function
    Dim dtmatrix As SAPbouiCOM.DataTable
    Public gridindex As String
    Public Function DoImportSeaLCLMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   ImportSeaLCLMenuEvent
        '   Purpose     :   This function will be providing to proceed validating for
        '                   Inventory [All] Menu Event information
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.MenuEvent
        '                       pVal =  set the SAP UI Menu Event Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' ***********************************************************************************
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoImportSeaLCLMenuEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            Select Case pVal.MenuUID
                Case "mnuImportSeaLCL"
                    If pVal.BeforeAction = False Then
                        'TODO: To ReArrange Loading SRF file
                        'TODO: To ReArrange Setting Up UserDataSources
                        'TODO: To ReArrange Adding and Setting Up ChooseFromList
                        'TODO: To ReArrange Loading Data from related table into related ComboBoxes
                        If Not LoadFromXML(p_oSBOApplication, "ImportSeaLCLv1.srf") Then Throw New ArgumentException(sErrDesc)
                        ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                        BubbleEvent = False
                        ImportSeaLCLForm.EnableMenu("1288", True)
                        ImportSeaLCLForm.EnableMenu("1289", True)
                        ImportSeaLCLForm.EnableMenu("1290", True)
                        ImportSeaLCLForm.EnableMenu("1291", True)
                        ImportSeaLCLForm.EnableMenu("1284", False)
                        ImportSeaLCLForm.EnableMenu("1286", False)

                        'ImportSeaLCLForm.DataBrowser.BrowseBy = "ed_JobNo"
                        ImportSeaLCLForm.DataBrowser.BrowseBy = "ed_DocNum"
                        ImportSeaLCLForm.Items.Item("fo_Prmt").Specific.Select()
                        ImportSeaLCLForm.Freeze(True)

                        '---------------------------------------------- Vouncher (OMM) 7-Mar-2011-------------------------------------------------------------------------'
                        Dim oColumn As SAPbouiCOM.Column
                        oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                        ImportSeaLCLForm.DataSources.DataTables.Add("DTCharges")
                        dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                        dtmatrix.Columns.Add("ChCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
                        dtmatrix.Columns.Add("AcCode", SAPbouiCOM.BoFieldsType.ft_Text)
                        dtmatrix.Columns.Add("Desc", SAPbouiCOM.BoFieldsType.ft_Text)
                        dtmatrix.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Price)
                        dtmatrix.Columns.Add("GST", SAPbouiCOM.BoFieldsType.ft_Text)
                        dtmatrix.Columns.Add("GSTAmt", SAPbouiCOM.BoFieldsType.ft_Price)
                        dtmatrix.Columns.Add("NoGST", SAPbouiCOM.BoFieldsType.ft_Price)
                        dtmatrix.Columns.Add("SeqNo", SAPbouiCOM.BoFieldsType.ft_Text)
                        dtmatrix.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text) 'To Add Item Code

                        oColumn = oMatrix.Columns.Item("colGST")
                        AddGSTComboData(oColumn)
                        oColumn.DataBind.Bind("DTCharges", "GST")
                        oColumn = oMatrix.Columns.Item("colChCode")
                        oColumn.DataBind.Bind("DTCharges", "ChCode")
                        ' AddUserDataSrc(oActiveForm, "test", "test", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 0)

                        'TO DO to drop the comment
                        'AddChooseFromList(ImportSeaLCLForm, "CCode", False, "UDOCHCODE")
                        'oColumn.ChooseFromListUID = "CCode"

                        oColumn = oMatrix.Columns.Item("colAcCode")
                        oColumn.DataBind.Bind("DTCharges", "AcCode")
                        oColumn = oMatrix.Columns.Item("colVDesc")
                        oColumn.DataBind.Bind("DTCharges", "Desc")
                        oColumn = oMatrix.Columns.Item("colAmount")
                        oColumn.DataBind.Bind("DTCharges", "Amount")
                        oColumn = oMatrix.Columns.Item("colGSTAmt")
                        oColumn.DataBind.Bind("DTCharges", "GSTAmt")
                        oColumn = oMatrix.Columns.Item("colNoGST")
                        oColumn.DataBind.Bind("DTCharges", "NoGST")
                        oColumn = oMatrix.Columns.Item("V_-1")
                        oColumn.DataBind.Bind("DTCharges", "SeqNo")
                        oColumn = oMatrix.Columns.Item("colICode")    'To Add Item Code
                        oColumn.DataBind.Bind("DTCharges", "ItemCode") 'To Add Item Code

                        dtmatrix.Rows.Add(1)
                        oMatrix.Clear()
                        dtmatrix.SetValue("SeqNo", 0, 1)
                        oMatrix.LoadFromDataSource()
                        '-------------------------------------Charge Code--------------------------------------------------------------'
                        'Vendor in Voucher
                        If AddChooseFromList(ImportSeaLCLForm, "PAYMENT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        ImportSeaLCLForm.Items.Item("ed_VedName").Specific.ChooseFromListUID = "PAYMENT"
                        ImportSeaLCLForm.Items.Item("ed_VedName").Specific.ChooseFromListAlias = "CardName"
                        ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                        'End

                        If AddUserDataSrc(ImportSeaLCLForm, "CASH", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "CHEQUE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        oOpt = ImportSeaLCLForm.Items.Item("op_Cash").Specific
                        oOpt.DataBind.SetBound(True, "", "CASH")
                        oOpt = ImportSeaLCLForm.Items.Item("op_Cheq").Specific
                        oOpt.DataBind.SetBound(True, "", "CHEQUE")
                        oOpt.GroupWith("op_Cash")
                        'OMM voucher
                        oCombo = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("Select BankName From ODSC")
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.MoveFirst()
                            While oRecordSet.EoF = False
                                oCombo.ValidValues.Add(oRecordSet.Fields.Item("BankName").Value, "")
                                oRecordSet.MoveNext()
                            End While
                        End If
                        oCombo = ImportSeaLCLForm.Items.Item("cb_PayCur").Specific
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("SELECT CurrCode FROM OCRN")
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.MoveFirst()
                            While oRecordSet.EoF = False
                                oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, "")
                                oRecordSet.MoveNext()
                            End While
                        End If
                        '------------------------------------------------------------------------------------------------------------------------------------------'
                        ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                        EnabledHeaderControls(ImportSeaLCLForm, False)
                        EnabledMaxtix(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("mx_TkrList").Specific, False)
                        ImportSeaLCLForm.PaneLevel = 7
                        ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Import Sea LCL"
                        ' ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN ").SetValue("U_TranMode", 0, "Sea") 'MSW LCL Change
                        ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                        ImportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                        ImportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        If AddChooseFromList(ImportSeaLCLForm, "cflBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddChooseFromList(ImportSeaLCLForm, "cflBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        ImportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListUID = "cflBP"
                        ImportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListAlias = "CardCode"
                        ImportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListUID = "cflBP2"
                        ImportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListAlias = "CardName"

                        If AddChooseFromList(ImportSeaLCLForm, "cflBP3", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddChooseFromList(ImportSeaLCLForm, "WRHSE", False, "WAREHOUSE") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddChooseFromList(ImportSeaLCLForm, "DSVES01", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddChooseFromList(ImportSeaLCLForm, "DSVES02", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                        ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                        ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListUID = "WRHSE"
                        ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListAlias = "Code"
                        ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                        ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                        ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                        ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                        '------------------------------- For Cargo Tab OMM & SYMA ------------------------------------------------'13 Jan 2011

                        AddChooseFromList(ImportSeaLCLForm, "cflCurCode", False, 37)
                        ImportSeaLCLForm.Items.Item("ed_CurCode").Specific.ChooseFromListUID = "cflCurCode"
                        '----------------------------------For Invoice Tab------------------------------------------------------'
                        AddChooseFromList(ImportSeaLCLForm, "cflCurCode1", False, 37)
                        oEditText = ImportSeaLCLForm.Items.Item("ed_CCharge").Specific
                        oEditText.ChooseFromListUID = "cflCurCode1"
                        AddChooseFromList(ImportSeaLCLForm, "cflCurCode2", False, 37)
                        oEditText = ImportSeaLCLForm.Items.Item("ed_Charge").Specific
                        oEditText.ChooseFromListUID = "cflCurCode2"
                        '-------------------------------------------------------------------------------------------------------'

                        oCombo = ImportSeaLCLForm.Items.Item("cb_PCode").Specific
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("SELECT Code, Name FROM [@OBT_TB004_PORTLIST]")
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.MoveFirst()
                            While oRecordSet.EoF = False
                                oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                                oRecordSet.MoveNext()
                            End While
                        End If

                        oCombo = ImportSeaLCLForm.Items.Item("cb_PType").Specific
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("SELECT PkgType FROM OPKG")
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.MoveFirst()
                            While oRecordSet.EoF = False
                                oCombo.ValidValues.Add(oRecordSet.Fields.Item("PkgType").Value, "")
                                oRecordSet.MoveNext()
                            End While
                        End If

                        oEditText = ImportSeaLCLForm.Items.Item("ed_JobNo").Specific
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEALCL'")
                        If oRecordSet.RecordCount > 0 Then
                            'ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW for Job Type Table
                            ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW for Job Type Table
                        End If

                        'fortruckingtab
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "TKREXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "DSINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "DSEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        oOpt = ImportSeaLCLForm.Items.Item("op_Inter").Specific
                        oOpt.DataBind.SetBound(True, "", "DSINTR")
                        oOpt = ImportSeaLCLForm.Items.Item("op_Exter").Specific
                        oOpt.DataBind.SetBound(True, "", "DSEXTR")
                        oOpt.GroupWith("op_Inter")

                        If AddUserDataSrc(ImportSeaLCLForm, "TKRDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "INSDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.DataBind.SetBound(True, "", "INSDATE")
                        ImportSeaLCLForm.Items.Item("ed_TkrDate").Specific.DataBind.SetBound(True, "", "TKRDATE")
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRATTE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRTEL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRFAX", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRMAIL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRCOL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddUserDataSrc(ImportSeaLCLForm, "TKRTO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        ImportSeaLCLForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
                        ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
                        ImportSeaLCLForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
                        ImportSeaLCLForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
                        ImportSeaLCLForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
                        ImportSeaLCLForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")

                        If AddUserDataSrc(ImportSeaLCLForm, "DSDISP", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.DataBind.SetBound(True, "", "DSDISP")
                        ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.DataBind.SetBound(True, "", "DSDISP")
                        ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.GroupWith("op_DspIntr")

                        If AddChooseFromList(ImportSeaLCLForm, "CFLTKRE", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        If AddChooseFromList(ImportSeaLCLForm, "CFLTKRV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        '---------------------------10-1-2011-------------------------------------
                        '----------Recordset for Binding colCType of Matrix (mx_Cont)-------------
                        '--------------------------SYMA & OMM-------------------------------------
                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
                        Dim oColum As SAPbouiCOM.Column
                        oColum = oMatrix.Columns.Item("colCType")
                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordSet.DoQuery("select U_Type  from [@OBT_LCL08_PCONTAINE] group by U_Type")
                        While oColum.ValidValues.Count > 0
                            oColum.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                        End While
                        If oRecordSet.RecordCount > 0 Then
                            Dim a As Integer = oRecordSet.RecordCount
                            oRecordSet.MoveFirst()
                            While oRecordSet.EoF = False
                                oColum.ValidValues.Add(oRecordSet.Fields.Item("U_Type").Value, " ")
                                oColum.DisplayDesc = False
                                oRecordSet.MoveNext()
                            End While
                        End If
                        oMatrix.AddRow()
                        oMatrix.Columns.Item("colCSeqNo").Cells.Item(1).Specific.Value = 1
                        '------------------------------For License Info SYMA & OMM (13/Jan/2011)---------------'
                        oMatrix = ImportSeaLCLForm.Items.Item("mx_License").Specific
                        oMatrix.AddRow()
                        oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
                        '--------------------------------------------------------------------------------------'
                        ImportSeaLCLForm.EnableMenu("1292", True)
                        ImportSeaLCLForm.EnableMenu("1293", True)
                        ImportSeaLCLForm.Freeze(False)
                        Select Case ImportSeaLCLForm.Mode
                            Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                ImportSeaLCLForm.Items.Item("bt_AddIns").Enabled = False
                                ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False
                                ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                            Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                ImportSeaLCLForm.Items.Item("bt_AddIns").Enabled = True
                                ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True
                                ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = True
                        End Select
                    End If

                Case "1281"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.Item("IMPORTSEALCL").Selected = True Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                            p_oSBOApplication.Forms.Item("IMPORTSEALCL").Items.Item("ed_JobNo").Enabled = True
                            p_oSBOApplication.Forms.Item("IMPORTSEALCL").Items.Item("ed_JobNo").Specific.Active = True
                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ImportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                                ImportSeaLCLForm.Items.Item("ed_Wrhse").Enabled = True
                            End If
                            If AddChooseFromListByOption(p_oSBOApplication.Forms.Item("IMPORTSEALCL"), True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                Case "1292"
                    If pVal.BeforeAction = False Then
                        'Export Voucher POP UP
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "LCLVOUCHER" Then
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                            If oMatrix.Columns.Item("colChCode1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                            End If
                            'Export Voucher POP UP
                        ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                            '-------------------------For Payment(omm)------------------------------------------'
                            'If (ImportSeaLCLForm.PaneLevel = 21) Then
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                            RowAddToMatrix(ImportSeaLCLForm, oMatrix)
                            'End If
                            '----------------------------------------------------------------------------------'
                        End If   
                    End If

                Case "1293"
                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "LCLVOUCHER" Then
                        oPayForm = p_oSBOApplication.Forms.ActiveForm
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                            If pVal.BeforeAction = True Then
                                If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                    BubbleEvent = False
                                End If

                                If BubbleEvent = True Then
                                    DeleteMatrixRow(oPayForm, oMatrix, "@OBT_LCL17_VDETAIL", "V_-1")
                                    BubbleEvent = False
                                    If Not oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    CalculateTotal(oPayForm, oMatrix)
                                End If
                            End If
                        Else
                            BubbleEvent = False
                        End If

                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Then
                        If pVal.BeforeAction = True Then
                            BubbleEvent = False
                        End If
                        If pVal.BeforeAction = False Then
                            '-------------------------For Payment(omm)------------------------------------------'
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                            dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                            If dtmatrix.Rows.Count > 0 Then
                                dtmatrix.Rows.Remove(gridindex - 1)
                            End If

                            'oMatrix.LoadFromDataSource()
                            '----------------------------------------------------------------------------------'
                        End If
                    End If
                    
                Case "1282"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.Item("IMPORTSEALCL").Selected = True Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                            EnabledHeaderControls(ImportSeaLCLForm, False) '25-3-2011
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEALCL'")
                            If oRecordSet.RecordCount > 0 Then
                                ' ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW for Job Type Table
                                ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString  'MSW for Job Type Table
                            End If
                            ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            ImportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                            ImportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ImportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                        End If
                    End If

                Case "1288", "1289", "1290", "1291"
                    If pVal.BeforeAction = True Then
                        If p_oSBOApplication.Forms.Item("IMPORTSEALCL").Selected = True Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If
                        End If
                    End If

                Case "EditVoc"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            LoadPaymentVoucher(ImportSeaLCLForm)
                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oPayForm.Items.Item("ed_DocNum").Visible = True
                            oPayForm.Items.Item("ed_DocNum").Enabled = True
                            oPayForm.Items.Item("ed_DocNum").Specific.Value = oMatrix.Columns.Item("colVDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString
                            oPayForm.Items.Item("cb_PayCur").Specific.Active = True
                            oPayForm.Items.Item("ed_DocNum").Visible = False
                            oPayForm.Items.Item("ed_DocNum").Enabled = False
                            oPayForm.DataBrowser.BrowseBy = "ed_DocNum"
                            oPayForm.Items.Item("ed_VedName").Enabled = False
                            oPayForm.Items.Item("ed_PayTo").Enabled = False
                            If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                            oPayForm.Items.Item("1").Click()
                        End If
                    End If



            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoImportSeaLCLMenuEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoImportSeaLCLMenuEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function
    '-------------------------For Payment(omm)------------------------------------------'
    Private Sub RowAddToMatrix(ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
        AddDataToDataTable(oActiveForm, oMatrix)
        dtmatrix.Rows.Add(1)
        If (oActiveForm.Items.Item("cb_GST").Specific.Value = "No") Then
            dtmatrix.SetValue("GST", dtmatrix.Rows.Count - 1, "None")
        End If
        dtmatrix.SetValue("SeqNo", dtmatrix.Rows.Count - 1, dtmatrix.Rows.Count)
        oMatrix.Clear()
        oMatrix.LoadFromDataSource()
    End Sub
    
    Private Sub AddDataToDataTable(ByVal oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
        Dim i As Integer = 0
        If dtmatrix.Rows.Count > 0 Then
            For i = 0 To dtmatrix.Rows.Count - 1
                dtmatrix.SetValue(0, i, oMatrix.Columns.Item("colChCode").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(1, i, oMatrix.Columns.Item("colAcCode").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(2, i, oMatrix.Columns.Item("colVDesc").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(3, i, oMatrix.Columns.Item("colAmount").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(4, i, oMatrix.Columns.Item("colGST").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(5, i, oMatrix.Columns.Item("colGSTAmt").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(6, i, oMatrix.Columns.Item("colNoGST").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(8, i, oMatrix.Columns.Item("colICode").Cells.Item(i + 1).Specific.Value) 'To Add Item Code
            Next
        End If

    End Sub
    Private Sub CalRate(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim RS As SAPbobsCOM.Recordset
        RS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
        RS.DoQuery("select Rate from ovtg where Code='" + oMatrix.Columns.Item("colGST1").Cells.Item(Row).Specific.Value + "'")
        Dim Rate As Double = 0.0
        Dim GSTAMT As Double = 0.0
        Dim NOGST As Double = 0.0
        If RS.RecordCount < 0 Then
            Rate = 0
        End If
        RS.MoveFirst()
        Rate = RS.Fields.Item("Rate").Value
        If Rate = 0 Then
            GSTAMT = 0
            NOGST = oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value
        Else
            GSTAMT = (Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) / 100.0) * Rate
            NOGST = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) - GSTAMT
        End If
       
        '====CalCulate Total======'
        oActiveForm.Freeze(True)
        oMatrix.Columns.Item("colGSTAmt").Editable = True
        oMatrix.Columns.Item("colNoGST").Editable = True
        oMatrix.Columns.Item("colGSTAmt").Cells.Item(Row).Specific.Value = Convert.ToString(GSTAMT)
        oMatrix.Columns.Item("colNoGST").Cells.Item(Row).Specific.Value = Convert.ToString(NOGST)
        oActiveForm.Items.Item("ed_VedName").Specific.Active = True
        oMatrix.Columns.Item("colGSTAmt").Editable = False
        oMatrix.Columns.Item("colNoGST").Editable = False
        CalculateTotal(oActiveForm, oMatrix)
        oActiveForm.Freeze(False)
    End Sub

    Private Sub CalculateTotal(ByRef oActiveForm As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix)
        Dim SubTotal As Double = 0.0
        Dim GSTTotal As Double = 0.0
        Dim Total As Double = 0.0
        For i As Integer = 1 To oMatrix.RowCount
            SubTotal = SubTotal + Convert.ToDouble(oMatrix.Columns.Item("colNoGST").Cells.Item(i).Specific.Value)
            GSTTotal = GSTTotal + Convert.ToDouble(oMatrix.Columns.Item("colGSTAmt").Cells.Item(i).Specific.Value)
        Next
        Total = SubTotal + GSTTotal
        oActiveForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_SubTotal", 0, SubTotal)
        oActiveForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_GSTAmt", 0, GSTTotal)
        oActiveForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_Total", 0, Total)
    End Sub

    Private Sub AddUpdateVoucher(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        '  ObjDBDataSource.Offset = 0

        rowIndex = pMatrix.GetNextSelectedRow

        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            If pMatrix.RowCount = 1 And pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                pMatrix.Clear()
            End If
        End If
        Try
            If ProcressedState = True Then
                'If ObjDBDataSource.GetValue("U_ConSeqNo", 0) = vbNullString Then pMatrix.Clear()
                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                        .SetValue("U_VSeqNo", .Offset, 1)
                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                        .SetValue("U_VSeqNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                    End If

                    .SetValue("U_RefNo", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value)
                    .SetValue("U_PVNo", .Offset, pForm.Items.Item("ed_VocNo").Specific.Value)
                    .SetValue("U_BPName", .Offset, pForm.Items.Item("ed_VedName").Specific.Value)
                    .SetValue("U_PayToAdd", .Offset, pForm.Items.Item("ed_PayTo").Specific.Value)
                    .SetValue("U_PayType", .Offset, IIf(pForm.Items.Item("op_Cash").Specific.Selected = True, "Cash", "Cheque").ToString)
                    .SetValue("U_BankName", .Offset, pForm.Items.Item("cb_BnkName").Specific.Value)
                    .SetValue("U_CheqNo", .Offset, pForm.Items.Item("ed_Cheque").Specific.Value)
                    .SetValue("U_Status", .Offset, "Draft")
                    .SetValue("U_CurCode", .Offset, pForm.Items.Item("cb_PayCur").Specific.Value)
                    .SetValue("U_PostDate", .Offset, pForm.Items.Item("ed_PosDate").Specific.Value)
                    .SetValue("U_GST", .Offset, pForm.Items.Item("cb_GST").Specific.Value)
                    .SetValue("U_Total", .Offset, pForm.Items.Item("ed_Total").Specific.Value)
                    .SetValue("U_PrepBy", .Offset, pForm.Items.Item("ed_VPrep").Specific.Value)
                    .SetValue("U_VDocNum", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource


                    .SetValue("LineId", .Offset, pForm.Items.Item("ed_VocNo").Specific.Value)
                    .SetValue("U_VSeqNo", .Offset, pForm.Items.Item("ed_VocNo").Specific.Value)
                    .SetValue("U_RefNo", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value)
                    .SetValue("U_PVNo", .Offset, pForm.Items.Item("ed_VocNo").Specific.Value)
                    .SetValue("U_BPName", .Offset, pForm.Items.Item("ed_VedName").Specific.Value)
                    .SetValue("U_PayToAdd", .Offset, pForm.Items.Item("ed_PayTo").Specific.Value)
                    .SetValue("U_PayType", .Offset, IIf(pForm.Items.Item("op_Cash").Specific.Selected = True, "Cash", "Cheque").ToString)
                    .SetValue("U_BankName", .Offset, pForm.Items.Item("cb_BnkName").Specific.Value)
                    .SetValue("U_CheqNo", .Offset, pForm.Items.Item("ed_Cheque").Specific.Value)
                    .SetValue("U_Status", .Offset, "Draft")
                    .SetValue("U_CurCode", .Offset, pForm.Items.Item("cb_PayCur").Specific.Value)
                    .SetValue("U_PostDate", .Offset, pForm.Items.Item("ed_PosDate").Specific.Value)
                    .SetValue("U_GST", .Offset, pForm.Items.Item("cb_GST").Specific.Value)
                    .SetValue("U_Total", .Offset, pForm.Items.Item("ed_Total").Specific.Value)
                    .SetValue("U_PrepBy", .Offset, pForm.Items.Item("ed_VPrep").Specific.Value)

                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub SetVoucherDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        Try
            pForm.Freeze(True)
            Dim oComboBank As SAPbouiCOM.ComboBox
            Dim oComboCurr As SAPbouiCOM.ComboBox
            Dim oComboGST As SAPbouiCOM.ComboBox
            oComboBank = pForm.Items.Item("cb_BnkName").Specific
            oComboCurr = pForm.Items.Item("cb_PayCur").Specific
            oComboGST = pForm.Items.Item("cb_GST").Specific

            oComboBank.Select(BankName, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oComboCurr.Select(Currency, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oComboGST.Select(GST, SAPbouiCOM.BoSearchKey.psk_ByValue)


            pForm.Items.Item("ed_PayTo").Specific.Value = PayTo
            pForm.Items.Item("ed_PayRate").Specific.Value = PayType

            If PayType = "Cash" Then
                pForm.DataSources.UserDataSources.Item("CASH").ValueEx = "1"
                pForm.DataSources.UserDataSources.Item("CHEQUE").ValueEx = "2"
            ElseIf PayType = "Cheque" Then
                pForm.DataSources.UserDataSources.Item("CHEQUE").ValueEx = "1"
                pForm.DataSources.UserDataSources.Item("CASH").ValueEx = "2"
            End If

            pForm.Items.Item("ed_VocNo").Specific.Value = VocNo
            pForm.Items.Item("ed_PosDate").Specific.Value = PaymentDate
            pForm.Items.Item("ed_PJobNo").Specific.Value = pForm.Items.Item("ed_JobNo").Specific.Value
            pForm.Items.Item("ed_VRemark").Specific.Value = Remark
            pForm.Items.Item("ed_VPrep").Specific.Value = PrepBy
            pForm.Items.Item("ed_SubTot").Specific.Value = SubTotal
            pForm.Items.Item("ed_GSTAmt").Specific.Value = GSTAmt
            pForm.Items.Item("ed_Total").Specific.Value = Total
            pForm.Items.Item("ed_Cheque").Specific.Value = CheqNo
            pForm.Items.Item("ed_PayRate").Specific.Value = ExRate
            pForm.Items.Item("ed_VedName").Specific.Value = VendorName
            If HolidayMarkUp(pForm, pForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, "") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            pForm.Freeze(False)

        Catch ex As Exception

        End Try
    End Sub
    Private Sub GetVoucherDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Try
            VocNo = pMatrix.Columns.Item("colVocNo").Cells.Item(Index).Specific.Value
            VendorName = pMatrix.Columns.Item("colVedName").Cells.Item(Index).Specific.Value
            PayTo = pMatrix.Columns.Item("colPayTo").Cells.Item(Index).Specific.Value
            PayType = pMatrix.Columns.Item("colPayType").Cells.Item(Index).Specific.Value
            BankName = pMatrix.Columns.Item("colBnkName").Cells.Item(Index).Specific.Value
            CheqNo = pMatrix.Columns.Item("colCheqNo").Cells.Item(Index).Specific.Value
            Status = pMatrix.Columns.Item("colStatus").Cells.Item(Index).Specific.Value
            Currency = pMatrix.Columns.Item("colCurCode").Cells.Item(Index).Specific.Value
            PaymentDate = pMatrix.Columns.Item("colPosDate").Cells.Item(Index).Specific.Value
            GST = pMatrix.Columns.Item("colGST").Cells.Item(Index).Specific.Value
            Total = Convert.ToDouble(pMatrix.Columns.Item("colTotal").Cells.Item(Index).Specific.Value)
            PrepBy = pMatrix.Columns.Item("colPrepBy").Cells.Item(Index).Specific.Value

            ExRate = pMatrix.Columns.Item("colExRate").Cells.Item(Index).Specific.Value
            GSTAmt = pMatrix.Columns.Item("colGSTAmt").Cells.Item(Index).Specific.Value
            SubTotal = pMatrix.Columns.Item("colSubTot").Cells.Item(Index).Specific.Value
            Remark = pMatrix.Columns.Item("colRemark").Cells.Item(Index).Specific.Value
        Catch ex As Exception

        End Try
    End Sub
    '----------------------------------------------------------------------------------'
    Public Function DoImportSeaLCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoImportSeaLCLItemEvent
        '   Purpose     :   This function will be providing to proceed validating for
        '                   Inventory [All] Menu Event information
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.ItemEvent
        '                       pVal =  set the SAP UI Menu Event Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oChMatrix As SAPbouiCOM.Matrix = Nothing
        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoImportSeaLCLItemEvent()"
        Dim sql As String = ""
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)

            Select Case pVal.FormTypeEx
                Case "LCLVOUCHER"
                    oPayForm = p_oSBOApplication.Forms.Item("LCLVOUCHER")

                    If pVal.BeforeAction = True Then
                        If pVal.ItemUID = "1" Then
                            Try
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                oPayForm = p_oSBOApplication.Forms.GetForm("LCLVOUCHER", 1)

                                If Not oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                    If oChMatrix.RowCount = 0 Then
                                        BubbleEvent = False
                                    End If
                                    If oChMatrix.RowCount = 1 And oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value = "" Then
                                        BubbleEvent = False
                                    End If

                                    If oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value = "" And BubbleEvent = True Then
                                        oChMatrix.DeleteRow(oChMatrix.RowCount)
                                    End If
                                End If

                            Catch ex As Exception
                                BubbleEvent = False
                                MessageBox.Show(ex.Message)
                            End Try
                        End If
                    End If
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount1" Then
                                CalRate(oPayForm, pVal.Row)
                            End If

                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "op_Cash" Then
                                oPayForm.Items.Item("cb_PayCur").Specific.Active = True
                                Dim oComboBank As SAPbouiCOM.ComboBox
                                oComboBank = oPayForm.Items.Item("cb_BnkName").Specific
                                For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                    oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                Next

                                oPayForm.Items.Item("cb_BnkName").Enabled = False
                                oPayForm.Items.Item("ed_Cheque").Enabled = False
                                oComboBank.Select(0, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                If Not oPayForm.Items.Item("ed_PayType").Specific.Value = "" Then
                                    oPayForm.Items.Item("ed_PayType").Specific.Value = "Cash"
                                End If
                                If Not oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oPayForm.Items.Item("ed_Cheque").Specific.Value = ""
                                End If
                            End If
                            If pVal.ItemUID = "op_Cheq" Then
                                Dim oComboBank As SAPbouiCOM.ComboBox
                                oPayForm.Items.Item("cb_BnkName").Specific.Active = True
                                oComboBank = oPayForm.Items.Item("cb_BnkName").Specific
                                If oComboBank.ValidValues.Count = 0 Then
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("Select BankName From ODSC")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oComboBank.ValidValues.Add(oRecordSet.Fields.Item("BankName").Value, "")
                                            oRecordSet.MoveNext()
                                        End While
                                    End If

                                End If
                                If Not oPayForm.Items.Item("ed_PayType").Specific.Value = "" Then
                                    oPayForm.Items.Item("ed_PayType").Specific.Value = "Cheque"
                                End If

                                oPayForm.Items.Item("cb_BnkName").Enabled = True
                                oPayForm.Items.Item("ed_Cheque").Enabled = True
                            End If
                            If pVal.ItemUID = "1" Then
                                If pVal.ActionSuccess = True Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)

                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "LCLVOUCHER" Then
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                                            If oPayForm.Items.Item("ed_PayType").Specific.Value = "Cash" Then
                                                oPayForm.Items.Item("op_Cash").Specific.Selected = True
                                                oPayForm.Items.Item("cb_BnkName").Enabled = False
                                                oPayForm.Items.Item("ed_Cheque").Enabled = False
                                            ElseIf oPayForm.Items.Item("ed_PayType").Specific.Value = "Cheque" Then
                                                oPayForm.Items.Item("op_Cheq").Specific.Selected = True
                                                oPayForm.Items.Item("cb_BnkName").Enabled = True
                                                oPayForm.Items.Item("ed_Cheque").Enabled = True
                                            End If
                                            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                            DisableChargeMatrix(oPayForm, oMatrix, False)
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            oPayForm.Items.Item("ed_PayTo").Enabled = False

                                        End If
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            sql = "Update [@OBT_LCL16_VHEADER] set U_APInvNo=" & Convert.ToInt32(strAPInvNo) & ",U_OutPayNo=" & Convert.ToInt32(strOutPayNo) & "" & _
                                            " ,U_FrDocNo=" & oPayForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oPayForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            oPayForm.Close()

                                            ImportSeaLCLForm.Items.Item("1").Click()
                                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If

                                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            oPayForm.Items.Item("ed_PayTo").Enabled = False
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "ed_VedName" Then
                                ObjDBDataSource = oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER") 'MSW To Add 18-3-2011
                                oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                                                                                       & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString()) 'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*

                                vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft
                                oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_BPCode", ObjDBDataSource.Offset, oDataTable.GetValue(0, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                            End If

                            If pVal.ColUID = "colChCode1" Then
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                Try
                                    oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                Catch ex As Exception

                                End Try
                                oChMatrix.Columns.Item("colAcCode").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_PAccCode").Cells.Item(0).Value.ToString
                                oChMatrix.Columns.Item("colICode").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_ItemCode").Cells.Item(0).Value.ToString

                                If oChMatrix.RowCount > 1 Then
                                    If oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value <> "" Then
                                        If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                                    End If
                                Else
                                    If oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                        If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                                    End If
                                End If

                            End If

                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            '-------------------------For Payment(omm)------------------------------------------'
                            If (pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colGST1") Then
                                CalRate(oPayForm, pVal.Row)
                            End If
                            'If (pVal.ItemUID = "cb_GST" And oPayForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft") Then
                            If (pVal.ItemUID = "cb_GST") And oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                ' dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
                                'oChMatrix.Columns.Item("colGST").Cells.Item(1).Specific.Value = "None"

                                Dim oCombo As SAPbouiCOM.ComboBox
                                oCombo = oChMatrix.Columns.Item("colGST1").Cells.Item(oChMatrix.RowCount).Specific
                                oCombo.Select("None", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                'dtmatrix.SetValue("GST", 0, "None")
                                'oMatrix.LoadFromDataSource()
                            End If
                            '----------------------------------------------------------------------------------'


                            If pVal.ItemUID = "cb_PayCur" Then
                                If oPayForm.Items.Item("cb_PayCur").Specific.Value <> "SGD" Then
                                    Dim Rate As String = String.Empty
                                    sql = "SELECT Rate FROM ORTT WHERE Currency = '" & oPayForm.Items.Item("cb_PayCur").Specific.Value & "' And DATENAME(YYYY,RateDate) = '" & _
                                            Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                            CInt(Today.ToString("dd"))
                                    oRecordSet.DoQuery(sql)
                                    If oRecordSet.RecordCount > 0 Then
                                        Rate = oRecordSet.Fields.Item("Rate").Value
                                    End If
                                    oPayForm.Items.Item("ed_PayRate").Enabled = True
                                    oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_ExRate", 0, Rate.ToString)
                                Else
                                    oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_ExRate", 0, Nothing)
                                    oPayForm.Items.Item("ed_PayRate").Enabled = False
                                    oPayForm.Items.Item("ed_PosDate").Specific.Active = True
                                End If
                            End If
                        End If
                    End If

                Case "IMPORTSEALCL"
                    ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")

                    'MSW for Job Type Table
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ItemUID = "ed_JobNo" Then

                                Dim str As String = ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value

                                Dim strErr As Integer
                                If str.Length <> 12 Then
                                    strErr = 1
                                Else
                                    Dim checkchar As String = Right(str, 6)
                                    Dim jobMode As String = Left(str, 2)
                                    Dim curYear As String = str.Substring(2, 4)
                                    Dim i As Integer = 0
                                    If jobMode.ToUpper <> "IM" Or curYear <> Now.Year.ToString() Then
                                        strErr = 2
                                    ElseIf Not IsNumeric(checkchar) Then
                                        strErr = 3
                                    Else
                                        sql = "select COUNT(*) count from [@OBT_TB002_IMPSEALCL] WHERE U_JobNum='" & str & "'"
                                        oRecordSet.DoQuery(sql)
                                        i = oRecordSet.Fields.Item("count").Value
                                        If i > 0 Then
                                            strErr = 4
                                        Else
                                            sql = "select COUNT(*) count from [@OBT_LCL01_IMPSEALCL] WHERE U_JobNum='" & str & "'"
                                            oRecordSet.DoQuery(sql)
                                            i = oRecordSet.Fields.Item("count").Value
                                            If i > 0 Then
                                                strErr = 4
                                            End If
                                        End If
                                    End If
                                End If

                                Select Case strErr
                                    Case 1
                                        ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Invalid Job Number.Job Number must be IM2011xxxxxx", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Case 2
                                        ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Invalid Job Number.Prefix must be ""IM and Current Year: IM2011 """, SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Case 3
                                        ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Invalid Job Number.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Case 4
                                        ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Job Number is already exist.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                End Select

                            End If
                        End If


                    End If

                    '-------------------------For Payment(omm)------------------------------------------'
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.Before_Action = False Then

                        If (pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colGST") Then
                            CalRate(ImportSeaLCLForm, pVal.Row)
                        End If
                        If (pVal.ItemUID = "cb_GST" And ImportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft") Then
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                            dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                            dtmatrix.SetValue("GST", 0, "None")
                            oMatrix.LoadFromDataSource()
                        End If

                        If pVal.ItemUID = "cb_PayCur" Then
                            If ImportSeaLCLForm.Items.Item("cb_PayCur").Specific.Value <> "SGD" Then
                                Dim Rate As String = String.Empty
                                sql = "SELECT Rate FROM ORTT WHERE Currency = '" & ImportSeaLCLForm.Items.Item("cb_PayCur").Specific.Value & "' And DATENAME(YYYY,RateDate) = '" & _
                                        Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                        CInt(Today.ToString("dd"))
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    Rate = oRecordSet.Fields.Item("Rate").Value
                                End If
                                ImportSeaLCLForm.Items.Item("ed_PayRate").Enabled = True
                                ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER").SetValue("U_ExRate", 0, Rate.ToString) ' * Change 
                            Else
                                ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER").SetValue("U_ExRate", 0, Nothing)       '* change
                                ImportSeaLCLForm.Items.Item("ed_PayRate").Enabled = False
                                ImportSeaLCLForm.Items.Item("cb_PayCur").Specific.Active = True
                            End If

                        End If
                    End If
                    '-----------------------------------------------------------------------------------'

                    If pVal.BeforeAction = False Then

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                            If Not RemoveFromAppList(ImportSeaLCLForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                        End If

                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                            Try
                                ImportSeaLCLForm.Items.Item("ed_Code").Enabled = False
                                ImportSeaLCLForm.Items.Item("ed_Name").Enabled = False
                                ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False 'MSW for Job Type Table
                            Catch ex As Exception
                            End Try
                        End If

                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            EnabledTrucker(ImportSeaLCLForm, False)
                        End If

                        '-------------------------For Payment(omm)------------------------------------------'
                        If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            CalRate(ImportSeaLCLForm, pVal.Row)
                        End If
                        '----------------------------------------------------------------------------------'

                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "fo_Prmt"
                                        ImportSeaLCLForm.PaneLevel = 7
                                        ImportSeaLCLForm.Items.Item("fo_PMain").Specific.Select()
                                    Case "fo_Dsptch"
                                        ImportSeaLCLForm.PaneLevel = 4
                                        ImportSeaLCLForm.Items.Item("ed_DspDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ImportSeaLCLForm.Items.Item("ed_DspDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        ImportSeaLCLForm.Items.Item("ed_DspHr").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                                        ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                    Case "fo_Trkng"
                                        ImportSeaLCLForm.PaneLevel = 6
                                        ImportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                            Else
                                                ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                            End If
                                        End If

                                    Case ("fo_Wrhse")
                                        ImportSeaLCLForm.PaneLevel = 3
                                    Case "fo_Vchr"
                                        ImportSeaLCLForm.PaneLevel = 20
                                        ImportSeaLCLForm.Items.Item("fo_VoView").Specific.Select()
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            Else
                                                ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                    Case "fo_VoView"
                                        ImportSeaLCLForm.PaneLevel = 20
                                    Case "fo_VoEdit"
                                        ImportSeaLCLForm.PaneLevel = 21
                                        ImportSeaLCLForm.Items.Item("cb_BnkName").Enabled = False
                                        ImportSeaLCLForm.Items.Item("ed_Cheque").Enabled = False
                                        ImportSeaLCLForm.Items.Item("ed_PayRate").Enabled = False
                                    Case "fo_TkrView"
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If ImportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If (oMatrix.Columns.Item("colInsDoc").Cells.Item(1).Specific.Value = "") Then
                                                ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                            Else
                                                ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                            End If
                                        End If

                                        ImportSeaLCLForm.PaneLevel = 6
                                    Case "fo_TkrEdit"
                                        ImportSeaLCLForm.PaneLevel = 5
                                        ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        If ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then
                                            SqlQuery = "SELECT U_WAdLine1,U_WAdLine2,U_WAdLine3,U_WState,U_WPostal,U_WCountry FROM [@OBT_TB003_WRHSE] WHERE Name = " & FormatString(ImportSeaLCLForm.Items.Item("ed_WName").Specific.Value.ToString) '* need to change [@OBT_TB003_WRHSE] Nyan Lin
                                            oRecordSet.DoQuery(SqlQuery)
                                            If oRecordSet.RecordCount > 0 Then
                                                ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = oRecordSet.Fields.Item("U_WAdLine1").Value.ToString
                                            End If
                                            oRecordSet.DoQuery("SELECT Address FROM OCRD WHERE CardCode = '" & ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                            If oRecordSet.RecordCount > 0 Then
                                                ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("Address").Value.ToString
                                            End If

                                            ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            ImportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True
                                        End If

                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = "" Then
                                            ImportSeaLCLForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty
                                            If (oMatrix.RowCount > 0) Then
                                                If (oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                    ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                                Else
                                                    ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                                End If
                                            Else
                                                ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                            End If
                                        End If


                                    Case "fo_PMain"
                                        ImportSeaLCLForm.PaneLevel = 8
                                    Case "fo_PCargo"
                                        ImportSeaLCLForm.PaneLevel = 9
                                    Case "fo_PCon"
                                        ImportSeaLCLForm.PaneLevel = 10
                                    Case "fo_PInv"
                                        ImportSeaLCLForm.PaneLevel = 11
                                    Case "fo_PLic"
                                        ImportSeaLCLForm.PaneLevel = 12
                                    Case "fo_PAttach"
                                        ImportSeaLCLForm.PaneLevel = 13
                                    Case "fo_PTotal"
                                        ImportSeaLCLForm.PaneLevel = 14

                                End Select

                                If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    ImportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                                    'ImportSeaLCLForm.Items.Item("ed_Wrhse").Enabled = True
                                End If

                                If pVal.ItemUID = "ch_POD" Then
                                    If ImportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = True Then
                                        ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
                                    End If
                                    If ImportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = False Then
                                        ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AddLic" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_License").Specific
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL10_PLICINFO").Clear()                          '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("colLicNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                End If

                                If pVal.ItemUID = "bt_DelLic" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_License").Specific
                                    Dim lRow As Long
                                    lRow = oMatrix.GetNextSelectedRow
                                    If lRow > -1 Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL10_PLICINFO").RemoveRecord(lRow - 1)   '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("@OBT_LCL10_PLICINFO")                                '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                            oUserTable.GetByKey(lRow)
                                            oUserTable.Remove()
                                            oUserTable = Nothing
                                        End If
                                        oMatrix.DeleteRow(lRow)
                                        SetMatrixSeqNo(oMatrix, "colLicNo")
                                        If lRow = 1 And oMatrix.RowCount = 0 Then
                                            oMatrix.FlushToDataSource()
                                            oMatrix.AddRow()
                                            oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
                                        End If
                                        oMatrix.FlushToDataSource()
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AddCon" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL08_PCONTAINE").Clear()                         '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("colCSeqNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                End If

                                If pVal.ItemUID = "bt_DelCon" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
                                    Dim lRow As Long
                                    lRow = oMatrix.GetNextSelectedRow
                                    If lRow > -1 Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL08_PCONTAINE").RemoveRecord(lRow - 1)    '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                            'oMatrix.AddRow(1)
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("OBT_LCL08_PCONTAINE")                                  '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                            'oMatrix.AddRow(1)
                                            oUserTable.GetByKey(lRow)
                                            oUserTable.Remove()
                                            oUserTable = Nothing
                                        End If
                                        oMatrix.DeleteRow(lRow)
                                        SetMatrixSeqNo(oMatrix, "colCSeqNo")
                                        If lRow = 1 And oMatrix.RowCount = 0 Then
                                            oMatrix.FlushToDataSource()
                                            oMatrix.AddRow()
                                            oMatrix.Columns.Item("colCSeqNo").Cells.Item(1).Specific.Value = 1
                                        End If
                                        oMatrix.FlushToDataSource()
                                    End If
                                End If

                                If pVal.ItemUID = "ch_CnUn" Then
                                    If ImportSeaLCLForm.Items.Item("ch_CnUn").Specific.Checked = True Then
                                        ImportSeaLCLForm.Freeze(True)
                                        ImportSeaLCLForm.Items.Item("ed_CnDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ImportSeaLCLForm.Items.Item("ed_CnDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        ImportSeaLCLForm.Items.Item("ed_CnHr").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ImportSeaLCLForm.Items.Item("ed_CnHr").Specific.Active = False
                                        ImportSeaLCLForm.Freeze(False)
                                    ElseIf ImportSeaLCLForm.Items.Item("ch_CnUn").Specific.Checked = False Then
                                        ImportSeaLCLForm.Freeze(True)
                                        ImportSeaLCLForm.Items.Item("ed_CnDate").Specific.Value = vbNullString
                                        ImportSeaLCLForm.Items.Item("ed_CnDay").Specific.Value = vbNullString
                                        ImportSeaLCLForm.Items.Item("ed_CnHr").Specific.Value = vbNullString
                                        ImportSeaLCLForm.Items.Item("ed_CnHr").Specific.Active = False
                                        ImportSeaLCLForm.Freeze(False)
                                    End If
                                End If

                                If pVal.ItemUID = "ch_CgCl" Then
                                    If ImportSeaLCLForm.Items.Item("ch_CgCl").Specific.Checked = True Then
                                        ImportSeaLCLForm.Freeze(True)
                                        ImportSeaLCLForm.Items.Item("ed_CgDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ImportSeaLCLForm.Items.Item("ed_CgDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        ImportSeaLCLForm.Items.Item("ed_CgHr").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ImportSeaLCLForm.Items.Item("ed_CgHr").Specific.Active = False
                                        ImportSeaLCLForm.Freeze(False)
                                    ElseIf ImportSeaLCLForm.Items.Item("ch_CgCl").Specific.Checked = False Then
                                        ImportSeaLCLForm.Freeze(True)
                                        ImportSeaLCLForm.Items.Item("ed_CgDate").Specific.Value = vbNullString
                                        ImportSeaLCLForm.Items.Item("ed_CgDay").Specific.Value = vbNullString
                                        ImportSeaLCLForm.Items.Item("ed_CgHr").Specific.Value = vbNullString
                                        ImportSeaLCLForm.Items.Item("ed_CgHr").Specific.Active = False
                                        ImportSeaLCLForm.Freeze(False)
                                    End If
                                End If

                                If pVal.ItemUID = "ch_Dsp" Then
                                    If ImportSeaLCLForm.Items.Item("ch_Dsp").Specific.Checked = True Then
                                        ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If ImportSeaLCLForm.Items.Item("ch_Dsp").Specific.Checked = False Then
                                        ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific.Active = True


                                    End If
                                End If

                                If pVal.ItemUID = "ed_ETADate" And ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_LCDDate" And ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific, ImportSeaLCLForm.Items.Item("ed_LCDDay").Specific, ImportSeaLCLForm.Items.Item("ed_LCDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_LCDDay").Specific, ImportSeaLCLForm.Items.Item("ed_LCDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_CnDate" And ImportSeaLCLForm.Items.Item("ed_CnDate").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_CgDate" And ImportSeaLCLForm.Items.Item("ed_CgDate").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_JbDate" And ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspDate" And ImportSeaLCLForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspCDte" And ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                                    If DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "bt_Payment" Then
                                    p_oSBOApplication.ActivateMenuItem("2818")
                                End If

                                If pVal.ItemUID = "op_Inter" Then
                                    ImportSeaLCLForm.Freeze(True)
                                    ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                    ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                    If AddChooseFromListByOption(ImportSeaLCLForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                    ImportSeaLCLForm.Freeze(False)
                                ElseIf pVal.ItemUID = "op_Exter" Then
                                    ImportSeaLCLForm.Freeze(True)
                                    ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = True
                                    If AddChooseFromListByOption(ImportSeaLCLForm, False, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Active = True
                                    ImportSeaLCLForm.Freeze(False)
                                End If

                                If pVal.ItemUID = "op_DspIntr" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If ClearComboData(ImportSeaLCLForm, "cb_Dspchr", "@OBT_LCL04_DISPATCH", "U_Dispatch") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT empID,lastName,firstName,middleName FROM OHEM")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        Do Until oRecordSet.EoF
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("firstName").Value.ToString & " " & _
                                                                   oRecordSet.Fields.Item("middleName").Value.ToString & " " & _
                                                                   oRecordSet.Fields.Item("lastName").Value.ToString, "")
                                            oRecordSet.MoveNext()
                                        Loop
                                    End If
                                ElseIf pVal.ItemUID = "op_DspExtr" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If ClearComboData(ImportSeaLCLForm, "cb_Dspchr", "@OBT_LCL04_DISPATCH", "U_Dispatch") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT CardCode,CardName FROM OCRD WHERE CardType = 'S'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        Do Until oRecordSet.EoF
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("CardName").Value.ToString, "")
                                            oRecordSet.MoveNext()
                                        Loop
                                    End If
                                End If

                                If pVal.ItemUID = "bt_GenPO" Then
                                    p_oSBOApplication.Menus.Item("2305").Activate()
                                    'SBO_Application.Menus.Item("6913").Activate()
                                    p_oSBOApplication.ActivateMenuItem("6913") 'MSW 04-04-2011
                                    p_oSBOApplication.Menus.Item("6913").Activate()

                                    Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                                    UDFAttachForm.Items.Item("U_JobNo").Specific.Value = ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    UDFAttachForm.Items.Item("U_InsDate").Specific.Value = ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value
                                End If

                                If pVal.ItemUID = "1" Then
                                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If pVal.ActionSuccess = True Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("ed_Code").Enabled = False 'MSW
                                            ImportSeaLCLForm.Items.Item("ed_Name").Enabled = False
                                            ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False 'MSW For Job Type Table
                                            Dim JobLastDocEntry As Integer
                                            Dim ObjectCode As String = String.Empty

                                            sql = "select top 1 Docentry from [@OBT_LCL01_IMPSEALCL] order by docentry desc"
                                            oRecordSet.DoQuery(sql)
                                            Dim FrDocEntry As Integer = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString)

                                            sql = "select top 1 Docentry from [@OBT_FREIGHTDOCNO] order by docentry desc"
                                            oRecordSet.DoQuery(sql)
                                            If oRecordSet.Fields.Item("Docentry").Value.ToString = "" Then
                                                JobLastDocEntry = 1
                                            Else
                                                JobLastDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString) + 1
                                            End If
                                            sql = "Insert Into [@OBT_FREIGHTDOCNO] (DocEntry,DocNum,U_JobNo,U_JobMode,U_JobType,U_JbStus,U_FrDocNo,U_JbDate,U_UDOName,U_CCode,U_CName,U_VCode,U_VName) Values " & _
                                                "(" & JobLastDocEntry & _
                                                    "," & JobLastDocEntry & _
                                                   "," & IIf(ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value), "NULL") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value), "NULL") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value), "NULL") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value), "NULL") & _
                                                    "," & FrDocEntry & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value), "Null") & _
                                                    "," & FormatString("IMPORTSEALCL") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value), "Null") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_Name").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_Name").Specific.Value), "Null") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_V").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_V").Specific.Value), "Null") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value), "Null") & ")"
                                            oRecordSet.DoQuery(sql)

                                            sql = "Update [@OBT_LCL01_IMPSEALCL] set U_JbDocNo=" & JobLastDocEntry & " Where DocEntry=" & FrDocEntry & ""
                                            oRecordSet.DoQuery(sql)
                                        End If
                                    ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.ToString & " And U_UDOName = " & FormatString("IMPORTSEALCL")
                                        oRecordSet.DoQuery(sql)
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AddIns" Then
                                    If String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.String) Then
                                        p_oSBOApplication.SetStatusBarMessage("Must Fill Trucker", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        'If ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = vbNullString Then
                                        If ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then
                                            modTrucking.AddUpdateInstructions(ImportSeaLCLForm, oMatrix, "@OBT_LCL03_TRUCKING", True)     '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                        Else
                                            modTrucking.AddUpdateInstructions(ImportSeaLCLForm, oMatrix, "@OBT_LCL03_TRUCKING", False)    '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                            ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                            ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                        End If
                                        ClearText(ImportSeaLCLForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                        ImportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                    End If
                                End If

                                If pVal.ItemUID = "bt_DelIns" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    modTrucking.DeleteByIndex(ImportSeaLCLForm, oMatrix, "@OBT_LCL03_TRUCKING")                             '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                    ClearText(ImportSeaLCLForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                    ImportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                End If

                                If pVal.ItemUID = "bt_AmdIns" Then
                                    ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Update Trucking Instruction"
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    'modTrucking.SetDataToEditTabByIndex(ImportSeaLCLForm)

                                    If (oMatrix.GetNextSelectedRow < 0) Then
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction", 1, "OK")
                                        'Exit Function
                                        BubbleEvent = False
                                        'edit by Chan
                                    Else
                                        modTrucking.GetDataFromMatrixByIndex(ImportSeaLCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        modTrucking.SetDataToEditTabByIndex(ImportSeaLCLForm)
                                        ImportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                        ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True 'MSW
                                    End If
                                    ImportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                End If

                                If pVal.ItemUID = "fo_TkrView" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    If oMatrix.RowCount > 1 Then
                                        ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                        Else
                                            ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                        End If
                                    End If
                                    ClearText(ImportSeaLCLForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                    ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                    ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                End If

                                '-------------------------Payment Vouncher (OMM)------------------------------------------------------'
                                If pVal.ItemUID = "fo_VoEdit" Then
                                    ImportSeaLCLForm.Freeze(True)
                                    ImportSeaLCLForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    ImportSeaLCLForm.Items.Item("op_Cash").Specific.Selected = True
                                    ImportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value = ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    ImportSeaLCLForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    'ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_VOUCHER").SetValue("U_DocNo", 0, ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value)

                                    If ImportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                        dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                        oMatrix.Clear()
                                    End If

                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = "" Then
                                        If (oMatrix.RowCount > 0) Then
                                            If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = 1
                                            Else
                                                ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                            End If
                                        Else
                                            ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = 1
                                        End If
                                    End If

                                    Dim oCombo As SAPbouiCOM.ComboBox
                                    oCombo = ImportSeaLCLForm.Items.Item("cb_PayCur").Specific
                                    If oCombo.ValidValues.Count = 0 Then
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CurrCode FROM OCRN")
                                        If oRecordSet.RecordCount > 0 Then
                                            oRecordSet.MoveFirst()
                                            While oRecordSet.EoF = False
                                                oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, "")
                                                oRecordSet.MoveNext()
                                            End While
                                        End If
                                    End If
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                    If dtmatrix.Rows.Count = 0 Then
                                        RowAddToMatrix(ImportSeaLCLForm, oMatrix)
                                    End If
                                    ImportSeaLCLForm.Items.Item("ed_VedName").Specific.Active = True
                                    ImportSeaLCLForm.Freeze(False)
                                End If

                                If pVal.ItemUID = "fo_VoView" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If oMatrix.RowCount > 1 Then
                                        ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                        Else
                                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                        End If
                                    End If
                                    ClearText(ImportSeaLCLForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    Dim oComboCurrency As SAPbouiCOM.ComboBox

                                    oComboBank = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    oComboCurrency = ImportSeaLCLForm.Items.Item("cb_PayCur").Specific

                                    For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                        oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                    For j As Integer = oComboCurrency.ValidValues.Count - 1 To 0 Step -1
                                        oComboCurrency.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                    ImportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" 'MSW 23-03-2011
                                End If
                                If pVal.ItemUID = "bt_Draft" Then
                                    ''''' Insert Voucher Item Table
                                    Dim DocEntry As Integer
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific                                        '* Change Nyan Lin   "[@[OBT_TB021_VOUCHER]]"
                                    oRecordSet.DoQuery("Delete From [@OBT_LCL15_VOUCHER] Where U_JobNo='" & ImportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value & "' And U_PVNo='" & ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value & "'")
                                    oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [@OBT_LCL15_VOUCHER] ")      '* Change Nyan Lin   "[@[OBT_TB021_VOUCHER]]"
                                    If oRecordSet.RecordCount > 0 Then
                                        DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                    End If
                                    'omm - Purchase Voucher Save To Draft 09-03-2010
                                    vocTotal = Convert.ToDouble(ImportSeaLCLForm.Items.Item("ed_Total").Specific.Value)
                                    gstTotal = Convert.ToDouble(ImportSeaLCLForm.Items.Item("ed_GSTAmt").Specific.Value)
                                    If ImportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                        SaveToPurchaseVoucher(ImportSeaLCLForm, True)
                                    Else
                                        SaveToPurchaseVoucher(ImportSeaLCLForm, False)
                                    End If
                                    SaveToDraftPurchaseVoucher(ImportSeaLCLForm)
                                    'End Purchase Voucher Save To Draft 09-03-2010
                                    dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                                    If oMatrix.RowCount > 0 Then
                                        For i As Integer = oMatrix.RowCount To 1 Step -1
                                            'To Add Item Code in Insert Statement
                                            '* Change Nyan Lin   "[@[OBT_TB021_VOUCHER]]"
                                            sql = "Insert Into [@OBT_LCL15_VOUCHER] (DocEntry,LineID,U_JobNo,U_PVNo,U_VSeqNo,U_ChCode,U_AccCode,U_ChDes,U_Amount,U_GST,U_GSTAmt,U_NoGST,U_ChrgCode) Values " & _
                                                   "(" & DocEntry & _
                                                    "," & i & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value), "NULL") & _
                                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value), "Null") & _
                                                    "," & IIf(oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.Value()), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colChCode").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colChCode").Cells.Item(i).Specific.Value()), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colAcCode").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colAcCode").Cells.Item(i).Specific.Value()), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colVDesc").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colVDesc").Cells.Item(i).Specific.Value()), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colAmount").Cells.Item(i).Specific.Value() <> "", oMatrix.Columns.Item("colAmount").Cells.Item(i).Specific.Value(), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colGST").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colGST").Cells.Item(i).Specific.Value()), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colGSTAmt").Cells.Item(i).Specific.Value() <> "", oMatrix.Columns.Item("colGSTAmt").Cells.Item(i).Specific.Value(), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colNoGST").Cells.Item(i).Specific.Value() <> "", oMatrix.Columns.Item("colNoGST").Cells.Item(i).Specific.Value(), "NULL") & _
                                                    "," & IIf(oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value()), "NULL") & ")"
                                            oRecordSet.DoQuery(sql)
                                            dtmatrix.Rows.Remove(i - 1)
                                        Next
                                    End If

                                    oMatrix.Clear()
                                    ''''' End Insert Voucher Item Table


                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If Convert.ToInt32(ImportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value) > oMatrix.RowCount Then
                                        AddUpdateVoucher(ImportSeaLCLForm, oMatrix, "@OBT_LCL05_VOUCHER", True)                               '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    Else
                                        AddUpdateVoucher(ImportSeaLCLForm, oMatrix, "@OBT_LCL05_VOUCHER", False)                              '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                        ImportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft"
                                    End If

                                    ClearText(ImportSeaLCLForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    Dim oComboCurrency As SAPbouiCOM.ComboBox

                                    oComboBank = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    oComboCurrency = ImportSeaLCLForm.Items.Item("cb_PayCur").Specific

                                    For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                        oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                    For j As Integer = oComboCurrency.ValidValues.Count - 1 To 0 Step -1
                                        oComboCurrency.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next


                                    ImportSeaLCLForm.Items.Item("fo_VoView").Specific.Select()
                                    ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True


                                End If

                                If pVal.ItemUID = "op_Cash" Then
                                    ImportSeaLCLForm.Items.Item("ed_VedName").Specific.Active = True
                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    oComboBank = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                        oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next

                                    ImportSeaLCLForm.Items.Item("cb_BnkName").Enabled = False
                                    ImportSeaLCLForm.Items.Item("ed_Cheque").Enabled = False
                                End If

                                If pVal.ItemUID = "op_Cheq" Then
                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    oComboBank = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    If oComboBank.ValidValues.Count = 0 Then
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("Select BankName,BankCode From ODSC")
                                        If oRecordSet.RecordCount > 0 Then
                                            oRecordSet.MoveFirst()
                                            While oRecordSet.EoF = False
                                                oComboBank.ValidValues.Add(oRecordSet.Fields.Item("BankName").Value, oRecordSet.Fields.Item("BankCode").Value)
                                                oRecordSet.MoveNext()
                                            End While
                                        End If

                                    End If

                                    ImportSeaLCLForm.Items.Item("cb_BnkName").Enabled = True
                                    ImportSeaLCLForm.Items.Item("ed_Cheque").Enabled = True
                                End If

                                If pVal.ItemUID = "bt_Cancel" Then
                                    ImportSeaLCLForm.Items.Item("fo_VoView").Specific.Select()
                                End If

                                If pVal.ItemUID = "mx_Voucher" And pVal.ColUID = "V_-1" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If oMatrix.GetNextSelectedRow > 0 Then
                                        If (oMatrix.IsRowSelected(oMatrix.GetNextSelectedRow)) = True Then
                                            GetVoucherDataFromMatrixByIndex(ImportSeaLCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        End If
                                    Else
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Payment Voucher.", 1, "&OK")
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AmdVoc" Then
                                    'POP UP Payment Voucher
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                        LoadPaymentVoucher(ImportSeaLCLForm)
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Voucher.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        'Exit Function
                                        BubbleEvent = False
                                    End If
                                    'Dim oComboBank As SAPbouiCOM.ComboBox
                                    'Dim oComboCurr As SAPbouiCOM.ComboBox
                                    'oComboBank = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    'oComboCurr = ImportSeaLCLForm.Items.Item("cb_PayCur").Specific
                                    'ImportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Update To Draft"

                                    'oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    'If oMatrix.GetNextSelectedRow < 0 Then
                                    '    p_oSBOApplication.MessageBox("Please Select One Row To Edit Payment Voucher.", 1, "&OK")
                                    '    Exit Function
                                    'End If

                                    'If oComboBank.ValidValues.Count = 0 Then
                                    '    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    '    oRecordSet.DoQuery("Select BankName From ODSC")
                                    '    If oRecordSet.RecordCount > 0 Then
                                    '        oRecordSet.MoveFirst()
                                    '        While oRecordSet.EoF = False
                                    '            oComboBank.ValidValues.Add(oRecordSet.Fields.Item("BankName").Value, "")
                                    '            oRecordSet.MoveNext()
                                    '        End While
                                    '    End If
                                    'End If

                                    'If oComboCurr.ValidValues.Count = 0 Then
                                    '    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    '    oRecordSet.DoQuery("SELECT CurrCode FROM OCRN")
                                    '    If oRecordSet.RecordCount > 0 Then
                                    '        oRecordSet.MoveFirst()
                                    '        While oRecordSet.EoF = False
                                    '            oComboCurr.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, "")
                                    '            oRecordSet.MoveNext()
                                    '        End While
                                    '    End If
                                    'End If


                                    ''To Add Item Code in Select Statement
                                    'sql = "Select U_ChCode as ChCode,U_AccCode as AcCode,U_ChDes as [Desc],U_Amount as Amount,U_GST as GST,U_GSTAmt As GSTAmt,U_NoGST As NoGST,U_VSeqNo as SeqNo,U_ChrgCode as ItemCode From [@OBT_LCL15_VOUCHER] " & _
                                    '    " Where U_JobNo = '" & ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value & "' And U_PVNo = '" & VocNo & "' Order By U_VSeqNo"             '* Change Nyan Lin   "[@OBT_TB021_VOUCHER]"

                                    'dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                                    'dtmatrix.ExecuteQuery(sql)
                                    'oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                    'oMatrix.LoadFromDataSource()

                                    'ImportSeaLCLForm.Items.Item("fo_VoEdit").Specific.Select()
                                    'SetVoucherDataToEditTabByIndex(ImportSeaLCLForm)
                                End If

                                '-----------------------------------------------------------------------------------------------------------'

                                If pVal.ItemUID = "mx_Cont" And ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If

                                If pVal.ItemUID = "bt_ChkList" Then
                                    Start(ImportSeaLCLForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                '-------------------------For Payment(omm)------------------------------------------'
                                If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "V_-1" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                    If pVal.Row > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            gridindex = CInt(pVal.Row)
                                        End If
                                    End If
                                End If
                                '----------------------------------------------------------------------------------'
                                If pVal.ItemUID = "mx_TkrList" Or pVal.ColUID = "V_1" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    If oMatrix.GetNextSelectedRow > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            modTrucking.rowIndex = CInt(pVal.Row)
                                            modTrucking.GetDataFromMatrixByIndex(ImportSeaLCLForm, oMatrix, modTrucking.rowIndex)
                                        End If
                                    Else
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction.", 1, "&OK")
                                    End If

                                End If

                                If pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize" And pVal.Before_Action = True And pVal.Row <> 0 Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
                                    Dim oColCombo As SAPbouiCOM.Column
                                    Dim omatCombo As SAPbouiCOM.ComboBox
                                    oColCombo = oMatrix.Columns.Item("colCSize")
                                    omatCombo = oColCombo.Cells.Item(pVal.Row).Specific
                                    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Value) Then
                                        Dim type As String = oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Selected.Value
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("select U_Size from [@OBT_LCL08_PCONTAINE] where U_Type='" & type & "'")                       '* Change Nyan Lin   "[@OBT_TB008_CONTAINER]" 
                                        oRecordSet.MoveFirst()
                                        While oColCombo.ValidValues.Count > 0
                                            oColCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                        End While
                                        If oRecordSet.RecordCount > 0 Then
                                            While Not oRecordSet.EoF
                                                omatCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Size").Value, " ")
                                                oRecordSet.MoveNext()
                                            End While
                                        End If
                                        oColCombo = Nothing
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                                Try
                                    '-------------------------For Payment(omm)------------------------------------------'
                                    If pVal.ItemUID = "ed_VedName" Then
                                        ObjDBDataSource = ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER")                 '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                                                                                               & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString())
                                        vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft                                        '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    End If
                                    If pVal.ColUID = "colChCode" Then
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                        dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                                        dtmatrix.SetValue(0, pVal.Row - 1, oDataTable.GetValue(0, 0).ToString)
                                        dtmatrix.SetValue(1, pVal.Row - 1, oDataTable.Columns.Item("U_PAccCode").Cells.Item(0).Value.ToString)
                                        dtmatrix.SetValue(2, pVal.Row - 1, oMatrix.Columns.Item("colVDesc").Cells.Item(pVal.Row).Specific.Value)
                                        dtmatrix.SetValue(3, pVal.Row - 1, oMatrix.Columns.Item("colAmount").Cells.Item(pVal.Row).Specific.Value)
                                        dtmatrix.SetValue(4, pVal.Row - 1, oMatrix.Columns.Item("colGST").Cells.Item(pVal.Row).Specific.Value)
                                        dtmatrix.SetValue(5, pVal.Row - 1, oMatrix.Columns.Item("colGSTAmt").Cells.Item(pVal.Row).Specific.Value)
                                        dtmatrix.SetValue(6, pVal.Row - 1, oMatrix.Columns.Item("colNoGST").Cells.Item(pVal.Row).Specific.Value)
                                        dtmatrix.SetValue(8, pVal.Row - 1, oDataTable.Columns.Item("U_ItemCode").Cells.Item(0).Value.ToString)  'To Add Item Code .. 7 for seqno
                                        'oMatrix.Clear()
                                        'oMatrix.AutoResizeColumns()
                                        oMatrix.LoadFromDataSource()
                                    End If
                                    '----------------------------------------------------------------------------------'
                                    If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString
                                        If String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value) Then
                                            ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            ImportSeaLCLForm.Items.Item("ed_ETADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            ImportSeaLCLForm.Items.Item("ed_ETAHr").Specific.Value = Now.ToString("HH:mm")
                                            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        If String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value) Then
                                            ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            ImportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            ImportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        EnabledHeaderControls(ImportSeaLCLForm, True)
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                                            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListUID = "WRHSE"
                                            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListAlias = "Code"
                                            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                                            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                                            ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                            ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                        End If
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_IComName", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_LCL06_PMAI]"
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_ShpAgt" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)     '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)      '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)            '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString) '* Change Nyan Lin   "[@OBT_LCL06_PMAIN]"
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_Wrhse" Then
                                        ImportSeaLCLForm.Items.Item("fo_Wrhse").Specific.Select()
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_WName", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)          '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WName", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)          '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WTel", 0, oDataTable.Columns.Item("U_WTel").Cells.Item(0).Value.ToString)         '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WCPrson", 0, oDataTable.Columns.Item("U_WPerson").Cells.Item(0).Value.ToString)      '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WMobile", 0, oDataTable.Columns.Item("U_WMobile").Cells.Item(0).Value.ToString)   '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WLast", 0, oDataTable.Columns.Item("U_WLat").Cells.Item(0).Value.ToString)         '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WLong", 0, oDataTable.Columns.Item("U_WLong").Cells.Item(0).Value.ToString)       '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WHr", 0, Trim(oDataTable.Columns.Item("U_WwhL1").Cells.Item(0).Value.ToString()))
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        sql = "SELECT U_WAdLine1,U_WAdLine2,U_WAdLine3,U_WState,U_WPostal,U_WCountry FROM [@OBT_TB003_WRHSE] WHERE Name = " & FormatString(oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)  '* need to Change
                                        Dim WAddress As String = String.Empty
                                        oRecordSet.DoQuery(sql)
                                        If oRecordSet.RecordCount > 0 Then
                                            WAddress = Trim(oRecordSet.Fields.Item("U_WAdLine1").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_WAdLine2").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_WAdLine3").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_WState").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_WPostal").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_WCountry").Value)
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WAddr", 0, WAddress)                                                         '* Change Nyan Lin   "[@OBT_TB017_WAREHOUSE]

                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = WAddress
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_RelName", 0, WAddress)  'NL LCL Change according to change latest design 
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_PoR", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                        End If
                                    End If
                                    If pVal.ItemUID = "ed_Trucker" Then
                                        If ImportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                            ImportSeaLCLForm.Freeze(True)
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                            oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = "" '25-3-2011
                                            ImportSeaLCLForm.Freeze(False)
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_Vessel" Or pVal.ItemUID = "ed_Voy" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)        '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)    '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_VName", 0, ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.String)                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    End If

                                    If pVal.ItemUID = "ed_CurCode" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL07_PCGDETAIL").SetValue("U_CurCode", 0, oDataTable.GetValue(0, 0).ToString)
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL07_PCGDETAIL").SetValue("U_CurRate", 0, Rate.ToString)
                                    End If
                                    If pVal.ItemUID = "ed_CCharge" Then
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL09_PINVDETAI").SetValue("U_Cchange", 0, oDataTable.GetValue(0, 0).ToString)                            '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL09_PINVDETAI").SetValue("U_CEchange", 0, Rate.ToString)                                               '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                    If pVal.ItemUID = "ed_Charge" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL09_PINVDETAI").SetValue("U_FCchange", 0, oDataTable.GetValue(0, 0).ToString)                          '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                Catch ex As Exception
                                End Try

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If pVal.ItemUID = "cb_PCode" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ImportSeaLCLForm.Items.Item("cb_PCode").Specific
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                             '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_PCode", 0, oCombo.Selected.Value.ToString)                                       '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                End If
                                If pVal.ItemUID = "cb_BnkName" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim test As String = "select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description
                                    oRecordSet.DoQuery("select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description)
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER").SetValue("U_GLAC", 0, oRecordSet.Fields.Item("GLAccount").Value)                            '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    'oCombo.Selected.Description
                                End If

                                If pVal.ItemUID = "cb_PType" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ImportSeaLCLForm.Items.Item("cb_PType").Specific
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL12_PTOTAL").SetValue("U_TUnit", 0, oCombo.Selected.Value.ToString)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN
                                If pVal.ItemUID = "ed_JobNo" And pVal.CharPressed = 13 Then
                                    ImportSeaLCLForm.Items.Item("1").Click()
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                'MSW 
                                If pVal.ItemUID = "ed_TkrDate" And pVal.Before_Action = False Then
                                    Dim strTime As SAPbouiCOM.EditText
                                    strTime = ImportSeaLCLForm.Items.Item("ed_TkrTime").Specific
                                    strTime.Value = Now.ToString("HH:mm")
                                End If
                                'End MSW
                                If pVal.ItemUID = "ed_InvNo" Then
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL07_PCGDETAIL").SetValue("U_InvNo", 0, ImportSeaLCLForm.Items.Item("ed_InvNo").Specific.String)         '* Change Nyan Lin   "[@OBT_TB0011_VOUCHER]"
                                End If
                                'NL LCL Change 24-03-2011
                                If pVal.ItemUID = "ed_NOP" Then
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL12_PTOTAL").SetValue("U_TotalOP", 0, ImportSeaLCLForm.Items.Item("ed_NOP").Specific.String)
                                End If
                                'End NL LCL Change 24-03-2011
                                Validateforform(pVal.ItemUID, ImportSeaLCLForm)

                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If BoolResize = False Then
                                    Try
                                        Dim oItemRet1 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEALCL").Items.Item("rt_Outer")
                                        Dim oItemRetInner As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEALCL").Items.Item("rt_Inner")
                                        oItemRetInner.Width = ImportSeaLCLForm.Items.Item("mx_Cont").Width + 15
                                        oItemRetInner.Height = ImportSeaLCLForm.Items.Item("mx_Cont").Height + 140
                                        oItemRet1.Top = ImportSeaLCLForm.Items.Item("mx_TkrList").Top - 29
                                        oItemRet1.Width = ImportSeaLCLForm.Items.Item("mx_TkrList").Width + 20
                                        oItemRet1.Height = ImportSeaLCLForm.Items.Item("mx_Voucher").Height + 90 '33
                                        ImportSeaLCLForm.Items.Item("155").Top = ImportSeaLCLForm.Items.Item("mx_TkrList").Top - 5
                                        ImportSeaLCLForm.Items.Item("155").Width = ImportSeaLCLForm.Items.Item("mx_TkrList").Width + 10
                                        ImportSeaLCLForm.Items.Item("155").Height = ImportSeaLCLForm.Items.Item("mx_TkrList").Height + 5
                                        BoolResize = True
                                    Catch ex As Exception
                                        ' p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try
                                ElseIf BoolResize = True Then
                                    Try
                                        Dim oItemRet2 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEALCL").Items.Item("rt_Outer")
                                        oItemRet2.Top = ImportSeaLCLForm.Items.Item("mx_TkrList").Top - 29
                                        Dim oItemRetInner2 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEALCL").Items.Item("rt_Inner")
                                        oItemRetInner2.Width = ImportSeaLCLForm.Items.Item("mx_Cont").Width + 15
                                        oItemRetInner2.Height = ImportSeaLCLForm.Items.Item("mx_Cont").Height + 140
                                        oItemRet2.Width = ImportSeaLCLForm.Items.Item("mx_TkrList").Width + 20
                                        oItemRet2.Height = ImportSeaLCLForm.Items.Item("mx_Voucher").Height + 90 '33
                                        ImportSeaLCLForm.Items.Item("155").Top = ImportSeaLCLForm.Items.Item("mx_TkrList").Top - 5
                                        ImportSeaLCLForm.Items.Item("155").Width = ImportSeaLCLForm.Items.Item("mx_TkrList").Width + 10
                                        ImportSeaLCLForm.Items.Item("155").Height = ImportSeaLCLForm.Items.Item("mx_TkrList").Height + 5
                                        BoolResize = False
                                    Catch ex As Exception
                                        'p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try
                                End If
                        End Select
                    End If
                    If pVal.BeforeAction = True Then
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "1" Then
                                    Dim PODFlag As String = String.Empty
                                    Dim JbStus As String = String.Empty
                                    Dim DispatchComplete As String = String.Empty
                                    If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        'develivery process by POD[Proof Of Delivery] check box
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_LCL01_IMPSEALCL] WHERE DocEntry = " & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value) 'MSW for Job Type Table
                                        If oRecordSet.RecordCount > 0 Then
                                            JbStus = oRecordSet.Fields.Item("U_JbStus").Value
                                            PODFlag = oRecordSet.Fields.Item("U_POD").Value
                                        End If
                                        If ImportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = True And JbStus = "Open" Then
                                            If p_oSBOApplication.MessageBox("Make sure that all entries trucking and vouchers are completed.(ensure no draft Payment in this job and " & _
                                                                       "ensure all external trucking transaction has generated the PO). Cannot edit or add after click POD check box. " & _
                                                                       "Do you want to continue?", 1, "&Yes", "&No") = 2 Then
                                                BubbleEvent = False
                                            End If
                                        End If
                                        If BubbleEvent = True Then
                                            'MSW 01-06-2011 for job type table
                                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            'End MSW 01-06-2011 for job type table
                                        End If
                                    End If
                                    If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        'handle for dispatch complete check box
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim ts As String = "SELECT U_Complete FROM [@OBT_LCL04_DISPATCH] WHERE DocEntry = " & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value
                                        oRecordSet.DoQuery(ts)
                                        If oRecordSet.RecordCount > 0 Then
                                            DispatchComplete = oRecordSet.Fields.Item("U_Complete").Value
                                        End If
                                        If ImportSeaLCLForm.Items.Item("ch_Dsp").Specific.Checked = True And DispatchComplete = "Y" Then
                                            BubbleEvent = False
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                                        End If
                                        If ImportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = True And PODFlag = "Y" Then
                                            BubbleEvent = False
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        End If
                                    End If
                                    If ImportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And ImportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If Validateforform(" ", ImportSeaLCLForm) Then
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                        End Select
                    End If
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoImportSeaLCLItemEvent = RTN_SUCCESS

        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoImportSeaLCLItemEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function

    Public Function DoImportSeaLCLRightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoImportSeaLCLItemEvent
        '   Purpose     :   This function is provider for ImportSeaLCL Right Click Event
        '               
        '   Parameters  :   ByRef eventInfo As SAPbouiCOM.ContextMenuInfo
        '                       pVal =  set the SAP Context Menu Info Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim bFlag As Boolean = False
        Dim FunctionName As String = "DoImportSeaLCLRightClickEvent"
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)

            If eventInfo.BeforeAction = True Then
                Select Case eventInfo.ItemUID
                    Case "[Matrix Name]"      'Enable the Add Row Menu and Delete Row function [Right Click Menu]
                        oMatrix = ImportSeaLCLForm.Items.Item("[Matrix Name]").Specific
                        If eventInfo.Row > 0 And (ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                            'bFlag = IIf(Not oMatrix.Columns.Item(oMatrix.Columns.Count - 2).Cells.Item(eventInfo.Row).Specific.Value.Equals(""), False, True)
                            If oMatrix.VisualRowCount = 1 Then bFlag = False
                        End If
                        'ImportSeaLCLForm.EnableMenu("1292", bFlag)
                        'ImportSeaLCLForm.EnableMenu("1293", bFlag)

                    Case "mx_Voucher"
                        If (eventInfo.BeforeAction = True) Then
                            'Dim oMenuItem As SAPbouiCOM.MenuItem
                            'Dim oMenus As SAPbouiCOM.Menus
                            Try
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = "EditVoc"
                                oCreationPackage.String = "Edit Payment Voucher"
                                oCreationPackage.Enabled = True
                                oMenuItem = p_oSBOApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("EditVoc") Then
                                    p_oSBOApplication.Menus.RemoveEx("EditVoc")
                                End If
                                oMenus.AddEx(oCreationPackage)
                            Catch ex As Exception
                                MessageBox.Show(ex.Message)
                            End Try
                        End If
                    Case Else
                        Try
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = "EditVoc"
                            oCreationPackage.String = "Edit Payment Voucher"
                            oCreationPackage.Enabled = True
                            oMenuItem = p_oSBOApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            If oMenus.Exists("EditVoc") Then
                                p_oSBOApplication.Menus.RemoveEx("EditVoc")
                            End If
                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                End Select
            End If

            DoImportSeaLCLRightClickEvent = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoImportSeaLCLRightClickEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function

    Private Sub EnabledHeaderControls(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)
        pForm.Items.Item("ed_ShpAgt").Enabled = pValue
        pForm.Items.Item("ed_OBL").Enabled = pValue
        pForm.Items.Item("ed_HBL").Enabled = pValue
        pForm.Items.Item("ed_Conn").Enabled = pValue
        pForm.Items.Item("ed_Vessel").Enabled = pValue
        pForm.Items.Item("ed_Voy").Enabled = pValue
        pForm.Items.Item("cb_PCode").Enabled = pValue
        pForm.Items.Item("ed_ETADate").Enabled = pValue
        pForm.Items.Item("ed_ETAHr").Enabled = pValue
        pForm.Items.Item("ed_CrgDsc").Enabled = pValue
        pForm.Items.Item("ed_TotalM3").Enabled = pValue
        pForm.Items.Item("ed_TotalWt").Enabled = pValue
        pForm.Items.Item("ed_NOP").Enabled = pValue
        pForm.Items.Item("cb_PType").Enabled = pValue

        pForm.Items.Item("cb_JobType").Enabled = pValue
        pForm.Items.Item("cb_JbStus").Enabled = pValue
        pForm.Items.Item("ed_Wrhse").Enabled = pValue
        pForm.Items.Item("ed_JobNo").Enabled = pValue

        pForm.Items.Item("ed_LCDDate").Enabled = pValue
        pForm.Items.Item("ed_LCDHr").Enabled = pValue

        pForm.Items.Item("ch_CnUn").Enabled = pValue
        pForm.Items.Item("ch_CgCl").Enabled = pValue
    End Sub

    Private Function CloseOpenForm(ByVal sFormId As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = String.Empty
        Dim iTemp As Int16
        Dim oForm As SAPbouiCOM.Form
        Try
            sFuncName = "CloseOpenForm()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Debug.Print("Forms Count  " + p_oSBOApplication.Forms.Count.ToString)
            For iTemp = 0 To p_oSBOApplication.Forms.Count - 1
                oForm = p_oSBOApplication.Forms.Item(iTemp)
                Debug.Print("Index =" + iTemp.ToString + "   TypeEx = " + oForm.TypeEx.ToString)
                If oForm.TypeEx = sFormId Then
                    Debug.Print("Index =" + iTemp.ToString + " Closing Form   " + oForm.TypeEx.ToString)
                    oForm.Close()
                End If
            Next iTemp
            CloseOpenForm = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with SUCCESS", sFuncName)
        Catch ex As Exception
            CloseOpenForm = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with ERROR", sFuncName)
        End Try
    End Function

    

    Private Sub ClearTruckingInfo(ByRef pForm As SAPbouiCOM.Form)
        pForm.Freeze(True)
        pForm.Items.Item("ed_PONo").Specific.Value = vbNullString
        pForm.Items.Item("ed_Trucker").Specific.Value = vbNullString
        pForm.Items.Item("ed_VehicNo").Specific.Value = vbNullString
        pForm.Items.Item("ed_EUC").Specific.Value = vbNullString
        pForm.Items.Item("ed_Attent").Specific.Value = vbNullString
        pForm.Items.Item("ed_TkrTel").Specific.Value = vbNullString
        pForm.Items.Item("ed_Fax").Specific.Value = vbNullString
        pForm.Items.Item("ed_Email").Specific.Value = vbNullString
        pForm.Items.Item("ed_TkrDate").Specific.Value = vbNullString
        pForm.Items.Item("ed_TkrTime").Specific.Value = vbNullString
        pForm.Items.Item("ee_ColFrm").Specific.Value = vbNullString
        pForm.Items.Item("ee_TkrTo").Specific.Value = vbNullString
        pForm.Items.Item("ee_TkrIns").Specific.Value = vbNullString
        pForm.Items.Item("ee_InsRmsk").Specific.Value = vbNullString
        pForm.Freeze(False)
    End Sub

    Private Sub ClearText(ByRef pForm As SAPbouiCOM.Form, ByVal ParamArray pControls() As String)
        Dim strTempUID As String = String.Empty
        Try
            If pControls.Length <= 0 Then Exit Sub
            pForm.Freeze(True)
            For i As Integer = 0 To pControls.Length - 1
                pForm.Items.Item(pControls.GetValue(i)).Specific.Value = vbNullString
            Next
            pForm.Freeze(False)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EnabledUIObjects(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean, ByVal ParamArray pControls() As String)
        Dim strTempUID As String = String.Empty
        Try
            If pControls.Length <= 0 Then Exit Sub
            pForm.Freeze(True)
            For i As Integer = 0 To pControls.Length - 1
                pForm.Items.Item(pControls.GetValue(i)).Enabled = pValue
            Next
            pForm.Freeze(False)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub SetMatrixSeqNo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal ColName As String)
        For i As Integer = 1 To oMatrix.RowCount
            oMatrix.Columns.Item(ColName).Cells.Item(i).Specific.Value = i
        Next
    End Sub

    Private Sub EnabledMaxtix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)
        Try
            pMatrix.Columns.Item("colInsDoc").Editable = pValue
            pMatrix.Columns.Item("colInsDoc").BackColor = 16645629
            pMatrix.Columns.Item("colPONo").Editable = pValue
            pMatrix.Columns.Item("colPONo").BackColor = 16645629
            pMatrix.Columns.Item("colInsDate").Editable = pValue
            pMatrix.Columns.Item("colInsDate").BackColor = 16645629
            pMatrix.Columns.Item("colMode").Editable = pValue
            pMatrix.Columns.Item("colMode").BackColor = 16645629
            pMatrix.Columns.Item("colTrucker").Editable = pValue
            pMatrix.Columns.Item("colTrucker").BackColor = 16645629
            pMatrix.Columns.Item("colVehNo").Editable = pValue
            pMatrix.Columns.Item("colVehNo").BackColor = 16645629
            pMatrix.Columns.Item("colEUC").Editable = pValue
            pMatrix.Columns.Item("colEUC").BackColor = 16645629
            pMatrix.Columns.Item("colAttent").Editable = pValue
            pMatrix.Columns.Item("colAttent").BackColor = 16645629
            pMatrix.Columns.Item("colTel").Editable = pValue
            pMatrix.Columns.Item("colTel").BackColor = 16645629
            pMatrix.Columns.Item("colFax").Editable = pValue
            pMatrix.Columns.Item("colFax").BackColor = 16645629
            pMatrix.Columns.Item("colEmail").Editable = pValue
            pMatrix.Columns.Item("colEmail").BackColor = 16645629
            pMatrix.Columns.Item("colTkrDate").Editable = pValue
            pMatrix.Columns.Item("colTkrDate").BackColor = 16645629
            pMatrix.Columns.Item("colTkrTime").Editable = pValue
            pMatrix.Columns.Item("colTkrTime").BackColor = 16645629
            pMatrix.Columns.Item("colColFrom").Editable = pValue
            pMatrix.Columns.Item("colColFrom").BackColor = 16645629
            pMatrix.Columns.Item("colTkrTo").Editable = pValue
            pMatrix.Columns.Item("colTkrTo").BackColor = 16645629
            pMatrix.Columns.Item("colTkrIns").Editable = pValue
            pMatrix.Columns.Item("colTkrIns").BackColor = 16645629
            pMatrix.Columns.Item("colRemarks").Editable = pValue
            pMatrix.Columns.Item("colRemarks").BackColor = 16645629
            pMatrix.Columns.Item("colPrepBy").Editable = pValue
            pMatrix.Columns.Item("colPrepBy").BackColor = 16645629
        Catch ex As Exception

        End Try
    End Sub

    Private Sub EnabledTrucker(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)
        Try
            pForm.Items.Item("ed_InsDoc").Enabled = pValue
            pForm.Items.Item("ed_InsDoc").BackColor = 16645629
            pForm.Items.Item("ed_PONo").Enabled = pValue
            pForm.Items.Item("ed_PONo").BackColor = 16645629
            pForm.Items.Item("ed_InsDate").Enabled = pValue
            pForm.Items.Item("ed_InsDate").BackColor = 16645629
            pForm.Items.Item("op_Inter").Enabled = pValue
            pForm.Items.Item("op_Exter").Enabled = pValue
            pForm.Items.Item("ed_Trucker").Enabled = pValue
            pForm.Items.Item("ed_Trucker").BackColor = 16645629
            pForm.Items.Item("ed_VehicNo").Enabled = pValue
            pForm.Items.Item("ed_VehicNo").BackColor = 16645629
            pForm.Items.Item("ed_EUC").Enabled = pValue
            pForm.Items.Item("ed_EUC").BackColor = 16645629
            pForm.Items.Item("ed_Attent").Enabled = pValue
            pForm.Items.Item("ed_Attent").BackColor = 16645629
            pForm.Items.Item("ed_TkrTel").Enabled = pValue
            pForm.Items.Item("ed_TkrTel").BackColor = 16645629
            pForm.Items.Item("ed_Fax").Enabled = pValue
            pForm.Items.Item("ed_Fax").BackColor = 16645629
            pForm.Items.Item("ed_Email").Enabled = pValue
            pForm.Items.Item("ed_Email").BackColor = 16645629
            pForm.Items.Item("ed_TkrDate").Enabled = pValue
            pForm.Items.Item("ed_TkrDate").BackColor = 16645629
            pForm.Items.Item("ed_TkrTime").Enabled = pValue
            pForm.Items.Item("ed_TkrTime").BackColor = 16645629
            pForm.Items.Item("ee_ColFrm").Enabled = pValue
            pForm.Items.Item("ee_ColFrm").BackColor = 16645629
            pForm.Items.Item("ee_TkrTo").Enabled = pValue
            pForm.Items.Item("ee_TkrTo").BackColor = 16645629
            pForm.Items.Item("ee_TkrIns").Enabled = pValue
            pForm.Items.Item("ee_TkrIns").BackColor = 16645629
            pForm.Items.Item("ee_InsRmsk").Enabled = pValue
            pForm.Items.Item("ee_InsRmsk").BackColor = 16645629
        Catch ex As Exception

        End Try
    End Sub

    Private Function Validateforform(ByVal ItemUID As String, ByVal ImportSeaLCLForm As SAPbouiCOM.Form) As Boolean
        If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_Name").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "cb_PCode" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("cb_PCode").Specific.Value) Then
            p_oSBOApplication.SetStatusBarMessage("Must Select Port Of Loading", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Shipping Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_Wrhse" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_LCDDate" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Last Clearance Date", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_WAddr" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_WAddr").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Fill Warehouse Address", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_JobNo" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Fill Job No", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub LoadHolidayMarkUp(ByVal ImportSeaLCLForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_LCDDay").Specific, ImportSeaLCLForm.Items.Item("ed_LCDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Private Function AddChooseFromListByOption(ByRef pForm As SAPbouiCOM.Form, ByVal pOption As Boolean, ByVal pObjID As String, ByVal pErrDesc As String) As Long
        Dim oEditText As SAPbouiCOM.EditText
        Try
            If pOption = True Then
                oEditText = pForm.Items.Item(pObjID).Specifice
                oEditText.DataBind.SetBound(True, "", "TKRINTR")
                oEditText.ChooseFromListUID = "CFLTKRE"
                oEditText.ChooseFromListAlias = "firstName"
            Else
                oEditText = pForm.Items.Item(pObjID).Specific
                oEditText.DataBind.SetBound(True, "", "TKREXTR")
                oEditText.ChooseFromListUID = "CFLTKRV"
                oEditText.ChooseFromListAlias = "CardName"
            End If
            AddChooseFromListByOption = RTN_SUCCESS
            AddChooseFromListByOption = RTN_SUCCESS
        Catch ex As Exception
            AddChooseFromListByOption = RTN_ERROR
        End Try
    End Function

#Region "---------- 'OMM - Purchase Voucher Save To Draft 09-03-2010"
    Private Sub SaveToPurchaseVoucher(ByRef pForm As SAPbouiCOM.Form, ByVal ProcessedState As Boolean)
        Dim nErr As Long
        Dim errMsg As String = String.Empty
        Dim ObjectCode As String = String.Empty
        Dim invDocEntry As Integer
        Dim Document As SAPbobsCOM.Documents
        Dim businessPartner As SAPbobsCOM.BusinessPartners
        Document = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        businessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        'Debug.Print(businessPartner.GetByKey(vendorCode))
        'Debug.Print(businessPartner.CardCode)
        vendorCode = pForm.Items.Item("ed_VedCode").Specific.Value
        If ProcessedState = False Then
            oRecordSet.DoQuery("Select DocEntry From OPCH Where U_JobNo='" & pForm.Items.Item("ed_PJobNo").Specific.Value & "' And U_PVNo='" & pForm.Items.Item("ed_VocNo").Specific.Value & "'")
            invDocEntry = oRecordSet.Fields.Item("DocEntry").Value
            If (Document.GetByKey(invDocEntry)) Then
                If (businessPartner.GetByKey(vendorCode)) Then
                    Document.CardCode = vendorCode
                    Document.CardName = pForm.Items.Item("ed_VedName").Specific.Value
                    Document.Address = pForm.Items.Item("ed_PayTo").Specific.Value
                    Document.DocCurrency = businessPartner.Currency
                    Document.DocDate = Now
                    Document.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    Document.JournalMemo = "A/P Invoices - " & vendorCode
                    Document.Series = 6
                    Document.TaxDate = Now
                End If
                If (Document.Update() <> 0) Then
                    MsgBox("Failed to add a payment")
                Else

                    'Alert() 'That is Alert Nayn Lin
                End If

            End If
        Else
            If (businessPartner.GetByKey(vendorCode)) Then
                Document.CardCode = vendorCode
                Document.CardName = pForm.Items.Item("ed_VedName").Specific.Value
                Document.Address = pForm.Items.Item("ed_PayTo").Specific.Value
                Document.DocCurrency = businessPartner.Currency
                Document.DocDate = Now
                Document.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                Document.JournalMemo = "A/P Invoices - " & vendorCode
                Document.Series = 6

                Document.TaxDate = Now
                Dim oMatrix As SAPbouiCOM.Matrix
                oMatrix = pForm.Items.Item("mx_ChCode").Specific
                dtmatrix = pForm.DataSources.DataTables.Item("DTCharges")
                If oMatrix.RowCount > 0 Then
                    For i As Integer = 1 To oMatrix.RowCount
                        Document.Lines.ItemCode = oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value
                        Document.Lines.ItemDescription = oMatrix.Columns.Item("colVDesc").Cells.Item(i).Specific.Value
                        Document.Lines.Quantity = 1
                        Document.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colAmount").Cells.Item(i).Specific.Value)
                        'MSW 23-03-2011 For VatCode GST None or Blank in GST Field if we didn't assign ZI ,system auto populate default SI 
                        If (dtmatrix.GetValue("GST", i - 1) = "" Or dtmatrix.GetValue("GST", i - 1) = "None") Then
                            Document.Lines.VatGroup = "ZI"
                        Else
                            Document.Lines.VatGroup = dtmatrix.GetValue("GST", i - 1)
                        End If
                        'Document.Lines.VatGroup = dtmatrix.GetValue("GST", i - 1) 'oMatrix.Columns.Item("colGST").Cells.Item(i).Specific.Value 'MSW 23-03-2011
                        Document.Lines.Add()
                    Next
                End If
            End If

            If (Document.Add() <> 0) Then
                MsgBox("Failed to add a payment")
            End If

        End If

        'Check Error 

        Call p_oDICompany.GetLastError(nErr, errMsg)
        If (0 <> nErr) Then
            MsgBox("Found error:" + Str(nErr) + "," + errMsg)
        Else
            ' MsgBox("Succeed in payment.add")
            p_oDICompany.GetNewObjectCode(ObjectCode)
            ObjectCode = p_oDICompany.GetNewObjectKey()
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sql = "Update OPCH set U_JobNo='" & pForm.Items.Item("ed_PJobNo").Specific.Value & "',U_PVNo='" & pForm.Items.Item("ed_VocNo").Specific.Value & "',U_FrDocNo='" & pForm.Items.Item("ed_DocNum").Specific.Value & "'" & _
            " Where DocEntry = " & Convert.ToInt32(ObjectCode) & ""
            oRecordSet.DoQuery(sql)
        End If

    End Sub
	
    Private Sub Alert()
        Dim code As String = "16"
        Dim AlertName As String = "Alert"
        'Dim QueryName As String = "AlertOne"
        'Dim QueryString As String = "select * from OPDF where  DocEntry =(select MAX(DocEntry )from OPDF)"
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("select * From OALT where Name='" & AlertName.ToString() & "'")
        If oRecordSet.RecordCount > 0 Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Update OALT set ExecTime='" & Now.Hour.ToString & Convert.ToString(Now.Minute + 2) & "', NextDate='" & Today.Date & "',NextTime='" & Now.Hour.ToString & Convert.ToString((Convert.ToInt32(Now.Minute.ToString()) + 2)) & "' where  Name='" & AlertName.ToString() & "'")
            'oRecordSet.DoQuery("Update OALT set  NextTime='" & QueryString.ToString() & "'  where  Code='" & QueryName.ToString() & "'")
            Status = "Update"
        End If
        If Status <> "Update" Then

            Dim oAlertManagement As SAPbobsCOM.AlertManagement
            Dim oAlertManagementParams As SAPbobsCOM.AlertManagementParams
            Dim oAlertManagementRecipients As SAPbobsCOM.AlertManagementRecipients
            Dim oAlertRecipient As SAPbobsCOM.AlertManagementRecipient
            Dim oAlertMangementService As SAPbobsCOM.AlertManagementService
            Dim oCompany As SAPbobsCOM.Company
            Dim sCookie As String
            Dim sConnStr As String
            Dim lRetval As String
            Dim Status As String = vbNullString
            oCompany = New SAPbobsCOM.Company
            sCookie = oCompany.GetContextCookie
            sConnStr = p_oUICompany.GetConnectionContext(sCookie)
            lRetval = oCompany.SetSboLoginContext(sConnStr)
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            oCompany.Connect()
            oCmpSrv = oCompany.GetCompanyService
            oAlertMangementService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.AlertManagementService)
            oAlertManagement = oAlertMangementService.GetDataInterface(SAPbobsCOM.AlertManagementServiceDataInterfaces.atsdiAlertManagement)
            oAlertManagement.Name = "Alert" 'Change
            oAlertManagement.QueryID = 145 'Change
            oAlertManagement.Active = SAPbobsCOM.BoYesNoEnum.tYES
            oAlertManagement.Priority = SAPbobsCOM.AlertManagementPriorityEnum.atp_High
            Try
                oAlertManagement.DayOfExecution = 1
                oAlertManagement.ExecutionTime = Convert.ToDateTime(Now.Hour.ToString + ":" + Now.Minute.ToString())
            Catch ex As Exception

            End Try
            oAlertManagement.FrequencyInterval = 1
            oAlertManagement.FrequencyType = SAPbobsCOM.AlertManagementFrequencyType.atfi_Monthly
            oAlertManagementRecipients = oAlertManagement.AlertManagementRecipients
            oAlertRecipient = oAlertManagementRecipients.Add()
            oAlertRecipient.UserCode = 1
            oAlertRecipient.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
            oAlertManagementParams = oAlertMangementService.AddAlertManagement(oAlertManagement)
            oRecordSet.DoQuery("Update OALT set NextDate='" & Today.Date & "',NextTime='" & Now.Hour.ToString & Convert.ToString((Convert.ToInt32(Now.Minute.ToString()) + 2)) & "' where  Code='" & code.ToString() & "'")
        End If

    End Sub
    Private Sub SaveToDraftPurchaseVoucher(ByRef pForm As SAPbouiCOM.Form)
        Dim nErr As Long
        Dim errMsg As String = String.Empty
        Dim ObjectCode As String = String.Empty
        Dim puchaseDocEntry As Integer
        Dim nextOutgoing As Integer
        Dim draftDocEntry As Integer
        Dim CurRate As Double 'To Calculate Amount
        Dim draftCurRate As Double 'Only To Save Data
        Dim ApplySys As Double
        Dim ApplyFC As Double
        Dim vatAppldSy As Double
        Dim vatAppldFC As Double

        p_oDICompany.GetNewObjectCode(ObjectCode)
        ObjectCode = p_oDICompany.GetNewObjectKey()
        puchaseDocEntry = Convert.ToInt32(ObjectCode)
        Dim vPay As SAPbobsCOM.Payments
        Dim businessPartner As SAPbobsCOM.BusinessPartners

        oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From OVPM ")
        If oRecordSet.RecordCount > 0 Then
            nextOutgoing = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
        End If
        vPay = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
        vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
        businessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        If (businessPartner.GetByKey(vendorCode)) Then

            vPay.DocNum = nextOutgoing
            vPay.CardCode = vendorCode
            vPay.CardName = pForm.Items.Item("ed_VedName").Specific.Value
            vPay.Address = pForm.Items.Item("ed_PayTo").Specific.Value
            vPay.ApplyVAT = 1

            vPay.DocCurrency = businessPartner.Currency
            vPay.DocDate = Now
            'vPay.DocRate = 0.0
            vPay.DocTypte = 2
            vPay.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
            vPay.JournalRemarks = "Outgoing Payments - " & vendorCode
            vPay.Series = 15
            vPay.TaxDate = Now

            vPay.TransferAccount = "161012"
            vPay.TransferDate = Now
            vPay.TransferSum = vocTotal
        End If

        If (vPay.Add() <> 0) Then
            MsgBox("Failed to add a payment")
        Else
            'Alert() 'That is Alert Nayn Lin

        End If
        Call p_oDICompany.GetLastError(nErr, errMsg)
        If (0 <> nErr) Then
            MsgBox("Found error:" + Str(nErr) + "," + errMsg)
        Else
            ' MsgBox("Succeed in Draft.add")
            p_oDICompany.GetNewObjectCode(ObjectCode)
            ObjectCode = p_oDICompany.GetNewObjectKey()
            draftDocEntry = Convert.ToInt32(ObjectCode)
            oRecordSet.DoQuery("Select SysRate As CurRate From OPDF")
            If oRecordSet.RecordCount > 0 Then
                CurRate = Convert.ToDouble(oRecordSet.Fields.Item("CurRate").Value)
            End If
            ApplySys = vocTotal / CurRate
            vatAppldSy = gstTotal / CurRate
            If businessPartner.Currency = "SGD" Then
                draftCurRate = 0.0
                ApplyFC = 0.0
                vatAppldFC = 0.0
            Else
                draftCurRate = CurRate
                ApplyFC = ApplySys
                vatAppldFC = vatAppldSy
            End If
            sql = "Insert Into PDF2 (DocNum,InvoiceID,DocEntry,SumApplied,AppliedFC,AppliedSys,vatApplied,vatAppldFC,vatAppldSy,DocRate,InvType) Values " & _
                                           "(" & draftDocEntry & _
                                            "," & 0 & _
                                            "," & puchaseDocEntry & _
                                            "," & vocTotal & _
                                            "," & ApplyFC & _
                                            "," & ApplySys & _
                                             "," & gstTotal & _
                                            "," & vatAppldFC & _
                                            "," & vatAppldSy & _
                                            "," & draftCurRate & _
                                            ",'" & 18 & "')"
            oRecordSet.DoQuery(sql)
        End If
    End Sub
#End Region

    Private Sub Start(ByRef pform As SAPbouiCOM.Form)
        Dim str As String = "dwdesk.exe"
        Dim myprocess As New Process
        Dim mainfolderpath As String = "C:\Users\PC-8\Documents\Fuji Xerox\DocuWorks\DWFolders\User Folder\"
        Dim foldername As String = pform.Items.Item("ed_JobNo").Specific.Value
        Dim docfolderpath As String = mainfolderpath & foldername
        Dim di As DirectoryInfo = New DirectoryInfo(docfolderpath)
        If Not di.Exists Then
            di.Create()
        End If

        Dim argument As String = "/f" & Chr(34) & docfolderpath
        Try
            myprocess.StartInfo.FileName = str
            myprocess.StartInfo.Arguments = argument
            myprocess.Start()
            myprocess.Refresh()
            If myprocess.HasExited = False Then
                myprocess.WaitForInputIdle(10000)
                MoveWindow(myprocess.MainWindowHandle, p_oSBOApplication.Forms.ActiveForm.Left + p_oSBOApplication.Forms.ActiveForm.Width + 6, p_oSBOApplication.Forms.ActiveForm.Top + 68, 300, 578, True)
            End If
        Catch ex As Exception

        End Try


    End Sub
#Region "Voucher POP Up"

    Private Sub LoadPaymentVoucher(ByRef oActiveForm As SAPbouiCOM.Form)
        Dim oPayForm As SAPbouiCOM.Form
        Dim oOptBtn As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        If Not LoadFromXML(p_oSBOApplication, "LCLPaymentVoucher.srf") Then Throw New ArgumentException(sErrDesc)
        oPayForm = p_oSBOApplication.Forms.ActiveForm
        oPayForm.Freeze(True)
        If AddChooseFromList(oPayForm, "PAYMENT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oPayForm.Items.Item("ed_VedName").Specific.ChooseFromListUID = "PAYMENT"
        oPayForm.Items.Item("ed_VedName").Specific.ChooseFromListAlias = "CardName"
        oPayForm.EnableMenu("1288", True)
        oPayForm.EnableMenu("1289", True)
        oPayForm.EnableMenu("1290", True)
        oPayForm.EnableMenu("1291", True)
        oPayForm.EnableMenu("1284", False)
        oPayForm.EnableMenu("1286", False)

        oPayForm.EnableMenu("1283", False) 'MSW 01-04-2011
        oPayForm.EnableMenu("1292", True)
        oPayForm.EnableMenu("1293", True)
        oPayForm.EnableMenu("5907", True)

        oPayForm.AutoManaged = True
        oPayForm.DataBrowser.BrowseBy = "ed_DocNum"
        oPayForm.Items.Item("cb_BnkName").Enabled = False
        oPayForm.Items.Item("ed_Cheque").Enabled = False
        oPayForm.Items.Item("ed_PayRate").Enabled = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPayForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("LCLVOUCHER", oRecordSet)
        oPayForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        oPayForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()

        If HolidayMarkUp(oPayForm, oPayForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oPayForm.DataSources.DBDataSources.Item("@OBT_LCL16_VHEADER").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)

        If AddUserDataSrc(oPayForm, "VCASH", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oPayForm, "VCHEQUE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oOptBtn = oPayForm.Items.Item("op_Cash").Specific
        oOptBtn.DataBind.SetBound(True, "", "VCASH")
        oOptBtn = oPayForm.Items.Item("op_Cheq").Specific
        oOptBtn.DataBind.SetBound(True, "", "VCHEQUE")
        oOptBtn.GroupWith("op_Cash")
        oPayForm.Items.Item("op_Cash").Specific.Selected = True
        oPayForm.Items.Item("ed_PayType").Specific.Value = "Cash"

        oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
        If oPayForm.Items.Item("ed_VocNo").Specific.Value = "" Then
            If (oMatrix.RowCount > 0) Then
                If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                    oPayForm.Items.Item("ed_VocNo").Specific.Value = 1
                Else
                    oPayForm.Items.Item("ed_VocNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                End If
            Else
                oPayForm.Items.Item("ed_VocNo").Specific.Value = 1
            End If
        End If

        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = oPayForm.Items.Item("cb_PayCur").Specific
        If oCombo.ValidValues.Count = 0 Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT CurrCode FROM OCRN")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, "")
                    oRecordSet.MoveNext()
                End While
            End If
        End If
        Dim oColumn As SAPbouiCOM.Column
        oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
        If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
        oColumn = oMatrix.Columns.Item("colChCode1")
        AddChooseFromList(oPayForm, "ChCode", False, "UDOCHCODE")
        oColumn.ChooseFromListUID = "ChCode"
        oColumn.ChooseFromListAlias = "Code"
        oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
        DisableChargeMatrix(oPayForm, oMatrix, True)
        oPayForm.Items.Item("ed_VedName").Specific.Active = True
        oPayForm.Freeze(False)
    End Sub

    Private Function AddNewRow(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        AddNewRow = False
        Dim sErrDesc As String = vbNullString
        Dim oDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
            oDbDataSource = oActiveForm.DataSources.DBDataSources.Item("@OBT_LCL17_VDETAIL")
            With oDbDataSource
                .SetValue("LineId", .Offset, oMatrix.RowCount + 1)
                .SetValue("U_VSeqNo", .Offset, oMatrix.RowCount + 1)
                .SetValue("U_ChCode", .Offset, String.Empty)
                .SetValue("U_AccCode", .Offset, String.Empty)
                .SetValue("U_ChDesc", .Offset, String.Empty)
                .SetValue("U_Amount", .Offset, String.Empty)
                .SetValue("U_GST", .Offset, String.Empty)
                .SetValue("U_GSTAmt", .Offset, String.Empty)
                .SetValue("U_NoGST", .Offset, String.Empty)
                .SetValue("U_ItemCode", .Offset, String.Empty)
            End With
            oMatrix.AddRow()
            AddGSTComboData(oMatrix.Columns.Item("colGST1"))
            If oActiveForm.Items.Item("cb_GST").Specific.Value = "No" Then
                Dim oCombo As SAPbouiCOM.ComboBox
                oCombo = oMatrix.Columns.Item("colGST1").Cells.Item(oMatrix.RowCount).Specific
                oCombo.Select("None", SAPbouiCOM.BoSearchKey.psk_ByValue)
                'oMatrix.Columns.Item("colGST").Cells.Item(oMatrix.RowCount).Specific.Value = "None"
            End If
            AddNewRow = True
        Catch ex As Exception
            AddNewRow = False
        End Try
    End Function

    Private Sub DisableChargeMatrix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)
        Try
            pMatrix.Columns.Item("colChCode1").Editable = pValue
            pMatrix.Columns.Item("colChCode1").BackColor = 16645886
            pMatrix.Columns.Item("colVDesc1").Editable = pValue
            pMatrix.Columns.Item("colVDesc1").BackColor = 16645886
            pMatrix.Columns.Item("colAmount1").Editable = pValue
            pMatrix.Columns.Item("colAmount1").BackColor = 16645886
            pMatrix.Columns.Item("colGST1").Editable = pValue
            pMatrix.Columns.Item("colGST1").BackColor = 16645886
        Catch ex As Exception

        End Try
    End Sub

    '======================================= Utilities ====================================
    Private Function GetNewKey(ByVal ObjectCode As String, ByRef oRecordSet As SAPbobsCOM.Recordset) As String
        '=============================================================================
        'Function   : GetNewKey()
        'Purpose    : This function get net next number of Manage Series, from
        '             ONNM and NNM1 tables
        'Parameters : ByVal ObjectCode As String
        '               ObjectCode = intended UDO code
        '             ByVal oRecordSet As SAPbobsCOM.Recordset
        '               oRecordSet = recordset object for which recordset to be created
        'Return     : next number from manage series
        'Author     : @channyeinkyaw
        '=============================================================================
        GetNewKey = vbNullString
        Try
            oRecordSet.DoQuery(String.Format("SELECT NNM1.NextNumber FROM ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode AND ONNM.DfltSeries = NNM1.Series) WHERE ONNM.ObjectCode = {0} ", FormatString(ObjectCode)))
            If oRecordSet.RecordCount > 0 Then
                GetNewKey = oRecordSet.Fields.Item("NextNumber").Value.ToString
            End If
        Catch ex As Exception
            GetNewKey = vbNullString
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub DeleteMatrixRow(ByRef oActiveForm As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix, ByVal objDataSource As String, ByVal oColumn As String)
        Dim tblname As String = objDataSource.Substring(1, objDataSource.Length - 1)
        Dim lRow As Long
        lRow = oMatrix.GetNextSelectedRow
        If lRow > -1 Then
            oActiveForm.DataSources.DBDataSources.Item(objDataSource).RemoveRecord(0)
            Dim oUserTable As SAPbobsCOM.UserTable
            oUserTable = p_oDICompany.UserTables.Item(tblname)
            oUserTable.GetByKey(lRow)
            oUserTable.Remove()
            oUserTable = Nothing
            oMatrix.DeleteRow(lRow)
            SetMatrixSeqNo(oMatrix, oColumn)
        End If

        If lRow = 1 And oMatrix.RowCount = 0 Then
            oMatrix.FlushToDataSource()
            oMatrix.AddRow()
            oMatrix.Columns.Item(oColumn).Cells.Item(1).Specific.Value = 1
        End If
        oMatrix.FlushToDataSource()
    End Sub
#End Region

    Public Sub LoadImportSeaLCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        Dim ImportSeaLCLForm As SAPbouiCOM.Form
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sErrDesc As String = vbNullString
        Dim oOpt As SAPbouiCOM.OptionBtn
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oEditText As SAPbouiCOM.EditText
        Try
            If Not LoadFromXML(p_oSBOApplication, "ImportSeaLCLv1.srf") Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
            'BubbleEvent = False
            ImportSeaLCLForm.EnableMenu("1288", True)
            ImportSeaLCLForm.EnableMenu("1289", True)
            ImportSeaLCLForm.EnableMenu("1290", True)
            ImportSeaLCLForm.EnableMenu("1291", True)
            ImportSeaLCLForm.EnableMenu("1284", False)
            ImportSeaLCLForm.EnableMenu("1286", False)
            If FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            End If
            ImportSeaLCLForm.Freeze(True)
            ImportSeaLCLForm.DataBrowser.BrowseBy = "ed_DocNum"
            ImportSeaLCLForm.Items.Item("fo_Prmt").Specific.Select()

            '---------------------------------------------- Vouncher (OMM) 7-Mar-2011 -------------------------------------------------------------------------'
            Dim oColumn As SAPbouiCOM.Column
            oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
            ImportSeaLCLForm.DataSources.DataTables.Add("DTCharges")
            dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
            dtmatrix.Columns.Add("ChCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
            dtmatrix.Columns.Add("AcCode", SAPbouiCOM.BoFieldsType.ft_Text)
            dtmatrix.Columns.Add("Desc", SAPbouiCOM.BoFieldsType.ft_Text)
            dtmatrix.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Price)
            dtmatrix.Columns.Add("GST", SAPbouiCOM.BoFieldsType.ft_Text)
            dtmatrix.Columns.Add("GSTAmt", SAPbouiCOM.BoFieldsType.ft_Price)
            dtmatrix.Columns.Add("NoGST", SAPbouiCOM.BoFieldsType.ft_Price)
            dtmatrix.Columns.Add("SeqNo", SAPbouiCOM.BoFieldsType.ft_Text)
            dtmatrix.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text) 'To Add Item Code

            oColumn = oMatrix.Columns.Item("colGST")
            AddGSTComboData(oColumn)
            oColumn.DataBind.Bind("DTCharges", "GST")
            oColumn = oMatrix.Columns.Item("colChCode")
            oColumn.DataBind.Bind("DTCharges", "ChCode")
            'AddUserDataSrc(oActiveForm, "test", "test", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 0)

            'TO DO to drop the comment
            'AddChooseFromList(ImportSeaLCLForm, "CCode", False, "UDOCHCODE")
            'oColumn.ChooseFromListUID = "CCode"

            oColumn = oMatrix.Columns.Item("colAcCode")
            oColumn.DataBind.Bind("DTCharges", "AcCode")
            oColumn = oMatrix.Columns.Item("colVDesc")
            oColumn.DataBind.Bind("DTCharges", "Desc")
            oColumn = oMatrix.Columns.Item("colAmount")
            oColumn.DataBind.Bind("DTCharges", "Amount")
            oColumn = oMatrix.Columns.Item("colGSTAmt")
            oColumn.DataBind.Bind("DTCharges", "GSTAmt")
            oColumn = oMatrix.Columns.Item("colNoGST")
            oColumn.DataBind.Bind("DTCharges", "NoGST")
            oColumn = oMatrix.Columns.Item("V_-1")
            oColumn.DataBind.Bind("DTCharges", "SeqNo")
            oColumn = oMatrix.Columns.Item("colICode")    'To Add Item Code
            oColumn.DataBind.Bind("DTCharges", "ItemCode") 'To Add Item Code

            dtmatrix.Rows.Add(1)
            oMatrix.Clear()
            dtmatrix.SetValue("SeqNo", 0, 1)
            oMatrix.LoadFromDataSource()
            '-------------------------------------Charge Code--------------------------------------------------------------'
            'Vendor in Voucher
            If AddChooseFromList(ImportSeaLCLForm, "PAYMENT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm.Items.Item("ed_VedName").Specific.ChooseFromListUID = "PAYMENT"
            ImportSeaLCLForm.Items.Item("ed_VedName").Specific.ChooseFromListAlias = "CardName"
            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
            'End

            If AddUserDataSrc(ImportSeaLCLForm, "CASH", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "CHEQUE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oOpt = ImportSeaLCLForm.Items.Item("op_Cash").Specific
            oOpt.DataBind.SetBound(True, "", "CASH")
            oOpt = ImportSeaLCLForm.Items.Item("op_Cheq").Specific
            oOpt.DataBind.SetBound(True, "", "CHEQUE")
            oOpt.GroupWith("op_Cash")
            'OMM voucher
            oCombo = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("Select BankName From ODSC")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("BankName").Value, "")
                    oRecordSet.MoveNext()
                End While
            End If
            oCombo = ImportSeaLCLForm.Items.Item("cb_PayCur").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT CurrCode FROM OCRN")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, "")
                    oRecordSet.MoveNext()
                End While
            End If
            '------------------------------------------------------------------------------------------------------------------------------------------'
            ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
            EnabledHeaderControls(ImportSeaLCLForm, False)
            EnabledMaxtix(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("mx_TkrList").Specific, False)
            ImportSeaLCLForm.PaneLevel = 7

            If FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If Not Title = vbNullString Then
                    ImportSeaLCLForm.Title = Title
                End If
                ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Import Sea LCL"
                'ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN ").SetValue("U_TranMode", 0, "Sea") 'MSW LCL Change
                ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                ImportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                ImportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
            End If

            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddChooseFromList(ImportSeaLCLForm, "cflBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "cflBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ImportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListUID = "cflBP"
            ImportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListAlias = "CardCode"
            ImportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListUID = "cflBP2"
            ImportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListAlias = "CardName"

            If AddChooseFromList(ImportSeaLCLForm, "cflBP3", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "WRHSE", False, "WAREHOUSE") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "DSVES01", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "DSVES02", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListUID = "WRHSE"
            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListAlias = "Code"
            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
            ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
            ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
            '------------------------------- For Cargo Tab OMM & SYMA ------------------------------------------------'13 Jan 2011

            AddChooseFromList(ImportSeaLCLForm, "cflCurCode", False, 37)
            ImportSeaLCLForm.Items.Item("ed_CurCode").Specific.ChooseFromListUID = "cflCurCode"
            '----------------------------------For Invoice Tab------------------------------------------------------'
            AddChooseFromList(ImportSeaLCLForm, "cflCurCode1", False, 37)
            oEditText = ImportSeaLCLForm.Items.Item("ed_CCharge").Specific
            oEditText.ChooseFromListUID = "cflCurCode1"
            AddChooseFromList(ImportSeaLCLForm, "cflCurCode2", False, 37)
            oEditText = ImportSeaLCLForm.Items.Item("ed_Charge").Specific
            oEditText.ChooseFromListUID = "cflCurCode2"
            '-------------------------------------------------------------------------------------------------------'

            oCombo = ImportSeaLCLForm.Items.Item("cb_PCode").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT Code, Name FROM [@OBT_TB004_PORTLIST]")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = ImportSeaLCLForm.Items.Item("cb_PType").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT PkgType FROM OPKG")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("PkgType").Value, "")
                    oRecordSet.MoveNext()
                End While
            End If

            oEditText = ImportSeaLCLForm.Items.Item("ed_JobNo").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEALCL'")
            If oRecordSet.RecordCount > 0 Then
                'ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW for Job Type Table
                ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW for Job Type Table
            End If

            'fortruckingtab
            If AddUserDataSrc(ImportSeaLCLForm, "TKRINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKREXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "DSINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "DSEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oOpt = ImportSeaLCLForm.Items.Item("op_Inter").Specific
            oOpt.DataBind.SetBound(True, "", "DSINTR")
            oOpt = ImportSeaLCLForm.Items.Item("op_Exter").Specific
            oOpt.DataBind.SetBound(True, "", "DSEXTR")
            oOpt.GroupWith("op_Inter")

            If AddUserDataSrc(ImportSeaLCLForm, "TKRDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "INSDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.DataBind.SetBound(True, "", "INSDATE")
            ImportSeaLCLForm.Items.Item("ed_TkrDate").Specific.DataBind.SetBound(True, "", "TKRDATE")
            If AddUserDataSrc(ImportSeaLCLForm, "TKRATTE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRTEL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRFAX", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRMAIL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRCOL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRTO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
            ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
            ImportSeaLCLForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
            ImportSeaLCLForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
            ImportSeaLCLForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
            ImportSeaLCLForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")

            If AddUserDataSrc(ImportSeaLCLForm, "DSDISP", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.DataBind.SetBound(True, "", "DSDISP")
            ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.DataBind.SetBound(True, "", "DSDISP")
            ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.GroupWith("op_DspIntr")

            If AddChooseFromList(ImportSeaLCLForm, "CFLTKRE", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "CFLTKRV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            '---------------------------10-1-2011-------------------------------------
            '----------Recordset for Binding colCType of Matrix (mx_Cont)-------------
            '--------------------------SYMA & OMM-------------------------------------
            oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
            Dim oColum As SAPbouiCOM.Column
            oColum = oMatrix.Columns.Item("colCType")
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select U_Type  from [@OBT_LCL08_PCONTAINE] group by U_Type")
            While oColum.ValidValues.Count > 0
                oColum.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
            End While
            If oRecordSet.RecordCount > 0 Then
                Dim a As Integer = oRecordSet.RecordCount
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oColum.ValidValues.Add(oRecordSet.Fields.Item("U_Type").Value, " ")
                    oColum.DisplayDesc = False
                    oRecordSet.MoveNext()
                End While
            End If
            oMatrix.AddRow()
            oMatrix.Columns.Item("colCSeqNo").Cells.Item(1).Specific.Value = 1
            '------------------------------For License Info SYMA & OMM (13/Jan/2011)---------------'
            oMatrix = ImportSeaLCLForm.Items.Item("mx_License").Specific
            oMatrix.AddRow()
            oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
            '--------------------------------------------------------------------------------------'
            ImportSeaLCLForm.EnableMenu("1292", True)
            ImportSeaLCLForm.EnableMenu("1293", True)

            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                Dim tempItem As SAPbouiCOM.Item
                tempItem = ImportSeaLCLForm.Items.Item("ed_JobNo")
                tempItem.Enabled = True
                ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = JobNo
                ImportSeaLCLForm.Items.Item("1").Click()
                tempItem.Enabled = False
            End If
            
            Select Case ImportSeaLCLForm.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_AddIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_AddIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = True
            End Select
            ImportSeaLCLForm.Freeze(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Module