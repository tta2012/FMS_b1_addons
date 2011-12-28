Option Explicit On

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
Imports CrystalDecisions.CrystalReports.Engine

Module modExportSeaLCL
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Private VocNo, VendorName, PayTo, PayType, BankName, CheqNo, Status, Currency, PaymentDate, GST, PrepBy, ExRate, Remark As String
    Private Total, GSTAmt, SubTotal, vocTotal, gstTotal As Double
    Private vendorCode As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sql As String = ""
    Dim strAPInvNo, strOutPayNo As String
    Dim dtmatrix As SAPbouiCOM.DataTable
    Public gridindex As String

    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public Function DoExportSeaLCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoExportSeaLCLFormDataEvent
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
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim FunctionName As String = "DoExportSeaLCLFormDataEvent"
        Dim sKeyValue As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSQLQuery As String = String.Empty
        Dim oDocument As SAPbobsCOM.Documents
        Dim oXmlReader As XmlTextReader
        Dim sDocNum As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oChMatrix As SAPbouiCOM.Matrix
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", FunctionName)
            Select BusinessObjectInfo.FormTypeEx
                Case "SHIPPINGINV"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                            oShpForm = p_oSBOApplication.Forms.GetForm("SHIPPINGINV", 1)
                            oMatrix = ExportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                            If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_LCL18_SHPINV", True)
                            ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_LCL18_SHPINV", False)
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                    'Voucher POP UP
                Case "VOUCHER"
                    If BusinessObjectInfo.BeforeAction = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            Try
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                                oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
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
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                                oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
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
                        ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                        oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_TB009_EVOUCHER", True)
                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_TB009_EVOUCHER", False)
                        End If
                    End If

                Case "142"
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            Dim oExportSeaLCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
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
                            oExportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = sDocNum
                            oEditText = oExportSeaLCLForm.Items.Item("ed_Trucker").Specific
                            oEditText.DataBind.SetBound(True, "", "TKREXTR")
                            oEditText.ChooseFromListUID = "CFLTKRV"
                            oEditText.ChooseFromListAlias = "CardName"
                            oExportSeaLCLForm.DataSources.UserDataSources.Item("TKREXTR").ValueEx = sName
                            oExportSeaLCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = sAttention
                            oExportSeaLCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = sPhone
                            oExportSeaLCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = sFax
                            oExportSeaLCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = sMail
                        End If
                    End If

                Case "EXPORTSEALCL"
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            LoadHolidayMarkUp(ExportSeaLCLForm)
                            ExportSeaLCLForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty 'MSW
                            If Not String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value) Then
                                If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            If Not String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_DspDate").Specific.Value) Then
                                If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoExportSeaLCLFormDataEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaLCLFormDataEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()
        End Try
    End Function

    Public Function DoExportSeaLCLMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoExportSeaLCLMenuEvent
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
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoExportSeaLCLMenuEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            Select Case pVal.MenuUID
                Case "menuExportSeaLCL"
                    If pVal.BeforeAction = False Then
                        LoadExportSeaFCLForm()
                    End If
            End Select

            DoExportSeaLCLMenuEvent = RTN_SUCCESS
        Catch ex As Exception
            DoExportSeaLCLMenuEvent = RTN_ERROR
        End Try
    End Function

    Public Function DoExportSeaLCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoExportSeaLCLItemEvent
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
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oChMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing

        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoExportSeaLCLItemEvent()"
        Dim sql As String = ""
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Select Case pVal.FormTypeEx
                Case "EXPORTSEALCL"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        If pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize" And pVal.Before_Action = True And pVal.Row <> 0 Then
                            oMatrix = ExportSeaLCLForm.Items.Item("mx_Cont").Specific
                            Dim oColCombo As SAPbouiCOM.Column
                            Dim omatCombo As SAPbouiCOM.ComboBox
                            oColCombo = oMatrix.Columns.Item("colCSize")
                            omatCombo = oColCombo.Cells.Item(pVal.Row).Specific
                            If Not String.IsNullOrEmpty(oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Value) Then
                                Dim type As String = oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Selected.Value
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("select U_Size from [@OBT_TB008_CONTAINER] where U_Type='" & type & "'")                       '* Change Nyan Lin   "[@OBT_TB008_CONTAINER]" 
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
                    End If


                    '-------------------------For Payment(omm)------------------------------------------'
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.Before_Action = False Then

                        If (pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colGST") Then
                            CalRate(ExportSeaLCLForm, pVal.Row)
                        End If
                        If (pVal.ItemUID = "cb_GST" And ExportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft") Then
                            oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
                            dtmatrix = ExportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                            dtmatrix.SetValue("GST", 0, "None")
                            oMatrix.LoadFromDataSource()
                        End If

                        If pVal.ItemUID = "cb_PayCur" Then
                            If ExportSeaLCLForm.Items.Item("cb_PayCur").Specific.Value <> "SGD" Then
                                Dim Rate As String = String.Empty
                                sql = "SELECT Rate FROM ORTT WHERE Currency = '" & ExportSeaLCLForm.Items.Item("cb_PayCur").Specific.Value & "' And DATENAME(YYYY,RateDate) = '" & _
                                        Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                        CInt(Today.ToString("dd"))
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    Rate = oRecordSet.Fields.Item("Rate").Value
                                End If
                                ExportSeaLCLForm.Items.Item("ed_PayRate").Enabled = True
                                ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER").SetValue("U_ExRate", 0, Rate.ToString) ' * Change 
                            Else
                                ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER").SetValue("U_ExRate", 0, Nothing)       '* change
                                ExportSeaLCLForm.Items.Item("ed_PayRate").Enabled = False
                                ExportSeaLCLForm.Items.Item("cb_PayCur").Specific.Active = True
                            End If

                        End If
                    End If
                    '-----------------------------------------------------------------------------------'

                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                            If Not RemoveFromAppList(ExportSeaLCLForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                        End If

                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                            Try
                                ExportSeaLCLForm.Items.Item("ed_Code").Enabled = False
                                ExportSeaLCLForm.Items.Item("ed_Name").Enabled = False
                                ExportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False
                            Catch ex As Exception
                            End Try
                        End If

                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            EnabledTrucker(ExportSeaLCLForm, False)
                        End If

                        '-------------------------For Payment(omm)------------------------------------------'
                        If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            CalRate(ExportSeaLCLForm, pVal.Row)
                        End If
                        '----------------------------------------------------------------------------------'

                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "fo_Prmt"
                                        ExportSeaLCLForm.PaneLevel = 7
                                        ExportSeaLCLForm.Items.Item("fo_PMain").Specific.Select()
                                    Case "fo_Dsptch"
                                        ExportSeaLCLForm.PaneLevel = 4
                                        ExportSeaLCLForm.Items.Item("ed_DspDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ExportSeaLCLForm.Items.Item("ed_DspDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        ExportSeaLCLForm.Items.Item("ed_DspHr").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                                        ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                    Case "fo_Trkng"
                                        ExportSeaLCLForm.PaneLevel = 6
                                        ExportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                            Else
                                                ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                            End If
                                        End If


                                    Case "fo_Vchr"
                                        ExportSeaLCLForm.PaneLevel = 20
                                        ExportSeaLCLForm.Items.Item("fo_VoView").Specific.Select()
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            Else
                                                '    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                    Case "fo_VoView"
                                        ExportSeaLCLForm.PaneLevel = 20


                                    Case "fo_VoEdit"
                                        ExportSeaLCLForm.PaneLevel = 21
                                        ExportSeaLCLForm.Items.Item("cb_BnkName").Enabled = False
                                        ExportSeaLCLForm.Items.Item("ed_Cheque").Enabled = False
                                        ExportSeaLCLForm.Items.Item("ed_PayRate").Enabled = False
                                    Case "fo_TkrView"
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If ExportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If (oMatrix.Columns.Item("V_1").Cells.Item(1).Specific.Value = "") Then
                                                ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                            Else
                                                ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                            End If
                                        End If

                                        ExportSeaLCLForm.PaneLevel = 6
                                    Case "fo_TkrEdit"
                                        ExportSeaLCLForm.PaneLevel = 5
                                        ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        If ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then

                                            oRecordSet.DoQuery("SELECT Address FROM OCRD WHERE CardCode = '" & ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                            If oRecordSet.RecordCount > 0 Then
                                                ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("Address").Value.ToString
                                            End If

                                            ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            ExportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True
                                        End If

                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If ExportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = "" Then
                                            ExportSeaLCLForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty
                                            If (oMatrix.RowCount > 0) Then
                                                If (oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                    ExportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                                Else
                                                    ExportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                                End If
                                            Else
                                                ExportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                            End If
                                            ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                        End If


                                    Case "fo_PMain"
                                        ExportSeaLCLForm.PaneLevel = 8
                                    Case "fo_PCargo"
                                        ExportSeaLCLForm.PaneLevel = 9
                                    Case "fo_PCon"
                                        ExportSeaLCLForm.PaneLevel = 10
                                    Case "fo_PInv"
                                        ExportSeaLCLForm.PaneLevel = 11
                                    Case "fo_PLic"
                                        ExportSeaLCLForm.PaneLevel = 12
                                    Case "fo_PAttach"
                                        ExportSeaLCLForm.PaneLevel = 13
                                    Case "fo_PTotal"
                                        ExportSeaLCLForm.PaneLevel = 14

                                    Case "fo_ShpInv"
                                        ExportSeaLCLForm.PaneLevel = 25
                                        ExportSeaLCLForm.Items.Item("fo_ShView").Specific.Select()
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                            Else
                                                '    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                    Case "fo_ShView"
                                        ExportSeaLCLForm.PaneLevel = 25
                                    Case "fo_BkVsl"
                                        ExportSeaLCLForm.PaneLevel = 26
                                    Case "fo_Crate"
                                        ExportSeaLCLForm.PaneLevel = 27
                                    Case "fo_Fumi"
                                        ExportSeaLCLForm.PaneLevel = 28
                                    Case "fo_OpBunk"
                                        ExportSeaLCLForm.PaneLevel = 29
                                    Case "fo_ArmEs"
                                        ExportSeaLCLForm.PaneLevel = 30
                                    Case "fo_Crane"
                                        ExportSeaLCLForm.PaneLevel = 31
                                    Case "fo_Fork"
                                        ExportSeaLCLForm.PaneLevel = 32

                                End Select

                                If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    ExportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                                    'ImportSeaLCLForm.Items.Item("ed_Wrhse").Enabled = True
                                End If

                                If pVal.ItemUID = "ch_POD" Then
                                    If ExportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = True Then
                                        ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
                                    End If
                                    If ExportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = False Then
                                        ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AddLic" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_License").Specific
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB014_EPLICINFO").Clear()                          '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("colLicNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                End If

                                If pVal.ItemUID = "bt_DelLic" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_License").Specific
                                    Dim lRow As Long
                                    lRow = oMatrix.GetNextSelectedRow
                                    If lRow > -1 Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB014_EPLICINFO").RemoveRecord(lRow - 1)   '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("@OBT_TB014_EPLICINFO")                                '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
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
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Cont").Specific
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB012_EPCONT").Clear()                         '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("colCSeqNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                End If

                                If pVal.ItemUID = "bt_DelCon" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Cont").Specific
                                    Dim lRow As Long
                                    lRow = oMatrix.GetNextSelectedRow
                                    If lRow > -1 Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB012_EPCONT").RemoveRecord(lRow - 1)    '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                            'oMatrix.AddRow(1)
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("OBT_TB012_EPCONT")                                  '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
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




                                If pVal.ItemUID = "ch_Dsp" Then
                                    If ExportSeaLCLForm.Items.Item("ch_Dsp").Specific.Checked = True Then
                                        ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    If ExportSeaLCLForm.Items.Item("ch_Dsp").Specific.Checked = False Then
                                        ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific.Active = True


                                    End If
                                End If

                                If pVal.ItemUID = "ed_ETADate" And ExportSeaLCLForm.Items.Item("ed_ETADate").Specific.String <> String.Empty Then
                                    If DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value
                                    If DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "ed_JbDate" And ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                                    If DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspDate" And ExportSeaLCLForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                                    If DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspCDte" And ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                                    If DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "bt_Payment" Then
                                    p_oSBOApplication.ActivateMenuItem("2818")
                                End If

                                If pVal.ItemUID = "op_Inter" Then
                                    ExportSeaLCLForm.Freeze(True)
                                    ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                    ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                    If AddChooseFromListByOption(ExportSeaLCLForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                    ExportSeaLCLForm.Freeze(False)
                                ElseIf pVal.ItemUID = "op_Exter" Then
                                    ExportSeaLCLForm.Freeze(True)
                                    ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = True
                                    If AddChooseFromListByOption(ExportSeaLCLForm, False, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaLCLForm.Items.Item("ed_PONo").Specific.Active = True
                                    ExportSeaLCLForm.Freeze(False)
                                End If

                                If pVal.ItemUID = "op_DspIntr" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If ClearComboData(ExportSeaLCLForm, "cb_Dspchr", "@OBT_TB007_EDISPATCH", "U_Dispatch") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
                                    Dim oCombo As SAPbouiCOM.ComboBox = ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If ClearComboData(ExportSeaLCLForm, "cb_Dspchr", "@OBT_TB007_EDISPATCH", "U_Dispatch") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
                                    p_oSBOApplication.ActivateMenuItem("6913") 'MSW 04-04-2011
                                    p_oSBOApplication.Menus.Item("6913").Activate()
                                    'SBO_Application.Menus.Item("6913").Activate()
                                    Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                                    UDFAttachForm.Items.Item("U_JobNo").Specific.Value = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    UDFAttachForm.Items.Item("U_InsDate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value
                                End If

                                If pVal.ItemUID = "1" Then
                                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If pVal.ActionSuccess = True Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("ed_Code").Enabled = False 'MSW
                                            ExportSeaLCLForm.Items.Item("ed_Name").Enabled = False
                                            ExportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False
                                            Dim JobLastDocEntry As Integer
                                            Dim ObjectCode As String = String.Empty

                                            'p_oDICompany.GetNewObjectCode(ObjectCode)
                                            'ObjectCode = p_oDICompany.GetNewObjectKey()
                                            sql = "select top 1 Docentry from [@OBT_TB002_EXPORT] order by docentry desc"
                                            oRecordSet.DoQuery(sql)
                                            Dim FrDocEntry As Integer = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString)

                                            sql = "select top 1 Docentry from [@OBT_FREIGHTDOCNO] order by docentry desc"
                                            oRecordSet.DoQuery(sql)
                                            If oRecordSet.Fields.Item("Docentry").Value.ToString = "" Then
                                                JobLastDocEntry = 1
                                            Else
                                                JobLastDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString) + 1
                                            End If



                                            sql = "Insert Into [@OBT_FREIGHTDOCNO] (DocEntry,DocNum,U_JobNo,U_JobMode,U_JobType,U_JbStus,U_FrDocNo,U_JbDate,U_ObjType,U_CusCode,U_CusName,U_ShpCode,U_ShpName) Values " & _
                                                "(" & JobLastDocEntry & _
                                                    "," & JobLastDocEntry & _
                                                   "," & IIf(ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value), "NULL") & _
                                                    "," & IIf(ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value), "NULL") & _
                                                    "," & IIf(ExportSeaLCLForm.Items.Item("cb_JobType").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("cb_JobType").Specific.Value), "NULL") & _
                                                    "," & IIf(ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value), "NULL") & _
                                                    "," & FrDocEntry & _
                                                     "," & IIf(ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value), "Null") & _
                                                    "," & IIf(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString() <> "", FormatString(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString()), "Null") & _
                                                     "," & IIf(ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value), "Null") & _
                                                      "," & IIf(ExportSeaLCLForm.Items.Item("ed_Name").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_Name").Specific.Value), "Null") & _
                                                       "," & IIf(ExportSeaLCLForm.Items.Item("ed_V").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_V").Specific.Value), "Null") & _
                                                    "," & IIf(ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value), "Null") & ")"
                                            oRecordSet.DoQuery(sql)

                                            sql = "Update [@OBT_TB002_EXPORT] set U_JbDocNo=" & JobLastDocEntry & " Where DocEntry=" & FrDocEntry & ""
                                            oRecordSet.DoQuery(sql)

                                            'p_oDICompany.GetNewObjectCode(ObjectCode)
                                            'ObjectCode = p_oDICompany.GetNewObjectKey()
                                            'Dim JobDocEntry As Integer = Convert.ToInt32(ObjectCode)

                                        End If
                                    Else
                                        sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                                        oRecordSet.DoQuery(sql)

                                    End If
                                End If

                                If pVal.ItemUID = "bt_AddIns" Then
                                    If String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.String) Then
                                        p_oSBOApplication.SetStatusBarMessage("Must Fill Trucker", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        'If ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = vbNullString Then
                                        If ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then
                                            modTrucking.AddUpdateInstructions(ExportSeaLCLForm, oMatrix, "@OBT_TB006_ETRUCKING", True)     '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                        Else
                                            modTrucking.AddUpdateInstructions(ExportSeaLCLForm, oMatrix, "@OBT_TB006_ETRUCKING", False)    '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                            ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                            ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                        End If
                                        ClearText(ExportSeaLCLForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                        ExportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                    End If
                                End If

                                If pVal.ItemUID = "bt_DelIns" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    modTrucking.DeleteByIndex(ExportSeaLCLForm, oMatrix, "@OBT_TB006_ETRUCKING")                             '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                    ClearText(ExportSeaLCLForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                    ExportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                End If

                                If pVal.ItemUID = "bt_AmdIns" Then
                                    ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Update Trucking Instruction"
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific

                                    If (oMatrix.GetNextSelectedRow < 0) Then
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction", 1, "OK")
                                        BubbleEvent = False
                                        Exit Function
                                    Else
                                        modTrucking.GetDataFromMatrixByIndex(ExportSeaLCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        modTrucking.SetDataToEditTabByIndex(ExportSeaLCLForm)
                                        ExportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                        ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True 'MSW
                                    End If
                                    ExportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                End If

                                If pVal.ItemUID = "fo_TkrView" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    If oMatrix.RowCount > 1 Then
                                        ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                        Else
                                            ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                                        End If
                                    End If
                                    ClearText(ExportSeaLCLForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                    ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                    ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                End If

                                '-------------------------Payment Vouncher (OMM)------------------------------------------------------'
                                If pVal.ItemUID = "fo_VoEdit" Then
                                    ExportSeaLCLForm.Freeze(True)
                                    ExportSeaLCLForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    ExportSeaLCLForm.Items.Item("op_Cash").Specific.Selected = True
                                    ExportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    ExportSeaLCLForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    'ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_VOUCHER").SetValue("U_DocNo", 0, ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value)

                                    If ExportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                        dtmatrix = ExportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                        oMatrix.Clear()
                                    End If

                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = "" Then
                                        If (oMatrix.RowCount > 0) Then
                                            If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = 1
                                            Else
                                                ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                            End If
                                        Else
                                            ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value = 1
                                        End If
                                    End If

                                    Dim oCombo As SAPbouiCOM.ComboBox
                                    oCombo = ExportSeaLCLForm.Items.Item("cb_PayCur").Specific
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
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                    If dtmatrix.Rows.Count = 0 Then
                                        RowAddToMatrix(ExportSeaLCLForm, oMatrix)
                                    End If
                                    ExportSeaLCLForm.Items.Item("ed_VedName").Specific.Active = True
                                    ExportSeaLCLForm.Freeze(False)
                                End If

                                If pVal.ItemUID = "fo_VoView" Then
                                    'oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    'If oMatrix.RowCount > 1 Then
                                    '    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    'ElseIf oMatrix.RowCount = 1 Then
                                    '    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    '        ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    '    Else
                                    '        '  ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                    '    End If
                                    'End If
                                    'ClearText(ExportSeaLCLForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

                                    'Dim oComboBank As SAPbouiCOM.ComboBox
                                    'Dim oComboCurrency As SAPbouiCOM.ComboBox

                                    'oComboBank = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    'oComboCurrency = ExportSeaLCLForm.Items.Item("cb_PayCur").Specific

                                    'For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                    '    oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    'Next
                                    'For j As Integer = oComboCurrency.ValidValues.Count - 1 To 0 Step -1
                                    '    oComboCurrency.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    'Next
                                    ' ExportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" 'MSW 23-03-2011
                                End If
                                If pVal.ItemUID = "bt_Draft" Then
                                    ''''' Insert Voucher Item Table
                                    Dim DocEntry As Integer
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific                                        '* Change Nyan Lin   "[@[OBT_TB021_VOUCHER]]"
                                    oRecordSet.DoQuery("Delete From [@OBT_LCL15_VOUCHER] Where U_JobNo='" & ExportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value & "' And U_PVNo='" & ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value & "'")
                                    oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [@OBT_LCL15_VOUCHER] ")      '* Change Nyan Lin   "[@[OBT_TB021_VOUCHER]]"
                                    If oRecordSet.RecordCount > 0 Then
                                        DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                    End If
                                    'omm - Purchase Voucher Save To Draft 09-03-2010
                                    vocTotal = Convert.ToDouble(ExportSeaLCLForm.Items.Item("ed_Total").Specific.Value)
                                    gstTotal = Convert.ToDouble(ExportSeaLCLForm.Items.Item("ed_GSTAmt").Specific.Value)
                                    If ExportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                        SaveToPurchaseVoucher(ExportSeaLCLForm, True)
                                    Else
                                        SaveToPurchaseVoucher(ExportSeaLCLForm, False)
                                    End If
                                    SaveToDraftPurchaseVoucher(ExportSeaLCLForm)
                                    'End Purchase Voucher Save To Draft 09-03-2010
                                    dtmatrix = ExportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                                    If oMatrix.RowCount > 0 Then
                                        For i As Integer = oMatrix.RowCount To 1 Step -1
                                            'To Add Item Code in Insert Statement
                                            '* Change Nyan Lin   "[@[OBT_TB021_VOUCHER]]"
                                            sql = "Insert Into [@OBT_LCL15_VOUCHER] (DocEntry,LineID,U_JobNo,U_PVNo,U_VSeqNo,U_ChCode,U_AccCode,U_ChDes,U_Amount,U_GST,U_GSTAmt,U_NoGST,U_ChrgCode) Values " & _
                                                   "(" & DocEntry & _
                                                    "," & i & _
                                                    "," & IIf(ExportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value), "NULL") & _
                                                    "," & IIf(ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value <> "", FormatString(ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value), "Null") & _
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

                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If Convert.ToInt32(ExportSeaLCLForm.Items.Item("ed_VocNo").Specific.Value) > oMatrix.RowCount Then
                                        AddUpdateVoucher(ExportSeaLCLForm, oMatrix, "@OBT_TB009_EVOUCHER", True)                               '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    Else
                                        AddUpdateVoucher(ExportSeaLCLForm, oMatrix, "@OBT_TB009_EVOUCHER", False)                              '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                        ExportSeaLCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft"
                                    End If

                                    ClearText(ExportSeaLCLForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    Dim oComboCurrency As SAPbouiCOM.ComboBox

                                    oComboBank = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    oComboCurrency = ExportSeaLCLForm.Items.Item("cb_PayCur").Specific

                                    For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                        oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next
                                    For j As Integer = oComboCurrency.ValidValues.Count - 1 To 0 Step -1
                                        oComboCurrency.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next

                                    ExportSeaLCLForm.Items.Item("fo_VoView").Specific.Select()
                                    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True

                                End If

                                If pVal.ItemUID = "op_Cash" Then
                                    ExportSeaLCLForm.Items.Item("ed_VedName").Specific.Active = True
                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    oComboBank = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                        oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next

                                    ExportSeaLCLForm.Items.Item("cb_BnkName").Enabled = False
                                    ExportSeaLCLForm.Items.Item("ed_Cheque").Enabled = False
                                End If

                                If pVal.ItemUID = "op_Cheq" Then
                                    Dim oComboBank As SAPbouiCOM.ComboBox
                                    oComboBank = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
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

                                    ExportSeaLCLForm.Items.Item("cb_BnkName").Enabled = True
                                    ExportSeaLCLForm.Items.Item("ed_Cheque").Enabled = True
                                End If

                                If pVal.ItemUID = "bt_Cancel" Then
                                    ExportSeaLCLForm.Items.Item("fo_VoView").Specific.Select()
                                End If

                                If pVal.ItemUID = "mx_Voucher" And pVal.ColUID = "V_-1" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If oMatrix.GetNextSelectedRow > 0 Then
                                        If (oMatrix.IsRowSelected(oMatrix.GetNextSelectedRow)) = True Then
                                            GetVoucherDataFromMatrixByIndex(ExportSeaLCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        End If
                                    Else
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Payment Voucher.", 1, "&OK")
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AmdVoc" Then
                                    'POP UP Payment Voucher
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                        LoadPaymentVoucher(ExportSeaLCLForm)
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Voucher.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If
                                'Shipping Invoice POP Up
                                If pVal.ItemUID = "bt_ShpInv" Then
                                    'POP UP Shipping Invoice
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                        LoadShippingInvoice(ExportSeaLCLForm)
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                '-----------------------------------------------------------------------------------------------------------'

                                If pVal.ItemUID = "mx_Cont" And ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If

                                If pVal.ItemUID = "bt_ChkList" Then
                                    Start(ExportSeaLCLForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                '-------------------------For Payment(omm)------------------------------------------'
                                If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "V_-1" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                    If pVal.Row > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            gridindex = CInt(pVal.Row)
                                        End If
                                    End If
                                End If
                                '----------------------------------------------------------------------------------'
                                If pVal.ItemUID = "mx_TkrList" Or pVal.ColUID = "V_1" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    If oMatrix.GetNextSelectedRow > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            modTrucking.rowIndex = CInt(pVal.Row)
                                            modTrucking.GetDataFromMatrixByIndex(ExportSeaLCLForm, oMatrix, modTrucking.rowIndex)
                                        End If
                                    Else
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction.", 1, "&OK")
                                    End If

                                End If



                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                                Try
                                    '-------------------------For Payment(omm)------------------------------------------'
                                    If pVal.ItemUID = "ed_VedName" Then
                                        ObjDBDataSource = ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER")                 '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                                                                                               & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString())
                                        vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft                                        '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    End If
                                    If pVal.ColUID = "colChCode" Then
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
                                        dtmatrix = ExportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
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
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString
                                        If String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value) Then
                                            ExportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            ExportSeaLCLForm.Items.Item("ed_ETADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific.Value = Now.ToString("HH:mm")
                                            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        If String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value) Then
                                            ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            ExportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            ExportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        EnabledHeaderControls(ExportSeaLCLForm, True)
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                            ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"

                                            ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                                            ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                                            ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                            ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                        End If
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_IComName", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_LCL06_PMAI]"
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_ShpAgt" Then
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)     '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)      '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)            '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString) '* Change Nyan Lin   "[@OBT_TB005_EPERMIT]"
                                        End If
                                    End If


                                    If pVal.ItemUID = "ed_Trucker" Then
                                        If ExportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                            ExportSeaLCLForm.Freeze(True)
                                            ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                            oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                            ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                            ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                            ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                            ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = "" '25-3-2011
                                            ExportSeaLCLForm.Freeze(False)
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_Vessel" Or pVal.ItemUID = "ed_Voy" Then
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)        '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)    '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_VName", 0, ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.String)                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    End If

                                    If pVal.ItemUID = "ed_CurCode" Then
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_CurCode", 0, oDataTable.GetValue(0, 0).ToString)
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_ExRate", 0, Rate.ToString)
                                    End If
                                    If pVal.ItemUID = "ed_CCharge" Then
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_Cchange", 0, oDataTable.GetValue(0, 0).ToString)                            '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_CEchange", 0, Rate.ToString)                                               '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                    If pVal.ItemUID = "ed_Charge" Then
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_FCchange", 0, oDataTable.GetValue(0, 0).ToString)                          '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                Catch ex As Exception
                                End Try

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If pVal.ItemUID = "cb_PCode" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ExportSeaLCLForm.Items.Item("cb_PCode").Specific
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                             '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_PCode", 0, oCombo.Selected.Value.ToString)                                       '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                End If
                                If pVal.ItemUID = "cb_BnkName" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim test As String = "select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description
                                    oRecordSet.DoQuery("select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description)
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER").SetValue("U_GLAC", 0, oRecordSet.Fields.Item("GLAccount").Value)                            '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    'oCombo.Selected.Description
                                End If

                                If pVal.ItemUID = "cb_PType" Then
                                    Dim oCombo As SAPbouiCOM.ComboBox = ExportSeaLCLForm.Items.Item("cb_PType").Specific
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_TUnit", 0, oCombo.Selected.Value.ToString)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN
                                If pVal.ItemUID = "ed_JobNo" And pVal.CharPressed = 13 Then
                                    ExportSeaLCLForm.Items.Item("1").Click()
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                'MSW 
                                If pVal.ItemUID = "ed_TkrDate" And pVal.Before_Action = False Then
                                    Dim strTime As SAPbouiCOM.EditText
                                    strTime = ExportSeaLCLForm.Items.Item("ed_TkrTime").Specific
                                    strTime.Value = Now.ToString("HH:mm")
                                End If
                                'End MSW
                                If pVal.ItemUID = "ed_InvNo" Then
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_InvNo", 0, ExportSeaLCLForm.Items.Item("ed_InvNo").Specific.String)         '* Change Nyan Lin   "[@OBT_TB0011_VOUCHER]"
                                End If
                                'NL LCL Change 24-03-2011
                                If pVal.ItemUID = "ed_NOP" Then
                                    ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_TotalOP", 0, ExportSeaLCLForm.Items.Item("ed_NOP").Specific.String)
                                End If
                                'End NL LCL Change 24-03-2011
                                If ExportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Validateforform(pVal.ItemUID, ExportSeaLCLForm)
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If BoolResize = False Then
                                    Try
                                        Dim oItemRet1 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("EXPORTSEALCL").Items.Item("rt_Outer")
                                        Dim oItemRetInner As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("EXPORTSEALCL").Items.Item("rt_Inner")
                                        oItemRetInner.Width = ExportSeaLCLForm.Items.Item("mx_Cont").Width + 15
                                        oItemRetInner.Height = ExportSeaLCLForm.Items.Item("mx_Cont").Height + 140
                                        oItemRet1.Top = ExportSeaLCLForm.Items.Item("mx_TkrList").Top - 29
                                        oItemRet1.Width = ExportSeaLCLForm.Items.Item("mx_TkrList").Width + 20
                                        oItemRet1.Height = ExportSeaLCLForm.Items.Item("mx_Voucher").Height + 90 '33
                                        ExportSeaLCLForm.Items.Item("155").Top = ExportSeaLCLForm.Items.Item("mx_TkrList").Top - 5
                                        ExportSeaLCLForm.Items.Item("155").Width = ExportSeaLCLForm.Items.Item("mx_TkrList").Width + 10
                                        ExportSeaLCLForm.Items.Item("155").Height = ExportSeaLCLForm.Items.Item("mx_TkrList").Height + 5
                                        BoolResize = True
                                    Catch ex As Exception
                                        ' p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try
                                ElseIf BoolResize = True Then
                                    Try
                                        Dim oItemRet2 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("EXPORTSEALCL").Items.Item("rt_Outer")
                                        oItemRet2.Top = ExportSeaLCLForm.Items.Item("mx_TkrList").Top - 29
                                        Dim oItemRetInner2 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("EXPORTSEALCL").Items.Item("rt_Inner")
                                        oItemRetInner2.Width = ExportSeaLCLForm.Items.Item("mx_Cont").Width + 15
                                        oItemRetInner2.Height = ExportSeaLCLForm.Items.Item("mx_Cont").Height + 140
                                        oItemRet2.Width = ExportSeaLCLForm.Items.Item("mx_TkrList").Width + 20
                                        oItemRet2.Height = ExportSeaLCLForm.Items.Item("mx_Voucher").Height + 90 '33
                                        ExportSeaLCLForm.Items.Item("155").Top = ExportSeaLCLForm.Items.Item("mx_TkrList").Top - 5
                                        ExportSeaLCLForm.Items.Item("155").Width = ExportSeaLCLForm.Items.Item("mx_TkrList").Width + 10
                                        ExportSeaLCLForm.Items.Item("155").Height = ExportSeaLCLForm.Items.Item("mx_TkrList").Height + 5
                                        BoolResize = False
                                    Catch ex As Exception
                                        'p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    End Try
                                End If
                        End Select
                    End If
                    If pVal.BeforeAction = True Then
                        Select Case pVal.EventType
                            'MSW for Job Type Table
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                If Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If pVal.ItemUID = "ed_JobNo" And pVal.InnerEvent = False Then
                                        ValidateJobNumber(ExportSeaLCLForm, BubbleEvent)
                                    End If
                                End If

                                'End MSW for Job Type Table
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                If pVal.ItemUID = "1" Then
                                    Dim PODFlag As String = String.Empty
                                    Dim JbStus As String = String.Empty
                                    Dim DispatchComplete As String = String.Empty
                                    If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        'develivery process by POD[Proof Of Delivery] check box
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        'oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_TB002_EXPORT] WHERE DocEntry = " & ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value)
                                        oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_TB002_EXPORT] WHERE DocEntry = " & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value) 'MSW 08-06-2011 for job Type Table
                                        If oRecordSet.RecordCount > 0 Then
                                            JbStus = oRecordSet.Fields.Item("U_JbStus").Value
                                            PODFlag = oRecordSet.Fields.Item("U_POD").Value
                                        End If
                                        If ExportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = True And JbStus = "Open" Then
                                            If p_oSBOApplication.MessageBox("Make sure that all entries trucking and vouchers are completed.(ensure no draft Payment in this job and " & _
                                                                       "ensure all external trucking transaction has generated the PO). Cannot edit or add after click POD check box. " & _
                                                                       "Do you want to continue?", 1, "&Yes", "&No") = 2 Then
                                                BubbleEvent = False
                                            End If
                                        End If
                                        If BubbleEvent = True Then
                                            'MSW 08-06-2011 for job type table
                                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            'End MSW 08-06-2011 for job type table
                                        End If
                                    End If
                                    If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        'handle for dispatch complete check box
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ' oRecordSet.DoQuery("SELECT U_Complete FROM [@OBT_TB007_EDISPATCH] WHERE DocEntry = " & ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value)'MSW 08-06-2011 for job Type Table
                                        oRecordSet.DoQuery("SELECT U_Complete FROM [@OBT_TB007_EDISPATCH] WHERE DocEntry = " & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value) 'MSW 08-06-2011 for job Type Table
                                        If oRecordSet.RecordCount > 0 Then
                                            DispatchComplete = oRecordSet.Fields.Item("U_Complete").Value
                                        End If
                                        If ExportSeaLCLForm.Items.Item("ch_Dsp").Specific.Checked = True And DispatchComplete = "Y" Then
                                            BubbleEvent = False
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

                                        End If
                                        If ExportSeaLCLForm.Items.Item("ch_POD").Specific.Checked = True And PODFlag = "Y" Then
                                            BubbleEvent = False
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        End If
                                    End If
                                    If ExportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And ExportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If Validateforform(" ", ExportSeaLCLForm) Then
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                        End Select
                    End If


            End Select

            DoExportSeaLCLItemEvent = RTN_SUCCESS
        Catch ex As Exception
            DoExportSeaLCLItemEvent = RTN_ERROR
        End Try
    End Function

    Public Function DoExportSeaLCLRightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoExportSeaLCLRightClickEvent
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
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
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
                        oMatrix = ExportSeaLCLForm.Items.Item("[Matrix Name]").Specific
                        If eventInfo.Row > 0 And (ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
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

                    Case "mx_ShpInv"
                        If (eventInfo.BeforeAction = True) Then
                            'Dim oMenuItem As SAPbouiCOM.MenuItem
                            'Dim oMenus As SAPbouiCOM.Menus
                            Try
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING

                                oCreationPackage.UniqueID = "EditShp"
                                oCreationPackage.String = "Edit Shipping Invoice"
                                oCreationPackage.Enabled = True

                                oMenuItem = p_oSBOApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("EditShp") Then
                                    p_oSBOApplication.Menus.RemoveEx("EditShp")
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

                            oCreationPackage.UniqueID = "EditShp"
                            oCreationPackage.String = "Edit Shipping Invoice"
                            oCreationPackage.Enabled = True
                            oMenuItem = p_oSBOApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            If oMenus.Exists("EditShp") Then
                                p_oSBOApplication.Menus.RemoveEx("EditShp")
                            End If

                        Catch ex As Exception
                            MessageBox.Show(ex.Message)
                        End Try
                End Select
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoExportSeaLCLRightClickEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaLCLRightClickEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()
        End Try
    End Function

    Public Sub LoadExportSeaFCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim sErrDesc As String = ""
        Dim sFuncName As String = "LoadExportSeaFCLForm"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            If Not LoadFromXML(p_oSBOApplication, "ExportSeaLCLv1.srf") Then Throw New ArgumentException(sErrDesc)
            ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
            ExportSeaLCLForm.EnableMenu("1288", True)
            ExportSeaLCLForm.EnableMenu("1289", True)
            ExportSeaLCLForm.EnableMenu("1290", True)
            ExportSeaLCLForm.EnableMenu("1291", True)
            ExportSeaLCLForm.EnableMenu("1284", False)
            ExportSeaLCLForm.EnableMenu("1286", False)
            ExportSeaLCLForm.DataBrowser.BrowseBy = "ed_DocNum"
            ExportSeaLCLForm.Items.Item("fo_Prmt").Specific.Select()
            ExportSeaLCLForm.Freeze(True)

            ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
            EnabledHeaderControls(ExportSeaLCLForm, False)
            EnabledMaxtix(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("mx_TkrList").Specific, False)
            ExportSeaLCLForm.PaneLevel = 7
            ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Sea LCL"
            ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
            ExportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
            ExportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, _
                             sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddChooseFromList(ExportSeaLCLForm, "cflBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ExportSeaLCLForm, "cflBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ExportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListUID = "cflBP"
            ExportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListAlias = "CardCode"
            ExportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListUID = "cflBP2"
            ExportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListAlias = "CardName"

            If AddChooseFromList(ExportSeaLCLForm, "cflBP3", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddChooseFromList(ExportSeaLCLForm, "DSVES01", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ExportSeaLCLForm, "DSVES02", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
            ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"

            ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
            ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
            ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
            ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"

            '-------------------------------For Cargo Tab OMM & SYMA------------------------------------------------'13 Jan 2011

            AddChooseFromList(ExportSeaLCLForm, "cflCurCode", False, 37)
            ExportSeaLCLForm.Items.Item("ed_CurCode").Specific.ChooseFromListUID = "cflCurCode"
            '----------------------------------For Invoice Tab------------------------------------------------------'
            AddChooseFromList(ExportSeaLCLForm, "cflCurCode1", False, 37)
            oEditText = ExportSeaLCLForm.Items.Item("ed_CCharge").Specific
            oEditText.ChooseFromListUID = "cflCurCode1"
            AddChooseFromList(ExportSeaLCLForm, "cflCurCode2", False, 37)
            oEditText = ExportSeaLCLForm.Items.Item("ed_Charge").Specific
            oEditText.ChooseFromListUID = "cflCurCode2"
            '-------------------------------------------------------------------------------------------------------'
            oCombo = ExportSeaLCLForm.Items.Item("cb_PCode").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT Code, Name FROM [@OBT_TB004_PORTLIST]")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = ExportSeaLCLForm.Items.Item("cb_PType").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT PkgType FROM OPKG")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("PkgType").Value, "")
                    oRecordSet.MoveNext()
                End While
            End If

            oEditText = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'EXPORTSEALCL'")
            If oRecordSet.RecordCount > 0 Then
                'ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW For JobType Table 
                ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW For JobType Table 
            End If

            'fortruckingtab
            If AddUserDataSrc(ExportSeaLCLForm, "TKRINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKREXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "DSINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "DSEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oOpt = ExportSeaLCLForm.Items.Item("op_Inter").Specific
            oOpt.DataBind.SetBound(True, "", "DSINTR")
            oOpt = ExportSeaLCLForm.Items.Item("op_Exter").Specific
            oOpt.DataBind.SetBound(True, "", "DSEXTR")
            oOpt.GroupWith("op_Inter")

            If AddUserDataSrc(ExportSeaLCLForm, "TKRDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "INSDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.DataBind.SetBound(True, "", "INSDATE")
            ExportSeaLCLForm.Items.Item("ed_TkrDate").Specific.DataBind.SetBound(True, "", "TKRDATE")
            If AddUserDataSrc(ExportSeaLCLForm, "TKRATTE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRTEL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRFAX", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRMAIL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRCOL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRTO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ExportSeaLCLForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
            ExportSeaLCLForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
            ExportSeaLCLForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
            ExportSeaLCLForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
            ExportSeaLCLForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
            ExportSeaLCLForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")

            If AddUserDataSrc(ExportSeaLCLForm, "DSDISP", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.DataBind.SetBound(True, "", "DSDISP")
            ExportSeaLCLForm.Items.Item("op_DspExtr").Specific.DataBind.SetBound(True, "", "DSDISP")
            ExportSeaLCLForm.Items.Item("op_DspExtr").Specific.GroupWith("op_DspIntr")

            If AddChooseFromList(ExportSeaLCLForm, "CFLTKRE", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ExportSeaLCLForm, "CFLTKRV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            '---------------------------10-1-2011-------------------------------------
            '----------Recordset for Binding colCType of Matrix (mx_Cont)-------------
            '--------------------------SYMA & OMM-------------------------------------
            oMatrix = ExportSeaLCLForm.Items.Item("mx_Cont").Specific
            Dim oColum As SAPbouiCOM.Column
            oColum = oMatrix.Columns.Item("colCType")
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select U_Type  from [@OBT_TB008_CONTAINER] group by U_Type")
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
            oMatrix = ExportSeaLCLForm.Items.Item("mx_License").Specific
            oMatrix.AddRow()
            oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
            '-------------------------------------------------------------------------------------'
            ExportSeaLCLForm.EnableMenu("1292", True)
            ExportSeaLCLForm.EnableMenu("1293", True)
            ExportSeaLCLForm.Freeze(False)
            Select Case ExportSeaLCLForm.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_AddIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_AddIns").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = True
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub EnabledHeaderControls(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)
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
        pForm.Items.Item("ed_JobNo").Enabled = pValue
        pForm.Items.Item("cb_JobType").Enabled = pValue
        pForm.Items.Item("cb_JbStus").Enabled = pValue
    End Sub

    Public Sub EnabledMaxtix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)
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
    End Sub

    Private Sub EnabledTrucker(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)
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
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_SubTotal", 0, SubTotal)
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_GSTAmt", 0, GSTTotal)
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_Total", 0, Total)
    End Sub

    Private Sub SetMatrixSeqNo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal ColName As String)
        For i As Integer = 1 To oMatrix.RowCount
            oMatrix.Columns.Item(ColName).Cells.Item(i).Specific.Value = i
        Next
    End Sub

    Private Function AddChooseFromListByOption(ByRef pForm As SAPbouiCOM.Form, ByVal pOption As Boolean, ByVal pObjID As String, ByVal pErrDesc As String) As Long
        Dim oEditText As SAPbouiCOM.EditText
        Try
            If pOption = True Then
                oEditText = pForm.Items.Item(pObjID).Specific
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

    Private Sub AddGSTComboData(ByVal oColumn As SAPbouiCOM.Column)
        Dim oForm As SAPbouiCOM.Form
        Dim RS As SAPbobsCOM.Recordset
        oForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
        RS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery("select Code,Name from ovtg Where Category='I'")
        RS.MoveFirst()
        For j As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
            oColumn.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oColumn.ValidValues.Add("None", "None")
        While RS.EoF = False
            oColumn.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
            RS.MoveNext()
        End While
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

    Private Sub SaveToPurchaseVoucher(ByRef pForm As SAPbouiCOM.Form, ByVal ProcessedState As Boolean)
        Dim nErr As Long
        Dim errMsg As String = String.Empty
        Dim ObjectCode As String = String.Empty
        Dim invDocEntry As Integer
        Dim Document As SAPbobsCOM.Documents
        Dim businessPartner As SAPbobsCOM.BusinessPartners
        Document = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        businessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        vendorCode = pForm.Items.Item("ed_VedCode").Specific.Value
        If ProcessedState = False Then
            oRecordSet.DoQuery("Select DocEntry From OPCH Where U_JobNo='" & pForm.Items.Item("ed_PJobNo").Specific.Value & "' And U_PVNo='" & pForm.Items.Item("ed_VocNo").Specific.Value & "' And U_FrDocNo='" & pForm.Items.Item("ed_DocNum").Specific.Value & "'")
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

                If oMatrix.RowCount > 0 Then
                    For i As Integer = 1 To oMatrix.RowCount
                        Document.Lines.ItemCode = oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value
                        Document.Lines.ItemDescription = oMatrix.Columns.Item("colVDesc1").Cells.Item(i).Specific.Value
                        Document.Lines.Quantity = 1
                        Document.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(i).Specific.Value)
                        'MSW 23-03-2011 For VatCode GST None or Blank in GST Field if we didn't assign ZI ,system auto populate default SI 
                        If oMatrix.Columns.Item("colGST1").Cells.Item(i).Specific.Value = "None" Then
                            'If (dtmatrix.GetValue("GST", i - 1) = "" Or dtmatrix.GetValue("GST", i - 1) = "None") Then
                            Document.Lines.VatGroup = "ZI"
                        Else
                            ' Document.Lines.VatGroup = dtmatrix.GetValue("GST", i - 1)
                            Document.Lines.VatGroup = oMatrix.Columns.Item("colGST1").Cells.Item(i).Specific.Value()
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

    Private Sub AddUpdateVoucher(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
        ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item(DataSource)
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
                    .SetValue("U_VDocNum", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value)

                    pMatrix.SetLineData(rowIndex)
                End With
            End If
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

    Private Sub LoadPaymentVoucher(ByRef oActiveForm As SAPbouiCOM.Form)
        Dim oPayForm As SAPbouiCOM.Form
        Dim oOptBtn As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        If Not LoadFromXML(p_oSBOApplication, "PaymentVoucher.srf") Then Throw New ArgumentException(sErrDesc)
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
        oPayForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("VOUCHER", oRecordSet)
        oPayForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        oPayForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()

        If HolidayMarkUp(oPayForm, oPayForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)

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

    Private Sub LoadShippingInvoice(ByRef oActiveForm As SAPbouiCOM.Form)
        Dim oShpForm As SAPbouiCOM.Form
        Dim sErrDesc As String = String.Empty
        'Dim oMatrix As SAPbouiCOM.Matrix

        If Not LoadFromXML(p_oSBOApplication, "ShipInvoice.srf") Then Throw New ArgumentException(sErrDesc)
        oShpForm = p_oSBOApplication.Forms.ActiveForm
        oShpForm.EnableMenu("1288", True)
        oShpForm.EnableMenu("1289", True)
        oShpForm.EnableMenu("1290", True)
        oShpForm.EnableMenu("1291", True)
        oShpForm.EnableMenu("1284", False)
        oShpForm.EnableMenu("1286", False)


        oShpForm.EnableMenu("1283", False) 'MSW 01-04-2011
        oShpForm.EnableMenu("1292", True)
        oShpForm.EnableMenu("1293", True)
        oShpForm.EnableMenu("5907", True)

        oShpForm.AutoManaged = True
        oShpForm.DataBrowser.BrowseBy = "ed_DocNum"
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oShpForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("SHIPPINGINV", oRecordSet)
        oShpForm.Items.Item("ed_ShDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oShpForm.Items.Item("ed_Job").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = "SELECT CardName,Address FROM OCRD WHERE CardCode = '" & oActiveForm.Items.Item("ed_Code").Specific.Value & "'"
        oRecordSet.DoQuery(sql)
        If oRecordSet.RecordCount > 0 Then
            Dim shipTo As String = String.Empty
            shipTo = Trim(oRecordSet.Fields.Item("CardName").Value) & Chr(13) & _
                    Trim(oRecordSet.Fields.Item("Address").Value)
            oShpForm.DataSources.DBDataSources.Item("@OBT_TB03_EXPSHPINV").SetValue("U_ShipTo", 0, shipTo)
            'oShpForm.Items.Item("ed_ShipTo").Specific.Value = shipTo
        End If

        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        ' oShpForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()

        If HolidayMarkUp(oShpForm, oShpForm.Items.Item("ed_ShDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        ' oShpForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)



        'oMatrix = oActiveForm.Items.Item("mx_ShpInv").Specific
        'If oShpForm.Items.Item("ed_VocNo").Specific.Value = "" Then
        '    If (oMatrix.RowCount > 0) Then
        '        If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
        '            oShpForm.Items.Item("ed_VocNo").Specific.Value = 1
        '        Else
        '            oShpForm.Items.Item("ed_VocNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
        '        End If
        '    Else
        '        oShpForm.Items.Item("ed_VocNo").Specific.Value = 1
        '    End If
        'End If

        oShpForm.Freeze(True)
        If AddUserDataSrc(oShpForm, "ExInv", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PONo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        If AddUserDataSrc(oShpForm, "Part", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PartDesp", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Qty", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Unit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Box", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "DL", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "DB", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "DH", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "M3", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TM3", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "NetKg", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TotKg", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "GroKg", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TotGKg", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "NEC", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TotNEC", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TotBox", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "BUnit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PBox", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PUnit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "UPrice", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TotV", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "SName", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PSName", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Ecc", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Lic", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "LicExDate", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Cls", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "UN", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "HSCode", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "DOM", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


        oShpForm.Items.Item("ed_ExInv").Specific.DataBind.SetBound(True, "", "ExInv")
        oShpForm.Items.Item("ed_PO").Specific.DataBind.SetBound(True, "", "PONo")

        oShpForm.Items.Item("ed_Part").Specific.DataBind.SetBound(True, "", "Part")
        oShpForm.Items.Item("ed_PartDes").Specific.DataBind.SetBound(True, "", "PartDesp")
        oShpForm.Items.Item("ed_Qty").Specific.DataBind.SetBound(True, "", "Qty")
        oShpForm.Items.Item("ed_Unit").Specific.DataBind.SetBound(True, "", "Unit")
        oShpForm.Items.Item("ed_Box").Specific.DataBind.SetBound(True, "", "Box")
        oShpForm.Items.Item("ed_L").Specific.DataBind.SetBound(True, "", "DL")
        oShpForm.Items.Item("ed_B").Specific.DataBind.SetBound(True, "", "DB")
        oShpForm.Items.Item("ed_H").Specific.DataBind.SetBound(True, "", "DH")
        oShpForm.Items.Item("ed_M3").Specific.DataBind.SetBound(True, "", "M3")
        oShpForm.Items.Item("ed_M3T").Specific.DataBind.SetBound(True, "", "TM3")
        oShpForm.Items.Item("ed_Net").Specific.DataBind.SetBound(True, "", "NetKg")
        oShpForm.Items.Item("ed_NetT").Specific.DataBind.SetBound(True, "", "TotKg")
        oShpForm.Items.Item("ed_Gross").Specific.DataBind.SetBound(True, "", "GroKg")
        oShpForm.Items.Item("ed_GrossT").Specific.DataBind.SetBound(True, "", "TotGKg")
        oShpForm.Items.Item("ed_Nec").Specific.DataBind.SetBound(True, "", "NEC")
        oShpForm.Items.Item("ed_NecT").Specific.DataBind.SetBound(True, "", "TotNEC")
        oShpForm.Items.Item("ed_TotBox").Specific.DataBind.SetBound(True, "", "TotBox")
        oShpForm.Items.Item("ed_Boxes").Specific.DataBind.SetBound(True, "", "BUnit")
        oShpForm.Items.Item("ed_PPBNo").Specific.DataBind.SetBound(True, "", "PBox")
        oShpForm.Items.Item("ed_PUnit").Specific.DataBind.SetBound(True, "", "PUnit")
        oShpForm.Items.Item("ed_UnPrice").Specific.DataBind.SetBound(True, "", "UPrice")
        oShpForm.Items.Item("ed_TotVal").Specific.DataBind.SetBound(True, "", "TotV")
        oShpForm.Items.Item("ed_ShName").Specific.DataBind.SetBound(True, "", "SName")
        oShpForm.Items.Item("ed_PShName").Specific.DataBind.SetBound(True, "", "PSName")
        oShpForm.Items.Item("ed_ECCN").Specific.DataBind.SetBound(True, "", "Ecc")
        oShpForm.Items.Item("ed_License").Specific.DataBind.SetBound(True, "", "Lic")
        oShpForm.Items.Item("ed_LExDate").Specific.DataBind.SetBound(True, "", "LicExDate")
        oShpForm.Items.Item("ed_Class").Specific.DataBind.SetBound(True, "", "Cls")
        oShpForm.Items.Item("ed_UN").Specific.DataBind.SetBound(True, "", "UN")
        oShpForm.Items.Item("ed_HSCode").Specific.DataBind.SetBound(True, "", "HSCode")
        oShpForm.Items.Item("ed_DOM").Specific.DataBind.SetBound(True, "", "DOM")

        AddChooseFromList(oShpForm, "cflPart", False, "PART")
        oShpForm.Items.Item("ed_Part").Specific.ChooseFromListUID = "cflPart"
        oShpForm.Items.Item("ed_Part").Specific.ChooseFromListAlias = "U_PartNo"
        oShpForm.Items.Item("bt_PPView").Visible = False
        oShpForm.Items.Item("ed_ShipTo").Specific.Active = True
        oShpForm.Freeze(False)
    End Sub

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
                'MoveWindow(myprocess.MainWindowHandle, p_oSBOApplication.Forms.ActiveForm.Left + p_oSBOApplication.Forms.ActiveForm.Width + 6, p_oSBOApplication.Forms.ActiveForm.Top + 68, 300, 578, True)
                MoveWindow(myprocess.MainWindowHandle, p_oSBOApplication.Forms.ActiveForm.Left + p_oSBOApplication.Forms.ActiveForm.Width + 6, p_oSBOApplication.Forms.ActiveForm.Top + 68, 300, 618, True)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function Validateforform(ByVal ItemUID As String, ByVal ExportSeaLCLForm As SAPbouiCOM.Form) As Boolean
        If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_Name").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "cb_PCode" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("cb_PCode").Specific.Value) Then
            p_oSBOApplication.SetStatusBarMessage("Must Select Port Of Loading", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Shipping Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_JobNo" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Fill Job No", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub ValidateJobNumber(ByRef oActiveForm As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim jobType As String = String.Empty
        Dim str As String = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        Dim strjobType As String = Left(oActiveForm.Items.Item("ed_JType").Specific.Value, 6)
        If strjobType = "Export" Then
            jobType = "EX"
        ElseIf strjobType = "Import" Then
            jobType = "IM"
        End If
        Dim curYear As String = String.Empty
        Dim strErr As Integer

        If str.Length <> 12 Then
            strErr = 1
        Else
            Dim checkchar As String = Right(str, 6)
            curYear = str.Substring(2, 4)
            Dim jobMode As String = Left(str, 2)
            Dim i As Integer = 0

            If jobMode.ToUpper <> jobType Or curYear <> Now.Year.ToString() Then
                strErr = 2
            ElseIf Not IsNumeric(checkchar) Then
                strErr = 3
            Else
                If strjobType = "Export" Then
                    sql = "select COUNT(*) count from [@OBT_TB002_EXPORT] WHERE U_JobNum='" & str & "'"
                    oRecordSet.DoQuery(sql)
                    i = oRecordSet.Fields.Item("count").Value
                    If i > 0 Then
                        strErr = 4
                    Else
                        'not use now job validation for export fcl table
                        'sql = "select COUNT(*) count from [@OBT_TB002_EXPORTFCL] WHERE U_JobNum='" & str & "'"
                        'oRecordSet.DoQuery(sql)
                        'i = oRecordSet.Fields.Item("count").Value
                        'If i > 0 Then
                        '    strErr = 4
                        'End If
                    End If
                ElseIf strjobType = "Import" Then
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
        End If

        Select Case strErr
            Case 1
                oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                p_oSBOApplication.SetStatusBarMessage("Invalid Job Number.Job Number must be " & jobType & curYear & "xxxxxx", SAPbouiCOM.BoMessageTime.bmt_Short)
                BubbleEvent = False
            Case 2
                oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                p_oSBOApplication.SetStatusBarMessage("Invalid Job Number.Prefix must be """ & jobType & " and Current Year: " & jobType & curYear & """", SAPbouiCOM.BoMessageTime.bmt_Short)
                BubbleEvent = False
            Case 3
                oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                p_oSBOApplication.SetStatusBarMessage("Invalid Job Number.", SAPbouiCOM.BoMessageTime.bmt_Short)
                BubbleEvent = False
            Case 4
                oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                p_oSBOApplication.SetStatusBarMessage("Job Number is already exist.", SAPbouiCOM.BoMessageTime.bmt_Short)
                BubbleEvent = False
        End Select
    End Sub

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

    Private Function AddNewRow(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        AddNewRow = False
        Dim sErrDesc As String = vbNullString
        Dim oDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
            oDbDataSource = oActiveForm.DataSources.DBDataSources.Item("@OBT_TB032_VDETAIL")
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

    Public Sub LoadHolidayMarkUp(ByVal ExportSeaLCLForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Public Sub AddUpdateShippingMatrix(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
        ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item(DataSource)
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
                        .SetValue("U_SSeqNo", .Offset, 1)
                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                        .SetValue("U_SSeqNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                    End If

                    .SetValue("U_DocNo", .Offset, pForm.Items.Item("ed_Job").Specific.Value)
                    .SetValue("U_ShipTo", .Offset, pForm.Items.Item("ed_ShipTo").Specific.Value)
                    .SetValue("U_ShipBy", .Offset, pForm.Items.Item("ed_ShipBy").Specific.Value)
                    .SetValue("U_ShipDate", .Offset, pForm.Items.Item("ed_ShDate").Specific.Value)
                    .SetValue("U_SInvNo", .Offset, pForm.Items.Item("ed_ShInvNo").Specific.Value)
                    .SetValue("U_SInvDate", .Offset, pForm.Items.Item("ed_Date").Specific.Value)
                    .SetValue("U_Remark", .Offset, pForm.Items.Item("ed_Remarks").Specific.Value)
                    .SetValue("U_ShDocNo", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value)

                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource


                    .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.GetNextSelectedRow).Specific.Value)
                    .SetValue("U_SSeqNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.GetNextSelectedRow).Specific.Value)
                    .SetValue("U_DocNo", .Offset, pForm.Items.Item("ed_Job").Specific.Value)
                    .SetValue("U_ShipTo", .Offset, pForm.Items.Item("ed_ShipTo").Specific.Value)
                    .SetValue("U_ShipBy", .Offset, pForm.Items.Item("ed_ShipBy").Specific.Value)
                    .SetValue("U_ShipDate", .Offset, pForm.Items.Item("ed_ShDate").Specific.Value)
                    .SetValue("U_SInvNo", .Offset, pForm.Items.Item("ed_ShInvNo").Specific.Value)
                    .SetValue("U_SInvDate", .Offset, pForm.Items.Item("ed_Date").Specific.Value)
                    .SetValue("U_Remark", .Offset, pForm.Items.Item("ed_Remarks").Specific.Value)
                    .SetValue("U_ShDocNo", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value)

                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

End Module
