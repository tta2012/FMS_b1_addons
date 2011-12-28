Option Explicit On

Imports SAPbobsCOM
Imports SAPbouiCOM
Imports System.Xml
Imports System.Threading
Imports System.Net.Mail
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.Diagnostics
Imports System.IO

Module modImportSeaFCL

    Public TaskDS As SAPbouiCOM.DBDataSource

    Dim docfolderpath As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oUDSs As SAPbouiCOM.UserDataSources
    Dim BoolResize As Boolean = False
    Dim sql As String = ""
    Dim rowIndex As Integer

    ' ============== Private Variable For Interface Item ===============
    Private oEdDocNum As SAPbouiCOM.EditText
    ' ============== Private Variable For Interface Item ===============
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Private LineId, ConSeqNo, ConNo, ConSealNo, ConSize, ConType, ConWt, ConDesc, ConDate, ConDay, ConTime, Conunstuff, ChStuff As String
    Private VocNo, VendorName, PayTo, PayType, BankName, CheqNo, Status, Currency, PaymentDate, GST, PrepBy, ExRate, Remark As String
    Private Total, GSTAmt, SubTotal As Double
    Private vocTotal As Double
    Private gstTotal As Double
    Private vendorCode As String
    Dim strAPInvNo, strOutPayNo As String
    Private currentRow As Integer
    Private ActiveMatrix As String
    Private DocLastKey As String

    Dim dtmatrix As SAPbouiCOM.DataTable
    Public gridindex As String

    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String
    Private Enum ImageType
        TIFF
        BMP
        JPG
        PNG
    End Enum

    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public Function DImportSeaFCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DImportSeaFCLFormDataEvent
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
        DImportSeaFCLFormDataEvent = RTN_ERROR
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oPOForm As SAPbouiCOM.Form = Nothing
        Dim sKeyValue As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSQLQuery As String = String.Empty
        Dim oDocument As SAPbobsCOM.Documents
        Dim oXmlReader As XmlTextReader
        Dim sDocNum As String = String.Empty
        Dim oEditText As SAPbouiCOM.EditText
        Dim oMatrix, oChMatrix As SAPbouiCOM.Matrix
        Dim sFuncName As String = "p_oSBOApplication_FormDataEvent"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            Select Case BusinessObjectInfo.FormTypeEx
                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oPOForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        'If BusinessObjectInfo.FormUID = "PURCHASEORDER" Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            'When action of BusinessObjectInfo is nicely done, need to do 2 tasks in action for Purcahse process
                            ' (1). need to show PopulatePurchaseHeader into the matrix of Main Export Form
                            ' (2). need to create PurchaseOrder into OPOR and POR1, related with main PurchaseProcess by using oPurchaseOrder document
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)

                            'oMatrix = oActiveForm.Items.Item("mx_Fumi").Specific
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                        "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If Not CreatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            If BusinessObjectInfo.FormTypeEx = "2000000021" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB033_FORKLIF", True) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB030_ARMES", True) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000010" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB034_FORKLIF", True) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB032_BUNKER", True) Then Throw New ArgumentException(sErrDesc)

                            End If
                            'If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql) Then Throw New ArgumentException(sErrDesc)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value))
                            SendAttachFile(oActiveForm, oPOForm)
                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                       "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If BusinessObjectInfo.FormTypeEx = "2000000021" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB033_FORKLIF", False) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB030_ARMES", False) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000010" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB034_FORKLIF", False) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB032_BUNKER", False) Then Throw New ArgumentException(sErrDesc)

                            End If
                            If Not UpdatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                        End If
                        'End If
                    End If

                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            If Not CreateGoodsReceiptPO(oActiveForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then

                        End If
                    End If

                Case "SHIPPINGINV"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            oShpForm = p_oSBOApplication.Forms.GetForm("SHIPPINGINV", 1)
                            oMatrix = oActiveForm.Items.Item("mx_ShpInv").Specific
                            If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_LCL20_SHPINV", True)
                            ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_LCL20_SHPINV", False)
                            End If
                        Catch ex As Exception

                        End Try
                    End If

                Case "VOUCHER"
                    If BusinessObjectInfo.BeforeAction = True Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                            Try
                                oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
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
                                oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
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
                        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                        oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                        oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_TB028_VOUC", True)
                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_TB028_VOUC", False)
                        End If

                    End If
                Case "142"
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            Dim oImportSeaLCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            oDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            Dim sCode As String = String.Empty
                            Dim sName As String = String.Empty
                            Dim sAttention As String = String.Empty
                            Dim sPhone As String = String.Empty
                            Dim sFax As String = String.Empty
                            Dim sMail As String = String.Empty
                            Dim sDate As Date
                            sDocNum = BusinessObjectInfo.ObjectKey
                            oXmlReader = New XmlTextReader(New IO.StringReader(sDocNum))
                            While oXmlReader.Read()
                                If oXmlReader.NodeType = XmlNodeType.XmlDeclaration Then
                                    sDocNum = oXmlReader.ReadElementString
                                End If
                            End While
                            oXmlReader.Close()
                            oRecordSet.DoQuery("SELECT CardCode,CardName FROM OPOR WHERE DocEntry = '" + sDocNum + "'")
                            If oRecordSet.RecordCount > 0 Then
                                sCode = oRecordSet.Fields.Item("CardCode").Value.ToString
                                sName = oRecordSet.Fields.Item("CardName").Value.ToString
                            End If
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

                Case "IMPORTSEAFCL"
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            LoadHolidayMarkUp(oActiveForm)
                            oActiveForm.Items.Item("ed_CunTime").Specific.Value = String.Empty 'MSW
                            oActiveForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty
                            'MSW
                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                                If oMatrix.RowCount > 1 Then
                                    oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                                ElseIf oMatrix.RowCount = 1 Then
                                    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                        oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                                    Else
                                        oActiveForm.Items.Item("bt_AmdCont").Enabled = False
                                    End If
                                End If
                                'MSW Voucher
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                                    If oMatrix.RowCount > 1 Then
                                        oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                        Else
                                            '   oActiveForm.Items.Item("bt_AmdVoc").Enabled = False
                                        End If
                                    End If
                                End If
                                oActiveForm.Items.Item("ed_VedName").Specific.Value = String.Empty
                                oActiveForm.Items.Item("ed_PayTo").Specific.Value = String.Empty
                                oActiveForm.Items.Item("ed_VocNo").Specific.Value = String.Empty
                                oActiveForm.Items.Item("ed_PosDate").Specific.Value = String.Empty
                                oActiveForm.Items.Item("ed_PJobNo").Specific.Value = String.Empty
                                oActiveForm.Items.Item("ed_VRemark").Specific.Value = String.Empty
                                'End MSW Voucher
                            End If
                        End If
                    End If
            End Select
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            DImportSeaFCLFormDataEvent = RTN_SUCCESS
        Catch ex As Exception
            DImportSeaFCLFormDataEvent = RTN_ERROR
        End Try
    End Function

    Public Function DoImportSeaFCLMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   ImportSeaFCLMenuEvent
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
        DoImportSeaFCLMenuEvent = RTN_ERROR
        Dim ImportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oOpt As SAPbouiCOM.OptionBtn
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CPOMatrix As SAPbouiCOM.Matrix = Nothing
        Dim CGRMatrix As SAPbouiCOM.Matrix = Nothing

        Dim sFuncName As String = "DoImportSeaFCLMenuEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            Select Case pVal.MenuUID
                Case "mnuImportSeaFCL"
                    If pVal.BeforeAction = False Then
                        LoadImportSeaFCLForm()
                    End If

                Case "1281"
                    If pVal.BeforeAction = False Then
                        'If SBO_Application.Forms.Item("IMPORTSEAFCL").Selected = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                            p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("ed_JobNo").Enabled = True
                            p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("ed_JobNo").Specific.Active = True
                            If ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ImportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                                'ImportSeaFCLForm.Items.Item("ed_Wrhse").Enabled = True
                            End If
                            If AddChooseFromListByOption(p_oSBOApplication.Forms.Item("IMPORTSEAFCL"), True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                Case "1282"
                    If pVal.BeforeAction = False Then
                        'If p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Selected = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                            EnabledHeaderControls(p_oSBOApplication, False) 'MSW 26-03-2011
                            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordset.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEAFCL'")
                            If oRecordset.RecordCount > 0 Then
                                '  ImportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                                ImportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordset.Fields.Item("NextNumber").Value.ToString 'MSW 01-06-2011  for Job Type Table
                            End If
                            ImportSeaFCLForm.Items.Item("ed_JType").Specific.Value = "Import Sea FCL" 'MSW 30-05-2011 for Job Type Table
                            ImportSeaFCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            ImportSeaFCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                            ImportSeaFCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                            If HolidayMarkUp(ImportSeaFCLForm, ImportSeaFCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, _
                                             ImportSeaFCLForm.Items.Item("ed_JbDay").Specific, ImportSeaFCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ImportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                        End If
                    End If

                Case "1288", "1289", "1290", "1291"
                    If pVal.BeforeAction = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                            If ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If
                        End If
                    End If

                Case "EditVoc"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            LoadPaymentVoucher(ImportSeaFCLForm)
                            ' If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ImportSeaFCLForm.Items.Item("mx_Voucher").Specific
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oPayForm.Items.Item("ed_DocNum").Visible = True
                            oPayForm.Items.Item("ed_DocNum").Enabled = True
                            oPayForm.Items.Item("ed_DocNum").Specific.Value = oMatrix.Columns.Item("colVDocNum").Cells.Item(currentRow).Specific.Value.ToString
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

                Case "EditShp"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            LoadShippingInvoice(ImportSeaFCLForm)
                            ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ImportSeaFCLForm.Items.Item("mx_ShpInv").Specific
                            oShpForm = p_oSBOApplication.Forms.ActiveForm
                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oShpForm.Items.Item("ed_DocNum").Visible = True
                            oShpForm.Items.Item("ed_DocNum").Enabled = True
                            oShpForm.Items.Item("ed_DocNum").Specific.Value = oMatrix.Columns.Item("colSDocNum").Cells.Item(currentRow).Specific.Value.ToString
                            oShpForm.Items.Item("ed_ShipBy").Specific.Active = True
                            oShpForm.Items.Item("ed_DocNum").Visible = False
                            oShpForm.Items.Item("ed_DocNum").Enabled = False
                            oShpForm.DataBrowser.BrowseBy = "ed_DocNum"
                            oShpForm.Items.Item("ed_ShipTo").Enabled = False
                            'If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                            oShpForm.Items.Item("1").Click()
                        End If
                    End If

                Case "EditCPO"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            LoadAndCreateCPO(ImportSeaFCLForm, RPOsrfname)
                            ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oMatrix = ImportSeaFCLForm.Items.Item(RPOmatrixname).Specific
                            CPOForm = p_oSBOApplication.Forms.ActiveForm
                            CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                            CPOForm.Items.Item("1").Click()
                        Else
                            p_oSBOApplication.MessageBox("Need to select the Row that you want to Edit")
                        End If
                    End If

                Case "CopyToCGR"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            ImportSeaFCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            LoadAndCreateCGR(ImportSeaFCLForm, RGRsrfname)
                            ImportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                            If Not FillDataToGoodsReceipt(ImportSeaFCLForm, RGRmatrixname, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                        Else
                            p_oSBOApplication.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                        End If
                    End If

            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            DoImportSeaFCLMenuEvent = RTN_SUCCESS
        Catch ex As Exception
            DoImportSeaFCLMenuEvent = RTN_ERROR
        End Try
    End Function

    Public Function DoImportSeaFCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoImportSeaFCLItemEvent
        '   Purpose     :   This function will be providing to proceed validating for
        '                   Inventory [All] Menu Event information
        '               f
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
        DoImportSeaFCLItemEvent = RTN_ERROR
        Dim oActiveForm, oPayForm, oShpForm As SAPbouiCOM.Form
        Dim oEditText As SAPbouiCOM.EditText
        Dim oMatrix, oChMatrix, oShpMatrix As SAPbouiCOM.Matrix
        Dim oDate As Date
        Dim oItem As SAPbouiCOM.Item
        Dim oLinkedButton As SAPbouiCOM.LinkedButton
        Dim oDBDataSource As SAPbouiCOM.DBDataSource
        Dim oOptBtn As SAPbouiCOM.OptionBtn
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CPOMatrix As SAPbouiCOM.Matrix = Nothing
        Dim CGRMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim sFuncName As String = "DoImportSeaFCLItemEvent()"
        Try
            Select Case pVal.FormTypeEx
                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016"       ' CGR --> Custom Goods Receipt
                    CGRForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Try
                                If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                                    CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)
                                    CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)
                                    oCombo = CGRForm.Items.Item("cb_Contact").Specific
                                    If Not ClearComboData(CGRForm, "cb_Contact", "@OBT_TB12_FFCGR", "U_CPerson") Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CGRForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                            oRecordSet.MoveNext()
                                        End While
                                    End If
                                End If
                                If pVal.ColUID = "colItemNo" Then
                                    CGRMatrix = CGRForm.Items.Item("mx_Item").Specific
                                    Try
                                        CGRMatrix.Columns.Item("colItemNo").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        CGRMatrix.Columns.Item("colIDesc").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(1, 0).ToString
                                    Catch ex As Exception
                                    End Try

                                    If CGRMatrix.RowCount > 1 Then
                                        If CGRMatrix.Columns.Item("colItemNo").Cells.Item(CGRMatrix.RowCount).Specific.Value <> "" Then
                                            If Not AddNewRowGR(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                        End If
                                    Else
                                        If CGRMatrix.Columns.Item("colItemNo").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                            If Not AddNewRowGR(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                        End If
                                    End If


                                End If
                            Catch ex As Exception
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            If pVal.ItemUID = "cb_Contact" Then
                                Dim oComboContact As SAPbouiCOM.ComboBox = CGRForm.Items.Item("cb_Contact").Specific
                                CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_CPerson", 0, oComboContact.Selected.Description.ToString)
                                CGRForm.Items.Item("ed_VRef").Specific.Active = True
                            End If

                            If pVal.ItemUID = "cb_SInA" Then
                                Dim comboStaff As SAPbouiCOM.ComboBox = CGRForm.Items.Item("cb_SInA").Specific
                                CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_SInA", 0, comboStaff.Selected.Description.ToString)
                                CGRForm.Items.Item("ed_TDate").Specific.Active = True
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "ed_TDate" And CGRForm.Items.Item("ed_TDate").Specific.String <> String.Empty Then
                                If Not DateTime(CGRForm, _
                                                CGRForm.Items.Item("ed_TDate").Specific, _
                                                CGRForm.Items.Item("ed_TDay").Specific, _
                                                CGRForm.Items.Item("ed_TTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                CGRForm.Items.Item("ed_TPlace").Specific.Active = True
                            End If

                            If pVal.ItemUID = "1" Then
                                If pVal.BeforeAction = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        'CGRForm = p_oSBOApplication.Forms.ActiveForm
                                        'If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        '    If Not CreateGoodsReceiptPO(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                        'End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015"      ' CPO --> Custom Purchase Order"
                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                    If pVal.BeforeAction = False Then
                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            CPOForm.Items.Item("ed_Code").Enabled = False
                            CPOForm.Items.Item("ed_Name").Enabled = False
                            CPOForm.Items.Item("cb_SInA").Enabled = False
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            'Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Try
                                If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                                    CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)
                                    CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_VName", 0, oDataTable.GetValue(1, 0).ToString)
                                    oCombo = CPOForm.Items.Item("cb_Contact").Specific
                                    If Not ClearComboData(CPOForm, "cb_Contact", "@OBT_TB08_FFCPO", "U_CPerson") Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                            oRecordSet.MoveNext()
                                        End While
                                    End If

                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT E_Mail FROM OCRD WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        Dim stremail As String = IIf(oRecordSet.Fields.Item("E_Mail").Value.ToString = vbNull.ToString, "", oRecordSet.Fields.Item("E_Mail").Value)
                                        Try
                                            CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_Email", 0, stremail)
                                            'CPOForm.Items.Item("ed_Email").Specific.Value = stremail
                                        Catch ex As Exception

                                        End Try

                                    End If
                                End If

                                If pVal.ColUID = "colItemNo" Then
                                    CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                                    Try
                                        CPOMatrix.Columns.Item("colItemNo").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        CPOMatrix.Columns.Item("colIDesc").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(1, 0).ToString
                                    Catch ex As Exception
                                    End Try


                                    If CPOMatrix.RowCount > 1 Then
                                        If CPOMatrix.Columns.Item("colItemNo").Cells.Item(CPOMatrix.RowCount).Specific.Value <> "" Then
                                            If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                        End If
                                    Else
                                        If CPOMatrix.Columns.Item("colItemNo").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                            If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                        End If
                                    End If
                                End If


                            Catch ex As Exception
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "2" Then
                                CPOForm = Nothing
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            If pVal.ColUID = "colIQty" Or pVal.ColUID = "colIPrice" Then
                                CalAmtPO(CPOForm, pVal.Row)
                                CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                                If CPOMatrix.Columns.Item("colIGST").Cells.Item(CPOMatrix.RowCount).Specific.Value <> "" Then
                                    CalRatePO(CPOForm, pVal.Row)
                                End If
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            If pVal.ColUID = "colIGST" Then
                                CalRatePO(CPOForm, pVal.Row)
                            End If

                            If pVal.ItemUID = "cb_Contact" Then
                                Dim oComboContact As SAPbouiCOM.ComboBox = CPOForm.Items.Item("cb_Contact").Specific
                                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_CPerson", 0, oComboContact.Selected.Description.ToString)
                                CPOForm.Items.Item("ed_VRef").Specific.Active = True
                            End If

                            If pVal.ItemUID = "cb_SInA" Then
                                Dim comboStaff As SAPbouiCOM.ComboBox = CPOForm.Items.Item("cb_SInA").Specific
                                Dim strempId As String = comboStaff.Selected.Value.ToString
                                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_SInA", 0, comboStaff.Selected.Description.ToString)
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                oRecordSet.DoQuery("SELECT email FROM OHEM WHERE empID = " & strempId & "")
                                If oRecordSet.RecordCount > 0 Then
                                    oRecordSet.MoveFirst()
                                    If CPOForm.Items.Item("ed_Email").Specific.Value <> "" Then
                                        CPOForm.Items.Item("ed_Email").Specific.Value = CPOForm.Items.Item("ed_Email").Specific.Value + "," + IIf(oRecordSet.Fields.Item("email").Value = vbNull.ToString, "", oRecordSet.Fields.Item("email").Value)
                                    Else
                                        CPOForm.Items.Item("ed_Email").Specific.Value = IIf(oRecordSet.Fields.Item("email").Value = vbNull.ToString, "", oRecordSet.Fields.Item("email").Value)
                                    End If

                                End If

                                CPOForm.Items.Item("ed_TDate").Specific.Active = True
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "ed_TDate" And CPOForm.Items.Item("ed_TDate").Specific.String <> String.Empty Then
                                If Not DateTime(CPOForm, _
                                                CPOForm.Items.Item("ed_TDate").Specific, _
                                                CPOForm.Items.Item("ed_TDay").Specific, _
                                                CPOForm.Items.Item("ed_TTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                CPOForm.Items.Item("ed_TPlace").Specific.Active = True
                            End If

                            If pVal.ItemUID = "1" Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                ' When click the Add button in AddMode of Custom Purchase Order form
                                ' need to also trigger the item pressed event of Main Export Form according by Customize Biz Logic
                                If pVal.Action_Success = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        CPOForm = p_oSBOApplication.Forms.ActiveForm

                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            'when retrieve spefific data for update, add new row in the matrix
                                            If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                            CPOForm.Items.Item("ed_Code").Enabled = False
                                            CPOForm.Items.Item("ed_Name").Enabled = False
                                            CPOForm.Items.Item("cb_SInA").Enabled = False
                                            CPOForm.Items.Item("bt_Preview").Visible = True
                                            CPOForm.Items.Item("bt_Resend").Visible = True
                                        End If

                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            CPOForm.Close()
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oActiveForm.Items.Item("1").Click()
                                        End If


                                    End If
                                End If
                            End If

                            If pVal.ItemUID = "bt_Preview" Then
                                PreviewPO(oActiveForm, CPOForm)
                            End If
                            If pVal.ItemUID = "bt_Resend" Then
                                SendAttachFile(oActiveForm, CPOForm)
                            End If

                            If pVal.ItemUID = "bt_Test" Then
                                CPOForm = p_oSBOApplication.Forms.ActiveForm
                                CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                                CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                                If CPOMatrix.Columns.Item("colItemNo").Cells.Item(CPOMatrix.RowCount).Specific.Value = "" Then
                                    CPOMatrix.DeleteRow(CPOMatrix.RowCount)
                                End If
                                If Not UpdatePurchaseOrder(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                    End If

                    If pVal.BeforeAction = True Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "1" Then
                                Try
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                                    If Not CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                                        If CPOMatrix.Columns.Item("colItemNo").Cells.Item(CPOMatrix.RowCount).Specific.Value = "" Then
                                            CPOMatrix.DeleteRow(CPOMatrix.RowCount)
                                        End If
                                    End If
                                Catch ex As Exception
                                    BubbleEvent = False
                                    MessageBox.Show(ex.Message)
                                End Try
                            End If
                        End If
                    End If
                Case "SHIPPINGINV"
                    oShpForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    ' oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                    Dim m3 As Decimal
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "1" Then
                                If pVal.ActionSuccess = True Then
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                    oShpForm.Items.Item("bt_PPView").Visible = True
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Then
                                        'If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        '    If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                                        '    If oPayForm.Items.Item("ed_PayType").Specific.Value = "Cash" Then
                                        '        oPayForm.Items.Item("op_Cash").Specific.Selected = True
                                        '        oPayForm.Items.Item("cb_BnkName").Enabled = False
                                        '        oPayForm.Items.Item("ed_Cheque").Enabled = False
                                        '    ElseIf oPayForm.Items.Item("ed_PayType").Specific.Value = "Cheque" Then
                                        '        oPayForm.Items.Item("op_Cheq").Specific.Selected = True
                                        '        oPayForm.Items.Item("cb_BnkName").Enabled = True
                                        '        oPayForm.Items.Item("ed_Cheque").Enabled = True
                                        '    End If
                                        '    oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                        '    DisableChargeMatrix(oPayForm, oMatrix, False)
                                        '    oPayForm.Items.Item("ed_VedName").Enabled = False
                                        '    oPayForm.Items.Item("ed_PayTo").Enabled = False

                                        'End If
                                        If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            sql = "Update [@OBT_TB03_EXPSHPINV] set " & _
                                            " U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oShpForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            oShpForm.Close()

                                            oActiveForm.Items.Item("1").Click()
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If

                                        ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Items.Item("ed_ShipTo").Enabled = False
                                            oShpForm.Items.Item("bt_PPView").Visible = True
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                End If
                            End If
                            If pVal.ItemUID = "bt_PPView" Then
                                'Save Report To specific JobFile as PDF File USE Code
                                Dim reportuti As New clsReportUtilities
                                Dim pdfFilename As String = "SHIPPING INV"
                                Dim mainFolder As String = "C:\Users\PC-8\Documents\Fuji Xerox\DocuWorks\DWFolders\User Folder\"
                                Dim jobNo As String = oShpForm.Items.Item("ed_Job").Specific.Value
                                Dim rptPath As String = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString & "\Shipping Invoice.rpt"
                                Dim pdffilepath As String = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                Dim rptDocument As ReportDocument = New ReportDocument()
                                rptDocument.Load(rptPath)
                                rptDocument.Refresh()
                                rptDocument.SetParameterValue("@DocEntry", oShpForm.Items.Item("ed_DocNum").Specific.Value)
                                reportuti.SetDBLogIn(rptDocument)
                                If Not pdffilepath = String.Empty Then
                                    reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument)
                                End If
                                'reportuti.PrintDoc(rptDocument) 'To Print To Printer
                                rptDocument.Close()
                                If Not reportuti.SendMailDoc("maysandar@sap-infotech.com,kyawzin@sap-infotech.com", "tuntunaung@sap-infotech.com", "ShippingInvoice", "Testing Mail", pdffilepath) Then
                                    p_oSBOApplication.MessageBox("Send Message Fail", 0, "OK")
                                Else
                                    p_oSBOApplication.MessageBox("Send Message Successfully", 0, "OK")
                                End If
                                'End Save Report To specific JobFile as PDF File
                            End If

                            If pVal.ItemUID = "bt_Add" Then
                                oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                                If ValidateforformShippingInv(oShpForm) = True Then
                                    Exit Function
                                Else
                                    If oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD" Then
                                        AddUpdateShippingInv(oShpForm, oShpMatrix, "@OBT_TB03_EXPSHPINVD", True)
                                    Else
                                        AddUpdateShippingInv(oShpForm, oShpMatrix, "@OBT_TB03_EXPSHPINVD", False)
                                        oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD"
                                    End If
                                End If
                            End If

                            If pVal.ItemUID = "bt_Clear" Then
                                ClearText(oShpForm, "ed_ExInv", "ed_PO", "ed_Part", "ed_PartDes", "ed_Qty", "ed_Unit", "ed_Box", "ed_L", "ed_B", "ed_H", "ed_M3", "ed_M3T", "ed_Net", "ed_NetT", "ed_Gross", "ed_GrossT", "ed_Nec", "ed_NecT", "ed_TotBox", "ed_Boxes", "ed_PPBNo", "ed_PUnit", "ed_UnPrice", "ed_TotVal", "ed_ShName", "ed_PShName", "ed_ECCN", "ed_License", "ed_LExDate", "ed_Class", "ed_UN", "ed_HSCode", "ed_DOM")
                            End If

                            If pVal.ItemUID = "mx_ShipInv" And pVal.ColUID = "V_-1" Then
                                oMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                                If oMatrix.GetNextSelectedRow > 0 Then
                                    If (oMatrix.IsRowSelected(oMatrix.GetNextSelectedRow)) = True Then
                                        oShpForm.Items.Item("bt_Add").Specific.Caption = "Edit"
                                        oShpForm.Freeze(True)
                                        SetShipInvDataToEditTabByIndex(oShpForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        oShpForm.Freeze(False)
                                    End If
                                Else
                                    p_oSBOApplication.MessageBox("Please Select One Row To Edit Shipping Invoice.", 1, "&OK")
                                End If
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Try
                                If pVal.ItemUID = "ed_Part" Then
                                    oShpForm.DataSources.UserDataSources.Item("Part").Value = IIf(oDataTable.Columns.Item("U_PartNo").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_PartNo").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("PartDesp").Value = IIf(oDataTable.Columns.Item("U_Desc").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_Desc").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("Qty").Value = IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("Unit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString)
                                    '   oActiveForm.DataSources.UserDataSources.Item("Box").Value = oDataTable.Columns.Item("U_Box").Cells.Item(0).Value.ToString
                                    oShpForm.DataSources.UserDataSources.Item("DL").Value = IIf(oDataTable.Columns.Item("U_Length").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Length").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("DB").Value = IIf(oDataTable.Columns.Item("U_Base").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Base").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("DH").Value = IIf(oDataTable.Columns.Item("U_Height").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Height").Cells.Item(0).Value.ToString)

                                    m3 = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Length").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Length").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Base").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Base").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Height").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Height").Cells.Item(0).Value)) / 1000000
                                    oShpForm.DataSources.UserDataSources.Item("M3").Value = m3
                                    oShpForm.DataSources.UserDataSources.Item("TM3").Value = m3 * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString))
                                    oShpForm.DataSources.UserDataSources.Item("NetKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("TotKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("GroKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value.ToString))
                                    oShpForm.DataSources.UserDataSources.Item("TotGKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("NEC").Value = 0
                                    oShpForm.DataSources.UserDataSources.Item("TotNEC").Value = 0

                                    oShpForm.DataSources.UserDataSources.Item("UPrice").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value = vbNull, 0.0, oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("TotV").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("SName").Value = IIf(oDataTable.Columns.Item("U_Shipping").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_Shipping").Cells.Item(0).Value.ToString)
                                    'oActiveForm.DataSources.UserDataSources.Item("PSName").Value = oDataTable.Columns.Item("U_PPSName").Cells.Item(0).Value.ToString
                                    oShpForm.DataSources.UserDataSources.Item("Ecc").Value = IIf(oDataTable.Columns.Item("U_ECCN").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_ECCN").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("Lic").Value = IIf(oDataTable.Columns.Item("U_ExpLic").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_ExpLic").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("LicExDate").Value = IIf(oDataTable.Columns.Item("U_LicDate").Cells.Item(0).Value = vbNull, " ", oDataTable.Columns.Item("U_LicDate").Cells.Item(0).Value)
                                    oShpForm.DataSources.UserDataSources.Item("Cls").Value = IIf(oDataTable.Columns.Item("U_Class").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_Class").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("UN").Value = IIf(oDataTable.Columns.Item("U_UNNo").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UNNo").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("HSCode").Value = IIf(oDataTable.Columns.Item("U_HSCode").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_HSCode").Cells.Item(0).Value.ToString)

                                End If
                            Catch ex As Exception

                            End Try
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                            If pVal.ItemUID = "ed_Qty" Then
                                If oShpForm.Items.Item("ed_Qty").Specific.value.ToString <> "" Then
                                    Dim qty As Double
                                    qty = IIf(Convert.ToDouble(oShpForm.Items.Item("ed_Qty").Specific.value.ToString) = 0, 0, Convert.ToDouble(oShpForm.Items.Item("ed_Qty").Specific.value.ToString))
                                    m3 = oShpForm.Items.Item("ed_L").Specific.value * oShpForm.Items.Item("ed_B").Specific.value * oShpForm.Items.Item("ed_H").Specific.value / 1000000
                                    oShpForm.DataSources.UserDataSources.Item("M3").Value = m3
                                    oShpForm.DataSources.UserDataSources.Item("TM3").Value = m3 * qty
                                    oShpForm.DataSources.UserDataSources.Item("NetKg").Value = oShpForm.Items.Item("ed_Net").Specific.value
                                    oShpForm.DataSources.UserDataSources.Item("TotKg").Value = oShpForm.Items.Item("ed_Net").Specific.value * qty
                                    oShpForm.DataSources.UserDataSources.Item("GroKg").Value = oShpForm.Items.Item("ed_Gross").Specific.value
                                    oShpForm.DataSources.UserDataSources.Item("TotGKg").Value = oShpForm.Items.Item("ed_Gross").Specific.value * qty
                                    ' oActiveForm.DataSources.UserDataSources.Item("UPrice").Value = oActiveForm.Items.Item("U_NetWt").Specific.value
                                    oShpForm.DataSources.UserDataSources.Item("TotV").Value = oShpForm.Items.Item("ed_UnPrice").Specific.value * qty
                                End If
                            End If
                        End If
                    End If

                Case "VOUCHER"
                    oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)

                    If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                        Try
                            oPayForm.Items.Item("ed_VedName").Enabled = False
                            oPayForm.Items.Item("ed_PayTo").Enabled = False
                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount1" Then
                            CalRate(oPayForm, pVal.Row)
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                        If pVal.ItemUID = "1" Then
                            Try
                                oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)

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

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
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
                            oPayForm.Items.Item("cb_BnkName").Specific.Active = True
                            oPayForm.Items.Item("ed_Cheque").Enabled = True
                        End If

                        If pVal.ItemUID = "1" Then
                            If pVal.ActionSuccess = True Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
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
                                        sql = "Update [@OBT_TB031_VHEADER] set U_APInvNo=" & Convert.ToInt32(strAPInvNo) & ",U_OutPayNo=" & Convert.ToInt32(strOutPayNo) & "" & _
                                        " ,U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oPayForm.Items.Item("ed_DocNum").Specific.Value & ""
                                        oRecordSet.DoQuery(sql)
                                        oPayForm.Close()

                                        oActiveForm.Items.Item("1").Click()
                                        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If

                                    ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        'SBO_Application.ActivateMenuItem("1291")
                                        oPayForm.Items.Item("ed_VedName").Enabled = False
                                        oPayForm.Items.Item("ed_PayTo").Enabled = False
                                        ' oPayForm.Close()
                                    End If
                                End If

                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.Before_Action = False Then
                        Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                        If pVal.ItemUID = "ed_VedName" Then
                            ObjDBDataSource = oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER") 'MSW To Add 18-3-2011
                            oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                            oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                                                                                   & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString()) 'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*

                            vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft
                            oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_BPCode", ObjDBDataSource.Offset, oDataTable.GetValue(0, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
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

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.Before_Action = False Then
                        '-------------------------For Payment(omm)------------------------------------------'
                        If (pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colGST1") Then
                            CalRate(oPayForm, pVal.Row)
                        End If
                        'If (pVal.ItemUID = "cb_GST" And oPayForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft") Then
                        If (pVal.ItemUID = "cb_GST") And oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                            ' dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
                            'oChMatrix.Columns.Item("colGST").Cells.Item(1).Specific.Value = "None"

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
                                oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_ExRate", 0, Rate.ToString)
                            Else
                                oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_ExRate", 0, Nothing)
                                oPayForm.Items.Item("ed_PayRate").Enabled = False
                                oPayForm.Items.Item("ed_PosDate").Specific.Active = True
                            End If
                        End If
                    End If

                Case "IMPORTSEAFCL"
                    oActiveForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    Dim index As Integer = 6
                    Dim folder As String

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(oActiveForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If

                    'MSW Vocuher Tab
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.Before_Action = False Then
                        '-------------------------For Payment(omm)------------------------------------------'
                        If (pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colGST") Then
                            CalRate(oActiveForm, pVal.Row)
                        End If
                        If (pVal.ItemUID = "cb_GST" And oActiveForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft") Then
                            oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                            dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
                            dtmatrix.SetValue("GST", 0, "None")
                            oMatrix.LoadFromDataSource()
                        End If
                        '----------------------------------------------------------------------------------'

                        If pVal.ItemUID = "cb_PayCur" Then
                            If oActiveForm.Items.Item("cb_PayCur").Specific.Value <> "SGD" Then
                                Dim Rate As String = String.Empty
                                sql = "SELECT Rate FROM ORTT WHERE Currency = '" & oActiveForm.Items.Item("cb_PayCur").Specific.Value & "' And DATENAME(YYYY,RateDate) = '" & _
                                        Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                        CInt(Today.ToString("dd"))
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    Rate = oRecordSet.Fields.Item("Rate").Value
                                End If
                                oActiveForm.Items.Item("ed_PayRate").Enabled = True
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_ExRate", 0, Rate.ToString)
                            Else
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_ExRate", 0, Nothing)
                                oActiveForm.Items.Item("ed_PayRate").Enabled = False
                                oActiveForm.Items.Item("ed_PosDate").Specific.Active = True
                            End If

                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then
                        oActiveForm.Freeze(True)
                        If pVal.ItemUID = "bt_Reset" Then
                            If oActiveForm.Items.Item("ch_Permit").Specific.Checked = True Then
                                index = index + 100
                                oActiveForm.Items.Item("fo_Prmt").Visible = True
                                'oActiveForm.Items.Item("fo_Prmt").Enabled = True
                                folder = "fo_Prmt"
                            Else
                                oActiveForm.Items.Item("fo_Prmt").Visible = False
                                'oActiveForm.Items.Item("fo_Prmt").Enabled = False
                            End If
                            If oActiveForm.Items.Item("ch_Voucher").Specific.Checked = True Then
                                oActiveForm.Items.Item("fo_Vchr").Visible = True
                                If folder Is Nothing Then
                                    folder = "fo_Vchr"
                                End If
                                oActiveForm.Items.Item("fo_Vchr").Left = index
                                index = index + 100
                            Else
                                oActiveForm.Items.Item("fo_Vchr").Visible = False
                            End If
                            If oActiveForm.Items.Item("ch_RYard").Specific.Checked = True Then
                                oActiveForm.Items.Item("fo_Yard").Visible = True
                                If folder Is Nothing Then
                                    folder = "fo_Yard"
                                End If
                                oActiveForm.Items.Item("fo_Yard").Left = index
                                index = index + 100
                            Else
                                oActiveForm.Items.Item("fo_Yard").Visible = False
                            End If
                            If oActiveForm.Items.Item("ch_Conta").Specific.Checked = True Then
                                oActiveForm.Items.Item("fo_Cont").Visible = True
                                If folder Is Nothing Then
                                    folder = "fo_Cont"
                                End If
                                oActiveForm.Items.Item("fo_Cont").Left = index
                                index = index + 100
                            Else
                                oActiveForm.Items.Item("fo_Cont").Visible = False
                            End If
                            If oActiveForm.Items.Item("ch_Truck").Specific.Checked = True Then
                                oActiveForm.Items.Item("fo_Trkng").Visible = True
                                If folder Is Nothing Then
                                    folder = "fo_Trkng"
                                End If
                                oActiveForm.Items.Item("fo_Trkng").Left = index
                                index = index + 100
                            Else
                                oActiveForm.Items.Item("fo_Trkng").Visible = False
                            End If
                            If oActiveForm.Items.Item("ch_Disp").Specific.Checked = True Then
                                oActiveForm.Items.Item("fo_Dsptch").Visible = True
                                If folder Is Nothing Then
                                    folder = "fo_Dsptch"
                                End If
                                oActiveForm.Items.Item("fo_Dsptch").Left = index
                                index = index + 100
                            Else
                                oActiveForm.Items.Item("fo_Dsptch").Visible = False
                            End If
                            If oActiveForm.Items.Item("ch_OCharge").Specific.Checked = vbTrue Then
                                oActiveForm.Items.Item("fo_Charge").Visible = True
                                If folder Is Nothing Then
                                    folder = "fo_Charge"
                                End If
                                oActiveForm.Items.Item("fo_Charge").Left = index
                                index = index + 100
                            Else
                                oActiveForm.Items.Item("fo_Charge").Visible = False
                            End If
                            If folder Is Nothing Then
                                oActiveForm.PaneLevel = 0
                                oActiveForm.Items.Item("rt_Outer").Visible = False
                            Else
                                oActiveForm.Items.Item("rt_Outer").Visible = True
                                oActiveForm.Items.Item(folder).Specific.Select()
                                If folder = "fo_Prmt" Then
                                    oActiveForm.PaneLevel = 8
                                End If
                            End If
                        End If
                        oActiveForm.Freeze(False)
                    End If

                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        EnabledTrucker(oActiveForm, False)
                    End If


                    If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                        Try
                            oActiveForm.Items.Item("ed_Code").Enabled = False
                            oActiveForm.Items.Item("ed_Name").Enabled = False
                            oActiveForm.Items.Item("ed_JobNo").Enabled = False 'MSW 30-05-2011
                        Catch ex As Exception
                        End Try
                    End If
                    '------------For Checklist-------------------------------------------------------------------'
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then

                        If pVal.ItemUID = "bt_ChkList" Then
                            'Dim oRecordSet As SAPbobsCOM.Recordset
                            'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            'oRecordSet.Command.Execute()
                            'Dim sql As String = String.Empty
                            'sql = "declare @xml nvarchar(max)set @xml = (SELECT COLUMN_NAME,DATA_TYPE,IS_NULLABLE,CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '@TESTSYMA3' for xml path,elements,root('root'))select @xml "

                            'oRecordSet.DoQuery(sql)
                            'oRecordSet.SaveXML(sql)
                            'Dim result As String
                            'result = oRecordSet.Fields.Item(0).Value
                            'result.Append(oRecordSet.Fields.Item(0).Value)
                            'Dim xmlDoc As New XmlDocument
                            'xmlDoc.LoadXml(result.ToString)
                            'xmlDoc.Save("d:\text.xml")

                            Start(oActiveForm) ' Docuwork Start 
                        End If
                    End If
                    '--------------------------------------------------------------------------------------'
                    '-------------------------For Payment(omm)------------------------------------------'
                    If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                        CalRate(oActiveForm, pVal.Row)
                    End If
                    '----------------------------------------------------------------------------------'


                    'MSW for Job Type Table
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ItemUID = "ed_JobNo" Then
                                ValidateJobNumber(oActiveForm, BubbleEvent)
                            End If
                        End If
                    End If

                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN) And pVal.BeforeAction = False Then

                        If pVal.ItemUID = "ed_InvDate" And oActiveForm.Items.Item("ed_InvDate").Specific.String <> String.Empty Then
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_InvDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                        'If pVal.ItemUID = "ed_ConLast" And oActiveForm.Items.Item("ed_ConLast").Specific.String <> String.Empty Then
                        '    If DateTime(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        '    If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        'End If
                        If pVal.ItemUID = "ed_InsDate" And oActiveForm.Items.Item("ed_InsDate").Specific.String <> String.Empty Then
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                        If pVal.ItemUID = "ed_TkrDate" And oActiveForm.Items.Item("ed_TkrDate").Specific.String <> String.Empty Then
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_TkrDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                        'MSW
                        If pVal.ItemUID = "ed_PosDate" And oActiveForm.Items.Item("ed_PosDate").Specific.String <> String.Empty Then
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If



                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_RESIZE And pVal.BeforeAction = False Then

                        If BoolResize = False Then 'Maximize
                            Try
                                Dim oItemRet1 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("rt_Outer")
                                Dim oItemRetInner As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("rt_Inner")
                                Dim ofoP As String = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_PMain").Height
                                oItemRet1.Top = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_Prmt").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_Prmt").Height - 1)
                                oItemRet1.Width = oActiveForm.Items.Item("mx_Cont").Width + 13 'oActiveForm.Items.Item("mx_TkrList").Width + 20
                                oItemRet1.Height = oActiveForm.Items.Item("mx_Cont").Height + 230 '33

                                oItemRetInner.Top = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_PMain").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_PMain").Height - 1)
                                oItemRetInner.Width = oActiveForm.Items.Item("mx_Cont").Width + 9
                                oItemRetInner.Height = oActiveForm.Items.Item("mx_Cont").Height + 153
                                'oItemRet1.Top = oActiveForm.Items.Item("mx_TkrList").Top - 29
                                ''oItemRet1.Top = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_Prmt").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_Prmt").Height - 1)
                                ''oItemRet1.Width = oActiveForm.Items.Item("mx_Cont").Width + 13 'oActiveForm.Items.Item("mx_TkrList").Width + 20
                                ''oItemRet1.Height = oActiveForm.Items.Item("mx_Voucher").Height + 90 '33
                                oActiveForm.Items.Item("155").Top = oActiveForm.Items.Item("fo_TkrView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_TkrView").Height - 1)
                                oActiveForm.Items.Item("155").Width = oActiveForm.Items.Item("mx_TkrList").Width + 9
                                oActiveForm.Items.Item("155").Height = oActiveForm.Items.Item("mx_TkrList").Height + 124
                                oActiveForm.Items.Item("bt_AmdIns").Top = oActiveForm.Items.Item("mx_TkrList").Top + oActiveForm.Items.Item("mx_TkrList").Height + 9
                                oActiveForm.Items.Item("bt_PrntIns").Top = oActiveForm.Items.Item("mx_TkrList").Top + oActiveForm.Items.Item("mx_TkrList").Height + 9
                                oActiveForm.Items.Item("bt_PrntIns").Left = (oActiveForm.Items.Item("155").Left + oActiveForm.Items.Item("155").Width) - (oActiveForm.Items.Item("bt_PrntIns").Width + 2)
                                'End MSW
                                'Container Tab
                                oActiveForm.Items.Item("rt_Cont").Top = oActiveForm.Items.Item("fo_ConView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_ConView").Height - 1)
                                oActiveForm.Items.Item("rt_Cont").Width = oActiveForm.Items.Item("mx_ConTab").Width + 9
                                oActiveForm.Items.Item("rt_Cont").Height = oActiveForm.Items.Item("mx_ConTab").Height + 124
                                oActiveForm.Items.Item("bt_AmdCont").Top = oActiveForm.Items.Item("mx_ConTab").Top + oActiveForm.Items.Item("mx_ConTab").Height + 9

                                'Dispatch Tab
                                oActiveForm.Items.Item("rt_Disp").Top = oActiveForm.Items.Item("fo_DisView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_DisView").Height - 1)
                                oActiveForm.Items.Item("rt_Disp").Width = oActiveForm.Items.Item("mx_DispTab").Width + 9
                                oActiveForm.Items.Item("rt_Disp").Height = oActiveForm.Items.Item("mx_DispTab").Height + 124
                                oActiveForm.Items.Item("bt_AmdDisp").Top = oActiveForm.Items.Item("mx_DispTab").Top + oActiveForm.Items.Item("mx_DispTab").Height + 9

                                'Other Charges Tab
                                oActiveForm.Items.Item("rt_Charge").Top = oActiveForm.Items.Item("fo_ChView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_ChView").Height - 1)
                                oActiveForm.Items.Item("rt_Charge").Width = oActiveForm.Items.Item("mx_Charge").Width + 9
                                oActiveForm.Items.Item("rt_Charge").Height = oActiveForm.Items.Item("mx_Charge").Height + 124
                                oActiveForm.Items.Item("bt_AmdCh").Top = oActiveForm.Items.Item("mx_Charge").Top + oActiveForm.Items.Item("mx_Charge").Height + 9

                                'Voucher Tab
                                oActiveForm.Items.Item("rt_Voucher").Top = oActiveForm.Items.Item("fo_VoView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_VoView").Height - 1)
                                oActiveForm.Items.Item("rt_Voucher").Width = oActiveForm.Items.Item("mx_Voucher").Width + 9
                                oActiveForm.Items.Item("rt_Voucher").Height = oActiveForm.Items.Item("mx_Voucher").Height + 124
                                oActiveForm.Items.Item("bt_AmdVoc").Top = oActiveForm.Items.Item("mx_Voucher").Top + oActiveForm.Items.Item("mx_Voucher").Height + 16

                                'MSW Voucher Edit Tab Text box under matrix (need to change depends on matrix height and width)
                                oActiveForm.Items.Item("405").Top = oActiveForm.Items.Item("mx_ChCode").Top + oActiveForm.Items.Item("mx_ChCode").Height + 16
                                oActiveForm.Items.Item("ed_VRemark").Top = oActiveForm.Items.Item("405").Top + 2
                                'oActiveForm.Items.Item("ed_VRemark").BackColor = 16645629
                                oActiveForm.Items.Item("406").Top = oActiveForm.Items.Item("ed_VRemark").Top + oActiveForm.Items.Item("ed_VRemark").Height
                                oActiveForm.Items.Item("ed_VPrep").Top = oActiveForm.Items.Item("406").Top + 2
                                oActiveForm.Items.Item("bt_Draft").Top = oActiveForm.Items.Item("406").Top + oActiveForm.Items.Item("406").Height + 16
                                oActiveForm.Items.Item("bt_Cancel").Top = oActiveForm.Items.Item("bt_Draft").Top

                                oActiveForm.Items.Item("399").Top = oActiveForm.Items.Item("405").Top
                                oActiveForm.Items.Item("ed_SubTot").Top = oActiveForm.Items.Item("399").Top + 2
                                oActiveForm.Items.Item("399").Left = oActiveForm.Items.Item("mx_ChCode").Left + oActiveForm.Items.Item("mx_ChCode").Width - 290
                                oActiveForm.Items.Item("ed_SubTot").Left = oActiveForm.Items.Item("399").Left + 125
                                oActiveForm.Items.Item("400").Top = oActiveForm.Items.Item("399").Top + oActiveForm.Items.Item("399").Height
                                oActiveForm.Items.Item("ed_GSTAmt").Top = oActiveForm.Items.Item("400").Top + 2
                                oActiveForm.Items.Item("400").Left = oActiveForm.Items.Item("mx_ChCode").Left + oActiveForm.Items.Item("mx_ChCode").Width - 290
                                oActiveForm.Items.Item("ed_GSTAmt").Left = oActiveForm.Items.Item("399").Left + 125
                                oActiveForm.Items.Item("401").Top = oActiveForm.Items.Item("400").Top + oActiveForm.Items.Item("400").Height
                                oActiveForm.Items.Item("ed_Total").Top = oActiveForm.Items.Item("401").Top + 2
                                oActiveForm.Items.Item("401").Left = oActiveForm.Items.Item("mx_ChCode").Left + oActiveForm.Items.Item("mx_ChCode").Width - 290
                                oActiveForm.Items.Item("ed_Total").Left = oActiveForm.Items.Item("401").Left + 125
                                'End Voucher Edit Tab Text box under matrix (need to change depends on matrix height and width)

                                oItemRet1.Height = oActiveForm.Items.Item("rt_Voucher").Height + 24
                                oItemRetInner.Height = oItemRetInner.Height
                                'End MSW
                                BoolResize = True
                            Catch ex As Exception
                                'To Change
                                'p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        ElseIf BoolResize = True Then
                            Try
                                Dim oItemRet2 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("rt_Outer")
                                'oItemRet2.Top = oActiveForm.Items.Item("mx_TkrList").Top - 29
                                oItemRet2.Top = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_Prmt").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_Prmt").Height - 1)
                                Dim oItemRetInner2 As SAPbouiCOM.Item = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("rt_Inner")
                                oItemRet2.Width = oActiveForm.Items.Item("mx_Cont").Width + 13 ' oActiveForm.Items.Item("mx_TkrList").Width + 20
                                'oItemRet2.Height = oActiveForm.Items.Item("mx_Voucher").Height + 90 '33
                                oItemRet2.Height = 332

                                oItemRetInner2.Top = p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_PMain").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_PMain").Height - 1)
                                oItemRetInner2.Width = oActiveForm.Items.Item("mx_Cont").Width + 9
                                'oItemRetInner2.Height = oActiveForm.Items.Item("mx_Voucher").Height + 13
                                oItemRetInner2.Height = 255

                                oActiveForm.Items.Item("155").Top = oActiveForm.Items.Item("fo_TkrView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_TkrView").Height - 1)
                                oActiveForm.Items.Item("155").Width = oActiveForm.Items.Item("mx_TkrList").Width + 9
                                oActiveForm.Items.Item("155").Height = oActiveForm.Items.Item("mx_TkrList").Height + 126
                                oActiveForm.Items.Item("bt_AmdIns").Top = oActiveForm.Items.Item("mx_TkrList").Top + oActiveForm.Items.Item("mx_TkrList").Height + 9
                                oActiveForm.Items.Item("bt_PrntIns").Top = oActiveForm.Items.Item("mx_TkrList").Top + oActiveForm.Items.Item("mx_TkrList").Height + 9
                                oActiveForm.Items.Item("bt_PrntIns").Left = (oActiveForm.Items.Item("155").Left + oActiveForm.Items.Item("155").Width) - (oActiveForm.Items.Item("bt_PrntIns").Width + 2)

                                'MSW
                                'Container Tab
                                oActiveForm.Items.Item("rt_Cont").Top = oActiveForm.Items.Item("fo_ConView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_ConView").Height - 1)
                                oActiveForm.Items.Item("rt_Cont").Width = oActiveForm.Items.Item("mx_ConTab").Width + 9
                                oActiveForm.Items.Item("rt_Cont").Height = oActiveForm.Items.Item("mx_ConTab").Height + 126
                                oActiveForm.Items.Item("bt_AmdCont").Top = oActiveForm.Items.Item("mx_ConTab").Top + oActiveForm.Items.Item("mx_ConTab").Height + 9

                                'Dispatch Tab
                                oActiveForm.Items.Item("rt_Disp").Top = oActiveForm.Items.Item("fo_DisView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_DisView").Height - 1)
                                oActiveForm.Items.Item("rt_Disp").Width = oActiveForm.Items.Item("mx_DispTab").Width + 9
                                oActiveForm.Items.Item("rt_Disp").Height = oActiveForm.Items.Item("mx_DispTab").Height + 126
                                oActiveForm.Items.Item("bt_AmdDisp").Top = oActiveForm.Items.Item("mx_DispTab").Top + oActiveForm.Items.Item("mx_DispTab").Height + 9

                                'Other Charges Tab
                                oActiveForm.Items.Item("rt_Charge").Top = oActiveForm.Items.Item("fo_ChView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_ChView").Height - 1)
                                oActiveForm.Items.Item("rt_Charge").Width = oActiveForm.Items.Item("mx_Charge").Width + 9
                                oActiveForm.Items.Item("rt_Charge").Height = oActiveForm.Items.Item("mx_Charge").Height + 126
                                oActiveForm.Items.Item("bt_AmdCh").Top = oActiveForm.Items.Item("mx_Charge").Top + oActiveForm.Items.Item("mx_Charge").Height + 9

                                'Voucher Tab
                                oActiveForm.Items.Item("rt_Voucher").Top = oActiveForm.Items.Item("fo_VoView").Top + (p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Items.Item("fo_VoView").Height - 1)
                                oActiveForm.Items.Item("rt_Voucher").Width = oActiveForm.Items.Item("mx_Voucher").Width + 9
                                oActiveForm.Items.Item("rt_Voucher").Height = oActiveForm.Items.Item("mx_Voucher").Height + 126
                                oActiveForm.Items.Item("bt_AmdVoc").Top = oActiveForm.Items.Item("mx_Voucher").Top + oActiveForm.Items.Item("mx_Voucher").Height + 16

                                'MSW Voucher Edit Tab Text box under matrix (need to change depends on matrix height and width)
                                oActiveForm.Items.Item("405").Top = oActiveForm.Items.Item("mx_ChCode").Top + oActiveForm.Items.Item("mx_ChCode").Height + 16
                                oActiveForm.Items.Item("ed_VRemark").Top = oActiveForm.Items.Item("405").Top + 2
                                'oActiveForm.Items.Item("ed_VRemark").BackColor = 16645629
                                oActiveForm.Items.Item("406").Top = oActiveForm.Items.Item("ed_VRemark").Top + oActiveForm.Items.Item("ed_VRemark").Height
                                oActiveForm.Items.Item("ed_VPrep").Top = oActiveForm.Items.Item("406").Top + 2
                                oActiveForm.Items.Item("bt_Draft").Top = oActiveForm.Items.Item("406").Top + oActiveForm.Items.Item("406").Height + 16
                                oActiveForm.Items.Item("bt_Cancel").Top = oActiveForm.Items.Item("bt_Draft").Top

                                oActiveForm.Items.Item("399").Top = oActiveForm.Items.Item("405").Top
                                oActiveForm.Items.Item("ed_SubTot").Top = oActiveForm.Items.Item("399").Top + 2
                                oActiveForm.Items.Item("399").Left = oActiveForm.Items.Item("mx_ChCode").Left + oActiveForm.Items.Item("mx_ChCode").Width - 290
                                oActiveForm.Items.Item("ed_SubTot").Left = oActiveForm.Items.Item("399").Left + 125
                                oActiveForm.Items.Item("400").Top = oActiveForm.Items.Item("399").Top + oActiveForm.Items.Item("399").Height
                                oActiveForm.Items.Item("ed_GSTAmt").Top = oActiveForm.Items.Item("400").Top + 2
                                oActiveForm.Items.Item("400").Left = oActiveForm.Items.Item("mx_ChCode").Left + oActiveForm.Items.Item("mx_ChCode").Width - 290
                                oActiveForm.Items.Item("ed_GSTAmt").Left = oActiveForm.Items.Item("399").Left + 125
                                oActiveForm.Items.Item("401").Top = oActiveForm.Items.Item("400").Top + oActiveForm.Items.Item("400").Height
                                oActiveForm.Items.Item("ed_Total").Top = oActiveForm.Items.Item("401").Top + 2
                                oActiveForm.Items.Item("401").Left = oActiveForm.Items.Item("mx_ChCode").Left + oActiveForm.Items.Item("mx_ChCode").Width - 290
                                oActiveForm.Items.Item("ed_Total").Left = oActiveForm.Items.Item("401").Left + 125
                                'Voucher Edit Tab Text box under matrix (need to change depends on matrix height and width)
                                'End MSW
                                BoolResize = False
                            Catch ex As Exception
                                'p_oSBOApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            End Try
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        Select Case pVal.ItemUID
                            Case "fo_Prmt"
                                oActiveForm.PaneLevel = 7
                                oActiveForm.Items.Item("fo_PMain").Specific.Select()

                            Case "fo_Dsptch"
                                oActiveForm.PaneLevel = 17
                                oActiveForm.Items.Item("fo_DisView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_DispTab").Specific

                                oActiveForm.Items.Item("bt_AmdDisp").Enabled = True
                                If (oMatrix.RowCount < 1) Then
                                    oActiveForm.Items.Item("bt_AmdDisp").Enabled = False
                                ElseIf (oMatrix.RowCount < 2) Then
                                    If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                        oActiveForm.Items.Item("bt_AmdDisp").Enabled = False
                                    End If
                                End If
                            Case "fo_Trkng"
                                oActiveForm.PaneLevel = 6
                                oActiveForm.Items.Item("fo_TkrView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                If oMatrix.RowCount > 1 Then
                                    oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                                ElseIf oMatrix.RowCount = 1 Then
                                    If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                        oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                                    Else
                                        oActiveForm.Items.Item("bt_AmdIns").Enabled = False
                                    End If
                                End If

                            Case ("fo_Cont")
                                oActiveForm.PaneLevel = 15
                                oActiveForm.Items.Item("fo_ConView").Specific.Select()

                                oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                                If oMatrix.RowCount > 1 Then
                                    oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                                ElseIf oMatrix.RowCount = 1 Then
                                    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                        oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                                    Else
                                        oActiveForm.Items.Item("bt_AmdCont").Enabled = False
                                    End If
                                End If

                            Case ("fo_Yard")
                                oActiveForm.PaneLevel = 21
                            Case "fo_Vchr"
                                oActiveForm.PaneLevel = 2
                                oActiveForm.Items.Item("fo_VoView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                                If oMatrix.RowCount > 1 Then
                                    oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ' oActiveForm.Items.Item("bt_PrPDF").Enabled = True 'MSW To Add 18-03-2011
                                    ' oActiveForm.Items.Item("bt_Email").Enabled = True 'MSW To Add 18-03-2011
                                ElseIf oMatrix.RowCount = 1 Then
                                    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                        oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                        'oActiveForm.Items.Item("bt_PrPDF").Enabled = True 'MSW To Add 18-03-2011
                                        'oActiveForm.Items.Item("bt_Email").Enabled = True 'MSW To Add 18-03-2011
                                    Else
                                        'oActiveForm.Items.Item("bt_AmdVoc").Enabled = False
                                        ' oActiveForm.Items.Item("bt_PrPDF").Enabled = False 'MSW To Add 18-03-2011
                                        ' oActiveForm.Items.Item("bt_Email").Enabled = False 'MSW To Add 18-03-2011
                                    End If
                                End If

                            Case "fo_VoView"
                                oActiveForm.PaneLevel = 2
                            Case "fo_VoEdit"
                                oActiveForm.PaneLevel = 3
                                oActiveForm.Items.Item("cb_BnkName").Enabled = False
                                oActiveForm.Items.Item("ed_Cheque").Enabled = False
                                oActiveForm.Items.Item("ed_PayRate").Enabled = False
                            Case "fo_TkrView"
                                oActiveForm.PaneLevel = 6
                            Case "fo_TkrEdit"
                                oActiveForm.PaneLevel = 5
                                oActiveForm.Items.Item("bt_GenPO").Enabled = False
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                'msw
                                If oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then
                                    'Dim sql As String = "SELECT U_WAdLine1,U_WAdLine2,U_WAdLine3,U_WState,U_WPostal,U_WCountry FROM [@OBT_TB003_WRHSE] WHERE Name = " & FormatString(oActiveForm.Items.Item("ed_WName").Specific.Value.ToString)
                                    Dim sql As String = "SELECT U_YAdLine1,U_YAdLine2,U_YAdLine3,U_YState,U_YPostal,U_YCountry FROM [@OBT_TB022_RYARDLIST] WHERE Name = " & FormatString(oActiveForm.Items.Item("ed_YName").Specific.Value.ToString)
                                    oRecordSet.DoQuery(sql)
                                    If oRecordSet.RecordCount > 0 Then
                                        oActiveForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = oRecordSet.Fields.Item("U_YAdLine1").Value.ToString
                                    End If
                                    oRecordSet.DoQuery("SELECT Address FROM OCRD WHERE CardCode = '" & oActiveForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oActiveForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("Address").Value.ToString
                                    End If
                                    oActiveForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")

                                    oActiveForm.Items.Item("op_Inter").Specific.Selected = True 'MSW 01-04-2011
                                End If
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                If oActiveForm.Items.Item("ed_InsDoc").Specific.Value = "" Then
                                    oActiveForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty
                                    If (oMatrix.RowCount > 0) Then
                                        If (oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                            oActiveForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                        Else
                                            oActiveForm.Items.Item("ed_InsDoc").Specific.Value = oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                        End If
                                    Else
                                        oActiveForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                    End If
                                    oActiveForm.Items.Item("ed_Trucker").Specific.Active = True
                                End If

                            Case ("fo_Charge")
                                oActiveForm.PaneLevel = 19
                                oActiveForm.Items.Item("fo_ChView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_Charge").Specific
                                oActiveForm.Items.Item("bt_AmdCh").Enabled = True
                                If (oMatrix.RowCount < 1) Then
                                    oActiveForm.Items.Item("bt_AmdCh").Enabled = False
                                ElseIf (oMatrix.RowCount < 2) Then
                                    If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                        oActiveForm.Items.Item("bt_AmdCh").Enabled = False
                                    End If
                                End If
                            Case "fo_ConView"
                                oActiveForm.PaneLevel = 15
                            Case "fo_ConEdit"
                                oActiveForm.PaneLevel = 16
                            Case "fo_DisView"
                                oActiveForm.PaneLevel = 17
                            Case "fo_DisEdit"
                                oActiveForm.PaneLevel = 18
                                oActiveForm.Freeze(True)
                                ClearText(oActiveForm, "ee_Instru")
                                oActiveForm.Items.Item("ed_DspDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                oActiveForm.Items.Item("ed_DspDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                oActiveForm.Items.Item("ed_DspHr").Specific.Value = Now.ToString("HH:mm")
                                oActiveForm.Items.Item("ch_Dsp").Specific.Checked = False
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                oActiveForm.Items.Item("op_DspExtr").Specific.Selected = True
                                oActiveForm.Items.Item("op_DspIntr").Specific.Selected = True
                                oActiveForm.Items.Item("bt_AddDisp").Specific.Caption = "Add Dispatch"
                                oActiveForm.Items.Item("bt_DelDisp").Enabled = False
                                oActiveForm.Freeze(False)
                            Case "fo_ChView"
                                oActiveForm.PaneLevel = 19

                            Case "fo_ChEdit"
                                oActiveForm.PaneLevel = 20
                                'omm
                                oMatrix = oActiveForm.Items.Item("mx_Charge").Specific
                                If (oMatrix.RowCount > 0) Then
                                    If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                        oActiveForm.Items.Item("ed_CSeqNo").Specific.Value = 1
                                    Else
                                        oActiveForm.Items.Item("ed_CSeqNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                    End If
                                Else
                                    oActiveForm.Items.Item("ed_CSeqNo").Specific.Value = 1
                                End If
                                oActiveForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges"
                                oActiveForm.Items.Item("bt_DelCh").Enabled = False

                                oCombo = oActiveForm.Items.Item("cb_Claim").Specific
                                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                ClearText(oActiveForm, "ed_ChCode", "ed_Remarks")
                            Case "fo_PMain"
                                oActiveForm.PaneLevel = 8
                            Case "fo_PCargo"
                                oActiveForm.PaneLevel = 9
                            Case "fo_PCon"
                                oActiveForm.PaneLevel = 10
                            Case "fo_PInv"
                                oActiveForm.PaneLevel = 11
                            Case "fo_PLic"
                                oActiveForm.PaneLevel = 12
                            Case "fo_PAttach"
                                oActiveForm.PaneLevel = 13
                            Case "fo_PTotal"
                                oActiveForm.PaneLevel = 14
                            Case "fo_ShpInv"
                                oActiveForm.PaneLevel = 25
                                oActiveForm.Items.Item("fo_ShView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_ShpInv").Specific
                                If oMatrix.RowCount > 1 Then
                                    oActiveForm.Items.Item("bt_ShpInv").Enabled = True
                                ElseIf oMatrix.RowCount = 1 Then
                                    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                        oActiveForm.Items.Item("bt_ShpInv").Enabled = True
                                    Else
                                        '    ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                    End If
                                End If
                            Case "fo_ShView"
                                oActiveForm.PaneLevel = 25


                            Case "fo_OpBunk"
                                oActiveForm.PaneLevel = 29
                                oActiveForm.Items.Item("fo_BkView").Specific.Select()
                                oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                            Case "fo_ArmEs"
                                oActiveForm.PaneLevel = 30
                                oActiveForm.Items.Item("fo_ArView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                            Case "fo_Crane"
                                oActiveForm.PaneLevel = 31
                                oActiveForm.Items.Item("fo_CrView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                oActiveForm.Items.Item("bt_CranePO").Enabled = True
                            Case "fo_Fork"
                                oActiveForm.PaneLevel = 32
                                oActiveForm.Items.Item("fo_FkView").Specific.Select()
                                oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                            Case "fo_FkView"
                                oActiveForm.PaneLevel = 32
                            Case "fo_ArView"
                                oActiveForm.PaneLevel = 30
                            Case "fo_CrView"
                                oActiveForm.PaneLevel = 31
                            Case "fo_BkView"
                                oActiveForm.PaneLevel = 29

                        End Select

                        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oActiveForm.Items.Item("ch_POD").Enabled = False
                            'oActiveForm.Items.Item("ed_Wrhse").Enabled = True
                        End If

                        If pVal.ItemUID = "ch_POD" Then
                            If oActiveForm.Items.Item("ch_POD").Specific.Checked = True Then
                                oActiveForm.Items.Item("cb_JbStus").Specific.Select(2, SAPbouiCOM.BoSearchKey.psk_Index)
                            End If
                            If oActiveForm.Items.Item("ch_POD").Specific.Checked = False Then
                                oActiveForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            End If
                        End If

                        If pVal.ItemUID = "bt_AddLic" Then
                            oMatrix = oActiveForm.Items.Item("mx_License").Specific
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB014_PLICINFO").Clear()
                            oMatrix.AddRow(1)
                            oMatrix.FlushToDataSource()
                            oMatrix.Columns.Item("colLicNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                        End If

                        If pVal.ItemUID = "bt_DelLic" Then
                            oMatrix = oActiveForm.Items.Item("mx_License").Specific
                            Dim lRow As Long
                            lRow = oMatrix.GetNextSelectedRow
                            If lRow > -1 Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB014_PLICINFO").RemoveRecord(lRow - 1)
                                    Dim oUserTable As SAPbobsCOM.UserTable
                                    oUserTable = p_oDICompany.UserTables.Item("OBT_TB014_PLICINFO")
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
                            oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB012_PCONTAINE").Clear()
                            oMatrix.AddRow(1)
                            oMatrix.FlushToDataSource()
                            oMatrix.Columns.Item("colCSeqNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                        End If

                        If pVal.ItemUID = "bt_DelCon" Then
                            oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
                            Dim lRow As Long
                            lRow = oMatrix.GetNextSelectedRow
                            If lRow > -1 Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB012_PCONTAINE").RemoveRecord(lRow - 1)
                                    Dim oUserTable As SAPbobsCOM.UserTable
                                    oUserTable = p_oDICompany.UserTables.Item("OBT_TB012_PCONTAINE")
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
                            If oActiveForm.Items.Item("ch_CnUn").Specific.Checked = True Then
                                oActiveForm.Freeze(True)
                                oActiveForm.Items.Item("ed_CnDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                oActiveForm.Items.Item("ed_CnDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                oActiveForm.Items.Item("ed_CnHr").Specific.Value = Now.ToString("HH:mm")
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CnDay").Specific, oActiveForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                oActiveForm.Items.Item("ed_CnHr").Specific.Active = False
                                oActiveForm.Freeze(False)
                            ElseIf oActiveForm.Items.Item("ch_CnUn").Specific.Checked = False Then
                                oActiveForm.Freeze(True)
                                oActiveForm.Items.Item("ed_CnDate").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CnDay").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CnHr").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CnHr").Specific.Active = False
                                oActiveForm.Freeze(False)
                            End If
                        End If

                        If pVal.ItemUID = "ch_CgCl" Then
                            If oActiveForm.Items.Item("ch_CgCl").Specific.Checked = True Then
                                oActiveForm.Freeze(True)
                                oActiveForm.Items.Item("ed_CgDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                oActiveForm.Items.Item("ed_CgDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                oActiveForm.Items.Item("ed_CgHr").Specific.Value = Now.ToString("HH:mm")
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CgDay").Specific, oActiveForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                oActiveForm.Items.Item("ed_CgHr").Specific.Active = False
                                oActiveForm.Freeze(False)
                            ElseIf oActiveForm.Items.Item("ch_CgCl").Specific.Checked = False Then
                                oActiveForm.Freeze(True)
                                oActiveForm.Items.Item("ed_CgDate").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CgDay").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CgHr").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CgHr").Specific.Active = False
                                oActiveForm.Freeze(False)
                            End If
                        End If

                        If pVal.ItemUID = "ch_Dsp" Then
                            If oActiveForm.Items.Item("ch_Dsp").Specific.Checked = True Then
                                oActiveForm.Items.Item("ed_DspCDte").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                oActiveForm.Items.Item("ed_DspCDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                oActiveForm.Items.Item("ed_DspCHr").Specific.Value = Now.ToString("HH:mm")
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            Else
                                oActiveForm.Items.Item("ed_DspCDte").Specific.Value = ""
                                oActiveForm.Items.Item("ed_DspCDay").Specific.Value = ""
                                oActiveForm.Items.Item("ed_DspCHr").Specific.Value = ""
                            End If
                        End If

                        If pVal.ItemUID = "ed_ETADate" And oActiveForm.Items.Item("ed_ETADate").Specific.String <> String.Empty Then
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            oActiveForm.Items.Item("ed_ADate").Specific.Value = oActiveForm.Items.Item("ed_ETADate").Specific.Value
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.ItemUID = "ed_ConLast" And oActiveForm.Items.Item("ed_ConLast").Specific.String <> String.Empty Then
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTim").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTim").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.ItemUID = "ed_CunDate" And oActiveForm.Items.Item("ed_CunDate").Specific.String <> String.Empty Then
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_CunDate").Specific, oActiveForm.Items.Item("ed_CunDay").Specific, oActiveForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CunDay").Specific, oActiveForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                        If pVal.ItemUID = "ed_ConDate" And oActiveForm.Items.Item("ed_ConDate").Specific.String <> String.Empty Then
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_ConDate").Specific, oActiveForm.Items.Item("ed_ConDay").Specific, oActiveForm.Items.Item("ed_ConTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConDay").Specific, oActiveForm.Items.Item("ed_ConTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.ItemUID = "ed_JbDate" And oActiveForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                        'If pVal.ItemUID = "ed_DspDate" And oActiveForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                        '    If DateTime(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        '    If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        'End If
                        If pVal.ItemUID = "ed_DspCDte" And oActiveForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                            If DateTime(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.ItemUID = "bt_Payment" Then
                            p_oSBOApplication.ActivateMenuItem("2818")
                        End If

                        If pVal.ItemUID = "op_Inter" Then
                            oActiveForm.Items.Item("bt_GenPO").Enabled = False
                            oActiveForm.Items.Item("ed_Trucker").Specific.Value = ""
                            If AddChooseFromListByOption(oActiveForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        ElseIf pVal.ItemUID = "op_Exter" Then
                            oActiveForm.Items.Item("ed_Trucker").Specific.Value = ""
                            oActiveForm.Items.Item("ed_TkrTel").Specific.Value = ""
                            oActiveForm.Items.Item("ed_Fax").Specific.Value = ""
                            oActiveForm.Items.Item("ed_Email").Specific.Value = ""
                            oActiveForm.Items.Item("bt_GenPO").Enabled = True
                            If AddChooseFromListByOption(oActiveForm, False, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.ItemUID = "op_DspIntr" Then
                            oCombo = oActiveForm.Items.Item("cb_Dspchr").Specific
                            If ClearComboData(oActiveForm, "cb_Dspchr", "@OBT_TB007_DISPATCH", "U_Dispatch") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
                            oCombo = oActiveForm.Items.Item("cb_Dspchr").Specific
                            If ClearComboData(oActiveForm, "cb_Dspchr", "@OBT_TB007_DISPATCH", "U_Dispatch") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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

                            Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                            UDFAttachForm.Items.Item("U_JobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
                            UDFAttachForm.Items.Item("U_InsDate").Specific.Value = oActiveForm.Items.Item("ed_InsDate").Specific.Value
                        End If

                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("ed_Code").Enabled = False 'MSW
                                    oActiveForm.Items.Item("ed_Name").Enabled = False
                                    oActiveForm.Items.Item("ed_JobNo").Enabled = False 'MSW For Job Type Table
                                    Dim JobLastDocEntry As Integer
                                    Dim ObjectCode As String = String.Empty

                                    'p_oDICompany.GetNewObjectCode(ObjectCode)
                                    'ObjectCode = p_oDICompany.GetNewObjectKey()
                                    sql = "select top 1 Docentry from [@OBT_TB002_IMPSEALCL] order by docentry desc"
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
                                           "," & IIf(oActiveForm.Items.Item("ed_JobNo").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_JobNo").Specific.Value), "NULL") & _
                                            "," & IIf(oActiveForm.Items.Item("ed_JType").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_JType").Specific.Value), "NULL") & _
                                            "," & IIf(oActiveForm.Items.Item("cb_JobType").Specific.Value <> "", FormatString(oActiveForm.Items.Item("cb_JobType").Specific.Value), "NULL") & _
                                            "," & IIf(oActiveForm.Items.Item("cb_JbStus").Specific.Value <> "", FormatString(oActiveForm.Items.Item("cb_JbStus").Specific.Value), "NULL") & _
                                            "," & FrDocEntry & _
                                             "," & IIf(oActiveForm.Items.Item("ed_JbDate").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_JbDate").Specific.Value), "Null") & _
                                                    "," & IIf(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString() <> "", FormatString(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString()), "Null") & _
                                                     "," & IIf(oActiveForm.Items.Item("ed_Code").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_Code").Specific.Value), "Null") & _
                                                      "," & IIf(oActiveForm.Items.Item("ed_Name").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_Name").Specific.Value), "Null") & _
                                                       "," & IIf(oActiveForm.Items.Item("ed_V").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_V").Specific.Value), "Null") & _
                                                    "," & IIf(oActiveForm.Items.Item("ed_ShpAgt").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_ShpAgt").Specific.Value), "Null") & ")"
                                    oRecordSet.DoQuery(sql)

                                    sql = "Update [@OBT_TB002_IMPSEALCL] set U_JbDocNo=" & JobLastDocEntry & " Where DocEntry=" & FrDocEntry & ""
                                    oRecordSet.DoQuery(sql)

                                End If
                            ElseIf pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & oActiveForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & ""
                                oRecordSet.DoQuery(sql)
                            End If
                        End If



                        If pVal.ItemUID = "bt_AddIns" Then
                            If String.IsNullOrEmpty(oActiveForm.Items.Item("ed_Trucker").Specific.String) Then
                                p_oSBOApplication.SetStatusBarMessage("Must Fill Trucker", SAPbouiCOM.BoMessageTime.bmt_Short)
                                BubbleEvent = False
                            Else
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                'If oActiveForm.Items.Item("ed_InsDoc").Specific.Value = vbNullString Then
                                'If oActiveForm.Items.Item("ed_InsDoc").Specific.Value > oMatrix.RowCount Then
                                If oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then
                                    modTrucking.AddUpdateInstructions(oActiveForm, oMatrix, "@OBT_TB006_TRUCKING", True)
                                Else
                                    modTrucking.AddUpdateInstructions(oActiveForm, oMatrix, "@OBT_TB006_TRUCKING", False)
                                    oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                    oActiveForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                End If
                                ClearText(oActiveForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                                oActiveForm.Items.Item("fo_TkrView").Specific.Select()
                                oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                            End If
                        End If

                        If pVal.ItemUID = "bt_DelIns" Then
                            oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                            modTrucking.DeleteByIndex(oActiveForm, oMatrix, "@OBT_TB006_TRUCKING")
                            ClearText(oActiveForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                            oActiveForm.Items.Item("fo_TkrView").Specific.Select()
                        End If

                        If pVal.ItemUID = "bt_AmdIns" Then
                            oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Update Trucking Instruction"
                            oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                            If (oMatrix.GetNextSelectedRow < 0) Then
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction", 1, "OK")
                                Exit Function
                            Else
                                modTrucking.GetDataFromMatrixByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                                modTrucking.SetDataToEditTabByIndex(oActiveForm)
                                oActiveForm.Items.Item("fo_TkrEdit").Specific.Select()
                                oActiveForm.Items.Item("bt_DelIns").Enabled = True 'MSW
                            End If
                            oActiveForm.Items.Item("fo_TkrEdit").Specific.Select()
                        End If

                        If pVal.ItemUID = "fo_TkrView" Then
                            oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                            If oMatrix.RowCount > 1 Then
                                oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                            ElseIf oMatrix.RowCount = 1 Then
                                If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                                Else
                                    oActiveForm.Items.Item("bt_AmdIns").Enabled = False
                                End If
                            End If
                            ClearText(oActiveForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk")
                            oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                            oActiveForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                        End If

                        'MSW 24-02-11 Voucher
                        If pVal.ItemUID = "fo_VoEdit" Then
                            oActiveForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            oActiveForm.Items.Item("op_Cash").Specific.Selected = True
                            oActiveForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
                            oActiveForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)

                            If oActiveForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                                dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
                                For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                    dtmatrix.Rows.Remove(i)
                                Next
                                oMatrix.Clear()
                            End If

                            oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                            If oActiveForm.Items.Item("ed_VocNo").Specific.Value = "" Then
                                If (oMatrix.RowCount > 0) Then
                                    If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                        oActiveForm.Items.Item("ed_VocNo").Specific.Value = 1
                                    Else
                                        oActiveForm.Items.Item("ed_VocNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                    End If
                                Else
                                    oActiveForm.Items.Item("ed_VocNo").Specific.Value = 1
                                End If
                            End If


                            oCombo = oActiveForm.Items.Item("cb_PayCur").Specific
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
                            oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                            If dtmatrix.Rows.Count = 0 Then
                                RowAddToMatrix(oActiveForm, oMatrix)
                            End If
                            oActiveForm.Items.Item("ed_VedName").Specific.Active = True
                        End If

                        If pVal.ItemUID = "fo_VoView" Then
                            oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                            If oMatrix.RowCount > 1 Then
                                oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                'oActiveForm.Items.Item("bt_PrPDF").Enabled = True 'MSW To Add 18-03-2011
                                'oActiveForm.Items.Item("bt_Email").Enabled = True 'MSW To Add 18-03-2011
                            ElseIf oMatrix.RowCount = 1 Then
                                If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                    '   oActiveForm.Items.Item("bt_PrPDF").Enabled = True 'MSW To Add 18-03-2011
                                    '  oActiveForm.Items.Item("bt_Email").Enabled = True 'MSW To Add 18-03-2011
                                Else
                                    'oActiveForm.Items.Item("bt_AmdVoc").Enabled = False
                                    ' oActiveForm.Items.Item("bt_PrPDF").Enabled = False 'MSW To Add 18-03-2011
                                    'oActiveForm.Items.Item("bt_Email").Enabled = False 'MSW To Add 18-03-2011
                                End If
                            End If
                            ClearText(oActiveForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

                            Dim oComboBank As SAPbouiCOM.ComboBox
                            Dim oComboCurrency As SAPbouiCOM.ComboBox

                            oComboBank = oActiveForm.Items.Item("cb_BnkName").Specific
                            oComboCurrency = oActiveForm.Items.Item("cb_PayCur").Specific

                            For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            For j As Integer = oComboCurrency.ValidValues.Count - 1 To 0 Step -1
                                oComboCurrency.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            oActiveForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" 'MSW 23-03-2011
                        End If

                        If pVal.ItemUID = "bt_Draft" Then
                            ''''' Insert Voucher Item Table
                            Dim DocEntry As Integer
                            oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                            'MSW 04-04-2011
                            If oMatrix.RowCount = 1 And oMatrix.Columns.Item("colChCode").Cells.Item(1).Specific.Value() = "" Then
                                Exit Function
                            End If
                            'End MSW 04-04-2011
                            oRecordSet.DoQuery("Delete From [@OBT_TB029_VOCITEM] Where U_JobNo='" & oActiveForm.Items.Item("ed_PJobNo").Specific.Value & "' And U_PVNo='" & oActiveForm.Items.Item("ed_VocNo").Specific.Value & "'") 'MSW To Add 18-03-2011
                            oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [@OBT_TB029_VOCITEM] ")
                            If oRecordSet.RecordCount > 0 Then
                                DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                            End If

                            'MSW - Purchase Voucher Save To Draft 09-03-2010
                            vocTotal = Convert.ToDouble(oActiveForm.Items.Item("ed_Total").Specific.Value)
                            gstTotal = Convert.ToDouble(oActiveForm.Items.Item("ed_GSTAmt").Specific.Value)
                            If oActiveForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                SaveToPurchaseVoucher(oActiveForm, True)
                            Else
                                SaveToPurchaseVoucher(oActiveForm, False)
                            End If
                            SaveToDraftPurchaseVoucher(oActiveForm)
                            'End Purchase Voucher Save To Draft 09-03-2010

                            dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
                            If oMatrix.RowCount > 0 Then
                                For i As Integer = oMatrix.RowCount To 1 Step -1
                                    'To Add Item Code in Insert Statement

                                    sql = "Insert Into [@OBT_TB029_VOCITEM] (DocEntry,LineID,U_JobNo,U_PVNo,U_VSeqNo,U_ChCode,U_AccCode,U_ChDes,U_Amount,U_GST,U_GSTAmt,U_NoGST,U_ItemCode) Values " & _
                                           "(" & DocEntry & _
                                            "," & i & _
                                            "," & IIf(oActiveForm.Items.Item("ed_PJobNo").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_PJobNo").Specific.Value), "NULL") & _
                                            "," & IIf(oActiveForm.Items.Item("ed_VocNo").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_VocNo").Specific.Value), "Null") & _
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


                            oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                            If Convert.ToInt32(oActiveForm.Items.Item("ed_VocNo").Specific.Value) > oMatrix.RowCount Then
                                AddUpdateVoucher(oActiveForm, oMatrix, "@OBT_TB028_VOUC", True)
                            Else
                                AddUpdateVoucher(oActiveForm, oMatrix, "@OBT_TB028_VOUC", False)
                                oActiveForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft"
                            End If


                            ClearText(oActiveForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

                            Dim oComboBank As SAPbouiCOM.ComboBox
                            Dim oComboCurrency As SAPbouiCOM.ComboBox

                            oComboBank = oActiveForm.Items.Item("cb_BnkName").Specific
                            oComboCurrency = oActiveForm.Items.Item("cb_PayCur").Specific

                            For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            For j As Integer = oComboCurrency.ValidValues.Count - 1 To 0 Step -1
                                oComboCurrency.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next


                            oActiveForm.Items.Item("fo_VoView").Specific.Select()
                            oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                            'oActiveForm.Items.Item("bt_PrPDF").Enabled = True 'MSW To Add 18-03-2011
                            'oActiveForm.Items.Item("bt_Email").Enabled = True 'MSW To Add 18-03-2011


                        End If

                        If pVal.ItemUID = "op_Cash" Then
                            oActiveForm.Items.Item("ed_VedName").Specific.Active = True
                            Dim oComboBank As SAPbouiCOM.ComboBox
                            oComboBank = oActiveForm.Items.Item("cb_BnkName").Specific
                            For j As Integer = oComboBank.ValidValues.Count - 1 To 0 Step -1
                                oComboBank.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next

                            oActiveForm.Items.Item("cb_BnkName").Enabled = False
                            oActiveForm.Items.Item("ed_Cheque").Enabled = False
                            oComboBank.Select(0, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            oActiveForm.Items.Item("ed_Cheque").Specific.Value = ""
                        End If

                        If pVal.ItemUID = "op_Cheq" Then
                            Dim oComboBank As SAPbouiCOM.ComboBox
                            oComboBank = oActiveForm.Items.Item("cb_BnkName").Specific
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

                            oActiveForm.Items.Item("cb_BnkName").Enabled = True
                            oActiveForm.Items.Item("ed_Cheque").Enabled = True
                        End If

                        If pVal.ItemUID = "bt_Cancel" Then
                            oActiveForm.Items.Item("fo_VoView").Specific.Select()
                        End If

                        If pVal.ItemUID = "mx_Voucher" And pVal.ColUID = "V_-1" Then
                            oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                            If oMatrix.GetNextSelectedRow > 0 Then
                                If (oMatrix.IsRowSelected(oMatrix.GetNextSelectedRow)) = True Then
                                    GetVoucherDataFromMatrixByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                                End If
                            Else
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Payment Voucher.", 1, "&OK")
                            End If
                        End If

                        If pVal.ItemUID = "bt_AmdVoc" Then
                            'POP UP Payment Voucher
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                LoadPaymentVoucher(oActiveForm)
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Voucher.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If
                        End If

                        'End MSW
                        'Shipping Invoice POP Up
                        If pVal.ItemUID = "bt_ShpInv" Then
                            'POP UP Shipping Invoice
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                LoadShippingInvoice(oActiveForm)
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If

                        End If

                        'Purchase Order and Goods Receipt POP UP

                        If pVal.ItemUID = "bt_ForkPO" Then 'sw
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                LoadAndCreateCPO(oActiveForm, "ForkPurchaseOrder.srf")
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If

                        End If
                        If pVal.ItemUID = "bt_ArmePO" Then 'sw
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                LoadAndCreateCPO(oActiveForm, "ArmedPurchaseOrder.srf")
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If

                        End If

                        If pVal.ItemUID = "bt_CranePO" Then 'sw
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                LoadAndCreateCPO(oActiveForm, "CranePurchaseOrder.srf")
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If

                        End If

                        If pVal.ItemUID = "bt_BunkPO" Then 'sw
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                                LoadAndCreateCPO(oActiveForm, "BunkPurchaseOrder.srf")
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If

                        End If
                        '-----------------------------------------------------------------------------------------------------------'


                        'MSW To Add 18-03-2011
                        If pVal.ItemUID = "bt_Email" Then
                            Dim rptDocument As ReportDocument = New ReportDocument()
                            Dim CrxReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                            Dim dtPayment As New System.Data.DataTable
                            Dim sqlconnect As SqlClient.SqlConnection = New SqlClient.SqlConnection("Data Source=midserver;Initial Catalog=SBODEMOSG;User ID=sa;Password=root")
                            sqlconnect.Open()
                            Dim sqladp As New SqlClient.SqlDataAdapter("Select * From [@OBT_TB029_VOCITEM] Where U_JobNo='" & oActiveForm.Items.Item("ed_JobNo").Specific.Value & "'", sqlconnect)
                            sqladp.Fill(dtPayment)
                            sqlconnect.Close()
                            Dim adjunto As String = "D:\VoucherItem.pdf"
                            Dim rptPath As String = "D:\VoucherItem.rpt"
                            'rptDocument.Load(sPathReporte)
                            Dim crExportOptions As New CrystalDecisions.Shared.ExportOptions
                            Dim crDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                            rptDocument.Load(rptPath)
                            rptDocument.Refresh()
                            rptDocument.SetDatabaseLogon("sa", "root")
                            rptDocument.SetDataSource(dtPayment)
                            crDiskFileDestinationOptions.DiskFileName = adjunto
                            crExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                            crExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                            crExportOptions.ExportDestinationOptions = crDiskFileDestinationOptions
                            rptDocument.Export(crExportOptions)
                            Dim msg As New MailMessage("maysandar@sap-infotech.com", "ohnmar@sap-infotech.com", "Testing PDF", "Testing ")
                            Dim smtpser As New SmtpClient("MIDSERVER", 25)

                            Dim attach As System.Net.Mail.Attachment = New System.Net.Mail.Attachment(adjunto)
                            msg.Attachments.Add(attach)
                            smtpser.Send(msg)
                            rptDocument.Close()
                            'System.Diagnostics.Process.Start("D:\VoucherItem.pdf")
                        End If
                        If pVal.ItemUID = "bt_PrPDF" Then
                            Dim rptDocument As ReportDocument = New ReportDocument()
                            Dim CrxReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                            Dim dtPayment As New System.Data.DataTable
                            Dim sqlconnect As SqlClient.SqlConnection = New SqlClient.SqlConnection("Data Source=midserver;Initial Catalog=SBODEMOSG;User ID=sa;Password=root")
                            sqlconnect.Open()
                            Dim sqladp As New SqlClient.SqlDataAdapter("Select * From [@OBT_TB029_VOCITEM] Where U_JobNo='" & oActiveForm.Items.Item("ed_JobNo").Specific.Value & "'", sqlconnect)
                            sqladp.Fill(dtPayment)
                            sqlconnect.Close()
                            Dim adjunto As String = "D:\VoucherItem1.pdf"
                            Dim rptPath As String = "D:\VoucherItem.rpt"
                            'rptDocument.Load(sPathReporte)
                            Dim crExportOptions As New CrystalDecisions.Shared.ExportOptions
                            Dim crDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions


                            rptDocument.Load(rptPath)

                            rptDocument.Refresh()
                            rptDocument.SetDatabaseLogon("sa", "root")
                            rptDocument.SetDataSource(dtPayment)

                            crDiskFileDestinationOptions.DiskFileName = adjunto
                            crExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                            crExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                            crExportOptions.ExportDestinationOptions = crDiskFileDestinationOptions
                            rptDocument.Export(crExportOptions)
                            rptDocument.Close()
                            System.Diagnostics.Process.Start("D:\VoucherItem1.pdf")

                        End If

                        'End MSW To Add 18-03-2011

                        'MSW 11-02-2011 Container View List & Edit
                        If pVal.ItemUID = "ch_CStuff" Then
                            Dim strTime As SAPbouiCOM.EditText
                            If oActiveForm.Items.Item("ch_CStuff").Specific.Checked = True Then
                                oActiveForm.Freeze(True)
                                oActiveForm.Items.Item("ed_CunDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                oActiveForm.Items.Item("ed_CunDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                'oActiveForm.DataSources.DBDataSources.Item("@OBT_TB020_HCONTAINE").SetValue("U_ContTime", oActiveForm.DataSources.DBDataSources.Item("@OBT_TB020_HCONTAINE").Offset, Now.ToString("HH:mm"))
                                strTime = oActiveForm.Items.Item("ed_CunTime").Specific
                                strTime.Value = Now.ToString("HH:mm")
                                'oActiveForm.Items.Item("ed_CunTime").Specific.Value = Now.ToString("HH:mm")
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CunDay").Specific, oActiveForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                '  oActiveForm.Items.Item("ed_CunTime").Specific.Active = True
                                oActiveForm.Freeze(False)
                            ElseIf oActiveForm.Items.Item("ch_CStuff").Specific.Checked = False Then
                                oActiveForm.Freeze(True)
                                oActiveForm.Items.Item("ed_CunDate").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CunDay").Specific.Value = vbNullString
                                oActiveForm.Items.Item("ed_CunTime").Specific.Value = vbNullString
                                ' oActiveForm.Items.Item("ed_CunTime").Specific.Active = True
                                oActiveForm.Freeze(False)
                            End If
                        End If

                        If pVal.ItemUID = "bt_AddCont" Then
                            oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                            'If Convert.ToInt32(oActiveForm.Items.Item("ed_ConNo").Specific.Value) > oMatrix.RowCount Then
                            If oActiveForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container" Then
                                AddUpdateContainer(oActiveForm, oMatrix, "@OBT_TB020_HCONTAINE", True)
                            Else
                                AddUpdateContainer(oActiveForm, oMatrix, "@OBT_TB020_HCONTAINE", False)
                                oActiveForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container"
                                oActiveForm.Items.Item("bt_DelCont").Enabled = False 'MSW
                            End If

                            ClearText(oActiveForm, "ed_ConNo", "ed_ContNo", "ed_SealNo", "ed_ContWt", "ed_CDesc", "ed_CunDate", "ed_CunDay", "ed_CunTime")
                            Dim oComboType As SAPbouiCOM.ComboBox
                            Dim oComboSize As SAPbouiCOM.ComboBox

                            oComboType = oActiveForm.Items.Item("cb_ConType").Specific
                            oComboSize = oActiveForm.Items.Item("cb_ConSize").Specific

                            For j As Integer = oComboType.ValidValues.Count - 1 To 0 Step -1
                                oComboType.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next

                            oComboSize.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            oActiveForm.Items.Item("fo_ConView").Specific.Select()
                            oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                            UpdateNoofContainer(oActiveForm, oMatrix)
                        End If

                        If pVal.ItemUID = "fo_ConEdit" Then
                            Dim oComboType As SAPbouiCOM.ComboBox
                            Dim oComboSize As SAPbouiCOM.ComboBox
                            oComboType = oActiveForm.Items.Item("cb_ConType").Specific
                            oComboSize = oActiveForm.Items.Item("cb_ConSize").Specific

                            If oComboType.ValidValues.Count = 0 Then
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("select U_ContType  from [@OBT_TB021_CONT] group by U_ContType")
                                If oRecordSet.RecordCount > 0 Then
                                    oRecordSet.MoveFirst()
                                    While oRecordSet.EoF = False
                                        oComboType.ValidValues.Add(oRecordSet.Fields.Item("U_ContType").Value, "")
                                        oRecordSet.MoveNext()
                                    End While
                                End If
                            End If
                            If oComboSize.ValidValues.Count = 0 Then
                                oComboSize.ValidValues.Add("20'", "20'")
                                oComboSize.ValidValues.Add("40'", "40'")
                                oComboSize.ValidValues.Add("45'", "45'")
                            End If
                            oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                            ' If oActiveForm.Items.Item("ed_ConNo").Specific.Value = "" Then
                            If (oMatrix.RowCount > 0) Then
                                If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                    oActiveForm.Items.Item("ed_ConNo").Specific.Value = 1
                                Else
                                    oActiveForm.Items.Item("ed_ConNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                End If
                            Else
                                oActiveForm.Items.Item("ed_ConNo").Specific.Value = 1
                            End If
                            'End If
                            If oActiveForm.Items.Item("ed_CunDate").Specific.Value = "" Then
                                oActiveForm.Items.Item("ch_CStuff").Specific.Checked = False
                            End If
                            oActiveForm.Items.Item("ed_ContNo").Specific.Active = True
                        End If

                        If pVal.ItemUID = "fo_ConView" Then
                            oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                            If oMatrix.RowCount > 1 Then
                                oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                            ElseIf oMatrix.RowCount = 1 Then
                                If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    oActiveForm.Items.Item("bt_AmdCont").Enabled = True
                                Else
                                    oActiveForm.Items.Item("bt_AmdCont").Enabled = False
                                End If
                            End If
                            ClearText(oActiveForm, "ed_ConNo", "ed_ContNo", "ed_SealNo", "ed_ContWt", "ed_CDesc", "ed_CunDate", "ed_CunDay", "ed_CunTime")
                            Dim oComboType As SAPbouiCOM.ComboBox
                            Dim oComboSize As SAPbouiCOM.ComboBox

                            oComboType = oActiveForm.Items.Item("cb_ConType").Specific
                            oComboSize = oActiveForm.Items.Item("cb_ConSize").Specific

                            For j As Integer = oComboType.ValidValues.Count - 1 To 0 Step -1
                                oComboType.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next
                            oComboSize.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            oActiveForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container"
                            oActiveForm.Items.Item("bt_DelCont").Enabled = False 'MSW
                        End If

                        If pVal.ItemUID = "bt_DelCont" Then
                            oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                            DeleteContainerByIndex(oActiveForm, oMatrix, "@OBT_TB020_HCONTAINE")
                            ClearText(oActiveForm, "ed_ConNo", "ed_ContNo", "ed_SealNo", "ed_ContWt", "ed_CDesc", "ed_CunDate", "ed_CunDay", "ed_CunTime")
                            Dim oComboType As SAPbouiCOM.ComboBox
                            Dim oComboSize As SAPbouiCOM.ComboBox

                            oComboType = oActiveForm.Items.Item("cb_ConType").Specific
                            oComboSize = oActiveForm.Items.Item("cb_ConSize").Specific

                            For j As Integer = oComboType.ValidValues.Count - 1 To 0 Step -1
                                oComboType.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                            Next

                            oComboSize.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            oActiveForm.Items.Item("fo_ConView").Specific.Select()
                            UpdateNoofContainer(oActiveForm, oMatrix)
                        End If

                        If pVal.ItemUID = "bt_AmdCont" Then
                            Dim oComboType As SAPbouiCOM.ComboBox
                            Dim oComboSize As SAPbouiCOM.ComboBox
                            oComboType = oActiveForm.Items.Item("cb_ConType").Specific
                            oComboSize = oActiveForm.Items.Item("cb_ConSize").Specific
                            oActiveForm.Items.Item("bt_AddCont").Specific.Caption = "Update Container"
                            oActiveForm.Items.Item("bt_DelCont").Enabled = True 'MSW
                            oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                            If oMatrix.GetNextSelectedRow < 0 Then
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Container.", 1, "&OK")
                                Exit Function
                            End If
                            If oComboType.ValidValues.Count = 0 Then
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("select U_ContType  from [@OBT_TB021_CONT] group by U_ContType")
                                If oRecordSet.RecordCount > 0 Then
                                    oRecordSet.MoveFirst()
                                    While oRecordSet.EoF = False
                                        oComboType.ValidValues.Add(oRecordSet.Fields.Item("U_ContType").Value, "")
                                        oRecordSet.MoveNext()
                                    End While
                                End If
                            End If
                            If oComboSize.ValidValues.Count = 0 Then
                                oComboSize.ValidValues.Add("20'", "20'")
                                oComboSize.ValidValues.Add("40'", "40'")
                                oComboSize.ValidValues.Add("45'", "45'")
                            End If
                            oActiveForm.Items.Item("fo_ConEdit").Specific.Select()
                            GetContainerDataFromMatrixByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                            SetContainerDataToEditTabByIndex(oActiveForm)
                        End If

                        If pVal.ItemUID = "mx_ConTab" And pVal.ColUID = "V_-1" Then
                            oMatrix = oActiveForm.Items.Item("mx_ConTab").Specific
                            If oMatrix.GetNextSelectedRow > 0 Then
                                If (oMatrix.IsRowSelected(oMatrix.GetNextSelectedRow)) = True Then
                                    GetContainerDataFromMatrixByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                                End If
                            Else
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Container.", 1, "&OK")
                            End If
                        End If

                        ''End Container View List & Edit

                        ''----------------------------------Dipatch Tab-------------------------------------------------------------------------------'
                        If pVal.ItemUID = "bt_AddDisp" Then

                            oMatrix = oActiveForm.Items.Item("mx_DispTab").Specific

                            If oActiveForm.Items.Item("bt_AddDisp").Specific.Caption = "Add Dispatch" Then
                                AddUpdateDisp(oActiveForm, oMatrix, "@OBT_TB007_DISPATCH", True, 0)

                            Else
                                AddUpdateDisp(oActiveForm, oMatrix, "@OBT_TB007_DISPATCH", False, oMatrix.GetNextSelectedRow)
                                oActiveForm.Items.Item("bt_AddDisp").Specific.Caption = "Add Dispatch"
                            End If

                            oActiveForm.Items.Item("op_DspIntr").Specific.Selected = True
                            ClearText(oActiveForm, "ee_Instru", "ed_DspCDte", "ed_DspCDay", "ed_DspCHr")
                            oActiveForm.Items.Item("bt_AmdDisp").Enabled = True
                            oActiveForm.Items.Item("fo_DisView").Specific.Select()

                        End If

                        If pVal.ItemUID = "bt_DelDisp" Then
                            oMatrix = oActiveForm.Items.Item("mx_DispTab").Specific
                            DeleteByIndex(oActiveForm, oMatrix, "@OBT_TB007_DISPATCH")
                            ClearText(oActiveForm, "ee_Instru", "ed_DspCDte", "ed_DspCDay", "ed_DspCHr")
                            oActiveForm.Items.Item("fo_DisView").Specific.Select()
                        End If

                        If pVal.ItemUID = "bt_AmdDisp" Then
                            oMatrix = oActiveForm.Items.Item("mx_DispTab").Specific
                            If (oMatrix.GetNextSelectedRow < 0) Then
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Dispatch", 1, "OK")
                            Else
                                oActiveForm.Items.Item("fo_DisEdit").Specific.Select()
                                oActiveForm.Items.Item("bt_AddDisp").Specific.Caption = "Update Dispatch"
                                oActiveForm.Items.Item("bt_DelDisp").Enabled = True
                                oActiveForm.Freeze(True)
                                SetDispatchDataToEditTabByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                                oActiveForm.Freeze(False)
                            End If

                        End If

                        ''------------------------------------------------------------------------------------------------------------------------'
                        ''----------------------------------Other Charges Tab-------------------------------------------------------------------------------'
                        If pVal.ItemUID = "bt_AddCh" Then

                            oMatrix = oActiveForm.Items.Item("mx_Charge").Specific

                            If oActiveForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges" Then
                                AddUpdateOtherCharges(oActiveForm, oMatrix, "@OBT_TB024_HCHARGES", True, 0)
                            Else
                                AddUpdateOtherCharges(oActiveForm, oMatrix, "@OBT_TB024_HCHARGES", False, oMatrix.GetNextSelectedRow)
                                oActiveForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges"
                            End If
                            ClearText(oActiveForm, "ed_CSeqNo", "ed_ChCode", "ed_Remarks")
                            oActiveForm.Items.Item("bt_AmdCh").Enabled = True
                            oActiveForm.Items.Item("fo_ChView").Specific.Select()

                        End If

                        If pVal.ItemUID = "bt_DelCh" Then
                            oMatrix = oActiveForm.Items.Item("mx_Charge").Specific
                            DeleteByIndexOtherCharges(oActiveForm, oMatrix, "@OBT_TB024_HCHARGES")
                            ClearText(oActiveForm, "ed_CSeqNo", "ed_ChCode", "ed_Remarks")
                            oActiveForm.Items.Item("fo_ChView").Specific.Select()
                        End If

                        If pVal.ItemUID = "bt_AmdCh" Then
                            oMatrix = oActiveForm.Items.Item("mx_Charge").Specific
                            If (oMatrix.GetNextSelectedRow < 0) Then
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Other Charges", 1, "OK")
                            Else
                                oActiveForm.Items.Item("fo_ChEdit").Specific.Select()
                                oActiveForm.Items.Item("bt_AddCh").Specific.Caption = "Update Charges"
                                oActiveForm.Items.Item("bt_DelCh").Enabled = True
                                oActiveForm.Freeze(True)
                                SetOtherChargesDataToEditTabByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                                oActiveForm.Freeze(False)
                            End If

                        End If

                        ''------------------------------------------------------------------------------------------------------------------------'

                        If pVal.ItemUID = "mx_Cont" And pVal.BeforeAction = False And oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If

                    End If



                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                        If pVal.ItemUID = "1" Then
                            Dim PODFlag As String = String.Empty
                            Dim JbStus As String = String.Empty
                            Dim DispatchComplete As String = String.Empty
                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                'develivery process by POD[Proof Of Delivery] check box
                                'oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_TB002_IMPSEALCL] WHERE DocEntry = " & oActiveForm.Items.Item("ed_JobNo").Specific.Value)
                                oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_TB002_IMPSEALCL] WHERE DocEntry = " & oActiveForm.Items.Item("ed_DocNum").Specific.Value) 'MSW 01-06-2011 for job Type Table
                                If oRecordSet.RecordCount > 0 Then
                                    JbStus = oRecordSet.Fields.Item("U_JbStus").Value
                                    PODFlag = oRecordSet.Fields.Item("U_POD").Value
                                End If
                                If oActiveForm.Items.Item("ch_POD").Specific.Checked = True And JbStus = "Open" Then
                                    If p_oSBOApplication.MessageBox("Make sure that all entries trucking and vouchers are completed.(ensure no draft Payment in this job and " & _
                                                               "ensure all external trucking transaction has generated the PO). Cannot edit or add after click POD check box. " & _
                                                               "Do you want to continue?", 1, "&Yes", "&No") = 2 Then
                                        BubbleEvent = False
                                    End If
                                End If
                                If BubbleEvent = True Then
                                    'MSW 01-06-2011 for job type table
                                    sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & oActiveForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & ""
                                    oRecordSet.DoQuery(sql)
                                    'End MSW 01-06-2011 for job type table
                                End If
                            End If
                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                'handle for dispatch complete check box
                                'Not Used for multiple dispatch
                                'oRecordSet.DoQuery("SELECT U_Complete FROM [@OBT_TB007_DISPATCH] WHERE DocEntry = " & oActiveForm.Items.Item("ed_JobNo").Specific.Value)
                                'If oRecordSet.RecordCount > 0 Then
                                '    DispatchComplete = oRecordSet.Fields.Item("U_Complete").Value
                                'End If
                                'If oActiveForm.Items.Item("ch_Dsp").Specific.Checked = True And DispatchComplete = "Y" Then
                                '    BubbleEvent = False
                                '    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                'End If
                                If oActiveForm.Items.Item("ch_POD").Specific.Checked = True And PODFlag = "Y" Then
                                    BubbleEvent = False
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                End If
                            End If
                            If oActiveForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oActiveForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If Validateforform(" ", oActiveForm) Then
                                    BubbleEvent = False
                                End If
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                        '-------------------------For Payment(omm)------------------------------------------'
                        If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "V_-1" Then
                            oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                            If pVal.Row > 0 Then
                                If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                    gridindex = CInt(pVal.Row)
                                End If
                            End If
                        End If
                        '----------------------------------------------------------------------------------'
                        If pVal.ItemUID = "mx_TkrList" And pVal.ColUID = "V_1" Then
                            oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                            If pVal.Row > 0 Then
                                If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                    modTrucking.rowIndex = CInt(pVal.Row)
                                    modTrucking.GetDataFromMatrixByIndex(oActiveForm, oMatrix, modTrucking.rowIndex)
                                End If
                            End If
                        End If
                    End If

                    'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.ItemUID = "cb_ConSize" And pVal.Before_Action = True Then
                    '    'Dim oComboType As SAPbouiCOM.ComboBox
                    '    Dim oComboSize As SAPbouiCOM.ComboBox
                    '    Dim sComType As String = String.Empty
                    '    If oActiveForm.Items.Item("cb_ConType").Specific.Value <> "" Then
                    '        sComType = oActiveForm.Items.Item("cb_ConType").Specific.Value
                    '        oComboSize = oActiveForm.Items.Item("cb_ConSize").Specific
                    '        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    '        oRecordSet.DoQuery("select U_Size from [@OBT_TB021_CONT] where U_Type='" & sComType & "'")
                    '        While oComboSize.ValidValues.Count > 0
                    '            oComboSize.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                    '        End While
                    '        If oRecordSet.RecordCount > 0 Then
                    '            oRecordSet.MoveFirst()
                    '            While oRecordSet.EoF = False
                    '                oComboSize.ValidValues.Add(oRecordSet.Fields.Item("U_Size").Value, "")
                    '                oRecordSet.MoveNext()
                    '            End While
                    '        End If
                    '    End If

                    'End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize" And pVal.Before_Action = True And pVal.Row <> 0 Then
                        oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
                        Dim oColCombo As SAPbouiCOM.Column
                        Dim omatCombo As SAPbouiCOM.ComboBox
                        oColCombo = oMatrix.Columns.Item("colCSize")
                        omatCombo = oColCombo.Cells.Item(pVal.Row).Specific
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Value) Then
                            Dim type As String = oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Selected.Value
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select U_Size from [@OBT_TB008_CONTAINER] where U_Type='" & type & "'")
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

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "ed_JobNo" And pVal.CharPressed = 13 Then
                            oActiveForm.Items.Item("1").Click()
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                        Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                        Try
                            '-------------------------For Payment(omm)------------------------------------------'
                            If pVal.ColUID = "colChCode" Then
                                oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                                dtmatrix = oActiveForm.DataSources.DataTables.Item("DTCharges")
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
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)
                                oActiveForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString
                                If (oActiveForm.Items.Item("ed_ETADate").Specific.Value = "") Then
                                    oActiveForm.Items.Item("ed_ETADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    oActiveForm.Items.Item("ed_ETADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                    oActiveForm.Items.Item("ed_ETAHr").Specific.Value = Now.ToString("HH:mm")
                                    If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    oActiveForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    oActiveForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                    oActiveForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                    If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                EnabledHeaderControls(oActiveForm, True)
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oActiveForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                    oActiveForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                                    oActiveForm.Items.Item("ed_Yard").Specific.ChooseFromListUID = "YARD"
                                    oActiveForm.Items.Item("ed_Yard").Specific.ChooseFromListAlias = "Code"
                                    oActiveForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                                    oActiveForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                                    oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                    oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                End If
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                If oRecordSet.RecordCount > 0 Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_IComName", 0, oDataTable.GetValue(1, 0).ToString)
                                End If
                            End If
                            If pVal.ItemUID = "ed_ShpAgt" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                If oRecordSet.RecordCount > 0 Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString)
                                End If
                            End If
                            If pVal.ItemUID = "ed_ChCode" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB024_HCHARGES").SetValue("U_CCode", oActiveForm.DataSources.DBDataSources.Item("@OBT_TB024_HCHARGES").Offset, oDataTable.Columns.Item("U_CName").Cells.Item(0).Value.ToString)
                            End If
                            If pVal.ItemUID = "ed_Yard" Then
                                oActiveForm.Items.Item("fo_Yard").Specific.Select()
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_YName", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YName", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YTel", 0, oDataTable.Columns.Item("U_YTel").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YCPrson", 0, oDataTable.Columns.Item("U_YPerson").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YMobile", 0, oDataTable.Columns.Item("U_YMobile").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YHr", 0, oDataTable.Columns.Item("U_YwhL1").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YLat", 0, oDataTable.Columns.Item("U_YLat").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YLong", 0, oDataTable.Columns.Item("U_YLong").Cells.Item(0).Value.ToString)
                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim sql As String = "SELECT U_YAdLine1,U_YAdLine2,U_YAdLine3,U_YState,U_YPostal,U_YCountry FROM [@OBT_TB022_RYARDLIST] WHERE Name = " & FormatString(oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                Dim WAddress As String = String.Empty
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    WAddress = Trim(oRecordSet.Fields.Item("U_YAdLine1").Value) & Chr(13) & _
                                                    Trim(oRecordSet.Fields.Item("U_YAdLine2").Value) & Chr(13) & _
                                                    Trim(oRecordSet.Fields.Item("U_YAdLine3").Value) & Chr(13) & _
                                                    Trim(oRecordSet.Fields.Item("U_YState").Value) & Chr(13) & _
                                                    Trim(oRecordSet.Fields.Item("U_YPostal").Value) & Chr(13) & _
                                                    Trim(oRecordSet.Fields.Item("U_YCountry").Value)
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB025_RYARDTAB").SetValue("U_YAddr", 0, WAddress)
                                    oActiveForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = WAddress
                                End If
                            End If
                            If pVal.ItemUID = "ed_Trucker" Then
                                If oActiveForm.Items.Item("op_Inter").Specific.Selected = True Then
                                    oActiveForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                    oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                    oActiveForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                    oActiveForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                    oActiveForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                End If
                            End If

                            If pVal.ItemUID = "ed_Vessel" Or pVal.ItemUID = "ed_Voy" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_VName", 0, oActiveForm.Items.Item("ed_Vessel").Specific.String)
                            End If

                            If pVal.ItemUID = "ed_CurCode" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_CurCode", 0, oDataTable.GetValue(0, 0).ToString)
                                Dim Rate As String = String.Empty
                                sql = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                        Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                        CInt(Today.ToString("dd"))
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    Rate = oRecordSet.Fields.Item("Rate").Value
                                End If
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_ExRate", 0, Rate.ToString)
                            End If
                            If pVal.ItemUID = "ed_CCharge" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_Cchange", 0, oDataTable.GetValue(0, 0).ToString)

                                Dim Rate As String = String.Empty
                                sql = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                        Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                        CInt(Today.ToString("dd"))
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    Rate = oRecordSet.Fields.Item("Rate").Value
                                End If
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_ChRate", 0, Rate.ToString)

                            End If
                            If pVal.ItemUID = "ed_Charge" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_FCchange", 0, oDataTable.GetValue(0, 0).ToString)
                            End If
                            'MSW Voucher
                            If pVal.ItemUID = "ed_VedName" Then
                                ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC") 'MSW To Add 18-3-2011
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                                                                                       & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString()) 'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*

                                vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft
                            End If
                            'End

                        Catch ex As Exception
                        End Try
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False Then


                        If pVal.ItemUID = "cb_PCode" Then
                            oCombo = oActiveForm.Items.Item("cb_PCode").Specific
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_PCode", 0, oCombo.Selected.Value.ToString)
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "ed_TkrDate" And pVal.Before_Action = False Then
                            Dim strTime As SAPbouiCOM.EditText
                            strTime = oActiveForm.Items.Item("ed_TkrTime").Specific
                            strTime.Value = Now.ToString("HH:mm")
                        End If
                        If pVal.ItemUID = "ed_InvNo" Then
                            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_InvNo", 0, oActiveForm.Items.Item("ed_InvNo").Specific.String)
                        End If
                        If BubbleEvent = False Then
                            Validateforform(pVal.ItemUID, oActiveForm)
                        End If

                    End If
            End Select
            DoImportSeaFCLItemEvent = RTN_SUCCESS
        Catch ex As Exception
            DoImportSeaFCLItemEvent = RTN_ERROR
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function DoImportSeaFCLRightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoImportSeaFCLItemEvent
        '   Purpose     :   This function will be providing to proceed validating for
        '                   Inventory [All] Menu Event information
        '               f
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
        DoImportSeaFCLRightClickEvent = RTN_ERROR
        Try

            DoImportSeaFCLRightClickEvent = RTN_SUCCESS
        Catch ex As Exception
            DoImportSeaFCLRightClickEvent = RTN_ERROR
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Sub LoadImportSeaFCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        Dim oActiveForm, oPayForm, oShpForm As SAPbouiCOM.Form
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oCheckBox As SAPbouiCOM.CheckBox
        Dim oOpt As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = ""
        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String = "LoadImportSeaFCLForm"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not LoadFromXML(p_oSBOApplication, "ImportSeaFCLv1.srf") Then Throw New ArgumentException(sErrDesc)
            oActiveForm = p_oSBOApplication.Forms.ActiveForm

            'BubbleEvent = False
            '-------------------------------------Charge Code--------------------------------------------------------------'
            If AddChooseFromList(oActiveForm, "Charge", False, "UDOCHCODE") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_ChCode").Specific.ChooseFromListUID = "Charge"
            oActiveForm.Items.Item("ed_ChCode").Specific.ChooseFromListAlias = "U_CName"
            '-----------------------------------------------------------------------------------------------------------------'
            'Vendor in Voucher

            'End
            oActiveForm.EnableMenu("1288", True)
            oActiveForm.EnableMenu("1289", True)
            oActiveForm.EnableMenu("1290", True)
            oActiveForm.EnableMenu("1291", True)
            oActiveForm.EnableMenu("1284", False)
            oActiveForm.EnableMenu("1286", False)

            oActiveForm.DataBrowser.BrowseBy = "ed_DocNum"
            oActiveForm.Items.Item("fo_Prmt").Specific.Select()

            oActiveForm.Freeze(True)
            oActiveForm.Items.Item("bt_GenPO").Enabled = False
            oActiveForm.Items.Item("bt_AmdCont").Enabled = False 'MSW
            oActiveForm.Items.Item("bt_DelCont").Enabled = False 'MSW
            oActiveForm.Items.Item("bt_DelIns").Enabled = False 'MSW
            EnabledHeaderControls(oActiveForm, False)
            EnabledMaxtix(oActiveForm, oActiveForm.Items.Item("mx_TkrList").Specific, False)
            EnabledDispatchMatrix(oActiveForm, oActiveForm.Items.Item("mx_DispTab").Specific, False)
            oActiveForm.PaneLevel = 7
            oActiveForm.Items.Item("ed_JType").Specific.Value = "Import Sea FCL"
            oActiveForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
            oActiveForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
            oActiveForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")

            '-----------------DISPATCH Date----------------------------------------------------------------------'
            'oActiveForm.Items.Item("ed_DspDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
            'oActiveForm.Items.Item("ed_DspDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
            'oActiveForm.Items.Item("ed_DspHr").Specific.Value = Now.ToString("HH:mm")
            'If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            '-----------------DISPATCH Date----------------------------------------------------------------------'

            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddChooseFromList(oActiveForm, "cflBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(oActiveForm, "cflBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Code").Specific.ChooseFromListUID = "cflBP"
            oActiveForm.Items.Item("ed_Code").Specific.ChooseFromListAlias = "CardCode"
            oActiveForm.Items.Item("ed_Name").Specific.ChooseFromListUID = "cflBP2"
            oActiveForm.Items.Item("ed_Name").Specific.ChooseFromListAlias = "CardName"

            If AddChooseFromList(oActiveForm, "cflBP3", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(oActiveForm, "YARD", False, "UDORYARD") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(oActiveForm, "DSVES01", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(oActiveForm, "DSVES02", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
            oActiveForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
            oActiveForm.Items.Item("ed_Yard").Specific.ChooseFromListUID = "YARD"
            oActiveForm.Items.Item("ed_Yard").Specific.ChooseFromListAlias = "Code"
            oActiveForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
            oActiveForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
            oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
            oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"

            '-------------------------------For Cargo Tab OMM & SYMA------------------------------------------------'13 Jan 2011
            AddChooseFromList(oActiveForm, "cflCurCode", False, 37)
            oEditText = oActiveForm.Items.Item("ed_CurCode").Specific
            oEditText.ChooseFromListUID = "cflCurCode"
            '----------------------------------For Invoice Tab------------------------------------------------------'
            AddChooseFromList(oActiveForm, "cflCurCode1", False, 37)
            oEditText = oActiveForm.Items.Item("ed_CCharge").Specific
            oEditText.ChooseFromListUID = "cflCurCode1"
            AddChooseFromList(oActiveForm, "cflCurCode2", False, 37)
            oEditText = oActiveForm.Items.Item("ed_Charge").Specific
            oEditText.ChooseFromListUID = "cflCurCode2"
            '-------------------------------------------------------------------------------------------------------'

            oCombo = oActiveForm.Items.Item("cb_PCode").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT Code, Name FROM [@OBT_TB004_PORTLIST]")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("Code").Value, oRecordSet.Fields.Item("Name").Value)
                    oRecordSet.MoveNext()
                End While
            End If

            'MSW Container 
            If AddUserDataSrc(oActiveForm, "ConSeqNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ConNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "SealNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddUserDataSrc(oActiveForm, "ConSize", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ConType", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ContWt", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ConDesc", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ConStuff", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ConDate", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "ConDay", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "ConHr", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            oActiveForm.Items.Item("ed_ConNo").Specific.DataBind.SetBound(True, "", "ConSeqNo")
            oActiveForm.Items.Item("ed_ContNo").Specific.DataBind.SetBound(True, "", "ConNo")
            oActiveForm.Items.Item("ed_SealNo").Specific.DataBind.SetBound(True, "", "SealNo")
            oActiveForm.Items.Item("cb_ConSize").Specific.DataBind.SetBound(True, "", "ConSize")
            oActiveForm.Items.Item("cb_ConType").Specific.DataBind.SetBound(True, "", "ConType")
            oActiveForm.Items.Item("ed_ContWt").Specific.DataBind.SetBound(True, "", "ContWt")
            oActiveForm.Items.Item("ed_CDesc").Specific.DataBind.SetBound(True, "", "ConDesc")
            oActiveForm.Items.Item("ch_CStuff").Specific.DataBind.SetBound(True, "", "ConStuff")
            oActiveForm.Items.Item("ed_CunDate").Specific.DataBind.SetBound(True, "", "ConDate")
            oActiveForm.Items.Item("ed_CunDay").Specific.DataBind.SetBound(True, "", "ConDay")
            'oActiveForm.Items.Item("ed_CunTime").Specific.DataBind.SetBound(True, "", "ConHr")

            oCombo = oActiveForm.Items.Item("cb_ConType").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select U_ContType  from [@OBT_TB021_CONT] group by U_ContType")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_ContType").Value, "")
                    oRecordSet.MoveNext()
                End While
            End If

            oCombo = oActiveForm.Items.Item("cb_ConSize").Specific
            oCombo.ValidValues.Add("20'", "20'")
            oCombo.ValidValues.Add("40'", "40'")
            oCombo.ValidValues.Add("45'", "45'")
            'End MSW Container

            'oCombo = oActiveForm.Items.Item("cb_PType").Specific
            'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery("SELECT PkgType FROM OPKG")
            'If oRecordSet.RecordCount > 0 Then
            '    oRecordSet.MoveFirst()
            '    While oRecordSet.EoF = False
            '        oCombo.ValidValues.Add(oRecordSet.Fields.Item("PkgType").Value, "")
            '        oRecordSet.MoveNext()
            '    End While
            'End If

            oEditText = oActiveForm.Items.Item("ed_JobNo").Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEAFCL'")
            If oRecordSet.RecordCount > 0 Then
                'oActiveForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                oActiveForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
            End If

            'fortruckingtab
            If AddUserDataSrc(oActiveForm, "TKRINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKREXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "DSINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "DSEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oOpt = oActiveForm.Items.Item("op_Inter").Specific
            oOpt.DataBind.SetBound(True, "", "DSINTR")
            oOpt = oActiveForm.Items.Item("op_Exter").Specific
            oOpt.DataBind.SetBound(True, "", "DSEXTR")
            oOpt.GroupWith("op_Inter")

            If AddUserDataSrc(oActiveForm, "TKRDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "INSDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_InsDate").Specific.DataBind.SetBound(True, "", "INSDATE")
            oActiveForm.Items.Item("ed_TkrDate").Specific.DataBind.SetBound(True, "", "TKRDATE")
            If AddUserDataSrc(oActiveForm, "TKRATTE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRTEL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRFAX", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRMAIL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRCOL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRTO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
            oActiveForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
            oActiveForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
            oActiveForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
            oActiveForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
            oActiveForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")

            If AddUserDataSrc(oActiveForm, "DSDISP", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("op_DspIntr").Specific.DataBind.SetBound(True, "", "DSDISP")
            oActiveForm.Items.Item("op_DspExtr").Specific.DataBind.SetBound(True, "", "DSDISP")
            oActiveForm.Items.Item("op_DspExtr").Specific.GroupWith("op_DspIntr")

            If AddChooseFromList(oActiveForm, "CFLTKRE", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(oActiveForm, "CFLTKRV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'msw
            If AddUserDataSrc(oActiveForm, "Permit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "Voucher", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "Yard", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "Cont", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "Truck", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "Dispatch", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "Charges", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ch_Permit").Specific.DataBind.SetBound(True, "", "Permit")
            oActiveForm.Items.Item("ch_Voucher").Specific.DataBind.SetBound(True, "", "Voucher")
            oActiveForm.Items.Item("ch_RYard").Specific.DataBind.SetBound(True, "", "Yard")
            oActiveForm.Items.Item("ch_Conta").Specific.DataBind.SetBound(True, "", "Cont")
            oActiveForm.Items.Item("ch_Truck").Specific.DataBind.SetBound(True, "", "Truck")
            oActiveForm.Items.Item("ch_Disp").Specific.DataBind.SetBound(True, "", "Dispatch")
            oActiveForm.Items.Item("ch_OCharge").Specific.DataBind.SetBound(True, "", "Charges")

            'oActiveForm.DataSources.UserDataSources.Item("Permit").Value = "Y"
            'oActiveForm.DataSources.UserDataSources.Item("Voucher").Value = "Y"

            oActiveForm.Items.Item("ch_Permit").Specific.Checked = True
            oActiveForm.Items.Item("ch_Voucher").Specific.Checked = True
            oActiveForm.Items.Item("ch_RYard").Specific.Checked = True
            oActiveForm.Items.Item("ch_Conta").Specific.Checked = True
            oActiveForm.Items.Item("ch_Truck").Specific.Checked = True
            oActiveForm.Items.Item("ch_Disp").Specific.Checked = True
            oActiveForm.Items.Item("ch_OCharge").Specific.Checked = True
            ' ---------------------------10-1-2011-------------------------------------
            '----------Recordset for Binding colCType of Matrix (mx_Cont)-------------
            '--------------------------SYMA & OMM-------------------------------------
            oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
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

            'If AddUserDataSrc(oActiveForm, "Code", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException
            'oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
            'oColum = oMatrix.Columns.Item("colChCode")
            'oColum.DataBind.SetBound(True, "", "Code")
            'oColum.ChooseFromListUID = "CFLTKRE"
            'oColum.ChooseFromListAlias = "Code"
            'oMatrix.LoadFromDataSource()
            oActiveForm.EnableMenu("1283", False) 'MSW 01-04-2011
            oActiveForm.EnableMenu("1292", True)
            oActiveForm.EnableMenu("1293", True)
            oActiveForm.EnableMenu("5907", True)
            '------------------------------For License Info SYMA & OMM (13/Jan/2011)---------------'
            oMatrix = oActiveForm.Items.Item("mx_License").Specific
            oMatrix.AddRow()
            oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
            '-------------------------------------------------------------------------------------'

            oActiveForm.Freeze(False)
            Select Case oActiveForm.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    oActiveForm.Items.Item("bt_AmdIns").Enabled = False
                    oActiveForm.Items.Item("bt_AddIns").Enabled = False
                    oActiveForm.Items.Item("bt_DelIns").Enabled = False
                    oActiveForm.Items.Item("bt_PrntIns").Enabled = False
                Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                    oActiveForm.Items.Item("bt_AddIns").Enabled = True
                    oActiveForm.Items.Item("bt_DelIns").Enabled = True
                    oActiveForm.Items.Item("bt_PrntIns").Enabled = True
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub EnabledDispatchMatrix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)
        Try
            pMatrix.Columns.Item("colDisp").Editable = pValue
            pMatrix.Columns.Item("colDisp").BackColor = 16645629
            pMatrix.Columns.Item("colMode").Editable = pValue
            pMatrix.Columns.Item("colMode").BackColor = 16645629
            pMatrix.Columns.Item("colDate").Editable = pValue
            pMatrix.Columns.Item("colDate").BackColor = 16645629
            pMatrix.Columns.Item("colTime").Editable = pValue
            pMatrix.Columns.Item("colTime").BackColor = 16645629
            pMatrix.Columns.Item("colInst").Editable = pValue
            pMatrix.Columns.Item("colInst").BackColor = 16645629
            pMatrix.Columns.Item("colComp").Editable = pValue
            pMatrix.Columns.Item("colComp").BackColor = 16645629
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub EnabledHeaderControls(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)
        pForm.Items.Item("ed_ShpAgt").Enabled = pValue
        pForm.Items.Item("ed_OBL").Enabled = pValue
        pForm.Items.Item("ed_HBL").Enabled = pValue
        pForm.Items.Item("ed_Vessel").Enabled = pValue
        pForm.Items.Item("ed_Voy").Enabled = pValue
        pForm.Items.Item("cb_PCode").Enabled = pValue
        pForm.Items.Item("ed_ETADate").Enabled = pValue
        pForm.Items.Item("ed_ETAHr").Enabled = pValue
        pForm.Items.Item("ed_ConDate").Enabled = pValue
        pForm.Items.Item("ed_ConTime").Enabled = pValue
        pForm.Items.Item("ed_ConLast").Enabled = pValue
        pForm.Items.Item("ed_ConLTim").Enabled = pValue
        pForm.Items.Item("cb_JobType").Enabled = pValue
        pForm.Items.Item("cb_JbStus").Enabled = pValue
        pForm.Items.Item("ed_Yard").Enabled = pValue
        pForm.Items.Item("ed_JobNo").Enabled = pValue
    End Sub

    Private Sub ValidateJobNumber(ByRef oActiveForm As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)
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

        Dim argument As String = "/f" & Chr(34) & docfolderpath & "\"
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
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#Region "---------- 'MSW - Purchase Voucher Save To Draft 09-03-2010"
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
                    MsgBox("Failed to update a payment")
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
                        'Document.Lines.VatGroup = dtmatrix.GetValue("GST", i - 1) 'oMatrix.Columns.Item("colGST").Cells.Item(i).Specific.Value
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
        If HolidayMarkUp(oActiveForm, oPayForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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

#Region "MSW Voucher POP Up"
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

    Private Sub DeleteMatrixRow(ByRef oActiveForm As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix, ByVal objDataSource As String, ByVal oColumn As String)
        Dim tblname As String = objDataSource.Substring(1, objDataSource.Length)
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

#End Region

#Region "SHIPPING INVOICE POP UP"
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
        ' oShpForm.DataSources.DBDataSources.Item("@OBT_TB017_EVHEADER").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)

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
#End Region

#Region "SHIPPING INV" 'syma
    Private Sub AddUpdateShippingInv(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        rowIndex = pMatrix.GetNextSelectedRow
        If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        Try
            If ProcressedState = True Then

                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                        .SetValue("U_SerNo", .Offset, 1)
                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                        .SetValue("U_SerNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                    End If

                    .SetValue("U_ExInv", .Offset, pForm.Items.Item("ed_ExInv").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value)
                    .SetValue("U_PartNo", .Offset, pForm.Items.Item("ed_Part").Specific.Value)
                    .SetValue("U_Desc", .Offset, pForm.Items.Item("ed_PartDes").Specific.Value)
                    .SetValue("U_Qty", .Offset, pForm.Items.Item("ed_Qty").Specific.Value)
                    .SetValue("U_UM", .Offset, pForm.Items.Item("ed_Unit").Specific.Value)
                    .SetValue("U_Box", .Offset, pForm.Items.Item("ed_Box").Specific.Value)
                    .SetValue("U_Length", .Offset, pForm.Items.Item("ed_L").Specific.Value)
                    .SetValue("U_Base", .Offset, pForm.Items.Item("ed_B").Specific.Value)
                    .SetValue("U_Height", .Offset, pForm.Items.Item("ed_H").Specific.Value)
                    .SetValue("U_M3", .Offset, pForm.Items.Item("ed_M3").Specific.Value)
                    .SetValue("U_TotM3", .Offset, pForm.Items.Item("ed_M3T").Specific.Value)
                    .SetValue("U_NetWt", .Offset, pForm.Items.Item("ed_Net").Specific.Value)
                    .SetValue("U_TotNet", .Offset, pForm.Items.Item("ed_NetT").Specific.Value)
                    .SetValue("U_GrWt", .Offset, pForm.Items.Item("ed_Gross").Specific.Value)
                    .SetValue("U_TGrWt", .Offset, pForm.Items.Item("ed_GrossT").Specific.Value)
                    .SetValue("U_NEC", .Offset, pForm.Items.Item("ed_Nec").Specific.Value)
                    .SetValue("U_TotNEC", .Offset, pForm.Items.Item("ed_NecT").Specific.Value)
                    .SetValue("U_TNBox", .Offset, pForm.Items.Item("ed_TotBox").Specific.Value)
                    .SetValue("U_TNBoxUN", .Offset, pForm.Items.Item("ed_Boxes").Specific.Value)
                    .SetValue("U_PPBox", .Offset, pForm.Items.Item("ed_PPBNo").Specific.Value)
                    .SetValue("U_PPBoxUN", .Offset, pForm.Items.Item("ed_PUnit").Specific.Value)
                    .SetValue("U_Rate", .Offset, pForm.Items.Item("ed_UnPrice").Specific.Value)
                    .SetValue("U_TValue", .Offset, pForm.Items.Item("ed_TotVal").Specific.Value)
                    .SetValue("U_Shipping", .Offset, pForm.Items.Item("ed_ShName").Specific.Value)
                    .SetValue("U_PPSName", .Offset, pForm.Items.Item("ed_PShName").Specific.Value)
                    .SetValue("U_ECCN", .Offset, pForm.Items.Item("ed_ECCN").Specific.Value)
                    .SetValue("U_License", .Offset, pForm.Items.Item("ed_License").Specific.Value)
                    .SetValue("U_LicDate", .Offset, pForm.Items.Item("ed_LExDate").Specific.Value)
                    .SetValue("U_Class", .Offset, pForm.Items.Item("ed_Class").Specific.Value)
                    .SetValue("U_UNNo", .Offset, pForm.Items.Item("ed_UN").Specific.Value)
                    .SetValue("U_HSCode", .Offset, pForm.Items.Item("ed_HSCode").Specific.Value)
                    .SetValue("U_DOM", .Offset, pForm.Items.Item("ed_DOM").Specific.Value)

                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource

                    .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.GetNextSelectedRow).Specific.Value)
                    .SetValue("U_SerNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.GetNextSelectedRow).Specific.Value)
                    .SetValue("U_ExInv", .Offset, pForm.Items.Item("ed_ExInv").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value)
                    .SetValue("U_PartNo", .Offset, pForm.Items.Item("ed_Part").Specific.Value)
                    .SetValue("U_Desc", .Offset, pForm.Items.Item("ed_PartDes").Specific.Value)
                    .SetValue("U_Qty", .Offset, pForm.Items.Item("ed_Qty").Specific.Value)
                    .SetValue("U_UM", .Offset, pForm.Items.Item("ed_Unit").Specific.Value)
                    .SetValue("U_Box", .Offset, pForm.Items.Item("ed_Box").Specific.Value)
                    .SetValue("U_Length", .Offset, pForm.Items.Item("ed_L").Specific.Value)
                    .SetValue("U_Base", .Offset, pForm.Items.Item("ed_B").Specific.Value)
                    .SetValue("U_Height", .Offset, pForm.Items.Item("ed_H").Specific.Value)
                    .SetValue("U_M3", .Offset, pForm.Items.Item("ed_M3").Specific.Value)
                    .SetValue("U_TotM3", .Offset, pForm.Items.Item("ed_M3T").Specific.Value)
                    .SetValue("U_NetWt", .Offset, pForm.Items.Item("ed_Net").Specific.Value)
                    .SetValue("U_TotNet", .Offset, pForm.Items.Item("ed_NetT").Specific.Value)
                    .SetValue("U_GrWt", .Offset, pForm.Items.Item("ed_Gross").Specific.Value)
                    .SetValue("U_TGrWt", .Offset, pForm.Items.Item("ed_GrossT").Specific.Value)
                    .SetValue("U_NEC", .Offset, pForm.Items.Item("ed_Nec").Specific.Value)
                    .SetValue("U_TotNEC", .Offset, pForm.Items.Item("ed_NecT").Specific.Value)
                    .SetValue("U_TNBox", .Offset, pForm.Items.Item("ed_TotBox").Specific.Value)
                    .SetValue("U_TNBoxUN", .Offset, pForm.Items.Item("ed_Boxes").Specific.Value)
                    .SetValue("U_PPBox", .Offset, pForm.Items.Item("ed_PPBNo").Specific.Value)
                    .SetValue("U_PPBoxUN", .Offset, pForm.Items.Item("ed_PUnit").Specific.Value)
                    .SetValue("U_Rate", .Offset, pForm.Items.Item("ed_UnPrice").Specific.Value)
                    .SetValue("U_TValue", .Offset, pForm.Items.Item("ed_TotVal").Specific.Value)
                    .SetValue("U_Shipping", .Offset, pForm.Items.Item("ed_ShName").Specific.Value)
                    .SetValue("U_PPSName", .Offset, pForm.Items.Item("ed_PShName").Specific.Value)
                    .SetValue("U_ECCN", .Offset, pForm.Items.Item("ed_ECCN").Specific.Value)
                    .SetValue("U_License", .Offset, pForm.Items.Item("ed_License").Specific.Value)
                    .SetValue("U_LicDate", .Offset, pForm.Items.Item("ed_LExDate").Specific.Value)
                    .SetValue("U_Class", .Offset, pForm.Items.Item("ed_Class").Specific.Value)
                    .SetValue("U_UNNo", .Offset, pForm.Items.Item("ed_UN").Specific.Value)
                    .SetValue("U_HSCode", .Offset, pForm.Items.Item("ed_HSCode").Specific.Value)
                    .SetValue("U_DOM", .Offset, pForm.Items.Item("ed_DOM").Specific.Value)

                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetShipInvDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Dim sErrDesc As String = String.Empty

        Try
            pForm.Freeze(True)

            'pForm.Items.Item("ed_ConNo").Specific.Value = pMatrix.Columns.Item("V_-1").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_ExInv").Specific.Value = pMatrix.Columns.Item("colExInv").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PO").Specific.Value = pMatrix.Columns.Item("colPO").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Part").Specific.Value = pMatrix.Columns.Item("colPart").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PartDes").Specific.Value = pMatrix.Columns.Item("colPartDes").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Qty").Specific.Value = pMatrix.Columns.Item("colQty").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Unit").Specific.Value = pMatrix.Columns.Item("colQtyUnit").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Box").Specific.Value = pMatrix.Columns.Item("colBox").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_L").Specific.Value = pMatrix.Columns.Item("colDLength").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_B").Specific.Value = pMatrix.Columns.Item("colDBase").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_H").Specific.Value = pMatrix.Columns.Item("colDHeight").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_M3").Specific.Value = pMatrix.Columns.Item("colM3").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_M3T").Specific.Value = pMatrix.Columns.Item("colTotM3").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Net").Specific.Value = pMatrix.Columns.Item("colNWeight").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_NetT").Specific.Value = pMatrix.Columns.Item("colTotNet").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Gross").Specific.Value = pMatrix.Columns.Item("colGrWt").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_GrossT").Specific.Value = pMatrix.Columns.Item("colTGrWt").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Nec").Specific.Value = pMatrix.Columns.Item("colNEC").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_NecT").Specific.Value = pMatrix.Columns.Item("colTotNEC").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_TotBox").Specific.Value = pMatrix.Columns.Item("colTNBox").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Boxes").Specific.Value = pMatrix.Columns.Item("colTNBpxUN").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PPBNo").Specific.Value = pMatrix.Columns.Item("colPPBox").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PUnit").Specific.Value = pMatrix.Columns.Item("colPPBoxUN").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_UnPrice").Specific.Value = pMatrix.Columns.Item("colUPrice").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_TotVal").Specific.Value = pMatrix.Columns.Item("colTValue").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_ShName").Specific.Value = pMatrix.Columns.Item("colSName").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PShName").Specific.Value = pMatrix.Columns.Item("colPSName").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_ECCN").Specific.Value = pMatrix.Columns.Item("colECCN").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_License").Specific.Value = pMatrix.Columns.Item("colLicense").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_LExDate").Specific.Value = pMatrix.Columns.Item("colLicDate").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Class").Specific.Value = pMatrix.Columns.Item("colClass").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_UN").Specific.Value = pMatrix.Columns.Item("colUNNo").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_HSCode").Specific.Value = pMatrix.Columns.Item("colHSCode").Cells.Item(Index).Specific.Value
            ' If HolidaysMarkUp(pForm, pForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, pForm.Items.Item("ed_CunDay").Specific, pForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            pForm.Freeze(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub AddUpdateShippingMatrix(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
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
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Function ValidateforformShippingInv(ByVal oActiveForm As SAPbouiCOM.Form) As Boolean
        If String.IsNullOrEmpty(oActiveForm.Items.Item("ed_ExInv").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Enter Ex-Invoice ", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf String.IsNullOrEmpty(oActiveForm.Items.Item("ed_PO").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Enter PO", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf String.IsNullOrEmpty(oActiveForm.Items.Item("ed_Part").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Part No", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "PO and GR"
    Private Sub LoadAndCreateCPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)
        Dim CPOForm As SAPbouiCOM.Form
        Dim CPOMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Try
            If Not LoadFromXML(p_oSBOApplication, srfName) Then Throw New ArgumentException(sErrDesc)
            CPOForm = p_oSBOApplication.Forms.ActiveForm
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                CPOForm.EnableMenu("1292", True)
                CPOForm.EnableMenu("1293", True)
                CPOForm.EnableMenu("1294", True)

                CPOForm.AutoManaged = True
                CPOForm.DataBrowser.BrowseBy = "ed_CPOID"
                CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                If AddChooseFromList(CPOForm, "CFLCODES", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddChooseFromList(CPOForm, "CFLNAMES", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                BindingChooseFromList(CPOForm, "CFLCODES", "ed_Code", "CardCode")
                BindingChooseFromList(CPOForm, "CFLNAMES", "ed_Name", "CardName")

                oCombo = CPOForm.Items.Item("cb_SInA").Specific
                oRecordSet.DoQuery("SELECT empID,lastName,firstName,middleName FROM OHEM")
                If oRecordSet.RecordCount > 0 Then
                    oRecordSet.MoveFirst()
                    While oRecordSet.EoF = False
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item("empID").Value, oRecordSet.Fields.Item("firstName").Value & " " _
                                                                               & oRecordSet.Fields.Item("middleName").Value & " " _
                                                                               & oRecordSet.Fields.Item("lastName").Value)
                        oRecordSet.MoveNext()
                    End While
                End If

                CPOForm.Items.Item("ed_CPOID").Specific.Value = GetNewKey("FCPO", oRecordSet)
                CPOForm.Items.Item("ed_PONo").Specific.Value = GetNewKey("22", oRecordSet)
                CPOForm.Items.Item("ed_Status").Specific.Value = "Open"
                CPOForm.Items.Item("ed_PODate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                CPOForm.Items.Item("ed_PODay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                CPOForm.Items.Item("ed_POTime").Specific.Value = Now.ToString("HH:mm")

                If AddUserDataSrc(CPOForm, "Email", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                CPOForm.Items.Item("ch_Email").Specific.DataBind.SetBound(True, "", "Email")

                If AddUserDataSrc(CPOForm, "Fax", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                CPOForm.Items.Item("ch_Fax").Specific.DataBind.SetBound(True, "", "Fax")

                If AddUserDataSrc(CPOForm, "Print", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                CPOForm.Items.Item("ch_Print").Specific.DataBind.SetBound(True, "", "Print")

                If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_DocNum").Specific.Value)

                CPOForm.Items.Item("bt_Preview").Visible = False
                CPOForm.Items.Item("bt_Resend").Visible = False
                CPOForm.Items.Item("ed_Code").Specific.Active = True
                ' ==================================== Custom Purchase Order ========================================
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadAndCreateCGR(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)
        Dim CGRForm As SAPbouiCOM.Form
        Dim CGRMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Try
            If Not LoadFromXML(p_oSBOApplication, srfName) Then Throw New ArgumentException(sErrDesc)
            CGRForm = p_oSBOApplication.Forms.ActiveForm
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                CGRForm.EnableMenu("1292", True)
                CGRForm.EnableMenu("1293", True)
                CGRForm.EnableMenu("1294", True)
                CGRForm.AutoManaged = True
                CGRForm.DataBrowser.BrowseBy = "ed_CGRID"
                CGRMatrix = CGRForm.Items.Item("mx_Item").Specific
                If AddChooseFromList(CGRForm, "CFLCODEGR", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddChooseFromList(CGRForm, "CFLNAMEGR", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                BindingChooseFromList(CGRForm, "CFLCODEGR", "ed_Code", "CardCode")
                BindingChooseFromList(CGRForm, "CFLNAMEGR", "ed_Name", "CardName")

                oCombo = CGRForm.Items.Item("cb_SInA").Specific
                oRecordSet.DoQuery("SELECT empID,lastName,firstName,middleName FROM OHEM")
                If oRecordSet.RecordCount > 0 Then
                    oRecordSet.MoveFirst()
                    While oRecordSet.EoF = False
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item("empID").Value, oRecordSet.Fields.Item("firstName").Value & " " _
                                                                               & oRecordSet.Fields.Item("middleName").Value & " " _
                                                                               & oRecordSet.Fields.Item("lastName").Value)
                        oRecordSet.MoveNext()
                    End While
                End If
            End If

            CGRForm.Items.Item("ed_CGRID").Specific.Value = GetNewKey("FCGR", oRecordSet)
            CGRForm.Items.Item("ed_GRNo").Specific.Value = GetNewKey("20", oRecordSet)
            CGRForm.Items.Item("ed_Status").Specific.Value = "Open"
            CGRForm.Items.Item("ed_GRDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
            CGRForm.Items.Item("ed_GRDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
            CGRForm.Items.Item("ed_GRTime").Specific.Value = Now.ToString("HH:mm")
            AddUserDataSrc(CGRForm, "PONo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 11)

            'If AddUserDataSrc(CGRForm, "Email", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'CGRForm.Items.Item("39").Specific.DataBind.SetBound(True, "", "Email")

            If Not AddNewRowGR(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
            CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_DocNum").Specific.Value)
            CGRForm.Items.Item("ed_Code").Specific.Active = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function FillDataToGoodsReceipt(ByRef oSourceForm As SAPbouiCOM.Form, ByVal SourceMatrixName As String, _
                                            ByVal SourceColName1 As String, ByVal SourceColName2 As String, ByVal ActiveRow As Integer, _
                                            ByVal DestDataSource As String, ByRef oDestForm As SAPbouiCOM.Form) As Boolean
        FillDataToGoodsReceipt = False
        Dim oMatrix, oDestMatrix As SAPbouiCOM.Matrix
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim oPODocument As SAPbobsCOM.Documents
        Dim oDBDataSource, oLineDBDataSource As SAPbouiCOM.DBDataSource
        Dim sErrDesc As String = vbNullString
        Dim Total As Long
        Try
            oMatrix = oSourceForm.Items.Item(SourceMatrixName).Specific
            oDestMatrix = oDestForm.Items.Item("mx_Item").Specific
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oPODocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oDBDataSource = oDestForm.DataSources.DBDataSources.Item(DestDataSource)
            oLineDBDataSource = oDestForm.DataSources.DBDataSources.Item("@OBT_TB13_FFCGRITEM")
            If oPODocument.GetByKey(oMatrix.Columns.Item(SourceColName1).Cells.Item(ActiveRow).Specific.Value) Then
                oDestForm.DataSources.UserDataSources.Item("PONo").ValueEx = oMatrix.Columns.Item(SourceColName1).Cells.Item(ActiveRow).Specific.Value
                With oDBDataSource
                    .SetValue("U_Code", .Offset, oPODocument.CardCode)
                    .SetValue("U_Name", .Offset, oPODocument.CardName)
                    .SetValue("U_CPerson", .Offset, GetContactPersonName(oRecordset, oPODocument.ContactPersonCode, oPODocument.CardCode))
                    Dim tsql As String = "SELECT U_SInA,U_TPlace,U_TDate,U_TDay,U_TTime,U_PORMKS,U_POIRMKS,U_CNo,U_Dest,U_LocWork,U_POITPD FROM [@OBT_TB08_FFCPO] WHERE DocEntry = " + FormatString(oMatrix.Columns.Item(SourceColName2).Cells.Item(ActiveRow).Specific.Value)
                    oRecordset.DoQuery(tsql)
                    If oRecordset.RecordCount > 0 Then
                        oRecordset.MoveFirst()
                        While oRecordset.EoF = False
                            .SetValue("U_SInA", .Offset, oRecordset.Fields.Item("U_SInA").Value)
                            .SetValue("U_TDate", .Offset, CDate(oRecordset.Fields.Item("U_TDate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_TDay", .Offset, oRecordset.Fields.Item("U_TDay").Value)
                            .SetValue("U_TTime", .Offset, oRecordset.Fields.Item("U_TTime").Value.ToString)
                            .SetValue("U_TPlace", .Offset, oRecordset.Fields.Item("U_TPlace").Value)
                            .SetValue("U_GRRMKS", .Offset, oRecordset.Fields.Item("U_PORMKS").Value)
                            .SetValue("U_GRIRMKS", .Offset, oRecordset.Fields.Item("U_POIRMKS").Value)
                            .SetValue("U_GRTPD", .Offset, oRecordset.Fields.Item("U_POITPD").Value)
                            .SetValue("U_Dest", .Offset, oRecordset.Fields.Item("U_Dest").Value)
                            .SetValue("U_LocWork", .Offset, oRecordset.Fields.Item("U_LocWork").Value)
                            .SetValue("U_CNo", .Offset, oRecordset.Fields.Item("U_CNo").Value)
                            oRecordset.MoveNext()
                        End While
                    End If
                End With
            End If
            Dim tempSQL As String = "SELECT * FROM POR1 WHERE DocEntry =" + FormatString(oPODocument.DocEntry)
            oRecordset.DoQuery(tempSQL)
            oDestMatrix.Clear()

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()
                While oRecordset.EoF = False
                    With oLineDBDataSource
                        .SetValue("LineId", .Offset, oDestMatrix.VisualRowCount + 1)
                        .SetValue("U_GRINO", .Offset, oRecordset.Fields.Item("ItemCode").Value)
                        .SetValue("U_GRIDesc", .Offset, oRecordset.Fields.Item("Dscription").Value)
                        .SetValue("U_GRIQty", .Offset, oRecordset.Fields.Item("Quantity").Value)
                        .SetValue("U_GRIPrice", .Offset, oRecordset.Fields.Item("Price").Value)
                        .SetValue("U_GRIAmt", .Offset, oRecordset.Fields.Item("OpenSum").Value)
                        .SetValue("U_GRIGST", .Offset, oRecordset.Fields.Item("VatGroup").Value)
                        .SetValue("U_GRITot", .Offset, oRecordset.Fields.Item("OpenSum").Value)
                    End With
                    oDestMatrix.AddRow()
                    oRecordset.MoveNext()
                End While
            End If
            FillDataToGoodsReceipt = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            FillDataToGoodsReceipt = False
        End Try
    End Function

    Private Function CreateGoodsReceiptPO(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        CreateGoodsReceiptPO = False
        Dim oPurchase As SAPbobsCOM.Documents
        Dim oPurchaseDeliveryNote As SAPbobsCOM.Documents
        Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sErrDesc As String = vbNullString
        Try
            oPurchase = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oPurchaseDeliveryNote = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
            oBusinessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oPurchase.GetByKey(oActiveForm.DataSources.UserDataSources.Item("PONo").Value) Then
                p_oDICompany.GetNewObjectCode("20")
                If oBusinessPartner.GetByKey(oActiveForm.Items.Item("ed_Code").Specific.Value) Then
                    oPurchaseDeliveryNote.CardCode = oActiveForm.Items.Item("ed_Code").Specific.Value
                    oPurchaseDeliveryNote.CardName = oActiveForm.Items.Item("ed_Name").Specific.Value
                    oPurchaseDeliveryNote.ContactPersonCode = GetContactPersonCode(oRecordset, Trim(oActiveForm.Items.Item("cb_Contact").Specific.Value.ToString), oActiveForm.Items.Item("ed_Code").Specific.Value.ToString)
                    oPurchaseDeliveryNote.NumAtCard = oActiveForm.Items.Item("ed_VRef").Specific.Value
                    oPurchaseDeliveryNote.DocCurrency = oBusinessPartner.Currency
                    oPurchaseDeliveryNote.DocDate = Now
                    oPurchaseDeliveryNote.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    oPurchaseDeliveryNote.TaxDate = Now
                    oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            oPurchaseDeliveryNote.Lines.BaseType = CInt(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            oPurchaseDeliveryNote.Lines.BaseEntry = Convert.ToInt32(oActiveForm.DataSources.UserDataSources.Item("PONo").Value)
                            oPurchaseDeliveryNote.Lines.BaseLine = i - 1
                            oPurchaseDeliveryNote.Lines.ItemCode = oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value
                            oPurchaseDeliveryNote.Lines.ItemDescription = oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value
                            oPurchaseDeliveryNote.Lines.Quantity = oMatrix.Columns.Item("colIQty").Cells.Item(i).Specific.Value
                            oPurchaseDeliveryNote.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colIPrice").Cells.Item(i).Specific.Value)
                            oPurchaseDeliveryNote.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
                            'oPurchaseDeliveryNote.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                            If oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value = "None" Then
                                oPurchaseDeliveryNote.Lines.VatGroup = "ZI"
                            Else
                                oPurchaseDeliveryNote.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                            End If
                            oPurchaseDeliveryNote.Lines.Add()
                        Next
                    End If
                End If
                Dim ret As Long = oPurchaseDeliveryNote.Add
                If ret <> 0 Then
                    p_oDICompany.GetLastError(ret, sErrDesc)
                    MessageBox.Show("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
                End If
            End If
            CreateGoodsReceiptPO = True
        Catch ex As Exception
            CreateGoodsReceiptPO = False
        End Try
    End Function

    Private Sub CalAmtPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim cMatrix As SAPbouiCOM.Matrix
        Dim itemTotal As Double
        cMatrix = oActiveForm.Items.Item("mx_Item").Specific
        cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(cMatrix.Columns.Item("colIPrice").Cells.Item(Row).Specific.Value)
        cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value)
        itemTotal += cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value
        ' oActiveForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_POITPD", 0, itemTotal)
        CalculateTotalPO(oActiveForm, cMatrix)
    End Sub

    Private Sub CalRatePO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim cMatrix As SAPbouiCOM.Matrix
        'Dim oRecordset As SAPbobsCOM.Recordset
        'Dim oEditText As SAPbouiCOM.EditText
        'Dim itemTotal As Double
        'Dim totalNoGST As Double
        cMatrix = oActiveForm.Items.Item("mx_Item").Specific
        CalculateTotalPO(oActiveForm, cMatrix)
        'oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oRecordset.DoQuery("select Rate from ovtg where Code='" + cMatrix.Columns.Item("colIGST").Cells.Item(Row).Specific.Value + "'")
        'Dim Rate As Double = 0.0
        'Dim GSTAMT As Double = 0.0
        'Dim total As Double = 0.0
        'If oRecordset.RecordCount < 0 Then
        '    Rate = 0
        'Else
        '    oRecordset.MoveFirst()
        '    Rate = oRecordset.Fields.Item("Rate").Value
        '    totalNoGST = Convert.ToDouble(cMatrix.Columns.Item("colIQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(cMatrix.Columns.Item("colIPrice").Cells.Item(Row).Specific.Value)
        '    GSTAMT = (Convert.ToDouble(cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value) / 100.0) * Rate
        '    total = totalNoGST + GSTAMT
        '    'oEditText = cMatrix.Columns.Item("colITotal").Cells.Item(Row)
        '    ' cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.String = total
        '    'cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Editable = False
        '    'TODO need to show total value 
        '    itemTotal += cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value
        '    CalculateTotalPO(oActiveForm, cMatrix)
        '    'oActiveForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_POITPD", 0, itemTotal)
        'End If
    End Sub

    Private Sub CalculateTotalPO(ByRef oActiveForm As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix)

        Dim SubTotal As Double = 0.0
        Dim GSTAmt As Double = 0.0
        Dim Total As Double = 0.0
        For i As Integer = 1 To oMatrix.RowCount
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select Rate from ovtg where Code='" + oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value + "'")
            Dim Rate As Double = 0.0
            If oRecordSet.RecordCount < 0 Then
                Rate = 0
            Else
                oRecordSet.MoveFirst()
                Rate = oRecordSet.Fields.Item("Rate").Value
                GSTAmt = (Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value) / 100.0) * Rate
                SubTotal = SubTotal + GSTAmt + Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
            End If
        Next
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_POITPD", 0, SubTotal)
    End Sub

    Private Function PopulatePurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean) As Boolean
        PopulatePurchaseHeader = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Try
            ObjDbDataSource = pForm.DataSources.DBDataSources.Item(tblName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNo").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                        pMatrix.Clear()
                    End If
                End If
                Do Until oRecordSet.EoF
                    If ProcessedState = True Then
                        With ObjDbDataSource
                            If pMatrix.RowCount = 0 Then
                                .SetValue("LineId", .Offset, 1)
                            Else
                                .SetValue("LineId", .Offset, pMatrix.VisualRowCount + 1)
                            End If
                            .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                            .SetValue("U_PONo", .Offset, DocLastKey)
                            '.SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                            .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_Vendor", .Offset, oRecordSet.Fields.Item("U_VCode").Value.ToString)
                            .SetValue("U_Place", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            .SetValue("U_FDate", .Offset, pForm.Items.Item("ed_JbDate").Specific.Value)
                            .SetValue("U_FTime", .Offset, pForm.Items.Item("ed_JbHr").Specific.Value)
                            .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                            .SetValue("U_Remarks", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                        End With
                        pMatrix.AddRow()
                    Else
                        With ObjDbDataSource
                            .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(currentRow).Specific.Value)
                            .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                            '.SetValue("U_PONo", .Offset, DocLastKey)
                            .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                            .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_Vendor", .Offset, oRecordSet.Fields.Item("U_VCode").Value.ToString)
                            .SetValue("U_Place", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            .SetValue("U_FDate", .Offset, pForm.Items.Item("ed_JbDate").Specific.Value)
                            .SetValue("U_FTime", .Offset, pForm.Items.Item("ed_JbHr").Specific.Value)
                            .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                            .SetValue("U_Remarks", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                        End With
                        pMatrix.SetLineData(currentRow)
                    End If

                    oRecordSet.MoveNext()
                Loop
            End If
            PopulatePurchaseHeader = True
        Catch ex As Exception
            PopulatePurchaseHeader = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function UpdatePurchaseOrder(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        UpdatePurchaseOrder = False
        Dim oPurchaseDocument As SAPbobsCOM.Documents
        Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sErrDesc As String = vbNullString
        Try
            oPurchaseDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oBusinessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oPurchaseDocument.GetByKey(oActiveForm.Items.Item("ed_PONo").Specific.Value.ToString) Then
                If oBusinessPartner.GetByKey(oActiveForm.Items.Item("ed_Code").Specific.Value.ToString) Then
                    'Debug.Print(oPurchaseDocument.CardCode + "/" + oPurchaseDocument.CardName + "/" + oPurchaseDocument.DocCurrency + "/" + oPurchaseDocument.DocDate)
                    oPurchaseDocument.ContactPersonCode = GetContactPersonCode(oRecordset, Trim(oActiveForm.Items.Item("cb_Contact").Specific.Value.ToString), oActiveForm.Items.Item("ed_Code").Specific.Value.ToString)
                    oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                    oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
                    For i As Integer = 0 To oPurchaseDocument.Lines.Count
                        oPurchaseDocument.Lines.Delete()
                    Next
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            oPurchaseDocument.Lines.ItemCode = oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value
                            oPurchaseDocument.Lines.ItemDescription = oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value
                            oPurchaseDocument.Lines.Quantity = oMatrix.Columns.Item("colIQty").Cells.Item(i).Specific.Value
                            oPurchaseDocument.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colIPrice").Cells.Item(i).Specific.Value)
                            oPurchaseDocument.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
                            oPurchaseDocument.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                            oPurchaseDocument.Lines.Add()
                        Next
                    End If

                End If
                Dim ret As Long = oPurchaseDocument.Update
                If ret <> 0 Then
                    p_oDICompany.GetLastError(ret, sErrDesc)
                    Debug.Print("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
                Else
                    p_oSBOApplication.MessageBox("Purchase Order document is successfully update!")
                End If
            End If
            UpdatePurchaseOrder = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            UpdatePurchaseOrder = False
        End Try
    End Function

    Private Function CreatePurchaseOrder(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        '=============================================================================
        'Function   : CreatePurchaseOrder
        'Purpose    : 
        'Parameters : 
        'Return     : 
        'Author     : @channyeinkyaw
        '=============================================================================
        Dim oPurchaseDocument As SAPbobsCOM.Documents
        Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sErrDesc As String = vbNullString
        CreatePurchaseOrder = False
        Try
            oPurchaseDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oBusinessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            p_oDICompany.GetNewObjectCode("22")
            If oBusinessPartner.GetByKey(oActiveForm.Items.Item("ed_Code").Specific.Value.ToString) Then
                oPurchaseDocument.CardCode = oActiveForm.Items.Item("ed_Code").Specific.Value.ToString
                oPurchaseDocument.CardName = oActiveForm.Items.Item("ed_Name").Specific.Value.ToString
                oPurchaseDocument.ContactPersonCode = GetContactPersonCode(oRecordset, Trim(oActiveForm.Items.Item("cb_Contact").Specific.Value.ToString), oActiveForm.Items.Item("ed_Code").Specific.Value.ToString)
                oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                oPurchaseDocument.DocDate = Now
                oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                oPurchaseDocument.TaxDate = Now
                oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
                If oMatrix.RowCount > 0 Then
                    For i As Integer = 1 To oMatrix.RowCount
                        oPurchaseDocument.Lines.ItemCode = oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value
                        oPurchaseDocument.Lines.ItemDescription = oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value
                        oPurchaseDocument.Lines.Quantity = oMatrix.Columns.Item("colIQty").Cells.Item(i).Specific.Value
                        oPurchaseDocument.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colIPrice").Cells.Item(i).Specific.Value)
                        'oPurchaseDocument.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
                        '   oPurchaseDocument.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                        If oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value = "None" Then
                            oPurchaseDocument.Lines.VatGroup = "ZI"
                        Else
                            oPurchaseDocument.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                        End If
                        oPurchaseDocument.Lines.Add()
                    Next
                End If
            End If
            Dim ret As Long = oPurchaseDocument.Add
            If ret <> 0 Then
                p_oDICompany.GetLastError(ret, sErrDesc)
                MessageBox.Show("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
            End If
            oRecordset.DoQuery("SELECT DocEntry FROM OPOR Order By DocEntry")
            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveLast()
                DocLastKey = oRecordset.Fields.Item("DocEntry").Value
            End If
            CreatePurchaseOrder = True
        Catch ex As Exception
            CreatePurchaseOrder = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function GetContactPersonCode(ByVal oRecordSet As SAPbobsCOM.Recordset, ByVal Name As String, ByVal CardCode As String) As Integer
        '=============================================================================
        'Function   : GetContactPersonCode()
        'Purpose    : This function get Contact Person Code from ComboBox
        'Parameters : ByVal oComboBox As SAPbouiCOM.ComboBox
        '               oComboBox = 
        'Return     : Contact Person Code integer
        'Author     : @channyeinkyaw
        '=============================================================================
        'TODO 2 b continue
        GetContactPersonCode = 0
        Try
            oRecordSet.DoQuery("SELECT CntctCode FROM OCPR WHERE Name = " & FormatString(Name) & " AND CardCode = " & FormatString(CardCode))
            If oRecordSet.RecordCount > 0 Then
                GetContactPersonCode = oRecordSet.Fields.Item("CntctCode").Value.ToString
            End If
        Catch ex As Exception
            GetContactPersonCode = 0
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function GetContactPersonName(ByVal oRecordSet As SAPbobsCOM.Recordset, ByVal ContactPersonCode As String, ByVal CardCode As String) As String
        '=============================================================================
        'Function   : GetContactPersonCode()
        'Purpose    : This function get Contact Person Code from ComboBox
        'Parameters : ByVal oComboBox As SAPbouiCOM.ComboBox
        '               oComboBox = 
        'Return     : Contact Person Code integer
        'Author     : @channyeinkyaw
        '=============================================================================
        'TODO 2 b continue
        GetContactPersonName = vbNullString
        Try
            oRecordSet.DoQuery("SELECT Name FROM OCPR WHERE CntctCode = " & FormatString(ContactPersonCode) & " AND CardCode = " & FormatString(CardCode))
            If oRecordSet.RecordCount > 0 Then
                GetContactPersonName = oRecordSet.Fields.Item("Name").Value.ToString
            End If
        Catch ex As Exception
            GetContactPersonName = vbNullString
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function AddFirstRow(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        AddFirstRow = False
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
            oMatrix.AddRow()
            AddGSTComboData(oMatrix.Columns.Item("colIGST"))
            oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.RowCount).Specific.Value = CStr(oMatrix.RowCount)
            AddFirstRow = True
        Catch ex As Exception
            AddFirstRow = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function AddNewRowPO(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        AddNewRowPO = False
        Dim sErrDesc As String = vbNullString
        Dim oDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
            oDbDataSource = oActiveForm.DataSources.DBDataSources.Item("@OBT_TB09_FFCPOITEM")
            With oDbDataSource
                .SetValue("LineId", .Offset, oMatrix.RowCount + 1)
                .SetValue("U_POINO", .Offset, String.Empty)
                .SetValue("U_POIDesc", .Offset, String.Empty)
                .SetValue("U_POIQty", .Offset, String.Empty)
                .SetValue("U_POIPrice", .Offset, String.Empty)
                .SetValue("U_POIAmt", .Offset, String.Empty)
                .SetValue("U_POIGST", .Offset, String.Empty)
                .SetValue("U_POITot", .Offset, String.Empty)
            End With
            oMatrix.AddRow()
            AddGSTComboData(oMatrix.Columns.Item("colIGST"))
            AddNewRowPO = True
        Catch ex As Exception
            AddNewRowPO = False
        End Try
    End Function

    Private Function AddNewRowGR(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        AddNewRowGR = False
        Dim sErrDesc As String = vbNullString
        Dim oDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
            oDbDataSource = oActiveForm.DataSources.DBDataSources.Item("@OBT_TB13_FFCGRITEM")
            With oDbDataSource
                .SetValue("LineId", .Offset, oMatrix.RowCount + 1)
                .SetValue("U_GRINO", .Offset, String.Empty)
                .SetValue("U_GRIDesc", .Offset, String.Empty)
                .SetValue("U_GRIQty", .Offset, String.Empty)
                .SetValue("U_GRIPrice", .Offset, String.Empty)
                .SetValue("U_GRIAmt", .Offset, String.Empty)
                .SetValue("U_GRIGST", .Offset, String.Empty)
                .SetValue("U_GRITot", .Offset, String.Empty)
            End With
            oMatrix.AddRow()
            AddGSTComboData(oMatrix.Columns.Item("colIGST"))
            AddNewRowGR = True
        Catch ex As Exception
            AddNewRowGR = False
        End Try
    End Function

    Private Sub RowFunction(ByRef oForm As SAPbouiCOM.Form, ByVal mxUID As String, ByVal FieldAlias As String, ByVal TableName As String)
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oForm.Items.Item(mxUID).Specific
            ObjDbDataSource = oForm.DataSources.DBDataSources.Item(TableName)
            Dim MCount As Integer = CInt(oMatrix.RowCount)
            For MStart As Integer = 1 To MCount Step 1
                With ObjDbDataSource
                    .SetValue(FieldAlias, MStart, MStart.ToString)
                    oMatrix.Columns.Item(FieldAlias.ToString()).Cells.Item(MStart).Specific.Value = MStart.ToString()
                End With
                Debug.Print(ObjDbDataSource.GetValue(FieldAlias, MStart))
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#End Region
    Dim reportuti As New clsReportUtilities
    Dim pdfFilename As String = ""
    Dim mainFolder As String = ""
    Dim jobNo As String = ""
    Dim rptPath As String = ""
    Dim pdffilepath As String = ""
    Dim rptDocument As ReportDocument = New ReportDocument()
    Private Sub PreviewPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef oActiveForm As SAPbouiCOM.Form)
        Dim PONo As Integer
        pdfFilename = "PURCHASE ORDER"
        mainFolder = "C:\Users\PC-8\Documents\Fuji Xerox\DocuWorks\DWFolders\User Folder\"
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = IO.Directory.GetParent(System.Windows.Forms.Application.StartupPath).ToString & "\Purchase Order.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        If Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            PONo = Convert.ToInt32(oActiveForm.Items.Item("ed_PONo").Specific.Value)
        Else
            PONo = DocLastKey
        End If

        rptDocument.SetParameterValue("@DocEntry", PONo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument)
        End If
    End Sub

    Private Sub SendAttachFile(ByRef ParentForm As SAPbouiCOM.Form, ByRef oActiveForm As SAPbouiCOM.Form)
        'Save Report To specific JobFile as PDF File USE Code
        PreviewPO(ParentForm, oActiveForm)
        If oActiveForm.Items.Item("ch_Email").Specific.Checked = True Then
            oRecordSet.DoQuery("Select email from ohem a inner join ousr b on a.userid=b.userid where user_Code='" & p_oDICompany.UserName.ToString() & "'")
            If oRecordSet.RecordCount > 0 Then
                Dim frommail As String = oRecordSet.Fields.Item("email").Value.ToString
                If Not reportuti.SendMailDoc(oActiveForm.Items.Item("ed_Email").Specific.Value, frommail, "Purchase Order", "Purchase Order Item", pdffilepath) Then
                    p_oSBOApplication.MessageBox("Send Message Fail", 0, "OK")
                Else
                    p_oSBOApplication.MessageBox("Send Message Successfully", 0, "OK")
                End If
            End If
        End If

        If oActiveForm.Items.Item("ch_Fax").Specific.Checked = True Then
            ConvertFunction(oActiveForm, pdffilepath)
        End If

        rptDocument.Close()
    End Sub
    Private Sub ConvertFunction(ByVal oActiveForm As SAPbouiCOM.Form, ByVal pdffilepath As String)

        Dim SaveName As String = IO.Directory.GetParent(pdffilepath).ToString
        Dim pagerange As String = "sdfsdf"
        DoConversion(pdffilepath, "efere", SaveName.ToString(), 200, pagerange.ToString(), ImageType.TIFF)
        FindSubFolders(oActiveForm, pdffilepath)

    End Sub

    Private Sub DoConversion(ByVal file As String, ByVal password As String, ByVal folder As String, ByVal dpi As Integer, ByVal pagerange__1 As String, ByVal iType As ImageType)

        Dim format As System.Drawing.Imaging.ImageFormat
        Dim extension As String = Nothing

        ' Setup the license
        SolidFramework.License.ActivateDeveloperLicense()

        ' Set the output image type
        Select Case iType
            Case ImageType.BMP
                format = System.Drawing.Imaging.ImageFormat.Bmp
                extension = "bmp"
                Exit Select
            Case ImageType.JPG
                format = System.Drawing.Imaging.ImageFormat.Jpeg
                extension = "jpg"
                Exit Select
            Case ImageType.PNG
                format = System.Drawing.Imaging.ImageFormat.Png
                extension = "png"
                Exit Select
            Case ImageType.TIFF
                format = System.Drawing.Imaging.ImageFormat.Tiff
                extension = "tif"
                Exit Select
            Case Else
                Throw New ArgumentException("DoConversion: ImageType not known")
        End Select

        ' Load up the document
        Dim doc As New SolidFramework.Pdf.PdfDocument(file, password)
        doc.Open()

        ' Setup the outputfolder
        If Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If

        ' Setup the file string.
        Dim filename As String = folder + Path.DirectorySeparatorChar & Path.GetFileNameWithoutExtension(file)

        ' Get our pages.
        Dim Pages__2 As New List(Of SolidFramework.Pdf.Plumbing.PdfPage)(doc.Catalog.Pages.PageCount)
        Dim catalog As SolidFramework.Pdf.Catalog = DirectCast(SolidFramework.Pdf.Catalog.Create(doc), SolidFramework.Pdf.Catalog)
        Dim pages__3 As SolidFramework.Pdf.Plumbing.PdfPages = DirectCast(catalog.Pages, SolidFramework.Pdf.Plumbing.PdfPages)
        ProcessPages(pages__3, Pages__2)

        ' Check for page ranges
        Dim ranges As SolidFramework.PageRange = Nothing
        Dim bHaveRanges As Boolean = False
        If Not String.IsNullOrEmpty(pagerange__1) Then
            bHaveRanges = SolidFramework.PageRange.TryParse(pagerange__1, ranges)
        End If

        If bHaveRanges Then
            Dim pageArray As Integer() = ranges.ToArray()
            For Each number As Integer In pageArray
                CreateImageFromPage(Pages__2(number), dpi, filename, number, extension, format)
                Console.WriteLine(String.Format("Processed page {0} of {1}", number, Pages__2.Count))
            Next
        Else
            ' For each page, save off a file.
            Dim pageIndex As Integer = 0
            For Each page As SolidFramework.Pdf.Plumbing.PdfPage In Pages__2
                ' Update the page number.
                pageIndex += 1

                CreateImageFromPage(page, dpi, filename, pageIndex, extension, format)
                Console.WriteLine(String.Format("Processed page {0} of {1}", pageIndex, Pages__2.Count))
            Next
        End If
    End Sub
    Private Sub ProcessPages(ByRef pages As SolidFramework.Pdf.Plumbing.PdfPages, ByRef listPages As List(Of SolidFramework.Pdf.Plumbing.PdfPage))
        ' Walk the Pages catalog and get all the page objects.  This will follow 
        ' the references and get the actual object that we can work 
        ' with recursively.
        For Each pdfItem As SolidFramework.Pdf.Plumbing.PdfItem In pages.Kids
            Dim dictionary As SolidFramework.Pdf.Plumbing.PdfDictionary = DirectCast(SolidFramework.Pdf.Plumbing.PdfItem.GetIndirectionItem(pdfItem), SolidFramework.Pdf.Plumbing.PdfDictionary)
            If dictionary.Type = "Pages" Then
                Dim nodePages As SolidFramework.Pdf.Plumbing.PdfPages = DirectCast(dictionary, SolidFramework.Pdf.Plumbing.PdfPages)
                ProcessPages(nodePages, listPages)
            ElseIf dictionary.Type = "Page" Then
                Dim page As SolidFramework.Pdf.Plumbing.PdfPage = DirectCast(dictionary, SolidFramework.Pdf.Plumbing.PdfPage)
                listPages.Add(page)
            End If
        Next
    End Sub
    Private Sub CreateImageFromPage(ByVal page As SolidFramework.Pdf.Plumbing.PdfPage, ByVal dpi As Integer, ByVal filename As String, ByVal pageIndex As Integer, ByVal extension As String, ByVal format As System.Drawing.Imaging.ImageFormat)
        ' Create a bitmap from the page with set dpi.
        Dim bm As Bitmap = page.DrawBitmap(dpi)

        ' Setup the filename.
        Dim filepath As String = String.Format(filename & "-{0}.{1}", pageIndex, extension)

        ' If the file exits already, delete it. I.E. Overwrite it.
        If File.Exists(filepath) Then
            File.Delete(filepath)
        End If

        ' Save the file.
        bm.Save(filepath, format)

        ' Cleanup.
        bm.Dispose()
    End Sub

    Private Sub FindSubFolders(ByVal oActiveForm As SAPbouiCOM.Form, ByVal pdffilepath As String)
        Dim objFaxDocument As New FAXCOMEXLib.FaxDocument
        Dim objFaxServer As New FAXCOMEXLib.FaxServer
        Dim strFileSize As String = ""
        Dim di As New IO.DirectoryInfo("C:\Users\UNIQUE\Desktop\SavePDF")
        Try
            di.GetFiles("*.TIF*", SearchOption.AllDirectories)
        Catch
        End Try

        Dim aryFi As IO.FileInfo() = di.GetFiles("*.TIFF*")
        Dim fi As IO.FileInfo

        For Each fi In aryFi
            strFileSize = (Math.Round(fi.Length / 1024)).ToString()
            Try
                p_oSBOApplication.MessageBox(fi.Name.ToString())
                objFaxServer.Connect("PC-8")
                objFaxDocument.Body = pdffilepath
                objFaxDocument.DocumentName = "SAPITDocument"
                objFaxDocument.Sender.FaxNumber = "206-350-2896" 'oForm.Items.Item("ed_DESC").Specific.Value.ToString()
                objFaxDocument.Priority = FAXCOMEXLib.FAX_PRIORITY_TYPE_ENUM.fptHIGH
                objFaxDocument.Recipients.Add(oActiveForm.Items.Item("ed_Email").Specific.Value.ToString())
                objFaxDocument.Sender.Company = "SAPIT"
                objFaxDocument.Sender.Email = "min@sap-infotech.com"
                objFaxDocument.Sender.StreetAddress = "http://voicemail.k7.net"
                objFaxDocument.Sender.SaveDefaultSender()
                objFaxDocument.ConnectedSubmit(objFaxServer)
                System.Threading.Thread.Sleep(5000)
                objFaxServer.Disconnect()
                p_oSBOApplication.MessageBox("Successful to Send Document", MsgBoxStyle.Information, "Fax")
            Catch ex As Exception
                p_oSBOApplication.MessageBox("Fail to Send Document !", MsgBoxStyle.Critical, "ok")
            End Try

        Next

        ' File.Delete("C:\Users\UNIQUE\Desktop\SavePDF\")

    End Sub

#Region "----------Container View List & Edit"
    Private Sub AddUpdateContainer(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        'ObjDBDataSource.Offset = 0
        Dim i As Integer
        Dim stuffDate As String
        Dim istuffHr As Integer
        Dim strstuffHr As String
        rowIndex = pMatrix.GetNextSelectedRow

        If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        Try
            If ProcressedState = True Then
                'If ObjDBDataSource.GetValue("U_ConSeqNo", 0) = vbNullString Then pMatrix.Clear()
                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                        .SetValue("U_ConSeqNo", .Offset, 1)
                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                        .SetValue("U_ConSeqNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                    End If


                    .SetValue("U_ConNo", .Offset, pForm.Items.Item("ed_ContNo").Specific.Value)
                    .SetValue("U_SealNo", .Offset, pForm.Items.Item("ed_SealNo").Specific.Value)
                    .SetValue("U_Size", .Offset, pForm.Items.Item("cb_ConSize").Specific.Value)
                    .SetValue("U_Type", .Offset, pForm.Items.Item("cb_ConType").Specific.Value)
                    .SetValue("U_ContWt", .Offset, pForm.Items.Item("ed_ContWt").Specific.Value)
                    .SetValue("U_CgDesc", .Offset, pForm.Items.Item("ed_CDesc").Specific.Value)
                    If pForm.Items.Item("ch_CStuff").Specific.Checked = True Then
                        stuffDate = Right(pForm.Items.Item("ed_CunDate").Specific.Value, 2) & "." & Mid(pForm.Items.Item("ed_CunDate").Specific.Value, 5, 2) & "." & Mid(pForm.Items.Item("ed_CunDate").Specific.Value, 3, 2)
                        i = Convert.ToInt32(Left(pForm.Items.Item("ed_CunTime").Specific.Value, 2))
                        If i >= 12 Then
                            istuffHr = i - 12
                            strstuffHr = istuffHr.ToString() & ":" & Right(pForm.Items.Item("ed_CunTime").Specific.Value, 2) & "PM"
                        Else
                            istuffHr = i
                            strstuffHr = istuffHr.ToString() & ":" & Right(pForm.Items.Item("ed_CunTime").Specific.Value, 2) & "AM"
                        End If
                        'stuffTime = Left(pForm.Items.Item("ed_CunTime").Specific.Value, 2) & ":" & Right(pForm.Items.Item("ed_CunTime").Specific.Value, 2)
                        .SetValue("U_CunStuff", .Offset, stuffDate & "," & pForm.Items.Item("ed_CunDay").Specific.Value & "," & strstuffHr)
                    Else
                        .SetValue("U_CunStuff", .Offset, "")
                    End If

                    .SetValue("U_ContDate", .Offset, pForm.Items.Item("ed_CunDate").Specific.Value)
                    .SetValue("U_ContDay", .Offset, pForm.Items.Item("ed_CunDay").Specific.Value)
                    .SetValue("U_ContTime", .Offset, pForm.Items.Item("ed_CunTime").Specific.Value)
                    .SetValue("U_CHStuff", .Offset, IIf(pForm.Items.Item("ch_CStuff").Specific.Checked = True, "Y", "N"))
                    'pForm.DataSources.DBDataSources.Item(DataSource).SetValue("U_Code", 0, pForm.Items.Item("ed_CunDay").Specific.Value)
                    'pForm.DataSources.DBDataSources.Item(DataSource).SetValue("U_Name", 0, pForm.Items.Item("ed_CunTime").Specific.Value)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource
                    If pForm.Items.Item("ch_CStuff").Specific.Checked = True Then
                        stuffDate = Right(pForm.Items.Item("ed_CunDate").Specific.Value, 2) & "." & Mid(pForm.Items.Item("ed_CunDate").Specific.Value, 5, 2) & "." & Mid(pForm.Items.Item("ed_CunDate").Specific.Value, 3, 2)
                        i = Convert.ToInt32(Left(pForm.Items.Item("ed_CunTime").Specific.Value, 2))
                        If i >= 12 Then
                            istuffHr = i - 12
                            strstuffHr = istuffHr.ToString() & ":" & Right(pForm.Items.Item("ed_CunTime").Specific.Value, 2) & "PM"
                        Else
                            istuffHr = i
                            strstuffHr = istuffHr.ToString() & ":" & Right(pForm.Items.Item("ed_CunTime").Specific.Value, 2) & "AM"
                        End If
                        'stuffTime = Left(pForm.Items.Item("ed_CunTime").Specific.Value, 2) & ":" & Right(pForm.Items.Item("ed_CunTime").Specific.Value, 2)
                        .SetValue("U_CunStuff", .Offset, stuffDate & "," & pForm.Items.Item("ed_CunDay").Specific.Value & "," & strstuffHr)
                    Else
                        .SetValue("U_CunStuff", .Offset, "")
                    End If

                    .SetValue("LineId", .Offset, pForm.Items.Item("ed_ConNo").Specific.Value)
                    .SetValue("U_ConSeqNo", .Offset, pForm.Items.Item("ed_ConNo").Specific.Value)
                    .SetValue("U_ConNo", .Offset, pForm.Items.Item("ed_ContNo").Specific.Value)
                    .SetValue("U_SealNo", .Offset, pForm.Items.Item("ed_SealNo").Specific.Value)
                    .SetValue("U_Size", .Offset, pForm.Items.Item("cb_ConSize").Specific.Value)
                    .SetValue("U_Type", .Offset, pForm.Items.Item("cb_ConType").Specific.Value)
                    .SetValue("U_ContWt", .Offset, pForm.Items.Item("ed_ContWt").Specific.Value)
                    .SetValue("U_CgDesc", .Offset, pForm.Items.Item("ed_CDesc").Specific.Value)
                    '.SetValue("U_CunStuff", .Offset, stuffDate & "," & pForm.Items.Item("ed_CunDay").Specific.Value & "," & strstuffHr)
                    .SetValue("U_ContDate", .Offset, pForm.Items.Item("ed_CunDate").Specific.Value)
                    .SetValue("U_ContDay", .Offset, pForm.Items.Item("ed_CunDay").Specific.Value)
                    .SetValue("U_ContTime", .Offset, pForm.Items.Item("ed_CunTime").Specific.Value)
                    .SetValue("U_CHStuff", .Offset, IIf(pForm.Items.Item("ch_CStuff").Specific.Checked = True, "Y", "N"))

                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub SetContainerDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        Try
            pForm.Freeze(True)
            Dim oComboType As SAPbouiCOM.ComboBox
            Dim oComboSize As SAPbouiCOM.ComboBox
            oComboType = pForm.Items.Item("cb_ConType").Specific
            oComboSize = pForm.Items.Item("cb_ConSize").Specific
            If Conunstuff <> "" Then
                pForm.Items.Item("ch_CStuff").Specific.Checked = True
            Else
                pForm.Items.Item("ch_CStuff").Specific.Checked = False
            End If
            pForm.Items.Item("ed_ConNo").Specific.Value = ConSeqNo
            pForm.Items.Item("ed_ContNo").Specific.Value = ConNo
            pForm.Items.Item("ed_SealNo").Specific.Value = ConSealNo
            ' pForm.Items.Item("cb_ConSize").Specific.Value = ConSize
            oComboType.Select(ConType, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oComboSize.Select(ConSize, SAPbouiCOM.BoSearchKey.psk_ByValue)
            'pForm.Items.Item("cb_ConType").Specific.Value = ConType
            pForm.Items.Item("ed_ContWt").Specific.Value = ConWt
            pForm.Items.Item("ed_CDesc").Specific.Value = ConDesc
            pForm.Items.Item("ed_CunDate").Specific.Value = ConDate
            pForm.Items.Item("ed_CunDay").Specific.Value = ConDay
            pForm.Items.Item("ed_CunTime").Specific.Value = ConTime
            If HolidayMarkUp(pForm, pForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, pForm.Items.Item("ed_CunDay").Specific, pForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            pForm.Freeze(False)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub GetContainerDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Try
            ConSeqNo = pMatrix.Columns.Item("V_-1").Cells.Item(Index).Specific.Value
            ConNo = pMatrix.Columns.Item("colCNo").Cells.Item(Index).Specific.Value
            ConSealNo = pMatrix.Columns.Item("colSealNo").Cells.Item(Index).Specific.Value
            ConSize = pMatrix.Columns.Item("colCSize").Cells.Item(Index).Specific.Value
            ConType = pMatrix.Columns.Item("colCType").Cells.Item(Index).Specific.Value
            ConWt = pMatrix.Columns.Item("colContWt").Cells.Item(Index).Specific.Value
            ConDesc = pMatrix.Columns.Item("colConDesc").Cells.Item(Index).Specific.Value
            ConDate = pMatrix.Columns.Item("colCDate").Cells.Item(Index).Specific.Value
            ConDay = pMatrix.Columns.Item("colCDay").Cells.Item(Index).Specific.Value
            ConTime = pMatrix.Columns.Item("colCTime").Cells.Item(Index).Specific.Value
            Conunstuff = pMatrix.Columns.Item("colUnstuff").Cells.Item(Index).Specific.Value
            If Conunstuff <> "" Then
                ChStuff = "Y"
            Else
                ChStuff = "N"
            End If
            'ConDay = pMatrix.Columns.Item("colTel").Cells.Item(Index).Specific.Value
            'ConTime = pMatrix.Columns.Item("colFax").Cells.Item(Index).Specific.Value
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DeleteContainerByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String)

        Try
            If pMatrix.IsRowSelected(pMatrix.GetNextSelectedRow) = True Then
                Try
                    If (pForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        pForm.DataSources.DBDataSources.Item("@OBT_TB020_HCONTAINE").RemoveRecord(pMatrix.GetNextSelectedRow - 1)
                    End If
                Catch ex As Exception

                End Try

                pMatrix.DeleteRow(pMatrix.GetNextSelectedRow)
                SetMatrixSeqNo(pMatrix, "V_-1")
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pMatrix.RowCount = 0 Then
                    pMatrix.FlushToDataSource()
                    pMatrix.AddRow(1)
                End If
                pForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container"
                pForm.Items.Item("bt_DelCont").Enabled = False
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub UpdateNoofContainer(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix)
        'MSW No of Container

        Dim iCon20 As Integer = 0
        Dim iCon40 As Integer = 0
        Dim iCon45 As Integer = 0
        pMatrix = pForm.Items.Item("mx_ConTab").Specific

        If pMatrix.RowCount > 0 Then
            For i As Integer = 1 To pMatrix.RowCount
                If pMatrix.Columns.Item("colCSize").Cells.Item(i).Specific.Value() = "20'" Then
                    iCon20 = iCon20 + 1
                ElseIf pMatrix.Columns.Item("colCSize").Cells.Item(i).Specific.Value() = "40'" Then
                    iCon40 = iCon40 + 1
                ElseIf pMatrix.Columns.Item("colCSize").Cells.Item(i).Specific.Value() = "45'" Then
                    iCon45 = iCon45 + 1
                End If
            Next
            pForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Con20", 0, iCon20)
            pForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Con40", 0, iCon40)
            pForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Con45", 0, iCon45)
        End If

        'End
    End Sub
#End Region

#Region "---------- 'MSW - Voucher Tab View List & Edit"

    Private Sub AddUpdateVoucher(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
        ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item(DataSource)
        '  ObjDBDataSource.Offset = 0

        rowIndex = pMatrix.GetNextSelectedRow

        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        'MSW Voucher POP Up
        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            If pMatrix.RowCount = 1 And pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                pMatrix.Clear()
            End If
        End If
        'End MSW Voucher POP Up
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
                    .SetValue("U_RefNo", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value) 'MSW Voucher POP Up
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
                    .SetValue("U_VDocNum", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value) 'MSW Voucher POP Up

                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource
                    .SetValue("LineId", .Offset, pForm.Items.Item("ed_VocNo").Specific.Value)
                    .SetValue("U_VSeqNo", .Offset, pForm.Items.Item("ed_VocNo").Specific.Value)
                    .SetValue("U_RefNo", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value) 'MSW Voucher POP Up
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
                    .SetValue("U_VDocNum", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value) 'MSW Voucher POP Up

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
#End Region

    Private Sub AddUpdateDisp(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer)
        Dim ObjDBDataSource As SAPbouiCOM.DBDataSource
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        Dim a As String = pForm.Items.Item("cb_Dspchr").Specific.Value
        Dim Row As Integer = 0
        Try
            If ProcressedState = True Then
                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                        .SetValue("U_DisSeqNo", .Offset, 1)
                        Row = 1
                    Else
                        If pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value <> Nothing Then
                            Row = pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value
                            .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                            .SetValue("U_DisSeqNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                        Else
                            .SetValue("LineId", .Offset, 1)
                            .SetValue("U_DisSeqNo", .Offset, 1)
                        End If

                    End If
                    .SetValue("U_Dispatch", .Offset, pForm.Items.Item("cb_Dspchr").Specific.Value)
                    .SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_DspDate").Specific.Value)
                    .SetValue("U_InsTime", .Offset, pForm.Items.Item("ed_DspHr").Specific.Value)
                    .SetValue("U_InsDay", .Offset, pForm.Items.Item("ed_DspDay").Specific.Value)
                    .SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_DspIntr").Specific.Selected = True, "Internal", "External").ToString)
                    .SetValue("U_Instruct", .Offset, pForm.Items.Item("ee_Instru").Specific.Value)
                    Dim DspDate As String = pForm.Items.Item("ed_DspCDte").Specific.Value
                    Dim DspTime As String = pForm.Items.Item("ed_DspCHr").Specific.Value
                    If DspDate <> Nothing Then
                        If CInt(DspTime.Substring(0, 2).ToString()) > 12 Then
                            .SetValue("U_Complete", .Offset, DspDate.Substring(6, 2).ToString() + "." + DspDate.Substring(4, 2).ToString() + "." + DspDate.Substring(2, 2).ToString() + "," + pForm.Items.Item("ed_DspCDay").Specific.Value + "," + (CInt(DspTime.Substring(0, 2).ToString()) - 12).ToString() + ":" + DspTime.Substring(2, 2) + " PM")
                        Else
                            .SetValue("U_Complete", .Offset, DspDate.Substring(6, 2).ToString() + "." + DspDate.Substring(4, 2).ToString() + "." + DspDate.Substring(2, 2).ToString() + "," + pForm.Items.Item("ed_DspCDay").Specific.Value + "," + DspTime.Substring(0, 2) + ":" + DspTime.Substring(2, 2) + " AM")
                        End If

                        .SetValue("U_ComDate", .Offset, pForm.Items.Item("ed_DspCDte").Specific.Value)
                        .SetValue("U_ComDay", .Offset, pForm.Items.Item("ed_DspCDay").Specific.Value)
                        .SetValue("U_ComTime", .Offset, pForm.Items.Item("ed_DspCHr").Specific.Value)
                    End If
                    If Row = 0 Then
                        pMatrix.SetLineData(1)
                    Else
                        pMatrix.AddRow()
                    End If
                End With
            Else
                With ObjDBDataSource
                    .SetValue("LineId", .Offset, Index)
                    .SetValue("U_DisSeqNo", .Offset, Index)
                    .SetValue("U_Dispatch", .Offset, pForm.Items.Item("cb_Dspchr").Specific.Value)
                    .SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_DspDate").Specific.Value)
                    .SetValue("U_InsTime", .Offset, pForm.Items.Item("ed_DspHr").Specific.Value)
                    .SetValue("U_InsDay", .Offset, pForm.Items.Item("ed_DspDay").Specific.Value)
                    .SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_DspIntr").Specific.Selected = True, "Internal", "External").ToString)
                    .SetValue("U_Instruct", .Offset, pForm.Items.Item("ee_Instru").Specific.Value)
                    Dim DspDate As String = pForm.Items.Item("ed_DspCDte").Specific.Value
                    Dim DspTime As String = pForm.Items.Item("ed_DspCHr").Specific.Value
                    If DspDate <> Nothing Then
                        If CInt(DspTime.Substring(0, 2).ToString()) > 12 Then
                            .SetValue("U_Complete", .Offset, DspDate.Substring(6, 2).ToString() + "." + DspDate.Substring(4, 2).ToString() + "." + DspDate.Substring(2, 2).ToString() + "," + pForm.Items.Item("ed_DspCDay").Specific.Value + "," + (CInt(DspTime.Substring(0, 2).ToString()) - 12).ToString() + ":" + DspTime.Substring(2, 2) + " PM")
                        Else
                            .SetValue("U_Complete", .Offset, DspDate.Substring(6, 2).ToString() + "." + DspDate.Substring(4, 2).ToString() + "." + DspDate.Substring(2, 2).ToString() + "," + pForm.Items.Item("ed_DspCDay").Specific.Value + "," + DspTime.Substring(0, 2) + ":" + DspTime.Substring(2, 2) + " AM")
                        End If

                        .SetValue("U_ComDate", .Offset, pForm.Items.Item("ed_DspCDte").Specific.Value)
                        .SetValue("U_ComDay", .Offset, pForm.Items.Item("ed_DspCDay").Specific.Value)
                        .SetValue("U_ComTime", .Offset, pForm.Items.Item("ed_DspCHr").Specific.Value)
                    End If
                    pMatrix.SetLineData(Index)
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub SetDispatchDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Try
            Dim sErrDesc As String
            Dim ocombo As SAPbouiCOM.ComboBox
            ocombo = pForm.Items.Item("cb_Dspchr").Specific

            pForm.Items.Item("ed_DspDate").Specific.String = pMatrix.Columns.Item("colDate").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_DspHr").Specific.String = pMatrix.Columns.Item("colTime").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_DspDay").Specific.String = pMatrix.Columns.Item("colInsDay").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ee_Instru").Specific.Value = pMatrix.Columns.Item("colInst").Cells.Item(Index).Specific.Value
            If (pMatrix.Columns.Item("colComDate").Cells.Item(Index).Specific.Value <> Nothing) Then
                pForm.Items.Item("ch_Dsp").Specific.Checked = True
                pForm.Items.Item("ed_DspCDte").Specific.Value = pMatrix.Columns.Item("colComDate").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_DspCDay").Specific.Value = pMatrix.Columns.Item("colComDay").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_DspCHr").Specific.Value = pMatrix.Columns.Item("colComTime").Cells.Item(Index).Specific.Value
            Else
                pForm.Items.Item("ch_Dsp").Specific.Checked = False
            End If
            Dim a As String = pForm.Items.Item("ed_DspCHr").Specific.String
            If pMatrix.Columns.Item("colMode").Cells.Item(Index).Specific.Value = "Internal" Then
                'pForm.DataSources.UserDataSources.Item("DSDisp").ValueEx = "1"
                pForm.Items.Item("op_DspIntr").Specific.Selected = True
                ocombo.Select(pMatrix.Columns.Item("colDisp").Cells.Item(Index).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                If (pMatrix.Columns.Item("colDisp").Cells.Item(Index).Specific.Value = Nothing) Then
                    pForm.Items.Item("op_DspExtr").Specific.Selected = True
                    pForm.Items.Item("op_DspIntr").Specific.Selected = True
                End If
            ElseIf pMatrix.Columns.Item("colMode").Cells.Item(Index).Specific.Value = "External" Then
                'pForm.DataSources.UserDataSources.Item("DSDisp").ValueEx = "2"
                pForm.Items.Item("op_DspExtr").Specific.Selected = True
                ocombo.Select(pMatrix.Columns.Item("colDisp").Cells.Item(Index).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                If (pMatrix.Columns.Item("colDisp").Cells.Item(Index).Specific.Value = Nothing) Then
                    pForm.Items.Item("op_DspIntr").Specific.Selected = True
                    pForm.Items.Item("op_DspExtr").Specific.Selected = True
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddUpdateOtherCharges(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer)
        Dim ObjDBDataSource As SAPbouiCOM.DBDataSource
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        Dim Row As Integer = 0
        Try
            If ProcressedState = True Then
                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                        .SetValue("U_ChSeqNo", .Offset, 1)
                        Row = 1
                    Else
                        .SetValue("LineId", .Offset, pForm.Items.Item("ed_CSeqNo").Specific.Value)
                        .SetValue("U_ChSeqNo", .Offset, pForm.Items.Item("ed_CSeqNo").Specific.Value)
                        If pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value <> Nothing Then
                            Row = pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value
                        End If

                    End If
                    .SetValue("U_CCode", .Offset, pForm.Items.Item("ed_ChCode").Specific.Value)
                    .SetValue("U_CClaim", .Offset, pForm.Items.Item("cb_Claim").Specific.Value)
                    .SetValue("U_Remarks", .Offset, pForm.Items.Item("ed_Remarks").Specific.Value)
                    If Row = 0 Then
                        pMatrix.SetLineData(1)
                    Else
                        pMatrix.AddRow()
                    End If
                End With
            Else
                With ObjDBDataSource
                    .SetValue("LineId", .Offset, Index)
                    .SetValue("U_ChSeqNo", .Offset, Index)
                    .SetValue("U_CCode", .Offset, pForm.Items.Item("ed_ChCode").Specific.Value)
                    .SetValue("U_CClaim", .Offset, pForm.Items.Item("cb_Claim").Specific.Value)
                    .SetValue("U_Remarks", .Offset, pForm.Items.Item("ed_Remarks").Specific.Value)
                    pMatrix.SetLineData(Index)
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DeleteByIndexOtherCharges(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String)
        Try
            Dim lRow As Long
            lRow = pMatrix.GetNextSelectedRow
            If lRow > -1 Then
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    Try
                        pForm.DataSources.DBDataSources.Item("@OBT_TB024_HCHARGES").RemoveRecord(lRow - 1)
                    Catch ex As Exception

                    End Try
                End If
                pMatrix.DeleteRow(lRow)
                If lRow = 1 And pMatrix.RowCount = 0 Then
                    pMatrix.FlushToDataSource()
                    pMatrix.AddRow()
                    pForm.Items.Item("bt_AmdCh").Enabled = False
                End If
                pMatrix.FlushToDataSource()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetOtherChargesDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Try
            Dim ocombo As SAPbouiCOM.ComboBox
            ocombo = pForm.Items.Item("cb_Claim").Specific
            pForm.Items.Item("ed_CSeqNo").Specific.Value = pMatrix.Columns.Item("V_-1").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_ChCode").Specific.Value = pMatrix.Columns.Item("colCCode").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Remarks").Specific.Value = pMatrix.Columns.Item("colCRemark").Cells.Item(Index).Specific.Value
            ocombo.Select(pMatrix.Columns.Item("colCClaim").Cells.Item(Index).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoadHolidayMarkUp(ByVal oActiveForm As SAPbouiCOM.Form)
        Dim sErrDesc As String = String.Empty
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Private Function AddChooseFromListByOption(ByRef pForm As SAPbouiCOM.Form, ByVal pOption As Boolean, ByVal pObjID As String, ByVal pErrDesc As String) As Long
        Dim oEditText As SAPbouiCOM.EditText
        AddChooseFromListByOption = RTN_ERROR
        Try
            If pOption = True Then
                oEditText = pForm.Items.Item(pObjID).Specific
                oEditText.DataBind.SetBound(True, "", "TKRINTR")
                oEditText.ChooseFromListUID = "CFLTKRE"
                oEditText.ChooseFromListAlias = "firstName"
            Else
                oEditText = pForm.Items.Item(pObjID).Specific
                oEditText.DataBind.SetBound(True, "", "TKREXTR")
                'oEditText.ChooseFromListUID = ""
                oEditText.ChooseFromListUID = "CFLTKRV"
                'oEditText.ChooseFromListAlias = "CardName"
            End If
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

        oMatrix.Columns.Item("colGSTAmt").Editable = True
        oMatrix.Columns.Item("colNoGST").Editable = True
        oMatrix.Columns.Item("colGSTAmt").Cells.Item(Row).Specific.Value = Convert.ToString(GSTAMT)
        oMatrix.Columns.Item("colNoGST").Cells.Item(Row).Specific.Value = Convert.ToString(NOGST)
        oActiveForm.Items.Item("ed_VedName").Specific.Active = True
        oMatrix.Columns.Item("colGSTAmt").Editable = False
        oMatrix.Columns.Item("colNoGST").Editable = False

        '====CalCulate Total======'
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

    Private Sub SetMatrixSeqNo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal ColName As String)
        For i As Integer = 1 To oMatrix.RowCount
            oMatrix.Columns.Item(ColName).Cells.Item(i).Specific.Value = i
        Next
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

    Private Function Validateforform(ByVal ItemUID As String, ByVal oActiveForm As SAPbouiCOM.Form) As Boolean
        If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_Name").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "cb_PCode" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("cb_PCode").Specific.Value) Then
            p_oSBOApplication.SetStatusBarMessage("Must Select Port Of Loading", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_ShpAgt").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Shipping Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_Yard" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_Yard").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Choose Return Yard", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
            'ElseIf (ItemUID = "ed_LCDDate" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_LCDDate").Specific.String) Then
            '    p_oSBOApplication.SetStatusBarMessage("Must Choose Last Clearance Date", SAPbouiCOM.BoMessageTime.bmt_Short)
            '    Return True
        ElseIf (ItemUID = "ed_YAddr" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_YAddr").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Fill Return Yard Address", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        ElseIf (ItemUID = "ed_JobNo" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_JobNo").Specific.String) Then
            p_oSBOApplication.SetStatusBarMessage("Must Fill Job No", SAPbouiCOM.BoMessageTime.bmt_Short)
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub EnabledMaxtix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)
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


End Module