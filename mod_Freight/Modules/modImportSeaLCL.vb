Option Explicit On

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
Imports CrystalDecisions.CrystalReports.Engine
Imports SAPbobsCOM


Module modImportSeaLCL

    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Private VocNo, VendorName, PayTo, PayType, BankName, CheqNo, Status, Currency, PaymentDate, GST, PrepBy, ExRate, Remark As String
    Private Total, GSTAmt, SubTotal, vocTotal, gstTotal As Double
    Private vendorCode As String
    Private sPicName As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sql As String = ""
    Dim strAPInvNo, strOutPayNo As String
    Private currentRow As Integer
    Private ActiveMatrix As String
    Private DocLastKey As String
    Private DocStatus As String

    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String
    Dim vedCurCode As String = String.Empty
    Dim rowIndex As Integer


    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public Function DoImportSeaLCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
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
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oGRForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
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
        Dim matrixName As String = ""

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", FunctionName)
            Select Case BusinessObjectInfo.FormTypeEx
                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000050", "2000000038"  'MSW to Edit New Ticket
                    'If BusinessObjectInfo.ActionSuccess = False Then
                    '    oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    '    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = True Then
                    '        If Not CreatePurchaseOrder(oActiveForm, "mx_Item") Then
                    '            Throw New ArgumentException(sErrDesc)
                    '            BubbleEvent = False
                    '        End If
                    '    End If
                    'End If
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then

                            'When action of BusinessObjectInfo is nicely done, need to do 2 tasks in action for Purcahse process
                            ' (1). need to show PopulatePurchaseHeader into the matrix of Main Export Form
                            ' (2). need to create PurchaseOrder into OPOR and POR1, related with main PurchaseProcess by using oPurchaseOrder document
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If

                            'oMatrix = ImportSeaLCLForm.Items.Item("mx_Fumi").Specific
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                        "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oActiveForm.Items.Item("ed_CPOID").Specific.Value)
                            If Not CreatePurchaseOrder(oActiveForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            If BusinessObjectInfo.FormTypeEx = "2000000005" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Fumi").Specific
                                matrixName = "mx_Fumi"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL21_FUMI", True, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000021" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Fork").Specific
                                matrixName = "mx_Fork"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL22_FORKLIFT", True, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Armed").Specific
                                matrixName = "mx_Armed"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL24_ARMES", True, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000010" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Crane").Specific
                                matrixName = "mx_Crane"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL25_CRANE", True, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Bunk").Specific
                                matrixName = "mx_Bunk"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL26_BUNKER", True, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000038" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Orider").Specific
                                matrixName = "mx_Orider"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL27_OUTRIDER", True, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                matrixName = "mx_TkrList"
                                If Not PopulateTruckingPOToEditTab(ImportSeaLCLForm, sql, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            End If
                            'If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql) Then Throw New ArgumentException(sErrDesc)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oActiveForm.Items.Item("ed_CPOID").Specific.Value))
                            ' SendAttachFile(ImportSeaLCLForm, oActiveForm)
                            If matrixName = "mx_Armed" Or matrixName = "mx_TkrList" Then 'MSW 14-09-2011 Truck PO
                                SendAttachFile(ImportSeaLCLForm, oActiveForm)
                            Else
                                CreatePDF(ImportSeaLCLForm, matrixName)
                            End If
                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                       "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oActiveForm.Items.Item("ed_CPOID").Specific.Value)
                            If BusinessObjectInfo.FormTypeEx = "2000000005" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Fumi").Specific
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL21_FUMI", False, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000021" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Fork").Specific
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL22_FORKLIFT", False, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL24_ARMES", False, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000010" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Crane").Specific
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL25_CRANE", False, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL26_BUNKER", False, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000038" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_Orider").Specific
                                matrixName = "mx_Orider"
                                If Not PopulatePurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL27_OUTRIDER", False, oActiveForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                       "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                       "from  [@OBT_TB08_FFCPO] a inner join [@OBT_LCL03_TRUCKING] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oActiveForm.Items.Item("ed_CPOID").Specific.Value)
                                If Not PopulateTruckPurchaseHeader(ImportSeaLCLForm, oMatrix, sql, "@OBT_LCL03_TRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            End If
                            If Not UpdatePurchaseOrder(oActiveForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                        End If
                        'End If
                    End If

                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000051", "2000000039"  'MSW to Edit New Ticket
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oGRForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("IMPORTSEALCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                        ElseIf AlreadyExist("IMPORTAIR") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                        ElseIf AlreadyExist("IMPORTLAND") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                        End If
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            If Not CreateGoodsReceiptPO(oGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)

                            If BusinessObjectInfo.FormTypeEx = "2000000012" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL22_FORKLIFT", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000013" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL24_ARMES", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000014" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL25_CRANE", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000016" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL26_BUNKER", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000039" Then
                                oMatrix = oActiveForm.Items.Item("mx_Orider").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL27_OUTRIDER", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000051" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                        "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                        "from  [@OBT_TB08_FFCPO] a inner join [@OBT_LCL03_TRUCKING] b on a.DocEntry=b.U_PODocNo where a.U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL03_TRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            End If
                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then

                        End If
                    End If
                Case "SHIPPINGINV"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            oShpForm = p_oSBOApplication.Forms.GetForm("SHIPPINGINV", 1)
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                            If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_LCL19_SHPINV", True)
                            ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_LCL19_SHPINV", False)
                            End If
                            ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_CGR").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                            ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011

                        Catch ex As Exception

                        End Try


                    End If
                    'Voucher POP UP
                Case "VOUCHER"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        If AlreadyExist("IMPORTSEALCL") Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                        ElseIf AlreadyExist("IMPORTAIR") Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                        ElseIf AlreadyExist("IMPORTLAND") Then
                            ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                        End If
                        oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_LCL05_VOUCHER", True)
                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_LCL05_VOUCHER", False)
                        End If

                    End If

               
                Case "IMPORTSEALCL", "IMPORTAIR", "IMPORTLAND"
                    ImportSeaLCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
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
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            Dim JobLastDocEntry As Integer
                            Dim ObjectCode As String = String.Empty
                            sql = "select top 1 Docentry from [@OBT_LCL01_IMPSEALCL] order by docentry desc"
                            oRecordSet.DoQuery(sql)
                            Dim FrDocEntry As Integer = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString)
                            Dim NewJobNo As String = GetJobNumber("IM")
                            sql = "select top 1 Docentry from [@OBT_FREIGHTDOCNO] order by docentry desc"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.Fields.Item("Docentry").Value.ToString = "" Then
                                JobLastDocEntry = 1
                            Else
                                JobLastDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString) + 1
                            End If
                            sql = "Update [@OBT_LCL01_IMPSEALCL] set U_JbDocNo=" & JobLastDocEntry & ",U_JobNum = '" & NewJobNo & "' Where DocEntry=" & FrDocEntry & ""
                            oRecordSet.DoQuery(sql)
                            p_oSBOApplication.SetStatusBarMessage("Actual Job Number is " & NewJobNo, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            sql = "Insert Into [@OBT_FREIGHTDOCNO] (DocEntry,DocNum,U_JobNo,U_JobMode,U_JobType,U_JbStus,U_FrDocNo,U_JbDate,U_ObjType,U_CusCode,U_CusName,U_ShpCode,U_ShpName) Values " & _
                                "(" & JobLastDocEntry & _
                                    "," & JobLastDocEntry & _
                                   "," & IIf(NewJobNo <> "", FormatString(NewJobNo), "NULL") & _
                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value), "NULL") & _
                                    "," & IIf(ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value), "NULL") & _
                                    "," & IIf(ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value), "NULL") & _
                                    "," & FrDocEntry & _
                                     "," & IIf(ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value), "Null") & _
                                    "," & IIf(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString() <> "", FormatString(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString()), "Null") & _
                                     "," & IIf(ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value), "Null") & _
                                      "," & IIf(ImportSeaLCLForm.Items.Item("ed_Name").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_Name").Specific.Value), "Null") & _
                                         "," & IIf(ImportSeaLCLForm.Items.Item("ed_V").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_V").Specific.Value), "Null") & _
                                    "," & IIf(ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value <> "", FormatString(ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value), "Null") & ")"
                            oRecordSet.DoQuery(sql)
                        End If
                    ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                            oRecordSet.DoQuery(sql)
                        End If

                    End If

                    'End MSW
                    'MSW New Button at Choose From List 22-03-2011
                Case "VESSEL"
                    Dim vesCode As String = String.Empty
                    Dim voyNo As String = String.Empty
                    Dim oImportSeaLCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select top 1* from [@OBT_TB018_VESSEL] order by DocEntry desc")
                            If oRecordSet.RecordCount > 0 Then
                                vesCode = oRecordSet.Fields.Item("Name").Value.ToString
                                voyNo = oRecordSet.Fields.Item("U_Voyage").Value.ToString
                            End If
                            oImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = vesCode
                            oImportSeaLCLForm.Items.Item("ed_Voy").Specific.Value = voyNo

                            'Try
                            '    oImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Active = True
                            'Catch ex As Exception

                            'End Try

                        End If
                    Else
                        If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE) And BusinessObjectInfo.BeforeAction = False Then
                            If BusinessObjectInfo.ActionSuccess = True Then
                                oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                                oImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = ""
                                oImportSeaLCLForm.Items.Item("ed_Voy").Specific.Value = ""

                                'Try
                                '    oImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Active = True
                                'Catch ex As Exception

                                'End Try
                            End If
                        End If
                    End If
                    'MSW New Button at Choose From List 22-03-2011
                    'MSW to Edit New Ticket
                Case "CHARGES"
                    Dim chDesc As String = String.Empty
                    Dim oCombo As SAPbouiCOM.ComboBox
                    Dim oImportSeaLCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            oCombo = oImportSeaLCLForm.Items.Item("cb_PayFor").Specific
                            While oCombo.ValidValues.Count > 0
                                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            End While
                            If oCombo.ValidValues.Count = 0 Then
                                oRecordSet.DoQuery("SELECT U_ChType FROM [@OBT_TB40_CHARGES]")
                                If oRecordSet.RecordCount > 0 Then
                                    oCombo.ValidValues.Add("", "")
                                    oRecordSet.MoveFirst()
                                    While oRecordSet.EoF = False
                                        oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_ChType").Value.ToString.Trim, "")
                                        oRecordSet.MoveNext()
                                    End While
                                    oCombo.ValidValues.Add("Define New", "")
                                End If
                            End If
                            If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                                oRecordSet.DoQuery("select top 1 U_ChType from [@OBT_TB40_CHARGES] order by DocEntry desc")
                                If oRecordSet.RecordCount > 0 Then
                                    chDesc = oRecordSet.Fields.Item("U_ChType").Value.ToString
                                End If
                                oCombo.Select(chDesc, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            Else
                                oRecordSet.DoQuery("select top 1 U_ChType from [@OBT_TB40_CHARGES] order by updateDate desc")
                                If oRecordSet.RecordCount > 0 Then
                                    chDesc = oRecordSet.Fields.Item("U_ChType").Value.ToString
                                End If
                                oCombo.Select(chDesc, SAPbouiCOM.BoSearchKey.psk_ByValue)
                            End If
                        End If
                    Else
                        If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE) And BusinessObjectInfo.BeforeAction = False Then
                            If BusinessObjectInfo.ActionSuccess = True Then
                                oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                                oCombo = oImportSeaLCLForm.Items.Item("cb_PayFor").Specific
                                While oCombo.ValidValues.Count > 0
                                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                End While
                                If oCombo.ValidValues.Count = 0 Then
                                    oRecordSet.DoQuery("SELECT U_ChType FROM [@OBT_TB40_CHARGES]")
                                    If oRecordSet.RecordCount > 0 Then
                                        oCombo.ValidValues.Add("", "")
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_ChType").Value.ToString.Trim, "")

                                            oRecordSet.MoveNext()
                                        End While
                                        oCombo.ValidValues.Add("Define New", "")
                                    End If
                                End If

                                oCombo.Select("", "")
                            End If
                        End If
                    End If
                    'MSW to Edit New Ticket


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
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing
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
                        LoadImportSeaLCLForm()
                    End If

                Case "1281"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTAIR" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTLAND" Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If

                            ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = True
                            ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Active = True
                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ImportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                                ImportSeaLCLForm.Items.Item("ed_Wrhse").Enabled = True
                            End If
                            If AddChooseFromListByOption(ImportSeaLCLForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                    End If
                Case "1292"
                    If pVal.BeforeAction = False Then

                        'Export Voucher POP UP
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                If oMatrix.Columns.Item("colChCode1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If
                            
                            'Export Voucher POP UP
                        ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Then
                            If pVal.BeforeAction = True Then
                                BubbleEvent = False
                            End If
                            'ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                            ''-------------------------For Payment(omm)------------------------------------------'
                            ''If (ImportSeaLCLForm.PaneLevel = 21) Then
                            'oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                            'RowAddToMatrix(ImportSeaLCLForm, oMatrix)
                            ''End If
                            ''----------------------------------------------------------------------------------'
                        End If
                       
                    End If

                Case "1293"
                    
                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                        oPayForm = p_oSBOApplication.Forms.ActiveForm
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                            If pVal.BeforeAction = True Then
                                If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                    BubbleEvent = False
                                End If

                                If BubbleEvent = True Then
                                    DeleteMatrixRow(oPayForm, oMatrix, "@OBT_TB032_VDETAIL", "V_-1")
                                    BubbleEvent = False
                                    If Not oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    CalculateTotal(oPayForm, oMatrix)
                                End If
                            End If
                        Else
                            BubbleEvent = False
                        End If
                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000005" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000020" Or _
                    p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000021" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000009" Or _
                    p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000010" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000015" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000050" Then
                        CPOForm = p_oSBOApplication.Forms.ActiveForm
                        'If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oMatrix = CPOForm.Items.Item("mx_Item").Specific
                        ' Dim lRow As Long
                        If pVal.BeforeAction = True Then
                            If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                BubbleEvent = False
                            End If

                            If BubbleEvent = True Then
                                DeleteMatrixRow(CPOForm, oMatrix, "@OBT_TB09_FFCPOITEM", "LineId")
                                BubbleEvent = False
                                If Not CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                                CalculateTotalPO(CPOForm, oMatrix)
                            End If
                        End If
                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Then
                        If pVal.BeforeAction = True Then
                            ' BubbleEvent = False
                            'MSW To Edit
                            Dim shpDocEntry As String
                            ImportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                            If oMatrix.GetNextSelectedRow > 0 Then
                                If oMatrix.Columns.Item("colDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString <> "" Then
                                    shpDocEntry = oMatrix.Columns.Item("colSDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString
                                    DeleteMatrixRow(ImportSeaLCLForm, oMatrix, "@OBT_LCL19_SHPINV", "V_-1")
                                    DeleteUDO(ImportSeaLCLForm, "SHIPPINGINV", shpDocEntry)
                                    ImportSeaLCLForm.Items.Item("1").Click()
                                    If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                End If
                            End If
                            BubbleEvent = False
                            'End MSW To Edit
                        End If
                        'If pVal.BeforeAction = False Then
                        '    '-------------------------For Payment(omm)------------------------------------------'
                        '    'oMatrix = ImportSeaLCLForm.Items.Item("mx_ChCode").Specific
                        '    'dtmatrix = ImportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
                        '    'If dtmatrix.Rows.Count > 0 Then
                        '    '    dtmatrix.Rows.Remove(gridindex - 1)
                        '    'End If

                        '    'oMatrix.LoadFromDataSource()
                        '    '----------------------------------------------------------------------------------'
                        'End If
                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Or _
                         p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000007" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000011" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000012" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000013" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000014" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000016" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000051" Then
                        If pVal.BeforeAction = True Then
                            BubbleEvent = False
                        End If
                    End If
                    
                Case "1282"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTAIR" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTLAND" Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If
                            EnabledHeaderControls(ImportSeaLCLForm, False) '25-3-2011
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEALCL'")
                            If oRecordSet.RecordCount > 0 Then
                                ' ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW for Job Type Table
                                ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString  'MSW for Job Type Table
                            End If
                            ImportSeaLCLForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString 'Prep By for Header
                            ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("IM")
                            ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Import Sea LCL" 'MSW 08-06-2011 for Job Type Table
                            ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            ImportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                            ImportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ImportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                        End If
                    End If

                Case "1288", "1289", "1290", "1291"
                    If pVal.BeforeAction = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTAIR" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTLAND" Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If
                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If
                        End If
                    End If
                Case "EditVoc"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTAIR" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTLAND" Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If
                            LoadPaymentVoucher(ImportSeaLCLForm)
                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oPayForm.Items.Item("ed_DocNum").Visible = True
                            oPayForm.Items.Item("ed_DocNum").Enabled = True
                            oPayForm.Items.Item("ed_DocNum").Specific.Value = oMatrix.Columns.Item("colVDocNum").Cells.Item(currentRow).Specific.Value.ToString
                            oPayForm.Items.Item("cb_PayCur").Specific.Active = True
                            oPayForm.Items.Item("ed_DocNum").Visible = False
                            oPayForm.Items.Item("ed_DocNum").Enabled = False
                            oPayForm.DataBrowser.BrowseBy = "ed_DocNum"
                            'MSW
                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                            oPayForm.Items.Item("ed_VedName").Enabled = False

                            If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                            oPayForm.Items.Item("1").Click()
                            If oPayForm.Items.Item("op_Cash").Specific.Selected = True Then
                                oPayForm.Items.Item("cb_BnkName").Enabled = False
                                oPayForm.Items.Item("ed_Cheque").Enabled = False
                            ElseIf oPayForm.Items.Item("op_Cheq").Specific.Selected = True Then
                                oPayForm.Items.Item("cb_BnkName").Enabled = True
                                oPayForm.Items.Item("ed_Cheque").Enabled = True
                            End If
                         
                        End If

                        'End If
                    End If

                Case "EditShp"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEALCL" Then

                            ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)

                            LoadShippingInvoice(ImportSeaLCLForm)
                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_ShpInv").Specific
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
                            oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                            'MSW to Edit
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If (oShpMatrix.RowCount > 0) Then
                                If (oShpMatrix.Columns.Item("V_-1").Cells.Item(oShpMatrix.RowCount).Specific.Value = Nothing) Then
                                    oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
                                    oShpForm.Items.Item("ed_Box").Specific.Value = 1
                                    oShpForm.Items.Item("ed_Part").Specific.Active = True
                                Else
                                    oShpForm.Items.Item("ed_ItemNo").Specific.Value = oShpMatrix.Columns.Item("V_-1").Cells.Item(oShpMatrix.RowCount).Specific.Value + 1
                                    oRecordSet.DoQuery("Select ISNULL(U_BoxLast,0)+1 as Box from [@OBT_TB03_EXPSHPINVD] Where LineId=" & oShpMatrix.Columns.Item("V_-1").Cells.Item(oShpMatrix.RowCount).Specific.Value & "And DocEntry= " & oShpForm.Items.Item("ed_DocNum").Specific.Value)
                                    If oRecordSet.RecordCount > 0 Then
                                        oShpForm.Items.Item("ed_Box").Specific.Value = oRecordSet.Fields.Item("Box").Value
                                        oShpForm.Items.Item("ed_Part").Specific.Active = True
                                    End If
                                End If
                            Else
                                oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
                                oShpForm.Items.Item("ed_Box").Specific.Value = 1
                                oShpForm.Items.Item("ed_Part").Specific.Active = True
                            End If

                        End If

                        'End If
                    End If

                Case "EditCPO"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If

                            If ActiveMatrix = "mx_Orider" Then
                                RPOsrfname = "OutPurchaseOrder.srf"
                            End If

                            oMatrix = ImportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                            'MSW 14-09-2011 Truck PO
                            If ActiveMatrix = "mx_TkrList" Then
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadTruckingPO(ImportSeaLCLForm, RPOsrfname)
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oMatrix = ImportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                    CPOForm = p_oSBOApplication.Forms.ActiveForm
                                    CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                                    CPOForm.Items.Item("1").Click()
                                    CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                    p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                End If
                                'MSW 14-09-2011 Truck PO
                            Else
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadAndCreateCPO(ImportSeaLCLForm, RPOsrfname)
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oMatrix = ImportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                    CPOForm = p_oSBOApplication.Forms.ActiveForm
                                    CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                                    CPOForm.Items.Item("1").Click()
                                    CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                End If
                            End If
                        Else
                            p_oSBOApplication.MessageBox("Need to select the Row that you want to Edit")
                        End If
                    End If

                Case "CopyToCGR"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then

                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If

                            If ActiveMatrix = "mx_Orider" Then
                                RGRsrfname = "OutriderGoodReceipt.srf"

                            End If

                            oMatrix = ImportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                            'Truck PO
                            If ActiveMatrix = "mx_TkrList" Then
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadAndCreateCGR(ImportSeaLCLForm, RGRsrfname)
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    CGRForm = p_oSBOApplication.Forms.ActiveForm
                                    If Not FillDataToGoodsReceipt(ImportSeaLCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                    p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                End If
                                'Truck PO
                            Else
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadAndCreateCGR(ImportSeaLCLForm, RGRsrfname)
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    CGRForm = p_oSBOApplication.Forms.ActiveForm
                                    If Not FillDataToGoodsReceipt(ImportSeaLCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                End If
                            End If


                        Else
                            p_oSBOApplication.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                        End If
                    End If
                Case "CancelPO"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            If AlreadyExist("IMPORTSEALCL") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                            ElseIf AlreadyExist("IMPORTAIR") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                            ElseIf AlreadyExist("IMPORTLAND") Then
                                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                            End If
                            oMatrix = ImportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                            If ActiveMatrix = "mx_TkrList" Then
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                      "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                      "from  [@OBT_TB08_FFCPO] a inner join [@OBT_LCL03_TRUCKING] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value)
                                    If Not CancelPO(ImportSeaLCLForm, sql, ActiveMatrix) Then Throw New ArgumentException(sErrDesc)
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    ImportSeaLCLForm.Items.Item("1").Click()
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                    p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                End If
                            Else
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                        "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value)
                                    If Not CancelPO(ImportSeaLCLForm, sql, ActiveMatrix) Then Throw New ArgumentException(sErrDesc)
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    ImportSeaLCLForm.Items.Item("1").Click()
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                    p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                End If
                            End If
                        End If
                    End If
                Case ""

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

        ' **********************************************************************************
        '   Function    :   RowAddToMatrix()
        '   Purpose     :   This function will be providing to proceed Add row  for
        '                   Purechase order matrix item value for Purchase Order process
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix    '                          
        '   Return      :   No
        '                   
        ' **********************************************************************************
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

        ' **********************************************************************************
        '   Function    :    AddGSTComboData()
        '   Purpose     :    This function will be providing to add GST Amount  for
        '                    Purechase order item value for Purchase Order process                   
        '   Parameters  :    ByVal oColumn As SAPbouiCOM.Column                       
        '   Return      :    No
        '                   
        ' **********************************************************************************


        Dim RS As SAPbobsCOM.Recordset

        RS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery("select Code,Name from ovtg Where Category='I'")
        RS.MoveFirst()
        For j As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
            oColumn.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        ' oColumn.ValidValues.Add("None", "None") MSW To Edit New Ticket
        While RS.EoF = False
            oColumn.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
            RS.MoveNext()
        End While
    End Sub
   
    Private Sub AddDataToDataTable(ByVal oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)

        ' **********************************************************************************
        '   Function    :   AddDataToDataTable()
        '   Purpose     :   This function will be providing add data to datatable  for
        '                   Purechase order item value for Purchase Order process
        '               
        '   Parameters  :  ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer        '                          
        '   Return      :   No
        '                   
        ' **********************************************************************************

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

        '=============================================================================
        'Function   : CalRatePO()
        'Purpose    : This function provide for calculate Rate for purchae order itmes
        'Parameters : ByVal oActiveForm As SAPbouiCOM.Form,
        '             ByVal Row As Integer
        'Return     : No
        '=============================================================================

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
            'NOGST = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) - GSTAMT
            'NOGST = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) 'MSW TO Edit New Ticket
            NOGST = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) + GSTAMT 'MSW TO Edit New Ticket
        End If

        '====CalCulate Total======'
        oActiveForm.Freeze(True)
        oMatrix.Columns.Item("colGSTAmt").Editable = True
        oMatrix.Columns.Item("colNoGST").Editable = True
        oMatrix.Columns.Item("colGSTAmt").Cells.Item(Row).Specific.Value = Convert.ToString(GSTAMT)
        oMatrix.Columns.Item("colNoGST").Cells.Item(Row).Specific.Value = Convert.ToString(NOGST)
        'p_oSBOApplication.Forms.Item("VOUCHER").Select()
        'Try
        '    oActiveForm.Items.Item("ed_VedCode").Specific.Active = True
        'Catch ex As Exception

        'End Try
        'MSW to edit
        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If Row > 1 And oMatrix.Columns.Item("colChCode1").Cells.Item(Row).Specific.Value = "" Then
                oMatrix.Columns.Item("colChCode1").Cells.Item(Row - 1).Click()
            Else
                oMatrix.Columns.Item("colChCode1").Cells.Item(Row).Click()
            End If
        Else
            oActiveForm.Items.Item("ed_InvNo").Specific.Active = True
        End If

        oMatrix.Columns.Item("colGSTAmt").Editable = False
        oMatrix.Columns.Item("colNoGST").Editable = False

        CalculateTotal(oActiveForm, oMatrix)
        oActiveForm.Freeze(False)

    End Sub

    Private Sub CalculateTotal(ByRef oActiveForm As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix)

        ' **********************************************************************************
        '   Function    :   CalculateTotal()
        '   Purpose     :   This function will be providing to proceed calculate  for
        '                    Purechase order item value for Purchase Order process
        '               
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form,
        '                   ByRef oMatrix As SAPbouiCOM.Matrix                  
        '   Return      :   No
        '                   
        ' **********************************************************************************

        Dim SubTotal As Double = 0.0
        Dim GSTTotal As Double = 0.0
        Dim Total As Double = 0.0
        For i As Integer = 1 To oMatrix.RowCount
            'MSW To Edit New Ticket
            SubTotal = SubTotal + (Convert.ToDouble(oMatrix.Columns.Item("colNoGST").Cells.Item(i).Specific.Value) - Convert.ToDouble(oMatrix.Columns.Item("colGSTAmt").Cells.Item(i).Specific.Value))
            GSTTotal = GSTTotal + Convert.ToDouble(oMatrix.Columns.Item("colGSTAmt").Cells.Item(i).Specific.Value)
            'End MSW To Edit New Ticket
        Next
        Total = SubTotal + GSTTotal
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_SubTotal", 0, SubTotal)
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_GSTAmt", 0, GSTTotal)
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_Total", 0, Total)
    End Sub
    Private Sub AddUpdateVoucher(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)

        ' **********************************************************************************
        '   Function    :   AddUpdateVoucher()
        '   Purpose     :   This function will be providing to Add and Update to Voucher form  
        '                   of ImportseaLcl Form
        '               
        '   Parameters  :  ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :  ByVal DataSource As String, ByVal ProcressedState As Boolean          
        '   Return      :  No
        ' **********************************************************************************
        Dim oActiveForm As SAPbouiCOM.Form
        If AlreadyExist("IMPORTSEALCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
        ElseIf AlreadyExist("IMPORTAIR") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
        ElseIf AlreadyExist("IMPORTLAND") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
        End If
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
                    '.SetValue("U_PayToAdd", .Offset, pForm.Items.Item("ed_PayTo").Specific.Value)
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
                    '.SetValue("U_PayToAdd", .Offset, pForm.Items.Item("ed_PayTo").Specific.Value)
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
    Private Sub SetVoucherDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   SetVoucherDataToEditTabByIndex()
        '   Purpose     :   This function will be providing to proceed  for to get data from maxtrix 
        '                   of ImportseaLcl Form
        '               
        '   Parameters  :  ByVal pForm As SAPbouiCOM.Form        '                         
        '   Return      :  No 
        ' **********************************************************************************
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


            'pForm.Items.Item("ed_PayTo").Specific.Value = PayTo
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

        ' **********************************************************************************
        '   Function    :   GetVoucherDataFromMatrixByIndex()
        '   Purpose     :   This function will be providing to proceed  for to get data from maxtrix 
        '                   of ImportseaLcl Form
        '               
        '   Parameters  :  ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix
        '                  ByVal Index As Integer                     
        '   Return      :  No
        '                   
        ' **********************************************************************************
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
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CPOMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oChMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing
        Dim CGRMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoImportSeaLCLItemEvent()"
        Dim sql As String = ""
        Dim strDsp As String = String.Empty

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)

            Select Case pVal.FormTypeEx
                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000051", "2000000039"  'MSW 14-09-2011 Truck PO   ' CGR --> Custom Goods Receipt
                    CGRForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        CGRForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.BeforeAction = False Then

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "mx_Item" And pVal.ColUID = "LineId" Then
                                CGRForm.EnableMenu("1292", True)
                                CGRForm.EnableMenu("1293", True)
                                CGRForm.EnableMenu("1294", False)
                                CGRForm.EnableMenu("1283", False)
                                CGRForm.EnableMenu("1284", False)
                                CGRForm.EnableMenu("1286", False)
                                CGRForm.EnableMenu("772", False)
                                CGRForm.EnableMenu("773", False)
                                CGRForm.EnableMenu("775", False)
                            Else
                                CGRForm.EnableMenu("1292", False)
                                CGRForm.EnableMenu("1293", False)
                                CGRForm.EnableMenu("1294", False)
                                CGRForm.EnableMenu("1283", False)
                                CGRForm.EnableMenu("1284", False)
                                CGRForm.EnableMenu("1286", False)
                                CGRForm.EnableMenu("772", True)
                                CGRForm.EnableMenu("773", True)
                                CGRForm.EnableMenu("775", True)
                            End If
                        End If

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
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            If pVal.ColUID = "colIQty" Or pVal.ColUID = "colIPrice" Then
                                CalAmtPO(CGRForm, pVal.Row)
                                CGRMatrix = CGRForm.Items.Item("mx_Item").Specific
                                If CGRMatrix.Columns.Item("colIGST").Cells.Item(CGRMatrix.RowCount).Specific.Value <> "" Then
                                    CalRatePO(CGRForm, pVal.Row)
                                End If
                            End If
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


                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                            CGRMatrix = CGRForm.Items.Item("mx_Item").Specific
                            If CGRMatrix.GetNextSelectedRow <> -1 Then
                                If pVal.ColUID = "colIQty" Then
                                    Try
                                        If CGRMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value < 0 Then
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            'CGRMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        End If
                                    Catch ex As Exception
                                        If CGRMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "" Then
                                            'CGRMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        Else
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        End If
                                    End Try
                                End If
                                If pVal.ColUID = "colIPrice" Then
                                    Try
                                        If CGRMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value < 0 Then
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ' CGRMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value = "0"
                                        End If
                                    Catch ex As Exception
                                        If CGRMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value = "" Then
                                            'CGRMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value = "0"
                                        Else
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        End If
                                    End Try
                                End If
                                If pVal.ColUID = "colIAmt" Then
                                    Try
                                        If CGRMatrix.Columns.Item("colIAmt").Cells.Item(pVal.Row).Specific.Value < 0 Then
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            'CGRMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        End If
                                    Catch ex As Exception
                                        If CGRMatrix.Columns.Item("colIAmt").Cells.Item(pVal.Row).Specific.Value = "" Then
                                            'CGRMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        Else
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        End If
                                    End Try
                                End If
                            End If


                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "ed_TDate" And CGRForm.Items.Item("ed_TDate").Specific.String <> String.Empty Then
                                If Not DateTime(CGRForm, _
                                                CGRForm.Items.Item("ed_TDate").Specific, _
                                                CGRForm.Items.Item("ed_TDay").Specific, _
                                                CGRForm.Items.Item("ed_TTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                ' CGRForm.Items.Item("ed_TPlace").Specific.Active = True
                            End If


                            'MSW 
                            If pVal.ItemUID = "1" Then
                                If AlreadyExist("IMPORTSEALCL") Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                ElseIf AlreadyExist("IMPORTAIR") Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                                ElseIf AlreadyExist("IMPORTLAND") Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                                End If
                                ' When click the Add button in AddMode of Custom Purchase Order form
                                ' need to also trigger the item pressed event of Main Export Form according by Customize Biz Logic
                                If pVal.Action_Success = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        CGRForm = p_oSBOApplication.Forms.ActiveForm
                                        CGRForm.Close()
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        ImportSeaLCLForm.Items.Item("1").Click()
                                        ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_CGR").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                        ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011

                                    End If
                                End If
                            End If

                            'If pVal.ItemUID = "1" Then
                            '    If pVal.BeforeAction = True Then
                            '        If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                            '            'CGRForm = p_oSBOApplication.Forms.ActiveForm
                            '            'If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            '            '    If Not CreateGoodsReceiptPO(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            '            'End If
                            '        End If
                            '    End If
                            'End If
                        End If

                    End If

                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000050", "2000000038"  ''MSW To Edit New Ticket     ' CPO --> Custom Purchase Order"
                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If AlreadyExist("IMPORTSEALCL") Then
                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                    ElseIf AlreadyExist("IMPORTAIR") Then
                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                    ElseIf AlreadyExist("IMPORTLAND") Then
                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                    End If
                    If pVal.BeforeAction = False Then

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "mx_Item" And pVal.ColUID = "LineId" Then
                                CPOForm.EnableMenu("1292", True)
                                CPOForm.EnableMenu("1293", True)
                                CPOForm.EnableMenu("1294", False)
                                CPOForm.EnableMenu("1283", False)
                                CPOForm.EnableMenu("1284", False)
                                CPOForm.EnableMenu("1286", False)
                                CPOForm.EnableMenu("772", False)
                                CPOForm.EnableMenu("773", False)
                                CPOForm.EnableMenu("775", False)
                            Else
                                CPOForm.EnableMenu("1292", False)
                                CPOForm.EnableMenu("1293", False)
                                CPOForm.EnableMenu("1294", False)
                                CPOForm.EnableMenu("1283", False)
                                CPOForm.EnableMenu("1284", False)
                                CPOForm.EnableMenu("1286", False)
                                CPOForm.EnableMenu("772", True)
                                CPOForm.EnableMenu("773", True)
                                CPOForm.EnableMenu("775", True)
                            End If
                        End If

                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            CPOForm.Items.Item("ed_Code").Enabled = False
                            CPOForm.Items.Item("ed_Name").Enabled = False
                            CPOForm.Items.Item("cb_SInA").Enabled = True
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
                                    oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                            oRecordSet.MoveNext()
                                        End While
                                    End If
                                    ''MSW Comment for email not need to send mail
                                    'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    'oRecordSet.DoQuery("SELECT E_Mail FROM OCRD WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    'If oRecordSet.RecordCount > 0 Then
                                    '    oRecordSet.MoveFirst()
                                    '    Dim stremail As String = IIf(oRecordSet.Fields.Item("E_Mail").Value.ToString = vbNull.ToString, "", oRecordSet.Fields.Item("E_Mail").Value)
                                    '    Try
                                    '        Dim tempMail As String = CPOForm.Items.Item("ed_Email").Specific.Value
                                    '        Dim arrTemp() As String
                                    '        If tempmail.Length > 0 Then
                                    '            arrTemp = tempMail.Split(",")
                                    '            If arrTemp.Count > 1 Then
                                    '                arrTemp(0) = stremail
                                    '                arrTemp(1) = arrTemp(1).ToString
                                    '                stremail = arrTemp(0) & "," & arrTemp(1)
                                    '            End If
                                    '        End If
                                    '        CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_Email", 0, stremail)
                                    '        'CPOForm.Items.Item("ed_Email").Specific.Value = stremail
                                    '    Catch ex As Exception

                                    '    End Try

                                    'End If
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


                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                            CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                            If CPOMatrix.GetNextSelectedRow <> -1 Then
                                If pVal.ColUID = "colIQty" Then
                                    Try
                                        If CPOMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value < 0 Then
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            'CPOMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        End If
                                    Catch ex As Exception
                                        If CPOMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "" Then
                                            'CPOMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        Else
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        End If
                                    End Try
                                End If
                                If pVal.ColUID = "colIPrice" Then
                                    Try
                                        If CPOMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value < 0 Then
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ' CPOMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value = "0"
                                        End If
                                    Catch ex As Exception
                                        If CPOMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value = "" Then
                                            'CPOMatrix.Columns.Item("colIPrice").Cells.Item(pVal.Row).Specific.Value = "0"
                                        Else
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        End If
                                    End Try
                                End If
                                If pVal.ColUID = "colIAmt" Then
                                    Try
                                        If CPOMatrix.Columns.Item("colIAmt").Cells.Item(pVal.Row).Specific.Value < 0 Then
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            'CPOMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        End If
                                    Catch ex As Exception
                                        If CPOMatrix.Columns.Item("colIAmt").Cells.Item(pVal.Row).Specific.Value = "" Then
                                            'CPOMatrix.Columns.Item("colIQty").Cells.Item(pVal.Row).Specific.Value = "0"
                                        Else
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        End If
                                    End Try
                                End If
                            End If


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
                                CPOForm.Items.Item("ed_TDate").Specific.Active = True
                                'MSW Comment for email no need to send mail
                                'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                'oRecordSet.DoQuery("SELECT email FROM OHEM WHERE empID = " & strempId & "")
                                'If oRecordSet.RecordCount > 0 Then
                                '    oRecordSet.MoveFirst()

                                '    Dim stremail As String = IIf(oRecordSet.Fields.Item("email").Value = vbNull.ToString, "", oRecordSet.Fields.Item("email").Value)
                                '    Try
                                '        Dim tempMail As String = CPOForm.Items.Item("ed_Email").Specific.Value
                                '        Dim arrTemp() As String
                                '        If tempMail.Length > 0 Then
                                '            arrTemp = tempMail.Split(",")
                                '            If arrTemp.Count > 1 Then
                                '                arrTemp(0) = arrTemp(0).ToString
                                '                arrTemp(1) = stremail
                                '                stremail = arrTemp(0) & "," & arrTemp(1)
                                '                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_Email", 0, stremail)
                                '            Else
                                '                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_Email", 0, tempMail & "," & stremail)
                                '            End If
                                '        Else

                                '            CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_Email", 0, stremail)
                                '        End If

                                '        'CPOForm.Items.Item("ed_Email").Specific.Value = stremail
                                '    Catch ex As Exception

                                '    End Try

                                '    'If CPOForm.Items.Item("ed_Email").Specific.Value <> "" Then
                                '    '    CPOForm.Items.Item("ed_Email").Specific.Value = CPOForm.Items.Item("ed_Email").Specific.Value + "," + IIf(oRecordSet.Fields.Item("email").Value = vbNull.ToString, "", oRecordSet.Fields.Item("email").Value)
                                '    'Else
                                '    '    CPOForm.Items.Item("ed_Email").Specific.Value = IIf(oRecordSet.Fields.Item("email").Value = vbNull.ToString, "", oRecordSet.Fields.Item("email").Value)
                                '    'End If

                                'End If

                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "ed_TDate" And CPOForm.Items.Item("ed_TDate").Specific.String <> String.Empty Then
                                If Not DateTime(CPOForm, _
                                                CPOForm.Items.Item("ed_TDate").Specific, _
                                                CPOForm.Items.Item("ed_TDay").Specific, _
                                                CPOForm.Items.Item("ed_TTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                ' CPOForm.Items.Item("ed_TPlace").Specific.Active = True
                            End If


                            If pVal.ItemUID = "1" Then
                                If AlreadyExist("IMPORTSEALCL") Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                ElseIf AlreadyExist("IMPORTAIR") Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                                ElseIf AlreadyExist("IMPORTLAND") Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                                End If
                                ' When click the Add button in AddMode of Custom Purchase Order form
                                ' need to also trigger the item pressed event of Main Export Form according by Customize Biz Logic
                                If pVal.Action_Success = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        CPOForm = p_oSBOApplication.Forms.ActiveForm
                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            'when retrieve spefific data for update, add new row in the matrix
                                            If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                            If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                                CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            CPOForm.Items.Item("ed_Code").Enabled = False
                                            CPOForm.Items.Item("ed_Name").Enabled = False
                                            CPOForm.Items.Item("cb_SInA").Enabled = True
                                            CPOForm.Items.Item("bt_Preview").Visible = True
                                            CPOForm.Items.Item("bt_Resend").Visible = False
                                        End If
                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            CPOForm.Close()
                                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            ImportSeaLCLForm.Items.Item("1").Click()
                                            'MSW 14-09-2011 Truck PO
                                            If pVal.FormTypeEx = "2000000050" Then
                                                EnabledTruckerForExternal(ImportSeaLCLForm, False)
                                                ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                                            End If
                                            'MSW 14-09-2011 Truck PO
                                        End If
                                    End If
                                    ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_CGR").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                    ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                                    If (pVal.FormTypeEx = "IMPORTAIR" Or pVal.FormTypeEx = "IMPORTLAND") Then
                                        ImportSeaLCLForm.Items.Item("bt_Orider").Enabled = True
                                    End If
                                End If
                            End If
                            If pVal.ItemUID = "bt_Preview" Then
                                PreviewPO(ImportSeaLCLForm, CPOForm)
                            End If
                            If pVal.ItemUID = "bt_Resend" Then
                                SendAttachFile(ImportSeaLCLForm, CPOForm)
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
                                    If AlreadyExist("IMPORTSEALCL") Then
                                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                    ElseIf AlreadyExist("IMPORTAIR") Then
                                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                                    ElseIf AlreadyExist("IMPORTLAND") Then
                                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                                    End If
                                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                                    If Not CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                                        If CheckQtyValue(CPOMatrix, "Purchase Order") = True Then
                                            p_oSBOApplication.SetStatusBarMessage("Document Total is Zero.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            BubbleEvent = False
                                        ElseIf CPOMatrix.RowCount = 1 And CPOMatrix.Columns.Item("colItemNo").Cells.Item(CPOMatrix.RowCount).Specific.Value = "" Then
                                            'msw
                                            p_oSBOApplication.SetStatusBarMessage("Document Total is Zero.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            BubbleEvent = False
                                        ElseIf CPOMatrix.Columns.Item("colItemNo").Cells.Item(CPOMatrix.RowCount).Specific.Value = "" Then
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
                    Try
                        oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    Dim m3 As Decimal
             
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            If pVal.ItemUID = "ed_ShDate" Then
                                If HolidayMarkUp(oShpForm, oShpForm.Items.Item("ed_ShDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                If HolidayMarkUp(oShpForm, oShpForm.Items.Item("ed_ShDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            
                            If pVal.ItemUID = "1" Then
                                If pVal.ActionSuccess = True Then
                                    ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                    oShpForm.Items.Item("bt_PPView").Visible = True
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Then

                                        If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB03_EXPSHPINV] set " & _
                                            " U_FrDocNo=" & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oShpForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            oShpForm.Close()

                                            ImportSeaLCLForm.Items.Item("1").Click()
                                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If

                                        ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            oShpForm.Items.Item("ed_ShipTo").Enabled = False
                                            oShpForm.Items.Item("bt_PPView").Visible = True
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                    ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                End If
                            End If
                            If pVal.ItemUID = "bt_PPView" Then

                                'Save Report To specific JobFile as PDF File USE Code
                                Dim reportuti As New clsReportUtilities
                                Dim pdfFilename As String = "SHIPPING INV"
                                Dim mainFolder As String = p_fmsSetting.DocuPath
                                Dim jobNo As String = oShpForm.Items.Item("ed_Job").Specific.Value
                                Dim rptPath As String = Application.StartupPath.ToString & "\Shipping Invoice.rpt"
                                Dim pdffilepath As String = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                Dim rptDocument As ReportDocument = New ReportDocument()
                                rptDocument.Load(rptPath)
                                rptDocument.Refresh()
                                rptDocument.SetParameterValue("@DocEntry", oShpForm.Items.Item("ed_DocNum").Specific.Value)
                                reportuti.SetDBLogIn(rptDocument)
                                If Not pdffilepath = String.Empty Then
                                    reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
                                End If
                                'reportuti.PrintDoc(rptDocument) 'To Print To Printer

                                rptDocument.Close()
                                'End Save Report To specific JobFile as PDF File

                                'Dim reportuti As New ReportUtilities
                                'Dim rptPath As String = IO.Directory.GetParent(Application.StartupPath).ToString & "\Shipping Invoice.rpt"
                                'Dim pdfPath As String = IO.Directory.GetParent(Application.StartupPath).ToString & "\Shipping Invoice.pdf"


                                ''Dim rptPath As String = "D:\LatestFCL\April\modFreightImportSeaFCL\modFreightImportSeaFCL\bin\test.rpt"
                                ''Dim pdfpath As String = "D:\LatestFCL\April\modFreightImportSeaFCL\modFreightImportSeaFCL\bin\test.pdf"
                                'Dim rptDocument As ReportDocument = New ReportDocument()

                                'rptDocument.Load(rptPath)
                                'rptDocument.Refresh()
                                'rptDocument.SetParameterValue("@DocEntry", oShpForm.Items.Item("ed_DocNum").Specific.Value)
                                'reportuti.SetDBLogIn(rptDocument)

                                'Dim crExportOptions As New CrystalDecisions.Shared.ExportOptions
                                'Dim crDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
                                'crDiskFileDestinationOptions.DiskFileName = pdfPath
                                'crExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                                'crExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                                'crExportOptions.ExportDestinationOptions = crDiskFileDestinationOptions
                                'rptDocument.Export(crExportOptions)


                            End If



                            If pVal.ItemUID = "bt_Add" Then
                                oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                                If ValidateforformShippingInv(oShpForm) = True Then
                                    Exit Function
                                Else
                                    If oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD" Then
                                        AddUpdateShippingInv(oShpForm, oShpMatrix, "@OBT_TB03_EXPSHPINVD", True)
                                        oShpForm.Items.Item("bt_Add").Specific.Caption = "Edit"
                                        CalculateNoOfBoxes(oShpForm, oShpMatrix)
                                    Else
                                        AddUpdateShippingInv(oShpForm, oShpMatrix, "@OBT_TB03_EXPSHPINVD", False)
                                        CalculateNoOfBoxes(oShpForm, oShpMatrix)
                                        'oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD" 'MSW To Edit
                                    End If

                                End If
                            End If
                            If pVal.ItemUID = "bt_Clear" Then
                                oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                                ClearText(oShpForm, "ed_PartDes", "ed_Qty", "ed_Unit", "ed_Box", "ed_BoxLast", "ed_L", "ed_B", "ed_H", "ed_M3", "ed_M3T", "ed_Net", "ed_NetT", "ed_Gross", "ed_GrossT", "ed_Nec", "ed_NecT", "ed_TotBox", "ed_Boxes", "ed_PPBNo", "ed_PUnit", "ed_UnPrice", "ed_TotVal", "ed_ShName", "ed_PShName", "ed_ECCN", "ed_License", "ed_LExDate", "ed_Class", "ed_UN", "ed_HSCode", "ed_DOM", "ed_Part")
                                'MSW To Edit
                                oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD" 'MSW To Edit
                                If (oShpMatrix.RowCount > 0) Then
                                    If (oShpMatrix.Columns.Item("V_-1").Cells.Item(oShpMatrix.RowCount).Specific.Value = Nothing) Then
                                        oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
                                        oShpForm.Items.Item("ed_Box").Specific.Value = 1
                                        oShpForm.Items.Item("ed_Part").Specific.Active = True
                                    Else
                                        oShpForm.Items.Item("ed_Box").Specific.Value = oShpMatrix.Columns.Item("colBoxLast").Cells.Item(oShpMatrix.RowCount).Specific.Value + 1
                                        oShpForm.Items.Item("ed_ItemNo").Specific.Value = oShpMatrix.Columns.Item("V_-1").Cells.Item(oShpMatrix.RowCount).Specific.Value + 1
                                        oShpForm.Items.Item("ed_Part").Specific.Active = True
                                    End If
                                Else
                                    oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
                                    oShpForm.Items.Item("ed_Box").Specific.Value = 1
                                    oShpForm.Items.Item("ed_Part").Specific.Active = True
                                End If
                                'End MSW To Edit

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
                                    'oShpForm.DataSources.UserDataSources.Item("Qty").Value = IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString) 'MSW To Edit 31-08-2011
                                    oShpForm.DataSources.UserDataSources.Item("Unit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString)

                                    oShpForm.DataSources.UserDataSources.Item("DL").Value = IIf(oDataTable.Columns.Item("U_Length").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Length").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("DB").Value = IIf(oDataTable.Columns.Item("U_Base").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Base").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("DH").Value = IIf(oDataTable.Columns.Item("U_Height").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Height").Cells.Item(0).Value.ToString)

                                    m3 = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Length").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Length").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Base").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Base").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Height").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Height").Cells.Item(0).Value)) / 1000000
                                    oShpForm.DataSources.UserDataSources.Item("M3").Value = m3
                                    ' oShpForm.DataSources.UserDataSources.Item("TM3").Value = m3 * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString))
                                    oShpForm.DataSources.UserDataSources.Item("NetKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value))
                                    ' oShpForm.DataSources.UserDataSources.Item("TotKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("GroKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value.ToString))
                                    ' oShpForm.DataSources.UserDataSources.Item("TotGKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    'MSW To Edit
                                    oShpForm.DataSources.UserDataSources.Item("NEC").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value.ToString)) 'MSW To Edit
                                    ' oShpForm.DataSources.UserDataSources.Item("TotNEC").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("PBox").Value = IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString) 'MSW To Edit 31-08-2011
                                    'oShpForm.DataSources.UserDataSources.Item("BUnit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("BUnit").Value = "Box" 'MSW 21-09-2011
                                    oShpForm.DataSources.UserDataSources.Item("PUnit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString & "/Box") 'MSW 21-09-2011
                                    'End MSW To Edit
                                    oShpForm.DataSources.UserDataSources.Item("UPrice").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value = vbNull, 0.0, oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value))
                                    'oShpForm.DataSources.UserDataSources.Item("TotV").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Rate").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("SName").Value = IIf(oDataTable.Columns.Item("U_Shipping").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_Shipping").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("Ecc").Value = IIf(oDataTable.Columns.Item("U_ECCN").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_ECCN").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("Lic").Value = IIf(oDataTable.Columns.Item("U_ExpLic").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_ExpLic").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("Cls").Value = IIf(oDataTable.Columns.Item("U_Class").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_Class").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("UN").Value = IIf(oDataTable.Columns.Item("U_UNNo").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UNNo").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("HSCode").Value = IIf(oDataTable.Columns.Item("U_HSCode").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_HSCode").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("LicExDate").Value = IIf(oDataTable.Columns.Item("U_LicDate").Cells.Item(0).Value = Nothing, " ", Convert.ToDateTime(oDataTable.Columns.Item("U_LicDate").Cells.Item(0).Value).ToString("yyyyMMdd"))
                                End If
                            Catch ex As Exception

                            End Try
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                            If pVal.ItemUID = "ed_Qty" Then
                                'MSW To Edit
                                If oShpForm.Items.Item("ed_Qty").Specific.value.ToString <> "" Then
                                    Dim qty As Double
                                    Dim noofPkg As Double
                                    Try
                                        qty = IIf(Convert.ToDouble(oShpForm.Items.Item("ed_Qty").Specific.value.ToString) = 0, 0, Convert.ToDouble(oShpForm.Items.Item("ed_Qty").Specific.value.ToString))
                                        noofPkg = qty / Convert.ToDouble(oShpForm.Items.Item("ed_PPBNo").Specific.value.ToString)
                                        ' m3 = oShpForm.Items.Item("ed_L").Specific.value * oShpForm.Items.Item("ed_B").Specific.value * oShpForm.Items.Item("ed_H").Specific.value / 1000000
                                        m3 = Convert.ToDouble(IIf(oShpForm.Items.Item("ed_L").Specific.Value.ToString = "", 0, oShpForm.Items.Item("ed_L").Specific.Value.ToString)) * Convert.ToDouble(IIf(oShpForm.Items.Item("ed_B").Specific.Value.ToString = "", 0, oShpForm.Items.Item("ed_B").Specific.Value.ToString)) * Convert.ToDouble(IIf(oShpForm.Items.Item("ed_H").Specific.Value.ToString = "", 0, oShpForm.Items.Item("ed_H").Specific.Value.ToString)) / 1000000
                                        oShpForm.DataSources.UserDataSources.Item("M3").Value = m3
                                        oShpForm.DataSources.UserDataSources.Item("TM3").Value = m3 * noofPkg
                                        oShpForm.DataSources.UserDataSources.Item("TotBox").Value = noofPkg
                                        'MSW 21-09-2011
                                        If noofPkg > 1 Then
                                            oShpForm.DataSources.UserDataSources.Item("BUnit").Value = "Boxes"
                                        Else
                                            oShpForm.DataSources.UserDataSources.Item("BUnit").Value = "Box"
                                        End If
                                        'MSW 21-09-2011
                                        oShpForm.DataSources.UserDataSources.Item("NetKg").Value = oShpForm.Items.Item("ed_Net").Specific.value
                                        oShpForm.DataSources.UserDataSources.Item("TotKg").Value = oShpForm.Items.Item("ed_Net").Specific.value * noofPkg
                                        oShpForm.DataSources.UserDataSources.Item("GroKg").Value = oShpForm.Items.Item("ed_Gross").Specific.value
                                        oShpForm.DataSources.UserDataSources.Item("TotGKg").Value = oShpForm.Items.Item("ed_Gross").Specific.value * noofPkg
                                        oShpForm.DataSources.UserDataSources.Item("TotNEC").Value = oShpForm.Items.Item("ed_Nec").Specific.value * qty
                                        oShpForm.DataSources.UserDataSources.Item("TotV").Value = oShpForm.Items.Item("ed_UnPrice").Specific.value * qty
                                        oShpForm.DataSources.UserDataSources.Item("BoxLast").Value = (oShpForm.Items.Item("ed_Box").Specific.Value + noofPkg) - 1
                                        'oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
                                        ' CalculateNoOfBoxes(oShpForm, oShpMatrix)


                                    Catch ex As Exception
                                        p_oSBOApplication.SetStatusBarMessage("Qty Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    End Try

                                End If
                            End If

                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.InnerEvent = False Then
                            If pVal.ItemUID = "ed_Qty" Then
                                If oShpForm.Items.Item("ed_Qty").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_Qty").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_Qty").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Qty Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_Qty").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.ActiveItem = "ed_Qty"
                                        oShpForm.Items.Item("ed_Qty").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Qty Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_Qty").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If
                            End If
                            If pVal.ItemUID = "ed_TotVal" Then
                                If oShpForm.Items.Item("ed_TotVal").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_TotVal").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_TotVal").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Total Value  Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_TotVal").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.Items.Item("ed_TotVal").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Total Value Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_TotVal").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If

                            End If
                            If pVal.ItemUID = "ed_L" Then
                                If oShpForm.Items.Item("ed_L").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_L").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_L").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Dimension Base Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_L").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.Items.Item("ed_L").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Dimension Base Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_L").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If

                            End If
                            If pVal.ItemUID = "ed_B" Then
                                If oShpForm.Items.Item("ed_B").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_B").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_B").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Dimension Base Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_B").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.Items.Item("ed_B").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Dimension Base Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_B").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If

                            End If
                            If pVal.ItemUID = "ed_H" Then
                                If oShpForm.Items.Item("ed_H").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_H").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_H").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Dimension Height Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_H").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.Items.Item("ed_H").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Dimension Height Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_H").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If

                            End If
                            If pVal.ItemUID = "ed_TotBox" Then
                                If oShpForm.Items.Item("ed_TotBox").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_TotBox").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_TotBox").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Totoal Box  Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_TotBox").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.Items.Item("ed_TotBox").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Totoal Box  Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_TotBox").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If

                            End If

                            If pVal.ItemUID = "ed_PPBNo" Then
                                If oShpForm.Items.Item("ed_PPBNo").Specific.Value <> "" Then
                                    Try
                                        If oShpForm.Items.Item("ed_PPBNo").Specific.Value < 0 Then
                                            oShpForm.Items.Item("ed_PPBNo").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Packing Per Boxes Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            oShpForm.Items.Item("ed_PPBNo").Specific.Value = ""
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oShpForm.Items.Item("ed_PPBNo").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Packing Per Boxes Must be greater than 0", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        oShpForm.Items.Item("ed_PPBNo").Specific.Value = ""
                                        BubbleEvent = False
                                    End Try
                                End If
                            End If

                        End If
                    End If
                Case "VOUCHER"
                    If AlreadyExist("IMPORTAIR") Then
                        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                    ElseIf AlreadyExist("IMPORTSEALCL") Then
                        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                    ElseIf AlreadyExist("IMPORTLAND") Then
                        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                    End If

                    oPayForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    Try
                        oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Before_Action = False Then
                        If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "V_-1" Then
                            oPayForm.EnableMenu("1292", True)
                            oPayForm.EnableMenu("1293", True)
                            oPayForm.EnableMenu("1294", False)
                            oPayForm.EnableMenu("1283", False)
                            oPayForm.EnableMenu("1284", False)
                            oPayForm.EnableMenu("1286", False)
                            oPayForm.EnableMenu("772", False)
                            oPayForm.EnableMenu("773", False)
                            oPayForm.EnableMenu("775", False)
                        Else
                            oPayForm.EnableMenu("1292", False)
                            oPayForm.EnableMenu("1293", False)
                            oPayForm.EnableMenu("1294", False)
                            oPayForm.EnableMenu("1283", False)
                            oPayForm.EnableMenu("1284", False)
                            oPayForm.EnableMenu("1286", False)
                            oPayForm.EnableMenu("772", True)
                            oPayForm.EnableMenu("773", True)
                            oPayForm.EnableMenu("775", True)
                        End If
                    End If

                    If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                        Try
                            oPayForm.Items.Item("ed_VedName").Enabled = False
                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                            oPayForm.Items.Item("bt_PayView").Visible = True
                        Catch ex As Exception
                        End Try
                    End If
                    Dim strPayFor As String = String.Empty
                    If pVal.BeforeAction = True And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                        If pVal.ItemUID = "1" Then
                            Try
                                If AlreadyExist("IMPORTAIR") Then
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                                ElseIf AlreadyExist("IMPORTSEALCL") Then
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                ElseIf AlreadyExist("IMPORTLAND") Then
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                                End If
                                oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)

                                If Not oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                    If oChMatrix.RowCount = 0 Then
                                        p_oSBOApplication.SetStatusBarMessage("Document Total is Zero.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                    If CheckQtyValue(oChMatrix, "Payment Voucher") = True Then
                                        p_oSBOApplication.SetStatusBarMessage("Document Total is Zero.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf oChMatrix.RowCount = 1 And oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("Document Total is Zero.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If

                                    If oChMatrix.Columns.Item("colChCode1").Cells.Item(oChMatrix.RowCount).Specific.Value = "" And BubbleEvent = True Then
                                        oChMatrix.DeleteRow(oChMatrix.RowCount)
                                    End If
                                    vocTotal = Convert.ToDouble(oPayForm.Items.Item("ed_Total").Specific.Value)
                                    gstTotal = Convert.ToDouble(oPayForm.Items.Item("ed_GSTAmt").Specific.Value)
                                    If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If Not SaveToPurchaseVoucher(oPayForm, True) Then
                                            BubbleEvent = False
                                        End If
                                        If BubbleEvent = True Then
                                            SaveToDraftPurchaseVoucher(oPayForm)
                                        End If
                                    ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If Not UpdateDraftPurchaseVoucher(oPayForm) Then
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If

                            Catch ex As Exception
                                BubbleEvent = False
                                MessageBox.Show(ex.Message)
                            End Try
                        End If
                    End If
                    If pVal.BeforeAction = False Then

                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                            Try
                                oPayForm.Items.Item("ed_VedName").Enabled = False
                                'MSW
                                'oPayForm.Items.Item("ed_PayTo").Enabled = False
                                oPayForm.Items.Item("ed_VedCode").Enabled = False
                            Catch ex As Exception
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount1" Then
                                CalRate(oPayForm, pVal.Row)
                            End If
                            'MSW To Edit
                            If pVal.ItemUID = "ed_VedCode" Then
                                If vedCurCode = "##" Then
                                    oCombo = oPayForm.Items.Item("cb_PayCur").Specific
                                    oCombo.Select("SGD", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If

                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then

                            If pVal.ItemUID = "bt_PayView" Then
                                PreviewPaymentVoucher(oActiveForm, oPayForm)
                            End If
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
                            'MSW To Edit New Ticket
                            If pVal.ItemUID = "op_Cheq" Then
                                Dim oComboBank As SAPbouiCOM.ComboBox
                                oPayForm.Items.Item("cb_BnkName").Specific.Active = True
                                oComboBank = oPayForm.Items.Item("cb_BnkName").Specific
                                If oComboBank.ValidValues.Count = 0 Then
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("Select AliasName From DSC1") 'MSW To Edit New Ticket
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oComboBank.ValidValues.Add(oRecordSet.Fields.Item("AliasName").Value, "")
                                            oRecordSet.MoveNext()
                                        End While
                                        'MSW To Edit
                                        oComboBank.Select("MayBank", SAPbouiCOM.BoSearchKey.psk_ByValue)  'MSW To Edit New Ticket
                                        oRecordSet.DoQuery("Select UsrNumber1 From DSC1 Where AliasName='MayBank'") 'MSW To Edit New Ticket
                                        If oRecordSet.RecordCount > 0 Then
                                            oCombo = oPayForm.Items.Item("cb_PayCur").Specific
                                            If oRecordSet.Fields.Item("UsrNumber1").Value.ToString.Trim.Length > 3 Then
                                                oCombo.Select("SGD", SAPbouiCOM.BoSearchKey.psk_ByValue)  'MSW To Edit New Ticket
                                            Else
                                                oCombo.Select(oRecordSet.Fields.Item("UsrNumber1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)  'MSW To Edit New Ticket
                                            End If
                                        End If
                                    End If

                                End If
                                If Not oPayForm.Items.Item("ed_PayType").Specific.Value = "" Then
                                    oPayForm.Items.Item("ed_PayType").Specific.Value = "Cheque"
                                End If

                                oPayForm.Items.Item("cb_BnkName").Enabled = True
                                oPayForm.Items.Item("ed_Cheque").Enabled = True
                            End If
                            'End MSW To Edit New Ticket


                            If pVal.ItemUID = "1" Then
                                If pVal.ActionSuccess = True Then
                                    If AlreadyExist("IMPORTAIR") Then
                                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTAIR", 1)
                                    ElseIf AlreadyExist("IMPORTSEALCL") Then
                                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
                                    ElseIf AlreadyExist("IMPORTLAND") Then
                                        ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm("IMPORTLAND", 1)
                                    End If

                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            'If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
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
                                            'DisableChargeMatrix(oPayForm, oMatrix, False)
                                            'MSW
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            'oPayForm.Items.Item("ed_PayTo").Enabled = False
                                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                                            oPayForm.Items.Item("bt_PayView").Visible = True

                                        End If
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB031_VHEADER] set U_APInvNo=" & Convert.ToInt32(strAPInvNo) & ",U_OutPayNo=" & Convert.ToInt32(strOutPayNo) & "" & _
                                            " ,U_FrDocNo='" & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "' Where DocEntry = " & oPayForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            PreviewPaymentVoucher(ImportSeaLCLForm, oPayForm)
                                            oPayForm.Close()

                                            ImportSeaLCLForm.Items.Item("1").Click()
                                            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If

                                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            '  p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                                            oPayForm.Items.Item("bt_PayView").Visible = True
                                            'oPayForm.Items.Item("ed_PayTo").Enabled = False
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                    ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                End If
                            End If
                        End If



                        'If oCombo.ValidValues.Count = 0 Then
                        '    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        '    oRecordSet.DoQuery("SELECT CurrCode,CurrName FROM OCRN")
                        '    If oRecordSet.RecordCount > 0 Then
                        '        oRecordSet.MoveFirst()
                        '        While oRecordSet.EoF = False
                        '            oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value.ToString.Trim, oRecordSet.Fields.Item("CurrName").Value.ToString)
                        '            oRecordSet.MoveNext()
                        '        End While

                        '    End If
                        'End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            'MSW 
                            If pVal.ItemUID = "ed_VedCode" Then
                                ObjDBDataSource = oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER") 'MSW To Add 18-3-2011
                                oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                'oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                '                                                                       & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString()) 'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*

                                vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft
                                oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_BPCode", ObjDBDataSource.Offset, oDataTable.GetValue(0, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                'If oDataTable.GetValue("Currency", 0).ToString = "##" Then
                                vedCurCode = oDataTable.GetValue("Currency", 0).ToString
                                'End If
                            End If

                            If pVal.ColUID = "colChCode1" Then
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                'MSW to Edit New Ticket
                                Try
                                    oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    'oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_CCode").Cells.Item(0).Value.ToString
                                Catch ex As Exception

                                End Try
                                Try
                                    oChMatrix.Columns.Item("colVDesc1").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_CName").Cells.Item(0).Value.ToString
                                Catch ex As Exception

                                End Try
                                'End MSW to Edit New Ticket
                                Try
                                    oChMatrix.Columns.Item("colAcCode").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_PAccCode").Cells.Item(0).Value.ToString
                                Catch ex As Exception

                                End Try
                                Try
                                    oChMatrix.Columns.Item("colICode").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_ItemCode").Cells.Item(0).Value.ToString
                                Catch ex As Exception

                                End Try


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
                            ''MSW To Edit New Ticket
                            ''If (pVal.ItemUID = "cb_GST") And oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            ''    oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                            ''    oCombo = oChMatrix.Columns.Item("colGST1").Cells.Item(oChMatrix.RowCount).Specific
                            ''    oCombo.Select("None", SAPbouiCOM.BoSearchKey.psk_ByValue) 'MSW To Edit New Ticket
                            ''End If
                            ''End MSW To Edit New Ticket
                            '----------------------------------------------------------------------------------'
                            'MSW TO Edit New Ticket
                            If pVal.ItemUID = "cb_BnkName" Then
                                oRecordSet.DoQuery("Select UsrNumber1 From DSC1 Where AliasName='" & oPayForm.Items.Item("cb_BnkName").Specific.Value & "'") 'MSW To Edit New Ticket
                                If oRecordSet.RecordCount > 0 Then
                                    oCombo = oPayForm.Items.Item("cb_PayCur").Specific
                                    If oRecordSet.Fields.Item("UsrNumber1").Value.ToString.Trim.Length > 3 Then
                                        oCombo.Select("SGD", SAPbouiCOM.BoSearchKey.psk_ByValue)  'MSW To Edit New Ticket
                                    Else
                                        oCombo.Select(oRecordSet.Fields.Item("UsrNumber1").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)  'MSW To Edit New Ticket
                                    End If

                                End If
                            End If
                            'End MSW TO Edit New Ticket
                            'MSW TO Edit New Ticket
                            If pVal.ItemUID = "cb_PayFor" Then
                                If oPayForm.Items.Item("cb_PayFor").Specific.Value.ToString.Trim = "Define New" Then
                                    oCombo = oPayForm.Items.Item("cb_PayFor").Specific
                                    oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    p_oSBOApplication.ActivateMenuItem("47617")
                                End If
                            End If
                            'End MSW TO Edit New Ticket

                            If pVal.ItemUID = "cb_PayCur" Then
                                If oPayForm.Items.Item("cb_PayCur").Specific.Value.ToString.Trim <> "SGD" Then
                                    Dim Rate As String = String.Empty
                                    sql = "SELECT Rate FROM ORTT WHERE Currency = '" & oPayForm.Items.Item("cb_PayCur").Specific.Value & "' And DATENAME(YYYY,RateDate) = '" & _
                                            Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                            CInt(Today.ToString("dd"))
                                    oRecordSet.DoQuery(sql)
                                    If oRecordSet.RecordCount > 0 Then
                                        Rate = oRecordSet.Fields.Item("Rate").Value
                                    End If
                                    oPayForm.Items.Item("ed_PayRate").Visible = True
                                    ' oPayForm.Items.Item("ed_PayRate").Enabled = True
                                    oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_ExRate", 0, Rate.ToString)
                                Else
                                    oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_ExRate", 0, Nothing)
                                    'oPayForm.Items.Item("ed_PayRate").Enabled = False
                                    oPayForm.Items.Item("ed_PayRate").Visible = False
                                    'oPayForm.Items.Item("ed_PosDate").Specific.Active = True
                                End If
                            End If
                        End If
                    End If
                Case "Image"
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            sPicName = pVal.ItemUID
                            LoadPopUpForm(pVal.ItemUID)
                        End If
                    End If

                Case "IMPORTSEALCL", "IMPORTAIR", "LOCAL", "IMPORTLAND"

                    ImportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)

                    Try
                        ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                            LoadHolidayMarkUp(ImportSeaLCLForm)
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(ImportSeaLCLForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If
                    If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_CGR").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                        ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                        If (pVal.FormTypeEx = "IMPORTAIR" Or pVal.FormTypeEx = "IMPORTLAND") Then
                            ImportSeaLCLForm.Items.Item("bt_Orider").Enabled = True
                        End If
                    End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        If pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize" And pVal.Before_Action = True And pVal.Row <> 0 Then
                            oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
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

                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "ed_FCharge" And pVal.Before_Action = True Then
                            If ImportSeaLCLForm.Items.Item("ed_FCharge").Specific.value <> "" Then
                                Try
                                    If ImportSeaLCLForm.Items.Item("ed_FCharge").Specific.value < 0 Then
                                        ImportSeaLCLForm.Items.Item("ed_FCharge").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_FCharge").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ImportSeaLCLForm.Items.Item("ed_FCharge").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_FCharge").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End Try
                            End If

                        End If
                        If pVal.ItemUID = "ed_Percent" And pVal.Before_Action = True Then
                            If ImportSeaLCLForm.Items.Item("ed_Percent").Specific.value <> "" Then
                                Try
                                    If ImportSeaLCLForm.Items.Item("ed_Percent").Specific.value < 0 Then
                                        ImportSeaLCLForm.Items.Item("ed_Percent").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_Percent").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ImportSeaLCLForm.Items.Item("ed_Percent").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Percent").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End Try
                            End If

                        End If
                        If pVal.ItemUID = "ed_GWt" And pVal.Before_Action = True Then
                            If ImportSeaLCLForm.Items.Item("ed_GWt").Specific.value <> "" Then
                                Try
                                    If ImportSeaLCLForm.Items.Item("ed_GWt").Specific.value < 0 Then
                                        ImportSeaLCLForm.Items.Item("ed_GWt").Specific.Active = True
                                        ImportSeaLCLForm.Items.Item("ed_GWt").Specific.Value = ""
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ImportSeaLCLForm.Items.Item("ed_GWt").Specific.Active = True
                                    ImportSeaLCLForm.Items.Item("ed_GWt").Specific.Value = ""
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False

                                End Try
                            End If

                        End If
                    End If


                    'MSW for Job Type Table
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ItemUID = "ed_JobNo" Then
                                ValidateJobNumber(ImportSeaLCLForm, BubbleEvent)
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
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            If pVal.ItemUID = "ed_TotalM3" Or pVal.ItemUID = "ed_TotalWt" Then
                                ' If Convert.ToDouble(ImportSeaLCLForm.Items.Item("ed_TotalM3").Specific.Value) <> 0.0 Then
                                If Convert.ToDouble(ImportSeaLCLForm.Items.Item("ed_TotalM3").Specific.Value) > Convert.ToDouble(ImportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Value) Then
                                    ImportSeaLCLForm.Items.Item("ed_TotChWt").Specific.Value = ImportSeaLCLForm.Items.Item("ed_TotalM3").Specific.Value
                                Else
                                    ImportSeaLCLForm.Items.Item("ed_TotChWt").Specific.Value = ImportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Value
                                End If
                                '  ImportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Active = True
                                'End If

                            End If
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
                                        If ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                                            ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                        ElseIf ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                                            ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                                        Else
                                            ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                        End If
                                        'ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                                        'ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                    Case "fo_Trkng"
                                        ImportSeaLCLForm.PaneLevel = 6
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_TkrList"
                                        '#1008 17-09-2011
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
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_Voucher"
                                        '#1008 17-09-2011
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                                End If

                                            Else
                                                'ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
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
                                        ImportSeaLCLForm.Freeze(True)
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
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRDATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                                            'ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            'MSW 14-09-2011 Truck PO
                                            EnabledTruckerForExternal(ImportSeaLCLForm, True)
                                            'ImportSeaLCLForm.Items.Item("op_Inter").Enabled = True
                                            'ImportSeaLCLForm.Items.Item("op_Exter").Enabled = True
                                            'MSW 14-09-2011 Truck PO
                                            ImportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True

                                        End If

                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = "" Then
                                            ImportSeaLCLForm.Items.Item("ed_TkrTime").Specific.Value = Now.ToString("HH:mm")
                                            If (oMatrix.RowCount > 0) Then
                                                If (oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                    ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                                Else
                                                    ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                                End If
                                            Else
                                                ImportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = 1
                                            End If
                                            ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                        End If
                                        ImportSeaLCLForm.Freeze(False)
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

                                    Case "fo_ShpInv"
                                        ImportSeaLCLForm.PaneLevel = 25
                                        ImportSeaLCLForm.Items.Item("fo_ShView").Specific.Select()
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_ShpInv"
                                        '#1008 17-09-2011
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                        End If


                                    Case "fo_ShView"
                                        ImportSeaLCLForm.PaneLevel = 25
                                    Case "fo_BkVsl"
                                        ImportSeaLCLForm.PaneLevel = 26
                                    Case "fo_Crate"
                                        ImportSeaLCLForm.PaneLevel = 27
                                        ImportSeaLCLForm.Items.Item("fo_CtView").Specific.Select()
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Crate").Specific
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                        End If

                                    Case "fo_CtView"
                                        ImportSeaLCLForm.PaneLevel = 27
                                    Case "fo_Fumi"
                                        ImportSeaLCLForm.PaneLevel = 28

                                        ImportSeaLCLForm.Items.Item("fo_FmView").Specific.Select()

                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Fumi").Specific
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                        End If

                                    Case "fo_FmView"
                                        ImportSeaLCLForm.PaneLevel = 28
                                    Case "fo_OpBunk"
                                        ImportSeaLCLForm.PaneLevel = 29
                                        ImportSeaLCLForm.Items.Item("fo_BkView").Specific.Select()
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_Bunk"
                                        '#1008 17-09-2011
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                        End If

                                    Case "fo_ArmEs"
                                        ImportSeaLCLForm.PaneLevel = 30
                                        ImportSeaLCLForm.Items.Item("fo_ArView").Specific.Select()
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_Armed"
                                        '#1008 17-09-2011
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Armed").Specific
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                        End If

                                    Case "fo_Crane"
                                        ImportSeaLCLForm.PaneLevel = 31
                                        ImportSeaLCLForm.Items.Item("fo_CrView").Specific.Select()
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_Crane"
                                        '#1008 17-09-2011
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Crane").Specific
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                                        End If

                                    Case "fo_Fork"
                                        ImportSeaLCLForm.PaneLevel = 32
                                        ImportSeaLCLForm.Items.Item("fo_FkView").Specific.Select()
                                        '#1008 17-09-2011
                                        ImportSeaLCLForm.Settings.Enabled = True
                                        ImportSeaLCLForm.Settings.EnableRowFormat = True
                                        ImportSeaLCLForm.Settings.MatrixUID = "mx_Fork"
                                        '#1008 17-09-2011
                                        oMatrix = ImportSeaLCLForm.Items.Item("mx_Fork").Specific
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                                        End If

                                    Case "fo_FkView"
                                        ImportSeaLCLForm.PaneLevel = 32
                                    Case "fo_ArView"
                                        ImportSeaLCLForm.PaneLevel = 30
                                    Case "fo_CrView"
                                        ImportSeaLCLForm.PaneLevel = 31
                                    Case "fo_BkView"
                                        ImportSeaLCLForm.PaneLevel = 29
                                    Case "fo_Orider"
                                        ImportSeaLCLForm.PaneLevel = 33
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ImportSeaLCLForm.Items.Item("bt_Orider").Enabled = True
                                        End If
                                End Select

                                ' MSW New Button at Choose From List 22-03-2011
                                If p_oSBOApplication.Forms.ActiveForm.Title = "List of VESSEL" Then
                                    If pVal.ItemUID = "ed_Vessel" Then
                                        AddNewBtToCFLFrm(p_oSBOApplication.Forms.ActiveForm, "bt_VNew")
                                    End If
                                    If pVal.ItemUID = "ed_Voy" Then
                                        AddNewBtToCFLFrm(p_oSBOApplication.Forms.ActiveForm, "bt_VoNew")
                                    End If
                                End If
                                ' MSW New Button at Choose From List 22-03-2011

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
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL10_PLICINFO").RemoveRecord(0)   '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("OBT_LCL10_PLICINFO")                                '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
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
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL08_PCONTAINE").RemoveRecord(0)    '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
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
                                'MSW to edit 10-09-2011
                                If pVal.ItemUID = "bt_PrntDis" Then
                                    PreviewDispatchInstruction(ImportSeaLCLForm)
                                End If
                                If pVal.ItemUID = "ed_ETADate" And ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.String <> String.Empty Then
                                    Try
                                        If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETATime").Specific) Then Throw New ArgumentException(sErrDesc)
                                        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value
                                        ImportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETADay").Specific.Value
                                        ImportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETATime").Specific.Value
                                        'ImportSeaLCLForm.ActiveItem = "ed_CrgDsc"
                                    Catch ex As Exception
                                        'ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value
                                        'ImportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETADay").Specific.Value
                                        'ImportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETATime").Specific.Value
                                        'ImportSeaLCLForm.ActiveItem = "ed_CrgDsc"
                                    End Try

                                    ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ADate").Specific, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_LCDDate" And ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific, ImportSeaLCLForm.Items.Item("ed_LCDDay").Specific, ImportSeaLCLForm.Items.Item("ed_LCDHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_LCDDay").Specific, ImportSeaLCLForm.Items.Item("ed_LCDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_CnDate" And ImportSeaLCLForm.Items.Item("ed_CnDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_CgDate" And ImportSeaLCLForm.Items.Item("ed_CgDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_JbDate" And ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspDate" And ImportSeaLCLForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspCDte" And ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                                    If Not DateTime(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "bt_Payment" Then
                                    p_oSBOApplication.ActivateMenuItem("2818")
                                End If

                                If pVal.ItemUID = "op_Inter" Then
                                    ImportSeaLCLForm.Freeze(True)
                                    'MSW 14-09-2011 Truck PO
                                    If ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Value <> "" Then
                                        If Not CancelTruckingPurchaseOrder(Convert.ToInt32(ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Value)) Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    'MSW 14-09-2011 Truck PO
                                    ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                    ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_PODocNo").Specific.Value = ""
                                    'MSW 14-09-2011 Truck PO
                                    ImportSeaLCLForm.Items.Item("ed_VehicNo").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_Attent").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = ""
                                    ImportSeaLCLForm.Items.Item("ed_EUC").Specific.Value = ""
                                    EnabledTruckerForExternal(ImportSeaLCLForm, True)
                                    'MSW 14-09-2011 Truck PO
                                    If AddChooseFromListByOption(ImportSeaLCLForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                    ImportSeaLCLForm.Freeze(False)
                                ElseIf pVal.ItemUID = "op_Exter" Then
                                    ImportSeaLCLForm.Freeze(True)
                                    If ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = "" Then 'MSW 14-09-2011 Truck PO
                                        ImportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                        'MSW 14-09-2011 Truck PO
                                        ImportSeaLCLForm.Items.Item("ed_VehicNo").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_Attent").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = ""
                                        ImportSeaLCLForm.Items.Item("ed_EUC").Specific.Value = ""
                                        EnabledTruckerForExternal(ImportSeaLCLForm, False)
                                        'MSW 14-09-2011 Truck PO
                                        ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = True
                                        If AddChooseFromListByOption(ImportSeaLCLForm, False, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    'ImportSeaLCLForm.Items.Item("ed_PONo").Specific.Active = True
                                    ImportSeaLCLForm.Freeze(False)
                                End If

                                If pVal.ItemUID = "op_DspIntr" Then
                                    If ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                                        strDsp = ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific.Value
                                    End If
                                    ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal"
                                    oCombo = ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If Not ClearComboData(ImportSeaLCLForm, "cb_Dspchr", "@OBT_LCL04_DISPATCH", "U_Dispatch") Then Throw New ArgumentException(sErrDesc)
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
                                        If strDsp <> "" Then
                                            oCombo.Select(strDsp.Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If
                                        strDsp = String.Empty
                                    End If


                                ElseIf pVal.ItemUID = "op_DspExtr" Then
                                    If ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                                        strDsp = ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific.Value
                                    End If
                                    ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External"
                                    oCombo = ImportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If Not ClearComboData(ImportSeaLCLForm, "cb_Dspchr", "@OBT_LCL04_DISPATCH", "U_Dispatch") Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT CardCode,CardName FROM OCRD WHERE CardType = 'S'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        Do Until oRecordSet.EoF
                                            oCombo.ValidValues.Add(oRecordSet.Fields.Item("CardName").Value.ToString, "")
                                            oRecordSet.MoveNext()
                                        Loop
                                        If strDsp <> "" Then
                                            oCombo.Select(strDsp.Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If
                                        strDsp = String.Empty
                                    End If

                                End If

                                If pVal.ItemUID = "bt_GenPO" Then

                                    'Purchase Order and Goods Receipt POP UP Truck PO

                                    ' ==================================== Creating Custom Purchase Order ==============================
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close"
                                        End If
                                        LoadTruckingPO(ImportSeaLCLForm, "TkrListPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                    ' ==================================== Creating Custom Purchase Order ==============================



                                    ''p_oSBOApplication.Menus.Item("2305").Activate()
                                    ' ''p_oSBOApplication.Menus.Item("6913").Activate()
                                    ''p_oSBOApplication.ActivateMenuItem("6913") 'MSW 04-04-2011
                                    ''p_oSBOApplication.Menus.Item("6913").Activate()

                                    ''Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                                    ''UDFAttachForm.Items.Item("U_JobNo").Specific.Value = ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    ''UDFAttachForm.Items.Item("U_InsDate").Specific.Value = ImportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value
                                End If

                                If pVal.ItemUID = "1" Then
                                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If pVal.ActionSuccess = True Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            ImportSeaLCLForm.Items.Item("ed_Code").Enabled = False 'MSW
                                            ImportSeaLCLForm.Items.Item("ed_Name").Enabled = False
                                            ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False 'MSW For Job Type Table
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_CGR").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                            ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                                            If (pVal.FormTypeEx = "IMPORTAIR" Or pVal.FormTypeEx = "IMPORTLAND") Then
                                                ImportSeaLCLForm.Items.Item("bt_Orider").Enabled = True
                                            End If
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                    
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
                                            ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False 'MSW to edit New Ticket 07-09-2011
                                        End If
                                        ClearText(ImportSeaLCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                                        ImportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        ImportSeaLCLForm.Items.Item("1").Click() 'MSW to edit New Ticket 07-09-2011
                                        ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                    End If
                                End If
                                'KM to edit
                                If pVal.ItemUID = "bt_PrntIns" Then
                                    PreviewInsDoc(ImportSeaLCLForm)
                                End If

                                If pVal.ItemUID = "bt_A6Label" Then
                                    PreviewA6Label(ImportSeaLCLForm)
                                End If
                                If pVal.ItemUID = "bt_DelIns" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    modTrucking.DeleteByIndex(ImportSeaLCLForm, oMatrix, "@OBT_LCL03_TRUCKING")                             '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                    ClearText(ImportSeaLCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                                    ImportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                End If

                                If pVal.ItemUID = "bt_AmdIns" Then

                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    'modTrucking.SetDataToEditTabByIndex(ImportSeaLCLForm)

                                    If (oMatrix.GetNextSelectedRow < 0) Then
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction", 1, "OK")
                                        Exit Function
                                    Else
                                        ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Update Trucking Instruction"
                                        modTrucking.GetDataFromMatrixByIndex(ImportSeaLCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        modTrucking.SetDataToEditTabByIndex(ImportSeaLCLForm)
                                        ImportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                        'MSW 14-09-2011 Truck PO
                                        If ImportSeaLCLForm.Items.Item("op_Exter").Specific.Selected = True Then
                                            EnabledTruckerForExternal(ImportSeaLCLForm, False)
                                        ElseIf ImportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                            EnabledTruckerForExternal(ImportSeaLCLForm, True)
                                        End If
                                        'MSW 14-09-2011 Truck PO
                                        'ImportSeaLCLForm.Items.Item("op_Inter").Enabled = False
                                        'ImportSeaLCLForm.Items.Item("op_Exter").Enabled = False
                                        ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True 'MSW
                                        ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = True 'MSW to edit New Ticket 07-09-2011
                                        ImportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                    End If

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
                                    ClearText(ImportSeaLCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                                    ImportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                    ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                    ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False 'MSW to edit New Ticket 07-09-2011
                                End If

                                '-------------------------Payment Vouncher (OMM)------------------------------------------------------'

                                If pVal.ItemUID = "fo_VoView" Then
                                    oMatrix = ImportSeaLCLForm.Items.Item("mx_Voucher").Specific
                                    If oMatrix.RowCount > 1 Then
                                        ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                        Else
                                            'ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                        End If
                                    End If
                                    ' ClearText(ImportSeaLCLForm, "ed_VedName", "ed_PayTo", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

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

                                If pVal.ItemUID = "bt_AmdVoc" Then
                                    'POP UP Payment Voucher
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadPaymentVoucher(ImportSeaLCLForm)
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Voucher.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
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



                                'Shipping Invoice POP UP
                                If pVal.ItemUID = "bt_ShpInv" Then
                                    'POP UP Shipping Invoice
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Items.Item("ed_WAddr").Specific.Value = "" And Not ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.Value = "" And Not ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific.Value = "" And Not ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = "" And Not ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value = "" And Not ImportSeaLCLForm.Items.Item("cb_PCode").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadShippingInvoice(ImportSeaLCLForm)
                                    Else
                                        If ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        ElseIf ImportSeaLCLForm.Items.Item("ed_Code").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("No Customer Code to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            'ElseIf ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" Then
                                            '    p_oSBOApplication.SetStatusBarMessage("No Shipping Agent Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        ElseIf ImportSeaLCLForm.Items.Item("cb_PCode").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("No Port of Loading to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)

                                        ElseIf ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("No Vessel/Voy Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        ElseIf ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("No WareHouse Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        ElseIf ImportSeaLCLForm.Items.Item("ed_WAddr").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("Please go to Warehouse Tab and  fill Warehouse Address to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        ElseIf ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific.Value = "" Then
                                            p_oSBOApplication.SetStatusBarMessage("No latest Clearance Date to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)

                                        End If
                                    End If
                                End If
                                'Purchase Order and Goods Receipt POP UP
                                If pVal.ItemUID = "bt_CPO" Then
                                    ' ==================================== Creating Custom Purchase Order ==============================
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ImportSeaLCLForm, "FumiPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                    ' ==================================== Creating Custom Purchase Order ==============================
                                End If

                                If pVal.ItemUID = "bt_CGR" Then
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCGR(ImportSeaLCLForm, "FumiGoodsReceipt.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_CrPO" Then 'sw
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ImportSeaLCLForm, "CratePurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_ForkPO" Then 'sw
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        'If ImportSeaLCLForm.Items.Item("ed_FVendor").Specific.Value = "TO001" Then
                                        '    If Not CreatePOPDF(ImportSeaLCLForm, "mx_Fork", "TO001", "F00001") Then Throw New ArgumentException(sErrDesc)
                                        'Else
                                        '    LoadAndCreateCPO(ImportSeaLCLForm, "ForkPurchaseOrder.srf")
                                        'End If
                                        LoadAndCreateCPO(ImportSeaLCLForm, "ForkPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If
                                If pVal.ItemUID = "bt_Orider" Then 'sw
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        'If ImportSeaLCLForm.Items.Item("ed_FVendor").Specific.Value = "TO001" Then
                                        '    If Not CreatePOPDF(ImportSeaLCLForm, "mx_Fork", "TO001", "F00001") Then Throw New ArgumentException(sErrDesc)
                                        'Else
                                        '    LoadAndCreateCPO(ImportSeaLCLForm, "ForkPurchaseOrder.srf")
                                        'End If
                                        LoadAndCreateCPO(ImportSeaLCLForm, "OutPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_ArmePO" Then 'sw
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ImportSeaLCLForm, "ArmedPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_CranePO" Then 'sw
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        'If ImportSeaLCLForm.Items.Item("ed_CVendor").Specific.Value = "TO001" Then
                                        '    If Not CreatePOPDF(ImportSeaLCLForm, "mx_Crane", "TO001", "CR0001") Then Throw New ArgumentException(sErrDesc)
                                        'Else
                                        '    LoadAndCreateCPO(ImportSeaLCLForm, "CranePurchaseOrder.srf")
                                        'End If
                                        LoadAndCreateCPO(ImportSeaLCLForm, "CranePurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_BunkPO" Then 'sw
                                    If Not ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        'If Not CreatePOPDF(ImportSeaLCLForm, "mx_Bunk", "TO001", "B00001") Then Throw New ArgumentException(sErrDesc)
                                        LoadAndCreateCPO(ImportSeaLCLForm, "BunkPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If
                                '-----------------------------------------------------------------------------------------------------------'

                                If pVal.ItemUID = "mx_Cont" And ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ImportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    If pVal.Row > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            modTrucking.rowIndex = CInt(pVal.Row)
                                            modTrucking.GetDataFromMatrixByIndex(ImportSeaLCLForm, oMatrix, modTrucking.rowIndex)
                                        End If
                                    End If

                                End If

                                'If pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize" And pVal.Before_Action = True And pVal.Row <> 0 Then
                                '    oMatrix = ImportSeaLCLForm.Items.Item("mx_Cont").Specific
                                '    Dim oColCombo As SAPbouiCOM.Column
                                '    Dim omatCombo As SAPbouiCOM.ComboBox
                                '    oColCombo = oMatrix.Columns.Item("colCSize")
                                '    omatCombo = oColCombo.Cells.Item(pVal.Row).Specific
                                '    If Not String.IsNullOrEmpty(oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Value) Then
                                '        Dim type As String = oMatrix.Columns.Item("colCType").Cells.Item(pVal.Row).Specific.Selected.Value
                                '        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '        oRecordSet.DoQuery("select U_Size from [@OBT_TB008_CONTAINER] where U_Type='" & type & "'")                       '* Change Nyan Lin   "[@OBT_TB008_CONTAINER]" 
                                '        oRecordSet.MoveFirst()
                                '        While oColCombo.ValidValues.Count > 0
                                '            oColCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                '        End While
                                '        If oRecordSet.RecordCount > 0 Then
                                '            While Not oRecordSet.EoF
                                '                omatCombo.ValidValues.Add(oRecordSet.Fields.Item("U_Size").Value, " ")
                                '                oRecordSet.MoveNext()
                                '            End While
                                '        End If
                                '        oColCombo = Nothing
                                '    End If
                                'End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
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
                                    'MSW
                                    ''If pVal.ItemUID = "ed_CVendor" Then
                                    ''    ImportSeaLCLForm.Items.Item("ed_CVendor").Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    ''End If
                                    ''If pVal.ItemUID = "ed_FVendor" Then
                                    ''    ImportSeaLCLForm.Items.Item("ed_FVendor").Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    ''End If

                                    If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString
                                        If String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value) Then
                                            'ImportSeaLCLForm.Items.Item("ed_ETADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            'ImportSeaLCLForm.Items.Item("ed_ETADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            'ImportSeaLCLForm.Items.Item("ed_ETATime").Specific.Value = Now.ToString("HH:mm")
                                            'If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        If String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value) Then
                                            'ImportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            'ImportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            'ImportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                            'If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        EnabledHeaderControls(ImportSeaLCLForm, True)
                                        If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                                            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListUID = "WRHSE"
                                            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListAlias = "Code"
                                            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                                            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                                            'ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                            'ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                        End If
                                        'ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)  ' MSW New Ticket 07-09-2011
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName,VatIdUnCmp FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'") ' MSW New Ticket 07-09-2011
                                        If oRecordSet.RecordCount > 0 Then
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_IUEN", 0, oRecordSet.Fields.Item("VatIdUnCmp").Value)  ' MSW New Ticket 07-09-2011
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_IComName", 0, oRecordSet.Fields.Item("CardName").Value)   '* Change Nyan Lin   "[@OBT_LCL06_PMAI]"
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_ShpAgt" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)     '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)      '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        'ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)            '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName,VatIdUnCmp FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'") ' MSW New Ticket 07-09-2011
                                        If oRecordSet.RecordCount > 0 Then
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_UEN", 0, oRecordSet.Fields.Item("VatIdUnCmp").Value) ' MSW New Ticket 07-09-2011
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString) '* Change Nyan Lin   "[@OBT_LCL02_PERMIT]"
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
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL13_WAREHOUSE").SetValue("U_WHr", 0, Trim(oDataTable.Columns.Item("U_Wwhl1").Cells.Item(0).Value.ToString()))
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
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_RelName", 0, WAddress)  'NL LCL Change according to change latest design 
                                            ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_PoR", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                        End If
                                    End If
                                    If pVal.ItemUID = "ed_Trucker" Then
                                        If ImportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                            'ImportSeaLCLForm.Freeze(True)
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                            oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                            ImportSeaLCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = "" '25-3-2011
                                            'ImportSeaLCLForm.Freeze(False)
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_Vessel" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)        '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ' ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)    '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_VName", 0, ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.String)                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    End If

                                    If pVal.ItemUID = "ed_CurCode" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_CurCode", 0, oDataTable.GetValue(0, 0).ToString)
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_ExRate", 0, Rate.ToString)
                                    End If
                                    If pVal.ItemUID = "ed_CCharge" Then
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_Cchange", 0, oDataTable.GetValue(0, 0).ToString)                            '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_CEchange", 0, Rate.ToString)                                               '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                    If pVal.ItemUID = "ed_Charge" Then
                                        ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_FCchange", 0, oDataTable.GetValue(0, 0).ToString)                          '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                Catch ex As Exception
                                End Try

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If pVal.ItemUID = "cb_PCode" Then
                                    oCombo = ImportSeaLCLForm.Items.Item("cb_PCode").Specific
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL01_IMPSEALCL").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                             '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_PCode", 0, oCombo.Selected.Value.ToString)                                       '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                End If
                                If pVal.ItemUID = "cb_BnkName" Then
                                    oCombo = ImportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim test As String = "select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description
                                    oRecordSet.DoQuery("select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description)
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL05_VOUCHER").SetValue("U_GLAC", 0, oRecordSet.Fields.Item("GLAccount").Value)                            '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    'oCombo.Selected.Description
                                End If

                                If pVal.ItemUID = "cb_PType" Then
                                    oCombo = ImportSeaLCLForm.Items.Item("cb_PType").Specific
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_TUnit", 0, oCombo.Selected.Value.ToString)
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
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_InvNo", 0, ImportSeaLCLForm.Items.Item("ed_InvNo").Specific.String)         '* Change Nyan Lin   "[@OBT_TB0011_VOUCHER]"
                                End If
                                'NL LCL Change 24-03-2011
                                If pVal.ItemUID = "ed_NOP" Then
                                    ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL02_PERMIT").SetValue("U_TotalOP", 0, ImportSeaLCLForm.Items.Item("ed_NOP").Specific.String)
                                End If
                                'End NL LCL Change 24-03-2011
                                If BubbleEvent = False Then
                                    Validateforform(pVal.ItemUID, ImportSeaLCLForm)
                                End If
                                'If ImportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Or ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                '    Validateforform(pVal.ItemUID, ImportSeaLCLForm)
                                'End If


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
                                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ImportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & ""  'MSW 08-06-2011 for job Type Table
                                            oRecordSet.DoQuery(sql)
                                            'End MSW 01-06-2011 for job type table
                                        End If
                                    End If
                                    If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        'handle for dispatch complete check box
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT U_Complete FROM [@OBT_LCL04_DISPATCH] WHERE DocEntry = " & ImportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value)
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
                    'MSW To Add New Button at ChooseFromList
                    'When User Click New button at ChooseFromList Form ,Setup Form Show to fill data.
                Case "9999"
                    oActiveForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_VNew" Then
                            oActiveForm.Close()
                            p_oSBOApplication.ActivateMenuItem("47644")
                        End If
                        If pVal.ItemUID = "bt_VoNew" Then
                            oActiveForm.Close()
                            p_oSBOApplication.ActivateMenuItem("47644")
                        End If
                    End If
                    'MSW To Add New Button at ChooseFromList
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

            If eventInfo.FormUID = "IMPORTSEALCL" Or eventInfo.FormUID = "IMPORTAIR" Or eventInfo.FormUID = "IMPORTLAND" Then
                oMenuItem = p_oSBOApplication.Menus.Item("1280")
                oMenus = oMenuItem.SubMenus
                ImportSeaLCLForm = p_oSBOApplication.Forms.GetForm(eventInfo.FormUID, 1)
                ImportSeaLCLForm.EnableMenu("772", True)
                ImportSeaLCLForm.EnableMenu("773", True)
                ImportSeaLCLForm.EnableMenu("775", True)
                If eventInfo.ItemUID = "mx_Fumi" Or eventInfo.ItemUID = "mx_Crate" Or eventInfo.ItemUID = "mx_Fork" Or eventInfo.ItemUID = "mx_Armed" Or eventInfo.ItemUID = "mx_Crane" Or eventInfo.ItemUID = "mx_Bunk" Or eventInfo.ItemUID = "mx_TkrList" Or eventInfo.ItemUID = "mx_Orider" Then 'MSW 14-09-2011 Truck PO
                    If eventInfo.BeforeAction = True Then
                        ImportSeaLCLForm.EnableMenu("772", False)
                        ImportSeaLCLForm.EnableMenu("773", False)
                        ImportSeaLCLForm.EnableMenu("775", False)
                        If eventInfo.ItemUID = "mx_TkrList" Then
                            oMatrix = ImportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                            If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                                If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colInsDoc").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                    If Not oMenus.Exists("EditCPO") Then
                                        oMenus.Add("EditCPO", "Edit Custom Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                        RPOmatrixname = eventInfo.ItemUID
                                        RPOsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "PurchaseOrder.srf"
                                    End If
                                    If Not oMenus.Exists("CancelPO") Then
                                        oMenus.Add("CancelPO", "Cancel Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                        RPOmatrixname = eventInfo.ItemUID
                                    End If
                                    If Not oMenus.Exists("CopyToCGR") Then
                                        oMenus.Add("CopyToCGR", "Copy To Custom Goods Receipt", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                        RGRmatrixname = eventInfo.ItemUID
                                        RGRsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "GoodsReceipt.srf"
                                    End If
                                    currentRow = eventInfo.Row
                                    ActiveMatrix = eventInfo.ItemUID
                                End If
                            End If
                        Else
                            oMatrix = ImportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                            If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then

                                If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                    If Not oMenus.Exists("EditCPO") Then
                                        oMenus.Add("EditCPO", "Edit Custom Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                        RPOmatrixname = eventInfo.ItemUID
                                        RPOsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "PurchaseOrder.srf"
                                    End If
                                    If Not oMenus.Exists("CancelPO") Then
                                        oMenus.Add("CancelPO", "Cancel Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                        RPOmatrixname = eventInfo.ItemUID
                                    End If
                                    If Not oMenus.Exists("CopyToCGR") Then
                                        oMenus.Add("CopyToCGR", "Copy To Custom Goods Receipt", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                        RGRmatrixname = eventInfo.ItemUID
                                        RGRsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "GoodsReceipt.srf"
                                    End If
                                    currentRow = eventInfo.Row
                                    ActiveMatrix = eventInfo.ItemUID
                                End If
                            End If

                        End If
                    Else
                        If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                            p_oSBOApplication.Menus.RemoveEx("EditCPO")
                            p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                            p_oSBOApplication.Menus.RemoveEx("CancelPO")
                        End If
                    End If
                End If

                If eventInfo.ItemUID = "mx_ShpInv" Then
                    ImportSeaLCLForm.EnableMenu("772", False)
                    ImportSeaLCLForm.EnableMenu("773", False)
                    ImportSeaLCLForm.EnableMenu("775", False)
                    If eventInfo.BeforeAction = True Then
                        oMatrix = ImportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNum").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If Not oMenus.Exists("EditShp") Then
                                    oMenus.Add("EditShp", "Edit Shipping Invoice", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = "ShipInvoice.srf"
                                End If
                                currentRow = eventInfo.Row
                                p_oSBOApplication.Forms.ActiveForm.EnableMenu("1293", True) 'MSW To Edit
                            End If
                        End If
                    Else
                        If oMenus.Exists("EditShp") Then
                            p_oSBOApplication.Menus.RemoveEx("EditShp")
                            p_oSBOApplication.Forms.ActiveForm.EnableMenu("1293", False)
                        End If

                    End If
                End If
                If eventInfo.ItemUID = "mx_Voucher" Then
                    ImportSeaLCLForm.EnableMenu("772", False)
                    ImportSeaLCLForm.EnableMenu("773", False)
                    ImportSeaLCLForm.EnableMenu("775", False)
                    If eventInfo.BeforeAction = True Then
                        oMatrix = ImportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colVocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If Not oMenus.Exists("EditVoc") Then
                                    oMenus.Add("EditVoc", "Edit Payment Voucher", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = "PaymentVoucher.srf"
                                End If
                                currentRow = eventInfo.Row
                            End If
                        End If
                    Else
                        If oMenus.Exists("EditVoc") Then
                            p_oSBOApplication.Menus.RemoveEx("EditVoc")
                        End If

                    End If

                End If

                If eventInfo.ItemUID = "mx_ShpInv" Then
                    ImportSeaLCLForm.EnableMenu("772", False)
                    ImportSeaLCLForm.EnableMenu("773", False)
                    ImportSeaLCLForm.EnableMenu("775", False)
                End If

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

        ' **********************************************************************************
        '   Function    :   EnabledHeaderControls()
        '   Purpose     :   This function will be providing to enable header control in
        '                   ImportseaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean
        '   Return      :   No

        '*************************************************************
        If pForm.Title.Substring(7, 3) = "Air" Then

            pForm.Items.Item("ed_MAWB").Enabled = pValue
            pForm.Items.Item("ed_HAWB").Enabled = pValue
            pForm.Items.Item("ed_AirLine").Enabled = pValue
            pForm.Items.Item("ed_FLNo").Enabled = pValue
            pForm.Items.Item("ed_AAgent").Enabled = pValue
            pForm.Items.Item("ed_CPCode").Enabled = pValue

            pForm.Items.Item("ed_ETATime").Enabled = pValue
        ElseIf pForm.Title.Substring(7, 4) = "Land" Then
            pForm.Items.Item("ed_PEntry").Enabled = pValue
            pForm.Items.Item("ed_ETATime").Enabled = pValue
        Else
            pForm.Items.Item("ed_ShpAgt").Enabled = pValue
            pForm.Items.Item("ed_OBL").Enabled = pValue
            pForm.Items.Item("ed_HBL").Enabled = pValue
            pForm.Items.Item("ed_Conn").Enabled = pValue
            pForm.Items.Item("ed_Vessel").Enabled = pValue
            pForm.Items.Item("ed_Voy").Enabled = pValue
            pForm.Items.Item("cb_PCode").Enabled = pValue
            ' pForm.Items.Item("ed_ETATime").Enabled = pValue
        End If

        pForm.Items.Item("ed_ETADate").Enabled = pValue
        pForm.Items.Item("ed_NOP").Enabled = pValue
        pForm.Items.Item("cb_PType").Enabled = pValue
        pForm.Items.Item("ed_CrgDsc").Enabled = pValue
        pForm.Items.Item("ed_TotalM3").Enabled = pValue
        pForm.Items.Item("ed_TotalWt").Enabled = pValue


        'MSW to Edit #1001
        If pValue = True And pForm.Items.Item("cb_JobType").Specific.Value = "" Then
            pForm.Items.Item("cb_JobType").Enabled = pValue
        End If
        'End MSW to Edit #1001
        'pForm.Items.Item("cb_JobType").Enabled = pValue
        pForm.Items.Item("cb_JbStus").Enabled = pValue
        pForm.Items.Item("ed_Wrhse").Enabled = pValue
        'pForm.Items.Item("ed_JobNo").Enabled = pValue

        pForm.Items.Item("ed_LCDDate").Enabled = pValue
        pForm.Items.Item("ed_LCDHr").Enabled = pValue

        pForm.Items.Item("ch_CnUn").Enabled = pValue
        pForm.Items.Item("ch_CgCl").Enabled = pValue
    End Sub

    Private Function CloseOpenForm(ByVal sFormId As String, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   CloseOpenForm()
        '   Purpose     :   This function will be providing to close and open forms for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByVal sFormId As String, ByRef sErrDesc As String
        '   Return      :   0 - FAILURE
        '               :   1 - SUCCESS
        '*************************************************************

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

    Private Function ClearComboData(ByRef pForm As SAPbouiCOM.Form, ByVal pComboUID As String, ByVal DataSource As String, ByVal FieldAlias As String) As Boolean
        ' **********************************************************************************
        '   Function    :   ClearComboData()
        '   Purpose     :   This function will be providing to clear Combo items  data for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pComboUID As String,
        '               :   ByVal DataSource As String, ByVal FieldAlias As String
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        '*************************************************************
        Dim sFuncName As String = "ClearComboData()"
        Dim sErrDesc As String = String.Empty
        ClearComboData = False
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If pForm.Items.Item(pComboUID).Specific.ValidValues.Count > 0 Then
                For i As Integer = pForm.Items.Item(pComboUID).Specific.ValidValues.Count - 1 To 0 Step -1
                    pForm.Items.Item(pComboUID).Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
                pForm.DataSources.DBDataSources.Item(DataSource).SetValue(FieldAlias, 0, vbNullString)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", sFuncName)
            ClearComboData = True
        Catch ex As Exception
            ClearComboData = False
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with ERROR", sFuncName)
        End Try
    End Function


    Private Sub ClearTruckingInfo(ByRef pForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   ClearTruckingInfo()
        '   Purpose     :   This function will be providing to clear items  data for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form
        '   Return      :  No
        '*************************************************************
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
        pForm.Items.Item("ee_Rmsk").Specific.Value = vbNullString 'MSW New Ticket 07-09-2011
        pForm.Freeze(False)
    End Sub

    Private Sub ClearText(ByRef pForm As SAPbouiCOM.Form, ByVal ParamArray pControls() As String)

        ' **********************************************************************************
        '   Function    :   ClearText()
        '   Purpose     :   This function will be providing to clear items  data for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal ParamArray pControls() As String
        '               
        '   Return      :   No
        '*************************************************************
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

        ' **********************************************************************************
        '   Function    :   EnabledUIObjects()
        '   Purpose     :   This function will be providing to clear items  data for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal ParamArray pControls() As String
        '               
        '   Return      :   No
        '*************************************************************
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
        ' **********************************************************************************
        '   Function    :   EnabledMaxtix
        '   Purpose     :   This function will be providing to enable maxtrix for main
        '                   Form.
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix,
        '                   ByVal pValue As Boolean 
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

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
        Try
            pForm.Items.Item("ed_InsDoc").Enabled = pValue
            pForm.Items.Item("ed_InsDoc").BackColor = 16645629
            'pForm.Items.Item("ed_PONo").Enabled = pValue
            'pForm.Items.Item("ed_PONo").BackColor = 16645629
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
            'MSW New Ticket 07-09-2011
            pForm.Items.Item("ee_Rmsk").Enabled = pValue
            pForm.Items.Item("ee_Rmsk").BackColor = 16645629
            'MSW New Ticket 07-09-2011
        Catch ex As Exception

        End Try
    End Sub

    Private Function Validateforform(ByVal ItemUID As String, ByVal ImportSeaLCLForm As SAPbouiCOM.Form) As Boolean
        ' **********************************************************************************
        '   Function    :   ValidateforformShippingInv()
        '   Purpose     :   This function provide to validate for Shipping Invoie items data in ExporeSeaLcl form  object
        '               
        '   Parameters  :  ByVal oActiveForm As SAPbouiCOM.Form
        '   Return      :  Fase - FAILURE
        '                  True - SUCCESS
        ' **********************************************************************************
        Try
            '*************************************************************
            If ImportSeaLCLForm.Title.Substring(7, 3) = "Air" Or ImportSeaLCLForm.Title.Substring(7, 4) = "Land" Then
                If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_Name").Specific.String) Then
                    p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
                    Return True
                Else
                    Return False
                End If
            Else
                If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_Name").Specific.String) Then
                    p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
                    Return True
                ElseIf (ItemUID = "cb_PCode" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("cb_PCode").Specific.Value) Then
                    p_oSBOApplication.SetStatusBarMessage("Must Select Port Of Loading", SAPbouiCOM.BoMessageTime.bmt_Short)
                    Return True
                    'ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.String) Then
                    '    p_oSBOApplication.SetStatusBarMessage("Must Choose Shipping Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                    '    Return True
                ElseIf (ItemUID = "ed_Wrhse" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.String) Then
                    p_oSBOApplication.SetStatusBarMessage("Must Choose Warehouse", SAPbouiCOM.BoMessageTime.bmt_Short)
                    Return True
                ElseIf (ItemUID = "ed_LCDDate" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific.String) Then
                    p_oSBOApplication.SetStatusBarMessage("Must Choose Last Day(Storage Incur)", SAPbouiCOM.BoMessageTime.bmt_Short)
                    Return True
                ElseIf (ItemUID = "ed_WAddr" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_WAddr").Specific.String) Then
                    p_oSBOApplication.SetStatusBarMessage("Must Fill Warehouse Address", SAPbouiCOM.BoMessageTime.bmt_Short)
                    Return True
                    'ElseIf (ItemUID = "ed_JobNo" Or ItemUID = " ") And String.IsNullOrEmpty(ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.String) Then
                    '    p_oSBOApplication.SetStatusBarMessage("Must Fill Job No", SAPbouiCOM.BoMessageTime.bmt_Short)
                    '    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Sub LoadHolidayMarkUp(ByVal ImportSeaLCLForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   LoadHolidayMarkUp()
        '   Purpose     :   This function will be providing to load Holiday Markup fomr for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByVal ExportSeaLCLForm As SAPbouiCOM.Form
        '               
        '   Return      :   No
        '                   
        ' **********************************************************************************

        Dim sErrDesc As String = String.Empty
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ETADay").Specific, ImportSeaLCLForm.Items.Item("ed_ETATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_LCDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_LCDDay").Specific, ImportSeaLCLForm.Items.Item("ed_LCDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CnDay").Specific, ImportSeaLCLForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_CgDay").Specific, ImportSeaLCLForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ImportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_ADay").Specific, ImportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Private Function AddChooseFromListByOption(ByRef pForm As SAPbouiCOM.Form, ByVal pOption As Boolean, ByVal pObjID As String, ByVal pErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   AddChooseFromListByOption()
        '   Purpose     :   This function will be providing to proceed validating for
        '                   choose choosefromlsit value  information
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pOption As Boolean, 
        '                       ByVal pObjID As String, ByVal pErrDesc As String        '                  
        '                   ByRef pErrDesc AS String 
        '                       pErrDesc = Error Description to be returned to calling function        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' ***********************************************************************************
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

#Region "---------- 'OMM - Purchase Voucher Save To Draft 09-03-2010"
    Private Function SaveToPurchaseVoucher(ByRef pForm As SAPbouiCOM.Form, ByVal ProcessedState As Boolean) As Boolean
        Dim nErr As Long
        Dim errMsg As String = String.Empty
        Dim ObjectCode As String = String.Empty
        Dim invDocEntry As Integer
        Dim invDate As String
        Dim Document As SAPbobsCOM.Documents
        Dim businessPartner As SAPbobsCOM.BusinessPartners
        SaveToPurchaseVoucher = False
        Try
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
                        Document.NumAtCard = pForm.Items.Item("ed_InvNo").Specific.Value

                        'Document.Address = pForm.Items.Item("ed_PayTo").Specific.Value
                        If businessPartner.Currency = "##" Then
                            ' Document.DocCurrency = "SGD"
                            Document.DocCurrency = pForm.Items.Item("cb_PayCur").Specific.Value().ToString.Trim
                        Else
                            Document.DocCurrency = businessPartner.Currency
                        End If

                        Document.DocDate = Now
                        Document.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                        Document.JournalMemo = "A/P Invoices - " & vendorCode
                        Document.Series = 6
                        If pForm.Items.Item("ed_InvDate").Specific.Value.ToString = "" Then
                            Document.TaxDate = Now
                        Else
                            invDate = Mid(pForm.Items.Item("ed_InvDate").Specific.Value, 5, 2) & "/" & Right(pForm.Items.Item("ed_InvDate").Specific.Value, 2) & "/" & Left(pForm.Items.Item("ed_InvDate").Specific.Value, 4)
                            Document.TaxDate = Convert.ToDateTime(invDate)
                        End If
                    End If
                    If (Document.Update() <> 0) Then
                        'MsgBox("Failed to add a payment")

                    Else

                        'Alert() 'That is Alert Nayn Lin
                    End If

                End If
            Else
                If (businessPartner.GetByKey(vendorCode)) Then
                    Document.CardCode = vendorCode
                    Document.CardName = pForm.Items.Item("ed_VedName").Specific.Value
                    Document.NumAtCard = pForm.Items.Item("ed_InvNo").Specific.Value

                    If businessPartner.Currency = "##" Then
                        Document.DocCurrency = pForm.Items.Item("cb_PayCur").Specific.Value().ToString.Trim
                    Else
                        Document.DocCurrency = businessPartner.Currency
                    End If

                    Document.DocDate = Now
                    Document.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    Document.JournalMemo = "A/P Invoices - " & vendorCode
                    Document.Series = 6
                    If pForm.Items.Item("ed_InvDate").Specific.Value.ToString = "" Then
                        Document.TaxDate = Now
                    Else
                        invDate = Mid(pForm.Items.Item("ed_InvDate").Specific.Value, 5, 2) & "/" & Right(pForm.Items.Item("ed_InvDate").Specific.Value, 2) & "/" & Left(pForm.Items.Item("ed_InvDate").Specific.Value, 4)
                        Document.TaxDate = Convert.ToDateTime(invDate)
                    End If
                    Dim oMatrix As SAPbouiCOM.Matrix
                    oMatrix = pForm.Items.Item("mx_ChCode").Specific
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            Document.Lines.ItemCode = oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value
                            Document.Lines.ItemDescription = oMatrix.Columns.Item("colVDesc1").Cells.Item(i).Specific.Value
                            Document.Lines.Quantity = 1
                            Document.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(i).Specific.Value)
                            'Document.Lines.Currency = "SGD"
                            'MSW 23-03-2011 For VatCode GST None or Blank in GST Field if we didn't assign ZI ,system auto populate default SI 
                            If oMatrix.Columns.Item("colGST1").Cells.Item(i).Specific.Value = "None" Then
                                Document.Lines.VatGroup = "ZI"
                            Else
                                Document.Lines.VatGroup = oMatrix.Columns.Item("colGST1").Cells.Item(i).Specific.Value()
                            End If
                            Document.Lines.Add()
                        Next
                    End If
                End If

                If (Document.Add() <> 0) Then
                    ' MsgBox("Failed to add a payment")
                End If

            End If

            'Check Error 

            Call p_oDICompany.GetLastError(nErr, errMsg)
            If (0 <> nErr) Then
                'MsgBox("Found error:" + Str(nErr) + "," + errMsg)
                p_oSBOApplication.SetStatusBarMessage(errMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                SaveToPurchaseVoucher = False
            Else
                ' MsgBox("Succeed in payment.add")
                p_oDICompany.GetNewObjectCode(ObjectCode)
                ObjectCode = p_oDICompany.GetNewObjectKey()
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                sql = "Update OPCH set U_JobNo='" & pForm.Items.Item("ed_PJobNo").Specific.Value & "',U_PVNo='" & pForm.Items.Item("ed_VocNo").Specific.Value & "',U_FrDocNo='" & pForm.Items.Item("ed_DocNum").Specific.Value & "'" & _
                " Where DocEntry = " & Convert.ToInt32(ObjectCode) & ""
                oRecordSet.DoQuery(sql)
                SaveToPurchaseVoucher = True
            End If
        Catch ex As Exception
            SaveToPurchaseVoucher = False
            MessageBox.Show(ex.Message)
        End Try


    End Function
    Private Function UpdateDraftPurchaseVoucher(ByRef pForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   UpdateDraftPurchaseVoucher()
        '   Purpose     :   This function will be providing to save  Draft purchase voucher form data  for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form
        '               
        '   Return      :   No
        ' **********************************************************************************


        Dim nErr As Long
        Dim errMsg As String = String.Empty
        Dim ObjectCode As String = String.Empty
        Dim puchaseDocEntry As Integer
        Dim nextOutgoing As Integer
        Dim draftDocEntry As Integer
        Dim updDocEntry As Integer
        Dim curCode As String
        Dim BnkCode As String
        Dim invDate As String
        Dim GLAcCode As String = String.Empty 'MSW To Edit New Ticket
        Dim sumApplied As Double = 0.0
        Dim appliedFC As Double = 0.0
        Dim docCur As String = String.Empty
        Dim vPay As SAPbobsCOM.Payments
        Dim businessPartner As SAPbobsCOM.BusinessPartners
        UpdateDraftPurchaseVoucher = False
        Try
            nextOutgoing = Convert.ToInt32(GetNewKey("46", oRecordSet).ToString)
            oRecordSet.DoQuery("select a.DocEntry as DraftDocEntry,b.DocEntry as purchaseDocEntry from OPDF a inner join PDF2 b on a.DocEntry=b.DocNum where a.U_FrPVNo='" & pForm.Items.Item("ed_DocNum").Specific.Value & "'")
            If oRecordSet.RecordCount > 0 Then
                updDocEntry = oRecordSet.Fields.Item("DraftDocEntry").Value
                puchaseDocEntry = oRecordSet.Fields.Item("purchaseDocEntry").Value
            End If

            If pForm.Items.Item("ed_InvDate").Specific.Value.ToString = "" Then
                invDate = Now
            Else
                invDate = Mid(pForm.Items.Item("ed_InvDate").Specific.Value, 5, 2) & "/" & Right(pForm.Items.Item("ed_InvDate").Specific.Value, 2) & "/" & Left(pForm.Items.Item("ed_InvDate").Specific.Value, 4)
            End If

            'MSW To Edit New Ticket
            oRecordSet.DoQuery("select GLAccount from DSC1 Where AliasName='" & pForm.Items.Item("cb_BnkName").Specific.Value.ToString.Trim & "'")
            If oRecordSet.RecordCount > 0 Then
                GLAcCode = oRecordSet.Fields.Item("GLAccount").Value
            End If
            'End MSW To Edit New Ticket
            oRecordSet.DoQuery("Update OPCH Set NumAtCard='" & pForm.Items.Item("ed_InvNo").Specific.Value & "',TaxDate='" & Convert.ToDateTime(invDate) & "' Where DocEntry=" & puchaseDocEntry)

            vPay = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
            businessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

            If (businessPartner.GetByKey(pForm.Items.Item("ed_VedCode").Specific.Value)) Then

                vPay.DocNum = nextOutgoing
                vPay.CardCode = pForm.Items.Item("ed_VedCode").Specific.Value
                vPay.CardName = pForm.Items.Item("ed_VedName").Specific.Value
                vPay.ApplyVAT = 1
                If businessPartner.Currency = "##" Then
                    vPay.DocCurrency = pForm.Items.Item("cb_PayCur").Specific.Value().ToString.Trim
                Else
                    vPay.DocCurrency = businessPartner.Currency
                End If

                vPay.DocDate = Now
                vPay.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier
                vPay.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                vPay.JournalRemarks = "Outgoing Payments - " & vendorCode
                vPay.Series = 15
                vPay.TaxDate = Now
                vPay.TransferAccount = "120101"
                vPay.CashSum = vocTotal
                vPay.CashAccount = "120301"
                vPay.Invoices.DocEntry = puchaseDocEntry
                vPay.Invoices.DocLine = 0
                vPay.Invoices.SumApplied = vocTotal
                vPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                vPay.CheckAccount = GLAcCode 'MSW To Edit New Ticket

            End If

            If (vPay.Add() <> 0) Then
                '   MsgBox("Failed to add a payment")
            Else

            End If
            Call p_oDICompany.GetLastError(nErr, errMsg)
            If (0 <> nErr) Then
                MsgBox("Found error:" + Str(nErr) + "," + errMsg)
                UpdateDraftPurchaseVoucher = True
            Else
                p_oDICompany.GetNewObjectCode(ObjectCode)
                ObjectCode = p_oDICompany.GetNewObjectKey()
                draftDocEntry = Convert.ToInt32(ObjectCode)
                sql = "Update OPDF set U_FrPVNo='" & pForm.Items.Item("ed_DocNum").Specific.Value & "'" & _
                   " Where DocEntry = " & Convert.ToInt32(ObjectCode) & ""
                oRecordSet.DoQuery(sql)
                If businessPartner.Currency = "SGD" Then
                    curCode = "SGD"
                ElseIf businessPartner.Currency = "##" Then
                    curCode = pForm.Items.Item("cb_PayCur").Specific.Value().ToString.Trim
                Else
                    curCode = businessPartner.Currency
                End If
                If pForm.Items.Item("op_Cash").Specific.Selected = False Then
                    'MSW to Edit New Ticket
                    'oRecordSet.DoQuery("select BankCode From ODSC where BankName='" & pForm.Items.Item("cb_BnkName").Specific.Value.ToString & "'")
                    oRecordSet.DoQuery("select BankCode From DSC1 where AliasName='" & pForm.Items.Item("cb_BnkName").Specific.Value.ToString & "'")
                    'End MSW to Edit New Ticket
                    If oRecordSet.RecordCount > 0 Then
                        BnkCode = oRecordSet.Fields.Item("BankCode").Value
                    Else
                        BnkCode = String.Empty
                    End If
                    oRecordSet.DoQuery("Update OPDF set [CheckSum]=CashSum,[CheckSumFC]=CashSumFC,CashAcct='',CheckSumSy=CashSumSy,CashSum=" & 0.0 & ",CashSumSy=" & 0.0 & ",CashSumFC=" & 0.0 & ",PayNoDoc='N',NoDocSumSy=" & 0.0 & ",NoDocSum=" & 0.0 & ",NoDocSumFC=" & 0.0 & "Where DocEntry=" & draftDocEntry)
                    oRecordSet.DoQuery("Select * from PDF1 Where DocNum=" & updDocEntry)
                    'MSW To Edit New Ticket Find GLAcCode in query
                    If oRecordSet.RecordCount > 0 Then
                        sql = "Update PDF1 set DueDate='" & Today.Date.ToString("yyyyMMdd") & "',CheckSum=" & vocTotal & ",Currency='" & curCode & "',CheckAct='" & GLAcCode & "',CountryCod='SG',CheckNum='" & pForm.Items.Item("ed_Cheque").Specific.Value & "',BankCode='" & BnkCode & "',ManualChk='Y'" & _
                        "Where DocNum=" & updDocEntry
                    Else
                        sql = "Insert Into PDF1 (DocNum,LineID,DueDate,CheckSum,Currency,CheckAct,CountryCod,CheckNum,BankCode,ManualChk) Values " & _
                                                       "(" & updDocEntry & _
                                                        "," & 0 & _
                                                        ",'" & Today.Date.ToString("yyyyMMdd") & _
                                                        "'," & vocTotal & _
                                                        ",'" & curCode & _
                                                        "','" & GLAcCode & "'" & _
                                                        ",'SG'" & _
                                                        ",'" & pForm.Items.Item("ed_Cheque").Specific.Value & _
                                                        "','" & BnkCode & _
                                                        "','Y')"
                    End If
                    oRecordSet.DoQuery(sql)
                    sql = "SELECT [CheckSum] as SumApplied,CheckSumFC as AppliedFC,DocCurr From OPDF WHERE Docentry=" & Convert.ToInt32(ObjectCode)
                Else
                    sql = "SELECT CashSum as SumApplied,CheckSumFC as AppliedFC,DocCurr From OPDF WHERE Docentry=" & Convert.ToInt32(ObjectCode)
                End If
                oRecordSet.DoQuery(sql)
                If oRecordSet.RecordCount > 0 Then
                    sumApplied = oRecordSet.Fields.Item("SumApplied").Value
                    appliedFC = oRecordSet.Fields.Item("AppliedFC").Value
                    docCur = oRecordSet.Fields.Item("DocCurr").Value
                End If
                If vPay.GetByKey(ObjectCode) Then
                    Try
                        sql = "UPDATE OPDF SET CashAcct = newdata.CashAcct,CashSum = newdata.CashSum,CashSumFC = newdata.CashSumFC,CheckAcct = newdata.CheckAcct,[CheckSum] = newdata.[CheckSum],CheckSumFC = newdata.CheckSumFC,DocCurr = newdata.DocCurr," & _
                            "DocRate = newdata.DocRate,SysRate = newdata.SysRate,DocTotal = newdata.DocTotal,DocTotalFC = newdata.DocTotalFC,CashSumSy = newdata.CashSumSy,CheckSumSy = newdata.CheckSumSy,DocTotalSy = newdata.DocTotalSy" & _
                            " From (SELECT CashAcct,CashSum,CashSumFC,CheckAcct,[CheckSum],CheckSumFC,DocCurr,DocRate,SysRate,DocTotal,DocTotalFC,CashSumSy,CheckSumSy,DocTotalSy FROM OPDF" & _
                            " WHERE Docentry=" & Convert.ToInt32(ObjectCode) & ") newdata WHERE Docentry=" & updDocEntry
                        oRecordSet.DoQuery(sql)
                        sql = "Update PDF2 Set SumApplied=newdata.SumApplied,AppliedFC=newdata.AppliedFC,AppliedSys=newdata.AppliedSys,DocRate=newdata.DocRate,vatApplied=newdata.vatApplied," & _
                                "vatAppldFC=newdata.vatAppldFC,vatAppldSy=newdata.vatAppldSy,BfDcntSum=newdata.BfDcntSum,BfDcntSumF=newdata.BfDcntSumF,BfDcntSumS=newdata.BfDcntSumS" & _
                                 " From (Select SumApplied,AppliedFC,AppliedSys,DocRate,vatApplied,vatAppldFC,vatAppldSy,BfDcntSum,BfDcntSumF,BfDcntSumS From PDF2 Where DocNum=" & Convert.ToInt32(ObjectCode) & ") newdata Where DocNum=" & updDocEntry
                        oRecordSet.DoQuery(sql)

                        If docCur <> "SGD" Then
                            sql = "Update PDF2 Set SumApplied=" & sumApplied & ",AppliedFC=" & appliedFC & ",AppliedSys=" & sumApplied & "," & _
                                "BfDcntSum=" & sumApplied & ",BfDcntSumF=" & appliedFC & ",BfDcntSumS=" & sumApplied & " Where DocNum=" & updDocEntry
                            oRecordSet.DoQuery(sql)
                        End If

                        If pForm.Items.Item("op_Cash").Specific.Selected = True Then
                            sql = "Delete From PDF1 WHERE DocNum=" & updDocEntry
                            oRecordSet.DoQuery(sql)
                        End If
                        sql = "Delete From OPDF WHERE Docentry=" & Convert.ToInt32(ObjectCode)
                        oRecordSet.DoQuery(sql)
                        sql = "Delete From PDF1 WHERE DocNum=" & Convert.ToInt32(ObjectCode)
                        oRecordSet.DoQuery(sql)
                        sql = "Delete From PDF2 WHERE DocNum=" & Convert.ToInt32(ObjectCode)
                        oRecordSet.DoQuery(sql)
                    Catch ex As Exception

                    End Try
                End If
                UpdateDraftPurchaseVoucher = True
            End If
        Catch ex As Exception
            UpdateDraftPurchaseVoucher = False
            MessageBox.Show(ex.Message)
        End Try


    End Function
    Private Function SaveToDraftPurchaseVoucher(ByRef pForm As SAPbouiCOM.Form) As Boolean
        Dim nErr As Long
        Dim errMsg As String = String.Empty
        Dim ObjectCode As String = String.Empty
        Dim puchaseDocEntry As Integer
        Dim nextOutgoing As Integer
        Dim draftDocEntry As Integer
        Dim curCode As String
        Dim BnkCode As String
        Dim GLAcCode As String = String.Empty 'MSW To Edit New Ticket
        Dim Document As SAPbobsCOM.Documents
        Dim docTotal As Double = 0.0
        Dim docTotalFC As Double = 0.0
        Dim docCur As String = String.Empty
        Dim curSource As String = String.Empty
        SaveToDraftPurchaseVoucher = False
        Try
            p_oDICompany.GetNewObjectCode(ObjectCode)
            ObjectCode = p_oDICompany.GetNewObjectKey()
            puchaseDocEntry = Convert.ToInt32(ObjectCode)
            Dim vPay As SAPbobsCOM.Payments
            Dim businessPartner As SAPbobsCOM.BusinessPartners

            nextOutgoing = Convert.ToInt32(GetNewKey("46", oRecordSet).ToString)
            Document = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
            vPay = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentsDrafts)
            vPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments
            businessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            'MSW To Edit New Ticket
            oRecordSet.DoQuery("select GLAccount from DSC1 Where AliasName='" & pForm.Items.Item("cb_BnkName").Specific.Value.ToString.Trim & "'")
            If oRecordSet.RecordCount > 0 Then
                GLAcCode = oRecordSet.Fields.Item("GLAccount").Value
            End If
            'End MSW To Edit New Ticket

            oRecordSet.DoQuery("select CurSource,DocCur,DocTotal,DocTotalFC from OPCH  Where DocEntry=" & puchaseDocEntry)
            If oRecordSet.RecordCount > 0 Then
                docTotal = oRecordSet.Fields.Item("DocTotal").Value
                docTotalFC = oRecordSet.Fields.Item("DocTotalFC").Value
                docCur = oRecordSet.Fields.Item("DocCur").Value
                curSource = oRecordSet.Fields.Item("CurSource").Value
            End If

            If (businessPartner.GetByKey(vendorCode)) Then

                vPay.DocNum = nextOutgoing
                vPay.CardCode = vendorCode
                vPay.CardName = pForm.Items.Item("ed_VedName").Specific.Value
                vPay.ApplyVAT = 1
                'If businessPartner.Currency = "##" Then
                '    vPay.DocCurrency = "SGD"
                '    vPay.DocCurrency = pForm.Items.Item("cb_PayCur").Specific.Value().ToString.Trim
                'Else
                '    vPay.DocCurrency = businessPartner.Currency
                'End If

                vPay.DocCurrency = docCur
                vPay.DocDate = Now
                vPay.DocTypte = SAPbobsCOM.BoRcptTypes.rSupplier
                vPay.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                vPay.JournalRemarks = "Outgoing Payments - " & vendorCode
                vPay.Series = 15
                vPay.TaxDate = Now
                vPay.TransferAccount = "120101"
                If docCur <> "SGD" Then
                    vPay.CashSum = docTotalFC
                    vPay.Invoices.SumApplied = docTotalFC
                Else
                    vPay.CashSum = docTotal
                    vPay.Invoices.SumApplied = docTotal
                End If

                '  vPay.CashSum = vocTotal

                vPay.CashAccount = "120301"
                vPay.Invoices.DocEntry = puchaseDocEntry
                vPay.Invoices.DocLine = 0
                'vPay.Invoices.SumApplied = vocTotal
                vPay.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice
                vPay.CheckAccount = GLAcCode 'MSW To Edit New Ticket

            End If

            If (vPay.Add() <> 0) Then
                ' MsgBox("Failed to add a payment")
            Else
                'Alert() 'That is Alert Nayn Lin

            End If
            Call p_oDICompany.GetLastError(nErr, errMsg)
            If (0 <> nErr) Then
                MsgBox("Found error:" + Str(nErr) + "," + errMsg)
                SaveToDraftPurchaseVoucher = False
            Else
                p_oDICompany.GetNewObjectCode(ObjectCode)
                ObjectCode = p_oDICompany.GetNewObjectKey()
                draftDocEntry = Convert.ToInt32(ObjectCode)
                sql = "Update OPDF set U_FrPVNo='" & pForm.Items.Item("ed_DocNum").Specific.Value & "'" & _
                " Where DocEntry = " & Convert.ToInt32(ObjectCode) & ""
                oRecordSet.DoQuery(sql)

                'If businessPartner.Currency = "SGD" Then
                '    curCode = "SGD"
                'ElseIf businessPartner.Currency = "##" Then
                '    curCode = pForm.Items.Item("cb_PayCur").Specific.Value().ToString.Trim
                'Else
                '    curCode = businessPartner.Currency
                'End If
                If pForm.Items.Item("op_Cash").Specific.Selected = False Then
                    'MSW to Edit New Ticket
                    'oRecordSet.DoQuery("select BankCode From ODSC where BankName='" & pForm.Items.Item("cb_BnkName").Specific.Value.ToString & "'")
                    oRecordSet.DoQuery("select BankCode From DSC1 where AliasName='" & pForm.Items.Item("cb_BnkName").Specific.Value.ToString & "'")
                    'End MSW to Edit New Ticket
                    If oRecordSet.RecordCount > 0 Then
                        BnkCode = oRecordSet.Fields.Item("BankCode").Value
                    Else
                        BnkCode = String.Empty
                    End If
                    Dim checkSum As Double = 0.0
                    oRecordSet.DoQuery("Update OPDF set [CheckSum]=CashSum,[CheckSumFC]=CashSumFC,CashAcct='',CheckSumSy=CashSumSy,CashSum=" & 0.0 & ",CashSumSy=" & 0.0 & ",CashSumFC=" & 0.0 & ",PayNoDoc='N',NoDocSumSy=" & 0.0 & ",NoDocSum=" & 0.0 & ",NoDocSumFC=" & 0.0 & "Where DocEntry=" & draftDocEntry)
                    oRecordSet.DoQuery("Select CheckSum,CheckSumFc from OPDF Where DocEntry=" & draftDocEntry)
                    If oRecordSet.RecordCount > 0 Then
                        If docCur <> "SGD" Then
                            checkSum = oRecordSet.Fields.Item("CheckSumFc").Value
                        Else
                            checkSum = oRecordSet.Fields.Item("CheckSum").Value
                        End If
                    End If

                    'MSW To Edit New Ticket find GLAcCode
                    sql = "Insert Into PDF1 (DocNum,LineID,DueDate,CheckSum,Currency,CheckAct,CountryCod,CheckNum,BankCode,ManualChk) Values " & _
                                                   "(" & draftDocEntry & _
                                                    "," & 0 & _
                                                    ",'" & Today.Date.ToString("yyyyMMdd") & _
                                                    "'," & checkSum & _
                                                    ",'" & docCur & _
                                                    "','" & GLAcCode & "'" & _
                                                    ",'SG'" & _
                                                    ",'" & pForm.Items.Item("ed_Cheque").Specific.Value & _
                                                    "','" & BnkCode & _
                                                    "','Y')"
                    oRecordSet.DoQuery(sql)
                End If
                SaveToDraftPurchaseVoucher = True
            End If
        Catch ex As Exception
            SaveToDraftPurchaseVoucher = False
            MessageBox.Show(ex.Message)
        End Try

    End Function
#End Region

    Private Sub Alert()
        Dim code As String = "14"
        Dim AlertName As String = "Pending PO2"
        Dim ObjectCode As String = String.Empty
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
            oAlertManagement.Name = "Pending PO2" 'Change
            oAlertManagement.QueryID = 103 'Change
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
            oAlertRecipient.UserCode = 16
            oAlertRecipient.SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
            oAlertManagementParams = oAlertMangementService.AddAlertManagement(oAlertManagement)
            oRecordSet.DoQuery("Update OALT set NextDate='" & Today.Date & "',NextTime='" & Now.Hour.ToString & Convert.ToString((Convert.ToInt32(Now.Minute.ToString()) + 2)) & "' where  Code='" & code.ToString() & "'")
        Else
            oRecordSet.DoQuery("Update OALT set NextDate='" & Today.Date & "',NextTime='" & Now.Hour.ToString & Convert.ToString((Convert.ToInt32(Now.Minute.ToString()) + 2)) & "' where  Code='" & code.ToString() & "'")
        End If

    End Sub

    Private Sub Start(ByRef pform As SAPbouiCOM.Form)

        '=============================================================================
        'Function   : Start()
        'Purpose    : This function to provide for To process with pdf form in ImportseaLcl

        'Parameters : ByRef oActiveForm As SAPbouiCOM.Form
        'Return     : No
        '             
        '==========================================

        Dim str As String = "dwdesk.exe"
        Dim myprocess As New Process
        Dim mainfolderpath As String = p_fmsSetting.DocuPath
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
                'MoveWindow(myprocess.MainWindowHandle, p_oSBOApplication.Forms.ActiveForm.Left + p_oSBOApplication.Forms.ActiveForm.Width + 6, p_oSBOApplication.Forms.ActiveForm.Top + 68, 500, 618, True)
            End If
        Catch ex As Exception

        End Try


    End Sub

#Region "Voucher POP Up"


    Private Sub LoadPaymentVoucher(ByRef oActiveForm As SAPbouiCOM.Form)
        '=============================================================================
        'Function   : LoadPaymentVoucher()
        'Purpose    : This function to provide for load payment voucher form in ImportSealCL 
        'Parameters : ByRef oActiveForm As SAPbouiCOM.Form
        'Return     : No
        '             
        '==========================================

        Dim oPayForm As SAPbouiCOM.Form
        Dim oOptBtn As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        LoadFromXML(p_oSBOApplication, "PaymentVoucher.srf")
        oPayForm = p_oSBOApplication.Forms.ActiveForm
        oPayForm.Freeze(True)
        If AddChooseFromList(oPayForm, "VPAYMENT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oPayForm.Items.Item("ed_VedCode").Specific.ChooseFromListUID = "VPAYMENT"
        oPayForm.Items.Item("ed_VedCode").Specific.ChooseFromListAlias = "CardCode"
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
        oPayForm.EnableMenu("15883", True)
        oPayForm.EnableMenu("1292", True)
        oPayForm.EnableMenu("1293", True)
        oPayForm.EnableMenu("1294", False)
        oPayForm.EnableMenu("772", True)
        oPayForm.EnableMenu("773", True)
        oPayForm.EnableMenu("775", True)
        oPayForm.AutoManaged = True
        oPayForm.DataBrowser.BrowseBy = "ed_DocNum"
        oPayForm.Items.Item("cb_BnkName").Enabled = False
        oPayForm.Items.Item("ed_Cheque").Enabled = False
        'oPayForm.Items.Item("ed_PayRate").Enabled = False
        oPayForm.Items.Item("ed_PayRate").Visible = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPayForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("VOUCHER", oRecordSet)
        oPayForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_InvDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        oPayForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()

        ' If HolidayMarkUp(oPayForm, oPayForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)

        If AddUserDataSrc(oPayForm, "VCASH", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oPayForm, "VCHEQUE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oOptBtn = oPayForm.Items.Item("op_Cash").Specific
        oOptBtn.DataBind.SetBound(True, "", "VCASH")
        oOptBtn = oPayForm.Items.Item("op_Cheq").Specific
        oOptBtn.DataBind.SetBound(True, "", "VCHEQUE")
        oOptBtn.GroupWith("op_Cash")

        oPayForm.Items.Item("op_Cheq").Specific.Selected = True
        oPayForm.Items.Item("ed_PayType").Specific.Value = "Cheque"

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
            oRecordSet.DoQuery("SELECT CurrCode,CurrName FROM OCRN")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value.ToString.Trim, oRecordSet.Fields.Item("CurrName").Value.ToString)
                    oRecordSet.MoveNext()
                End While
                oCombo.Select("SGD", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
        End If
        'MSW to Edit New Ticket
        oCombo = oPayForm.Items.Item("cb_PayFor").Specific
        If oCombo.ValidValues.Count = 0 Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT U_ChType FROM [@OBT_TB40_CHARGES]")
            If oRecordSet.RecordCount > 0 Then
                oCombo.ValidValues.Add("", "")
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_ChType").Value.ToString.Trim, "")
                    oRecordSet.MoveNext()
                End While
                oCombo.ValidValues.Add("Define New", "")
            Else
                oCombo.ValidValues.Add("", "")
                oCombo.ValidValues.Add("Define New", "")
            End If
        End If
        'End MSW to Edit New Ticket
        oPayForm.Items.Item("bt_PayView").Visible = False
        Dim oColumn As SAPbouiCOM.Column
        oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
        If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
        oColumn = oMatrix.Columns.Item("colChCode1")
        AddChooseFromList(oPayForm, "ChCode", False, "UDOCHCODE")
        oColumn.ChooseFromListUID = "ChCode"
        oColumn.ChooseFromListAlias = "Code" 'MSW To Edit New Ticket
        oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
        DisableChargeMatrix(oPayForm, oMatrix, True)
        oPayForm.Items.Item("ed_VedCode").Specific.Active = True
        oPayForm.Freeze(False)
    End Sub
    Private Function AddNewRow(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
        ' **********************************************************************************
        '   Function    :   AddNewRowPO()
        '   Purpose     :   This function will be providing to Create new row for Purchase Order form  maxtrix item in  ExporeSeaLCL Form  
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form,
        '               :   ByVal MatrixUID As String
        '   Return      :   False- FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

        AddNewRow = False
        Dim sErrDesc As String = vbNullString
        Dim oDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
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
            'MSW to edit
            'MSW to Edit New Ticket
            oCombo = oMatrix.Columns.Item("colGST1").Cells.Item(oMatrix.RowCount).Specific
            oRecordSet.DoQuery("Select ECVatGroup from OCRD where CardCode='" & oActiveForm.Items.Item("ed_VedCode").Specific.Value & "'")
            If oRecordSet.RecordCount > 0 Then
                If oRecordSet.Fields.Item("ECVatGroup").Value = "" Then
                    oCombo.Select("NI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    oCombo.Select(oRecordSet.Fields.Item("ECVatGroup").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
            Else
                oCombo.Select("NI", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            'If oActiveForm.Items.Item("cb_GST").Specific.Value.ToString.Trim = "No" Then
            '    oCombo.Select("None", SAPbouiCOM.BoSearchKey.psk_ByValue)
            'Else
            '    oRecordSet.DoQuery("Select ECVatGroup from OCRD where CardCode='" & oActiveForm.Items.Item("ed_VedCode").Specific.Value & "'")
            '    If oRecordSet.RecordCount > 0 Then
            '        If oRecordSet.Fields.Item("ECVatGroup").Value = "" Then
            '            oCombo.Select("NI", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '        Else
            '            oCombo.Select(oRecordSet.Fields.Item("ECVatGroup").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
            '        End If
            '    Else
            '        oCombo.Select("None", SAPbouiCOM.BoSearchKey.psk_ByValue)
            '    End If
            'End If
            'End MSW to Edit New Ticket
            AddNewRow = True
        Catch ex As Exception
            AddNewRow = False
        End Try
    End Function

    Private Sub DisableChargeMatrix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)
        '=============================================================================
        'Function   : DisableChargeMatrix()
        'Purpose    : This function to disable charge Matrix 
        '             ExporeSealCL form and other Purchase Order form 
        'Parameters : ByRef pForm As SAPbouiCOM.Form,ByRef pMatrix As SAPbouiCOM.Matrix,
        '             ByVal pValue As Boolean               
        'Return     : No
        '==========================================

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
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
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

#Region "Validate Job Number"
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
#End Region

#Region "SHIPPING INVOICE POP UP"
    Private Sub LoadShippingInvoice(ByRef oActiveForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   LoadShippingInvoice()
        '   Purpose     :   This function  to load and show shipping Invoice form in 
        '                   ExportSeaLCL Main fom for process of shipping Invoice data
        '               
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form
        '                                 
        '               
        '   Return      :  No
        '                 
        ' **********************************************************************************
        Dim oShpForm As SAPbouiCOM.Form
        Dim sErrDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix

        LoadFromXML(p_oSBOApplication, "ShipInvoice.srf")
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
        oShpForm = p_oSBOApplication.Forms.ActiveForm
        oShpForm.Freeze(True)
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
        'MSW to edit
        oShpForm.EnableMenu("1292", False)
        oShpForm.EnableMenu("1293", False)
        oShpForm.EnableMenu("4870", False)
        oShpForm.EnableMenu("771", False)
        oShpForm.EnableMenu("772", True)
        oShpForm.EnableMenu("773", True)
        oShpForm.EnableMenu("774", False)
        oShpForm.EnableMenu("775", True)
        'End MSW to edit

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


        If HolidayMarkUp(oShpForm, oShpForm.Items.Item("ed_ShDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


        ' oShpForm.Freeze(True)
        If AddUserDataSrc(oShpForm, "ItemNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Part", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PartDesp", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Qty", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Unit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Box", sErrDesc, SAPbouiCOM.BoDataType.dt_PRICE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "BoxLast", sErrDesc, SAPbouiCOM.BoDataType.dt_PRICE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
        If AddUserDataSrc(oShpForm, "TotBox", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "BUnit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PBox", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PUnit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "UPrice", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "TotV", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "SName", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PSName", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Ecc", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Lic", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "LicExDate", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Cls", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "UN", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "HSCode", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "DOM", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oShpForm.Items.Item("ed_ItemNo").Specific.DataBind.SetBound(True, "", "ItemNo")
        oShpForm.Items.Item("ed_Part").Specific.DataBind.SetBound(True, "", "Part")
        oShpForm.Items.Item("ed_PartDes").Specific.DataBind.SetBound(True, "", "PartDesp")
        oShpForm.Items.Item("ed_Qty").Specific.DataBind.SetBound(True, "", "Qty")
        oShpForm.Items.Item("ed_Unit").Specific.DataBind.SetBound(True, "", "Unit")
        oShpForm.Items.Item("ed_Box").Specific.DataBind.SetBound(True, "", "Box")
        oShpForm.Items.Item("ed_BoxLast").Specific.DataBind.SetBound(True, "", "BoxLast")
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

        'AddChooseFromList(oShpForm, "cflPart", False, "PART")
        If AddChooseFromListByFilter(oShpForm, "cflPart", False, "PART", "U_BPCode", SAPbouiCOM.BoConditionOperation.co_EQUAL, oActiveForm.Items.Item("ed_Code").Specific.Value) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oShpForm.Items.Item("ed_Part").Specific.ChooseFromListUID = "cflPart"
        oShpForm.Items.Item("ed_Part").Specific.ChooseFromListAlias = "U_PartNo"
        oShpForm.Items.Item("bt_PPView").Visible = False

        oMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
        If oShpForm.Items.Item("ed_ItemNo").Specific.Value = "" Then
            'If (oMatrix.RowCount > 0) Then
            '    If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
            '        oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
            '    Else
            '        oShpForm.Items.Item("ed_ItemNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
            '    End If
            'Else
            '    oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
            'End If
            If (oMatrix.RowCount > 0) Then
                If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                    oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
                    oShpForm.Items.Item("ed_Box").Specific.Value = 1
                    oShpForm.Items.Item("ed_Part").Specific.Active = True
                Else
                    oShpForm.Items.Item("ed_ItemNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                    oRecordSet.DoQuery("Select ISNULL(U_BoxLast,0)+1 as Box from [@OBT_TB03_EXPSHPINVD] Where LinedID=" & oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value)
                    If oRecordSet.RecordCount > 0 Then
                        oShpForm.Items.Item("ed_Box").Specific.Value = oRecordSet.Fields.Item("Box").Value
                    End If
                    oShpForm.Items.Item("ed_Part").Specific.Active = True
                End If
            Else
                oShpForm.Items.Item("ed_ItemNo").Specific.Value = 1
                oShpForm.Items.Item("ed_Box").Specific.Value = 1
                oShpForm.Items.Item("ed_Part").Specific.Active = True
            End If
        End If
        oShpForm.Items.Item("ed_ShipTo").Specific.Active = True


        oShpForm.Freeze(False)
    End Sub
#End Region

#Region "SHIPPING INV" 'syma
    Private Sub AddUpdateShippingInv(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)

        ' **********************************************************************************
        '   Function    :   AddUpdateShippingInv()
        '   Purpose     :   This function  to add and update shipping Invoice form in 
        '                   ExportSeaLCL Main fom for update process of shipping Invoice
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,    
        '               :   ByVal DataSource As String, ByVal ProcressedState As Boolean
        '   Return      :  No
        '                 
        ' **********************************************************************************


        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        'MSW to edit
        If pMatrix.GetNextSelectedRow < 0 Then
            rowIndex = Convert.ToInt32(pForm.Items.Item("ed_ItemNo").Specific.Value)
        Else
            rowIndex = pMatrix.GetNextSelectedRow
        End If
        'End MSW to edit

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
                    '.SetValue("U_PartNo", .Offset, pForm.Items.Item("ed_Part").Specific.Value)
                    .SetValue("U_Desc", .Offset, pForm.Items.Item("ed_PartDes").Specific.Value)
                    .SetValue("U_Qty", .Offset, pForm.Items.Item("ed_Qty").Specific.Value)
                    .SetValue("U_UM", .Offset, pForm.Items.Item("ed_Unit").Specific.Value)
                    .SetValue("U_Box", .Offset, pForm.Items.Item("ed_Box").Specific.Value)
                    .SetValue("U_BoxLast", .Offset, pForm.Items.Item("ed_BoxLast").Specific.Value)
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
                    .SetValue("U_PartNo", .Offset, pForm.Items.Item("ed_Part").Specific.Value)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource

                    .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(rowIndex).Specific.Value) 'MSW to edit
                    .SetValue("U_SerNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(rowIndex).Specific.Value) 'MSW to edit
                    .SetValue("U_ExInv", .Offset, pForm.Items.Item("ed_ExInv").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value)
                    '.SetValue("U_PartNo", .Offset, pForm.Items.Item("ed_Part").Specific.Value)
                    .SetValue("U_Desc", .Offset, pForm.Items.Item("ed_PartDes").Specific.Value)
                    .SetValue("U_Qty", .Offset, pForm.Items.Item("ed_Qty").Specific.Value)
                    .SetValue("U_UM", .Offset, pForm.Items.Item("ed_Unit").Specific.Value)
                    .SetValue("U_Box", .Offset, pForm.Items.Item("ed_Box").Specific.Value)
                    .SetValue("U_BoxLast", .Offset, pForm.Items.Item("ed_BoxLast").Specific.Value)
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
                    .SetValue("U_PartNo", .Offset, pForm.Items.Item("ed_Part").Specific.Value)
                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub SetShipInvDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)

        ' **********************************************************************************
        '   Function    :   SetShipInvDataToEditTabByIndex()
        '   Purpose     :   This function  to Edit e shipping Invoice data form in 
        '                   ExportSeaLCL Main fom for edit  process of shipping Invoice data
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, 
        '               :    ByVal Index As Integer
        '   Return      :  No
        '                 
        ' **********************************************************************************

        Dim sErrDesc As String = String.Empty

        Try
            pForm.Freeze(True)

            'pForm.Items.Item("ed_ConNo").Specific.Value = pMatrix.Columns.Item("V_-1").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_ExInv").Specific.Value = pMatrix.Columns.Item("colExInv").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PO").Specific.Value = pMatrix.Columns.Item("colPO").Cells.Item(Index).Specific.Value

            pForm.Items.Item("ed_ItemNo").Specific.Value = pMatrix.Columns.Item("V_-1").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Part").Specific.Value = pMatrix.Columns.Item("colPart").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PartDes").Specific.Value = pMatrix.Columns.Item("colPartDes").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Qty").Specific.Value = pMatrix.Columns.Item("colQty").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Unit").Specific.Value = pMatrix.Columns.Item("colQtyUnit").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Box").Specific.Value = pMatrix.Columns.Item("colBox").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_BoxLast").Specific.Value = pMatrix.Columns.Item("colBoxLast").Cells.Item(Index).Specific.Value
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
            pForm.Items.Item("ed_DOM").Specific.Value = pMatrix.Columns.Item("colDOM").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Part").Specific.Active = True

            ' If HolidaysMarkUp(pForm, pForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, pForm.Items.Item("ed_CunDay").Specific, pForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            pForm.Freeze(False)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddUpdateShippingMatrix(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)


        ' **********************************************************************************
        '   Function    :   AddUpdateShippingMatrix()
        '   Purpose     :   This function  process for add Shipping data from matrix to save into 
        '                   ExporeSeaLcl  object
        '               
        '   Parameters  :  ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '                  ByVal DataSource As String, ByVal ProcressedState As Boolean)        '                   
        '               
        '   Return      :  Fase - FAILURE
        '                  True - SUCCESS
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEALCL", 1)
        ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item(DataSource)
        '  ObjDBDataSource.Offset = 0

        rowIndex = pMatrix.GetNextSelectedRow

        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            If pMatrix.RowCount = 1 Then
                If pMatrix.Columns.Item("colDocNum").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                    pMatrix.Clear()
                End If
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
    Private Function ValidateforformShippingInv(ByVal oActiveForm As SAPbouiCOM.Form) As Boolean
        ' **********************************************************************************
        '   Function    :   ValidateforformShippingInv()
        '   Purpose     :   This function provide to validate for Shipping Invoie items data in ExporeSeaLcl form  object
        '               
        '   Parameters  :  ByVal oActiveForm As SAPbouiCOM.Form
        '   Return      :  Fase - FAILURE
        '                  True - SUCCESS
        ' **********************************************************************************

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

        ' **********************************************************************************
        '   Function    :   LoadAndCreateCPO()
        '   Purpose     :   This function will be providing to create and show Purchase Order form 
        '                   for Purchase Order form  
        '               
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,                       
        '                   ByRef srfName As String
        '               
        '   Return      :   No
        '                   
        ' **********************************************************************************

        Dim CPOForm As SAPbouiCOM.Form
        Dim CPOMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Try
            LoadFromXML(p_oSBOApplication, srfName)
            CPOForm = p_oSBOApplication.Forms.ActiveForm
            CPOForm.Freeze(True)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                CPOForm.EnableMenu("1292", True)
                CPOForm.EnableMenu("1293", True)

                CPOForm.EnableMenu("1294", False)
                CPOForm.EnableMenu("1283", False)
                CPOForm.EnableMenu("1284", False)
                CPOForm.EnableMenu("1286", False)
                CPOForm.EnableMenu("772", True)
                CPOForm.EnableMenu("773", True)
                CPOForm.EnableMenu("775", True)
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
                'If srfName = "CranePurchaseOrder.srf" Then
                '    CPOForm.Items.Item("ed_Code").Specific.Value = ParentForm.Items.Item("ed_CVendor").Specific.Value
                'ElseIf srfName = "ForkPurchaseOrder.srf" Then
                '    CPOForm.Items.Item("ed_Code").Specific.Value = ParentForm.Items.Item("ed_FVendor").Specific.Value
                'End If
                CPOForm.Items.Item("ed_Code").Specific.Active = True
                CPOForm.Freeze(False)
                ' ==================================== Custom Purchase Order ========================================
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub LoadAndCreateCGR(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)

        ' **********************************************************************************
        '   Function    :   LoadAndCreateCGR()
        '   Purpose     :   This function will be providing to Create and Load GoodReceipt data form for ExporeSeaLCL Form  
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,
        '               :   ByRef srfName As String
        '   Return      :   No
        '                   
        ' **********************************************************************************

        Dim CGRForm As SAPbouiCOM.Form
        Dim CGRMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Try
            LoadFromXML(p_oSBOApplication, srfName)
            CGRForm = p_oSBOApplication.Forms.ActiveForm
            CGRForm.Freeze(True)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                CGRForm.EnableMenu("1292", True)
                CGRForm.EnableMenu("1293", True)

                CGRForm.EnableMenu("1294", False)
                CGRForm.EnableMenu("1283", False)
                CGRForm.EnableMenu("1284", False)
                CGRForm.EnableMenu("1286", False)
                CGRForm.EnableMenu("772", True)
                CGRForm.EnableMenu("773", True)
                CGRForm.EnableMenu("775", True)
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
            If Not AddNewRowGR(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
            CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_DocNum").Specific.Value)
            CGRForm.Items.Item("ed_Code").Specific.Active = True
            CGRForm.Freeze(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function FillDataToGoodsReceipt(ByRef oSourceForm As SAPbouiCOM.Form, ByVal SourceMatrixName As String, _
                                            ByVal SourceColName1 As String, ByVal SourceColName2 As String, ByVal ActiveRow As Integer, _
                                            ByVal DestDataSource As String, ByRef oDestForm As SAPbouiCOM.Form) As Boolean


        ' **********************************************************************************
        '   Function    :   FillDataToGoodsReceipt()
        '   Purpose     :   This function provide to save goodreceipt  items from purchase order data in ExporeSeaLcl form  object
        '               
        '   Parameters  :  ByRef oSourceForm As SAPbouiCOM.Form, ByVal SourceMatrixName As String,
        '               :  ByVal SourceColName1 As String, ByVal SourceColName2 As String, ByVal ActiveRow As Integer,
        '               :  ByVal DestDataSource As String, ByRef oDestForm As SAPbouiCOM.Form 
        '   Return      :  Fase - FAILURE 
        '                  True - SUCCESS
        ' **********************************************************************************

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
                    Dim tsql As String = "SELECT U_SInA,U_TPlace,U_TDate,U_TDay,U_TTime,U_PORMKS,U_POIRMKS,U_ColFrm,U_TkrIns,U_TkrTo,U_CNo,U_Dest,U_LocWork,U_POITPD FROM [@OBT_TB08_FFCPO] WHERE DocEntry = " + FormatString(oMatrix.Columns.Item(SourceColName2).Cells.Item(ActiveRow).Specific.Value) 'MSW 14-09-2011 Truck PO
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
                            'MSW 14-09-2011 Truck PO
                            .SetValue("U_ColFrm", .Offset, oRecordset.Fields.Item("U_ColFrm").Value)
                            .SetValue("U_TkrIns", .Offset, oRecordset.Fields.Item("U_TkrIns").Value)
                            .SetValue("U_TkrTo", .Offset, oRecordset.Fields.Item("U_TkrTo").Value)
                            'MSW 14-09-2011 Truck PO
                            .SetValue("U_GRTPD", .Offset, oRecordset.Fields.Item("U_POITPD").Value)
                            .SetValue("U_Dest", .Offset, oRecordset.Fields.Item("U_Dest").Value)
                            .SetValue("U_LocWork", .Offset, oRecordset.Fields.Item("U_LocWork").Value)
                            .SetValue("U_CNo", .Offset, oRecordset.Fields.Item("U_CNo").Value)
                            oRecordset.MoveNext()
                        End While
                    End If
                End With
            End If
            Dim tempSQL As String = "SELECT * FROM POR1 WHERE DocEntry =" + FormatString(oPODocument.DocEntry) + " And OpenQty <> 0 "
            oRecordset.DoQuery(tempSQL)
            oDestMatrix.Clear()

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()
                While oRecordset.EoF = False
                    With oLineDBDataSource
                        .SetValue("LineId", .Offset, oDestMatrix.VisualRowCount + 1)
                        .SetValue("U_GRINO", .Offset, oRecordset.Fields.Item("ItemCode").Value)
                        .SetValue("U_GRIDesc", .Offset, oRecordset.Fields.Item("Dscription").Value)
                        '.SetValue("U_GRIQty", .Offset, oRecordset.Fields.Item("Quantity").Value)
                        .SetValue("U_GRIQty", .Offset, oRecordset.Fields.Item("OpenQty").Value)
                        .SetValue("U_GRIPrice", .Offset, oRecordset.Fields.Item("Price").Value)
                        .SetValue("U_GRIAmt", .Offset, oRecordset.Fields.Item("OpenSum").Value)
                        .SetValue("U_GRIGST", .Offset, oRecordset.Fields.Item("VatGroup").Value)
                        .SetValue("U_GRITot", .Offset, oRecordset.Fields.Item("OpenSum").Value)
                        .SetValue("U_POLineId", .Offset, oRecordset.Fields.Item("LineNum").Value)
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

        ' **********************************************************************************
        '   Function    :   CreateGoodsReceiptPO()
        '   Purpose     :   This function will be providing to create for GoodsReceipt
        '                    Purchase order form  
        '               
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form,ByVal MatrixUID As String        '                   
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

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
                    'p_oSBOApplication.MessageBox(oBusinessPartner.Currency)
                    If oBusinessPartner.Currency = "##" Then
                        oPurchaseDeliveryNote.DocCurrency = "SGD"
                    Else
                        oPurchaseDeliveryNote.DocCurrency = oBusinessPartner.Currency
                    End If

                    oPurchaseDeliveryNote.DocDate = Now
                    oPurchaseDeliveryNote.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    oPurchaseDeliveryNote.TaxDate = Now
                    oMatrix = oActiveForm.Items.Item(MatrixUID).Specific
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            oPurchaseDeliveryNote.Lines.BaseType = CInt(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            oPurchaseDeliveryNote.Lines.BaseEntry = Convert.ToInt32(oActiveForm.DataSources.UserDataSources.Item("PONo").Value)
                            oPurchaseDeliveryNote.Lines.BaseLine = Convert.ToInt32(oMatrix.Columns.Item("colLineId").Cells.Item(i).Specific.Value)
                            oPurchaseDeliveryNote.Lines.ItemCode = oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value
                            oPurchaseDeliveryNote.Lines.ItemDescription = oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value
                            oPurchaseDeliveryNote.Lines.Quantity = oMatrix.Columns.Item("colIQty").Cells.Item(i).Specific.Value
                            oPurchaseDeliveryNote.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colIPrice").Cells.Item(i).Specific.Value)
                            ' oPurchaseDeliveryNote.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
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
                oRecordset.DoQuery("SELECT DocStatus FROM OPOR Where DocEntry=" + FormatString(oActiveForm.DataSources.UserDataSources.Item("PONo").Value))
                If oRecordset.RecordCount > 0 Then
                    DocStatus = IIf(oRecordset.Fields.Item("DocStatus").Value = "C", "Closed", "Open")

                    oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oRecordset.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_POStatus = " + FormatString(DocStatus) + " WHERE U_PONo = " + FormatString(oActiveForm.DataSources.UserDataSources.Item("PONo").Value))
                End If
            End If
            CreateGoodsReceiptPO = True
        Catch ex As Exception
            CreateGoodsReceiptPO = False
        End Try
    End Function

    Private Sub CalAmtPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)

        '=============================================================================
        'Function   : CalAmtPO()
        'Purpose    : This function provide for calculate amount for purchae order itmes
        'Parameters : ByVal oActiveForm As SAPbouiCOM.Form,
        '             ByVal Row As Integer
        'Return     : No
        '============================================================================

        Dim cMatrix As SAPbouiCOM.Matrix
        oActiveForm.Freeze(True)
        cMatrix = oActiveForm.Items.Item("mx_Item").Specific
        cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(cMatrix.Columns.Item("colIPrice").Cells.Item(Row).Specific.Value)
        cMatrix.Columns.Item("colITotal").Editable = True
        cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value)
        'oActiveForm.Items.Item("ed_Code").Specific.Active = True
        cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Click()
        cMatrix.Columns.Item("colITotal").Editable = False
        oActiveForm.Freeze(False)
        CalculateTotalPO(oActiveForm, cMatrix)
    End Sub

    Private Sub CalRatePO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)


        '=============================================================================
        'Function   : CalRatePO()
        'Purpose    : This function provide for calculate Rate for purchae order itmes
        'Parameters : ByVal oActiveForm As SAPbouiCOM.Form,
        '             ByVal Row As Integer
        'Return     : No
        '=============================================================================

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

        '=============================================================================
        'Function   : CalculateTotalPO()
        'Purpose    : This function provide for calculate Total  amount for purchae order itmes
        'Parameters : ByVal oActiveForm As SAPbouiCOM.Form,
        '             ByRef oMatrix As SAPbouiCOM.Matrix
        'Return     : No
        '=============================================================================

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

    Private Function PopulatePurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean, ByRef oPOForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   PopulatePurchaseHeader()
        '   Purpose     :   This function  to insert data to main matrix 
        '                   from purchase form(PO)  matrix 
        '               
        '   Parameters  :  ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix,
        '                  ByVal pStrSQL As String, ByVal tblName As String,ByVal ProcessedState As Boolean        '                   
        '               
        '   Return      :  Fase - FAILURE
        '                  True - SUCCESS
        ' **********************************************************************************
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
                            If oPOForm.Title.Contains("Goods Receipt") = True Then 'MSW To Edit New Ticket
                                .SetValue("U_FDate", .Offset, oPOForm.Items.Item("ed_GRDate").Specific.Value)
                                .SetValue("U_FTime", .Offset, oPOForm.Items.Item("ed_GRTime").Specific.Value)
                            Else
                                .SetValue("U_FDate", .Offset, oPOForm.Items.Item("ed_PODate").Specific.Value)
                                .SetValue("U_FTime", .Offset, oPOForm.Items.Item("ed_POTime").Specific.Value)
                            End If
                            .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                            .SetValue("U_Remarks", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                        End With
                        pMatrix.AddRow()
                    Else
                        With ObjDbDataSource
                            .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(currentRow).Specific.Value)
                            .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                            .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                            .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_Vendor", .Offset, oRecordSet.Fields.Item("U_VCode").Value.ToString)
                            .SetValue("U_Place", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            If oPOForm.Title.Contains("Goods Receipt") = True Then 'MSW To Edit New Ticket
                                .SetValue("U_FDate", .Offset, oPOForm.Items.Item("ed_GRDate").Specific.Value)
                                .SetValue("U_FTime", .Offset, oPOForm.Items.Item("ed_GRTime").Specific.Value)
                            Else
                                .SetValue("U_FDate", .Offset, oPOForm.Items.Item("ed_PODate").Specific.Value)
                                .SetValue("U_FTime", .Offset, oPOForm.Items.Item("ed_POTime").Specific.Value)
                            End If
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

        ' **********************************************************************************
        '   Function    :   UpdatePurchaseOrder()
        '   Purpose     :   This function will be proceess to  Update Purchase
        '                   Order form Sap Data base table .
        '               
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String        '                       
        '                   sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   False- FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

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
                    'oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                    'oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    If oBusinessPartner.Currency = "##" Then
                        oPurchaseDocument.DocCurrency = "SGD"
                    Else
                        oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                    End If



                    oPurchaseDocument.DocDate = Now
                    oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                    oPurchaseDocument.TaxDate = Now
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
                            ' oPurchaseDocument.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
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
        'Function   : CreatePurchaseOrder()
        'Purpose    : This function to create Purchase Order items from purchae order fomr data from ExporeSeaCL Form
        'Parameters : ByRef oActiveForm As SAPbouiCOM.Form,
        '           : ByVal MatrixUID As String
        'Return     : Fase - FAILURE
        '           : True - SUCCESS

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
                If oBusinessPartner.Currency = "##" Then
                    oPurchaseDocument.DocCurrency = "SGD"
                Else
                    oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                End If

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
                p_oSBOApplication.SetStatusBarMessage("Error Code" + ret.ToString + " / " + sErrDesc.ToString, SAPbouiCOM.BoMessageTime.bmt_Short)
                'MessageBox.Show("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
                CreatePurchaseOrder = False
            Else
                oRecordset.DoQuery("SELECT DocEntry FROM OPOR Order By DocEntry")
                If oRecordset.RecordCount > 0 Then
                    oRecordset.MoveLast()
                    DocLastKey = oRecordset.Fields.Item("DocEntry").Value
                End If
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
        Dim oCombo As SAPbouiCOM.ComboBox
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
            oCombo = oMatrix.Columns.Item("colIGST").Cells.Item(oMatrix.RowCount).Specific
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sql = "Select ECVatGroup from OCRD where CardCode='" & oActiveForm.Items.Item("ed_Code").Specific.Value & "'"
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                If oRecordSet.Fields.Item("ECVatGroup").Value = "" Then
                    oCombo.Select("SI", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    oCombo.Select(oRecordSet.Fields.Item("ECVatGroup").Value.ToString, SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
            Else
                oCombo.Select("SI", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If

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
    Dim originpdf As String = ""
    Dim jobNo As String = ""
    Dim rptPath As String = ""
    Dim pdffilepath As String = ""
    Dim rptDocument As ReportDocument
    Private Sub PreviewPaymentVoucher(ByRef ParentForm As SAPbouiCOM.Form, ByRef oActiveForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************

        Dim PVNo As Integer
        rptDocument = New ReportDocument
        pdfFilename = "PAYMENT VOUCHER"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Payment Voucher.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        'If Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
        PVNo = Convert.ToInt32(oActiveForm.Items.Item("ed_DocNum").Specific.Value)
        'Else
        'PONo = DocLastKey
        ' End If

        rptDocument.SetParameterValue("@DocEntry", PVNo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If
        'reportuti.PrintDoc(rptDocument)
    End Sub

    Private Sub PreviewPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef oActiveForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************

        Dim PONo As Integer
        rptDocument = New ReportDocument
        pdfFilename = "PURCHASE ORDER"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Purchase Order.rpt"
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
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If
    End Sub

    Private Sub CreatePDF(ByRef ParentForm As SAPbouiCOM.Form, ByVal matrixName As String)

        ' **********************************************************************************
        '   Function    :   CreatePOPDF
        '   Purpose     :   This function will be providing to create  purchase
        '                   order PDF Form.
        '               
        '   Parameters  :   ByVal oActiveForm As SAPbouiCOM.Form,ByVal matrixName As String,ByVal iCode As String
        '                    
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************



        If matrixName = "mx_Bunk" Then
            pdfFilename = "Bunker"
            originpdf = "BunkTOPS.pdf"
        ElseIf matrixName = "mx_Fork" Then

            pdfFilename = "ForkLift"
            originpdf = "F&CSOPS.pdf"
        ElseIf matrixName = "mx_Crane" Then
            pdfFilename = "Crane"
            originpdf = "F&CSOPS.pdf"
        ElseIf matrixName = "mx_Orider" Then
            pdfFilename = "OutRider"
            originpdf = "OutRider (2).pdf"
        End If

        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
        Dim pdfoutputfile As FileStream = New FileStream(pdffilepath, System.IO.FileMode.Create)
        Dim formfiller As iTextSharp.text.pdf.PdfStamper = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
        Dim ac As iTextSharp.text.pdf.AcroFields = formfiller.AcroFields
        'ac.SetField("txtDocEntry", oActiveForm.Items.Item("ed_DocNum").Specific.Value)
        'ac.SetField("txtLineId", i.ToString)
        formfiller.Close()
        reader.Close()
        Process.Start(pdffilepath)
    End Sub

    Private Sub SendAttachFile(ByRef ParentForm As SAPbouiCOM.Form, ByRef oActiveForm As SAPbouiCOM.Form)
        '=============================================================================
        'Function   : SendAttachFile()
        'Purpose    : This function to sent email file and fax document
        '             and convert to pdf to image while senting fax documnet
        'Parameters : ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '              
        'Return     : No
        '=============================================================================


        'Save Report To specific JobFile as PDF File USE Code
        PreviewPO(ParentForm, oActiveForm)
        If oActiveForm.Items.Item("ch_Email").Specific.Checked = True Then
            'oRecordSet.DoQuery("Select email from ohem a inner join ousr b on a.userid=b.userid where user_Code='" & p_oDICompany.UserName.ToString() & "'")
            'If oRecordSet.RecordCount > 0 Then
            '    Dim frommail As String = oRecordSet.Fields.Item("email").Value.ToString
            '    If Not reportuti.SendMailDoc(oActiveForm.Items.Item("ed_Email").Specific.Value, frommail, "Purchase Order", "Purchase Order Item", pdffilepath) Then
            '        p_oSBOApplication.MessageBox("Send Message Fail", 0, "OK")
            '    Else
            '        p_oSBOApplication.MessageBox("Send Message Successfully", 0, "OK")
            '    End If
            'End If
        End If

        If oActiveForm.Items.Item("ch_Fax").Specific.Checked = True Then
            'ConvertFunction(oActiveForm, pdffilepath)
        End If

        rptDocument.Close()
    End Sub

    Private Function CreatePurchaseOrderPDF(ByRef oActiveForm As SAPbouiCOM.Form, ByVal CardCode As String, ByVal iCode As String) As Boolean
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
        Dim dblPrice As Double = 0.0
        Dim sErrDesc As String = vbNullString
        CreatePurchaseOrderPDF = False
        Try
            oPurchaseDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oBusinessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            p_oDICompany.GetNewObjectCode("22")

            If oBusinessPartner.GetByKey(CardCode) Then
                oRecordset.DoQuery("Select CardName From OCRD Where CardCode= '" & CardCode & "'")
                oPurchaseDocument.CardCode = CardCode
                oPurchaseDocument.CardName = oRecordset.Fields.Item("CardName").Value
                'oPurchaseDocument.ContactPersonCode = GetContactPersonCode(oRecordset, Trim(oActiveForm.Items.Item("cb_Contact").Specific.Value.ToString), oActiveForm.Items.Item("ed_Code").Specific.Value.ToString)
                If oBusinessPartner.Currency = "##" Then
                    oPurchaseDocument.DocCurrency = "SGD"
                Else
                    oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                End If

                oPurchaseDocument.DocDate = Now
                oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                oPurchaseDocument.TaxDate = Now
                oRecordset.DoQuery("Select ITM1.Price,OITM.ItemCode,OITM.ItemName from OITM Inner Join ITM1 on OITM.itemcode=ITM1.itemcode Where OITM.itemcode='" & iCode & "' and ITM1.pricelist=1")
                If oRecordset.RecordCount > 0 Then
                    'dblPrice = oRecordset.Fields.Item("Price").Value
                    oPurchaseDocument.Lines.ItemCode = oRecordset.Fields.Item("ItemCode").Value
                    oPurchaseDocument.Lines.ItemDescription = oRecordset.Fields.Item("ItemName").Value
                    oPurchaseDocument.Lines.Quantity = 1
                    oPurchaseDocument.Lines.UnitPrice = oRecordset.Fields.Item("Price").Value
                    'oPurchaseDocument.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colIAmt").Cells.Item(i).Specific.Value)
                    '   oPurchaseDocument.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                    ' If oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value = "None" Then
                    oPurchaseDocument.Lines.VatGroup = "ZI"
                    'Else
                    ' oPurchaseDocument.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                    'End If
                    oPurchaseDocument.Lines.Add()
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
            CreatePurchaseOrderPDF = True
        Catch ex As Exception
            CreatePurchaseOrderPDF = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function PopulatePurchaseHeaderPDF(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean) As Boolean

        ' **********************************************************************************
        '   Function    :   PopulatePurchaseHeaderPDF
        '   Purpose     :   This function will be providing to create  purchase
        '                   order PDF Form.
        '               
        '   Parameters  :   ByVal oActiveForm As SAPbouiCOM.Form,ByVal matrixName As String,ByVal cardcode As String
        '               :   ByVal iCode As String 
        '                    
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

        PopulatePurchaseHeaderPDF = False
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
            PopulatePurchaseHeaderPDF = True
        Catch ex As Exception
            PopulatePurchaseHeaderPDF = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function GetJobNumber(ByVal prefix As String) As String

        ' **********************************************************************************
        '   Function    :   GetJobNumber
        '   Purpose     :   This function will be providing to create  new job number
        '                   for purchase order process.
        '               
        '   Parameters  :   ByVal prefix As String
        '                    
        '               
        '   Return      :  No
        ' **********************************************************************************


        GetJobNumber = vbNullString
        Try
            Dim jobSrNo As Integer = 0
            Dim postFix As String = String.Empty
            Dim strJobNo As String = String.Empty
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("select U_JobNo from [@OBT_FREIGHTDOCNO] Order by docentry asc")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveLast()
                jobSrNo = Convert.ToInt32(Right(oRecordSet.Fields.Item("U_JobNo").Value, 6)) + 1
            Else
                jobSrNo = 1
            End If
            For i = 1 To 6 - jobSrNo.ToString.Length
                postFix = postFix + "0"
            Next
            GetJobNumber = prefix + Now.ToString("yyyy") + postFix + jobSrNo.ToString
        Catch ex As Exception
            GetJobNumber = vbNullString
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Sub LoadImportSeaLCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)

        ' **********************************************************************************
        '   Function    :   LoadImportSeaLCLForm
        '   Purpose     :   This function will be providing to load and show ImportSeaLCLForm        '                
        '               
        '   Parameters  :   Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString,
        '               :   Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        '                    
        '               
        '   Return      :  No
        ' **********************************************************************************

        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim sErrDesc As String = String.Empty
        Dim jobType As String = String.Empty
        Try
            If Left(Title, 5) = "Local" Then
                If Not LoadFromXML(p_oSBOApplication, "Localv1.srf") Then Throw New ArgumentException(sErrDesc)
            ElseIf Title.Substring(7, 3) = "Sea" Then
                If Not LoadFromXML(p_oSBOApplication, "ImportSeaLCLv1.srf") Then Throw New ArgumentException(sErrDesc)
            ElseIf Title.Substring(7, 3) = "Air" Then
                If Not LoadFromXML(p_oSBOApplication, "ImportAirv1.srf") Then Throw New ArgumentException(sErrDesc)
            ElseIf Title.Substring(7, 4) = "Land" Then
                If Not LoadFromXML(p_oSBOApplication, "ImportLandv1.srf") Then Throw New ArgumentException(sErrDesc)
            End If

            'If Not LoadFromXML(p_oSBOApplication, "ImportSeaLCLv1.srf") Then Throw New ArgumentException(sErrDesc)

            'LoadFromXML(p_oSBOApplication, "ImportSeaLCLv1.srf")


            If AlreadyExist("IMPORTSEALCL") Then
                ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")
                jobType = "Import Sea LCL"
            ElseIf AlreadyExist("LOCAL") Then
                ImportSeaLCLForm = p_oSBOApplication.Forms.Item("LOCAL")
                jobType = "Local"
            ElseIf AlreadyExist("IMPORTAIR") Then
                ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTAIR")
                jobType = "Import Air"
            ElseIf AlreadyExist("IMPORTLAND") Then
                ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTLAND")
                jobType = "Import Land"
            End If

            'ImportSeaLCLForm = p_oSBOApplication.Forms.Item("IMPORTSEALCL")

            ImportSeaLCLForm.Title = Title ' MSW To Edit New Ticket 07-09-2011
            ImportSeaLCLForm.EnableMenu("1288", True)
            ImportSeaLCLForm.EnableMenu("1289", True)
            ImportSeaLCLForm.EnableMenu("1290", True)
            ImportSeaLCLForm.EnableMenu("1291", True)
            ImportSeaLCLForm.EnableMenu("1284", False)
            ImportSeaLCLForm.EnableMenu("1286", False)
            ImportSeaLCLForm.EnableMenu("1281", True)
            ImportSeaLCLForm.EnableMenu("1283", False)
            ImportSeaLCLForm.EnableMenu("1292", False)
            ImportSeaLCLForm.EnableMenu("1293", False)
            ImportSeaLCLForm.EnableMenu("4870", False)
            ImportSeaLCLForm.EnableMenu("771", False)
            ImportSeaLCLForm.EnableMenu("772", True)
            ImportSeaLCLForm.EnableMenu("773", True)
            ImportSeaLCLForm.EnableMenu("775", False)
            ImportSeaLCLForm.EnableMenu("774", False)


            If FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            End If
            ImportSeaLCLForm.DataBrowser.BrowseBy = "ed_DocNum"
            ImportSeaLCLForm.Items.Item("fo_Prmt").Specific.Select()
            'ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = True

            ImportSeaLCLForm.Freeze(True)

            ImportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
            EnabledHeaderControls(ImportSeaLCLForm, False)
            EnabledMaxtix(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("mx_TkrList").Specific, False)
            ImportSeaLCLForm.PaneLevel = 7

            If FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If Not Title = vbNullString Then
                    ImportSeaLCLForm.Title = Title
                End If
                ImportSeaLCLForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString 'Prep By for Header
                ImportSeaLCLForm.Items.Item("ed_JType").Specific.Value = jobType
                ' ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_LCL06_PMAIN ").SetValue("U_TranMode", 0, "Sea") 'MSW LCL Change
                ImportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                ImportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                oCombo = ImportSeaLCLForm.Items.Item("cb_TMode").Specific
                If ImportSeaLCLForm.Title.Substring(7, 3) = "Air" Then
                    oCombo.Select("Air", SAPbouiCOM.BoSearchKey.psk_ByValue)
                ElseIf ImportSeaLCLForm.Title.Substring(7, 4) = "Land" Then
                    oCombo.Select("Land", SAPbouiCOM.BoSearchKey.psk_ByValue)
                Else
                    oCombo.Select("Sea", SAPbouiCOM.BoSearchKey.psk_ByValue)
                End If
                ImportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                ImportSeaLCLForm.Items.Item("ed_Code").Specific.Active = True
            End If
            If HolidayMarkUp(ImportSeaLCLForm, ImportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ImportSeaLCLForm.Items.Item("ed_JbDay").Specific, ImportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddChooseFromList(ImportSeaLCLForm, "cflBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "cflBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ImportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListUID = "cflBP"
            ImportSeaLCLForm.Items.Item("ed_Code").Specific.ChooseFromListAlias = "CardCode"
            ImportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListUID = "cflBP2"
            ImportSeaLCLForm.Items.Item("ed_Name").Specific.ChooseFromListAlias = "CardName"

            If AddChooseFromList(ImportSeaLCLForm, "cflBP3", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "WRHSE", False, "UDOWH") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "DSVES01", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ImportSeaLCLForm, "DSVES02", False, "VESSEL") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
            ImportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListUID = "WRHSE"
            ImportSeaLCLForm.Items.Item("ed_Wrhse").Specific.ChooseFromListAlias = "Code"
            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
            ImportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
            ' ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
            'ImportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"

            'MSW to add

            ''If AddChooseFromListCondition(ImportSeaLCLForm, "CVendor", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S", "U_VType", "Crane") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ''ImportSeaLCLForm.Items.Item("ed_CVendor").Specific.ChooseFromListUID = "CVendor"
            ''ImportSeaLCLForm.Items.Item("ed_CVendor").Specific.ChooseFromListAlias = "CardCode"


            ''If AddChooseFromListCondition(ImportSeaLCLForm, "FVendor", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S", "U_VType", "ForkLift") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ''ImportSeaLCLForm.Items.Item("ed_FVendor").Specific.ChooseFromListUID = "FVendor"
            ''ImportSeaLCLForm.Items.Item("ed_FVendor").Specific.ChooseFromListAlias = "CardCode"
            'End MSW to add

            '-------------------------------For Cargo Tab OMM & SYMA------------------------------------------------'13 Jan 2011

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
            'Get Job No
            If Not ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("IM")
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
            If AddUserDataSrc(ImportSeaLCLForm, "EUC", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) ' MSW To Edit New Ticket

            'MSW 14-09-2011 Truck PO
            If AddUserDataSrc(ImportSeaLCLForm, "TKRINS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRIRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "TKRRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ImportSeaLCLForm, "PODOCNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'MSW 14-09-2011 Truck PO

            ImportSeaLCLForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
            ImportSeaLCLForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
            ImportSeaLCLForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
            ImportSeaLCLForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
            ImportSeaLCLForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
            ImportSeaLCLForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")
            ImportSeaLCLForm.Items.Item("ed_EUC").Specific.DataBind.SetBound(True, "", "EUC") ' MSW To Edit New Ticket
            ImportSeaLCLForm.Items.Item("ee_InsRmsk").Specific.DataBind.SetBound(True, "", "TKRIRMK") 'MSW 14-09-2011 Truck PO
            ImportSeaLCLForm.Items.Item("ee_TkrIns").Specific.DataBind.SetBound(True, "", "TKRINS") 'MSW 14-09-2011 Truck PO
            ImportSeaLCLForm.Items.Item("ee_Rmsk").Specific.DataBind.SetBound(True, "", "TKRRMK") 'MSW 14-09-2011 Truck PO
            ImportSeaLCLForm.Items.Item("ed_PODocNo").Specific.DataBind.SetBound(True, "", "PODOCNO") 'MSW 14-09-2011 Truck PO



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
            oMatrix = ImportSeaLCLForm.Items.Item("mx_License").Specific
            oMatrix.AddRow()
            oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
            '-------------------------------------------------------------------------------------'


            If ImportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = True
                ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = JobNo
                ImportSeaLCLForm.Items.Item("1").Click()
                ImportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False
                If ImportSeaLCLForm.Title.Substring(7, 3) = "Air" Then
                    ImportSeaLCLForm.Title = "Import Air " + ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value ' MSW To Edit New Ticket 07-09-2011
                ElseIf ImportSeaLCLForm.Title.Substring(7, 4) = "Land" Then
                    ImportSeaLCLForm.Title = "Import Land " + ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value ' MSW To Edit New Ticket 07-09-2011
                Else
                    ImportSeaLCLForm.Title = "Import Sea-LCL " + ImportSeaLCLForm.Items.Item("cb_JobType").Specific.Value ' MSW To Edit New Ticket 07-09-2011
                End If
            End If
            If ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
            ElseIf ImportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                ImportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
            Else
                ImportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
            End If

            ImportSeaLCLForm.Freeze(False)
            Select Case ImportSeaLCLForm.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_AddIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                    ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                    ImportSeaLCLForm.Items.Item("bt_A6Label").Enabled = False
                Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    ImportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_AddIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                    ImportSeaLCLForm.Items.Item("bt_A6Label").Enabled = True

                    ImportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_CGR").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_ForkPO").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_CranePO").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                    ImportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                    If (ImportSeaLCLForm.Title.Substring(7, 3) = "Air" Or ImportSeaLCLForm.Title.Substring(7, 4) = "Land") Then
                        ImportSeaLCLForm.Items.Item("bt_Orider").Enabled = True
                    End If
            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    'MSW To Edit
    Private Sub CalculateNoOfBoxes(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix)
        If pMatrix.RowCount > 1 Then
            For i As Integer = Convert.ToInt32(pForm.Items.Item("ed_ItemNo").Specific.Value) To pMatrix.RowCount - 1
                pMatrix.Columns.Item("colBox").Cells.Item(i + 1).Specific.Value() = pMatrix.Columns.Item("colBoxLast").Cells.Item(i).Specific.Value() + 1
                pMatrix.Columns.Item("colBoxLast").Cells.Item(i + 1).Specific.Value() = (Convert.ToDouble(pMatrix.Columns.Item("colBox").Cells.Item(i + 1).Specific.Value()) + Convert.ToDouble(pMatrix.Columns.Item("colTNBox").Cells.Item(i + 1).Specific.Value())) - 1
            Next
        End If

    End Sub
    'MSW To Edit New Ticket
    Private Function CheckQtyValue(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal formName As String) As Boolean
        CheckQtyValue = False
        For i As Integer = 1 To oMatrix.RowCount
            If formName = "Purchase Order" Then
                If oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value <> "" And Convert.ToDouble(oMatrix.Columns.Item("colIQty").Cells.Item(i).Specific.Value) = 0.0 Then
                    CheckQtyValue = True
                    Exit Function
                End If
            ElseIf formName = "Payment Voucher" Then
                If oMatrix.Columns.Item("colChCode1").Cells.Item(i).Specific.Value <> "" And Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(i).Specific.Value) = 0.0 Then
                    CheckQtyValue = True
                    Exit Function
                End If
            End If

        Next
        CheckQtyValue = False
    End Function

    Private Sub EnabledTruckerForExternal(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)

        ' **********************************************************************************
        '   Function    :   EnabledTruckerForExternal
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
        Try
            pForm.Freeze(True)
            pForm.Items.Item("ed_Trucker").Specific.Active = True
            pForm.Items.Item("ed_EUC").Enabled = pValue
            pForm.Items.Item("ed_Attent").Enabled = pValue
            pForm.Items.Item("ed_TkrTel").Enabled = pValue
            pForm.Items.Item("ed_Fax").Enabled = pValue
            pForm.Items.Item("ed_Email").Enabled = pValue
            pForm.Items.Item("ed_TkrDate").Enabled = pValue
            pForm.Items.Item("ed_TkrTime").Enabled = pValue
            pForm.Items.Item("ee_ColFrm").Enabled = pValue
            pForm.Items.Item("ee_TkrTo").Enabled = pValue
            pForm.Items.Item("ee_TkrIns").Enabled = pValue
            pForm.Items.Item("ee_Rmsk").Enabled = pValue
            pForm.Items.Item("ee_InsRmsk").Enabled = pValue
            pForm.Items.Item("ed_Trucker").Specific.Active = True
            pForm.Freeze(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

#Region "Trucking PO"
    Private Sub LoadTruckingPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)

        ' **********************************************************************************
        '   Function    :   LoadAndCreateCPO()
        '   Purpose     :   This function will be providing to create and show Purchase Order form 
        '                   for Purchase Order form  
        '               
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,                       
        '                   ByRef srfName As String
        '               
        '   Return      :   No
        '                   
        ' **********************************************************************************

        Dim CPOForm As SAPbouiCOM.Form
        Dim CPOMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Try
            LoadFromXML(p_oSBOApplication, srfName)
            CPOForm = p_oSBOApplication.Forms.ActiveForm
            CPOForm.Freeze(True)
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

                CPOForm.Items.Item("ed_TrkTo").Specific.Value = ParentForm.Items.Item("ee_TkrTo").Specific.Value


                If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_DocNum").Specific.Value)
                CPOForm.Items.Item("bt_Preview").Visible = False
                CPOForm.Items.Item("bt_Resend").Visible = False

                CPOForm.Items.Item("ed_Code").Specific.Active = True
                CPOForm.Freeze(False)
                ' ==================================== Custom Purchase Order ========================================
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Function PopulateTruckingPOToEditTab(ByVal pForm As SAPbouiCOM.Form, ByVal pStrSQL As String, ByVal oPOForm As SAPbouiCOM.Form) As Boolean
        PopulateTruckingPOToEditTab = False

        Dim sCode As String = String.Empty
        Dim sName As String = String.Empty
        Dim sAttention As String = String.Empty
        Dim sPhone As String = String.Empty
        Dim sFax As String = String.Empty
        Dim sMail As String = String.Empty
        Dim UEN As String = String.Empty
        Dim oEditText As SAPbouiCOM.EditText

        Try
            ' oRecordSet.DoQuery("SELECT CardCode,CardName FROM OPOR WHERE DocEntry = '" + pForm.Items.Item("ed_Code").Specific.Value + "'")
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If oRecordSet.RecordCount > 0 Then
                sCode = oRecordSet.Fields.Item("U_VCode").Value.ToString
                'sName = oRecordSet.Fields.Item("CardName").Value.ToString
            End If
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            pForm.DataSources.UserDataSources.Item("PODOCNO").ValueEx = oPOForm.Items.Item("ed_CPOID").Specific.Value
            oRecordSet.DoQuery("select  OCRD.CardName,OCPR.Name,OCPR.Tel1,OCPR.Fax,OCPR.E_MailL,OCRD.VatIdUnCmp from OCPR LEFT OUTER JOIN OCRD ON OCPR.Name = OCRD.CntctPrsn where OCRD.CardCode = '" + sCode + "'")
            If oRecordSet.RecordCount > 0 Then
                sName = oRecordSet.Fields.Item("CardName").Value.ToString
                sAttention = oRecordSet.Fields.Item("Name").Value.ToString
                sPhone = oRecordSet.Fields.Item("Tel1").Value.ToString
                sFax = oRecordSet.Fields.Item("Fax").Value.ToString
                sMail = oRecordSet.Fields.Item("E_MailL").Value.ToString
                UEN = oRecordSet.Fields.Item("VatIdUnCmp").Value.ToString  ' MSW To Edit New Ticket
            End If
            pForm.Items.Item("ed_PONo").Specific.Value = DocLastKey
            oEditText = pForm.Items.Item("ed_Trucker").Specific
            oEditText.DataBind.SetBound(True, "", "TKREXTR")
            oEditText.ChooseFromListUID = "CFLTKRV"
            oEditText.ChooseFromListAlias = "CardName"
            pForm.DataSources.UserDataSources.Item("TKREXTR").ValueEx = sName
            pForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = sAttention
            pForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = sPhone
            pForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = sFax
            pForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = sMail
            pForm.DataSources.UserDataSources.Item("EUC").ValueEx = UEN
            sql = "Select U_PORMKS,U_POIRMKS,U_ColFrm,U_TkrTo,U_TkrIns from [@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                pForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = oRecordSet.Fields.Item("U_ColFrm").Value
                pForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("U_TkrTo").Value
                pForm.DataSources.UserDataSources.Item("TKRINS").ValueEx = oRecordSet.Fields.Item("U_TkrIns").Value
                pForm.DataSources.UserDataSources.Item("TKRIRMK").ValueEx = oRecordSet.Fields.Item("U_POIRMKS").Value
                pForm.DataSources.UserDataSources.Item("TKRRMK").ValueEx = oRecordSet.Fields.Item("U_PORMKS").Value
            End If
            PopulateTruckingPOToEditTab = True
        Catch ex As Exception
            PopulateTruckingPOToEditTab = False
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function PopulateTruckPurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String) As Boolean

        ' **********************************************************************************
        '   Function    :   PopulatePurchaseHeader()
        '   Purpose     :   This function  to insert data to main matrix 
        '                   from purchase form(PO)  matrix 
        '               
        '   Parameters  :  ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix,
        '                  ByVal pStrSQL As String, ByVal tblName As String,ByVal ProcessedState As Boolean        '                   
        '               
        '   Return      :  Fase - FAILURE
        '                  True - SUCCESS
        ' **********************************************************************************
        PopulateTruckPurchaseHeader = False
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

                    With ObjDbDataSource
                        'MSW 14-09-2011 Truck PO
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_1").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_InsDocNo", .Offset, pMatrix.Columns.Item("V_1").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                        '.SetValue("U_PONo", .Offset, DocLastKey)
                        .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                        .SetValue("U_InsDate", .Offset, CDate(oRecordSet.Fields.Item("U_InsDate").Value).ToString("yyyyMMdd"))

                        .SetValue("U_Mode", .Offset, oRecordSet.Fields.Item("U_Mode").Value.ToString)
                        .SetValue("U_Trucker", .Offset, oRecordSet.Fields.Item("U_Trucker").Value.ToString)
                        .SetValue("U_VehNo", .Offset, oRecordSet.Fields.Item("U_VehNo").Value.ToString)
                        .SetValue("U_EUC", .Offset, oRecordSet.Fields.Item("U_EUC").Value.ToString)
                        .SetValue("U_Attent", .Offset, oRecordSet.Fields.Item("U_Attent").Value.ToString)
                        .SetValue("U_Tel", .Offset, oRecordSet.Fields.Item("U_Tel").Value.ToString)
                        .SetValue("U_Fax", .Offset, oRecordSet.Fields.Item("U_Fax").Value.ToString)
                        .SetValue("U_Email", .Offset, oRecordSet.Fields.Item("U_Email").Value.ToString)
                        .SetValue("U_TkrDate", .Offset, CDate(oRecordSet.Fields.Item("U_TkrDate").Value).ToString("yyyyMMdd"))
                        .SetValue("U_TkrTime", .Offset, oRecordSet.Fields.Item("U_TkrTime").Value)
                        .SetValue("U_ColFrm", .Offset, oRecordSet.Fields.Item("U_ColFrm").Value.ToString)
                        .SetValue("U_TkrTo", .Offset, oRecordSet.Fields.Item("U_TkrTo").Value.ToString)
                        .SetValue("U_TkrIns", .Offset, oRecordSet.Fields.Item("U_TkrIns").Value.ToString)
                        .SetValue("U_InsRemsk", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                        .SetValue("U_Remarks", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString) 'MSW 14-09-2011 Truck PO
                        .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString) 'MSW 14-09-2011 Truck PO
                        .SetValue("U_PrepBy", .Offset, p_oDICompany.UserName.ToString)
                    End With
                    pMatrix.SetLineData(currentRow)

                    oRecordSet.MoveNext()
                Loop
            End If
            PopulateTruckPurchaseHeader = True
        Catch ex As Exception
            PopulateTruckPurchaseHeader = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function PopulateOtherPurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String) As Boolean

        ' **********************************************************************************
        '   Function    :   PopulatePurchaseHeader()
        '   Purpose     :   This function  to insert data to main matrix 
        '                   from purchase form(PO)  matrix 
        '               
        '   Parameters  :  ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix,
        '                  ByVal pStrSQL As String, ByVal tblName As String,ByVal ProcessedState As Boolean        '                   
        '               
        '   Return      :  Fase - FAILURE
        '                  True - SUCCESS
        ' **********************************************************************************
        PopulateOtherPurchaseHeader = False
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
                    With ObjDbDataSource
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                        '.SetValue("U_PONo", .Offset, DocLastKey)
                        .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                        .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                        .SetValue("U_Vendor", .Offset, oRecordSet.Fields.Item("U_VCode").Value.ToString)
                        .SetValue("U_Place", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                        .SetValue("U_FDate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                        .SetValue("U_FTime", .Offset, oRecordSet.Fields.Item("U_POTime").Value)
                        .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                        .SetValue("U_Remarks", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                    End With
                    pMatrix.SetLineData(currentRow)

                    oRecordSet.MoveNext()
                Loop
            End If
            PopulateOtherPurchaseHeader = True
        Catch ex As Exception
            PopulateOtherPurchaseHeader = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function CancelTruckingPurchaseOrder(ByVal PONo As Integer) As Boolean

        ' **********************************************************************************
        '   Function    :   CancelTruckingPurchaseOrder()
        '   Purpose     :   This function will be proceess to  Update Purchase
        '                   Order form Sap Data base table .
        '               
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String        '                       
        '                   sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   False- FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

        CancelTruckingPurchaseOrder = False
        Dim oPurchaseDocument As SAPbobsCOM.Documents
        Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Try
            oPurchaseDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oBusinessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oPurchaseDocument.GetByKey(PONo) Then
                Dim ret As Long = oPurchaseDocument.Cancel
                If ret <> 0 Then
                    p_oDICompany.GetLastError(ret, sErrDesc)
                    Debug.Print("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
                Else
                    'p_oSBOApplication.MessageBox("Purchase Order document is successfully Cancelled!")
                End If
            End If
            CancelTruckingPurchaseOrder = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            CancelTruckingPurchaseOrder = False
        End Try
    End Function
#End Region

    Private Function CancelPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal sql As String, ByVal matrixName As String) As Boolean

        ' **********************************************************************************
        '   Function    :   CancelPO()
        '   Purpose     :   This function will be proceess to  Cancel Purchase
        '                   Order form Sap Data base table .
        '               
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String        '                       
        '                   sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   False- FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim dbSource As String = String.Empty
        Dim sErrDesc As String = String.Empty
        oMatrix = oActiveForm.Items.Item(matrixName).Specific
        Try
            CancelPO = False
            oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            oActiveForm.Items.Item("2").Specific.Caption = "Close"
            If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value)) Then Throw New ArgumentException(sErrDesc)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_POStatus = 'Cancelled' WHERE U_PONo = " + FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value))
            If matrixName = "mx_TkrList" Then
                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_LCL03_TRUCKING") Then Throw New ArgumentException(sErrDesc)
            Else
                If matrixName = "mx_Fork" Then
                    dbSource = "@OBT_LCL22_FORKLIFT"
                ElseIf matrixName = "mx_Armed" Then
                    dbSource = "@OBT_LCL24_ARMES"
                ElseIf matrixName = "mx_Crane" Then
                    dbSource = "@OBT_LCL25_CRANE"
                ElseIf matrixName = "mx_Bunk" Then
                    dbSource = "@OBT_LCL26_BUNKER"
                ElseIf matrixName = "mx_Orider" Then
                    dbSource = "@OBT_LCL27_OUTRIDER"
                End If
                If Not PopulateOtherPurchaseHeader(oActiveForm, oMatrix, sql, dbSource) Then Throw New ArgumentException(sErrDesc)
            End If
            oActiveForm.Items.Item(1).Click()
            CancelPO = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            CancelPO = False
        End Try
    End Function

    'KM to edit
    Private Sub PreviewDispatchInstruction(ByRef ParentForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************

        'Dim DocNum As Integer
        'rptDocument = New ReportDocument
        'pdfFilename = "Dispatch Instruction"
        'mainFolder = p_fmsSetting.DocuPath
        'jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        'rptPath = Application.StartupPath.ToString & "\Dispatch Instruction ImportLCL.rpt"
        'pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        'rptDocument.Load(rptPath)
        'rptDocument.Refresh()
        'DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        'rptDocument.SetParameterValue("@DocEntry", DocNum)
        'reportuti.SetDBLogIn(rptDocument)
        'If Not pdffilepath = String.Empty Then
        '    reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument,true)
        'End If
    End Sub

    Private Sub PreviewInsDoc(ByRef ParentForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************

        Dim DocNum As Integer
        Dim InsDoc As Integer
        rptDocument = New ReportDocument
        pdfFilename = "Trucking Instruction"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Trucking Instruction ImportLCL.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()

        DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        InsDoc = Convert.ToInt32(ParentForm.Items.Item("ed_InsDoc").Specific.Value)

        rptDocument.SetParameterValue("@DocEntry", DocNum)
        rptDocument.SetParameterValue("@InsDocNo", InsDoc)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If
    End Sub

    Private Sub PreviewA6Label(ByRef ParentForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************

        Dim DocNum As Integer
        rptDocument = New ReportDocument
        pdfFilename = "A6 Label"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\A6 Label Import Sea LCL.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        rptDocument.SetParameterValue("@DocEntry", DocNum)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If
    End Sub

  

  
End Module


