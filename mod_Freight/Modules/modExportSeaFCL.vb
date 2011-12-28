Option Explicit On

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Drawing.Printing
Imports System.Threading

Module modExportSeaFCL
    Private DocLastKey As String
    Private currentRow As Integer
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim DocStatus As String = String.Empty
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Private ShpInvoice, PO, Details, ShipTo, AWB, Box, Weight, POD, DesCargo As String
    Private Total, GSTAmt, SubTotal, vocTotal, gstTotal As Double
    Private vendorCode As String
    Dim strAPInvNo, strOutPayNo As String
    Dim sql As String = ""
    Private sPicName, sImgePath As String
    Private ActiveForm As SAPbouiCOM.Form
    Dim dtmatrix As SAPbouiCOM.DataTable
    Private gridindex As String
    Private ActiveMatrix As String
    Private Activesrf As String
    Dim vedCurCode As String = String.Empty
    Private LineId, ConSeqNo, ConNo, ConSealNo, ConSize, ConType, ConWt, ConDesc, ConDate, ConDay, ConTime, Conunstuff, ChStuff As String
    Dim MainForm As SAPbouiCOM.Form = Nothing 'Fumigation

    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String
    Dim IsAmend As Boolean = False
    Dim selectedRow As Integer
    Dim sErrDesc As String
    Private ObjMatrix As SAPbouiCOM.Matrix
    Private DspLineId, DocumentNo, PurchaseOrderNo, PurchaseDocNo, POSerialNo, POStus, PrepBy, DMulti, DOrigin, PODate, InstructionDate, Mode, DspCode, Dispatch, EUC, Attention, Telephone, Fax, Email, DspDate, DspTime, DspIns, DspIRemarks, Remarks, PreparedBy As String
    Dim JobStus As String = String.Empty

    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public Function DoExportSeaFCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   DoImportSeaFCLFormDataEvent
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
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim ExportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oPOForm As SAPbouiCOM.Form = Nothing
        Dim oGRForm As SAPbouiCOM.Form = Nothing
        Dim FunctionName As String = "DoExportSeaFCLFormDataEvent"
        Dim sKeyValue As String = String.Empty
        Dim sMessage As String = String.Empty
        Dim sSQLQuery As String = String.Empty
        Dim oDocument As SAPbobsCOM.Documents
        Dim oXmlReader As XmlTextReader
        Dim sDocNum As String = String.Empty
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oChMatrix As SAPbouiCOM.Matrix
        Dim oMatrixName As String = ""
        Dim TableName As String = ""
        Dim tblHeader As String = ""
        Dim source As String = ""
        Dim jMode As String = ""
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", FunctionName)
            Select Case BusinessObjectInfo.FormTypeEx

                Case "2000000021", "2000000009", "2000000010", "2000000043", "2000000050", "2000000052"  'MSW to Edit New Ticket
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oPOForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        'If BusinessObjectInfo.FormUID = "PURCHASEORDER" Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            'When action of BusinessObjectInfo is nicely done, need to do 2 tasks in action for Purcahse process
                            ' (1). need to show PopulatePurchaseHeader into the matrix of Main Export Form
                            ' (2). need to create PurchaseOrder into OPOR and POR1, related with main PurchaseProcess by using oPurchaseOrder document
                            If AlreadyExist("EXPORTSEAFCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                            End If

                            'oMatrix = oActiveForm.Items.Item("mx_Fumi").Specific
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                        "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If Not CreatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            If BusinessObjectInfo.FormTypeEx = "2000000020" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crate").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL13_CRATE", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL15_BUNKER", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL16_ARMESCORT", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000043" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bok").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL12_BOOKING", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                'oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                'If Not PopulateTruckingPOToEditTab(oActiveForm, sql) Then Throw New ArgumentException(sErrDesc)
                                ''MSW 14-09-2011 Truck PO
                            End If
                            'If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql) Then Throw New ArgumentException(sErrDesc)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value))
                            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
                            If BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                'CreatePOPDF(oActiveForm, "mx_Armed", "")
                            Else
                                SendAttachFile(oActiveForm, oPOForm)
                            End If

                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                            If AlreadyExist("EXPORTSEAFCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                            End If
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                       "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If BusinessObjectInfo.FormTypeEx = "2000000020" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crate").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL13_CRATE", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL15_BUNKER", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL16_ARMESCORT", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000043" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bok").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL12_BOOKING", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_TkrCode,b.U_Trucker,b.U_VehNo," & _
                                        "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                        "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL03_ETRUCKING] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_FCL03_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000052" Then
                                oMatrix = oActiveForm.Items.Item("mx_DspList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_DspCode,b.U_Dispatch," & _
                                        "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_DspDate,b.U_DspTime " & _
                                        "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL04_EDISPATCH] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                                If Not PopulateDispatchPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_FCL04_EDISPATCH") Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Dispatch PO
                            End If
                            If Not UpdatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)

                        End If
                        'End If
                    End If

                Case "2000000013", "2000000014", "2000000044", "2000000051", "2000000053"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oGRForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("EXPORTSEAFCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRFCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                        End If
                        'oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            If Not CreateGoodsReceiptPO(oGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)

                            If BusinessObjectInfo.FormTypeEx = "2000000011" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crate").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL13_CRATE", False, oGRForm) Then Throw New ArgumentException(sErrDesc)

                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000016" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL15_BUNKER", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000013" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL16_ARMESCORT", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000044" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bok").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_FCL12_BOOKING", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000051" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_TkrCode,b.U_Trucker,b.U_VehNo," & _
                                       "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime,b.U_PO " & _
                                       "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL03_ETRUCKING] b on a.DocEntry=b.U_PODocNo where a.U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_FCL03_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                If oMatrix.Columns.Item("colMulti").Cells.Item(currentRow).Specific.Value() = "Y" Then
                                    modMultiJobForNormal.UpdateMultiJobPOStatus(ExportSeaFCLForm, "Closed", oMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value, "[@OBT_FCL03_ETRUCKING]")
                                End If
                                oMatrix = oActiveForm.Items.Item("mx_PO").Specific
                                If Not EditPOTab(oActiveForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc)

                                'Truck PO

                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000053" Then
                                oMatrix = oActiveForm.Items.Item("mx_DspList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_DspCode,b.U_Dispatch," & _
                                       "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_DspDate,b.U_DspTime,b.U_PO " & _
                                       "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL04_EDISPATCH] b on a.DocEntry=b.U_PODocNo where a.U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulateDispatchPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_FCL04_EDISPATCH") Then Throw New ArgumentException(sErrDesc)
                                If oMatrix.Columns.Item("colMulti").Cells.Item(currentRow).Specific.Value() = "Y" Then
                                    modMultiJobForNormal.UpdateMultiJobPOStatus(ExportSeaFCLForm, "Closed", oMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value, "[@OBT_FCL04_EDISPATCH]")
                                End If
                                oMatrix = oActiveForm.Items.Item("mx_PO").Specific
                                If Not EditPOTab(oActiveForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc)
                                'Dispatch PO
                            End If

                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then

                        End If
                    End If

                Case "2000000026", "2000000030", "2000000032", "2000000035", "2000000038", "2000000038", "2000000041", "2000000005", "2000000015", "2000000060", "2000000021", "2000000020"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oPOForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("EXPORTSEAFCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRFCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                        End If

                        Select Case BusinessObjectInfo.FormUID

                            Case "CRANEPURCHASEORDER"
                                oMatrixName = "mx_Crane"
                                TableName = "@OBT_TB33_CRANE"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                source = "Crane"
                            Case "BUNKPURCHASEORDER"
                                oMatrixName = "mx_Bunk"
                                TableName = "@OBT_TB01_BUNKER"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                source = "Bunker"
                            Case "TOLLPURCHASEORDER"
                                oMatrixName = "mx_Toll"
                                TableName = "@OBT_TB01_TOLL"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                source = "Toll"
                            Case "CourierPURCHASEORDER"
                                oMatrixName = "mx_Courier"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                TableName = "@OBT_FCL21_COURIER"
                            Case "DGLPPURCHASEORDER"
                                oMatrixName = "mx_DGLP"
                                TableName = "@OBT_FCL23_DGLP"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                            Case "OutriderPURCHASEORDER"
                                oMatrixName = "mx_Outer"
                                tblHeader = "@OUTRIDER"
                                TableName = "@OBT_TBL03_OUTRIDER"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                source = "Outrider"
                            Case "COOPURCHASEORDER"
                                oMatrixName = "mx_COO"
                                TableName = "@OBT_FCL20_COO"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                            Case "FUMIPURCHASEORDER" 'Fumigation
                                oMatrixName = "mx_Fumi"
                                tblHeader = "@FUMIGATION"
                                TableName = "@OBT_TBL01_FUMIGAT"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                source = "Fumigation"
                            Case "FORKLIFTPURCHASEORDER"  'to combine
                                source = "Forklift"
                                oMatrixName = "mx_Fork"
                                tblHeader = "@FORKLIFT"
                                TableName = "@OBT_TBL05_FORKLIFT"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                            Case "CRATEPURCHASEORDER"  'to combine
                                source = "Crate"
                                oMatrixName = "mx_Crate"
                                tblHeader = "@CRATE"
                                TableName = "@OBT_TBL08_CRATE"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)

                        End Select
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            oMatrix = ExportSeaFCLForm.Items.Item(oMatrixName).Specific
                            Dim sql As String = "select DocEntry,U_PONo,U_VCode,U_VName,U_VRef,U_SInA,U_TDate,U_TTime,U_TPlace,U_CPerson,U_PODate,U_POTime,U_POStatus,U_PORMKS,U_POIRMKS from " & _
                                                               "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If Not CreatePurchaseOrder(oPOForm, "mx_Item") Then
                                BubbleEvent = False
                                Throw New ArgumentException(sErrDesc)
                            End If
                            If Not PopulatePurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, True, oPOForm, source) Then Throw New ArgumentException(sErrDesc)

                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value))

                        End If
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then

                            oMatrix = ExportSeaFCLForm.Items.Item(oMatrixName).Specific
                            If Not UpdatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            Dim sql As String = "select DocEntry,U_PONo,U_VCode,U_VName,U_VRef,U_SInA,U_TDate,U_TTime,U_TPlace,U_CPerson,U_PODate,U_POTime,U_POStatus,U_PORMKS,U_POIRMKS from " & _
                                                               "[@OBT_TB08_FFCPO] where  U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                            If Not PopulatePurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, False, oPOForm, source) Then Throw New ArgumentException(sErrDesc)

                        End If
                    End If

                Case "2000000027", "2000000029", "2000000033", "2000000036", "2000000039", "2000000027", "2000000042", "2000000007", "2000000016", "2000000061", "2000000012", "2000000011"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oGRForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("EXPORTSEAFCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRLCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                        End If

                        Select Case BusinessObjectInfo.FormUID
                            Case "CRANEGOODSRECEIPT"
                                oMatrixName = "mx_Crane"
                                tblHeader = "@CRANE"
                                TableName = "@OBT_TB33_CRANE"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                source = "Crane"
                            Case "BUNKGOODSRECEIPT"
                                oMatrixName = "mx_Bunk"
                                tblHeader = "@BUNKER"
                                TableName = "@OBT_TB01_BUNKER"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                source = "Bunk"
                            Case "TOLLGOODSRECEIPT"
                                oMatrixName = "mx_Toll"
                                tblHeader = "@TOLL"
                                TableName = "@OBT_TB01_TOLL"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                source = "Toll"
                            Case "FORKLIFTGOODSRECEIPT" 'to combine
                                oMatrixName = "mx_Fork"
                                TableName = "@OBT_TBL05_FORKLIFT"
                                tblHeader = "@FORKLIFT"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                source = "Forklift"
                            Case "CRATEGOODSRECEIPT" 'to combine
                                oMatrixName = "mx_Crate"
                                TableName = "@OBT_TBL08_CRATE"
                                tblHeader = "@CRATE"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                source = "Crate"
                            Case "COURIERGOODSRECEIPT"
                                oMatrixName = "mx_Courier"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                TableName = "@OBT_FCL21_COURIER"
                            Case "DGLPGOODSRECEIPT"
                                oMatrixName = "mx_DGLP"
                                TableName = "@OBT_FCL23_DGLP"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                            Case "OUTRIDERGOODSRECEIPT"
                                oMatrixName = "mx_Outer"
                                tblHeader = "@OUTRIDER"
                                TableName = "@OBT_TBL03_OUTRIDER"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                source = "Outrider"
                            Case "COOGOODSRECEIPT"
                                oMatrixName = "mx_COO"
                                TableName = "@OBT_FCL20_COO"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                            Case "FUMIGOODSRECEIPT" 'Fumigation
                                oMatrixName = "mx_Fumi"
                                tblHeader = "@FUMIGATION"
                                TableName = "@OBT_TBL01_FUMIGAT"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                source = "Fumigation"
                        End Select


                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            oMatrix = ExportSeaFCLForm.Items.Item(oMatrixName).Specific
                            If Not CreateGoodsReceiptPO(oGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            Dim sql As String = "select DocEntry,U_PONo,U_PO,U_VCode,U_VName,U_VRef,U_SInA,U_TDate,U_TTime,U_TPlace,U_CPerson,U_PODate,U_POTime,U_POStatus,U_PORMKS,U_POIRMKS from " & _
                                                "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())

                            If Not PopulatePurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, False, oGRForm, source) Then Throw New ArgumentException(sErrDesc)

                            If oMatrix.Columns.Item("colMultiJb").Cells.Item(currentRow).Specific.Value() = "Y" Then
                                modMultiJobForNormal.UpdateMultiJobPOStatusButton(ExportSeaFCLForm, "Closed", oMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value, TableName, tblHeader)
                            End If
                            oMatrix = oActiveForm.Items.Item("mx_PO").Specific
                            If Not EditPOTab(oActiveForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc)

                        End If

                    End If
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                        ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                        ExportSeaFCLForm.Close()
                    End If


                Case "SHIPPINGINV"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            If AlreadyExist("EXPORTSEAFCL") Then
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                            End If

                            oShpForm = p_oSBOApplication.Forms.GetForm("SHIPPINGINV", 1)
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ShpInv").Specific
                            If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_FCL07_SHPINV", True)
                            ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                AddUpdateShippingMatrix(oShpForm, oMatrix, "@OBT_FCL07_SHPINV", False)
                            End If
                        Catch ex As Exception

                        End Try


                    End If
            'Voucher POP UP
                Case "VOUCHER"
                    

                    If BusinessObjectInfo.ActionSuccess = True Then
                        'ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        If AlreadyExist("EXPORTSEAFCL") Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRFCL") Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                        End If
                        oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_FCL05_EVOUCHER", True)
                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_FCL05_EVOUCHER", False)
                        End If

                    End If
                Case "EXPORTSEAFCL", "EXPORTAIRFCL"
                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            If AlreadyExist("EXPORTSEAFCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                            End If
                            ExportSeaFCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            LoadHolidayMarkUp(ExportSeaFCLForm)
                            ExportSeaFCLForm.Items.Item("ed_TkrTime").Specific.Value = String.Empty 'MSW
                            If Not String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_InsDate").Specific.Value) Then
                                If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            'If Not String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_DspDate").Specific.Value) Then
                            '    If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_DspDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            'End If
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

                            End If
                        End If
                    End If
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        ExportSeaFCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If BusinessObjectInfo.ActionSuccess = True Then
                            Dim JobLastDocEntry As Integer
                            Dim ObjectCode As String = String.Empty
                            Dim NewJobNo As String = ""
                            sql = "select top 1 Docentry from [@OBT_FCL01_EXPORT] order by docentry desc"
                            oRecordSet.DoQuery(sql)
                            Dim FrDocEntry As Integer = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString)
                            'Google Doc
                            If ExportSeaFCLForm.Items.Item("cb_TMode").Specific.Value.ToString.Trim = "Air" Then
                                jMode = "A"
                            ElseIf ExportSeaFCLForm.Items.Item("cb_TMode").Specific.Value.ToString.Trim = "Sea" Then
                                jMode = "S"
                            ElseIf ExportSeaFCLForm.Items.Item("cb_TMode").Specific.Value.ToString.Trim = "Land" Then
                                jMode = "L"

                            End If
                            If ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Substring(0, 6) = "Import" Then
                                NewJobNo = GetJobNumber("I" & jMode.ToString)
                            ElseIf ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Substring(0, 6) = "Export" Then
                                NewJobNo = GetJobNumber("E" & jMode.ToString)
                            ElseIf ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Substring(0, 5) = "Local" Then
                                NewJobNo = GetJobNumber("L" & jMode.ToString)
                            ElseIf ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Substring(0, 12) = "Transhipment" Then
                                NewJobNo = GetJobNumber("TS")
                            End If
                            sql = "select top 1 Docentry from [@OBT_FREIGHTDOCNO] order by docentry desc"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.Fields.Item("Docentry").Value.ToString = "" Then
                                JobLastDocEntry = 1
                            Else
                                JobLastDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString) + 1
                            End If
                            sql = "Update [@OBT_FCL01_EXPORT] set U_JbDocNo=" & JobLastDocEntry & ",U_JobNum = '" & NewJobNo & "' Where DocEntry=" & FrDocEntry & ""
                            oRecordSet.DoQuery(sql)
                            p_oSBOApplication.SetStatusBarMessage("Actual Job Number is " & NewJobNo, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            sql = "Insert Into [@OBT_FREIGHTDOCNO] (DocEntry,DocNum,U_JobNo,U_JobType,U_JbStus,U_FrDocNo,U_JbDate,U_ObjType,U_CusCode,U_CusName,U_ShpCode,U_IShpCode,U_ShpName,U_IShpName) Values " & _
                                "(" & JobLastDocEntry & _
                                    "," & JobLastDocEntry & _
                                   "," & IIf(NewJobNo <> "", FormatString(NewJobNo), "NULL") & _
                                    "," & IIf(ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value), "NULL") & _
                                    "," & IIf(ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value), "NULL") & _
                                    "," & FrDocEntry & _
                                     "," & IIf(ExportSeaFCLForm.Items.Item("ed_JbDate").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_JbDate").Specific.Value), "Null") & _
                                    "," & IIf(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString() <> "", FormatString(p_oSBOApplication.Forms.ActiveForm.TypeEx.ToString()), "Null") & _
                                     "," & IIf(ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value), "Null") & _
                                      "," & IIf(ExportSeaFCLForm.Items.Item("ed_Name").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_Name").Specific.Value), "Null") & _
                                       "," & IIf(ExportSeaFCLForm.Items.Item("ed_V").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_V").Specific.Value), "Null") & _
                                       "," & IIf(ExportSeaFCLForm.Items.Item("ed_IV").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_IV").Specific.Value), "Null") & _
                                       "," & IIf(ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.Value), "Null") & _
                                    "," & IIf(ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.Value <> "", FormatString(ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.Value), "Null") & ")"
                            oRecordSet.DoQuery(sql)
                        End If
                    ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                        ExportSeaFCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If BusinessObjectInfo.ActionSuccess = True Then
                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                            oRecordSet.DoQuery(sql)
                        End If
                    End If

                    'MSW New Button at Choose From List 22-03-2011
                Case "VESSEL"
                    Dim vesCode As String = String.Empty
                    Dim voyNo As String = String.Empty
                    Dim oExportSeaFCLForm As SAPbouiCOM.Form = Nothing
                    If AlreadyExist("EXPORTSEALCL") Then
                        oExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                    ElseIf AlreadyExist("EXPORTAIRLCL") Then
                        oExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                    End If
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select top 1* from [@OBT_TB018_VESSEL] order by DocEntry desc")
                            vesCode = oRecordSet.Fields.Item("Name").Value.ToString
                            voyNo = oRecordSet.Fields.Item("U_Voyage").Value.ToString
                        End If
                        oExportSeaFCLForm.Items.Item("ed_Vessel").Specific.Value = vesCode
                        oExportSeaFCLForm.Items.Item("ed_Voy").Specific.Value = voyNo

                    Else
                        If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE) And BusinessObjectInfo.BeforeAction = False Then
                            If BusinessObjectInfo.ActionSuccess = True Then
                                oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                                oExportSeaFCLForm.Items.Item("ed_Vessel").Specific.Value = ""
                                oExportSeaFCLForm.Items.Item("ed_Voy").Specific.Value = ""

                            End If
                        End If
                    End If
                    'MSW New Button at Choose From List 22-03-2011


                    'MSW to Edit New Ticket
                Case "134"
                    Dim vesCode As String = String.Empty
                    Dim voyNo As String = String.Empty
                    Dim oExportSeaFCLForm As SAPbouiCOM.Form = Nothing

                    oExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                    oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select OSLP.SlpName as SlpName,OCRD.CardFName as Country from OCRD left join OSLP on OCRD.SlpCode =OSLP.SlpCode  WHERE OCRD.CardCode = '" & oExportSeaFCLForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                            If oRecordSet.RecordCount > 0 Then
                                oExportSeaFCLForm.Items.Item("ed_Sales").Specific.Value = oRecordSet.Fields.Item("SlpName").Value.ToString
                                oExportSeaFCLForm.Items.Item("ed_Country").Specific.Value = oRecordSet.Fields.Item("Country").Value.ToString
                            End If
                        End If

                    End If
                Case "CHARGES"
                    Dim chDesc As String = String.Empty
                    Dim oCombo As SAPbouiCOM.ComboBox
                    Dim oImportSeaFCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            oCombo = oImportSeaFCLForm.Items.Item("cb_PayFor").Specific
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
                                oCombo = oImportSeaFCLForm.Items.Item("cb_PayFor").Specific
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
            DoExportSeaFCLFormDataEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaFCLFormDataEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()
        End Try
    End Function

    Private Function PopulatePurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal TableName As String, ByVal ProcessedState As Boolean, ByRef oPOForm As SAPbouiCOM.Form) As Boolean
        PopulatePurchaseHeader = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Try
            ObjDbDataSource = pForm.DataSources.DBDataSources.Item(TableName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If pMatrix.RowCount > 0 Then
                        If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNo").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                            pMatrix.Clear()
                        End If
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
                            '.SetValue("LineId", .Offset, pMatrix.VisualRowCount + 1)
                            .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                            .SetValue("U_PONo", .Offset, DocLastKey)
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
                            .SetValue("U_Remark", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                            ' jobNo = pForm.Items.Item("ed_JobNo").Specific.value
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
                            If oPOForm.Title.Contains("Goods Receipt") = True Then 'MSW To Edit New Ticket
                                .SetValue("U_FDate", .Offset, oPOForm.Items.Item("ed_GRDate").Specific.Value)
                                .SetValue("U_FTime", .Offset, oPOForm.Items.Item("ed_GRTime").Specific.Value)
                            Else
                                .SetValue("U_FDate", .Offset, oPOForm.Items.Item("ed_PODate").Specific.Value)
                                .SetValue("U_FTime", .Offset, oPOForm.Items.Item("ed_POTime").Specific.Value)
                            End If
                            .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                            .SetValue("U_Remark", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
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

    Private Function PopulateOtherPurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal TableName As String) As Boolean
        PopulateOtherPurchaseHeader = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Try
            ObjDbDataSource = pForm.DataSources.DBDataSources.Item(TableName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If pMatrix.RowCount > 0 Then
                        If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNo").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                            pMatrix.Clear()
                        End If
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
                        .SetValue("U_FTime", .Offset, oRecordSet.Fields.Item("U_PODate").Value)

                        .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                        .SetValue("U_Remark", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
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


    Public Function DoExportSeaFCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
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
        Dim ExportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oChMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim oComboPO As SAPbouiCOM.ComboBox = Nothing
        Dim CPOMatrix As SAPbouiCOM.Matrix = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim CGRMatrix As SAPbouiCOM.Matrix = Nothing

        Dim CraneForm As SAPbouiCOM.Form = Nothing
        Dim ForkForm As SAPbouiCOM.Form = Nothing
        Dim DGLPForm As SAPbouiCOM.Form = Nothing
        Dim OutForm As SAPbouiCOM.Form = Nothing
        Dim CourForm As SAPbouiCOM.Form = Nothing
        Dim COOForm As SAPbouiCOM.Form = Nothing
        Dim MultiJobForm As SAPbouiCOM.Form = Nothing
        Dim FumigationForm As SAPbouiCOM.Form = Nothing 'Fumigation
        Dim DetachJobForm As SAPbouiCOM.Form = Nothing
        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoExportSeaFCLItemEvent()"
        Dim sql As String = ""
        Dim formUid As String = String.Empty
        Dim strDsp As String = String.Empty


        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            If AlreadyExist("EXPORTSEAFCL") Then
                formUid = "EXPORTSEAFCL"
            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                formUid = "EXPORTAIRFCL"
            ElseIf AlreadyExist("EXPORTLAND") Then
                formUid = "EXPORTLAND"
            End If
            Select Case pVal.FormTypeEx

                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000044", _
                    "2000000027", "2000000029", "2000000033", "2000000036", "2000000039", "2000000042", "2000000051", "2000000053", "2000000061" 'MSW 14-09-2011 Truck PO  ' CGR --> Custom Goods Receipt
                    CGRForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        CGRForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "mx_Item" And (pVal.ColUID = "LineId" Or pVal.ColUID = "V_-1") Then
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
                                    If CGRMatrix.Columns.Item("colItemNo").Cells.Item(pVal.Row).Specific.Value <> "" Then
                                        If Not AddNewRowGR(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
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
                                'CGRForm.Items.Item("ed_TDate").Specific.Active = True
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
                            'If pVal.ItemUID = "ed_TDate" And CGRForm.Items.Item("ed_TDate").Specific.String <> String.Empty Then
                            '    '    If Not DateTime(CGRForm, _
                            '    '                    CGRForm.Items.Item("ed_TDate").Specific, _
                            '    '                    CGRForm.Items.Item("ed_TDay").Specific, _
                            '    '                    CGRForm.Items.Item("ed_TTime").Specific) Then Throw New ArgumentException(sErrDesc)
                            '    '    CGRForm.Items.Item("ed_TPlace").Specific.Active = True
                            'End If

                            If pVal.ItemUID = "1" Then
                                If pVal.BeforeAction = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                    End If
                                End If
                            End If

                            If pVal.ItemUID = "1" Then

                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                ' When click the Add button in AddMode of Custom Purchase Order form
                                ' need to also trigger the item pressed event of Main Export Form according by Customize Biz Logic
                                If pVal.Action_Success = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        CGRForm = p_oSBOApplication.Forms.ActiveForm
                                        If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            'when retrieve spefific data for update, add new row in the matrix
                                            ' If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            CGRForm.Close()

                                            Try
                                                If pVal.FormTypeEx = "2000000027" Then 'Crane
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                                ElseIf pVal.FormTypeEx = "2000000016" Then 'Bunker
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                                ElseIf pVal.FormTypeEx = "2000000061" Then 'Toll
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                                ElseIf pVal.FormTypeEx = "2000000029" Then 'Forklift
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                                ElseIf pVal.FormTypeEx = "2000000033" Then
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                                ElseIf pVal.FormTypeEx = "2000000036" Then
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                                ElseIf pVal.FormTypeEx = "2000000039" Then
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                                ElseIf pVal.FormTypeEx = "2000000042" Then
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                                                ElseIf pVal.FormTypeEx = "2000000007" Then
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1) 'Fumigation
                                                ElseIf pVal.FormTypeEx = "2000000012" Then 'to combine
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                                ElseIf pVal.FormTypeEx = "2000000011" Then 'to combine
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                                Else
                                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                                End If
                                                If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                                End If
                                            Catch ex As Exception

                                            End Try

                                            ExportSeaFCLForm.Items.Item("1").Click()
                                            If MainForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                MainForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            MainForm.Items.Item("1").Click() 'Fumigation 
                                            MainForm.Items.Item("2").Specific.Caption = "Close"
                                        End If
                                    End If
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DrBL").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_SpOrder").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_CrPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_CPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_BunkPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_ArmePO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    'ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                                    ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = True

                                    ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_A6Label").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                                    'ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                    ' ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                                End If
                            End If
                        End If

                    End If
                Case "2000000026", "2000000030", "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000029", "2000000032",
                    "2000000035", "2000000038", "2000000041", "2000000043", "2000000050", "2000000052", "2000000060"  ''MSW To Edit New Ticket   ' CPO --> Custom Purchase Order"
                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE Then
                        Try

                        Catch ex As Exception

                        End Try
                    End If

                    If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        CPOForm.Items.Item("ed_Code").Enabled = False
                        CPOForm.Items.Item("ed_Name").Enabled = False
                    End If

                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            If pVal.ItemUID = "mx_Item" And (pVal.ColUID = "LineId" Or pVal.ColUID = "V_-1") Then
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

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            'Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Try
                                If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                                    CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)
                                    CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_VName", 0, oDataTable.GetValue(1, 0).ToString)
                                    oComboPO = CPOForm.Items.Item("cb_Contact").Specific
                                    If Not ClearComboData(CPOForm, "cb_Contact", "@OBT_TB08_FFCPO", "U_CPerson") Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oComboPO.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                            oRecordSet.MoveNext()
                                        End While
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
                                'End If

                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            Try
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
                                    CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_SInA", 0, comboStaff.Selected.Description.ToString)
                                    If pVal.FormUID = "DGLPPURCHASEORDER" Then
                                        CPOForm.Items.Item("ed_TDate").Specific.Active = True
                                    End If
                                End If

                            Catch ex As Exception
                                MessageBox.Show(ex.ToString())
                            End Try

                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "ed_TDate" Then
                                If CPOForm.Items.Item("ed_TDate").Specific.String <> String.Empty Then
                                    If Not DateTime(CPOForm, _
                                                CPOForm.Items.Item("ed_TDate").Specific, _
                                                CPOForm.Items.Item("ed_TDay").Specific, _
                                                CPOForm.Items.Item("ed_TTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                End If
                                'CPOForm.Items.Item("ed_TPlace").Specific.Active = True
                            End If

                            If pVal.ItemUID = "1" Then

                                If pVal.Action_Success = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        CPOForm = p_oSBOApplication.Forms.ActiveForm

                                        Try
                                            If pVal.FormTypeEx = "2000000026" Then 'Crane
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                            ElseIf pVal.FormTypeEx = "2000000015" Then 'BUNKER
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                            ElseIf pVal.FormTypeEx = "2000000060" Then 'TOLL
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                            ElseIf pVal.FormTypeEx = "2000000030" Then 'Forklift
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                            ElseIf pVal.FormTypeEx = "2000000035" Then 'Forklift
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                            ElseIf pVal.FormTypeEx = "2000000038" Then 'Outrider
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                            ElseIf pVal.FormTypeEx = "2000000041" Then 'Forklift
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                                            ElseIf pVal.FormTypeEx = "2000000032" Then 'Forklift
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                            ElseIf pVal.FormTypeEx = "2000000005" Then 'Fumigation
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                            ElseIf pVal.FormTypeEx = "2000000021" Then 'to combine
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                            ElseIf pVal.FormTypeEx = "2000000020" Then 'to combine
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                            Else
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                            End If
                                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                        Catch ex As Exception

                                        End Try
                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            'when retrieve spefific data for update, add new row in the matrix
                                            If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                            If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                ExportSeaFCLForm.Items.Item("1").Click()
                                                CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            End If
                                            CPOForm.Items.Item("ed_Code").Enabled = False
                                            CPOForm.Items.Item("ed_Name").Enabled = False
                                            CPOForm.Items.Item("cb_SInA").Enabled = True
                                            CPOForm.Items.Item("bt_Preview").Visible = True

                                            CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If



                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            CPOForm.Close()


                                            ExportSeaFCLForm.Items.Item("1").Click()
                                            'MSW 14-09-2011 Truck PO
                                            If pVal.FormTypeEx = "2000000050" Then
                                                ' EnabledTruckerForExternal(ExportSeaFCLForm, False)
                                                ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = False
                                            End If
                                            'MSW 14-09-2011 Truck PO
                                            If pVal.FormTypeEx = "2000000052" Then
                                                ' EnabledDispatchForExternal(ExportSeaFCLForm, False)
                                                ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False
                                            End If
                                        End If
                                    End If
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DrBL").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_SpOrder").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_CrPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_CPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_BunkPO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_ArmePO").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    'ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                                    ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_A6Label").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                                    'ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                    'ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If
                            If pVal.ItemUID = "bt_Preview" Then
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                PreviewPO(ExportSeaFCLForm, CPOForm)
                            End If
                        End If
                    End If

                    If pVal.BeforeAction = True Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "1" Then
                                Try
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                                    If Not CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        CPOMatrix = CPOForm.Items.Item("mx_Item").Specific
                                        'MSW To Edit New Ticket
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
                    Dim m3 As Decimal
                    Try
                        oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "1" Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    oShpForm.Items.Item("bt_PPView").Visible = True
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Then
                                        If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            sql = "Update [@OBT_TB03_EXPSHPINV] set " & _
                                            " U_FrDocNo=" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oShpForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            oShpForm.Close()

                                            ExportSeaFCLForm.Items.Item("1").Click()
                                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

                                            End If

                                        ElseIf oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Items.Item("ed_ShipTo").Enabled = False
                                            oShpForm.Items.Item("bt_PPView").Visible = True
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                    ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
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
                                Else
                                    If oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD" Then
                                        AddUpdateShippingInv(oShpForm, oShpMatrix, "@OBT_TB03_EXPSHPINVD", True)
                                        oShpForm.Items.Item("bt_Add").Specific.Caption = "Edit"
                                        CalculateNoOfBoxes(oShpForm, oShpMatrix)
                                    Else
                                        AddUpdateShippingInv(oShpForm, oShpMatrix, "@OBT_TB03_EXPSHPINVD", False)
                                        CalculateNoOfBoxes(oShpForm, oShpMatrix) 'oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD" 'MSW To Edit
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
                                    'oShpForm.DataSources.UserDataSources.Item("Qty").Value = IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString)'MSW To Edit 31-08-2011
                                    oShpForm.DataSources.UserDataSources.Item("Unit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString)
                                    '   oActiveForm.DataSources.UserDataSources.Item("Box").Value = oDataTable.Columns.Item("U_Box").Cells.Item(0).Value.ToString
                                    oShpForm.DataSources.UserDataSources.Item("DL").Value = IIf(oDataTable.Columns.Item("U_Length").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Length").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("DB").Value = IIf(oDataTable.Columns.Item("U_Base").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Base").Cells.Item(0).Value.ToString)
                                    oShpForm.DataSources.UserDataSources.Item("DH").Value = IIf(oDataTable.Columns.Item("U_Height").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Height").Cells.Item(0).Value.ToString)

                                    m3 = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Length").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Length").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Base").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Base").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Height").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Height").Cells.Item(0).Value)) / 1000000
                                    oShpForm.DataSources.UserDataSources.Item("M3").Value = m3
                                    'oShpForm.DataSources.UserDataSources.Item("TM3").Value = m3 * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString))
                                    oShpForm.DataSources.UserDataSources.Item("NetKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value))
                                    'oShpForm.DataSources.UserDataSources.Item("TotKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NetWt").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("GroKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value.ToString))
                                    'oShpForm.DataSources.UserDataSources.Item("TotGKg").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_GrWt").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    'MSW To Edit

                                    oShpForm.DataSources.UserDataSources.Item("NEC").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value.ToString)) 'MSW To Edit
                                    ' oShpForm.DataSources.UserDataSources.Item("TotNEC").Value = Convert.ToDouble(IIf(oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_NEQ").Cells.Item(0).Value)) * Convert.ToDouble(IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value))
                                    oShpForm.DataSources.UserDataSources.Item("PBox").Value = IIf(oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value = vbNull, 0, oDataTable.Columns.Item("U_Qty").Cells.Item(0).Value.ToString) 'MSW To Edit 31-08-2011
                                    'oShpForm.DataSources.UserDataSources.Item("BUnit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString)
                                    'oShpForm.DataSources.UserDataSources.Item("PUnit").Value = IIf(oDataTable.Columns.Item("U_UM").Cells.Item(0).Value = vbNull.ToString, " ", oDataTable.Columns.Item("U_UM").Cells.Item(0).Value.ToString)
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
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                            If pVal.ItemUID = "ed_Qty" Then
                                'MSW To Edit
                                If oShpForm.Items.Item("ed_Qty").Specific.value.ToString <> "" Then
                                    Dim qty As Double
                                    Dim noofPkg As Double
                                    Try
                                        qty = IIf(Convert.ToDouble(oShpForm.Items.Item("ed_Qty").Specific.value.ToString) = 0, 0, Convert.ToDouble(oShpForm.Items.Item("ed_Qty").Specific.value.ToString))
                                        If oShpForm.Items.Item("ed_PPBNo").Specific.value <> "0" And oShpForm.Items.Item("ed_PPBNo").Specific.value <> "" Then
                                            noofPkg = qty / Convert.ToDouble(oShpForm.Items.Item("ed_PPBNo").Specific.value.ToString)
                                        End If

                                        m3 = IIf(oShpForm.Items.Item("ed_L").Specific.value = "", 0, oShpForm.Items.Item("ed_L").Specific.value) * IIf(oShpForm.Items.Item("ed_B").Specific.value = "", 0, oShpForm.Items.Item("ed_B").Specific.value) * IIf(oShpForm.Items.Item("ed_H").Specific.value = "", 0, oShpForm.Items.Item("ed_H").Specific.value)
                                        If (m3 <> 0) Then
                                            m3 = IIf(oShpForm.Items.Item("ed_L").Specific.value = vbNull, 0, oShpForm.Items.Item("ed_L").Specific.value) * IIf(oShpForm.Items.Item("ed_B").Specific.value = vbNull, 0, oShpForm.Items.Item("ed_B").Specific.value) * IIf(oShpForm.Items.Item("ed_H").Specific.value = vbNull, 0, oShpForm.Items.Item("ed_H").Specific.value) / 1000000
                                        End If
                                        oShpForm.DataSources.UserDataSources.Item("M3").Value = m3
                                        oShpForm.DataSources.UserDataSources.Item("TM3").Value = m3 * noofPkg
                                        oShpForm.DataSources.UserDataSources.Item("TotBox").Value = noofPkg
                                        'MSW(21 - 9 - 2011)
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


                    End If
                Case "VOUCHER"
                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                    oPayForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    Try
                        oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(oPayForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If

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
                 
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                        If pVal.ItemUID = "1" Then
                            Try
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
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
                                'oPayForm.Items.Item("ed_PayTo").Enabled = False
                                oPayForm.Items.Item("cb_PayCur").Enabled = False 'to km
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
                                PreviewPaymentVoucher(ExportSeaFCLForm, oPayForm)
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
                                        'to km
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
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
                                        'to km
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
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)

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
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                                            oPayForm.Items.Item("bt_PayView").Visible = True
                                            oPayForm.Items.Item("bt_PayView").Visible = True
                                            'oPayForm.Items.Item("ed_PayTo").Enabled = False

                                        End If
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB031_VHEADER] set U_APInvNo=" & Convert.ToInt32(strAPInvNo) & ",U_OutPayNo=" & Convert.ToInt32(strOutPayNo) & "" & _
                                            " ,U_FrDocNo=" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oPayForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            PreviewPaymentVoucher(ExportSeaFCLForm, oPayForm)
                                            oPayForm.Close()

                                            ExportSeaFCLForm.Items.Item("1").Click()
                                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If

                                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            ' p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            oPayForm.Items.Item("bt_PayView").Visible = True
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                    ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                End If
                            End If
                        End If
                        'to km
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                            If pVal.ItemUID = "ed_VedCode" Then
                                If oPayForm.Items.Item("ed_VedCode").Specific.Value.ToString <> "" Then
                                    oRecordSet.DoQuery("Select Currency from OCRD Where CardCode='" & oPayForm.Items.Item("ed_VedCode").Specific.Value & "' ")
                                    If oRecordSet.RecordCount > 0 Then
                                        If oRecordSet.Fields.Item("Currency").Value.ToString <> "##" Then
                                            oCombo = oPayForm.Items.Item("cb_PayCur").Specific
                                            oCombo.Select(oRecordSet.Fields.Item("Currency").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        'to km
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects

                            If pVal.ItemUID = "ed_VedCode" Then
                                ObjDBDataSource = oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER") 'MSW To Add 18-3-2011
                                oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                'oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                '                                                                       & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString()) 'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*

                                vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft
                                oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_BPCode", ObjDBDataSource.Offset, oDataTable.GetValue(0, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                'New UI
                                If oDataTable.Columns.Item("Currency").Cells.Item(0).Value.ToString() <> "##" Then
                                    oPayForm.Items.Item("cb_PayCur").Enabled = False
                                Else
                                    oPayForm.Items.Item("cb_PayCur").Enabled = True
                                End If
                                'New UI
                            End If

                            If pVal.ColUID = "colChCode1" Then
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                'MSW to Edit New Ticket
                                Try
                                    'oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString

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
                                    '  oChMatrix.Columns.Item("colICode").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_ItemCode").Cells.Item(0).Value.ToString
                                    oChMatrix.Columns.Item("colICode").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
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
                                    oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_ExRate", 0, Rate.ToString)
                                Else
                                    oPayForm.DataSources.DBDataSources.Item("@OBT_TB031_VHEADER").SetValue("U_ExRate", 0, Nothing)
                                    oPayForm.Items.Item("ed_PayRate").Visible = False
                                End If
                            End If
                        End If
                    End If
                Case "Image"
                    If pVal.BeforeAction = True Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                            sPicName = pVal.ItemUID
                            LoadPopUpForm(pVal.ItemUID)
                        End If
                    End If
                Case "ShowBigImg"
                    If pVal.BeforeAction = True Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If (pVal.ItemUID = "btn_Close") Then
                                ActiveForm.Close()
                            End If
                            If (pVal.ItemUID = "btn_Print") Then
                                Try
                                    PrintDoc(p_fmsSetting.PicturePath & sPicName.Substring(0, 1) & "\" & sPicName & ".jpg")
                                    'ShowPrintDialogBrowser("D:\FMS\" & sPicName.Substring(0, 1) & "\" & sPicName & ".jpg")
                                    sImgePath = ""
                                Catch ex As Exception

                                End Try

                            End If

                        End If
                    End If


                Case "EXPORTSEAFCL", "EXPORTAIRFCL", "EXPORTLAND"
                    ExportSeaFCLForm = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
                    Try
                        ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(ExportSeaFCLForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If
                    If Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                            If pVal.ItemUID = "1" Then
                                oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                If oMatrix.RowCount > 1 Then
                                    If oMatrix.Columns.Item("colCNo1").Cells.Item(oMatrix.RowCount).Specific.Value = "" Then
                                        oMatrix.DeleteRow(oMatrix.RowCount)
                                    End If
                                End If
                                oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                If oMatrix.RowCount > 1 Then
                                    If oMatrix.Columns.Item("colCCode1").Cells.Item(oMatrix.RowCount).Specific.Value = "" Then
                                        oMatrix.DeleteRow(oMatrix.RowCount)
                                    End If
                                End If

                            End If
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "mx_ConTab" Then
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                            If pVal.ColUID = "colCNo1" Then
                                If oMatrix.Columns.Item("colCNo1").Cells.Item(pVal.Row).Specific.Value <> "" And oMatrix.Columns.Item("colCNo1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL19_HCONTAINE").Clear()
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                    oCombo = oMatrix.Columns.Item("colCSize1").Cells.Item(oMatrix.RowCount).Specific
                                    oCombo.Select("20'", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If
                            End If
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "mx_ConTab" Then
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                            If pVal.ColUID = "colCSize1" Then
                                UpdateNoofContainer(ExportSeaFCLForm, oMatrix)
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "mx_Voucher" Or pVal.ItemUID = "mx_TkrList" Or pVal.ItemUID = "mx_DspList" Or pVal.ItemUID = "mx_PO" Then
                            oMatrix = ExportSeaFCLForm.Items.Item(pVal.ItemUID).Specific
                            If pVal.Row > 0 Then
                                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                oMatrix.SelectRow(pVal.Row, True, False)
                                selectedRow = pVal.Row
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK And pVal.BeforeAction = False Then
                        Dim Code As String = ""
                        Dim Name As String = ""

                        If pVal.ItemUID = "mx_TkrList" Then
                            If pVal.Row > 0 Then
                                ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Update"

                                oMatrix = ExportSeaFCLForm.Items.Item("mx_TkrList").Specific
                                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                oMatrix.SelectRow(pVal.Row, True, False)
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value() = "Closed" Or oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value() = "Cancelled" Then
                                    p_oSBOApplication.SetStatusBarMessage("This PO is already " & oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value & ".", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Else
                                    modTrucking.GetDataFromMatrixByIndex(ExportSeaFCLForm, oMatrix, pVal.Row)
                                    modTrucking.SetDataToEditTabByIndex(ExportSeaFCLForm)

                                    Code = ExportSeaFCLForm.Items.Item("ed_TkrCode").Specific.Value()
                                    Name = ExportSeaFCLForm.Items.Item("ed_Trucker").Specific.Value()

                                    'MSW 14-09-2011 Truck PO
                                    If ExportSeaFCLForm.Items.Item("op_Exter").Specific.Selected = True Then
                                        If ExportSeaFCLForm.Items.Item("ed_TOrigin").Specific.Value = "Y" Then
                                            ExportSeaFCLForm.Items.Item("bt_TkrDPO").Enabled = True
                                        Else
                                            ExportSeaFCLForm.Items.Item("bt_TkrDPO").Enabled = False
                                        End If
                                        ExportSeaFCLForm.Items.Item("op_Exter").Enabled = False
                                        ExportSeaFCLForm.Items.Item("op_Inter").Enabled = False
                                        ExportSeaFCLForm.Items.Item("bt_PO").Enabled = False
                                        If AddChooseFromListByOption(ExportSeaFCLForm, False, "ed_Trucker", "TKRINTR", "CFLTKRE", "TKREXTR", "CFLTKRV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRCODE").ValueEx = Code
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("TKREXTR").ValueEx = Name
                                        ExportSeaFCLForm.Items.Item("ed_Trucker").Enabled = False
                                    ElseIf ExportSeaFCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                        ExportSeaFCLForm.Items.Item("op_Exter").Enabled = True
                                        ExportSeaFCLForm.Items.Item("op_Inter").Enabled = True
                                        If AddChooseFromListByOption(ExportSeaFCLForm, True, "ed_Trucker", "TKRINTR", "CFLTKRE", "TKREXTR", "CFLTKRV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = Name
                                    End If
                                    'MSW 14-09-2011 Truck PO

                                    ExportSeaFCLForm.Items.Item("bt_Tkview").Enabled = True 'MSW to edit New Ticket 07-09-2011
                                    ExportSeaFCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                End If

                               
                            End If
                        ElseIf pVal.ItemUID = "mx_DspList" Then

                            If pVal.Row > 0 Then
                                ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Update"
                                oMatrix = ExportSeaFCLForm.Items.Item("mx_DspList").Specific
                                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                oMatrix.SelectRow(pVal.Row, True, False)
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value() = "Closed" Or oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value() = "Cancelled" Then
                                    p_oSBOApplication.SetStatusBarMessage("This PO is already " & oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value & ".", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Else
                                    GetDispatchDataFromMatrixByIndex(ExportSeaFCLForm, oMatrix, pVal.Row)
                                    SetDispatchDataToEditTabByIndex(ExportSeaFCLForm)
                                    Code = ExportSeaFCLForm.Items.Item("ed_DspCode").Specific.Value()
                                    Name = ExportSeaFCLForm.Items.Item("ed_Dspatch").Specific.Value()
                                    If ExportSeaFCLForm.Items.Item("op_DExter").Specific.Selected = True Then
                                        If ExportSeaFCLForm.Items.Item("ed_DOrigin").Specific.Value = "Y" Then
                                            ExportSeaFCLForm.Items.Item("bt_DspDPO").Enabled = True
                                        Else
                                            ExportSeaFCLForm.Items.Item("bt_DspDPO").Enabled = False
                                        End If
                                        ExportSeaFCLForm.Items.Item("op_DExter").Enabled = False '2-12
                                        ExportSeaFCLForm.Items.Item("op_DInter").Enabled = False
                                        ExportSeaFCLForm.Items.Item("bt_DPO").Enabled = False
                                        If AddChooseFromListByOption(ExportSeaFCLForm, False, "ed_Dspatch", "DSPINTR", "CFLDSP", "DSPEXTR", "CFLDSPV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPCODE").ValueEx = Code
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPEXTR").ValueEx = Name
                                        ExportSeaFCLForm.Items.Item("ed_Dspatch").Enabled = False
                                    ElseIf ExportSeaFCLForm.Items.Item("op_DInter").Specific.Selected = True Then
                                        ExportSeaFCLForm.Items.Item("op_DExter").Enabled = True '2-12
                                        ExportSeaFCLForm.Items.Item("op_DInter").Enabled = True
                                        If AddChooseFromListByOption(ExportSeaFCLForm, True, "ed_Dspatch", "DSPINTR", "CFLDSP", "DSPEXTR", "CFLDSPV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPINTR").ValueEx = Name
                                    End If
                                    ExportSeaFCLForm.Items.Item("bt_Dspview").Enabled = True
                                    ExportSeaFCLForm.Items.Item("fo_DspEdit").Specific.Select()

                                End If
                            End If

                            ElseIf pVal.ItemUID = "mx_Voucher" Then
                                If pVal.Row > 0 Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                    oMatrix.SelectRow(pVal.Row, True, False)
                                    LoadPaymentVoucher(ExportSeaFCLForm)
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                                    oPayForm = p_oSBOApplication.Forms.ActiveForm
                                    oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    oPayForm.Items.Item("ed_DocNum").Visible = True
                                    oPayForm.Items.Item("ed_DocNum").Enabled = True
                                    oPayForm.Items.Item("ed_DocNum").Specific.Value = oMatrix.Columns.Item("colVDocNum").Cells.Item(pVal.Row).Specific.Value.ToString
                                    oPayForm.Items.Item("cb_PayCur").Specific.Active = True
                                    oPayForm.Items.Item("ed_DocNum").Visible = False
                                    oPayForm.Items.Item("ed_DocNum").Enabled = False
                                    oPayForm.DataBrowser.BrowseBy = "ed_DocNum"
                                    oPayForm.Items.Item("ed_VedName").Enabled = False
                                    oPayForm.Items.Item("ed_VedCode").Enabled = False
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
                            End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_GOT_FOCUS And pVal.Before_Action = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                            Dim JobType As String = ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value
                            'Google Doc
                            If JobType.Contains("Import") Then

                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, False, True)
                            ElseIf JobType.Contains("Export") Then

                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, True, False)
                            ElseIf JobType.Contains("Local") Then

                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, True, True)
                            ElseIf JobType.Contains("Transhipment") Then

                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, True, True)
                            End If
                        End If
                    End If


                    'MSW to edit New Ticket 07-09-2011
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                        If pVal.ItemUID = "ed_TotalM3" Or pVal.ItemUID = "ed_TotalWt" Then
                            If Convert.ToDouble(ExportSeaFCLForm.Items.Item("ed_TotalM3").Specific.Value) > Convert.ToDouble(ExportSeaFCLForm.Items.Item("ed_TotalWt").Specific.Value) Then
                                ExportSeaFCLForm.Items.Item("ed_TotChWt").Specific.Value = ExportSeaFCLForm.Items.Item("ed_TotalM3").Specific.Value
                            Else
                                ExportSeaFCLForm.Items.Item("ed_TotChWt").Specific.Value = ExportSeaFCLForm.Items.Item("ed_TotalWt").Specific.Value
                            End If
                            ' ExportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Active = True
                        End If
                    End If
                    'End MSW to edit New Ticket 07-09-2011


                    If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_DrBL").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_SpOrder").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_CrPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_CPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_BunkPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_ArmePO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                    End If


                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                        If pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize" And pVal.Before_Action = True And pVal.Row <> 0 Then
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_Cont").Specific
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
                            If ExportSeaFCLForm.Items.Item("ed_FCharge").Specific.value <> "" Then
                                Try
                                    If ExportSeaFCLForm.Items.Item("ed_FCharge").Specific.value < 0 Then
                                        ExportSeaFCLForm.Items.Item("ed_FCharge").Specific.Value = ""
                                        ExportSeaFCLForm.Items.Item("ed_FCharge").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ExportSeaFCLForm.Items.Item("ed_FCharge").Specific.Value = ""
                                    ExportSeaFCLForm.Items.Item("ed_FCharge").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End Try
                            End If

                        End If
                        If pVal.ItemUID = "ed_Percent" And pVal.Before_Action = True Then
                            If ExportSeaFCLForm.Items.Item("ed_Percent").Specific.value <> "" Then
                                Try
                                    If ExportSeaFCLForm.Items.Item("ed_Percent").Specific.value < 0 Then
                                        ExportSeaFCLForm.Items.Item("ed_Percent").Specific.Value = ""
                                        ExportSeaFCLForm.Items.Item("ed_Percent").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ExportSeaFCLForm.Items.Item("ed_Percent").Specific.Value = ""
                                    ExportSeaFCLForm.Items.Item("ed_Percent").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End Try
                            End If

                        End If
                        If pVal.ItemUID = "ed_GWt" And pVal.Before_Action = True Then
                            If ExportSeaFCLForm.Items.Item("ed_GWt").Specific.value <> "" Then
                                Try
                                    If ExportSeaFCLForm.Items.Item("ed_GWt").Specific.value < 0 Then
                                        ExportSeaFCLForm.Items.Item("ed_GWt").Specific.Value = ""
                                        ExportSeaFCLForm.Items.Item("ed_GWt").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ExportSeaFCLForm.Items.Item("ed_GWt").Specific.Value = ""
                                    ExportSeaFCLForm.Items.Item("ed_GWt").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False

                                End Try
                            End If

                        End If
                    End If

                    '-------------------------For Payment(omm)------------------------------------------'
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.Before_Action = False Then

                        If (pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colGST") Then
                            CalRate(ExportSeaFCLForm, pVal.Row)
                        End If
                        If (pVal.ItemUID = "cb_GST" And ExportSeaFCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft") Then
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ChCode").Specific
                            dtmatrix = ExportSeaFCLForm.DataSources.DataTables.Item("DTCharges")
                            dtmatrix.SetValue("GST", 0, "None")
                            oMatrix.LoadFromDataSource()
                        End If

                        If pVal.ItemUID = "cb_PayCur" Then
                            If ExportSeaFCLForm.Items.Item("cb_PayCur").Specific.Value <> "SGD" Then
                                Dim Rate As String = String.Empty
                                sql = "SELECT Rate FROM ORTT WHERE Currency = '" & ExportSeaFCLForm.Items.Item("cb_PayCur").Specific.Value & "' And DATENAME(YYYY,RateDate) = '" & _
                                        Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                        CInt(Today.ToString("dd"))
                                oRecordSet.DoQuery(sql)
                                If oRecordSet.RecordCount > 0 Then
                                    Rate = oRecordSet.Fields.Item("Rate").Value
                                End If
                                ExportSeaFCLForm.Items.Item("ed_PayRate").Enabled = True
                                ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL05_EVOUCHER").SetValue("U_ExRate", 0, Rate.ToString) ' * Change 
                            Else
                                ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL05_EVOUCHER").SetValue("U_ExRate", 0, Nothing)       '* change
                                ExportSeaFCLForm.Items.Item("ed_PayRate").Enabled = False
                                ExportSeaFCLForm.Items.Item("cb_PayCur").Specific.Active = True
                            End If

                        End If
                    End If
                    '-----------------------------------------------------------------------------------'

                    If pVal.BeforeAction = False Then
                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                            Try
                                ExportSeaFCLForm.Items.Item("ed_Code").Enabled = False
                                ExportSeaFCLForm.Items.Item("ed_Name").Enabled = False
                                ExportSeaFCLForm.Items.Item("ed_JobNo").Enabled = False
                                'Check PO Status
                                If ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Closed" Or ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Cancelled" Then
                                    ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                                    ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = False
                                    ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = False
                                ElseIf ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Open" Then
                                    ExportSeaFCLForm.Items.Item("ch_POD").Enabled = True
                                    ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = True
                                    ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                                End If
                               

                                'If ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value <> "" Then
                                '    If JobStus = "" Then
                                '        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '        oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_FCL01_EXPORT] WHERE DocEntry = " & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value) 'MSW 08-06-2011 for job Type Table
                                '        If oRecordSet.RecordCount > 0 Then
                                '            JobStus = oRecordSet.Fields.Item("U_JbStus").Value
                                '        End If
                                '        If JobStus = "Closed" Or JobStus = "Cancelled" Then
                                '            ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = False
                                '            ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = False
                                '        End If
                                '    End If
                                'End If
                            Catch ex As Exception
                            End Try
                        End If
                       
                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            EnabledTrucker(ExportSeaFCLForm, False)
                        End If

                        '-------------------------For Payment(omm)------------------------------------------'
                        If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "colAmount" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS Then
                            CalRate(ExportSeaFCLForm, pVal.Row)
                        End If
                        '----------------------------------------------------------------------------------'

                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "fo_Prmt"
                                        ExportSeaFCLForm.PaneLevel = 7
                                        ExportSeaFCLForm.Items.Item("fo_PMain").Specific.Select()

                                    Case "fo_Dsptch"
                                        ExportSeaFCLForm.PaneLevel = 41
                                        ExportSeaFCLForm.Items.Item("fo_DspView").Specific.Select()
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_DspList").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ' ExportSeaFCLForm.Items.Item("bt_AmdIns").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                                            Else
                                                ' ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                                            End If
                                        ElseIf oMatrix.RowCount = 0 Then
                                            ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False
                                        End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_DspList"
                                    Case "fo_DspView"
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_DspList").Specific
                                        If ExportSeaFCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If (oMatrix.Columns.Item("V_1").Cells.Item(1).Specific.Value = "") Then
                                                ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False
                                            Else
                                                ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = True
                                            End If
                                        End If
                                        ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Add"
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_DspList"
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.PaneLevel = 41

                                    Case "fo_DspEdit"
                                        ExportSeaFCLForm.PaneLevel = 42
                                        If ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Add" Then
                                            NewPOEditTab(ExportSeaFCLForm, "mx_DspList")
                                        End If

                                    Case "fo_DspSet"
                                        ExportSeaFCLForm.PaneLevel = 43
                                        ExportSeaFCLForm.Items.Item("fo_DspSet").Specific.Select()
                                    Case "fo_Trkng"
                                        ExportSeaFCLForm.PaneLevel = 38
                                        ExportSeaFCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_TkrList").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                            Else
                                                ' ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                            End If
                                        ElseIf oMatrix.RowCount = 0 Then
                                            '  ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = False

                                        End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_TkrList"
                                        '#1008 17-09-2011
                                    Case "fo_Vchr"
                                        ExportSeaFCLForm.PaneLevel = 20
                                        ExportSeaFCLForm.Items.Item("fo_VoView").Specific.Select()
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                            Else
                                                '    ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        ElseIf oMatrix.RowCount = 0 Then
                                            ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = False
                                        End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Voucher"
                                        '#1008 17-09-2011
                                    Case "fo_VoView"
                                        ExportSeaFCLForm.PaneLevel = 20

                                    Case "fo_VoEdit"
                                        ExportSeaFCLForm.PaneLevel = 21
                                        ExportSeaFCLForm.Items.Item("cb_BnkName").Enabled = False
                                        ExportSeaFCLForm.Items.Item("ed_Cheque").Enabled = False
                                        ExportSeaFCLForm.Items.Item("ed_PayRate").Enabled = False
                                    Case "fo_TkrView"
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_TkrList").Specific
                                        If ExportSeaFCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            If (oMatrix.Columns.Item("V_1").Cells.Item(1).Specific.Value = "") Then
                                                ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = False
                                            Else
                                                ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                            End If
                                        End If
                                        ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Add"
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_TkrList"
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.PaneLevel = 38

                                    Case "fo_TkrEdit"
                                        ExportSeaFCLForm.PaneLevel = 39
                                        If ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Add" Then
                                            NewPOEditTab(ExportSeaFCLForm, "mx_TkrList")
                                        End If


                                        ' ExportSeaFCLForm.Items.Item("bt_GenPO").Enabled = False

                                    Case "fo_TkrSet"
                                        ExportSeaFCLForm.PaneLevel = 40
                                        ExportSeaFCLForm.Items.Item("fo_TkrSet").Specific.Select()
                                    Case "fo_PMain"
                                        ExportSeaFCLForm.PaneLevel = 8
                                    Case "fo_PCargo"
                                        ExportSeaFCLForm.PaneLevel = 9
                                    Case "fo_PCon"
                                        ExportSeaFCLForm.PaneLevel = 10
                                    Case "fo_PInv"
                                        ExportSeaFCLForm.PaneLevel = 11
                                    Case "fo_PLic"
                                        ExportSeaFCLForm.PaneLevel = 12
                                    Case "fo_PAttach"
                                        ExportSeaFCLForm.PaneLevel = 13
                                    Case "fo_PTotal"
                                        ExportSeaFCLForm.PaneLevel = 14

                                    Case "fo_ShpInv"
                                        ExportSeaFCLForm.PaneLevel = 25
                                        ExportSeaFCLForm.Items.Item("fo_ShView").Specific.Select()
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_ShpInv").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                                End If

                                            Else
                                                '    ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_ShpInv"
                                        '#1008 17-09-2011
                                    Case "fo_ShView"
                                        ExportSeaFCLForm.PaneLevel = 25
                                    Case "fo_BkVsl"
                                        ExportSeaFCLForm.PaneLevel = 26
                                        ExportSeaFCLForm.Items.Item("fo_BView").Specific.Select()
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Bok").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                                                End If

                                            Else
                                                'ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Bok"
                                        '#1008 17-09-2011
                                    Case "fo_Crate"
                                        ExportSeaFCLForm.PaneLevel = 27
                                        ExportSeaFCLForm.Items.Item("fo_Cr").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Crate"
                                        '#1008 17-09-2011
                                    Case "fo_Fumi"
                                        ExportSeaFCLForm.PaneLevel = 28
                                        ExportSeaFCLForm.Items.Item("fo_FV").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Fumi"
                                        '#1008 17-09-2011
                                    Case "fo_OpBunk"
                                        ExportSeaFCLForm.PaneLevel = 29
                                        ExportSeaFCLForm.Items.Item("fo_Buk").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Bunk"
                                        '#1008 17-09-2011
                                    Case "fo_ArmEs"
                                        ExportSeaFCLForm.PaneLevel = 30
                                        ExportSeaFCLForm.Items.Item("fo_Arm").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Armed"
                                        '#1008 17-09-2011
                                    Case "fo_PO"
                                        ExportSeaFCLForm.PaneLevel = 37
                                        ExportSeaFCLForm.Items.Item("fo_Arm").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_PO"
                                        '#1008 17-09-2011

                                    Case "fo_Yard"
                                        ExportSeaFCLForm.PaneLevel = 31


                                    Case "fo_Cont"
                                        ExportSeaFCLForm.PaneLevel = 32
                                        ExportSeaFCLForm.Items.Item("fo_ConView").Specific.Select()

                                        'oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                        'If oMatrix.RowCount > 1 Then
                                        '    ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = True
                                        'ElseIf oMatrix.RowCount = 1 Then
                                        '    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                        '        ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = True
                                        '    Else
                                        '        ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = False
                                        '    End If
                                        'End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_ConTab"
                                        '#1008 17-09-2011
                                    Case "fo_ConView"
                                        ExportSeaFCLForm.PaneLevel = 32
                                    Case "fo_ConEdit"
                                        ExportSeaFCLForm.PaneLevel = 33

                                    Case "fo_Charge"
                                        ExportSeaFCLForm.PaneLevel = 34
                                        ExportSeaFCLForm.Items.Item("fo_ChView").Specific.Select()
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                        ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = True
                                        If (oMatrix.RowCount < 1) Then
                                            ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = False
                                        ElseIf (oMatrix.RowCount < 2) Then
                                            If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = False
                                            End If
                                        End If
                                        '#1008 17-09-2011
                                        ExportSeaFCLForm.Settings.Enabled = True
                                        ExportSeaFCLForm.Settings.EnableRowFormat = True
                                        ExportSeaFCLForm.Settings.MatrixUID = "mx_Charge"
                                        '#1008 17-09-2011
                                    Case "fo_ChView"
                                        ExportSeaFCLForm.PaneLevel = 34
                                    Case "fo_ChEdit"
                                        ExportSeaFCLForm.PaneLevel = 35

                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                        If (oMatrix.RowCount > 0) Then
                                            If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                ExportSeaFCLForm.Items.Item("ed_CSeqNo").Specific.Value = 1
                                            Else
                                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oRecordSet.DoQuery("select count(*)  from [@OBT_FCL17_HCHARGES] where U_ChSeqNo=  '" + oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + "'")
                                                If oRecordSet.RecordCount > 0 And IsAmend = False Then
                                                    ExportSeaFCLForm.Items.Item("ed_CSeqNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                                End If
                                            End If
                                        Else
                                            ExportSeaFCLForm.Items.Item("ed_CSeqNo").Specific.Value = 1
                                        End If
                                        ExportSeaFCLForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges"
                                        ExportSeaFCLForm.Items.Item("bt_DelCh").Enabled = False

                                        oCombo = ExportSeaFCLForm.Items.Item("cb_Claim").Specific
                                        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                        ClearText(ExportSeaFCLForm, "ed_ChCode", "ed_Remarks")


                                End Select

                                If pVal.FormTypeEx = "EXPORTSEAFCL" Then
                                    ' MSW New Button at Choose From List 22-03-2011
                                    If p_oSBOApplication.Forms.ActiveForm.Title = "List of VESSEL" Then
                                        If pVal.ItemUID = "ed_Vessel" Then
                                            AddNewBtToCFLFrm(p_oSBOApplication.Forms.ActiveForm, "bt_VNew")
                                        End If
                                        'If pVal.ItemUID = "ed_Voy" Then
                                        '    AddNewBtToCFLFrm(p_oSBOApplication.Forms.ActiveForm, "bt_VoNew")
                                        'End If
                                    End If

                                    ' MSW New Button at Choose From List 22-03-2011
                                End If

                                '========== Bunker ========='''Bunker
                                If pVal.ItemUID = "bt_Bunk" Then
                                    'LoadFumigation(ExportSeaFCLForm, "Bunker.srf", "BUNKER")
                                    If Not AlreadyExist("BUNKER") Then
                                        LoadFumigation(ExportSeaFCLForm, "Bunker.srf", "BUNKER")
                                    Else
                                        p_oSBOApplication.Forms.Item("BUNKER").Select()

                                    End If
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                    'modFumigation.CreateDTDetail(FumigationForm)
                                    oRecordSet.DoQuery("Select * from [@BUNKER] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")

                                    If oRecordSet.RecordCount > 0 Then
                                        FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        FumigationForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        'FumigationForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("U_JobNo").Value
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If


                                End If
                                '========== Toll ========='''Toll
                                If pVal.ItemUID = "bt_Toll" Then

                                    '        LoadFumigation(ExportSeaFCLForm, "Toll.srf", "TOLL")
                                    If Not AlreadyExist("TOLL") Then
                                        LoadFumigation(ExportSeaFCLForm, "Toll.srf", "TOLL")
                                    Else
                                        p_oSBOApplication.Forms.Item("TOLL").Select()

                                    End If
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                    oRecordSet.DoQuery("Select * from [@Toll] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")

                                    If oRecordSet.RecordCount > 0 Then
                                        FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        FumigationForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        'FumigationForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("U_JobNo").Value
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If


                                End If

                                '========== Crane ========='''button
                                If pVal.ItemUID = "bt_Crane" Then
                                    If Not AlreadyExist("CRANE") Then
                                        LoadFumigation(ExportSeaFCLForm, "Crane.srf", "CRANE")
                                    Else
                                        p_oSBOApplication.Forms.Item("CRANE").Select()

                                    End If
                                    'LoadFumigation(ExportSeaFCLForm, "Crane.srf", "CRANE")
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                    'modFumigation.CreateDTDetail(FumigationForm)
                                    oRecordSet.DoQuery("Select * from [@CRANE] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")

                                    If oRecordSet.RecordCount > 0 Then
                                        FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        FumigationForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        'FumigationForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("U_JobNo").Value
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If


                                End If
                                If pVal.ItemUID = "bt_DGD" Then
                                    CreateDGDPDF(ExportSeaFCLForm)
                                End If
                                '========== Fumigation ========='
                                If pVal.ItemUID = "bt_Fumi" Then
                                    If Not AlreadyExist("FUMIGATION") Then
                                        LoadFumigation(ExportSeaFCLForm, "Fumigation.srf", "FUMIGATION")
                                    Else
                                        p_oSBOApplication.Forms.Item("FUMIGATION").Select()

                                    End If
                                    'LoadFumigation(ExportSeaFCLForm, "Fumigation.srf", "FUMIGATION")
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    oRecordSet.DoQuery("Select * from [@FUMIGATION] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        FumigationForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        'FumigationForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("U_JobNo").Value
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If



                                End If
                                '========== End Fumigation ========='

                                If pVal.ItemUID = "479" Then
                                    ExportSeaFCLForm.Items.Item("fo_Cont").Specific.Select()
                                End If

                                If pVal.ItemUID = "bt_Foklit" Then
                                    LoadFumigation(ExportSeaFCLForm, "Forklift.srf", "FORKLIFT")
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                    'modFumigation.CreateDTDetail(FumigationForm)
                                    oRecordSet.DoQuery("Select * from [@FORKLIFT] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")

                                    If oRecordSet.RecordCount > 0 Then
                                        FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        FumigationForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        'FumigationForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("U_JobNo").Value
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Crate" Then 'to combine

                                    LoadFumigation(ExportSeaFCLForm, "Crate.srf", "CRATE")
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                    'modFumigation.CreateDTDetail(FumigationForm)
                                    oRecordSet.DoQuery("Select * from [@CRATE] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")

                                    If oRecordSet.RecordCount > 0 Then
                                        FumigationForm = p_oSBOApplication.Forms.ActiveForm
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        FumigationForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        'FumigationForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("U_JobNo").Value
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Courier" Then
                                    LoadButtonForm(ExportSeaFCLForm, "CourierExport.srf", "COURIEREXPORT")
                                    oRecordSet.DoQuery("Select * from  [@COURIEREXPORT] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        CourForm = p_oSBOApplication.Forms.ActiveForm
                                        CourForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        CourForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        CourForm.Items.Item("1").Click()
                                        CourForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        CourForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If

                                If pVal.ItemUID = "bt_DGLP" Then
                                    LoadButtonForm(ExportSeaFCLForm, "DGLPExport.srf", "DGLPEXPORT")
                                    oRecordSet.DoQuery("Select * from  [@DGLPEXPORT] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        DGLPForm = p_oSBOApplication.Forms.ActiveForm
                                        DGLPForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        DGLPForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        DGLPForm.Items.Item("1").Click()
                                        DGLPForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        DGLPForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Outer" Then
                                    If Not AlreadyExist("OUTRIDER") Then
                                        LoadFumigation(ExportSeaFCLForm, "Outrider.srf", "OUTRIDER")
                                    Else
                                        p_oSBOApplication.Forms.Item("OUTRIDER").Select()

                                    End If
                                    'LoadFumigation(ExportSeaFCLForm, "Outrider.srf", "OUTRIDER")

                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                    oRecordSet.DoQuery("Select * from  [@OUTRIDER] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        OutForm = p_oSBOApplication.Forms.ActiveForm
                                        OutForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        OutForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        OutForm.Items.Item("1").Click()
                                        OutForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        OutForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If

                                If pVal.ItemUID = "bt_Certifi" Then
                                    LoadButtonForm(ExportSeaFCLForm, "CertificateOfOriginExport.srf", "COOEXPORT")
                                    oRecordSet.DoQuery("Select * from  [@COOEXPORT] Where U_DocNum='" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        COOForm = p_oSBOApplication.Forms.ActiveForm
                                        COOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        COOForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        COOForm.Items.Item("1").Click()
                                        COOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        COOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Alert" Then
                                    PDFAlert(p_oSBOApplication.Forms.ActiveForm)
                                End If
                                If pVal.ItemUID = "bt_Excel" Then
                                    If pVal.BeforeAction = False Then
                                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                            ActiveForm = p_oSBOApplication.Forms.Item(pVal.FormUID.ToString())
                                            Dim excel As Microsoft.Office.Interop.Excel.Application
                                            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
                                            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet
                                            Dim CustomerCode As String

                                            Try
                                                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oRecordSet.DoQuery("Select * from OINV")
                                                CustomerCode = oRecordSet.Fields.Item("CardCode").Value

                                                excel = New Microsoft.Office.Interop.Excel.Application
                                                If Not Directory.Exists("C:\Billing_Summary") Then
                                                    Directory.CreateDirectory("C:\Billing_Summary")
                                                End If
                                                Dim Source As String = Application.StartupPath.ToString() & "\00 EX-SLB master 09.02.xls"
                                                Dim filepath As String = "C:\Billing_Summary" & "\00 EX-SLB master 09.02.xls"
                                                If File.Exists(filepath) Then
                                                    wBook = excel.Workbooks.Open(filepath)
                                                Else
                                                    CopyFile(Source, "C:\Billing_Summary\00 EX-SLB master 09.02.xls")
                                                    wBook = excel.Workbooks.Open(filepath)

                                                End If
                                                wSheet = wBook.Sheets("x out Loyang jetty")
                                                wSheet.Cells(8, 3) = CustomerCode
                                                excel.Visible = True
                                                'wSheet.Activate()

                                                Marshal.ReleaseComObject(excel)
                                                Marshal.ReleaseComObject(wBook)
                                                Marshal.ReleaseComObject(wSheet)


                                            Catch ex As COMException
                                                MessageBox.Show("Error accessing Excel: " + ex.ToString())

                                            Catch ex As Exception
                                                MessageBox.Show("Error: " + ex.ToString())

                                            End Try
                                        ElseIf pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE Then
                                            System.Windows.Forms.Application.Exit()
                                        End If
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Label" Then
                                    LoadFormXML("ImageN.srf")
                                    ActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    LoadData()
                                End If


                                If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False

                                End If

                                If pVal.ItemUID = "ch_POD" Then
                                    If ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = True Then
                                        'Check PO Status and Draft Voucher Status
                                        If CheckPOandVoucherStatus(ExportSeaFCLForm) = True Then
                                            p_oSBOApplication.MessageBox("Need To Closed Open PO First.")
                                            ExportSeaFCLForm.Items.Item("ch_POD").Enabled = True
                                            ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                            ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = False
                                        Else
                                            Dim chk As Integer = 0
                                            chk = p_oSBOApplication.MessageBox("You cannot change this Job after you have closed it.Continue?", 1, "Yes", "No")
                                            If chk = 1 Then
                                                ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = True
                                                ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index)
                                                ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                                                ExportSeaFCLForm.Items.Item("ed_xRef").Specific.Active = True
                                                ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = False
                                            Else
                                                ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = False
                                                ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                ExportSeaFCLForm.Items.Item("ch_POD").Enabled = True
                                            End If

                                        End If

                                    End If
                                    'If ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = False Then
                                    '    ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    'End If
                                End If

                                If pVal.ItemUID = "bt_AddLic" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_License").Specific
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL08_EPLICINFO").Clear()                          '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("colLicNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                End If

                                If pVal.ItemUID = "bt_DelLic" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_License").Specific
                                    Dim lRow As Long
                                    lRow = oMatrix.GetNextSelectedRow
                                    If lRow > -1 Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL08_EPLICINFO").RemoveRecord(0)   '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("OBT_FCL08_EPLICINFO")                                '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
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
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Cont").Specific
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL06_EPCONT").Clear()                         '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                    oMatrix.AddRow(1)
                                    oMatrix.FlushToDataSource()
                                    oMatrix.Columns.Item("colCSeqNo").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                End If

                                If pVal.ItemUID = "bt_DelCon" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Cont").Specific
                                    Dim lRow As Long
                                    lRow = oMatrix.GetNextSelectedRow
                                    If lRow > -1 Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL06_EPCONT").RemoveRecord(0)    '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
                                            'oMatrix.AddRow(1)
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("OBT_FCL06_EPCONT")                                  '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
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
                                'MSW to edit 10-09-2011

                                'If pVal.ItemUID = "ch_Dsp" Then
                                '    If ExportSeaFCLForm.Items.Item("ch_Dsp").Specific.Checked = True Then
                                '        ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                '        ExportSeaFCLForm.Items.Item("ed_DspCDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                '        ExportSeaFCLForm.Items.Item("ed_DspCHr").Specific.Value = Now.ToString("HH:mm")
                                '        If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                '    End If
                                '    If ExportSeaFCLForm.Items.Item("ch_Dsp").Specific.Checked = False Then
                                '        ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific.Value = ""
                                '        ExportSeaFCLForm.Items.Item("ed_DspCDay").Specific.Value = ""
                                '        ExportSeaFCLForm.Items.Item("ed_DspCHr").Specific.Value = ""
                                '        ExportSeaFCLForm.Items.Item("cb_Dspchr").Specific.Active = True
                                '    End If
                                'End If

                                If pVal.ItemUID = "ed_ETDDate" And ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific.String <> String.Empty Then
                                    'If DateTime(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific, ExportSeaFCLForm.Items.Item("ed_ETDDay").Specific, ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Try
                                        If Not DateTimeWithoutDay(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific, ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific.Value

                                        ExportSeaFCLForm.Items.Item("ed_ATime").Specific.Value = ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific.Value
                                        'ExportSeaFCLForm.ActiveItem = "ed_CrgDsc"
                                    Catch ex As Exception
                                        'ExportSeaFCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific.Value
                                        'ExportSeaFCLForm.Items.Item("ed_ADay").Specific.Value = ExportSeaFCLForm.Items.Item("ed_ETDDay").Specific.Value
                                        'ExportSeaFCLForm.Items.Item("ed_ATime").Specific.Value = ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific.Value
                                        'ExportSeaFCLForm.ActiveItem = "ed_CrgDsc"
                                    End Try
                                    If HolidayMarkUpWithoutDay(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaFCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific.Value
                                    If Not DateTimeWithoutDay(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific, ExportSeaFCLForm.Items.Item("ed_ATime").Specific) Then Throw New ArgumentException(sErrDesc)
                                End If


                                If pVal.ItemUID = "ed_CunDate" And ExportSeaFCLForm.Items.Item("ed_CunDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_CunDate").Specific, ExportSeaFCLForm.Items.Item("ed_CunDay").Specific, ExportSeaFCLForm.Items.Item("ed_CunTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_CunDay").Specific, ExportSeaFCLForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_JbDate" And ExportSeaFCLForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_JbDate").Specific, ExportSeaFCLForm.Items.Item("ed_JbDay").Specific, ExportSeaFCLForm.Items.Item("ed_JbHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_JbDay").Specific, ExportSeaFCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                'If pVal.ItemUID = "ed_DspDate" And ExportSeaFCLForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                                '    If Not DateTime(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspDate").Specific, ExportSeaFCLForm.Items.Item("ed_DspDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                '    If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_DspDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                'End If
                                'If pVal.ItemUID = "ed_DspCDte" And ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                                '    If Not DateTime(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific, ExportSeaFCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspCHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                '    If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                'End If

                                If pVal.ItemUID = "bt_Payment" Then
                                    p_oSBOApplication.ActivateMenuItem("2818")
                                End If

                                If pVal.ItemUID = "op_Inter" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    'MSW 14-09-2011 Truck PO
                                    If ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value <> "" Then
                                        If Not CancelTruckingPurchaseOrder(Convert.ToInt32(ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value.ToString)) Then Throw New ArgumentException(sErrDesc)
                                        If Not UpdateForCancelStatus(ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                        sql = "select U_POStatus,U_PO from [@OBT_TB08_FFCPO] Where U_PONo = '" & ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value & "'"
                                        If Not EditPOTab(ExportSeaFCLForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Items.Item("ed_Trucker").Enabled = True
                                    End If
                                    'MSW 14-09-2011 Truck PO
                                    ExportSeaFCLForm.Items.Item("bt_PO").Enabled = False
                                    ClearText(ExportSeaFCLForm, "ed_Trucker", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_PODocNo", "ed_VehicNo", "ed_Attent", "ed_PONo", "ed_EUC", "ed_PO", "ed_PStus", "ed_TMulti", "ed_TOrigin")
                                    If AddChooseFromListByOption(ExportSeaFCLForm, True, "ed_Trucker", "TKRINTR", "CFLTKRE", "TKREXTR", "CFLTKRV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaFCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                    ExportSeaFCLForm.Freeze(False)
                                ElseIf pVal.ItemUID = "op_Exter" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    If ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value = "" Then 'MSW 14-09-2011 Truck PO
                                        ClearText(ExportSeaFCLForm, "ed_Trucker", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_EUC", "ed_TMulti", "ed_TOrigin")
                                        ExportSeaFCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                        ExportSeaFCLForm.Items.Item("bt_PO").Enabled = True
                                        If AddChooseFromListByOption(ExportSeaFCLForm, False, "ed_Trucker", "TKRINTR", "CFLTKRE", "TKREXTR", "CFLTKRV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Freeze(False)
                                    End If
                                End If

                                If pVal.ItemUID = "op_DInter" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    'MSW 14-09-2011 Truck PO
                                    If ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value <> "" Then
                                        If Not CancelTruckingPurchaseOrder(Convert.ToInt32(ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value.ToString)) Then Throw New ArgumentException(sErrDesc)
                                        If Not UpdateForCancelStatus(ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                        sql = "select U_POStatus,U_PO from [@OBT_TB08_FFCPO] Where U_PONo = '" & ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value & "'"
                                        If Not EditPOTab(ExportSeaFCLForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Items.Item("ed_Dspatch").Enabled = True
                                    End If
                                    'MSW 14-09-2011 Truck PO
                                    ClearText(ExportSeaFCLForm, "ed_Dspatch", "ed_DspTel", "ed_DFax", "ed_DEmail", "ed_DPDocNo", "ed_DPO", "ed_DPStus", "ed_DAttent", "ed_DPONo", "ed_DEUC", "ed_DMulti", "ed_DOrigin")
                                    ExportSeaFCLForm.Items.Item("bt_DPO").Enabled = False
                                    If AddChooseFromListByOption(ExportSeaFCLForm, True, "ed_Dspatch", "DSPINTR", "CFLDSP", "DSPEXTR", "CFLDSPV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaFCLForm.Items.Item("ed_Dspatch").Specific.Active = True
                                    ExportSeaFCLForm.Freeze(False)
                                ElseIf pVal.ItemUID = "op_DExter" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    If ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value = "" Then 'MSW 14-09-2011 Truck PO
                                        ClearText(ExportSeaFCLForm, "ed_Dspatch", "ed_DspTel", "ed_DFax", "ed_DEmail", "ed_DEUC", "ed_DMulti", "ed_DOrigin")
                                        ExportSeaFCLForm.Items.Item("bt_DPO").Enabled = True
                                        If AddChooseFromListByOption(ExportSeaFCLForm, False, "ed_Dspatch", "DSPINTR", "CFLDSP", "DSPEXTR", "CFLDSPV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Items.Item("ed_Dspatch").Specific.Active = True
                                        ExportSeaFCLForm.Freeze(False)
                                    End If
                                End If

                                'If pVal.ItemUID = "bt_GenPO" Then
                                '    'Purchase Order and Goods Receipt POP UP 

                                '    ' ==================================== Creating Custom Purchase Order ==============================
                                '    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                '        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                '            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                '            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                '        End If
                                '        LoadTruckingPO(ExportSeaFCLForm, "TkrListPurchaseOrder.srf")
                                '    Else
                                '        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                '        Exit Function
                                '    End If

                                '    ' ==================================== Creating Custom Purchase Order ==============================

                                '    'p_oSBOApplication.Menus.Item("2305").Activate()
                                '    'p_oSBOApplication.ActivateMenuItem("6913") 'MSW 04-04-2011
                                '    'p_oSBOApplication.Menus.Item("6913").Activate()
                                '    ''SBO_Application.Menus.Item("6913").Activate()
                                '    'Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                                '    'UDFAttachForm.Items.Item("U_JobNo").Specific.Value = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                '    'UDFAttachForm.Items.Item("U_InsDate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value
                                'End If

                                'Create Dispatch PO
                                If pVal.ItemUID = "bt_DPO" Then
                                    If ExportSeaFCLForm.Items.Item("ed_DspCode").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("There is no Vendor to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf ExportSeaFCLForm.Items.Item("ed_DICode").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("There is no Item to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        If Not CreateDspGenPO(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspCode").Specific.Value, ExportSeaFCLForm.Items.Item("ed_DICode").Specific.Value) Then Throw New ArgumentException(sErrDesc)

                                        sql = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                             "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(ExportSeaFCLForm.Items.Item("ed_DPDocNo").Specific.Value)


                                        If Not PopulateDispatchPOToEditTab(ExportSeaFCLForm, sql) Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Items.Item("bt_DspAdd").Click()
                                        SavePOPDFInEditTab(ExportSeaFCLForm, "Dispatch") 'Button
                                        ExportSeaFCLForm.Items.Item("ed_Dspatch").Enabled = False
                                        ExportSeaFCLForm.Items.Item("op_DInter").Enabled = False
                                        ExportSeaFCLForm.Items.Item("op_DExter").Enabled = False
                                    End If
                                End If

                                'Create Trucking PO
                                If pVal.ItemUID = "bt_PO" Then
                                    If ExportSeaFCLForm.Items.Item("ed_TkrCode").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("There is no Vendor to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf ExportSeaFCLForm.Items.Item("ed_POICode").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("There is no Item to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        If Not CreateGenPO(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_TkrCode").Specific.Value, ExportSeaFCLForm.Items.Item("ed_POICode").Specific.Value) Then Throw New ArgumentException(sErrDesc)

                                        sql = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                             "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(ExportSeaFCLForm.Items.Item("ed_PODocNo").Specific.Value)


                                        If Not PopulateTruckingPOToEditTab(ExportSeaFCLForm, sql) Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Items.Item("bt_TkrAdd").Click()
                                        SavePOPDFInEditTab(ExportSeaFCLForm, "Trucking") 'Button
                                        ExportSeaFCLForm.Items.Item("ed_Trucker").Enabled = False
                                        ExportSeaFCLForm.Items.Item("op_Inter").Enabled = False
                                        ExportSeaFCLForm.Items.Item("op_Exter").Enabled = False
                                    End If

                                End If

                                If pVal.ItemUID = "1" Then
                                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If pVal.ActionSuccess = True Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            ExportSeaFCLForm.Items.Item("ed_Code").Enabled = False 'MSW
                                            ExportSeaFCLForm.Items.Item("ed_Name").Enabled = False
                                            ExportSeaFCLForm.Items.Item("ed_JobNo").Enabled = False
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            
                                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                            AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCNo1", "@OBT_FCL19_HCONTAINE")
                                            oCombo = oMatrix.Columns.Item("colCSize1").Cells.Item(oMatrix.RowCount).Specific
                                            oCombo.Select("20'", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                            oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                            AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCCode1", "@OBT_FCL17_HCHARGES")
                                            oCombo = oMatrix.Columns.Item("colCClaim1").Cells.Item(oMatrix.RowCount).Specific
                                            oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_DrBL").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_SpOrder").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_CrPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_CPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_BunkPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_ArmePO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ' ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                            'ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                                            'ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

                                            
                                          
                                        End If

                                    End If
                                    If pVal.FormMode = 1 And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                                        If pVal.ActionSuccess = True Then
                                            Dim JobType As String = ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value
                                            'Google Doc
                                            If JobType.Contains("Import") Then
                                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                                EnableExportImport(ExportSeaFCLForm, False, True)
                                            ElseIf JobType.Contains("Export") Then
                                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                                EnableExportImport(ExportSeaFCLForm, True, False)
                                            ElseIf JobType.Contains("Local") Then
                                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                                EnableExportImport(ExportSeaFCLForm, True, True)
                                            ElseIf JobType.Contains("Transhipment") Then
                                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                                EnableExportImport(ExportSeaFCLForm, True, True)
                                            End If
                                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                            AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCNo1", "@OBT_FCL19_HCONTAINE")
                                            oCombo = oMatrix.Columns.Item("colCSize1").Cells.Item(oMatrix.RowCount).Specific
                                            oCombo.Select("20'", SAPbouiCOM.BoSearchKey.psk_ByValue)

                                            oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                            AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCCode1", "@OBT_FCL17_HCHARGES")
                                            oCombo = oMatrix.Columns.Item("colCClaim1").Cells.Item(oMatrix.RowCount).Specific
                                            oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_DrBL").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_SpOrder").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_CrPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_CPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_BunkPO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_ArmePO").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ' ExportSeaFCLForm.Items.Item("bt_PVoc").Enabled = True
                                            ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                            'ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                                            'ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

                                            'Check PO
                                            If ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Closed" Or ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Cancelled" Then
                                                ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                                                ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = False
                                            ElseIf ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Open" Then
                                                ExportSeaFCLForm.Items.Item("ch_POD").Enabled = True
                                                ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = True
                                                ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                                            End If

                                        End If
                                    End If
                                End If
                                'KM to edit
                                If pVal.ItemUID = "bt_PrntDis" Then
                                    'PreviewInsDoc(ExportSeaFCLForm)
                                    Dim DocNum As Integer
                                    Dim InsDoc As Integer
                                    rptDocument = New ReportDocument
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_DspList").Specific
                                    pdfFilename = "Dispatch Instruction"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                    rptPath = Application.StartupPath.ToString & "\Dispatch Instruction.rpt"
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    rptDocument.Load(rptPath)
                                    rptDocument.Refresh()
                                    DocNum = Convert.ToInt32(ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value)
                                    InsDoc = Convert.ToInt32(oMatrix.Columns.Item("colInsDoc").Cells.Item(selectedRow).Specific.Value.ToString)
                                    rptDocument.SetParameterValue("@DocEntry", DocNum)
                                    rptDocument.SetParameterValue("@InsDocNo", InsDoc)
                                    reportuti.SetDBLogIn(rptDocument)
                                    If Not pdffilepath = String.Empty Then
                                        reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
                                    End If
                                End If
                                If pVal.ItemUID = "bt_A6Label" Then
                                    PreviewA6Label(ExportSeaFCLForm)
                                End If


                                If pVal.ItemUID = "bt_DspDPO" Then
                                    LoadDetachJobNormalForm(ExportSeaFCLForm, "DetachJobNormal.srf", "Dispatch", "ed_DPO")
                                    ActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    modDetachJobNormal.LoadDetachJForm()
                                End If


                                If pVal.ItemUID = "bt_Dspview" Then
                                    PreviewDispatchInstruction(ExportSeaFCLForm)
                                End If
                                If pVal.ItemUID = "bt_DspMul" Then
                                    LoadMultiJobNormalForm(ExportSeaFCLForm, "MultiJobNormal.srf", "Dispatch")
                                    ActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    modMultiJobForNormal.LoadMultiJForm()
                                End If
                                If pVal.ItemUID = "bt_DspNew" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    ClearText(ExportSeaFCLForm, "ed_DInsDoc", "ed_DMulti", "ed_DOrigin", "ed_DPDocNo", "ed_DPONo", "ed_DPO", "ed_DPStus", "ed_DInDate", "ed_Dspatch", "ed_DEUC", "ed_DAttent", "ed_DspTel", "ed_DFax", "ed_DEmail", "ed_DspDate", "ed_DspTime", "ee_DspIns", "ee_DIRmsk", "ee_DRmsk") 'MSW New Ticket 07-09-2011
                                    ExportSeaFCLForm.Freeze(False)
                                    ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Add"
                                    NewPOEditTab(ExportSeaFCLForm, "mx_DspList")
                                End If

                                If pVal.ItemUID = "bt_DspAdd" Then
                                    Dim poMatrix As SAPbouiCOM.Matrix
                                    If String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_Dspatch").Specific.String) Then
                                        p_oSBOApplication.SetStatusBarMessage("Must Fill Dispatcher", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf ExportSeaFCLForm.Items.Item("op_DExter").Specific.Selected = True And ExportSeaFCLForm.Items.Item("ed_DPO").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("Need to Create PO for External.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        poMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_DspList").Specific
                                        If ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Add" Then
                                            AddUpdateInstructions(ExportSeaFCLForm, oMatrix, "@OBT_FCL04_EDISPATCH", True)
                                            If ExportSeaFCLForm.Items.Item("op_DExter").Specific.Selected = True Then
                                                oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                                SaveToPOTab(ExportSeaFCLForm, poMatrix, True, ExportSeaFCLForm.Items.Item("ed_DPO").Specific.Value, ExportSeaFCLForm.Items.Item("ed_DPDocNo").Specific.Value, ExportSeaFCLForm.Items.Item("ed_Dspatch").Specific.Value, ExportSeaFCLForm.Items.Item("ed_DDate").Specific.Value, "Dispatch", ExportSeaFCLForm.Items.Item("ed_DPStus").Specific.Value)
                                                ExportSeaFCLForm.Items.Item("ed_Dspatch").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_DInter").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_DExter").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_DPO").Enabled = False
                                            End If
                                        Else
                                            AddUpdateInstructions(ExportSeaFCLForm, oMatrix, "@OBT_FCL04_EDISPATCH", False)
                                            If ExportSeaFCLForm.Items.Item("op_DExter").Specific.Selected = True Then
                                                EditPOInEditTab(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value, "mx_DspList") '
                                                ExportSeaFCLForm.Items.Item("ed_Dspatch").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_DInter").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_DExter").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_DPO").Enabled = False
                                            End If

                                            '  ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Add"
                                        End If
                                        'ClearText(ExportSeaFCLForm, "ed_DInsDoc", "ed_DPDocNo", "ed_DPONo", "ed_DPO", "ed_DPStus", "ed_DInDate", "ed_Dspatch", "ed_DEUC", "ed_DAttent", "ed_DspTel", "ed_DFax", "ed_DEmail", "ed_DspDate", "ed_DspTime", "ee_DspIns", "ee_DIRmsk", "ee_DRmsk") 'MSW New Ticket 07-09-2011
                                        'ExportSeaFCLForm.Items.Item("fo_DspView").Specific.Select()
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                        End If
                                        ExportSeaFCLForm.Items.Item("bt_DspAdd").Specific.Caption = "Update"
                                        ExportSeaFCLForm.Items.Item("1").Click()
                                        SaveDispatchInstruction(ExportSeaFCLForm) 'to km
                                        ExportSeaFCLForm.Items.Item("bt_DPO").Enabled = False 'to km
                                        If ExportSeaFCLForm.Items.Item("ed_DMulti").Specific.Value = "Y" Then
                                            ExportSeaFCLForm.Items.Item("bt_DspDPO").Enabled = True
                                        Else
                                            ExportSeaFCLForm.Items.Item("bt_DspDPO").Enabled = False
                                        End If
                                    End If
                                End If

                                If pVal.ItemUID = "bt_PrntPO" Then
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                    sql = "Select DocEntry From OPOR where DocNum='" & oMatrix.Columns.Item("colPONo").Cells.Item(selectedRow).Specific.Value.ToString & "'"
                                    oRecordSet.DoQuery(sql)
                                    If oRecordSet.RecordCount > 0 Then
                                        PreviewPOFromPOTab(ExportSeaFCLForm, Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value.ToString))
                                    End If

                                End If
                                If pVal.ItemUID = "bt_PrntIns" Then
                                    Dim DocNum As Integer
                                    Dim InsDoc As Integer
                                    rptDocument = New ReportDocument
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_TkrList").Specific
                                    pdfFilename = "Trucking Instruction"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                    If ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-LCL") Or ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Air") Or _
                                        ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Land") Or ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-LCL") Or _
                                        ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Air") Or ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Land") Or _
                                        ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Transhipment") Or ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Local") Then
                                        rptPath = Application.StartupPath.ToString & "\Trucking Instruction.rpt"
                                    ElseIf ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-FCL") Or ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-FCL") Then
                                        rptPath = Application.StartupPath.ToString & "\Trucking Instruction FCL.rpt"
                                    End If

                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    rptDocument.Load(rptPath)
                                    rptDocument.Refresh()

                                    DocNum = Convert.ToInt32(ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value)
                                    InsDoc = Convert.ToInt32(oMatrix.Columns.Item("colInsDoc").Cells.Item(selectedRow).Specific.Value.ToString)

                                    rptDocument.SetParameterValue("@DocEntry", DocNum)
                                    rptDocument.SetParameterValue("@InsDocNo", InsDoc)
                                    reportuti.SetDBLogIn(rptDocument)
                                    If Not pdffilepath = String.Empty Then
                                        reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
                                    End If
                                End If

                                If pVal.ItemUID = "bt_TkrDPO" Then
                                    LoadDetachJobNormalForm(ExportSeaFCLForm, "DetachJobNormal.srf", "Trucking", "ed_PO")
                                    ActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    modDetachJobNormal.LoadDetachJForm()
                                End If

                                If pVal.ItemUID = "bt_Tkview" Then
                                    PreviewInsDoc(ExportSeaFCLForm)
                                End If
                                If pVal.ItemUID = "bt_TkrMul" Then
                                    LoadMultiJobNormalForm(ExportSeaFCLForm, "MultiJobNormal.srf", "Trucking")
                                    ActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    modMultiJobForNormal.LoadMultiJForm()
                                End If
                                If pVal.ItemUID = "bt_TkrNew" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    ClearText(ExportSeaFCLForm, "ed_InsDoc", "ed_TMulti", "ed_TOrigin", "ed_PODocNo", "ed_PONo", "ed_PO", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk", "ee_ColFrm") 'MSW New Ticket 07-09-2011
                                    ExportSeaFCLForm.Freeze(False)

                                    ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Add"
                                    NewPOEditTab(ExportSeaFCLForm, "mx_TkrList")
                                End If

                                If pVal.ItemUID = "bt_TkrAdd" Then
                                    Dim poMatrix As SAPbouiCOM.Matrix
                                    If String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_Trucker").Specific.String) Then
                                        p_oSBOApplication.SetStatusBarMessage("Must Fill Trucker", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf ExportSeaFCLForm.Items.Item("op_Exter").Specific.Selected = True And ExportSeaFCLForm.Items.Item("ed_PO").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("Need to Create PO for External.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        poMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_TkrList").Specific
                                        If ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Add" Then
                                            modTrucking.AddUpdateInstructions(ExportSeaFCLForm, oMatrix, "@OBT_FCL03_ETRUCKING", True)
                                            If ExportSeaFCLForm.Items.Item("op_Exter").Specific.Selected = True Then
                                                oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                                SaveToPOTab(ExportSeaFCLForm, poMatrix, True, ExportSeaFCLForm.Items.Item("ed_PO").Specific.Value, ExportSeaFCLForm.Items.Item("ed_PODocNo").Specific.Value, ExportSeaFCLForm.Items.Item("ed_Trucker").Specific.Value, ExportSeaFCLForm.Items.Item("ed_Date").Specific.Value, "Trucking", ExportSeaFCLForm.Items.Item("ed_PStus").Specific.Value)
                                                ExportSeaFCLForm.Items.Item("ed_Trucker").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_PO").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_Inter").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_Exter").Enabled = False
                                            End If

                                        Else
                                            modTrucking.AddUpdateInstructions(ExportSeaFCLForm, oMatrix, "@OBT_FCL03_ETRUCKING", False)    '
                                            If ExportSeaFCLForm.Items.Item("op_Exter").Specific.Selected = True Then
                                                EditPOInEditTab(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value, "mx_TkrList") '
                                                ExportSeaFCLForm.Items.Item("ed_Trucker").Enabled = False
                                                ExportSeaFCLForm.Items.Item("bt_PO").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_Inter").Enabled = False
                                                ExportSeaFCLForm.Items.Item("op_Exter").Enabled = False
                                            End If

                                            ' ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Add"

                                        End If
                                        ' ClearText(ExportSeaFCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_PO", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk", "ee_ColFrm") 'MSW New Ticket 07-09-2011
                                        'ExportSeaFCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        ExportSeaFCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Update"
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close"
                                        End If
                                        ExportSeaFCLForm.Items.Item("1").Click() 'MSW to edit New Ticket 07-09-2011
                                        SaveInsDoc(ExportSeaFCLForm) 'to km
                                        ExportSeaFCLForm.Items.Item("bt_PO").Enabled = False 'to km
                                        If ExportSeaFCLForm.Items.Item("ed_TMulti").Specific.Value = "Y" Then
                                            ExportSeaFCLForm.Items.Item("bt_TkrDPO").Enabled = True
                                        Else
                                            ExportSeaFCLForm.Items.Item("bt_TkrDPO").Enabled = False
                                        End If
                                    End If
                                End If


                                '-------------------------Payment Vouncher (OMM)------------------------------------------------------'
                                If pVal.ItemUID = "fo_VoEdit" Then
                                    ExportSeaFCLForm.Freeze(True)
                                    ExportSeaFCLForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    ExportSeaFCLForm.Items.Item("op_Cash").Specific.Selected = True
                                    ExportSeaFCLForm.Items.Item("ed_PJobNo").Specific.Value = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                    ExportSeaFCLForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()
                                    If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    'ImportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_VOUCHER").SetValue("U_DocNo", 0, ImportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value)

                                    If ExportSeaFCLForm.Items.Item("bt_Draft").Specific.Caption = "Save To Draft" Then
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_ChCode").Specific
                                        dtmatrix = ExportSeaFCLForm.DataSources.DataTables.Item("DTCharges")
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                        oMatrix.Clear()
                                    End If

                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                                    If ExportSeaFCLForm.Items.Item("ed_VocNo").Specific.Value = "" Then
                                        If (oMatrix.RowCount > 0) Then
                                            If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                                ExportSeaFCLForm.Items.Item("ed_VocNo").Specific.Value = 1
                                            Else
                                                ExportSeaFCLForm.Items.Item("ed_VocNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                            End If
                                        Else
                                            ExportSeaFCLForm.Items.Item("ed_VocNo").Specific.Value = 1
                                        End If
                                    End If


                                    oCombo = ExportSeaFCLForm.Items.Item("cb_PayCur").Specific
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
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ChCode").Specific
                                    If dtmatrix.Rows.Count = 0 Then
                                        RowAddToMatrix(ExportSeaFCLForm, oMatrix)
                                    End If
                                    ExportSeaFCLForm.Items.Item("ed_VedName").Specific.Active = True
                                    ExportSeaFCLForm.Freeze(False)
                                End If

                                'If pVal.ItemUID = "mx_Voucher" Then
                                '    oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                                '    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                '    oMatrix.SelectRow(pVal.Row, True, False)
                                '    vocRow = pVal.Row
                                'End If

                                If pVal.ItemUID = "bt_PVoc" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
                                    Dim PVNo As Integer
                                    rptDocument = New ReportDocument
                                    pdfFilename = "PAYMENT VOUCHER"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                    rptPath = Application.StartupPath.ToString & "\Payment Voucher.rpt"
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    rptDocument.Load(rptPath)
                                    rptDocument.Refresh()
                                    'If Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    PVNo = Convert.ToInt32(oMatrix.Columns.Item("colVDocNum").Cells.Item(selectedRow).Specific.Value.ToString)
                                    'Else
                                    'PONo = DocLastKey
                                    ' End If

                                    rptDocument.SetParameterValue("@DocEntry", PVNo)
                                    reportuti.SetDBLogIn(rptDocument)
                                    If Not pdffilepath = String.Empty Then
                                        reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
                                    End If
                                End If
                                If pVal.ItemUID = "bt_AmdVoc" Then
                                    'POP UP Payment Voucher
                                    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        If Not AlreadyExist("VOUCHER") Then
                                            LoadPaymentVoucher(ExportSeaFCLForm)
                                        Else
                                            p_oSBOApplication.Forms.Item("VOUCHER").Select()

                                        End If

                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Voucher.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        'Exit Function
                                    End If

                                End If
                                'Shipping Invoice POP Up
                                If pVal.ItemUID = "bt_ShpInv" Then
                                    'POP UP Shipping Invoice
                                    'If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If formUid = "EXPORTSEAFCL" Then
                                        If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And _
                                        (Not ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" Or Not ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.Value = "") And _
                                        Not ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value = "" And _
                                        Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            LoadShippingInvoice(ExportSeaFCLForm)
                                        Else
                                            If ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Customer Code to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" Or ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Shipping Agent Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            End If
                                            'Exit Function
                                        End If
                                    ElseIf formUid = "EXPORTAIRFCL" Then
                                        If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And _
                                        Not ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" And _
                                        Not ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value = "" And _
                                         Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            LoadShippingInvoice(ExportSeaFCLForm)
                                        Else
                                            If ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Customer Code to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Shipping Agent Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            End If
                                            'Exit Function
                                        End If
                                    End If


                                End If

                                'Purchase Order and Goods Receipt POP UP

                                '---------------------------------- Booking Preview-----------------------------'
                                If pVal.ItemUID = "bt_DrBL" Then
                                    'ExportSeaFCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                    'DeletePDF()
                                    'CopyFolder("DraftBL")
                                    pdfFilename = "DraftBL"
                                    originpdf = "DraftBL.pdf"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    PutDBValueToDBAndPreviewForDraftBL("DraftBL", ExportSeaFCLForm, pdffilepath, originpdf)
                                End If
                                'If pVal.ItemUID = "bt_DGD" Then
                                '    'ExportSeaFCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                '    'DeletePDF()
                                '    'CopyFolder("DGD")
                                '    pdfFilename = "DGD"
                                '    originpdf = "DGD-Air.pdf"
                                '    mainFolder = p_fmsSetting.DocuPath
                                '    jobNo = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                '    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                '    PutDBValueToDBAndPreviewForDGD("DGD", ExportSeaFCLForm, pdffilepath, originpdf)
                                'End If
                                If pVal.ItemUID = "bt_SpOrder" Then
                                    'ExportSeaFCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                    'DeletePDF()
                                    'CopyFolder("ShippingOrder")
                                    pdfFilename = "ShippingOrder"
                                    originpdf = "ShippingOrder.pdf"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    PutDBValueToDBAndPreviewForShippingOrder("ShippingOrder", ExportSeaFCLForm, pdffilepath, originpdf)
                                End If

                                If pVal.ItemUID = "bt_BokPO" Then 'sw
                                    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaFCLForm, "BokPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        'Exit Function
                                    End If
                                End If

                                If pVal.ItemUID = "bt_CPO" Then
                                    ' ==================================== Creating Custom Purchase Order ==============================
                                    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaFCLForm, "FumiPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If
                                    ' ==================================== Creating Custom Purchase Order ==============================
                                End If

                                If pVal.ItemUID = "bt_CrPO" Then 'sw
                                    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaFCLForm, "CratePurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If
                                End If

                                If pVal.ItemUID = "bt_BunkPO" Then 'sw
                                    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        'If Not CreatePOPDF(ImportSeaLCLForm, "mx_Bunk", "B00001") Then Throw New ArgumentException(sErrDesc)
                                        LoadAndCreateCPO(ExportSeaFCLForm, "BunkPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_ArmePO" Then 'sw
                                    If Not ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaFCLForm, "ArmedPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If
                                End If
                                '-----------------------------------------------------------------------------------------------------------'

                                '=====  Container View List & Edit ====='
                                If pVal.ItemUID = "ch_CStuff" Then
                                    Dim strTime As SAPbouiCOM.EditText
                                    If ExportSeaFCLForm.Items.Item("ch_CStuff").Specific.Checked = True Then
                                        ExportSeaFCLForm.Freeze(True)
                                        ExportSeaFCLForm.Items.Item("ed_CunDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        ExportSeaFCLForm.Items.Item("ed_CunDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                        'ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_TB020_HCONTAINE").SetValue("U_ContTime", ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_TB020_HCONTAINE").Offset, Now.ToString("HH:mm"))
                                        strTime = ExportSeaFCLForm.Items.Item("ed_CunTime").Specific
                                        strTime.Value = Now.ToString("HH:mm")
                                        'ExportSeaFCLForm.Items.Item("ed_CunTime").Specific.Value = Now.ToString("HH:mm")
                                        If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_CunDay").Specific, ExportSeaFCLForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        '  ExportSeaFCLForm.Items.Item("ed_CunTime").Specific.Active = True
                                        ExportSeaFCLForm.Freeze(False)
                                    ElseIf ExportSeaFCLForm.Items.Item("ch_CStuff").Specific.Checked = False Then
                                        ExportSeaFCLForm.Freeze(True)
                                        ExportSeaFCLForm.Items.Item("ed_CunDate").Specific.Value = vbNullString
                                        ExportSeaFCLForm.Items.Item("ed_CunDay").Specific.Value = vbNullString
                                        ExportSeaFCLForm.Items.Item("ed_CunTime").Specific.Value = vbNullString
                                        ' ExportSeaFCLForm.Items.Item("ed_CunTime").Specific.Active = True
                                        ExportSeaFCLForm.Freeze(False)
                                    End If
                                End If

                                If pVal.ItemUID = "bt_AddCont" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                    'If Convert.ToInt32(ExportSeaFCLForm.Items.Item("ed_ConNo").Specific.Value) > oMatrix.RowCount Then
                                    If ExportSeaFCLForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container" Then
                                        AddUpdateContainer(ExportSeaFCLForm, oMatrix, "@OBT_FCL19_HCONTAINE", True)
                                    Else
                                        AddUpdateContainer(ExportSeaFCLForm, oMatrix, "@OBT_FCL19_HCONTAINE", False)
                                        ExportSeaFCLForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container"
                                        ExportSeaFCLForm.Items.Item("bt_DelCont").Enabled = False 'MSW
                                    End If

                                    ClearText(ExportSeaFCLForm, "ed_ConNo", "ed_ContNo", "ed_SealNo", "ed_ContWt", "ed_CDesc", "ed_CunDate", "ed_CunDay", "ed_CunTime")
                                    Dim oComboType As SAPbouiCOM.ComboBox
                                    Dim oComboSize As SAPbouiCOM.ComboBox

                                    oComboType = ExportSeaFCLForm.Items.Item("cb_ConType").Specific
                                    oComboSize = ExportSeaFCLForm.Items.Item("cb_ConSize").Specific

                                    For j As Integer = oComboType.ValidValues.Count - 1 To 0 Step -1
                                        oComboType.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next

                                    oComboSize.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    ExportSeaFCLForm.Items.Item("fo_ConView").Specific.Select()
                                    ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = True
                                    UpdateNoofContainer(ExportSeaFCLForm, oMatrix)
                                End If

                                If pVal.ItemUID = "fo_ConEdit" Then
                                    Dim oComboType As SAPbouiCOM.ComboBox
                                    Dim oComboSize As SAPbouiCOM.ComboBox
                                    oComboType = ExportSeaFCLForm.Items.Item("cb_ConType").Specific
                                    oComboSize = ExportSeaFCLForm.Items.Item("cb_ConSize").Specific

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
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                    ' If ExportSeaFCLForm.Items.Item("ed_ConNo").Specific.Value = "" Then
                                    If (oMatrix.RowCount > 0) Then
                                        If (oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                                            ExportSeaFCLForm.Items.Item("ed_ConNo").Specific.Value = 1
                                        Else
                                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet.DoQuery("select count(*)  from [@OBT_FCL19_HCONTAINE] where U_ConNo=  '" + oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + "'")

                                            If oRecordSet.RecordCount > 0 And IsAmend = False Then
                                                ExportSeaFCLForm.Items.Item("ed_ConNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                            End If
                                        End If
                                    Else
                                        ExportSeaFCLForm.Items.Item("ed_ConNo").Specific.Value = 1
                                    End If
                                    'End If
                                    If ExportSeaFCLForm.Items.Item("ed_CunDate").Specific.Value = "" Then
                                        ExportSeaFCLForm.Items.Item("ch_CStuff").Specific.Checked = False
                                    End If
                                    ExportSeaFCLForm.Items.Item("ed_ContNo").Specific.Active = True
                                End If

                                If pVal.ItemUID = "fo_ConView" Then
                                    'oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                    'If oMatrix.RowCount > 1 Then
                                    '    ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = True
                                    'ElseIf oMatrix.RowCount = 1 Then
                                    '    If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    '        ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = True
                                    '    Else
                                    '        ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = False
                                    '    End If
                                    'End If
                                    'ClearText(ExportSeaFCLForm, "ed_ConNo", "ed_ContNo", "ed_SealNo", "ed_ContWt", "ed_CDesc", "ed_CunDate", "ed_CunDay", "ed_CunTime")
                                    'Dim oComboType As SAPbouiCOM.ComboBox
                                    'Dim oComboSize As SAPbouiCOM.ComboBox

                                    'oComboType = ExportSeaFCLForm.Items.Item("cb_ConType").Specific
                                    'oComboSize = ExportSeaFCLForm.Items.Item("cb_ConSize").Specific

                                    'For j As Integer = oComboType.ValidValues.Count - 1 To 0 Step -1
                                    '    oComboType.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    'Next
                                    'oComboSize.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    'ExportSeaFCLForm.Items.Item("bt_AddCont").Specific.Caption = "Add Container"
                                    'ExportSeaFCLForm.Items.Item("bt_DelCont").Enabled = False 'MSW
                                    'IsAmend = False
                                End If

                                If pVal.ItemUID = "bt_DelCont" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                    DeleteContainerByIndex(ExportSeaFCLForm, oMatrix, "@OBT_FCL19_HCONTAINE")
                                    ClearText(ExportSeaFCLForm, "ed_ConNo", "ed_ContNo", "ed_SealNo", "ed_ContWt", "ed_CDesc", "ed_CunDate", "ed_CunDay", "ed_CunTime")
                                    Dim oComboType As SAPbouiCOM.ComboBox
                                    Dim oComboSize As SAPbouiCOM.ComboBox

                                    oComboType = ExportSeaFCLForm.Items.Item("cb_ConType").Specific
                                    oComboSize = ExportSeaFCLForm.Items.Item("cb_ConSize").Specific

                                    For j As Integer = oComboType.ValidValues.Count - 1 To 0 Step -1
                                        oComboType.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
                                    Next

                                    oComboSize.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                    ExportSeaFCLForm.Items.Item("fo_ConView").Specific.Select()
                                    UpdateNoofContainer(ExportSeaFCLForm, oMatrix)
                                End If

                                If pVal.ItemUID = "bt_AmdCont" Then
                                    Dim oComboType As SAPbouiCOM.ComboBox
                                    Dim oComboSize As SAPbouiCOM.ComboBox
                                    oComboType = ExportSeaFCLForm.Items.Item("cb_ConType").Specific
                                    oComboSize = ExportSeaFCLForm.Items.Item("cb_ConSize").Specific
                                    ExportSeaFCLForm.Items.Item("bt_AddCont").Specific.Caption = "Update Container"
                                    ExportSeaFCLForm.Items.Item("bt_DelCont").Enabled = True 'MSW
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
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
                                    IsAmend = True
                                    ExportSeaFCLForm.Items.Item("fo_ConEdit").Specific.Select()
                                    GetContainerDataFromMatrixByIndex(ExportSeaFCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                    SetContainerDataToEditTabByIndex(ExportSeaFCLForm)
                                End If

                                If pVal.ItemUID = "mx_ConTab" And pVal.ColUID = "V_-1" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                                    If oMatrix.GetNextSelectedRow > 0 Then
                                        If (oMatrix.IsRowSelected(oMatrix.GetNextSelectedRow)) = True Then
                                            GetContainerDataFromMatrixByIndex(ExportSeaFCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        End If
                                    Else
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Container.", 1, "&OK")
                                    End If
                                End If

                                '===== End Container View List & Edit ====='

                                ''===== Other Charges Tab ====='
                                If pVal.ItemUID = "bt_AddCh" Then

                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific

                                    If ExportSeaFCLForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges" Then
                                        AddUpdateOtherCharges(ExportSeaFCLForm, oMatrix, "@OBT_FCL17_HCHARGES", True, 0)
                                    Else
                                        AddUpdateOtherCharges(ExportSeaFCLForm, oMatrix, "@OBT_FCL17_HCHARGES", False, oMatrix.GetNextSelectedRow)
                                        ExportSeaFCLForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges"
                                    End If
                                    ClearText(ExportSeaFCLForm, "ed_CSeqNo", "ed_ChCode", "ed_Remarks")
                                    ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = True
                                    ExportSeaFCLForm.Items.Item("fo_ChView").Specific.Select()

                                End If

                                If pVal.ItemUID = "bt_DelCh" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                    DeleteByIndexOtherCharges(ExportSeaFCLForm, oMatrix, "@OBT_FCL17_HCHARGES")
                                    ClearText(ExportSeaFCLForm, "ed_CSeqNo", "ed_ChCode", "ed_Remarks")
                                    ExportSeaFCLForm.Items.Item("fo_ChView").Specific.Select()
                                End If

                                If pVal.ItemUID = "fo_ChView" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                    If oMatrix.RowCount > 1 Then
                                        ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = True
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = True
                                        Else
                                            ExportSeaFCLForm.Items.Item("bt_AmdCh").Enabled = False
                                        End If
                                    End If
                                    ClearText(ExportSeaFCLForm, "ed_CSeqNo", "ed_ChCode", "ed_Remarks")

                                    ExportSeaFCLForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges"
                                    ExportSeaFCLForm.Items.Item("bt_DelCh").Enabled = False 'MSW
                                    IsAmend = False
                                End If

                                If pVal.ItemUID = "bt_AmdCh" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                    If (oMatrix.GetNextSelectedRow < 0) Then
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Other Charges", 1, "OK")
                                    Else
                                        IsAmend = True
                                        ExportSeaFCLForm.Items.Item("fo_ChEdit").Specific.Select()
                                        ExportSeaFCLForm.Items.Item("bt_AddCh").Specific.Caption = "Update Charges"
                                        ExportSeaFCLForm.Items.Item("bt_DelCh").Enabled = True
                                        ExportSeaFCLForm.Freeze(True)
                                        SetOtherChargesDataToEditTabByIndex(ExportSeaFCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        ExportSeaFCLForm.Freeze(False)
                                    End If

                                End If

                                ''===== End Other Charges Tab ===== '



                                If pVal.ItemUID = "mx_Cont" And ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If

                                If pVal.ItemUID = "bt_ChkList" Then
                                    Start(ExportSeaFCLForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                '-------------------------For Payment(omm)------------------------------------------'
                                If pVal.ItemUID = "mx_ChCode" And pVal.ColUID = "V_-1" Then
                                    oMatrix = ExportSeaFCLForm.Items.Item("mx_ChCode").Specific
                                    If pVal.Row > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            gridindex = CInt(pVal.Row)
                                        End If
                                    End If
                                End If
                                '----------------------------------------------------------------------------------'
                                'If pVal.ItemUID = "mx_TkrList" And pVal.ColUID = "V_1" Then
                                '    oMatrix = ExportSeaFCLForm.Items.Item("mx_TkrList").Specific
                                '    'If oMatrix.GetNextSelectedRow > 0 Then
                                '    If pVal.Row > 0 Then
                                '        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                '            modTrucking.rowIndex = CInt(pVal.Row)
                                '            modTrucking.GetDataFromMatrixByIndex(ExportSeaFCLForm, oMatrix, modTrucking.rowIndex)
                                '        End If
                                '    End If
                                '    'Else
                                '    '   p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction.", 1, "&OK")
                                '    ' End If

                                'End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                                Try
                                    '-------------------------For Payment(omm)------------------------------------------'
                                    'New UI
                                    'Dispatch Setting
                                    If pVal.ItemUID = "ed_DICode" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_TB01_DSPPOSET").SetValue("U_ICode", 0, oDataTable.GetValue(0, 0).ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_TB01_DSPPOSET").SetValue("U_IDesc", 0, oDataTable.GetValue(1, 0).ToString)
                                    End If

                                    'Trucking Setting

                                    If pVal.ItemUID = "ed_POICode" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_TB01_TRKPOSET").SetValue("U_ICode", 0, oDataTable.GetValue(0, 0).ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_TB01_TRKPOSET").SetValue("U_IDesc", 0, oDataTable.GetValue(1, 0).ToString)
                                    End If

                                    If pVal.ColUID = "colCCode1" Then
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                                        Try
                                            oMatrix.Columns.Item("colCCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                        Catch ex As Exception
                                        End Try
                                        Try
                                            oMatrix.Columns.Item("colCDesc1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(1, 0).ToString
                                        Catch ex As Exception
                                        End Try

                                        If oMatrix.Columns.Item("colCCode1").Cells.Item(pVal.Row).Specific.Value <> "" And oMatrix.Columns.Item("colCCode1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL17_HCHARGES").Clear()
                                            oMatrix.AddRow(1)
                                            oMatrix.FlushToDataSource()
                                            oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                                            oCombo = oMatrix.Columns.Item("colCClaim1").Cells.Item(oMatrix.RowCount).Specific
                                            oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                        End If

                                    End If

                                    If pVal.ItemUID = "ed_VedName" Then
                                        ObjDBDataSource = ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL05_EVOUCHER")                 '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL05_EVOUCHER").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL05_EVOUCHER").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                                                                                               & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString())
                                        vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft                                        '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    End If
                                    If pVal.ColUID = "colChCode" Then
                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_ChCode").Specific
                                        dtmatrix = ExportSeaFCLForm.DataSources.DataTables.Item("DTCharges")
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
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString) '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_SaleCode", 0, oDataTable.GetValue(0, 0).ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString

                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_ID", 0, oDataTable.Columns.Item("CardFName").Cells.Item(0).Value.ToString)
                                        'If String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific.Value) Then
                                        '    ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                        '    ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific.Value = Now.ToString("HH:mm")
                                        '    If HolidayMarkUpWithoutDay(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        'End If
                                        If String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_ADate").Specific.Value) Then
                                            ExportSeaFCLForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            ExportSeaFCLForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                            ExportSeaFCLForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                            If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_ADay").Specific, ExportSeaFCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        EnabledHeaderControls(ExportSeaFCLForm, True)
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                            ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                                            ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.ChooseFromListUID = "cflBP4"
                                            ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.ChooseFromListAlias = "CardName"



                                            'ExportSeaFCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                            'ExportSeaFCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                        End If
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_IComName", 0, oDataTable.GetValue(1, 0).ToString)   '* Change Nyan Lin   "[@OBT_LCL06_PMAI]"
                                        End If

                                        oRecordSet.DoQuery("select OSLP.SlpName as SlpName from OCRD left join OSLP on OCRD.SlpCode =OSLP.SlpCode  WHERE OCRD.CardCode = '" & ExportSeaFCLForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ExportSeaFCLForm.Items.Item("ed_Sales").Specific.Value = oRecordSet.Fields.Item("SlpName").Value.ToString
                                        End If
                                    End If
                                    'Google Doc
                                    If pVal.ItemUID = "ed_C2Code" Or pVal.ItemUID = "ed_Client2" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_C2Code", 0, oDataTable.GetValue(0, 0).ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Name2", 0, oDataTable.GetValue(1, 0).ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_C2Ctry", 0, oDataTable.Columns.Item("CardFName").Cells.Item(0).Value.ToString)
                                    End If
                                    If pVal.ItemUID = "ed_ShpAgt" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)     '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)      '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)            '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString) '* Change Nyan Lin   "[@OBT_TB005_EPERMIT]"
                                        End If
                                    End If
                                    If pVal.ItemUID = "ed_IShpAgt" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_IMShpAgt", 0, oDataTable.GetValue(1, 0).ToString)     '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_IVCode", 0, oDataTable.GetValue(0, 0).ToString)      '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)            '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("SELECT CardName FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString) '* Change Nyan Lin   "[@OBT_TB005_EPERMIT]"
                                        End If
                                    End If


                                    If pVal.ItemUID = "ed_ChCode" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL17_HCHARGES").SetValue("U_CCode", ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL17_HCHARGES").Offset, oDataTable.Columns.Item("U_CName").Cells.Item(0).Value.ToString)
                                    End If

                                    If pVal.ItemUID = "ed_Yard" Then
                                        ExportSeaFCLForm.Items.Item("fo_Yard").Specific.Select()
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_YName", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YName", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YTel", 0, oDataTable.Columns.Item("U_YTel").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YCPrson", 0, oDataTable.Columns.Item("U_YPerson").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YMobile", 0, oDataTable.Columns.Item("U_YMobile").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YHr", 0, oDataTable.Columns.Item("U_YwhL1").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YLat", 0, oDataTable.Columns.Item("U_YLat").Cells.Item(0).Value.ToString)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YLong", 0, oDataTable.Columns.Item("U_YLong").Cells.Item(0).Value.ToString)
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim sql1 As String = "SELECT U_YAdLine1,U_YAdLine2,U_YAdLine3,U_YState,U_YPostal,U_YCountry FROM [@OBT_TB022_RYARDLIST] WHERE Name = " & FormatString(oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                        Dim WAddress As String = String.Empty
                                        oRecordSet.DoQuery(sql1)
                                        If oRecordSet.RecordCount > 0 Then
                                            WAddress = Trim(oRecordSet.Fields.Item("U_YAdLine1").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_YAdLine2").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_YAdLine3").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_YState").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_YPostal").Value) & Chr(13) & _
                                                            Trim(oRecordSet.Fields.Item("U_YCountry").Value)
                                            ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL18_RYARDTAB").SetValue("U_YAddr", 0, WAddress)
                                            ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = WAddress
                                        End If
                                    End If

                                    If pVal.ItemUID = "ed_Trucker" Then
                                        Try
                                            If ExportSeaFCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                                'ExportSeaFCLForm.Freeze(True)
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                                oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = "" '25-3-2011
                                                'ExportSeaFCLForm.Freeze(False)
                                            ElseIf ExportSeaFCLForm.Items.Item("op_Exter").Specific.Selected = True Then
                                                ExportSeaFCLForm.Items.Item("ed_TkrCode").Specific.Value = oDataTable.GetValue(0, 0).ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKREXTR").ValueEx = oDataTable.Columns.Item("CardName").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = oDataTable.Columns.Item("CntctPrsn").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("Phone1").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("Fax").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("E_Mail").Cells.Item(0).Value.ToString
                                            End If
                                        Catch ex As Exception

                                        End Try

                                    End If
                                    'Dispatch Tab
                                    If pVal.ItemUID = "ed_Dspatch" Then
                                        Try
                                            If ExportSeaFCLForm.Items.Item("op_DInter").Specific.Selected = True Then
                                                'ExportSeaFCLForm.Freeze(True)
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                                oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPATTE").ValueEx = "" '25-3-2011
                                                'ExportSeaFCLForm.Freeze(False)
                                            ElseIf ExportSeaFCLForm.Items.Item("op_DExter").Specific.Selected = True Then
                                                ExportSeaFCLForm.Items.Item("ed_DspCode").Specific.Value = oDataTable.GetValue(0, 0).ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPEXTR").ValueEx = oDataTable.Columns.Item("CardName").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPATTE").ValueEx = oDataTable.Columns.Item("CntctPrsn").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPTEL").ValueEx = oDataTable.Columns.Item("Phone1").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPFAX").ValueEx = oDataTable.Columns.Item("Fax").Cells.Item(0).Value.ToString
                                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DSPMAIL").ValueEx = oDataTable.Columns.Item("E_Mail").Cells.Item(0).Value.ToString
                                            End If
                                        Catch ex As Exception

                                        End Try

                                    End If

                                    If pVal.ItemUID = "ed_Vessel" Or pVal.ItemUID = "ed_Voy" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)        '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        'ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)    '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_VName", 0, ExportSeaFCLForm.Items.Item("ed_Vessel").Specific.String)                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    End If

                                    If pVal.ItemUID = "ed_CurCode" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_CurCode", 0, oDataTable.GetValue(0, 0).ToString)
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_ExRate", 0, Rate.ToString)
                                    End If
                                    If pVal.ItemUID = "ed_CCharge" Then
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_Cchange", 0, oDataTable.GetValue(0, 0).ToString)                            '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                        Dim Rate As String = String.Empty
                                        SqlQuery = "SELECT Rate FROM ORTT WHERE Currency = '" & oDataTable.GetValue(0, 0).ToString & "' And DATENAME(YYYY,RateDate) = '" & _
                                                Today.ToString("yyyy") & "' And DATENAME(MM,RateDate) = '" & Today.ToString("MMMM") & "' And DATENAME(DD,RateDate) = " & _
                                                CInt(Today.ToString("dd"))
                                        oRecordSet.DoQuery(SqlQuery)
                                        If oRecordSet.RecordCount > 0 Then
                                            Rate = oRecordSet.Fields.Item("Rate").Value
                                        End If
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_CEchange", 0, Rate.ToString)                                               '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                    If pVal.ItemUID = "ed_Charge" Then
                                        ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_FCchange", 0, oDataTable.GetValue(0, 0).ToString)                          '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    End If
                                Catch ex As Exception
                                End Try

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                If pVal.ItemUID = "cb_JbStus" Then 'Check PO
                                    If ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim <> "Open" Then
                                        If ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Closed" And ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = False Then
                                            p_oSBOApplication.MessageBox("Need to check PO check box first.")
                                            ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                            BubbleEvent = False
                                        ElseIf ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim = "Closed" And ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = True Then
                                            BubbleEvent = False
                                        Else
                                            If CheckPOandVoucherStatus(ExportSeaFCLForm) = True Then
                                                p_oSBOApplication.MessageBox("Need To Closed Open PO First.")
                                                ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = False
                                                ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                            Else
                                                Dim chk As Integer = 0
                                                chk = p_oSBOApplication.MessageBox("You cannot change this Job after you have " & ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim & " it.Continue?", 1, "Yes", "No")
                                                If chk = 1 Then
                                                    ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = True
                                                    ExportSeaFCLForm.Items.Item("ed_xRef").Specific.Active = True
                                                    ' ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value.ToString.Trim, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                                    ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False

                                                    ExportSeaFCLForm.Items.Item("cb_JbStus").Enabled = False
                                                Else
                                                    ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = False
                                                    ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                                                    ExportSeaFCLForm.Items.Item("ch_POD").Enabled = True
                                                End If

                                            End If
                                        End If
                                    End If

                                End If
                                If pVal.ItemUID = "cb_PCode" Then
                                    oCombo = ExportSeaFCLForm.Items.Item("cb_PCode").Specific
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                             '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_PCode", 0, oCombo.Selected.Value.ToString)                                       '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                End If
                                If pVal.ItemUID = "cb_BnkName" Then
                                    oCombo = ExportSeaFCLForm.Items.Item("cb_BnkName").Specific
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim test As String = "select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description
                                    oRecordSet.DoQuery("select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description)
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL05_EVOUCHER").SetValue("U_GLAC", 0, oRecordSet.Fields.Item("GLAccount").Value)                            '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                    'oCombo.Selected.Description
                                End If

                                If pVal.ItemUID = "cb_PType" Then
                                    oCombo = ExportSeaFCLForm.Items.Item("cb_PType").Specific
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_TUnit", 0, oCombo.Selected.Value.ToString)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN
                                If pVal.ItemUID = "ed_JobNo" And pVal.CharPressed = 13 Then
                                    ExportSeaFCLForm.Items.Item("1").Click()
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                'MSW 
                                If pVal.ItemUID = "ed_TkrDate" And pVal.Before_Action = False Then
                                    Dim strTime As SAPbouiCOM.EditText
                                    strTime = ExportSeaFCLForm.Items.Item("ed_TkrTime").Specific
                                    strTime.Value = Now.ToString("HH:mm")
                                End If
                                'End MSW
                                If pVal.ItemUID = "ed_InvNo" Then
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_InvNo", 0, ExportSeaFCLForm.Items.Item("ed_InvNo").Specific.String)         '* Change Nyan Lin   "[@OBT_TB0011_VOUCHER]"
                                End If
                                'NL LCL Change 24-03-2011
                                If pVal.ItemUID = "ed_NOP" Then
                                    ExportSeaFCLForm.DataSources.DBDataSources.Item("@OBT_FCL02_EPERMIT").SetValue("U_TotalOP", 0, ExportSeaFCLForm.Items.Item("ed_NOP").Specific.String)
                                End If
                                'End NL LCL Change 24-03-2011
                                'If ExportSeaFCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Or ExportSeaFCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                '    Validateforform(pVal.ItemUID, ExportSeaFCLForm)
                                'End If
                                If BubbleEvent = False Then
                                    Validateforform(pVal.ItemUID, ExportSeaFCLForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                        End Select
                    End If
                    If pVal.BeforeAction = True Then
                        Select Case pVal.EventType
                            'MSW for Job Type Table
                            Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                                If Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If pVal.ItemUID = "ed_JobNo" And pVal.InnerEvent = False Then
                                        ValidateJobNumber(ExportSeaFCLForm, BubbleEvent)
                                    End If
                                End If

                                'End MSW for Job Type Table
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                If pVal.ItemUID = "1" Then
                                    Dim PODFlag As String = String.Empty
                                    Dim JbStus As String = String.Empty
                                    Dim DispatchComplete As String = String.Empty
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("SELECT U_JbStus,U_POD FROM [@OBT_FCL01_EXPORT] WHERE DocEntry = " & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value) 'MSW 08-06-2011 for job Type Table
                                    If oRecordSet.RecordCount > 0 Then
                                        JbStus = oRecordSet.Fields.Item("U_JbStus").Value
                                        PODFlag = oRecordSet.Fields.Item("U_POD").Value
                                    End If
                                    If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        'develivery process by POD[Proof Of Delivery] check box
                                        
                                        'If ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = True And JbStus = "Open" Then
                                        '    If p_oSBOApplication.MessageBox("Make sure that all entries trucking and vouchers are completed.(ensure no draft Payment in this job and " & _
                                        '                               "ensure all external trucking transaction has generated the PO). Cannot edit or add after click POD check box. " & _
                                        '                               "Do you want to continue?", 1, "&Yes", "&No") = 2 Then
                                        '        BubbleEvent = False
                                        '    End If
                                        'End If
                                        If BubbleEvent = True Then
                                            'MSW 08-06-2011 for job type table
                                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ExportSeaFCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            'End MSW 08-06-2011 for job type table
                                        End If
                                    End If
                                    If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        
                                        If ExportSeaFCLForm.Items.Item("ch_POD").Specific.Checked = True And JbStus = "Closed" Then
                                            p_oSBOApplication.MessageBox("System cannot update your data because job is already closed.")
                                            BubbleEvent = False
                                        ElseIf JbStus = "Cancelled" Then
                                            p_oSBOApplication.MessageBox("System cannot update your data because job is already Cancelled.")
                                            BubbleEvent = False
                                        End If

                                    End If
                                   
                                    If ExportSeaFCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And ExportSeaFCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If Validateforform(" ", ExportSeaFCLForm) Then
                                            BubbleEvent = False
                                        End If
                                    End If
                                End If
                        End Select
                    End If

                Case "MULTIJOBN"
                    MultiJobForm = p_oSBOApplication.Forms.Item(pVal.FormUID)

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then

                        If pVal.ItemUID = "ed_frmDate" And MultiJobForm.Items.Item("ed_frmDate").Specific.String <> String.Empty Then
                            Try
                                If Not modMultiJobForNormal.DateTimeMJ(MultiJobForm, MultiJobForm.Items.Item("ed_frmDate").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(MultiJobForm, MultiJobForm.Items.Item("ed_frmDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            Catch ex As Exception

                            End Try
                        End If
                        If pVal.ItemUID = "ed_toDate" And MultiJobForm.Items.Item("ed_toDate").Specific.String <> String.Empty Then
                            Try
                                If Not modMultiJobForNormal.DateTimeMJ(MultiJobForm, MultiJobForm.Items.Item("ed_toDate").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(MultiJobForm, MultiJobForm.Items.Item("ed_toDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            Catch ex As Exception

                            End Try
                        End If

                        If pVal.ItemUID = "bt_Search" Then
                            modMultiJobForNormal.Bind_MxSelect()
                        End If
                        If pVal.ItemUID = "bt_Choose" Then
                            modMultiJobForNormal.Bind_MxAdd()
                        End If

                        If pVal.ItemUID = "bt_Add" Then
                            oMatrix = MultiJobForm.Items.Item("mx_Add").Specific
                            ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                            If MultiJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Then
                                ExportSeaFCLForm.DataSources.UserDataSources.Item("TMULTI").ValueEx = "Y"
                                modMultiJobForNormal.UpdateRemarkForMultiJob(ExportSeaFCLForm, oMatrix, "ee_Rmsk")
                                modMultiJobForNormal.UpdateRemarkForPOTable(ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value, ExportSeaFCLForm.Items.Item("ee_Rmsk").Specific.Value)
                                modMultiJobForNormal.UpdateB1RemarkForMultiPO(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ee_Rmsk").Specific.Value, "ed_PONo")
                                modMultiJobForNormal.SavetoMultiPOTable(ExportSeaFCLForm, oMatrix, "ed_PO", "ed_PODocNo")

                                modMultiJobForNormal.LineAddtoTable(ExportSeaFCLForm, oMatrix, MultiJobForm.Items.Item("ed_Source").Specific.Value)
                                ExportSeaFCLForm.Items.Item("bt_TkrAdd").Click()
                            ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
                                ExportSeaFCLForm.DataSources.UserDataSources.Item("DMULTI").ValueEx = "Y"
                                modMultiJobForNormal.UpdateRemarkForMultiJob(ExportSeaFCLForm, oMatrix, "ee_DRmsk")
                                modMultiJobForNormal.UpdateRemarkForPOTable(ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value, ExportSeaFCLForm.Items.Item("ee_DRmsk").Specific.Value)
                                modMultiJobForNormal.UpdateB1RemarkForMultiPO(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ee_DRmsk").Specific.Value, "ed_DPONo")
                                modMultiJobForNormal.SavetoMultiPOTable(ExportSeaFCLForm, oMatrix, "ed_DPO", "ed_DPDocNo")

                                modMultiJobForNormal.LineAddtoTable(ExportSeaFCLForm, oMatrix, MultiJobForm.Items.Item("ed_Source").Specific.Value)
                                ExportSeaFCLForm.Items.Item("bt_DspAdd").Click()
                            Else  'Fumigation
                                If MultiJobForm.Items.Item("ed_Source").Specific.Value = "Fumigation" Then
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Crane" Then
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Outrider" Then
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Forklift" Then 'to combine
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Crate" Then 'to combine
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Bunker" Then
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                ElseIf MultiJobForm.Items.Item("ed_Source").Specific.Value = "Toll" Then
                                    FumigationForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                End If
                                FumigationForm.DataSources.UserDataSources.Item("MultiJob").ValueEx = "Y"
                                modMultiJobForNormal.UpdateRemarkForMultiJob(FumigationForm, oMatrix, "ed_Remark")
                                modMultiJobForNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modMultiJobForNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                modMultiJobForNormal.SavetoMultiPOTable(FumigationForm, oMatrix, "ed_PO", "ed_PODocNo")
                                modMultiJobForNormal.LineAddtoTable(FumigationForm, oMatrix, MultiJobForm.Items.Item("ed_Source").Specific.Value)
                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                FumigationForm.Items.Item("bt_DJob").Enabled = True
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            End If

                            MultiJobForm.Items.Item("2").Click()
                        End If
                    End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                        Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                        Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                        If pVal.ItemUID = "ed_Cust" Then
                            MultiJobForm.DataSources.UserDataSources.Item("CUST").ValueEx() = oDataTable.GetValue(1, 0).ToString
                            '  MultiJobForm.DataSources.DBDataSources.Item("@OBT_TB42_DGLPBP").SetValue("U_DGLPName", 0, oDataTable.Columns.Item("U_DGLPName").Cells.Item(0).Value.ToString)

                        End If
                    End If


                Case "DETACHJOBN"
                    DetachJobForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_Detach" Then
                            oMatrix = DetachJobForm.Items.Item("mx_Select").Specific
                            If DetachJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Then
                                modDetachJobNormal.DetachMultiJob(ExportSeaFCLForm, "ee_Rmsk", "ed_TMulti")
                                modDetachJobNormal.UpdateRemarkForPOTable(ExportSeaFCLForm.Items.Item("ed_PONo").Specific.Value, ExportSeaFCLForm.Items.Item("ee_Rmsk").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ee_Rmsk").Specific.Value, "ed_PONo")
                                ExportSeaFCLForm.Items.Item("bt_TkrAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
                                modDetachJobNormal.DetachMultiJob(ExportSeaFCLForm, "ee_DRmsk", "ed_DMulti")
                                modDetachJobNormal.UpdateRemarkForPOTable(ExportSeaFCLForm.Items.Item("ed_DPONo").Specific.Value, ExportSeaFCLForm.Items.Item("ee_DRmsk").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ee_DRmsk").Specific.Value, "ed_DPONo")
                                ExportSeaFCLForm.Items.Item("bt_DspAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Fumigation" Then
                                FumigationForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Outrider" Then '15/12/2011
                                FumigationForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Crane" Then 'Button
                                FumigationForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Forklift" Then 'to combine
                                FumigationForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Crate" Then 'to combine
                                FumigationForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Bunker" Then 'Button
                                FumigationForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            ElseIf DetachJobForm.Items.Item("ed_Source").Specific.Value = "Toll" Then 'Button
                                FumigationForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                modDetachJobNormal.DetachMultiJob(FumigationForm, "ed_Remark", "ed_MultiJb")
                                modDetachJobNormal.UpdateRemarkForPOTable(FumigationForm.Items.Item("ed_PONo").Specific.Value, FumigationForm.Items.Item("ed_Remark").Specific.Value)
                                modDetachJobNormal.UpdateB1RemarkForMultiPO(FumigationForm, FumigationForm.Items.Item("ed_Remark").Specific.Value, "ed_PONo")
                                FumigationForm.Items.Item("bt_FumiAdd").Click()
                            End If
                            DetachJobForm.Items.Item("2").Click()
                        End If
                    End If
                   
                    'MSW To Add New Button at ChooseFromList
                    'When User Click New button at ChooseFromList Form ,Setup Form Show to fill data.
                Case "9999"
                    ExportSeaFCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_VNew" Then
                            ExportSeaFCLForm.Close()
                            p_oSBOApplication.ActivateMenuItem("47644")
                        End If
                        If pVal.ItemUID = "bt_VoNew" Then
                            ExportSeaFCLForm.Close()
                            p_oSBOApplication.ActivateMenuItem("47644")
                        End If
                    End If
                    'MSW To Add New Button at ChooseFromList
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoExportSeaFCLItemEvent = RTN_SUCCESS

        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaFCLItemEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function

    Private Function PDFAlert(ByVal oActiveForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   PDFAlert()
        '   Purpose     :   This function will be providing to show PDF Alert Function   for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByVal oActiveForm As SAPbouiCOM.Form
        '   Return      :   False- FAILURE
        '               :   True - SUCCESS
        ' **********************************************************************************


        PDFAlert = False
        Dim originpdf As String = String.Empty
        Dim PDFDate As String = String.Empty
        Try
            mainFolder = p_fmsSetting.DocuPath
            pdfFilename = "PreAlert"
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            originpdf = "Alert - PDF.pdf"
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            'Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(IO.Directory.GetParent(Application.StartupPath).ToString & "\" & originpdf)
            Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            Dim pdfoutputfile As FileStream = New FileStream(pdffilepath, System.IO.FileMode.Create)
            Dim formfiller As iTextSharp.text.pdf.PdfStamper = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            Dim ac As iTextSharp.text.pdf.AcroFields = formfiller.AcroFields
            PDFDate = oActiveForm.Items.Item("ed_JbDate").Specific.value.ToString()
            ac.SetField("txtDate", PDFDate.Substring(6, 2).ToString() & "." & PDFDate.Substring(4, 2).ToString() & "." & PDFDate.Substring(0, 4).ToString())
            ac.SetField("txtjobFileNo", oActiveForm.Items.Item("ed_JobNo").Specific.value)
            ac.SetField("txtCargo", oActiveForm.Items.Item("ed_CrgDsc").Specific.value)
            ac.SetField("txtShpInvoice", ShpInvoice)
            ac.SetField("txtPO#", PO)
            ac.SetField("txtRoute", "")
            ac.SetField("txtDetails", "")
            ac.SetField("txtShipTo", ShipTo)
            ac.SetField("txtAWB", "") 'oActiveForm.Items.Item("ed_AWBNo").Specific.value)
            ac.SetField("txtBox", Box)
            ac.SetField("txtWeight", Weight)
            ac.SetField("txtPOD", "")
            ac.SetField("txtPODClearing Agent", "")
            ac.SetField("txtFrom", "")
            ac.SetField("txtTo", "")
            ac.SetField("txtRouting1", "")
            ac.SetField("txtRouting2", "")
            ac.SetField("txtRoutingDate", "")
            ac.SetField("txtAirPortLegand1", "")
            ac.SetField("txtAirPortLegand2", "")
            ac.SetField("txtAirPortLegand3", "")
            ac.SetField("txtAirPortLegand4", "")
            ac.SetField("txtEmail1", "")
            ac.SetField("txtEmail2", "")
            formfiller.Close()
            reader.Close()
            Process.Start(pdffilepath)
            PDFAlert = True
        Catch ex As Exception
            PDFAlert = False
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Public Function DoExportSeaFCLMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
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
        Dim ExportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oComboPO As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing
        Dim FunctionName As String = "DoExportSeaFCLMenuEvent()"
        Dim oMatrixName As String = ""
        Dim TableName As String = ""


        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Select Case pVal.MenuUID
                Case "mnuExportSeaFCL"
                    If pVal.BeforeAction = False Then
                        LoadExportSeaFCLForm()
                    End If

                Case "1281"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx).Items.Item("ed_JobNo").Enabled = True
                            p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx).Items.Item("ed_JobNo").Specific.Active = True
                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False

                            End If
                            ' If AddChooseFromListByOption(p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx), True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            If AddChooseFromListByOption(ExportSeaFCLForm, True, "ed_Trucker", "TKRINTR", "CFLTKRE", "TKREXTR", "CFLTKRV", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                        End If
                    End If
                Case "1292"
                    If pVal.BeforeAction = False Then
                        'Export Voucher POP UP
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                            If oMatrix.Columns.Item("colChCode1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                            End If
                            'Export Voucher POP UP
                        ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            If ActiveMatrix = "mx_ConTab" Or ActiveMatrix = "mx_Charge" Then
                                oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                If ActiveMatrix = "mx_ConTab" Then
                                    AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCNo1", "@OBT_FCL19_HCONTAINE")
                                    oCombo = oMatrix.Columns.Item("colCSize1").Cells.Item(oMatrix.RowCount).Specific
                                    oCombo.Select("20'", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                ElseIf ActiveMatrix = "mx_Charge" Then
                                    AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCCode1", "@OBT_FCL17_HCHARGES")
                                    oCombo = oMatrix.Columns.Item("colCClaim1").Cells.Item(oMatrix.RowCount).Specific
                                    oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                End If

                            ElseIf pVal.BeforeAction = True Then
                                BubbleEvent = False
                            End If
                            ''-------------------------For Payment(omm)------------------------------------------'
                            ''If (ImportSeaLCLForm.PaneLevel = 21) Then
                            'oMatrix = ExportSeaFCLForm.Items.Item("mx_ChCode").Specific
                            'RowAddToMatrix(ExportSeaFCLForm, oMatrix)
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
                     p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000010" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000015" Or _
                       p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000032" Or _
                     p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000035" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000038" Or _
                       p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000041" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000043" Or _
                     p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000026" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000030" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000050" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000052" Then

                        ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                        'If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oMatrix = ExportSeaFCLForm.Items.Item("mx_Item").Specific
                        ' Dim lRow As Long
                        If pVal.BeforeAction = True Then
                            If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                BubbleEvent = False
                            End If

                            If BubbleEvent = True Then
                                DeleteMatrixRow(ExportSeaFCLForm, oMatrix, "@OBT_TB09_FFCPOITEM", "LineId")
                                BubbleEvent = False
                                If Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                                CalculateTotalPO(ExportSeaFCLForm, oMatrix)
                            End If
                        End If

                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                        If pVal.BeforeAction = True Then
                            ' BubbleEvent = False
                            'MSW To Edit
                            ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                            If ActiveMatrix = "mx_ShpInv" Then
                                Dim shpDocEntry As String
                                oMatrix = ExportSeaFCLForm.Items.Item("mx_ShpInv").Specific
                                If oMatrix.GetNextSelectedRow > 0 Then
                                    If oMatrix.Columns.Item("colDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString <> "" Then
                                        shpDocEntry = oMatrix.Columns.Item("colSDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString
                                        DeleteMatrixRow(ExportSeaFCLForm, oMatrix, "@OBT_FCL07_SHPINV", "V_-1")
                                        DeleteUDO(ExportSeaFCLForm, "SHIPPINGINV", shpDocEntry)
                                        ExportSeaFCLForm.Items.Item("1").Click()
                                        If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                                    End If
                                End If
                            ElseIf ActiveMatrix = "mx_ConTab" Or ActiveMatrix = "mx_Charge" Then
                                oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                    BubbleEvent = False
                                End If

                                If BubbleEvent = True Then
                                    If ActiveMatrix = "mx_ConTab" Then
                                        DeleteMatrixRow(ExportSeaFCLForm, oMatrix, "@OBT_FCL19_HCONTAINE", "V_-1")
                                        UpdateNoofContainer(ExportSeaFCLForm, oMatrix)
                                        BubbleEvent = False
                                    ElseIf ActiveMatrix = "mx_Charge" Then
                                        DeleteMatrixRow(ExportSeaFCLForm, oMatrix, "@OBT_FCL17_HCHARGES", "V_-1")
                                        BubbleEvent = False
                                    End If
                                    If Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                            End If
                            

                            'End MSW To Edit
                        End If

                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Or _
                         p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000007" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000011" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000012" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000013" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000014" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000016" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000044" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000051" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000053" Then
                        If pVal.BeforeAction = True Then
                            BubbleEvent = False
                        End If
                    End If

                Case "1282"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            EnabledHeaderControls(ExportSeaFCLForm, False) '25-3-2011
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'EXPORTSEAFCL'")
                            If oRecordSet.RecordCount > 0 Then
                                ' ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                                ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                            End If
                            ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("EX")
                            If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Then
                                ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value = "Export Sea FCL" 'MSW 08-06-2011 for Job Type Table
                            ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                                ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value = "Export Air FCL" 'MSW 08-06-2011 for Job Type Table
                            End If
                            ExportSeaFCLForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString
                            ExportSeaFCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            ExportSeaFCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                            ExportSeaFCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                            If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_JbDay").Specific, ExportSeaFCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ExportSeaFCLForm.Items.Item("ch_POD").Enabled = False
                        End If
                    End If

                Case "1288", "1289", "1290", "1291"
                    If pVal.BeforeAction = False Then
                        ' If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If

                            Dim JobType As String = ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value
                            'Google Doc
                            If JobType.Contains("Import") Then
                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, False, True)
                            ElseIf JobType.Contains("Export") Then
                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, True, False)
                            ElseIf JobType.Contains("Local") Then
                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, True, True)
                            ElseIf JobType.Contains("Transhipment") Then
                                ExportSeaFCLForm.Title = JobType + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                                EnableExportImport(ExportSeaFCLForm, True, True)
                            End If
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                            AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCNo1", "@OBT_FCL19_HCONTAINE")
                            oCombo = oMatrix.Columns.Item("colCSize1").Cells.Item(oMatrix.RowCount).Specific
                            oCombo.Select("20'", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                            AddTabMatrixRow(ExportSeaFCLForm, oMatrix, "V_-1", "colCCode1", "@OBT_FCL17_HCHARGES")
                            oCombo = oMatrix.Columns.Item("colCClaim1").Cells.Item(oMatrix.RowCount).Specific
                            oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        End If

                    End If



                Case "EditVoc"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                            ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(p_oSBOApplication.Forms.ActiveForm.TypeEx, 1)
                            LoadPaymentVoucher(ExportSeaFCLForm)
                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ExportSeaFCLForm.Items.Item("mx_Voucher").Specific
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
                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                            'oPayForm.Items.Item("ed_PayTo").Enabled = False
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
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                            Try
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm(p_oSBOApplication.Forms.ActiveForm.TypeEx, 1)
                                LoadShippingInvoice(ExportSeaFCLForm)
                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                                oMatrix = ExportSeaFCLForm.Items.Item("mx_ShpInv").Specific
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

                            Catch ex As Exception
                                MessageBox.Show(ex.ToString())
                            End Try

                        End If

                        'End If
                    End If
                Case "EditCPO"
                    If pVal.BeforeAction = False Then
                        Select Case ActiveMatrix
                            Case "mx_Bunk"
                                RPOsrfname = "BunkPurchaseOrder.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                            Case "mx_Toll"
                                RPOsrfname = "TollPurchaseOrder.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                            Case "mx_Crane"
                                RPOsrfname = "CranePurchaseOrder_Form.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                            Case "mx_Armed"
                                RPOsrfname = "ArmedPurchaseOrder.srf"
                            Case "mx_Fumi"
                                RPOsrfname = "FumiPurchaseOrder.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                            Case "mx_Courier"
                                RPOsrfname = "CourierPurchaseOrder.srf"
                            Case "mx_DGLP"
                                RPOsrfname = "DGLPPurchaseOrder.srf"
                            Case "mx_Fork"
                                RPOsrfname = "ForkPurchaseOrder.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                            Case "mx_Crate"  'to combine
                                RPOsrfname = "CratePurchaseOrder.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                            Case "mx_Outer"
                                RPOsrfname = "OutPurchaseOrder.srf"
                                ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                            Case "mx_COO"
                                RPOsrfname = "CertificateOfOriginPPurchaseOrder.srf"
                        End Select

                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000025" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000026" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000032" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000035" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000038" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000028" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000031" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000034" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000037" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000040" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "FUMIGATION" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "OUTRIDER" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "CRANE" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "BUNKER" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "TOLL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "FORKLIFT" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "CRATE" Then           'Purchase
                            If currentRow > 0 Then
                                oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadAndCreate_PurChaseOrder(ExportSeaFCLForm, RPOsrfname)
                                    CPOForm = p_oSBOApplication.Forms.ActiveForm
                                    CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                                    CPOForm.Items.Item("1").Click()

                                    oComboPO = CPOForm.Items.Item("cb_Contact").Specific
                                    If Not ClearComboData(CPOForm, "cb_Contact", "@OBT_TB08_FFCPO", "U_CPerson") Then Throw New ArgumentException(sErrDesc)
                                    oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        oRecordSet.MoveFirst()
                                        While oRecordSet.EoF = False
                                            oComboPO.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                            oRecordSet.MoveNext()

                                        End While
                                    End If
                                    ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                End If
                            Else
                                p_oSBOApplication.MessageBox("Need to select the Row that you want to Edit")
                            End If
                        ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRFCL" Then
                            If currentRow > 0 Then
                                If AlreadyExist("EXPORTSEAFCL") Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                                ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                                End If
                                oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                If ActiveMatrix = "mx_TkrList" Or ActiveMatrix = "mx_DspList" Then
                                    If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                        LoadTruckingPO(ExportSeaFCLForm, RPOsrfname)
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
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
                                        LoadAndCreateCPO(ExportSeaFCLForm, RPOsrfname)
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        CPOForm = p_oSBOApplication.Forms.ActiveForm
                                        CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                                        CPOForm.Items.Item("1").Click()
                                        CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

                                        oComboPO = CPOForm.Items.Item("cb_Contact").Specific
                                        If Not ClearComboData(CPOForm, "cb_Contact", "@OBT_TB08_FFCPO", "U_CPerson") Then Throw New ArgumentException(sErrDesc)
                                        oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CPOForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            oRecordSet.MoveFirst()
                                            While oRecordSet.EoF = False
                                                oComboPO.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                                oRecordSet.MoveNext()

                                            End While
                                        End If
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                    End If
                                End If
                            Else
                                p_oSBOApplication.MessageBox("Need to select the Row that you want to Edit")
                            End If
                        End If

                    End If

                Case "CopyToCGR"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            Select Case ActiveMatrix
                                Case "mx_Crane"
                                    RGRsrfname = "CraneGoodsReceipt_Form.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                Case "mx_Bunk"
                                    RGRsrfname = "BunkGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                Case "mx_Toll"
                                    RGRsrfname = "TollGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                              
                                Case "mx_Courier"
                                    RGRsrfname = "CourierGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                Case "mx_DGLP"
                                    RGRsrfname = "DGLPGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                Case "mx_Outer"
                                    RGRsrfname = "OutriderGoodReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                Case "mx_COO"
                                    RGRsrfname = "CertificateOfOriginGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                                Case "mx_Fumi" 'Fumigation
                                    RGRsrfname = "FumiGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                Case "mx_Fork" 'to combine
                                    RGRsrfname = "ForkGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                Case "mx_Crate" 'to combine
                                    RGRsrfname = "CrateGoodsReceipt.srf"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                            End Select

                            If ActiveMatrix = "mx_TkrList" Or ActiveMatrix = "mx_DspList" Then
                                If currentRow > 0 Then
                                    If AlreadyExist("EXPORTSEAFCL") Then
                                        ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                                    ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                        ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                                    End If
                                    oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                    'Truck PO
                                    If ActiveMatrix = "mx_TkrList" Or ActiveMatrix = "mx_DspList" Then
                                        If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                            LoadAndCreateCGR(ExportSeaFCLForm, RGRsrfname)
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-201
                                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                                            If Not FillDataToGoodsReceipt(ExportSeaFCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
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
                                            LoadAndCreateCGR(ExportSeaFCLForm, RGRsrfname)
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                                            If Not FillDataToGoodsReceipt(ExportSeaFCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                        End If
                                    End If
                                Else
                                    p_oSBOApplication.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                                End If
                            Else
                                If currentRow > 0 Then
                                    oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                    If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                        LoadAndCreate_GoodReceipt(ExportSeaFCLForm, RGRsrfname)
                                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000027" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000030" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000033" Or _
                                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000036" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000039" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000042" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000007" Or _
                                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000016" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000061" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000012" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000011" Then

                                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                                            If Not FillDataToGoodsReceipt(ExportSeaFCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                                        End If
                                        oComboPO = CGRForm.Items.Item("cb_Contact").Specific
                                        If Not ClearComboData(CGRForm, "cb_Contact", "@OBT_TB12_FFCGR", "U_CPerson") Then Throw New ArgumentException(sErrDesc)
                                        oRecordSet.DoQuery("SELECT CntctCode, Name FROM OCPR WHERE CardCode = '" & CGRForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                        If oRecordSet.RecordCount > 0 Then
                                            oRecordSet.MoveFirst()
                                            While oRecordSet.EoF = False
                                                oComboPO.ValidValues.Add(oRecordSet.Fields.Item("CntctCode").Value, oRecordSet.Fields.Item("Name").Value)
                                                oRecordSet.MoveNext()
                                            End While
                                        End If

                                        MainForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE   'Fumigation
                                        MainForm.Items.Item("2").Specific.Caption = "Close"

                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                    End If
                                Else
                                    p_oSBOApplication.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                                End If

                            End If

                        Else
                            ExportSeaFCLForm.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                        End If
                    End If




                    'If pVal.BeforeAction = False Then
                    '    If currentRow > 0 Then
                    '        ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                    'LoadAndCreateCGR(ExportSeaFCLForm)
                    ''        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    '        CGRForm = p_oSBOApplication.Forms.ActiveForm
                    '        If Not FillDataToGoodsReceipt(ExportSeaFCLForm, "mx_Bok", "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                    '    Else
                    '        p_oSBOApplication.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                    '    End If
                    'End If

                Case "CancelPO"
                    Dim tblHeader As String = ""
                    Dim source As String = ""
                    If pVal.BeforeAction = False Then

                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000025" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000026" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000032" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000035" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000038" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000028" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000031" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000034" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000037" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000040" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "FUMIGATION" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "OUTRIDER" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "CRANE" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "BUNKER" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "TOLL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "FORKLIFT" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "CRATE" Then   'BUTTON           'Purchase
                            If currentRow > 0 Then
                                If ActiveMatrix = "mx_Crane" Then
                                    oMatrixName = "mx_Crane"
                                    tblHeader = "@CRANE"
                                    TableName = "@OBT_TB33_CRANE"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                    source = "Crane"
                                ElseIf ActiveMatrix = "mx_Fork" Then 'to combine
                                    oMatrixName = "mx_Fork"
                                    tblHeader = "@FORKLIFT"
                                    TableName = "@OBT_TBL05_FORKLIFT"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                    source = "Forklift"
                                ElseIf ActiveMatrix = "mx_Crate" Then 'to combine
                                    oMatrixName = "mx_Crate"
                                    tblHeader = "@CRATE"
                                    TableName = "@OBT_TBL08_CRATE"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                    source = "Crate"
                                ElseIf ActiveMatrix = "mx_Bunk" Then
                                    oMatrixName = "mx_Bunk"
                                    tblHeader = "@BUNKER"
                                    TableName = "@OBT_TB01_BUNKER"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                    source = "Bunker"
                                ElseIf ActiveMatrix = "mx_Toll" Then
                                    oMatrixName = "mx_Toll"
                                    tblHeader = "@TOLL"
                                    TableName = "@OBT_TB01_TOLL"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                    source = "Toll"
                                ElseIf ActiveMatrix = "mx_forklif" Then
                                    oMatrixName = "mx_forklif"
                                    TableName = "@OBT_FCL24_FORKLIFT"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                    source = "Forklift"
                                ElseIf ActiveMatrix = "mx_Courier" Then
                                    oMatrixName = "mx_Courier"
                                    TableName = "@OBT_FCL21_COURIER"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                ElseIf ActiveMatrix = "mx_DGLP" Then
                                    oMatrixName = "mx_DGLP"
                                    TableName = "@OBT_FCL23_DGLP"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                ElseIf ActiveMatrix = "mx_Outer" Then
                                    oMatrixName = "mx_Outer"
                                    tblHeader = "@OUTRIDER"
                                    TableName = "@OBT_TBL03_OUTRIDER"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("OUTRIDER", 1)
                                    source = "Outrider"
                                ElseIf ActiveMatrix = "mx_COO" Then
                                    oMatrixName = "mx_COO"
                                    TableName = "@OBT_FCL20_COO"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)

                                ElseIf ActiveMatrix = "mx_Fumi" Then 'Fumigation
                                    oMatrixName = "mx_Fumi"
                                    tblHeader = "@FUMIGATION"
                                    TableName = "@OBT_TBL01_FUMIGAT"
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("FUMIGATION", 1)
                                    source = "Fumigation"
                                End If

                                oMatrix = ExportSeaFCLForm.Items.Item(oMatrixName).Specific
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    If oMatrix.Columns.Item("colOriPO").Cells.Item(currentRow).Specific.Value.ToString = "Y" Then 'to km
                                        If oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value <> "" Then
                                            If oMatrix.Columns.Item("colMultiJb").Cells.Item(currentRow).Specific.Value() = "Y" Then
                                                Dim chk As Integer = 0
                                                chk = p_oSBOApplication.MessageBox("There are multi job in this PO.If you cancel this PO,PO status will change on other corresponding Job.Do you want to Cancel?", 1, "Yes", "No")
                                                If chk = 1 Then

                                                    modMultiJobForNormal.UpdateMultiJobPOStatusButton(ExportSeaFCLForm, "Cancelled", oMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value, TableName, tblHeader)

                                                    If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                                    If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                                    Dim sql As String = "select DocEntry,U_PONo,U_VCode,U_VName,U_VRef,U_SInA,U_TDate,U_TTime,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS,U_POIRMKS,U_PO from " & _
                                                               "[@OBT_TB08_FFCPO] where  DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())

                                                    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, source) Then Throw New ArgumentException(sErrDesc)

                                                    'If oMatrixName = "mx_Fumi" Then 'Fumigation
                                                    '    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, "Fumigation") Then Throw New ArgumentException(sErrDesc)
                                                    'ElseIf oMatrixName = "mx_Outer" Then
                                                    '    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, "Outrider") Then Throw New ArgumentException(sErrDesc)
                                                    'ElseIf oMatrixName = "mx_Crane" Then
                                                    '    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, "Crane") Then Throw New ArgumentException(sErrDesc)
                                                    'End If
                                                    oMatrix = MainForm.Items.Item("mx_PO").Specific 'Fumigation
                                                    If Not EditPOTab(MainForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc) 'New UI Edit Status to PO Tab
                                                End If
                                            Else '2-12
                                                If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                                If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                                Dim sql As String = "select DocEntry,U_PONo,U_VCode,U_VName,U_VRef,U_SInA,U_TDate,U_TTime,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS,U_POIRMKS,U_PO from " & _
                                                               "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())

                                                If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, source) Then Throw New ArgumentException(sErrDesc)

                                                'If oMatrixName = "mx_Fumi" Then 'Fumigation
                                                '    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, "Fumigation") Then Throw New ArgumentException(sErrDesc)
                                                'ElseIf oMatrixName = "mx_Outer" Then
                                                '    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, "Outrider") Then Throw New ArgumentException(sErrDesc)
                                                'ElseIf oMatrixName = "mx_Crane" Then
                                                '    If Not PopulateOtherPurchaseHeaderButton(ExportSeaFCLForm, oMatrix, sql, TableName, "Crane") Then Throw New ArgumentException(sErrDesc)
                                                'End If
                                                oMatrix = MainForm.Items.Item("mx_PO").Specific 'Fumigation
                                                If Not EditPOTab(MainForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc) 'New UI Edit Status to PO Tab
                                            End If
                                        End If

                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                        ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                                        ExportSeaFCLForm.Items.Item("1").Click()

                                        MainForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE   'Fumigation
                                        MainForm.Items.Item("2").Specific.Caption = "Close"
                                        MainForm.Items.Item("1").Click()
                                    Else
                                        p_oSBOApplication.MessageBox("You have no permission to Cancel this PO.")
                                    End If

                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                End If

                                ' p_oSBOApplication.ActivateMenuItem("1291")
                            Else
                                p_oSBOApplication.MessageBox("Need to select the Row that you want to Cancel")
                            End If
                        Else
                            If currentRow > 0 Then
                                If AlreadyExist("EXPORTSEAFCL") Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                                ElseIf AlreadyExist("EXPORTAIRFCL") Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
                                End If
                                oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                If ActiveMatrix = "mx_TkrList" Or ActiveMatrix = "mx_DspList" Then
                                    If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                        If oMatrix.Columns.Item("colOrigin").Cells.Item(currentRow).Specific.Value.ToString = "Y" Then 'to km
                                            If oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value <> "" Then
                                                If oMatrix.Columns.Item("colMulti").Cells.Item(currentRow).Specific.Value() = "Y" Then
                                                    Dim chk As Integer = 0
                                                    chk = p_oSBOApplication.MessageBox("There are multi job in this PO.If you cancel this PO,PO status will change on other corresponding Job.Do you want to Cancel?", 1, "Yes", "No")
                                                    If chk = 1 Then
                                                        If ActiveMatrix = "mx_TkrList" Then
                                                            modMultiJobForNormal.UpdateMultiJobPOStatus(ExportSeaFCLForm, "Cancelled", oMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value, "[@OBT_FCL03_ETRUCKING]")
                                                        ElseIf ActiveMatrix = "mx_DspList" Then
                                                            modMultiJobForNormal.UpdateMultiJobPOStatus(ExportSeaFCLForm, "Cancelled", oMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value, "[@OBT_FCL04_EDISPATCH]")
                                                        End If
                                                        If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                                        If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                                        oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                                        If ActiveMatrix = "mx_TkrList" Then
                                                            sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_TkrCode,b.U_Trucker,b.U_VehNo," & _
                                                               "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime,b.U_PO " & _
                                                               "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL03_ETRUCKING] b on a.DocEntry=b.U_PODocNo where b.U_PODocNo = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                                            If Not PopulateTruckPurchaseHeader(ExportSeaFCLForm, oMatrix, sql, "@OBT_FCL03_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                                            oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                                            If Not EditPOTab(ExportSeaFCLForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc) 'New UI Edit Status to PO Tab
                                                        ElseIf ActiveMatrix = "mx_DspList" Then
                                                            sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_DspCode,b.U_Dispatch," & _
                                                             "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_DspDate,b.U_DspTime,b.U_PO " & _
                                                             "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL04_EDISPATCH] b on a.DocEntry=b.U_PODocNo where b.U_PODocNo = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                                            If Not PopulateDispatchPurchaseHeader(ExportSeaFCLForm, oMatrix, sql, "@OBT_FCL04_EDISPATCH") Then Throw New ArgumentException(sErrDesc)
                                                            oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                                            If Not EditPOTab(ExportSeaFCLForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc) 'New UI Edit Status to PO Tab
                                                        End If
                                                    End If
                                                Else '2-12
                                                    If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                                    If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                                    oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                                    If ActiveMatrix = "mx_TkrList" Then
                                                        sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_TkrCode,b.U_Trucker,b.U_VehNo," & _
                                                           "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime,b.U_PO " & _
                                                           "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL03_ETRUCKING] b on a.DocEntry=b.U_PODocNo where b.U_PODocNo = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                                        If Not PopulateTruckPurchaseHeader(ExportSeaFCLForm, oMatrix, sql, "@OBT_FCL03_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                                        If Not EditPOTab(ExportSeaFCLForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc) 'New UI Edit Status to PO Tab
                                                    ElseIf ActiveMatrix = "mx_DspList" Then
                                                        sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_DspCode,b.U_Dispatch," & _
                                                         "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_DspDate,b.U_DspTime,b.U_PO " & _
                                                         "from  [@OBT_TB08_FFCPO] a inner join [@OBT_FCL04_EDISPATCH] b on a.DocEntry=b.U_PODocNo where b.U_PODocNo = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                                        If Not PopulateDispatchPurchaseHeader(ExportSeaFCLForm, oMatrix, sql, "@OBT_FCL04_EDISPATCH") Then Throw New ArgumentException(sErrDesc)
                                                        oMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                                        If Not EditPOTab(ExportSeaFCLForm, oMatrix, sql, "@OBT_TB01_POLIST") Then Throw New ArgumentException(sErrDesc) 'New UI Edit Status to PO Tab
                                                    End If
                                                End If
                                                ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                                oMatrix = ExportSeaFCLForm.Items.Item(ActiveMatrix).Specific
                                                ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
                                                ExportSeaFCLForm.Items.Item("1").Click()

                                            End If
                                        Else
                                            p_oSBOApplication.MessageBox("You have no permission to Cancel this PO.")
                                        End If


                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                        p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                    End If

                                End If

                            Else
                                p_oSBOApplication.MessageBox("Need to select the Row that you want to Cancel")
                            End If
                        End If

                    End If

            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoExportSeaFCLMenuEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaFCLMenuEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function

    Public Function DoExportSeaFCLRightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
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
        Dim ExportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim bFlag As Boolean = False
        Dim FunctionName As String = "DoImportSeaLCLRightClickEvent"
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim formuid As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)

            ExportSeaFCLForm = p_oSBOApplication.Forms.ActiveForm
            oMenuItem = p_oSBOApplication.Menus.Item("1280")
            ExportSeaFCLForm.EnableMenu("772", True)
            ExportSeaFCLForm.EnableMenu("773", True)
            ExportSeaFCLForm.EnableMenu("775", True)
            oMenus = oMenuItem.SubMenus
            If (p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL") And eventInfo.ItemUID = "mx_TkrList" Or eventInfo.ItemUID = "mx_DspList" Then 'Truck PO
                oMatrix = ExportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
                If eventInfo.BeforeAction = True Then
                    ExportSeaFCLForm.EnableMenu("772", False)
                    ExportSeaFCLForm.EnableMenu("773", False)
                    ExportSeaFCLForm.EnableMenu("775", False)
                    If oMenus.Exists("EditVoc") Then
                        p_oSBOApplication.Menus.RemoveEx("EditVoc")
                    ElseIf oMenus.Exists("EditShp") Then
                        p_oSBOApplication.Menus.RemoveEx("EditShp")
                    End If

                    If eventInfo.ItemUID = "mx_TkrList" Or eventInfo.ItemUID = "mx_DspList" Then
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colInsDoc").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                                    p_oSBOApplication.Menus.RemoveEx("EditCPO")
                                    p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                                    p_oSBOApplication.Menus.RemoveEx("CancelPO")
                                End If
                                If Not oMenus.Exists("EditCPO") Then
                                    oMenus.Add("EditCPO", "Edit Custom Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "PurchaseOrder.srf"
                                End If
                                If Not oMenus.Exists("CopyToCGR") Then
                                    oMenus.Add("CopyToCGR", "Copy To Custom Goods Receipt", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RGRmatrixname = eventInfo.ItemUID
                                    RGRsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "GoodsReceipt.srf"
                                End If
                                If Not oMenus.Exists("CancelPO") Then
                                    oMenus.Add("CancelPO", "Cancel Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                End If
                                currentRow = eventInfo.Row
                                ActiveMatrix = eventInfo.ItemUID

                            End If
                        End If

                    Else
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                                    p_oSBOApplication.Menus.RemoveEx("EditCPO")
                                    p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                                    p_oSBOApplication.Menus.RemoveEx("CancelPO")
                                End If
                                If Not oMenus.Exists("EditCPO") Then
                                    oMenus.Add("EditCPO", "Edit Custom Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "PurchaseOrder.srf"
                                End If
                                If Not oMenus.Exists("CopyToCGR") Then
                                    oMenus.Add("CopyToCGR", "Copy To Custom Goods Receipt", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RGRmatrixname = eventInfo.ItemUID
                                    RGRsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "GoodsReceipt.srf"
                                End If
                                If Not oMenus.Exists("CancelPO") Then
                                    oMenus.Add("CancelPO", "Cancel Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
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
            If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000025" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000028" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000031" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000034" Or _
                p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000037" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000040" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "FUMIGATION" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "OUTRIDER" Or _
                p_oSBOApplication.Forms.ActiveForm.TypeEx = "CRANE" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "BUNKER" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "TOLL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "FORKLIFT" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "CRATE" Then
                If eventInfo.ItemUID = "mx_Fumi" Or eventInfo.ItemUID = "mx_Outer" Or eventInfo.ItemUID = "mx_Crane" Or eventInfo.ItemUID = "mx_Bunk" Or eventInfo.ItemUID = "mx_Toll" Or eventInfo.ItemUID = "mx_Fork" Or eventInfo.ItemUID = "mx_Crate" Then 'Truck PO

                    oMatrix = ExportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
                    If eventInfo.BeforeAction = True Then
                        ExportSeaFCLForm.EnableMenu("772", False)
                        ExportSeaFCLForm.EnableMenu("773", False)
                        ExportSeaFCLForm.EnableMenu("775", False)
                        If oMenus.Exists("EditVoc") Then
                            p_oSBOApplication.Menus.RemoveEx("EditVoc")
                        ElseIf oMenus.Exists("EditShp") Then
                            p_oSBOApplication.Menus.RemoveEx("EditShp")
                        End If
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                                    p_oSBOApplication.Menus.RemoveEx("EditCPO")
                                    p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                                    p_oSBOApplication.Menus.RemoveEx("CancelPO")
                                End If
                                If Not oMenus.Exists("EditCPO") Then
                                    oMenus.Add("EditCPO", "Edit Custom Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "PurchaseOrder.srf"
                                End If
                                If Not oMenus.Exists("CopyToCGR") Then
                                    oMenus.Add("CopyToCGR", "Copy To Custom Goods Receipt", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RGRmatrixname = eventInfo.ItemUID
                                    RGRsrfname = eventInfo.ItemUID.Substring(3, eventInfo.ItemUID.Length - 3) + "GoodsReceipt.srf"
                                End If
                                If Not oMenus.Exists("CancelPO") Then
                                    oMenus.Add("CancelPO", "Cancel Purchase Order", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                End If
                                currentRow = eventInfo.Row
                                ActiveMatrix = eventInfo.ItemUID

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
            End If

            If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEAFCL" Then
                If eventInfo.ItemUID = "mx_ShpInv" Then
                    oMatrix = ExportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
                    If eventInfo.BeforeAction = True Then
                        ExportSeaFCLForm.EnableMenu("772", False)
                        ExportSeaFCLForm.EnableMenu("773", False)
                        ExportSeaFCLForm.EnableMenu("775", False)
                        If oMenus.Exists("EditVoc") Then
                            p_oSBOApplication.Menus.RemoveEx("EditVoc")
                        End If
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNum").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If Not oMenus.Exists("EditShp") Then
                                    oMenus.Add("EditShp", "Edit Shipping Invoice", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = "ShipInvoice.srf"
                                End If
                                currentRow = eventInfo.Row
                                ActiveMatrix = eventInfo.ItemUID 'New UI
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
                If eventInfo.ItemUID = "mx_ConTab" Or eventInfo.ItemUID = "mx_Charge" Then
                    Dim strCol As String = ""
                    oMatrix = ExportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
                    ExportSeaFCLForm.EnableMenu("772", False)
                    ExportSeaFCLForm.EnableMenu("773", False)
                    ExportSeaFCLForm.EnableMenu("775", False)

                    If oMenus.Exists("EditShp") Then
                        p_oSBOApplication.Menus.RemoveEx("EditShp")
                    End If

                    If eventInfo.BeforeAction = True Then
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If eventInfo.ItemUID = "mx_ConTab" Then
                                strCol = "colCNo1"
                            ElseIf eventInfo.ItemUID = "mx_Charge" Then
                                strCol = "colCCode1"
                            End If
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item(strCol).Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                ExportSeaFCLForm.EnableMenu("1292", True)
                                ExportSeaFCLForm.EnableMenu("1293", True)
                            End If
                            currentRow = eventInfo.Row
                            ActiveMatrix = eventInfo.ItemUID
                        End If
                    Else
                        ExportSeaFCLForm.EnableMenu("1292", False)
                        ExportSeaFCLForm.EnableMenu("1293", False)
                    End If
                End If
                If eventInfo.ItemUID = "mx_ShpInv" Or eventInfo.ItemUID = "mx_Voucher" Then
                    ExportSeaFCLForm.EnableMenu("772", False)
                    ExportSeaFCLForm.EnableMenu("773", False)
                    ExportSeaFCLForm.EnableMenu("775", False)
                End If
            End If

            DoExportSeaFCLRightClickEvent = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaFCLRightClickEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function

    Private Sub LoadAndCreate_GoodReceipt(ByRef ParentForm As SAPbouiCOM.Form, ByVal FormName As String)
        Dim CGRForm As SAPbouiCOM.Form
        Dim CGRMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Try
            oMenuItem = p_oSBOApplication.Menus.Item("1280")
            oMenus = oMenuItem.SubMenus
            LoadFromXML(p_oSBOApplication, FormName)
            If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                p_oSBOApplication.Menus.RemoveEx("EditCPO")
                p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                p_oSBOApplication.Menus.RemoveEx("CancelPO")
            End If
            CGRForm = p_oSBOApplication.Forms.ActiveForm
            CGRForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
            CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_JobNo").Specific.Value)
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
        Dim itemTotal As Long
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
                    'Dim tsql As String = "SELECT U_SInA,U_TPlace,U_vref,U_TDate,U_POITPD,U_TDay,U_TTime,U_PORMKS,U_POIRMKS,U_CNo,U_Dest FROM [@OBT_TB08_FFCPO] WHERE DocEntry = " + FormatString(oMatrix.Columns.Item(SourceColName2).Cells.Item(ActiveRow).Specific.Value)
                    Dim tsql As String = "SELECT U_SInA,U_TPlace,U_TDate,U_vref,U_TDay,U_TTime,U_PORMKS,U_POIRMKS,U_ColFrm,U_TkrIns,U_TkrTo,U_CNo,U_Dest,U_LocWork,U_POITPD FROM [@OBT_TB08_FFCPO] WHERE DocEntry = " + FormatString(oMatrix.Columns.Item(SourceColName2).Cells.Item(ActiveRow).Specific.Value) 'MSW 14-09-2011 Truck PO
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
                            .SetValue("U_VRef", .Offset, oRecordset.Fields.Item("U_vref").Value)
                            .SetValue("U_GRTPD", .Offset, oRecordset.Fields.Item("U_POITPD").Value)
                            .SetValue("U_CNo", .Offset, oRecordset.Fields.Item("U_CNo").Value)
                            oRecordset.MoveNext()
                        End While
                    End If
                End With
            End If
            Dim tempSQL As String = "SELECT * FROM POR1 WHERE DocEntry =" + FormatString(oPODocument.DocEntry) + " And OpenQty <> 0 "
            Dim K As Integer = 1
            oRecordset.DoQuery(tempSQL)
            oDestMatrix.Clear()

            If oRecordset.RecordCount > 0 Then
                oRecordset.MoveFirst()
                While oRecordset.EoF = False
                    With oLineDBDataSource
                        .SetValue("LineId", .Offset, K)
                        .SetValue("U_GRINO", .Offset, oRecordset.Fields.Item("ItemCode").Value)
                        .SetValue("U_GRIDesc", .Offset, oRecordset.Fields.Item("Dscription").Value)
                        '.SetValue("U_GRIQty", .Offset, oRecordset.Fields.Item("Quantity").Value)
                        .SetValue("U_GRIQty", .Offset, oRecordset.Fields.Item("OpenQty").Value)
                        .SetValue("U_GRIPrice", .Offset, oRecordset.Fields.Item("Price").Value)
                        .SetValue("U_GRIAmt", .Offset, oRecordset.Fields.Item("OpenSum").Value)
                        .SetValue("U_GRITot", .Offset, oRecordset.Fields.Item("OpenSum").Value)
                        .SetValue("U_GRIGST", .Offset, oRecordset.Fields.Item("VatGroup").Value)
                        .SetValue("U_POLineId", .Offset, oRecordset.Fields.Item("LineNum").Value)
                        itemTotal += oRecordset.Fields.Item("OpenSum").Value
                        K = K + 1
                    End With
                    oDestMatrix.AddRow()
                    oRecordset.MoveNext()
                End While
                oDestForm.Items.Item("ed_TPDue").Specific.value = itemTotal
            End If
            FillDataToGoodsReceipt = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            FillDataToGoodsReceipt = False
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
            'p_oSBOApplication.MessageBox("SELECT Name FROM OCPR WHERE CntctCode = " & FormatString(ContactPersonCode) & " AND CardCode = " & FormatString(CardCode))
            If oRecordSet.RecordCount > 0 Then
                GetContactPersonName = oRecordSet.Fields.Item("Name").Value.ToString
            End If
        Catch ex As Exception
            GetContactPersonName = vbNullString
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Sub LoadAndCreateCGR(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)
        Dim CGRForm As SAPbouiCOM.Form
        Dim CGRMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Try
            oMenuItem = p_oSBOApplication.Menus.Item("1280")
            oMenus = oMenuItem.SubMenus
            LoadFromXML(p_oSBOApplication, srfName)
            If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                p_oSBOApplication.Menus.RemoveEx("EditCPO")
                p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                p_oSBOApplication.Menus.RemoveEx("CancelPO")
            End If
            CGRForm = p_oSBOApplication.Forms.ActiveForm
            CGRForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
        Catch ex As Exception
            MessageBox.Show(ex.Message)

        End Try
    End Sub

    Private Sub DeleteMatrixRow(ByRef ExportSeaFCLForm As SAPbouiCOM.Form, ByRef oMatrix As SAPbouiCOM.Matrix, ByVal objDataSource As String, ByVal oColumn As String)
        '=============================================================================
        'Function   : DeleteMatrixRow()
        'Purpose    : This function to delete matrixrow of line Id 
        'Parameters : ByRef ExportSeaFCLForm As SAPbouiCOM.Form,ByRef oMatrix As SAPbouiCOM.Matrix,
        '           : ByVal objDataSource As String, ByVal oColumn As String       
        'Return     : No
        '=============================================================================

        Dim tblname As String = objDataSource.Substring(1, objDataSource.Length - 1)
        Dim lRow As Long
        lRow = oMatrix.GetNextSelectedRow
        If lRow > -1 Then
            ExportSeaFCLForm.DataSources.DBDataSources.Item(objDataSource).RemoveRecord(0)
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

    Private Sub LoadAndCreate_PurChaseOrder(ByRef ParentForm As SAPbouiCOM.Form, ByVal FormName As String)
        Dim CPOForm As SAPbouiCOM.Form
        Dim CPOMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Try
            oMenuItem = p_oSBOApplication.Menus.Item("1280")
            oMenus = oMenuItem.SubMenus
            LoadFromXML(p_oSBOApplication, FormName)
            If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                p_oSBOApplication.Menus.RemoveEx("EditCPO")
                p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                p_oSBOApplication.Menus.RemoveEx("CancelPO")
            End If
            CPOForm = p_oSBOApplication.Forms.ActiveForm
            CPOForm.Freeze(True)
            CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_JobNo").Specific.Value)
                CPOForm.Items.Item("ed_Code").Specific.Active = True


                If AddUserDataSrc(CPOForm, "Email", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                CPOForm.Items.Item("ch_Email").Specific.DataBind.SetBound(True, "", "Email")

                If AddUserDataSrc(CPOForm, "Fax", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                CPOForm.Items.Item("ch_Fax").Specific.DataBind.SetBound(True, "", "Fax")

                If AddUserDataSrc(CPOForm, "Print", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                CPOForm.Items.Item("ch_Print").Specific.DataBind.SetBound(True, "", "Print")
                CPOForm.Freeze(False)

                'If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                'CPOForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_DocNum").Specific.Value)
                'CPOForm.Items.Item("ed_Code").Specific.Active = True



                ' ==================================== Custom Purchase Order ========================================
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function Validateforform(ByVal ItemUID As String, ByVal ExportSeaFCLForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   Validateforform()
        '   Purpose     :   This function will be providing to validate Form   for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  : ByVal ItemUID As String, ByVal ExportSeaFCLForm As SAPbouiCOM.Form
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************]
        Try
            If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_Name").Specific.String) Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            ElseIf (ItemUID = "ed_IShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.String) And ExportSeaFCLForm.Title.Contains("Import") Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Carrier Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.String) And ExportSeaFCLForm.Title.Contains("Export") Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Carrier Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            ElseIf (String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.String) And String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.String)) And ExportSeaFCLForm.Title.Contains("Transhipment") Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Carrier Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            ElseIf (String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.String) And String.IsNullOrEmpty(ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.String)) And ExportSeaFCLForm.Title.Contains("Local") Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Carrier Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
        
    End Function

    Private Sub Start(ByRef pform As SAPbouiCOM.Form)

        '=============================================================================
        'Function   : Start()
        'Purpose    : This function to provide for To process with pdf form in ExporeSealCL 
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
                MoveWindow(myprocess.MainWindowHandle, p_oSBOApplication.Forms.ActiveForm.Left + p_oSBOApplication.Forms.ActiveForm.Width + 6, p_oSBOApplication.Forms.ActiveForm.Top + 68, 300, 675, True)
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub LoadAndCreateCPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)
        Dim CPOForm As SAPbouiCOM.Form
        Dim CPOMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sErrDesc As String = vbNullString
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Try
            oMenuItem = p_oSBOApplication.Menus.Item("1280")
            oMenus = oMenuItem.SubMenus
            LoadFromXML(p_oSBOApplication, srfName)
            If oMenus.Exists("EditCPO") And oMenus.Exists("CancelPO") And oMenus.Exists("CopyToCGR") Then
                p_oSBOApplication.Menus.RemoveEx("EditCPO")
                p_oSBOApplication.Menus.RemoveEx("CopyToCGR")
                p_oSBOApplication.Menus.RemoveEx("CancelPO")
            End If
            CPOForm = p_oSBOApplication.Forms.ActiveForm
            CPOForm.Freeze(True)
            CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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

                If srfName = "CranePurchaseOrder.srf" Then
                    CPOForm.Items.Item("ed_Code").Specific.Value = ParentForm.Items.Item("ed_CVendor").Specific.Value
                ElseIf srfName = "ForkPurchaseOrder.srf" Then
                    CPOForm.Items.Item("ed_Code").Specific.Value = ParentForm.Items.Item("ed_FVendor").Specific.Value
                End If
                CPOForm.Items.Item("ed_Code").Specific.Active = True
                CPOForm.Freeze(False)
                ' ==================================== Custom Purchase Order ========================================
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub PutDBValueToDBAndPreviewForShippingOrder(ByVal pdfname As String, ByVal pform As SAPbouiCOM.Form, ByVal pdfFilePath As String, ByVal originpdf As String)


        ' **********************************************************************************
        '   Function    :   PutDBValueToDBAndPreviewForShippingOrder()
        '   Purpose     :   This function will be providing to Save Database and preview shipping order form 
        '   Parameters  :   ByVal pdfname As String, ByVal pform As SAPbouiCOM.Form,
        '               :   ByVal pdfFilePath As String, ByVal originpdf As String
        '   Return      :   No
        '                 
        ' **********************************************************************************

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Dim outputFilePath As String = AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName + "\" + pdfname + "1.pdf"
        'Get pdf from project directory
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing

        Try
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)

            ' Create the form filler
            Using pdfOutputFile As New IO.FileStream(pdfFilePath, IO.FileMode.Create)
                Dim formFiller As iTextSharp.text.pdf.PdfStamper = Nothing
                Try
                    formFiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfOutputFile)
                    ' Get the form fields
                    Dim addressChangeForm As iTextSharp.text.pdf.AcroFields = formFiller.AcroFields

                    ' Fill Shipper 
                    Dim sqlshipper As String = "SELECT CompnyName, CompnyAddr FROM dbo.OADM"
                    oRecordSet.DoQuery(sqlshipper)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            addressChangeForm.SetField("Shipper", oRecordSet.Fields.Item("CompnyName").Value.ToString())
                            addressChangeForm.SetField("ShipperAddress", oRecordSet.Fields.Item("CompnyAddr").Value.ToString())

                            oRecordSet.MoveNext()
                        End While
                    End If

                    ' Fill the form
                    Dim sql As String = "SELECT dbo.[@OBT_FCL01_EXPORT].U_Vessel,dbo.[@OBT_FCL01_EXPORT].U_Voyage, dbo.[@OBT_FCL01_EXPORT].U_PName, dbo.OCRD.CardCode, dbo.OCRD.CardName, dbo.OCRD.Address, dbo.OCRD.Phone1,dbo.OCRD.Fax FROM dbo.[@OBT_FCL01_EXPORT] INNER JOIN dbo.OCRD ON dbo.[@OBT_FCL01_EXPORT].U_Code = dbo.OCRD.CardCode WHERE dbo.[@OBT_FCL01_EXPORT].U_JobNum =" & FormatString(pform.Items.Item("ed_JobNo").Specific.Value)
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            addressChangeForm.SetField("ConsigneeName", oRecordSet.Fields.Item("CardName").Value.ToString())
                            addressChangeForm.SetField("ConsigneeAddress", oRecordSet.Fields.Item("Address").Value.ToString())
                            addressChangeForm.SetField("Telephone", oRecordSet.Fields.Item("Phone1").Value.ToString())
                            addressChangeForm.SetField("Fax", oRecordSet.Fields.Item("Fax").Value.ToString())
                            addressChangeForm.SetField("Vessel", oRecordSet.Fields.Item("U_Vessel").Value.ToString())
                            addressChangeForm.SetField("Voyage", oRecordSet.Fields.Item("U_Voyage").Value.ToString())
                            addressChangeForm.SetField("PortOfLoading", oRecordSet.Fields.Item("U_PName").Value.ToString())
                            oRecordSet.MoveNext()
                        End While
                    End If
                Catch

                Finally
                    If formFiller IsNot Nothing Then
                        formFiller.Close()
                    End If
                End Try
            End Using
        Finally
            reader.Close()
        End Try
        Process.Start(pdfFilePath)
    End Sub

    Private Sub PutDBValueToDBAndPreviewForDGD(ByVal pdfname As String, ByVal pform As SAPbouiCOM.Form, ByVal pdfFilePath As String, ByVal originpdf As String)



        ' **********************************************************************************
        '   Function    :   PutDBValueToDBAndPreviewForDGD()
        '   Purpose     :   This function will be providing to Save DB to draft table in ExportSeaLCL form 
        '   Parameters  :   ByVal pdfname As String, ByVal pform As SAPbouiCOM.Form,
        '               :   ByVal pdfFilePath As String, ByVal originpdf As String
        '   Return      :   No
        '                 
        ' **********************************************************************************
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' Dim outputFilePath As String = AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName + "\" + pdfname + "1.pdf"
        'Get pdf from project directory
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing

        Try
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)

            ' Create the form filler
            Using pdfOutputFile As New IO.FileStream(pdfFilePath, IO.FileMode.Create)
                Dim formFiller As iTextSharp.text.pdf.PdfStamper = Nothing
                Try
                    formFiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfOutputFile)
                    ' Get the form fields
                    Dim addressChangeForm As iTextSharp.text.pdf.AcroFields = formFiller.AcroFields

                    ' Fill Shipper 
                    Dim sqlshipper As String = "SELECT CompnyName, CompnyAddr FROM dbo.OADM"
                    oRecordSet.DoQuery(sqlshipper)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            addressChangeForm.SetField("ShipperName", oRecordSet.Fields.Item("CompnyName").Value.ToString())
                            addressChangeForm.SetField("ShipperAddress", oRecordSet.Fields.Item("CompnyAddr").Value.ToString())

                            oRecordSet.MoveNext()
                        End While
                    End If

                    ' Fill the form
                    Dim sql As String = "SELECT dbo.[@OBT_FCL01_EXPORT].U_Vessel, dbo.[@OBT_FCL01_EXPORT].U_PName, dbo.OCRD.CardCode, dbo.OCRD.CardName, dbo.OCRD.Address, dbo.OCRD.Phone1,dbo.OCRD.Fax FROM dbo.[@OBT_FCL01_EXPORT] INNER JOIN dbo.OCRD ON dbo.[@OBT_FCL01_EXPORT].U_Code = dbo.OCRD.CardCode WHERE dbo.[@OBT_FCL01_EXPORT].U_JobNum =" & FormatString(pform.Items.Item("ed_JobNo").Specific.Value)
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            addressChangeForm.SetField("ConsigneeName", oRecordSet.Fields.Item("CardName").Value.ToString())
                            addressChangeForm.SetField("ConsigneeAddress", oRecordSet.Fields.Item("Address").Value.ToString())
                            addressChangeForm.SetField("Telephone", oRecordSet.Fields.Item("Phone1").Value.ToString())
                            addressChangeForm.SetField("Fax", oRecordSet.Fields.Item("Fax").Value.ToString())
                            addressChangeForm.SetField("Vessel", oRecordSet.Fields.Item("U_Vessel").Value.ToString())
                            addressChangeForm.SetField("PortOfLoading", oRecordSet.Fields.Item("U_PName").Value.ToString())
                            oRecordSet.MoveNext()
                        End While
                    End If
                Catch

                Finally
                    If formFiller IsNot Nothing Then
                        formFiller.Close()
                    End If
                End Try
            End Using
        Finally
            reader.Close()
        End Try
        Process.Start(pdfFilePath)
    End Sub

    Private Sub PutDBValueToDBAndPreviewForDraftBL(ByVal pdfname As String, ByVal pform As SAPbouiCOM.Form, ByVal pdfFilePath As String, ByVal originpdf As String)


        ' **********************************************************************************
        '   Function    :   PutDBValueToDBAndPreviewForDraftBL()
        '   Purpose     :   This function will be providing to copy value from DB to draft table in ExportSeaLCL form 
        '   Parameters  :   ByVal pdfname As String, ByVal pform As SAPbouiCOM.Form,
        '               :   ByVal pdfFilePath As String, ByVal originpdf As String
        '   Return      :   No
        '                 
        ' **********************************************************************************


        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Dim outputFilePath As String = AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName + "\" + pdfname + "1.pdf"
        'Get pdf from project directory
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing

        Try
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)

            ' Create the form filler
            Using pdfOutputFile As New IO.FileStream(pdfFilePath, IO.FileMode.Create)
                Dim formFiller As iTextSharp.text.pdf.PdfStamper = Nothing
                Try
                    formFiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfOutputFile)
                    ' Get the form fields
                    Dim addressChangeForm As iTextSharp.text.pdf.AcroFields = formFiller.AcroFields


                    ' Fill Shipper 
                    Dim sqlshipper As String = "SELECT CompnyName, CompnyAddr FROM dbo.OADM"
                    oRecordSet.DoQuery(sqlshipper)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            addressChangeForm.SetField("Shipper", oRecordSet.Fields.Item("CompnyName").Value.ToString())
                            addressChangeForm.SetField("ShipperAddress", oRecordSet.Fields.Item("CompnyAddr").Value.ToString())

                            oRecordSet.MoveNext()
                        End While
                    End If

                    ' Fill the form
                    Dim sql As String = "SELECT dbo.[@OBT_FCL01_EXPORT].U_Vessel, dbo.[@OBT_FCL01_EXPORT].U_Voyage, dbo.[@OBT_FCL01_EXPORT].U_PName, dbo.[@OBT_FCL01_EXPORT].U_OceanBL, dbo.[@OBT_FCL01_EXPORT].U_HouseBL, dbo.[@OBT_FCL01_EXPORT].U_Conn, dbo.[@OBT_FCL01_EXPORT].U_TotalWt, dbo.[@OBT_FCL01_EXPORT].U_TotalM3,dbo.[@OBT_FCL01_EXPORT].U_CrgDsc, dbo.OCRD.CardCode, dbo.OCRD.CardName, dbo.OCRD.Address, dbo.OCRD.Fax, dbo.OCRD.Phone1, dbo.OCRD.Phone2 FROM dbo.[@OBT_FCL01_EXPORT] INNER JOIN dbo.OCRD ON dbo.[@OBT_FCL01_EXPORT].U_Code = dbo.OCRD.CardCode WHERE dbo.[@OBT_FCL01_EXPORT].U_JobNum =" & FormatString(pform.Items.Item("ed_JobNo").Specific.Value)
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            addressChangeForm.SetField("ConsigneeName", oRecordSet.Fields.Item("CardName").Value.ToString())
                            addressChangeForm.SetField("ConsigneeAddress", oRecordSet.Fields.Item("Address").Value.ToString())
                            addressChangeForm.SetField("Telephone", oRecordSet.Fields.Item("Phone1").Value.ToString())
                            addressChangeForm.SetField("Fax", oRecordSet.Fields.Item("Fax").Value.ToString())
                            addressChangeForm.SetField("Vessel", oRecordSet.Fields.Item("U_Vessel").Value.ToString())
                            addressChangeForm.SetField("VoyNo", oRecordSet.Fields.Item("U_Voyage").Value.ToString())
                            addressChangeForm.SetField("Pol", oRecordSet.Fields.Item("U_PName").Value.ToString())
                            addressChangeForm.SetField("NumberofBL", oRecordSet.Fields.Item("U_OceanBL").Value.ToString())
                            addressChangeForm.SetField("ContainerNo", oRecordSet.Fields.Item("U_Conn").Value.ToString())
                            addressChangeForm.SetField("TotalWt", oRecordSet.Fields.Item("U_TotalWt").Value.ToString())
                            addressChangeForm.SetField("TotalM3", oRecordSet.Fields.Item("U_TotalM3").Value.ToString())
                            addressChangeForm.SetField("CargoDescrip", oRecordSet.Fields.Item("U_CrgDsc").Value.ToString())

                            oRecordSet.MoveNext()
                        End While
                    End If

                Catch

                Finally
                    If formFiller IsNot Nothing Then
                        formFiller.Close()
                    End If
                End Try
            End Using
        Finally
            reader.Close()
        End Try
        Process.Start(pdfFilePath)
    End Sub

    Private Sub LoadPaymentVoucher(ByRef oActiveForm As SAPbouiCOM.Form)
        '=============================================================================
        'Function   : LoadPaymentVoucher()
        'Purpose    : This function to provide for load payment voucher form in ExporeSealCL 
        'Parameters : ByRef oActiveForm As SAPbouiCOM.Form
        'Return     : No

        '==========================================

        Dim oPayForm As SAPbouiCOM.Form
        Dim oOptBtn As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        LoadFromXML(p_oSBOApplication, "PaymentVoucher.srf")
        oPayForm = p_oSBOApplication.Forms.ActiveForm
        oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
        oPayForm.Freeze(True)
        If AddChooseFromList(oPayForm, "PAYMENT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oPayForm.Items.Item("ed_VedCode").Specific.ChooseFromListUID = "PAYMENT"
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
        oPayForm.Items.Item("ed_PayRate").Enabled = True
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPayForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("VOUCHER", oRecordSet)
        oPayForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_InvDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        oPayForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()


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
        oPayForm.Items.Item("bt_PayView").Visible = False
        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = oPayForm.Items.Item("cb_PayCur").Specific
        If oCombo.ValidValues.Count = 0 Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT CurrCode,CurrName FROM OCRN Where CurrCode In ('SGD','USD')")
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

            Dim oColumn As SAPbouiCOM.Column
            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
            If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
            oColumn = oMatrix.Columns.Item("colChCode1")
            AddChooseFromList(oPayForm, "ChCode", False, "UDOCHCODE")
            oColumn.ChooseFromListUID = "ChCode"
            oColumn.ChooseFromListAlias = "Code" 'MSW To Edit New Ticket
            oMatrix = oPayForm.Items.Item("mx_ChCode").Specific
            DisableChargeMatrix(oPayForm, oMatrix, True)
            'oPayForm.Items.Item("ed_VedName").Specific.Active = True
            oPayForm.Items.Item("ed_VedCode").Specific.Active = True
            oPayForm.Freeze(False)
    End Sub

    Private Sub RowAddToMatrix(ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        ' **********************************************************************************
        '   Function    :   RowAddToMatrix()
        '   Purpose     :   This function will be providing to proceed Add row  for
        '                    Purechase order matrix item value for Purchase Order process
        '   Parameters  : ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix    '                          
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

    Private Function AddChooseFromListByOption(ByRef pForm As SAPbouiCOM.Form, ByVal pOption As Boolean, ByVal pObjID As String, ByVal objInter As String, ByVal cflInter As String, ByVal objExter As String, ByVal cflExter As String, ByVal pErrDesc As String) As Long
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
                oEditText.DataBind.SetBound(True, "", objInter)
                oEditText.ChooseFromListUID = cflInter
                oEditText.ChooseFromListAlias = "firstName"
            Else
                oEditText = pForm.Items.Item(pObjID).Specific
                oEditText.DataBind.SetBound(True, "", objExter)
                oEditText.ChooseFromListUID = cflExter
                oEditText.ChooseFromListAlias = "CardName"
            End If
            AddChooseFromListByOption = RTN_SUCCESS

        Catch ex As Exception
            AddChooseFromListByOption = RTN_ERROR
        End Try
    End Function

    Private Sub SetMatrixSeqNo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal ColName As String)
        For i As Integer = 1 To oMatrix.RowCount
            oMatrix.Columns.Item(ColName).Cells.Item(i).Specific.Value = i
        Next
    End Sub

    Private Sub LoadButtonForm(ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String, ByVal KeyName As String)

        ' **********************************************************************************
        '   Function    :   LoadButtonForm()
        '   Purpose     :   This function provide to Load and show Button Forms when Button clikede of main form  
        '   Parameters  :   ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String,
        '               :   ByVal KeyName As String 
        '   return      :   No          
        ' **********************************************************************************

        Dim oActiveForm As SAPbouiCOM.Form
        Dim sErrDesc As String = ""
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sFuncName As String = "SBO_Application_MenuEvent()"
        ' Dim oCombo, oComboMain As SAPbouiCOM.ComboBox
        If Not p_oDICompany.Connected Then
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
            If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        End If
        LoadFromXML(p_oSBOApplication, FormName)
        oActiveForm = p_oSBOApplication.Forms.ActiveForm
        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

        oActiveForm.EnableMenu("1288", True)
        oActiveForm.EnableMenu("1289", True)
        oActiveForm.EnableMenu("1290", True)
        oActiveForm.EnableMenu("1291", True)
        oActiveForm.EnableMenu("771", False)
        oActiveForm.EnableMenu("774", False)
        oActiveForm.EnableMenu("1284", False)
        oActiveForm.EnableMenu("1286", False)
        oActiveForm.EnableMenu("1283", False) 'MSW 01-04-2011
        oActiveForm.EnableMenu("772", False)
        oActiveForm.EnableMenu("4870", False)
        'If KeyName = "Crane" Or KeyName = "ForkLift" Then
        '    If AddChooseFromList(oActiveForm, "CFLCODEC", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '    If AddChooseFromList(oActiveForm, "CFLNAMEC", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '    If AddChooseFromList(oActiveForm, "CFLSHPAGT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        '    BindingChooseFromList(oActiveForm, "CFLCODEC", "ed_Code", "CardCode")
        '    BindingChooseFromList(oActiveForm, "CFLNAMEC", "ed_Name", "CardName")
        '    BindingChooseFromList(oActiveForm, "CFLSHPAGT", "ed_ShpAgt", "CardName")
        '    oActiveForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        '    oActiveForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
        '    oActiveForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
        '    oActiveForm.Items.Item("ed_JType").Specific.Value = parentForm.Items.Item("ed_JType").Specific.Value
        '    oCombo = oActiveForm.Items.Item("cb_JobType").Specific
        '    oComboMain = parentForm.Items.Item("cb_JobType").Specific
        '    oCombo.ValidValues.Add(parentForm.Items.Item("cb_JobType").Specific.Selected.Value.ToString(), "")
        '    oCombo.Select(0)

        'End If

        oActiveForm.DataBrowser.BrowseBy = "ed_DocNum"
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        jobNo = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Items.Item("ed_JobNo").Specific.Value = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Items.Item("ed_DocID").Specific.Value = GetNewKey(KeyName, oRecordSet)
        oActiveForm.Items.Item("ed_DocNum").Specific.Value = parentForm.Items.Item("ed_DocNum").Specific.Value



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

    Private Sub EnabledTrucker(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)

        ' **********************************************************************************
        '   Function    :   EnabledTrucker()
        '   Purpose     :   This function will be providing to enable items of
        '                   ImportSeaFCL form
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean
        '                  
        '   Return      :   No
        '                  
        ' ***********************************************************************************
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

    Private Sub ValidateJobNumber(ByRef oActiveForm As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim jobType As String = String.Empty
        Dim strJobNum As String = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        oRecordSet.DoQuery(" SELECT U_JOBMODE FROM [@OBT_FCL01_EXPORT] WHERE U_JOBNUM='" & strJobNum & "'")
        If oRecordSet.RecordCount > 0 Then
            jobType = oRecordSet.Fields.Item("U_JOBMODE").Value.ToString.Substring(7, 3)
            If AlreadyExist("EXPORTSEAFCL") Then
                If jobType = "Air" Then
                    oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                    p_oSBOApplication.SetStatusBarMessage("Cannot Load Air Job Number in Export Sea FCL.", SAPbouiCOM.BoMessageTime.bmt_Short)
                    BubbleEvent = False
                End If
            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                If jobType = "Sea" Then
                    oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                    p_oSBOApplication.SetStatusBarMessage("Cannot Load Sea Job Number in Export Air FCL.", SAPbouiCOM.BoMessageTime.bmt_Short)
                    BubbleEvent = False
                End If
            End If
        End If


    End Sub

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

    Private Function AddNewRow(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean

        '=============================================================================
        'Function   : AddNewRow()
        'Purpose    : This function get net new row of matrix  from
        '             ExporeSealCL form and other Purchase Order form 
        'Parameters : ByRef oActiveForm As SAPbouiCOM.Form,
        '             ByVal MatrixUID As String        '            
        'Return     : False- FAILURE
        '             True - SUCCESS
        '==========================================

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
            'End MSW to Edit New Ticket
            AddNewRow = True
        Catch ex As Exception
            AddNewRow = False
        End Try
    End Function

    Private Sub CalRate(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)

        ' **********************************************************************************
        '   Function    :   CalRate()
        '   Purpose     :   This function will be providing to proceed calculate  Rate for
        '                    Purechase order item value for Purchase Order process
        '               
        '   Parameters  :  ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer        '                          
        '   Return      :   No
        '                   
        ' **********************************************************************************

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
            NOGST = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) + GSTAMT 'MSW TO Edit New Ticket
        End If
        oActiveForm.Freeze(True)
        oMatrix.Columns.Item("colGSTAmt").Editable = True
        oMatrix.Columns.Item("colNoGST").Editable = True
        oMatrix.Columns.Item("colGSTAmt").Cells.Item(Row).Specific.Value = Convert.ToString(GSTAMT)
        oMatrix.Columns.Item("colNoGST").Cells.Item(Row).Specific.Value = Convert.ToString(NOGST)
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
        If Right(oActiveForm.UniqueID, 12) = "GOODSRECEIPT" Then
            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_GRTPD", 0, SubTotal)
        Else
            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_POITPD", 0, SubTotal)
        End If

    End Sub

    Private Sub CalAmtPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Try
            Dim cMatrix As SAPbouiCOM.Matrix
            oActiveForm.Freeze(True)
            cMatrix = oActiveForm.Items.Item("mx_Item").Specific
            cMatrix.Columns.Item("colIQty").Editable = True
            cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(cMatrix.Columns.Item("colIPrice").Cells.Item(Row).Specific.Value)
            cMatrix.Columns.Item("colITotal").Editable = True
            cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Specific.Value)
            cMatrix.Columns.Item("colIAmt").Cells.Item(Row).Click()
            cMatrix.Columns.Item("colITotal").Editable = False
            oActiveForm.Freeze(False)
            CalculateTotalPO(oActiveForm, cMatrix)

        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

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

    Private Sub LoadHolidayMarkUp(ByVal ExportSeaFCLForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   LoadHolidayMarkUp()
        '   Purpose     :   This function will be providing to load Holiday Markup fomr for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByVal ExportSeaFCLForm As SAPbouiCOM.Form
        '               
        '   Return      :   No
        '                   
        ' **********************************************************************************

        Dim sErrDesc As String = String.Empty
        If HolidayMarkUpWithoutDay(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ETDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_JbDay").Specific, ExportSeaFCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        'If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_DspDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        'If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaFCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_ADay").Specific, ExportSeaFCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Private Sub AddUpdateVoucher(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)

        ' **********************************************************************************
        '   Function    :   AddUpdateVoucher()
        '   Purpose     :   This function will be providing to Add and Update to Voucher form  
        '                   of ExportSealCL Form
        '               
        '   Parameters  :  ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :  ByVal DataSource As String, ByVal ProcressedState As Boolean          
        '   Return      :  No
        ' **********************************************************************************


        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        If AlreadyExist("EXPORTSEAFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
        ElseIf AlreadyExist("EXPORTAIRFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
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
        ' **********************************************************************************

        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        If AlreadyExist("EXPORTSEAFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
        ElseIf AlreadyExist("EXPORTAIRFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
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
                    ShpInvoice = pForm.Items.Item("ed_ShInvNo").Specific.Value
                    PO = pForm.Items.Item("ed_PO").Specific.Value
                    ShipTo = pForm.Items.Item("ed_ShipTo").Specific.Value
                    Box = pForm.Items.Item("ed_Box").Specific.Value
                    Weight = pForm.Items.Item("ed_Net").Specific.Value
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
                    Try
                        oPurchaseDeliveryNote.CardCode = oActiveForm.Items.Item("ed_Code").Specific.Value
                        oPurchaseDeliveryNote.CardName = oActiveForm.Items.Item("ed_Name").Specific.Value
                        oPurchaseDeliveryNote.ContactPersonCode = GetContactPersonCode(oRecordset, Trim(oActiveForm.Items.Item("cb_Contact").Specific.Value.ToString), oActiveForm.Items.Item("ed_Code").Specific.Value.ToString)
                        oPurchaseDeliveryNote.NumAtCard = oActiveForm.Items.Item("ed_VRef").Specific.Value
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
                                If Not (oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value = "") Then
                                    oPurchaseDeliveryNote.Lines.BaseType = CInt(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                    oPurchaseDeliveryNote.Lines.BaseLine = Convert.ToInt32(oMatrix.Columns.Item("colLineId").Cells.Item(i).Specific.Value)
                                    'p_oSBOApplication.MessageBox(oActiveForm.DataSources.UserDataSources.Item("PONo").Value.ToString())
                                    oPurchaseDeliveryNote.Lines.BaseEntry = Convert.ToInt32(oActiveForm.DataSources.UserDataSources.Item("PONo").Value)
                                    oPurchaseDeliveryNote.Lines.ItemCode = oMatrix.Columns.Item("colItemNo").Cells.Item(i).Specific.Value
                                    oPurchaseDeliveryNote.Lines.ItemDescription = oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value
                                    oPurchaseDeliveryNote.Lines.Quantity = oMatrix.Columns.Item("colIQty").Cells.Item(i).Specific.Value
                                    oPurchaseDeliveryNote.Lines.UnitPrice = Convert.ToDouble(oMatrix.Columns.Item("colIPrice").Cells.Item(i).Specific.Value)
                                    If oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value = "None" Then
                                        oPurchaseDeliveryNote.Lines.VatGroup = "ZI"
                                    Else
                                        oPurchaseDeliveryNote.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                                    End If
                                    oPurchaseDeliveryNote.Lines.Add()
                                End If
                            Next
                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString())
                    End Try

                End If
                Dim ret As Long = oPurchaseDeliveryNote.Add
                If ret <> 0 Then
                    p_oDICompany.GetLastError(ret, sErrDesc)
                    MessageBox.Show("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
                Else
                    oRecordset.DoQuery("SELECT DocStatus FROM OPOR Where DocEntry=" + FormatString(oActiveForm.DataSources.UserDataSources.Item("PONo").Value))
                    If oRecordset.RecordCount > 0 Then
                        DocStatus = IIf(oRecordset.Fields.Item("DocStatus").Value = "C", "Closed", "Open")
                        oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRecordset.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_POStatus = " + FormatString(DocStatus) + " WHERE U_PONo = " + FormatString(oActiveForm.DataSources.UserDataSources.Item("PONo").Value))
                    End If
                End If

               
            End If
            CreateGoodsReceiptPO = True
        Catch ex As Exception
            CreateGoodsReceiptPO = False
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
                    oPurchaseDocument.ContactPersonCode = GetContactPersonCode(oRecordset, Trim(oActiveForm.Items.Item("cb_Contact").Specific.Value.ToString), oActiveForm.Items.Item("ed_Code").Specific.Value.ToString)
                    If oBusinessPartner.Currency = "##" Then
                        oPurchaseDocument.DocCurrency = "SGD"
                    Else
                        oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                    End If

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
                            oPurchaseDocument.Lines.RowTotalFC = Convert.ToDouble(oMatrix.Columns.Item("colITotal").Cells.Item(i).Specific.Value)
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

    Public Sub PreviewPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef oActiveForm As SAPbouiCOM.Form)

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
        'pdffilepath = "C:\Users\UNIQUE\Desktop\For NL\Outrider.pdf"

        rptDocument.SetParameterValue("@DocEntry", PONo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If

    End Sub


    Public Sub PreviewPOInViewList(ByRef ParentForm As SAPbouiCOM.Form, ByVal index As Integer, ByVal oMatrix As SAPbouiCOM.Matrix)

        ' **********************************************************************************
        '   Function    :   PreviewPOInViewList()
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
        PONo = Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(index).Specific.Value())
        rptDocument.SetParameterValue("@DocEntry", PONo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If

    End Sub

    Private Sub PreviewPOFromPOTab(ByRef ParentForm As SAPbouiCOM.Form, ByRef PONo As Integer)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************


        rptDocument = New ReportDocument
        pdfFilename = "PURCHASE ORDER"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Purchase Order.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        rptDocument.SetParameterValue("@DocEntry", PONo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If

    End Sub

    Public Function CreatePOPDF(ByVal oActiveForm As SAPbouiCOM.Form, ByVal matrixName As String, ByVal status As String, Optional ByVal index As Integer = 0) As Boolean

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


        Dim sErrDesc As String = ""
        Dim i As Integer = 0
        Dim dblPrice As Double = 0.0
        Dim itemCode As String = String.Empty
        Dim itemDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oForm As SAPbouiCOM.Form
        Dim originpdf As String = String.Empty
        Dim strDate As String
        Dim UserTel As String = ""
        CreatePOPDF = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If matrixName = "mx_Armed" Then
                pdfFilename = "ArmEdit"
                originpdf = "ArmEdit.pdf"
            ElseIf matrixName = "mx_Outer" Then 'to combine
                pdfFilename = "Outrider"
                originpdf = "Aetos Outrider Escort.pdf"
            ElseIf matrixName = "mx_Bunk" Then 'to combine
                pdfFilename = "Bunker"
                originpdf = "Toll Bunker.pdf"
            ElseIf matrixName = "mx_Toll" Then 'to combine
                pdfFilename = "Toll"
                originpdf = "Toll.pdf"
            ElseIf matrixName = "mx_COO" Then
                pdfFilename = "COO"
                originpdf = "Certificate.pdf"
            ElseIf matrixName = "mx_DGLP" Then
                pdfFilename = "DGLP"
                originpdf = "DGLP_0002.pdf"
            Else

            End If
            mainFolder = p_fmsSetting.DocuPath
            'jobNo = jobNo 'oActiveForm.Items.Item("ed_JobNo").Specific.Value (Must be change Job No)
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            Dim pdfoutputfile As FileStream = New FileStream(pdffilepath, System.IO.FileMode.Create)
            Dim formfiller As iTextSharp.text.pdf.PdfStamper = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            Dim ac As iTextSharp.text.pdf.AcroFields = formfiller.AcroFields
            If AlreadyExist("EXPORTSEAFCL") Then
                oForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
            ElseIf AlreadyExist("EXPORTAIRFCL") Then
                oForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
            End If


            If matrixName = "mx_Armed" Then
                'ac.SetField("txtName", oActiveForm.Items.Item("cb_SInA").Specific.value)
                'ac.SetField("txtcontact", oActiveForm.Items.Item("ed_CNo").Specific.value)
                'ac.SetField("txtReqDate", Today.Date.ToString("dd/mm/yyyy")) 'Convert.ToDateTime(oActiveForm.Items.Item("ed_TDate").Specific.value.ToString()).ToString("dd/mm/yyyy"))
                'ac.SetField("txtFromTo", oActiveForm.Items.Item("ed_Dest").Specific.value & "To" & "Yangon")
                'ac.SetField("Date", Today.Date.Date.ToString())
                'ac.SetField("Authorisation_No", "")
                'ac.SetField("Company_Name_Owner_of_Vehicle", "Midwest Freight & Transportation Pte Ltd")
                'ac.SetField("Vehicle_Registration_No", "")
                'ac.SetField("Vehicle_Registration_No", "")
                'ac.SetField("Vehicle_Registration_No", "")
                'ac.SetField("Reserve_Vehicle_Registration_No", "")
                'ac.SetField("Trailers_LL_Regist_No._1", "")
                'ac.SetField("Trailers_LL_Regist_No._2", "")
                'ac.SetField("Trailers_LL_Regist_No._3", "")
                'ac.SetField("Width", "")
                'ac.SetField("load_from_Vehicle", "")
                'ac.SetField("NRIC_No", "")

            ElseIf matrixName = "mx_Outer" Then
                If status = "View" Then
                    oMatrix = oActiveForm.Items.Item("mx_Outer").Specific
                    If oMatrix.Columns.Item("colPO").Cells.Item(index).Specific.Value() <> "" And index <> 0 Then
                        ac.SetField("txtCompanyName", "Midwest Freight & Transportation Pte Ltd")
                        ac.SetField("txtApplicantName", oMatrix.Columns.Item("colSIA").Cells.Item(index).Specific.Value())
                        ac.SetField("txtContactNo", oMatrix.Columns.Item("colCPerson").Cells.Item(index).Specific.Value())
                        ac.SetField("txtSerialNo", "")
                        ac.SetField("txtPermitNo", "")
                        ac.SetField("txtHandPhone", oMatrix.Columns.Item("colTelNo").Cells.Item(index).Specific.Value())
                        ac.SetField("txtVehicleNo", "")
                        ac.SetField("txtTaxDate", "")
                        ac.SetField("txtLength", "")
                        ac.SetField("txtWidth", "")
                        ac.SetField("txtHeight", "")
                        ac.SetField("txtFaxNo", "")
                        ac.SetField("txtInsuranceName", "")
                        ac.SetField("txtIPolicyNo", "")
                        ac.SetField("dtpTaxExpiry", "")
                        ac.SetField("txtUnit", "")
                        ac.SetField("txtEscortFrom", oMatrix.Columns.Item("colLocFrom").Cells.Item(index).Specific.Value())
                        ac.SetField("txtEscortTo", oMatrix.Columns.Item("colLocTo").Cells.Item(index).Specific.Value())
                        strDate = IIf(oMatrix.Columns.Item("colDate").Cells.Item(index).Specific.Value().ToString() = "", Today.Date.ToString("yyyyMMdd"), oMatrix.Columns.Item("colDate").Cells.Item(index).Specific.Value().ToString())
                        ac.SetField("dtpEscort", strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4))
                        ac.SetField("txtTimeEscort", oMatrix.Columns.Item("colTime").Cells.Item(index).Specific.Value())
                    End If
                Else
                    If oActiveForm.Items.Item("ed_PO").Specific.Value <> "" Then
                        ac.SetField("txtCompanyName", "Midwest Freight & Transportation Pte Ltd")
                        ac.SetField("txtApplicantName", oActiveForm.Items.Item("ed_FSIA").Specific.value)
                        ac.SetField("txtContactNo", oActiveForm.Items.Item("ed_FCntact").Specific.value)
                        ac.SetField("txtSerialNo", "")
                        ac.SetField("txtPermitNo", "")
                        ac.SetField("txtHandPhone", oActiveForm.Items.Item("ed_SIATel").Specific.value)
                        ac.SetField("txtVehicleNo", "")
                        ac.SetField("txtTaxDate", "")
                        ac.SetField("txtLength", "")
                        ac.SetField("txtWidth", "")
                        ac.SetField("txtHeight", "")
                        ac.SetField("txtFaxNo", "")
                        ac.SetField("txtInsuranceName", "")
                        ac.SetField("txtIPolicyNo", "")
                        ac.SetField("dtpTaxExpiry", "")
                        ac.SetField("txtUnit", "")
                        ac.SetField("txtEscortFrom", oActiveForm.Items.Item("ed_LocFrom").Specific.value)
                        ac.SetField("txtEscortTo", oActiveForm.Items.Item("ed_LocTo").Specific.value)
                        strDate = IIf(oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString = "", Today.Date.ToString("yyyyMMdd"), oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString)
                        ac.SetField("dtpEscort", strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4))
                        ac.SetField("txtTimeEscort", oActiveForm.Items.Item("ed_FJbTime").Specific.value)
                    End If
                End If


            ElseIf matrixName = "mx_Bunk" Then

                If status = "Edit" Then
                    If oActiveForm.Items.Item("ed_PO").Specific.Value <> "" Then
                        ac.SetField("txtDate", Today.Date.Date.ToString("dd-MM-yyyy"))
                        ac.SetField("txtCargo", oActiveForm.Items.Item("ed_CDesc").Specific.Value)
                        ac.SetField("txtCPerson", oActiveForm.Items.Item("ed_FSIA").Specific.Value)
                        ac.SetField("txtTel", oActiveForm.Items.Item("ed_SIATel").Specific.Value)
                        strDate = IIf(oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString = "", Today.Date.ToString("yyyyMMdd"), oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString)
                        ac.SetField("txtJDate", strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4))
                        ac.SetField("txtJTime", oActiveForm.Items.Item("ed_FJbTime").Specific.Value)
                        ac.SetField("txtSInstr", oActiveForm.Items.Item("ed_PO").Specific.Value)
                        If oActiveForm.Items.Item("cb_Act").Specific.Value.ToString.Trim = "Delivery" Then
                            ac.SetField("chkDelivery", 1)
                        ElseIf oActiveForm.Items.Item("cb_Act").Specific.Value.ToString.Trim = "Collection" Then
                            ac.SetField("chkCollection", 1)
                        ElseIf oActiveForm.Items.Item("cb_Act").Specific.Value.ToString.Trim = "Prepare for Shipment" Then
                            ac.SetField("chkPrep", 1)
                        ElseIf oActiveForm.Items.Item("cb_Act").Specific.Value.ToString.Trim = "Auditing" Then
                            ac.SetField("chkAudit", 1)
                        End If
                        If oActiveForm.Items.Item("chk_1").Specific.Checked = True Then
                            ac.SetField("chk_1", 1)
                        End If
                        If oActiveForm.Items.Item("chk_1a").Specific.Checked = True Then
                            ac.SetField("chk_1a", 1)
                        End If
                        If oActiveForm.Items.Item("chk_2").Specific.Checked = True Then
                            ac.SetField("chk_2", 1)
                        End If
                        If oActiveForm.Items.Item("chk_2a").Specific.Checked = True Then
                            ac.SetField("chk_2a", 1)
                        End If

                        ac.SetField("txtPODate", oActiveForm.Items.Item("ed_FPODate").Specific.Value.Substring(6, 2) + "-" + oActiveForm.Items.Item("ed_FPODate").Specific.Value.Substring(4, 2) + "-" + oActiveForm.Items.Item("ed_FPODate").Specific.Value.Substring(0, 4))
                        oRecordSet.DoQuery("select PortNum from ousr where USER_CODE='" & oActiveForm.Items.Item("ed_Create").Specific.Value & "'")
                        If oRecordSet.RecordCount > 0 Then
                            UserTel = IIf(oRecordSet.Fields.Item("PortNum").Value.ToString = "", "", oRecordSet.Fields.Item("PortNum").Value.ToString)
                        End If
                        ac.SetField("txtUser", oActiveForm.Items.Item("ed_Create").Specific.Value & " " & UserTel)
                        oMatrix = oActiveForm.Items.Item("mx_BDetail").Specific
                        For j As Integer = 1 To oMatrix.RowCount
                            ac.SetField("txtPermit" & j, oMatrix.Columns.Item("colPermit").Cells.Item(j).Specific.Value())
                            ac.SetField("txtJob" & j, oMatrix.Columns.Item("colJobNo").Cells.Item(j).Specific.Value())
                            ac.SetField("txtQty" & j, Left(oMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value(), oMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value.ToString.Length - 2))
                            ac.SetField("txtUOM" & j, oMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value())
                            ac.SetField("txtKgs" & j, Left(oMatrix.Columns.Item("colKgs").Cells.Item(j).Specific.Value(), oMatrix.Columns.Item("colKgs").Cells.Item(j).Specific.Value.ToString.Length - 2))
                            ac.SetField("txtM3" & j, Left(oMatrix.Columns.Item("colM3").Cells.Item(j).Specific.Value(), oMatrix.Columns.Item("colM3").Cells.Item(j).Specific.Value.ToString.Length - 2))
                            ac.SetField("txtNEQ" & j, Left(oMatrix.Columns.Item("colNEQ").Cells.Item(j).Specific.Value(), oMatrix.Columns.Item("colNEQ").Cells.Item(j).Specific.Value.ToString.Length - 2))
                        Next

                        ac.SetField("txtTQty", Left(oActiveForm.Items.Item("ed_TQty").Specific.Value, oActiveForm.Items.Item("ed_TQty").Specific.Value.ToString.Length - 2))
                        ac.SetField("txtTKgs", Left(oActiveForm.Items.Item("ed_TKgs").Specific.Value, oActiveForm.Items.Item("ed_TKgs").Specific.Value.ToString.Length - 2))
                        ac.SetField("txtTM3", Left(oActiveForm.Items.Item("ed_TM3").Specific.Value, oActiveForm.Items.Item("ed_TM3").Specific.Value.ToString.Length - 2))
                        ac.SetField("txtTNEQ", Left(oActiveForm.Items.Item("ed_TNEQ").Specific.Value, oActiveForm.Items.Item("ed_TNEQ").Specific.Value.ToString.Length - 2))

                    End If

                ElseIf status = "View" Then
                    oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                    If oMatrix.Columns.Item("colPO").Cells.Item(index).Specific.Value() <> "" And index <> 0 Then
                        ac.SetField("txtDate", Today.Date.Date.ToString("dd-MM-yyyy"))
                        ac.SetField("txtCargo", oMatrix.Columns.Item("colCDesc").Cells.Item(index).Specific.Value())
                        ac.SetField("txtCPerson", oMatrix.Columns.Item("colCPerson").Cells.Item(index).Specific.Value())
                        ac.SetField("txtTel", oMatrix.Columns.Item("colTelNo").Cells.Item(index).Specific.Value())
                        strDate = IIf(oMatrix.Columns.Item("colDate").Cells.Item(index).Specific.Value = "", Today.Date.ToString("yyyyMMdd"), oMatrix.Columns.Item("colDate").Cells.Item(index).Specific.Value.ToString)
                        ac.SetField("txtJDate", strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4))
                        ac.SetField("txtJTime", oMatrix.Columns.Item("colTime").Cells.Item(index).Specific.Value())
                        ac.SetField("txtSInstr", oMatrix.Columns.Item("colPO").Cells.Item(index).Specific.Value())
                        If oMatrix.Columns.Item("colActive").Cells.Item(index).Specific.Value().ToString.Trim = "Delivery" Then
                            ac.SetField("chkDelivery", 1)
                        ElseIf oMatrix.Columns.Item("colActive").Cells.Item(index).Specific.Value().ToString.Trim = "Collection" Then
                            ac.SetField("chkCollection", 1)
                        ElseIf oMatrix.Columns.Item("colActive").Cells.Item(index).Specific.Value().ToString.Trim = "Prepare for Shipment" Then
                            ac.SetField("chkPrep", 1)
                        ElseIf oMatrix.Columns.Item("colActive").Cells.Item(index).Specific.Value().ToString.Trim = "Auditing" Then
                            ac.SetField("chkAudit", 1)
                        End If
                        If oMatrix.Columns.Item("colS1").Cells.Item(index).Specific.Value() = "Y" Then
                            ac.SetField("chkStore1", 1)
                        End If
                        If oMatrix.Columns.Item("colS1a").Cells.Item(index).Specific.Value() = "Y" Then
                            ac.SetField("chkStore1a", 1)
                        End If
                        If oMatrix.Columns.Item("colS2").Cells.Item(index).Specific.Value() = "Y" Then
                            ac.SetField("chkStore2", 1)
                        End If
                        If oMatrix.Columns.Item("colS2a").Cells.Item(index).Specific.Value() = "Y" Then
                            ac.SetField("chkStore2a", 1)
                        End If
                        ac.SetField("txtPODate", oMatrix.Columns.Item("colPODate").Cells.Item(index).Specific.Value().Substring(6, 2) + "-" + oMatrix.Columns.Item("colPODate").Cells.Item(index).Specific.Value().Substring(4, 2) + "-" + oMatrix.Columns.Item("colPODate").Cells.Item(index).Specific.Value().Substring(0, 4))
                        oRecordSet.DoQuery("select PortNum from ousr where USER_CODE='" & oMatrix.Columns.Item("colPrepare").Cells.Item(index).Specific.Value() & "'")
                        If oRecordSet.RecordCount > 0 Then
                            UserTel = IIf(oRecordSet.Fields.Item("PortNum").Value.ToString = "", "", oRecordSet.Fields.Item("PortNum").Value.ToString)
                        End If
                        ac.SetField("txtUser", oMatrix.Columns.Item("colPrepare").Cells.Item(index).Specific.Value() & " " & UserTel)

                        ac.SetField("txtTQty", Left(oMatrix.Columns.Item("colTQty").Cells.Item(index).Specific.Value(), oMatrix.Columns.Item("colTQty").Cells.Item(index).Specific.Value().ToString.Length - 2))
                        ac.SetField("txtTKgs", Left(oMatrix.Columns.Item("colTKgs").Cells.Item(index).Specific.Value(), oMatrix.Columns.Item("colTKgs").Cells.Item(index).Specific.Value().ToString.Length - 2))
                        ac.SetField("txtTM3", Left(oMatrix.Columns.Item("colTM3").Cells.Item(index).Specific.Value(), oMatrix.Columns.Item("colTM3").Cells.Item(index).Specific.Value().ToString.Length - 2))
                        ac.SetField("txtTNEQ", Left(oMatrix.Columns.Item("colTNEQ").Cells.Item(index).Specific.Value(), oMatrix.Columns.Item("colTNEQ").Cells.Item(index).Specific.Value().ToString.Length - 2))
                        sql = "Select * from [@OBT_TB01_BNKDETAIL] where U_PO='" & oMatrix.Columns.Item("colPO").Cells.Item(index).Specific.Value() & "' and U_DocN='" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & "'"
                        oRecordSet.DoQuery(sql)
                        If oRecordSet.RecordCount > 0 Then
                            oRecordSet.MoveFirst()
                            i = 1
                            While oRecordSet.EoF = False
                                Dim str As String = oRecordSet.Fields.Item("U_Qty").Value.ToString
                                ac.SetField("txtPermit" & i, oRecordSet.Fields.Item("U_Permit").Value)
                                ac.SetField("txtJob" & i, oRecordSet.Fields.Item("U_JobNo").Value)
                                ac.SetField("txtQty" & i, oRecordSet.Fields.Item("U_Qty").Value)
                                ac.SetField("txtUOM" & i, oRecordSet.Fields.Item("U_UOM").Value)
                                ac.SetField("txtKgs" & i, oRecordSet.Fields.Item("U_Kgs").Value)
                                ac.SetField("txtM3" & i, oRecordSet.Fields.Item("U_M3").Value)
                                ac.SetField("txtNEQ" & i, oRecordSet.Fields.Item("U_NEQ").Value)
                                oRecordSet.MoveNext()
                                i = i + 1
                            End While
                        End If
                    End If

                End If

            ElseIf matrixName = "mx_Toll" Then
                oMatrix = oActiveForm.Items.Item("mx_Toll").Specific
                If status = "Edit" Then
                    If oActiveForm.Items.Item("ed_PO").Specific.Value <> "" Then
                        ac.SetField("txtClient", "MIDWEST FREIGHT & TRANSPORTATION PTE LTD")
                        ac.SetField("txtUser", oActiveForm.Items.Item("ed_Create").Specific.Value)
                        ac.SetField("txtLoc", oActiveForm.Items.Item("ed_Loc").Specific.value)
                        ac.SetField("txtPO", oActiveForm.Items.Item("ed_PO").Specific.value)
                        ac.SetField("txtSIATel", oActiveForm.Items.Item("ed_SIATel").Specific.value)
                        strDate = IIf(oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString = "", Today.Date.ToString("yyyyMMdd"), oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString)
                        ac.SetField("txtDate", strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4))
                        ac.SetField("txtRemark", oActiveForm.Items.Item("ed_Remark").Specific.value)
                    End If
                ElseIf status = "View" Then
                    If oMatrix.Columns.Item("colPO").Cells.Item(index).Specific.Value() <> "" And index <> 0 Then
                        ac.SetField("txtClient", "MIDWEST FREIGHT & TRANSPORTATION PTE LTD")
                        ac.SetField("txtUser", oMatrix.Columns.Item("colPrepare").Cells.Item(index).Specific.Value)
                        ac.SetField("txtLoc", oMatrix.Columns.Item("colLoc").Cells.Item(index).Specific.Value)
                        ac.SetField("txtPO", oMatrix.Columns.Item("colPO").Cells.Item(index).Specific.Value)
                        ac.SetField("txtSIATel", oMatrix.Columns.Item("colTelNo").Cells.Item(index).Specific.Value)
                        strDate = IIf(oMatrix.Columns.Item("colDate").Cells.Item(index).Specific.Value = "", Today.Date.ToString("yyyyMMdd"), oMatrix.Columns.Item("colDate").Cells.Item(index).Specific.Value.ToString)
                        ac.SetField("txtDate", strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4))
                        ac.SetField("txtRemark", oMatrix.Columns.Item("colRemark").Cells.Item(index).Specific.Value)
                    End If

                End If


            ElseIf matrixName = "mx_COO" Then
                ac.SetField("Departure Date", Today.Date.Date.ToString("dd/mm/yyyy"))
                ac.SetField("Vessel Name/Flight No", oForm.Items.Item("ed_Vessel").Specific.value) 'oActiveForm.Items.Item("ed_TDate").Specific.value)
                ac.SetField("Port of Discharge", "") ' oActiveForm.Items.Item("ed_TTime").Specific.value)
                ac.SetField("Country of Final Destination", "")
                ac.SetField("Country of Origin of Goods", "")
                ac.SetField("No", "")
                ac.SetField("Country of Origin of Goods", "")
                ac.SetField("Name", "")
                ac.SetField("Date", "")
                ac.SetField("Marks & Numbers", "")
                ac.SetField("PD of goods", "")
                ac.SetField("Exporter", "")
                ac.SetField("Quantity & Unit", oForm.Items.Item("ed_NOP").Specific.value & oForm.Items.Item("cb_PType").Specific.value)
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("select CardName ,Address,Country   from ocrd where CardCode= '" & oActiveForm.Items.Item("ed_Code").Specific.value & "'")
                If oRecordSet.RecordCount > 0 Then
                    ac.SetField("Consignee", oRecordSet.Fields.Item("CardName").Value.ToString & "," & oRecordSet.Fields.Item("Address").Value.ToString & "," & oRecordSet.Fields.Item("Country").Value.ToString)
                End If
            ElseIf matrixName = "mx_DGLP" Then
                Dim StrDATEReQ As String
                If oActiveForm.Items.Item("ed_TDate").Specific.value <> "" Then
                    StrDATEReQ = oActiveForm.Items.Item("ed_TDate").Specific.value.ToString()
                    ac.SetField("txtVessel", oForm.Items.Item("ed_Vessel").Specific.value)
                    ac.SetField("txtDATE_REQUIRED", StrDATEReQ.Substring(6, 2) + "/" + StrDATEReQ.Substring(5, 2) + "/" + StrDATEReQ.Substring(0, 4))
                    ac.SetField("txtNOS_OF_PACKAGES", oForm.Items.Item("ed_NOP").Specific.value)
                    ac.SetField("txtTotalGrossWeight", oForm.Items.Item("ed_TotalWt").Specific.value)
                    ac.SetField("txtTOTAL_MSEAUREMElH", oForm.Items.Item("ed_TotalM3").Specific.value)
                    ac.SetField("txtNAME_OF_APPlICANT", oActiveForm.Items.Item("cb_SInA").Specific.value)
                    ac.SetField("txtContact", oActiveForm.Items.Item("ed_CNo").Specific.value)
                    ac.SetField("txtContact", oActiveForm.Items.Item("ed_CNo").Specific.value)
                    ac.SetField("txtNAME_AND_ADDRESS_OF_COMPANY", "Midwest Freight & Transportation Pte Ltd")
                    ac.SetField("txtAPPLICATION_NO", "")
                    ac.SetField("txtBEAM", "")
                    ac.SetField("txtLENGTH", "")
                    ac.SetField("txtFROM", "")
                    ac.SetField("txtlOURS_TO", "")
                    ac.SetField("NETr_EXPLOSIVES_CONTENT_NEQl", "")
                    ac.SetField("txtN/A", "")
                    ac.SetField("TYPE_OF_PACKING", "")
                    ac.SetField("SHIPPlNG_MARKSRow1", "")
                    ac.SetField("txtundefined", "")
                    ac.SetField("txtundefined_2", "")
                    ac.SetField("txtRELATIONSHIP_TO_THE_LlGfITER_VESSEL", "")
                    ac.SetField("undefined_5", "")
                    ac.SetField("NAME_AND_ADDRESS_OF_IMPORTERtxt_EXPORTER_AND_OR_CONSIGNEE", "Midwest Freight & Transportation Pte Ltd")
                End If
            End If
            formfiller.Close()
            reader.Close()
            Process.Start(pdffilepath)
            CreatePOPDF = True
        Catch ex As Exception
            CreatePOPDF = False
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function PopulatePurchaseHeaderFromMain(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean, ByRef oPOForm As SAPbouiCOM.Form) As Boolean
        PopulatePurchaseHeaderFromMain = False
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
                            '.SetValue("U_PONo", .Offset, DocLastKey)
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
            PopulatePurchaseHeaderFromMain = True
        Catch ex As Exception
            PopulatePurchaseHeaderFromMain = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function PopulateOtherPurchaseHeaderFromMain(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String) As Boolean
        PopulateOtherPurchaseHeaderFromMain = False
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
            PopulateOtherPurchaseHeaderFromMain = True
        Catch ex As Exception
            PopulateOtherPurchaseHeaderFromMain = False
            MessageBox.Show(ex.Message)
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
                ' p_oSBOApplication.MessageBox(oBusinessPartner.Currency)
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
                MessageBox.Show("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
            End If
            oRecordset.DoQuery("SELECT DocEntry FROM OPOR ORDER BY DocEntry")
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

    Public Sub LoadExportSeaFCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal JType As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)

        '=============================================================================
        'Function   : LoadExportSeaFCLForm()
        'Purpose    : This function will be providing to call and load ExporeSealLcLForm  
        '             and show  ExporeSealLcLForm for process
        'Parameters : Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString
        '             Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        'Return     : No
        '=============================================================================

        Dim ExportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oMenuItem As SAPbouiCOM.IMenuItem
        Dim oMenus As SAPbouiCOM.IMenus
        Dim jobType As String = String.Empty
        Dim oColumn As SAPbouiCOM.Column
        Dim jMode As String = "" 'Google Doc

        Dim FunctionName As String = "DoExportSeaFCLMenuEvent()"
        Dim sErrDesc As String = String.Empty
        Try
          
            LoadFromXML(p_oSBOApplication, "ExportSeaFCLv1.srf")
            ExportSeaFCLForm = p_oSBOApplication.Forms.Item("EXPORTSEAFCL")
            If Title.Contains("Export") Then
                EnableExportImport(ExportSeaFCLForm, True, False)
            ElseIf Title.Contains("Import") Then
                EnableExportImport(ExportSeaFCLForm, False, True)
            Else
                EnableExportImport(ExportSeaFCLForm, True, True)
            End If

          
                'ExportSeaFCLForm.Title = Title ' MSW To Edit New Ticket 07-09-2011
                ExportSeaFCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                Try
                    ExportSeaFCLForm.EnableMenu("1288", True)
                    ExportSeaFCLForm.EnableMenu("1289", True)
                    ExportSeaFCLForm.EnableMenu("1290", True)
                    ExportSeaFCLForm.EnableMenu("1291", True)
                    ExportSeaFCLForm.EnableMenu("1284", False)
                    ExportSeaFCLForm.EnableMenu("1286", False)
                    ExportSeaFCLForm.EnableMenu("1281", True)
                    ExportSeaFCLForm.EnableMenu("1282", False)
                    ExportSeaFCLForm.EnableMenu("1284", False)
                    ExportSeaFCLForm.EnableMenu("1286", False)
                    ExportSeaFCLForm.EnableMenu("1283", False) 'MSW 01-04-2011
                    ExportSeaFCLForm.EnableMenu("1292", False)
                    ExportSeaFCLForm.EnableMenu("1293", False)
                    ExportSeaFCLForm.EnableMenu("4870", False)
                    ExportSeaFCLForm.EnableMenu("771", False)
                    ExportSeaFCLForm.EnableMenu("772", True)
                    ExportSeaFCLForm.EnableMenu("773", True)
                    ExportSeaFCLForm.EnableMenu("774", False)
                    ExportSeaFCLForm.EnableMenu("775", True)

                    If FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    End If
                    ExportSeaFCLForm.DataBrowser.BrowseBy = "ed_DocNum"
                    ExportSeaFCLForm.Items.Item("fo_Vchr").Specific.Select()
                    ExportSeaFCLForm.Items.Item("bt_Crane").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_Foklit").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_DGLP").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_Excel").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_A6Label").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_AmdCont").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_DelCont").Enabled = False
                    ExportSeaFCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                ExportSeaFCLForm.Items.Item("bt_Fumi").Enabled = False
                ExportSeaFCLForm.Items.Item("bt_Crate").Enabled = False
                ExportSeaFCLForm.Items.Item("bt_Outer").Enabled = False
                ExportSeaFCLForm.Items.Item("bt_Bunk").Enabled = False
                ExportSeaFCLForm.Items.Item("bt_Toll").Enabled = False
                ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = False
                    ExportSeaFCLForm.Freeze(True)
                Catch ex As Exception
                    MessageBox.Show(ex.ToString())
                End Try


                ExportSeaFCLForm.Items.Item("bt_PO").Enabled = False
                EnabledHeaderControls(ExportSeaFCLForm, False)
                EnabledMaxtix(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("mx_TkrList").Specific, False)
                ExportSeaFCLForm.PaneLevel = 20

                If FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    If Not Title = vbNullString Then
                        ExportSeaFCLForm.Items.Item("ed_JType").Specific.Value = Title
                    End If

                    oCombo = ExportSeaFCLForm.Items.Item("cb_TMode").Specific
                If Title.Contains("Air") Then
                    oCombo.Select("Air", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    jMode = "A" 'Google Doc
                ElseIf Title.Contains("Land") Then
                    oCombo.Select("Land", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    jMode = "L" 'Google Doc
                ElseIf Title.Contains("Sea") Then
                    oCombo.Select("Sea", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    jMode = "S" 'Google Doc
                ElseIf Title.Contains("Local") Then
                    oCombo.Select("Land", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    jMode = "L" 'Google Doc
                End If

                    ExportSeaFCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                    ExportSeaFCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                    ExportSeaFCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                    ExportSeaFCLForm.Items.Item("ed_UserID").Specific.Value = p_oDICompany.UserName.ToString
                    ExportSeaFCLForm.Items.Item("ed_Code").Specific.Active = True

                End If



                If HolidayMarkUp(ExportSeaFCLForm, ExportSeaFCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaFCLForm.Items.Item("ed_JbDay").Specific, ExportSeaFCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If AddChooseFromList(ExportSeaFCLForm, "cflBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddChooseFromList(ExportSeaFCLForm, "cflBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

           

                ExportSeaFCLForm.Items.Item("ed_Code").Specific.ChooseFromListUID = "cflBP"
                ExportSeaFCLForm.Items.Item("ed_Code").Specific.ChooseFromListAlias = "CardCode"
                ExportSeaFCLForm.Items.Item("ed_Name").Specific.ChooseFromListUID = "cflBP2"
                ExportSeaFCLForm.Items.Item("ed_Name").Specific.ChooseFromListAlias = "CardName"
            'Google Doc
            If AddChooseFromList(ExportSeaFCLForm, "cflC2BP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(ExportSeaFCLForm, "cflC2BP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ExportSeaFCLForm.Items.Item("ed_C2Code").Specific.ChooseFromListUID = "cflC2BP"
            ExportSeaFCLForm.Items.Item("ed_C2Code").Specific.ChooseFromListAlias = "CardCode"
            ExportSeaFCLForm.Items.Item("ed_Client2").Specific.ChooseFromListUID = "cflC2BP2"
            ExportSeaFCLForm.Items.Item("ed_Client2").Specific.ChooseFromListAlias = "CardName"

                If AddChooseFromList(ExportSeaFCLForm, "cflBP3", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                ExportSeaFCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"

                If AddChooseFromList(ExportSeaFCLForm, "cflBP4", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.ChooseFromListUID = "cflBP4"
                ExportSeaFCLForm.Items.Item("ed_IShpAgt").Specific.ChooseFromListAlias = "CardName"

                ''-------------------------------For Cargo Tab OMM & SYMA------------------------------------------------'13 Jan 2011

                AddChooseFromList(ExportSeaFCLForm, "cflCurCode", False, 37)
                ExportSeaFCLForm.Items.Item("ed_CurCode").Specific.ChooseFromListUID = "cflCurCode"
                ''----------------------------------For Invoice Tab------------------------------------------------------'
                AddChooseFromList(ExportSeaFCLForm, "cflCurCode1", False, 37)
                oEditText = ExportSeaFCLForm.Items.Item("ed_CCharge").Specific
                oEditText.ChooseFromListUID = "cflCurCode1"
                AddChooseFromList(ExportSeaFCLForm, "cflCurCode2", False, 37)
                oEditText = ExportSeaFCLForm.Items.Item("ed_Charge").Specific
                oEditText.ChooseFromListUID = "cflCurCode2"
                ''-------------------------------------------------------------------------------------------------------'

                ''-------------------------------------Charge Code--------------------------------------------------------------'
                If AddChooseFromList(ExportSeaFCLForm, "Charge", False, "UDOCHCODE") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ExportSeaFCLForm.Items.Item("ed_ChCode").Specific.ChooseFromListUID = "Charge"
                ExportSeaFCLForm.Items.Item("ed_ChCode").Specific.ChooseFromListAlias = "U_CName"
                ''--------------------------------------------------------------------------------------------------------------'

                ''===== Other Charges Tab ====='
                oMatrix = ExportSeaFCLForm.Items.Item("mx_Charge").Specific
                oColumn = oMatrix.Columns.Item("colCClaim1")
                oColumn.ValidValues.Add("Yes", "Yes")
                oColumn.ValidValues.Add("No", "No")
                oMatrix.AddRow()
                oCombo = oMatrix.Columns.Item("colCClaim1").Cells.Item(1).Specific
                oCombo.Select("Yes", SAPbouiCOM.BoSearchKey.psk_ByValue)
                oMatrix.Columns.Item("V_-1").Cells.Item(1).Specific.Value = 1
                oColumn = oMatrix.Columns.Item("colCCode1")
                If AddChooseFromList(ExportSeaFCLForm, "CHITEM", False, 4) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                oColumn.ChooseFromListUID = "CHITEM"
                oColumn.ChooseFromListAlias = "ItemCode" 'MSW To Edit New Ticket

                ''===== End Other Charges Tab ====='


                ''===== Container Tab ====='

                oMatrix = ExportSeaFCLForm.Items.Item("mx_ConTab").Specific
                oColumn = oMatrix.Columns.Item("colCType1")
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("select U_ContType  from [@OBT_TB021_CONT] group by U_ContType")
                While oColumn.ValidValues.Count > 0
                    oColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index)
                End While
                If oRecordSet.RecordCount > 0 Then
                    Dim a As Integer = oRecordSet.RecordCount
                    oRecordSet.MoveFirst()
                    While oRecordSet.EoF = False
                        oColumn.ValidValues.Add(oRecordSet.Fields.Item("U_ContType").Value, " ")
                        oColumn.DisplayDesc = False
                        oRecordSet.MoveNext()
                    End While
                End If
                oColumn = oMatrix.Columns.Item("colCSize1")
                If oColumn.ValidValues.Count = 0 Then
                    oColumn.ValidValues.Add("20'", "20'")
                    oColumn.ValidValues.Add("40'", "40'")
                    oColumn.ValidValues.Add("45'", "45'")
                End If
                oMatrix.AddRow()
                oCombo = oMatrix.Columns.Item("colCSize1").Cells.Item(1).Specific
                oCombo.Select("20'", SAPbouiCOM.BoSearchKey.psk_ByValue)
                oMatrix.Columns.Item("V_-1").Cells.Item(1).Specific.Value = 1

            Dim optPC As SAPbouiCOM.OptionBtn
            optPC = ExportSeaFCLForm.Items.Item("op_FrtC").Specific
            optPC = ExportSeaFCLForm.Items.Item("op_FrtP").Specific
            optPC.GroupWith("op_FrtC")

                'If AddUserDataSrc(ExportSeaFCLForm, "ConSeqNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ConNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "SealNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                'If AddUserDataSrc(ExportSeaFCLForm, "ConSize", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ConType", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ContWt", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ConDesc", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ConStuff", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ConDate", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'If AddUserDataSrc(ExportSeaFCLForm, "ConDay", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ''If AddUserDataSrc(oActiveForm, "ConHr", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                'ExportSeaFCLForm.Items.Item("ed_ConNo").Specific.DataBind.SetBound(True, "", "ConSeqNo")
                'ExportSeaFCLForm.Items.Item("ed_ContNo").Specific.DataBind.SetBound(True, "", "ConNo")
                'ExportSeaFCLForm.Items.Item("ed_SealNo").Specific.DataBind.SetBound(True, "", "SealNo")
                'ExportSeaFCLForm.Items.Item("cb_ConSize").Specific.DataBind.SetBound(True, "", "ConSize")
                'ExportSeaFCLForm.Items.Item("cb_ConType").Specific.DataBind.SetBound(True, "", "ConType")
                'ExportSeaFCLForm.Items.Item("ed_ContWt").Specific.DataBind.SetBound(True, "", "ContWt")
                'ExportSeaFCLForm.Items.Item("ed_CDesc").Specific.DataBind.SetBound(True, "", "ConDesc")
                'ExportSeaFCLForm.Items.Item("ch_CStuff").Specific.DataBind.SetBound(True, "", "ConStuff")
                'ExportSeaFCLForm.Items.Item("ed_CunDate").Specific.DataBind.SetBound(True, "", "ConDate")
                'ExportSeaFCLForm.Items.Item("ed_CunDay").Specific.DataBind.SetBound(True, "", "ConDay")
                ''oActiveForm.Items.Item("ed_CunTime").Specific.DataBind.SetBound(True, "", "ConHr")

                'oCombo = ExportSeaFCLForm.Items.Item("cb_ConType").Specific
                'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oRecordSet.DoQuery("select U_ContType  from [@OBT_TB021_CONT] group by U_ContType")
                'If oRecordSet.RecordCount > 0 Then
                '    oRecordSet.MoveFirst()
                '    While oRecordSet.EoF = False
                '        oCombo.ValidValues.Add(oRecordSet.Fields.Item("U_ContType").Value, "")
                '        oRecordSet.MoveNext()
                '    End While

                'End If

                oCombo = ExportSeaFCLForm.Items.Item("cb_PType").Specific
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("SELECT PkgType FROM OPKG")
                If oRecordSet.RecordCount > 0 Then
                    oRecordSet.MoveFirst()
                    While oRecordSet.EoF = False
                        oCombo.ValidValues.Add(oRecordSet.Fields.Item("PkgType").Value, "")
                        oRecordSet.MoveNext()
                    End While
                End If
                oCombo = ExportSeaFCLForm.Items.Item("cb_ESur").Specific
                oCombo.ValidValues.Add("", "")
                oCombo.ValidValues.Add("Yes", "Yes")
                oCombo.ValidValues.Add("No", "No")
                oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)

                oCombo = ExportSeaFCLForm.Items.Item("cb_ISur").Specific
                oCombo.ValidValues.Add("", "")
                oCombo.ValidValues.Add("Yes", "Yes")
                oCombo.ValidValues.Add("No", "No")
                oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                'oCombo = ExportSeaFCLForm.Items.Item("cb_ConSize").Specific
                'oCombo.ValidValues.Add("20'", "20'")
                'oCombo.ValidValues.Add("40'", "40'")
                'oCombo.ValidValues.Add("45'", "45'")

                ''===== End MSW Container ====='



                oEditText = ExportSeaFCLForm.Items.Item("ed_JobNo").Specific
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'EXPORTSEAFCL'")
                If oRecordSet.RecordCount > 0 Then

                    ExportSeaFCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW For JobType Table 

                End If
            If Not ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                'Google Doc
                If JType = "Import" Then
                    ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("I" & jMode.ToString)
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                ElseIf JType = "Export" Then
                    ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("E" & jMode.ToString)
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                ElseIf JType = "Local" Then
                    ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("L" & jMode.ToString)
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                ElseIf JType = "Transhipment" Then
                    ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("TS")
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                End If

            End If



                ''for Dispatch Tab

                'Setting Tab new UI
                AddChooseFromList(ExportSeaFCLForm, "cflDICode", False, 4)
                ExportSeaFCLForm.Items.Item("ed_DICode").Specific.ChooseFromListUID = "cflDICode"
                ExportSeaFCLForm.Items.Item("ed_DICode").Specific.ChooseFromListAlias = "ItemCode"


                If AddUserDataSrc(ExportSeaFCLForm, "DSPINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                oOpt = ExportSeaFCLForm.Items.Item("op_DInter").Specific
                oOpt.DataBind.SetBound(True, "", "DSINTR")
                oOpt = ExportSeaFCLForm.Items.Item("op_DExter").Specific
                oOpt.DataBind.SetBound(True, "", "DSEXTR")
                oOpt.GroupWith("op_DInter")

                If AddUserDataSrc(ExportSeaFCLForm, "DSPDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DINSDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ExportSeaFCLForm.Items.Item("ed_DInDate").Specific.DataBind.SetBound(True, "", "DINSDATE")
                ExportSeaFCLForm.Items.Item("ed_DspDate").Specific.DataBind.SetBound(True, "", "DSPDATE")
                If AddUserDataSrc(ExportSeaFCLForm, "DSPATTE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPTEL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPFAX", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPMAIL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DEUC", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) ' MSW To Edit New Ticket

                ''MSW 14-09-2011 Truck PO
                If AddUserDataSrc(ExportSeaFCLForm, "DSPINS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPIRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "DSPPODOCNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'MSW 14-09-2011 Truck PO

                'New UI
                If AddUserDataSrc(ExportSeaFCLForm, "DSPCODE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DPO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DSTATUS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DCREATEDBY", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DMULTI", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DORIGIN", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                'New UI

                ExportSeaFCLForm.Items.Item("ed_DAttent").Specific.DataBind.SetBound(True, "", "DSPATTE")
                ExportSeaFCLForm.Items.Item("ed_DspTel").Specific.DataBind.SetBound(True, "", "DSPTEL")
                ExportSeaFCLForm.Items.Item("ed_DFax").Specific.DataBind.SetBound(True, "", "DSPFAX")
                ExportSeaFCLForm.Items.Item("ed_DEmail").Specific.DataBind.SetBound(True, "", "DSPMAIL")
                ExportSeaFCLForm.Items.Item("ed_DEUC").Specific.DataBind.SetBound(True, "", "DEUC") ' MSW To Edit New Ticket
                ExportSeaFCLForm.Items.Item("ee_DIRmsk").Specific.DataBind.SetBound(True, "", "DSPIRMK") 'MSW 14-09-2011 Truck PO
                ExportSeaFCLForm.Items.Item("ee_DspIns").Specific.DataBind.SetBound(True, "", "DSPINS") 'MSW 14-09-2011 Truck PO
                ExportSeaFCLForm.Items.Item("ee_DRmsk").Specific.DataBind.SetBound(True, "", "DSPRMK") 'MSW 14-09-2011 Truck PO
                ExportSeaFCLForm.Items.Item("ed_DPDocNo").Specific.DataBind.SetBound(True, "", "DSPPODOCNO") 'MSW 14-09-2011 Truck PO
                'New UI
                ExportSeaFCLForm.Items.Item("ed_DspCode").Specific.DataBind.SetBound(True, "", "DSPCODE")
                ExportSeaFCLForm.Items.Item("ed_DPO").Specific.DataBind.SetBound(True, "", "DPO")
                ExportSeaFCLForm.Items.Item("ed_DPStus").Specific.DataBind.SetBound(True, "", "DSTATUS")
                ExportSeaFCLForm.Items.Item("ed_DCreate").Specific.DataBind.SetBound(True, "", "DCREATEDBY")
                ExportSeaFCLForm.Items.Item("ed_DDate").Specific.DataBind.SetBound(True, "", "DDATE")
                ExportSeaFCLForm.Items.Item("ed_DMulti").Specific.DataBind.SetBound(True, "", "DMULTI")
                ExportSeaFCLForm.Items.Item("ed_DOrigin").Specific.DataBind.SetBound(True, "", "DORIGIN")
                'New UI

                If AddChooseFromList(ExportSeaFCLForm, "CFLDSP", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddChooseFromList(ExportSeaFCLForm, "CFLDSPV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


            'to km
            ExportSeaFCLForm.Items.Item("ed_POICode").Specific.Value = "1101"
            ExportSeaFCLForm.Items.Item("ed_POIDesc").Specific.Value = "Trucking-Local Delivery(Min)"
            ExportSeaFCLForm.Items.Item("ed_POQty").Specific.Value = "1"
            ExportSeaFCLForm.Items.Item("ed_POPrice").Specific.Value = "1"

            'to km
            ExportSeaFCLForm.Items.Item("ed_DICode").Specific.Value = "9104"
            ExportSeaFCLForm.Items.Item("ed_DIDesc").Specific.Value = "Courier Charges"
            ExportSeaFCLForm.Items.Item("ed_DQty").Specific.Value = "1"
            ExportSeaFCLForm.Items.Item("ed_DPrice").Specific.Value = "1"



                ''fortruckingtab

                'Setting Tab new UI
                AddChooseFromList(ExportSeaFCLForm, "cflTICode", False, 4)
                ExportSeaFCLForm.Items.Item("ed_POICode").Specific.ChooseFromListUID = "cflTICode"
                ExportSeaFCLForm.Items.Item("ed_POICode").Specific.ChooseFromListAlias = "ItemCode"

                If AddUserDataSrc(ExportSeaFCLForm, "TKRINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKREXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKINTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKEXTR", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                oOpt = ExportSeaFCLForm.Items.Item("op_Inter").Specific
                oOpt.DataBind.SetBound(True, "", "TKINTR")
                oOpt = ExportSeaFCLForm.Items.Item("op_Exter").Specific
                oOpt.DataBind.SetBound(True, "", "TKEXTR")
                oOpt.GroupWith("op_Inter")

                If AddUserDataSrc(ExportSeaFCLForm, "TKRDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "INSDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                ExportSeaFCLForm.Items.Item("ed_InsDate").Specific.DataBind.SetBound(True, "", "INSDATE")
                ExportSeaFCLForm.Items.Item("ed_TkrDate").Specific.DataBind.SetBound(True, "", "TKRDATE")
                If AddUserDataSrc(ExportSeaFCLForm, "TKRATTE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRTEL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRFAX", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRMAIL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRCOL", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRTO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "EUC", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) ' MSW To Edit New Ticket

                ''MSW 14-09-2011 Truck PO
                If AddUserDataSrc(ExportSeaFCLForm, "TKRINS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRIRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "TKRRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddUserDataSrc(ExportSeaFCLForm, "PODOCNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                'MSW 14-09-2011 Truck PO

                'New UI
                If AddUserDataSrc(ExportSeaFCLForm, "TKRCODE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "PO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "STATUS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "CREATEDBY", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "DATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "TMULTI", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                If AddUserDataSrc(ExportSeaFCLForm, "TORIGIN", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) 'OPOR DocNum (Serial No.)
                'New UI

                ExportSeaFCLForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
                ExportSeaFCLForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
                ExportSeaFCLForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
                ExportSeaFCLForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
                ExportSeaFCLForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
                ExportSeaFCLForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")
                ExportSeaFCLForm.Items.Item("ed_EUC").Specific.DataBind.SetBound(True, "", "EUC") ' MSW To Edit New Ticket
                ExportSeaFCLForm.Items.Item("ee_InsRmsk").Specific.DataBind.SetBound(True, "", "TKRIRMK") 'MSW 14-09-2011 Truck PO
                ExportSeaFCLForm.Items.Item("ee_TkrIns").Specific.DataBind.SetBound(True, "", "TKRINS") 'MSW 14-09-2011 Truck PO
                ExportSeaFCLForm.Items.Item("ee_Rmsk").Specific.DataBind.SetBound(True, "", "TKRRMK") 'MSW 14-09-2011 Truck PO
                ExportSeaFCLForm.Items.Item("ed_PODocNo").Specific.DataBind.SetBound(True, "", "PODOCNO") 'MSW 14-09-2011 Truck PO
                'New UI
                ExportSeaFCLForm.Items.Item("ed_TkrCode").Specific.DataBind.SetBound(True, "", "TKRCODE")
                ExportSeaFCLForm.Items.Item("ed_PO").Specific.DataBind.SetBound(True, "", "PO")
                ExportSeaFCLForm.Items.Item("ed_PStus").Specific.DataBind.SetBound(True, "", "STATUS")
                ExportSeaFCLForm.Items.Item("ed_Created").Specific.DataBind.SetBound(True, "", "CREATEDBY")
                ExportSeaFCLForm.Items.Item("ed_Date").Specific.DataBind.SetBound(True, "", "DATE")
                ExportSeaFCLForm.Items.Item("ed_TMulti").Specific.DataBind.SetBound(True, "", "TMULTI")
                ExportSeaFCLForm.Items.Item("ed_TOrigin").Specific.DataBind.SetBound(True, "", "TORIGIN")
                'New UI



                If AddChooseFromList(ExportSeaFCLForm, "CFLTKRE", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                If AddChooseFromList(ExportSeaFCLForm, "CFLTKRV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                ''---------------------------10-1-2011-------------------------------------
                ''----------Recordset for Binding colCType of Matrix (mx_Cont)-------------
                ''--------------------------SYMA & OMM-------------------------------------
                oMatrix = ExportSeaFCLForm.Items.Item("mx_Cont").Specific
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
                oMatrix = ExportSeaFCLForm.Items.Item("mx_License").Specific
                oMatrix.AddRow()
                oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
                ''-------------------------------------------------------------------------------------'


            If ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                Dim tempItem As SAPbouiCOM.Item
                tempItem = ExportSeaFCLForm.Items.Item("ed_JobNo")
                tempItem.Enabled = True
                ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value = JobNo
                ExportSeaFCLForm.Items.Item("1").Click()
                'Google Doc
                If JType = "Import" Then
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                ElseIf JType = "Export" Then
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                ElseIf JType = "Local" Then
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                ElseIf JType = "Transhipment" Then
                    ExportSeaFCLForm.Title = Title + " - " + ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Substring(0, ExportSeaFCLForm.Items.Item("ed_JobNo").Specific.Value.ToString.Length - 3)
                End If
                tempItem.Enabled = False

                ExportSeaFCLForm.Items.Item("ed_Client2").Specific.Active = True
            End If

                'If ExportSeaFCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                '    ExportSeaFCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                'ElseIf ExportSeaFCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                '    ExportSeaFCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                'Else
                '    ExportSeaFCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                'End If
                ExportSeaFCLForm.Freeze(False)

                Select Case ExportSeaFCLForm.Mode
                    Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                        ExportSeaFCLForm.Items.Item("bt_AmdIns").Enabled = False
                        ExportSeaFCLForm.Items.Item("bt_AddIns").Enabled = False
                        ExportSeaFCLForm.Items.Item("bt_DelIns").Enabled = False
                        ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = False
                        ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                    Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        ExportSeaFCLForm.Items.Item("bt_AmdIns").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_AddIns").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_DelIns").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_PrntIns").Enabled = True

                        ExportSeaFCLForm.Items.Item("bt_BokPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_DrBL").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_SpOrder").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_DGD").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_CrPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_CPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_BunkPO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_ArmePO").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_AmdVoc").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_ShpInv").Enabled = True
                        ExportSeaFCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011

                End Select
        Catch ex As Exception

        End Try
    End Sub

    Private Sub EnabledMaxtix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)

        ' **********************************************************************************
        '   Function    :   EnabledMaxtix()
        '   Purpose     :   This function will be providing to clear items  data for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal pValue As Boolean
        '               
        '   Return      :   No
        '*************************************************************
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

    Private Sub EnabledHeaderControls(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)


        ' **********************************************************************************
        '   Function    :   EnabledHeaderControls()
        '   Purpose     :   This function will be providing to enable header control in
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean
        '   Return      :   No

        '*************************************************************


        pForm.Items.Item("ed_Client2").Enabled = pValue
        pForm.Items.Item("ed_C2Code").Enabled = pValue 'Google Doc
        pForm.Items.Item("ed_CrgDsc").Enabled = pValue
        pForm.Items.Item("ed_TotalM3").Enabled = pValue
        pForm.Items.Item("ed_TotalWt").Enabled = pValue
        pForm.Items.Item("ed_NOP").Enabled = pValue
        pForm.Items.Item("cb_PType").Enabled = pValue
        pForm.Items.Item("ed_L").Enabled = pValue
        pForm.Items.Item("ed_B").Enabled = pValue
        pForm.Items.Item("ed_H").Enabled = pValue
        pForm.Items.Item("ed_JDesc").Enabled = pValue
        pForm.Items.Item("ed_ETADate").Enabled = pValue
        pForm.Items.Item("ed_ETAHr").Enabled = pValue
        pForm.Items.Item("ed_IVsl1").Enabled = pValue
        pForm.Items.Item("ed_IFrom1").Enabled = pValue
        pForm.Items.Item("ed_ITo1").Enabled = pValue
        pForm.Items.Item("ed_IDate2").Enabled = pValue
        pForm.Items.Item("ed_IETA2").Enabled = pValue
        pForm.Items.Item("ed_IVsl2").Enabled = pValue
        pForm.Items.Item("ed_IFrom2").Enabled = pValue
        pForm.Items.Item("ed_ITo2").Enabled = pValue
        pForm.Items.Item("ed_IShpAgt").Enabled = pValue
        pForm.Items.Item("ed_IOBL").Enabled = pValue
        pForm.Items.Item("cb_ISur").Enabled = pValue
        pForm.Items.Item("ed_IHBL").Enabled = pValue
        pForm.Items.Item("ed_ICutOff").Enabled = pValue
        pForm.Items.Item("ed_ETDDate").Enabled = pValue
        pForm.Items.Item("ed_ETDHr").Enabled = pValue
        pForm.Items.Item("ed_Vessel").Enabled = pValue
        pForm.Items.Item("ed_EFrom1").Enabled = pValue
        pForm.Items.Item("ed_ETo1").Enabled = pValue
        pForm.Items.Item("ed_EDate2").Enabled = pValue
        pForm.Items.Item("ed_EETA2").Enabled = pValue
        pForm.Items.Item("ed_EVsl2").Enabled = pValue
        pForm.Items.Item("ed_EFrom2").Enabled = pValue
        pForm.Items.Item("ed_ETo2").Enabled = pValue
        pForm.Items.Item("ed_ShpAgt").Enabled = pValue
        pForm.Items.Item("ed_OBL").Enabled = pValue
        pForm.Items.Item("cb_ESur").Enabled = pValue
        pForm.Items.Item("ed_HBL").Enabled = pValue
        pForm.Items.Item("ed_ECutOff").Enabled = pValue
        pForm.Items.Item("ed_BokRef").Enabled = pValue
        pForm.Items.Item("ed_ICtOffT").Enabled = pValue
        pForm.Items.Item("ed_ECtOffT").Enabled = pValue
        pForm.Items.Item("ed_IBokRef").Enabled = pValue
        pForm.Items.Item("ch_POD").Enabled = pValue
        pForm.Items.Item("ed_xRef").Enabled = pValue
        pForm.Items.Item("ed_BokRef").Enabled = pValue
        pForm.Items.Item("chk_Frt").Enabled = pValue
        pForm.Items.Item("chk_BG").Enabled = pValue
        pForm.Items.Item("chk_Stuff").Enabled = pValue
        pForm.Items.Item("chk_UStuff").Enabled = pValue
        pForm.Items.Item("chk_Cust").Enabled = pValue
        pForm.Items.Item("ed_CustRef").Enabled = pValue
        pForm.Items.Item("cb_JbStus").Enabled = pValue
        pForm.Items.Item("op_FrtC").Enabled = pValue
        pForm.Items.Item("op_FrtP").Enabled = pValue


    End Sub

    Private Sub DeleteByIndexOtherCharges(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String)

        ' **********************************************************************************
        '   Function    :   DeleteByIndexOtherCharges()
        '   Purpose     :   This function will be providing to Delete index 0ther charges form  for
        '                   ImportSeaFCL Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal DataSource As String
        '   Return      :   No
        '               
        '                   
        ' **********************************************************************************
        Try
            Dim lRow As Long
            lRow = pMatrix.GetNextSelectedRow
            If lRow > -1 Then
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    Try
                        pForm.DataSources.DBDataSources.Item("@OBT_FCL17_HCHARGES").RemoveRecord(lRow - 1)
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

        ' **********************************************************************************
        '   Function    :   SetOtherChargesDataToEditTabByIndex()
        '   Purpose     :   This function will be providing to ChargesData to editTab by Index form  for
        '                   ImportSeaFCL Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal Index As Integer
        '   Return      :   No
        '               
        '                   
        ' **********************************************************************************
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

    Public Sub LoadData()
        Dim DirList As New ArrayList

        GetDirectories(p_fmsSetting.PicturePath, DirList)
        DirList.Sort()

        For Each item In DirList
            FindSubFolders(item.ToString())
        Next
    End Sub

    Sub GetDirectories(ByVal StartPath As String, ByRef DirectoryList As ArrayList)
        Dim Dirs() As String = Directory.GetDirectories(StartPath)
        DirectoryList.AddRange(Dirs)

        For Each Dir As String In Dirs
            GetDirectories(Dir, DirectoryList)
        Next
    End Sub

    Private Sub FindSubFolders(ByRef Path As String)

        Dim di As New IO.DirectoryInfo(Path)
        Try
            di.GetFiles("*.jpg*", SearchOption.AllDirectories)
        Catch
        End Try
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.jpg*")
        Dim fi As IO.FileInfo
        For Each fi In aryFi
            Try
                Dim oPictureBox As SAPbouiCOM.PictureBox = ActiveForm.Items.Item((fi.Name.Substring(0, fi.Name.Length - 4))).Specific
                oPictureBox.Picture = fi.FullName
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Next
    End Sub

    Private Sub AddGSTComboData(ByVal oColumn As SAPbouiCOM.Column)
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

    Private Sub PreviewDispatchInstruction(ByRef ParentForm As SAPbouiCOM.Form)

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
        pdfFilename = "Dispatch Instruction"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Dispatch Instruction.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        InsDoc = Convert.ToInt32(ParentForm.Items.Item("ed_DInsDoc").Specific.Value)
        rptDocument.SetParameterValue("@DocEntry", DocNum)
        rptDocument.SetParameterValue("@InsDocNo", InsDoc)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If
    End Sub

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
            pForm.Items.Item("ed_Trucker").Specific.Active = True
            pForm.Items.Item("ed_EUC").Enabled = pValue
            pForm.Items.Item("ed_Attent").Enabled = pValue
            pForm.Items.Item("ed_TkrTel").Enabled = pValue
            pForm.Items.Item("ed_Fax").Enabled = pValue
            pForm.Items.Item("ed_Email").Enabled = pValue
            pForm.Items.Item("ed_Trucker").Specific.Active = True

        Catch ex As Exception

        End Try
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


        If ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-LCL") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Air") Or _
             ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Land") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-LCL") Or _
            ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Air") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Land") Or _
            ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Transhipment") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Local") Then

            rptPath = Application.StartupPath.ToString & "\Trucking Instruction.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-FCL") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-FCL") Then
            rptPath = Application.StartupPath.ToString & "\Trucking Instruction FCL.rpt"
        End If

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
        If ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-LCL") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Import Sea LCL.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-FCL") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Import Sea FCL.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Air") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Import Air.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Land") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Import Land.rpt"

        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-LCL") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Export Sea LCL.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-FCL") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Export Sea FCL.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Air") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Export Air.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Land") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Export Land.rpt"

        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Transhipment") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Transhipment.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Local") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Local.rpt"

        End If




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

    'MSW To Edit New Ticket
    Public Function CheckQtyValue(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal formName As String) As Boolean
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
            ElseIf formName = "Crane" Then
                If oMatrix.Columns.Item("colCType").Cells.Item(i).Specific.Value <> "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                ElseIf oMatrix.RowCount = 1 And oMatrix.Columns.Item("colCType").Cells.Item(i).Specific.Value = "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                End If
            ElseIf formName = "Forklift" Then
                If oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value <> "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                ElseIf oMatrix.RowCount = 1 And oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value = "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                End If
            ElseIf formName = "Crate" Then
                If oMatrix.Columns.Item("colDimen").Cells.Item(i).Specific.Value <> "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                ElseIf oMatrix.RowCount = 1 And oMatrix.Columns.Item("colDimen").Cells.Item(i).Specific.Value = "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                End If
            ElseIf formName = "Bunker" Then
                If oMatrix.Columns.Item("colPermit").Cells.Item(i).Specific.Value <> "" And Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Then
                    CheckQtyValue = True
                    Exit Function
                ElseIf oMatrix.RowCount = 1 And oMatrix.Columns.Item("colPermit").Cells.Item(i).Specific.Value = "" And Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Then
                    CheckQtyValue = True
                    Exit Function
                End If
            ElseIf formName = "Toll" Then
                If oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value <> "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                ElseIf oMatrix.RowCount = 1 And oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value = "" And (Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) = 0.0 Or Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value) = 0.0) Then
                    CheckQtyValue = True
                    Exit Function
                End If
            End If
        Next
        CheckQtyValue = False
    End Function

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
    Private Function PopulateTruckingPOToEditTab(ByVal pForm As SAPbouiCOM.Form, ByVal pStrSQL As String) As Boolean
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
            pForm.DataSources.UserDataSources.Item("PODOCNO").ValueEx = pForm.Items.Item("ed_PODocNo").Specific.Value
            oRecordSet.DoQuery("select  OCRD.CardName,OCPR.Name,OCPR.Tel1,OCPR.Fax,OCPR.E_MailL,OCRD.VatIdUnCmp from OCPR LEFT OUTER JOIN OCRD ON OCPR.Name = OCRD.CntctPrsn where OCRD.CardCode = '" + sCode + "'")
            If oRecordSet.RecordCount > 0 Then
                sName = oRecordSet.Fields.Item("CardName").Value.ToString
                sAttention = oRecordSet.Fields.Item("Name").Value.ToString
                sPhone = oRecordSet.Fields.Item("Tel1").Value.ToString
                sFax = oRecordSet.Fields.Item("Fax").Value.ToString
                sMail = oRecordSet.Fields.Item("E_MailL").Value.ToString
                UEN = oRecordSet.Fields.Item("VatIdUnCmp").Value.ToString  ' MSW To Edit New Ticket
            End If
            ' pForm.Items.Item("ed_PONo").Specific.Value = DocLastKey
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
            sql = "Select U_PORMKS,U_POIRMKS,U_ColFrm,U_TkrTo,U_TkrIns from [@OBT_TB08_FFCPO] where DocEntry = " & FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value)
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                pForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = oRecordSet.Fields.Item("U_ColFrm").Value
                pForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("U_TkrTo").Value
                pForm.DataSources.UserDataSources.Item("TKRINS").ValueEx = oRecordSet.Fields.Item("U_TkrIns").Value
                pForm.DataSources.UserDataSources.Item("TKRIRMK").ValueEx = oRecordSet.Fields.Item("U_POIRMKS").Value
                pForm.DataSources.UserDataSources.Item("TKRRMK").ValueEx = oRecordSet.Fields.Item("U_PORMKS").Value
            End If
            pForm.DataSources.UserDataSources.Item("TORIGIN").ValueEx = "Y" '2-12
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
                        .SetValue("U_PO", .Offset, pMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                        .SetValue("U_InsDate", .Offset, CDate(oRecordSet.Fields.Item("U_InsDate").Value).ToString("yyyyMMdd"))

                        .SetValue("U_Mode", .Offset, oRecordSet.Fields.Item("U_Mode").Value.ToString)
                        .SetValue("U_TkrCode", .Offset, oRecordSet.Fields.Item("U_TkrCode").Value.ToString)
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
                        .SetValue("U_PODate", .Offset, pMatrix.Columns.Item("colDate").Cells.Item(currentRow).Specific.Value) 'MSW 14-09-2011 Truck PO
                        .SetValue("U_MultiJob", .Offset, pMatrix.Columns.Item("colMulti").Cells.Item(currentRow).Specific.Value) 'MSW 14-09-2011 Truck PO
                        .SetValue("U_OriginPO", .Offset, pMatrix.Columns.Item("colOrigin").Cells.Item(currentRow).Specific.Value) 'MSW 14-09-2011 Truck PO

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
    Private Function CancelTruckingPurchaseOrder(ByRef PONO As Integer) As Boolean

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
            If oPurchaseDocument.GetByKey(PONO) Then
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


    Private Sub AddUpdateOtherCharges(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer)

        ' **********************************************************************************
        '   Function    :   SetOtherChargesDataToEditTabByIndex()
        '   Purpose     :   This function will be providing to update other charges form  for
        '                   ImportSeaFCL Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,SAPbouiCOM.Matrix,
        '               :   ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer
        '   Return      :   No
        '               
        '                   
        ' **********************************************************************************

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

#Region "----------Container View List & Edit"

    Private Sub AddUpdateContainer(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)


        ' **********************************************************************************
        '   Function    :   AddUpdateContainer()
        '   Purpose     :   This function will be providing to update  container information  items  data for
        '                   importSeaFcl Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal DataSource As String, ByVal ProcressedState As Boolean
        '   Return      :   No
        '*************************************************************

        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        '  ObjDBDataSource.Offset = 0
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
                If ObjDBDataSource.GetValue("U_ConSeqNo", 0) = vbNullString Then pMatrix.Clear()
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

        End Try
    End Sub
    Private Sub SetContainerDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)


        ' **********************************************************************************
        '   Function    :   SetContainerDataToEditTabByIndex()
        '   Purpose     :   This function will be providing to set  container data by tab index  for
        '                   imporeSeaFcl Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form
        '             
        '   Return      :   No
        '*************************************************************
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

        End Try
    End Sub

    Private Sub GetContainerDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)

        ' **********************************************************************************
        '   Function    :   GetContainerDataFromMatrixByIndex()
        '   Purpose     :   This function will be providing to get Container data with matrix index  for
        '                   imporeSeaFcl Form 
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix
        '               :   ByVal Index As Integer

        '   Return      :   No
        '*************************************************************
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

        ' **********************************************************************************
        '   Function    :   DeleteContainerByIndex()
        '   Purpose     :   This function will be providing to delete Container data in
        '                   imporeSeaFcl Form 
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal DataSource As String

        '   Return      :   No
        '*************************************************************

        Try
            If pMatrix.IsRowSelected(pMatrix.GetNextSelectedRow) = True Then
                Try
                    If (pForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                        pForm.DataSources.DBDataSources.Item("@OBT_FCL19_HCONTAINE").RemoveRecord(pMatrix.GetNextSelectedRow - 1)
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

        ' **********************************************************************************
        '   Function    :   UpdateNoofContainer()
        '   Purpose     :   This function will be providing to update Container data in
        '                   imporeSeaFcl Form 
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '             
        '   Return      :   No
        '*************************************************************
        'MSW No of Container

        Dim iCon20 As Integer = 0
        Dim iCon40 As Integer = 0
        Dim iCon45 As Integer = 0
        pMatrix = pForm.Items.Item("mx_ConTab").Specific

        If pMatrix.RowCount > 0 Then
            For i As Integer = 1 To pMatrix.RowCount
                If pMatrix.Columns.Item("colCNo1").Cells.Item(i).Specific.Value() <> "" Then
                    If pMatrix.Columns.Item("colCSize1").Cells.Item(i).Specific.Value() = "20'" Then
                        iCon20 = iCon20 + 1
                    ElseIf pMatrix.Columns.Item("colCSize1").Cells.Item(i).Specific.Value() = "40'" Then
                        iCon40 = iCon40 + 1
                    ElseIf pMatrix.Columns.Item("colCSize1").Cells.Item(i).Specific.Value() = "45'" Then
                        iCon45 = iCon45 + 1
                    End If
                End If
                
            Next
            pForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Con20", 0, iCon20)
            pForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Con40", 0, iCon40)
            pForm.DataSources.DBDataSources.Item("@OBT_FCL01_EXPORT").SetValue("U_Con45", 0, iCon45)
        End If
    End Sub
#End Region

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
                            'Document.DocCurrency = "SGD"
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
                        'Document.DocCurrency = "SGD"
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

#Region "Email"
    Dim jobNo As String = ""
    Dim pdfFilename As String = ""
    Dim mainFolder As String = ""
    Dim pdffilepath As String = ""
    Dim reportuti As New clsReportUtilities
    Dim rptDocument As ReportDocument
    Dim rptPath As String = ""
    Dim originpdf As String = ""
#End Region

#Region "Shipping Inv"
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
        If AlreadyExist("EXPORTSEAFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
        ElseIf AlreadyExist("EXPORTAIRFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRFCL", 1)
        End If

        oShpForm = p_oSBOApplication.Forms.ActiveForm
        oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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

        If AddUserDataSrc(oShpForm, "ItemNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Part", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "PartDesp", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Qty", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_NUMBER) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Unit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oShpForm, "Box", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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


        oShpForm.Items.Item("ed_ExInv").Specific.DataBind.SetBound(True, "", "ExInv")
        oShpForm.Items.Item("ed_PO").Specific.DataBind.SetBound(True, "", "PONo")

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
        Dim aliasName As String = ""
        sql = "Select AliasName From OCRD where CardCOde='" & oActiveForm.Items.Item("ed_Code").Specific.Value & "'"
        oRecordSet.DoQuery(sql)
        If oRecordSet.RecordCount > 0 Then
            aliasName = oRecordSet.Fields.Item("AliasName").Value
        End If

        If AddChooseFromListByFilter(oShpForm, "cflPart", False, "PART", "U_BPCode", SAPbouiCOM.BoConditionOperation.co_EQUAL, aliasName) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oShpForm.Items.Item("ed_Part").Specific.ChooseFromListUID = "cflPart"
        oShpForm.Items.Item("ed_Part").Specific.ChooseFromListAlias = "U_PartNo"
        oShpForm.Items.Item("bt_PPView").Visible = False
        oMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
        If oShpForm.Items.Item("ed_ItemNo").Specific.Value = "" Then
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
            pForm.Items.Item("ed_DOM").Specific.Value = pMatrix.Columns.Item("colDOM").Cells.Item(Index).Specific.Value
            'pForm.Items.Item("ed_Part").Specific.Value = pMatrix.Columns.Item("colPart").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Part").Specific.Active = True

            ' If HolidaysMarkUp(pForm, pForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, pForm.Items.Item("ed_CunDay").Specific, pForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            pForm.Freeze(False)

        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

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

                    .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.GetNextSelectedRow).Specific.Value)
                    .SetValue("U_SerNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.GetNextSelectedRow).Specific.Value)
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
#End Region

    Private Sub CalculateNoOfBoxes(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix)
        If pMatrix.RowCount > 1 Then
            For i As Integer = Convert.ToInt32(pForm.Items.Item("ed_ItemNo").Specific.Value) To pMatrix.RowCount - 1
                pMatrix.Columns.Item("colBox").Cells.Item(i + 1).Specific.Value() = pMatrix.Columns.Item("colBoxLast").Cells.Item(i).Specific.Value() + 1
                pMatrix.Columns.Item("colBoxLast").Cells.Item(i + 1).Specific.Value() = (Convert.ToDouble(pMatrix.Columns.Item("colBox").Cells.Item(i + 1).Specific.Value()) + Convert.ToDouble(pMatrix.Columns.Item("colTNBox").Cells.Item(i + 1).Specific.Value())) - 1
            Next
        End If

    End Sub

    Private Function UpdateForCancelStatus(ByVal PONo As String) As Boolean
        UpdateForCancelStatus = False
        Try
            Dim Status As String = "Cancelled"
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_POStatus = " + FormatString(Status) + " WHERE U_PONo = " + FormatString(PONo))
            UpdateForCancelStatus = True
        Catch ex As Exception
            UpdateForCancelStatus = False
        End Try
    End Function

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
    End Sub

  

    Private Sub EnableExportImport(ByVal pForm As SAPbouiCOM.Form, ByVal t As Boolean, ByVal f As Boolean)
        pForm.Items.Item("ed_ETADate").Visible = f
        pForm.Items.Item("ed_ETAHr").Visible = f
        pForm.Items.Item("ed_IVsl1").Visible = f
        pForm.Items.Item("ed_IFrom1").Visible = f
        pForm.Items.Item("ed_ITo1").Visible = f
        pForm.Items.Item("ed_IDate2").Visible = f
        pForm.Items.Item("ed_IETA2").Visible = f
        pForm.Items.Item("ed_IVsl2").Visible = f
        pForm.Items.Item("ed_IFrom2").Visible = f
        pForm.Items.Item("ed_ITo2").Visible = f
        pForm.Items.Item("ed_IShpAgt").Visible = f
        pForm.Items.Item("ed_IOBL").Visible = f
        pForm.Items.Item("cb_ISur").Visible = f
        pForm.Items.Item("ed_IHBL").Visible = f
        pForm.Items.Item("ed_ICutOff").Visible = f
        pForm.Items.Item("ed_ICtOffT").Visible = f
        pForm.Items.Item("ed_IBokRef").Visible = f
        pForm.Items.Item("381").Visible = f
        pForm.Items.Item("382").Visible = f
        pForm.Items.Item("406").Visible = f
        pForm.Items.Item("407").Visible = f
        pForm.Items.Item("408").Visible = f
        pForm.Items.Item("409").Visible = f
        pForm.Items.Item("377").Visible = f
        pForm.Items.Item("383").Visible = f
        pForm.Items.Item("385").Visible = f
        pForm.Items.Item("384").Visible = f
        pForm.Items.Item("stBokRef").Visible = f
        pForm.Items.Item("lk_IShp").Visible = f

        pForm.Items.Item("ed_ETDDate").Visible = t
        pForm.Items.Item("ed_ETDHr").Visible = t
        pForm.Items.Item("ed_Vessel").Visible = t
        pForm.Items.Item("ed_EFrom1").Visible = t
        pForm.Items.Item("ed_ETo1").Visible = t
        pForm.Items.Item("ed_EDate2").Visible = t
        pForm.Items.Item("ed_EETA2").Visible = t
        pForm.Items.Item("ed_EVsl2").Visible = t
        pForm.Items.Item("ed_EFrom2").Visible = t
        pForm.Items.Item("ed_ETo2").Visible = t
        pForm.Items.Item("ed_ShpAgt").Visible = t
        pForm.Items.Item("ed_OBL").Visible = t
        pForm.Items.Item("cb_ESur").Visible = t
        pForm.Items.Item("ed_HBL").Visible = t
        pForm.Items.Item("ed_ECutOff").Visible = t
        pForm.Items.Item("ed_ECtOffT").Visible = t
        pForm.Items.Item("ed_BokRef").Visible = t
        pForm.Items.Item("397").Visible = t
        pForm.Items.Item("422").Visible = t
        pForm.Items.Item("431").Visible = t
        pForm.Items.Item("432").Visible = t
        pForm.Items.Item("433").Visible = t
        pForm.Items.Item("434").Visible = t
        pForm.Items.Item("421").Visible = t
        pForm.Items.Item("423").Visible = t
        pForm.Items.Item("425").Visible = t
        pForm.Items.Item("424").Visible = t
        pForm.Items.Item("488").Visible = t
        pForm.Items.Item("lk_EShp").Visible = t

    End Sub
    Private Sub AddTabMatrixRow(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal firstCol As String, ByVal secCol As String, ByVal dbSource As String)

        ' **********************************************************************************
        '   Function    :   AddTabMatrixRow()
        '   Purpose     :   This function will be providing to update Container data in
        '                   imporeSeaFcl Form 
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '             
        '   Return      :   No
        '*************************************************************
        'MSW No of Container
        Try
            If pMatrix.RowCount = 0 Then
                pForm.DataSources.DBDataSources.Item(dbSource).Clear()
                pMatrix.AddRow(1)
                pMatrix.FlushToDataSource()
                pMatrix.Columns.Item(firstCol).Cells.Item(pMatrix.RowCount).Specific.Value = pMatrix.RowCount.ToString
            ElseIf pMatrix.RowCount = 1 And pMatrix.Columns.Item(firstCol).Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                pMatrix.Columns.Item(firstCol).Cells.Item(pMatrix.RowCount).Specific.Value = pMatrix.RowCount.ToString
            ElseIf pMatrix.Columns.Item(secCol).Cells.Item(pMatrix.RowCount).Specific.Value <> "" Then
                pForm.DataSources.DBDataSources.Item(dbSource).Clear()
                pMatrix.AddRow(1)
                pMatrix.FlushToDataSource()
                pMatrix.Columns.Item(firstCol).Cells.Item(pMatrix.RowCount).Specific.Value = pMatrix.RowCount.ToString
            End If
        Catch ex As Exception

        End Try

    End Sub


#Region "Create PO"
    Public Function CreateGenPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal cardcode As String, ByVal iCode As String) As Boolean
        Dim sErrDesc As String = ""
        Dim i As Integer = 0
        Dim dblPrice As Double = 0.0
        Dim itemCode As String = String.Empty
        Dim itemDesc As String = String.Empty
        Dim originpdf As String = String.Empty
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection

        CreateGenPO = False
        Try

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CreatePO(oActiveForm, cardcode, "ed_POICode", "ed_POIDesc", "ed_POQty", "ed_POPrice", "ed_PONo", "ed_PO", "ed_PStus") Then Throw New ArgumentException(sErrDesc)
            oGeneralService = p_oDICompany.GetCompanyService.GetGeneralService("FCPO")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            dblPrice = Convert.ToDouble(oActiveForm.Items.Item("ed_POQty").Specific.Value) * Convert.ToDouble(oActiveForm.Items.Item("ed_POPrice").Specific.Value)
            oGeneralData.SetProperty("U_VCode", oActiveForm.Items.Item("ed_TkrCode").Specific.Value)
            oGeneralData.SetProperty("U_VName", oActiveForm.Items.Item("ed_Trucker").Specific.Value)
            oGeneralData.SetProperty("U_PONo", oActiveForm.Items.Item("ed_PONo").Specific.Value)
            oGeneralData.SetProperty("U_PO", oActiveForm.Items.Item("ed_PO").Specific.Value)
            oGeneralData.SetProperty("U_PODate", Today.Date)
            oGeneralData.SetProperty("U_PODay", Today.DayOfWeek.ToString.Substring(0, 3))
            oGeneralData.SetProperty("U_POTime", Convert.ToDateTime(Now.ToString("HH:mm")))
            oGeneralData.SetProperty("U_POITPD", dblPrice)
            oGeneralData.SetProperty("U_POStatus", "Open")
            oGeneralData.SetProperty("U_PORMKS", oActiveForm.Items.Item("ee_Rmsk").Specific.Value)
            oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ee_InsRmsk").Specific.Value)
            oGeneralData.SetProperty("U_ColFrm", oActiveForm.Items.Item("ee_ColFrm").Specific.Value)
            oGeneralData.SetProperty("U_TkrTo", oActiveForm.Items.Item("ee_TkrTo").Specific.Value)
            oGeneralData.SetProperty("U_TkrIns", oActiveForm.Items.Item("ee_TkrIns").Specific.Value)

            oChildren = oGeneralData.Child("OBT_TB09_FFCPOITEM")
            oChild = oChildren.Add
            oChild.SetProperty("U_POINO", oActiveForm.Items.Item("ed_POICode").Specific.Value)
            oChild.SetProperty("U_POIDesc", oActiveForm.Items.Item("ed_POIDesc").Specific.Value)
            oChild.SetProperty("U_POIPrice", dblPrice)
            oChild.SetProperty("U_POIQty", oActiveForm.Items.Item("ed_POQty").Specific.Value)
            oChild.SetProperty("U_POIAmt", dblPrice)
            oChild.SetProperty("U_POIGST", "XI")
            oChild.SetProperty("U_POITot", dblPrice)

           
            oGeneralService.Add(oGeneralData)

            oRecordSet.DoQuery("SELECT DocEntry,DocNum FROM [@OBT_TB08_FFCPO] Order By DocEntry")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveLast()
                oActiveForm.Items.Item("ed_PODocNo").Specific.Value = oRecordSet.Fields.Item("DocEntry").Value
            End If
            CreateGenPO = True
        Catch ex As Exception
            CreateGenPO = False
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Public Function CreatePO(ByRef oActiveForm As SAPbouiCOM.Form, ByVal CardCode As String, ByVal iCode As String, ByVal iDesc As String, ByVal iQty As String, ByVal iPrice As String, ByVal txtDocEntry As String, ByVal txtDocNum As String, ByVal txtStus As String) As Boolean
        '2S001 Card code
        '4601 Icode

        Dim oPurchaseDocument As SAPbobsCOM.Documents
        Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim dblPrice As Double = 0.0
        Dim sErrDesc As String = vbNullString
        CreatePO = False
        Try
            oPurchaseDocument = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            oBusinessPartner = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            p_oDICompany.GetNewObjectCode("22")

            If oBusinessPartner.GetByKey(CardCode) Then
                oRecordset.DoQuery("Select CardName From OCRD Where CardCode= '" & CardCode & "'")
                oPurchaseDocument.CardCode = CardCode
                oPurchaseDocument.CardName = oRecordset.Fields.Item("CardName").Value
                If oBusinessPartner.Currency = "##" Then
                    oPurchaseDocument.DocCurrency = "SGD"
                Else
                    oPurchaseDocument.DocCurrency = oBusinessPartner.Currency
                End If

                oPurchaseDocument.DocDate = Now
                oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                oPurchaseDocument.TaxDate = Now

                If oRecordset.RecordCount > 0 Then
                    oPurchaseDocument.Lines.ItemCode = oActiveForm.Items.Item(iCode).Specific.Value
                    oPurchaseDocument.Lines.ItemDescription = oActiveForm.Items.Item(iDesc).Specific.Value
                    oPurchaseDocument.Lines.Quantity = oActiveForm.Items.Item(iQty).Specific.Value
                    oPurchaseDocument.Lines.UnitPrice = oActiveForm.Items.Item(iPrice).Specific.Value
                    oPurchaseDocument.Lines.VatGroup = "XI"
                    oPurchaseDocument.Lines.Add()
                End If

            End If
            Dim ret As Long = oPurchaseDocument.Add
            If ret <> 0 Then
                p_oDICompany.GetLastError(ret, sErrDesc)
                MessageBox.Show("Error Code" + ret.ToString + " / " + sErrDesc.ToString)
            Else
                sql = ""
                oRecordset.DoQuery("SELECT DocEntry,DocNum FROM OPOR Order By DocEntry")
                If oRecordset.RecordCount > 0 Then
                    oRecordset.MoveLast()
                    oActiveForm.Items.Item(txtDocEntry).Specific.Value = oRecordset.Fields.Item("DocEntry").Value
                    oActiveForm.Items.Item(txtDocNum).Specific.Value = oRecordset.Fields.Item("DocNum").Value
                    oActiveForm.Items.Item(txtStus).Specific.Value = "Open"
                    sql = "Update OPOR Set U_JobNo='" & oActiveForm.Items.Item("ed_JobNo").Specific.Value & "' where DocEntry='" & oRecordset.Fields.Item("DocEntry").Value & "'"
                End If
                If sql <> "" Then
                    oRecordset.DoQuery(sql)
                End If
            End If

            CreatePO = True
        Catch ex As Exception
            CreatePO = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Sub NewPOEditTab(ByRef oActiveForm As SAPbouiCOM.Form, ByVal matrixName As String)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim sErrDesc As String = ""
        oActiveForm.Freeze(True)

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If matrixName = "mx_TkrList" Then
            oActiveForm.Items.Item("bt_PO").Enabled = False
            If oActiveForm.Items.Item("bt_TkrAdd").Specific.Caption = "Add" Then
                oActiveForm.Items.Item("ed_InsDoc").Specific.Value = ""
                oActiveForm.Items.Item("ed_PONo").Specific.Value = ""
                oActiveForm.Items.Item("ed_Trucker").Specific.Value = ""
                oActiveForm.Items.Item("ed_VehicNo").Specific.Value = ""
                oActiveForm.Items.Item("ed_TkrTime").Specific.Value = ""
                oActiveForm.DataSources.UserDataSources.Item("TMULTI").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TORIGIN").ValueEx = ""

                oActiveForm.DataSources.UserDataSources.Item("PODOCNO").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("PO").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("INSDATE").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("EUC").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = ""

                oActiveForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRDATE").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRINS").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRIRMK").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRRMK").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("TKRCOL").ValueEx = ""
              

                'ClearText(oActiveForm, "ed_InsDoc", "ed_TMulti", "ed_TOrigin", "ed_PODocNo", "ed_PONo", "ed_PO", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk", "ee_ColFrm") 'MSW New Ticket 07-09-2011
                oRecordSet.DoQuery("SELECT Address FROM OCRD WHERE CardCode = '" & oActiveForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                If oRecordSet.RecordCount > 0 Then
                    oActiveForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("Address").Value.ToString
                End If
                If oActiveForm.Items.Item("chk_Cust").Specific.Checked = True Then
                    oActiveForm.DataSources.UserDataSources.Item("TKRINS").ValueEx = """This permit require custom endorsement"""
                End If
                oActiveForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                oActiveForm.DataSources.UserDataSources.Item("TKRDATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                oActiveForm.DataSources.UserDataSources.Item("DATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                oActiveForm.DataSources.UserDataSources.Item("CREATEDBY").ValueEx = p_oDICompany.UserName.ToString
                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                '2-12
                oActiveForm.Items.Item("op_Inter").Specific.Selected = True
                oActiveForm.Items.Item("op_Exter").Enabled = True
                oActiveForm.Items.Item("op_Inter").Enabled = True
                oActiveForm.Items.Item("ed_Trucker").Enabled = True
                'oActiveForm.Items.Item("bt_PO").Enabled = True
                ' EnabledTruckerForExternal(oActiveForm, True)
            End If
            oMatrix = oActiveForm.Items.Item(matrixName).Specific
            If oActiveForm.Items.Item("ed_InsDoc").Specific.Value = "" Then
                oActiveForm.Items.Item("ed_TkrTime").Specific.Value = Now.ToString("HH:mm")
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
        ElseIf matrixName = "mx_DspList" Then
            oActiveForm.Items.Item("bt_DPO").Enabled = False
            If oActiveForm.Items.Item("bt_DspAdd").Specific.Caption = "Add" Then
                oActiveForm.Items.Item("ed_DInsDoc").Specific.Value = ""
                oActiveForm.Items.Item("ed_DPONo").Specific.Value = ""
                oActiveForm.Items.Item("ed_Dspatch").Specific.Value = ""
                oActiveForm.Items.Item("ed_DspTime").Specific.Value = ""
                oActiveForm.DataSources.UserDataSources.Item("DMULTI").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DORIGIN").ValueEx = ""

                oActiveForm.DataSources.UserDataSources.Item("DSPPODOCNO").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DPO").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DINSDATE").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DEUC").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPATTE").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPTEL").ValueEx = ""

                oActiveForm.DataSources.UserDataSources.Item("DSPFAX").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPMAIL").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPDATE").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPINS").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPIRMK").ValueEx = ""
                oActiveForm.DataSources.UserDataSources.Item("DSPRMK").ValueEx = ""

                'ClearText(oActiveForm, "ed_DInsDoc", "ed_DMulti", "ed_DOrigin", "ed_DPDocNo", "ed_DPONo", "ed_DPO", "ed_DPStus", "ed_DInDate", "ed_Dspatch", "ed_DEUC", "ed_DAttent", "ed_DspTel", "ed_DFax", "ed_DEmail", "ed_DspDate", "ed_DspTime", "ee_DspIns", "ee_DIRmsk", "ee_DRmsk") 'MSW New Ticket 07-09-2011
                oActiveForm.Items.Item("ed_DInDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                oActiveForm.DataSources.UserDataSources.Item("DSPDATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                oActiveForm.DataSources.UserDataSources.Item("DDATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                oActiveForm.DataSources.UserDataSources.Item("DCREATEDBY").ValueEx = p_oDICompany.UserName.ToString
                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DInDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                '2-12
                oActiveForm.Items.Item("op_DInter").Specific.Selected = True
                oActiveForm.Items.Item("op_DExter").Enabled = True
                oActiveForm.Items.Item("op_DInter").Enabled = True
                oActiveForm.Items.Item("ed_Dspatch").Enabled = True
                'oActiveForm.Items.Item("bt_DPO").Enabled = True
                'EnabledDispatchForExternal(oActiveForm, True)
            End If
            oMatrix = oActiveForm.Items.Item(matrixName).Specific
            If oActiveForm.Items.Item("ed_DInsDoc").Specific.Value = "" Then
                oActiveForm.Items.Item("ed_DspTime").Specific.Value = Now.ToString("HH:mm")
                If (oMatrix.RowCount > 0) Then
                    If (oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value = Nothing) Then
                        oActiveForm.Items.Item("ed_DInsDoc").Specific.Value = 1
                    Else
                        oActiveForm.Items.Item("ed_DInsDoc").Specific.Value = oMatrix.Columns.Item("colInsDoc").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                    End If
                Else
                    oActiveForm.Items.Item("ed_DInsDoc").Specific.Value = 1
                End If
                oActiveForm.Items.Item("ed_Dspatch").Specific.Active = True
            End If
        End If

        oActiveForm.Freeze(False)
    End Sub
#End Region

#Region "Save to PO Tab"
    Public Sub SaveToPOTab(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal ProcressedState As Boolean, ByVal PONo As String, ByVal PODocNo As String, _
                           ByVal Vendor As String, ByVal PODate As String, ByVal Desc As String, ByVal Status As String)

        ObjDBDataSource = pForm.DataSources.DBDataSources.Item("@OBT_TB01_POLIST")
        If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        Try
            If ProcressedState = True Then
                If ObjDBDataSource.GetValue("U_PODocNo", 0) = vbNullString Then pMatrix.Clear()
                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)

                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)

                    End If
                    .SetValue("U_PODocNo", .Offset, PODocNo) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PONo", .Offset, PONo)
                    .SetValue("U_VName", .Offset, Vendor)
                    .SetValue("U_PODate", .Offset, PODate)
                    .SetValue("U_Desc", .Offset, Desc)
                    .SetValue("U_POStatus", .Offset, Status)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource

                    '.SetValue("LineId", .Offset, pForm.Items.Item("ed_InsDoc").Specific.Value)
                    '.SetValue("U_InsDocNo", .Offset, pForm.Items.Item("ed_InsDoc").Specific.Value)
                    '.SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_PODocNo").Specific.Value) 'MSW 14-09-2011 Truck PO
                    '.SetValue("U_PONo", .Offset, pForm.Items.Item("ed_PONo").Specific.Value)
                    '.SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value) 'New UI PO Serial No
                    '.SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_InsDate").Specific.Value)
                    '.SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_Inter").Specific.Selected = True, "Internal", "External").ToString)
                    '.SetValue("U_Trucker", .Offset, pForm.Items.Item("ed_Trucker").Specific.Value)
                    '.SetValue("U_VehNo", .Offset, pForm.Items.Item("ed_VehicNo").Specific.Value)
                    '.SetValue("U_EUC", .Offset, pForm.Items.Item("ed_EUC").Specific.Value)
                    '.SetValue("U_Attent", .Offset, pForm.Items.Item("ed_Attent").Specific.Value)
                    '.SetValue("U_Tel", .Offset, pForm.Items.Item("ed_TkrTel").Specific.Value)
                    '.SetValue("U_Fax", .Offset, pForm.Items.Item("ed_Fax").Specific.Value)
                    '.SetValue("U_Email", .Offset, pForm.Items.Item("ed_Email").Specific.Value)
                    '.SetValue("U_TkrDate", .Offset, pForm.Items.Item("ed_TkrDate").Specific.Value)
                    ''.SetValue("U_TkrTime", .Offset, TruckingTime.ToString())
                    '.SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value)
                    ''If Not pForm.Items.Item("ed_TkrTime").Specific.Value.ToString = "" Then
                    ''    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(2, 2).ToString())
                    ''End If
                    ''MSW 14-09-2011 Truck PO
                    'If pForm.Items.Item("op_Exter").Specific.Selected = True Then
                    '    .SetValue("U_Status", .Offset, pForm.Items.Item("ed_PStus").Specific.Value)
                    'Else
                    '    .SetValue("U_Status", .Offset, "")
                    'End If
                    ''MSW 14-09-2011 Truck PO
                    '.SetValue("U_ColFrm", .Offset, pForm.Items.Item("ee_ColFrm").Specific.Value)
                    '.SetValue("U_TkrTo", .Offset, pForm.Items.Item("ee_TkrTo").Specific.Value)
                    '.SetValue("U_TkrIns", .Offset, pForm.Items.Item("ee_TkrIns").Specific.Value)
                    '.SetValue("U_InsRemsk", .Offset, pForm.Items.Item("ee_InsRmsk").Specific.Value)
                    '.SetValue("U_Remarks", .Offset, pForm.Items.Item("ee_Rmsk").Specific.Value) 'MSW 14-09-2011 Truck PO
                    '.SetValue("U_PrepBy", .Offset, pForm.Items.Item("ed_Created").Specific.Value)
                    '.SetValue("U_PODate", .Offset, pForm.Items.Item("ed_Date").Specific.Value)
                    'pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "Dispatch Tab"
    Private Function PopulateDispatchPOToEditTab(ByVal pForm As SAPbouiCOM.Form, ByVal pStrSQL As String) As Boolean
        PopulateDispatchPOToEditTab = False

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
            pForm.DataSources.UserDataSources.Item("DSPPODOCNO").ValueEx = pForm.Items.Item("ed_DPDocNo").Specific.Value
            oRecordSet.DoQuery("select  OCRD.CardName,OCPR.Name,OCPR.Tel1,OCPR.Fax,OCPR.E_MailL,OCRD.VatIdUnCmp from OCPR LEFT OUTER JOIN OCRD ON OCPR.Name = OCRD.CntctPrsn where OCRD.CardCode = '" + sCode + "'")
            If oRecordSet.RecordCount > 0 Then
                sName = oRecordSet.Fields.Item("CardName").Value.ToString
                sAttention = oRecordSet.Fields.Item("Name").Value.ToString
                sPhone = oRecordSet.Fields.Item("Tel1").Value.ToString
                sFax = oRecordSet.Fields.Item("Fax").Value.ToString
                sMail = oRecordSet.Fields.Item("E_MailL").Value.ToString
                UEN = oRecordSet.Fields.Item("VatIdUnCmp").Value.ToString  ' MSW To Edit New Ticket
            End If
            ' pForm.Items.Item("ed_PONo").Specific.Value = DocLastKey
            oEditText = pForm.Items.Item("ed_Dspatch").Specific
            oEditText.DataBind.SetBound(True, "", "DSPEXTR")
            oEditText.ChooseFromListUID = "CFLDSPV"
            oEditText.ChooseFromListAlias = "CardName"
            pForm.DataSources.UserDataSources.Item("DSPEXTR").ValueEx = sName
            pForm.DataSources.UserDataSources.Item("DSPATTE").ValueEx = sAttention
            pForm.DataSources.UserDataSources.Item("DSPTEL").ValueEx = sPhone
            pForm.DataSources.UserDataSources.Item("DSPFAX").ValueEx = sFax
            pForm.DataSources.UserDataSources.Item("DSPMAIL").ValueEx = sMail
            pForm.DataSources.UserDataSources.Item("DEUC").ValueEx = UEN
            sql = "Select U_PORMKS,U_POIRMKS,U_TkrIns from [@OBT_TB08_FFCPO] where DocEntry = " & FormatString(pForm.Items.Item("ed_DPDocNo").Specific.Value)
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
      
                pForm.DataSources.UserDataSources.Item("DSPINS").ValueEx = oRecordSet.Fields.Item("U_TkrIns").Value
                pForm.DataSources.UserDataSources.Item("DSPIRMK").ValueEx = oRecordSet.Fields.Item("U_POIRMKS").Value
                pForm.DataSources.UserDataSources.Item("DSPRMK").ValueEx = oRecordSet.Fields.Item("U_PORMKS").Value
            End If
            pForm.DataSources.UserDataSources.Item("DORIGIN").ValueEx = "Y"
            PopulateDispatchPOToEditTab = True
        Catch ex As Exception
            PopulateDispatchPOToEditTab = False
            MessageBox.Show(ex.Message)
        End Try

    End Function
    Private Function PopulateDispatchPurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String) As Boolean

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
        PopulateDispatchPurchaseHeader = False
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
                        .SetValue("U_PO", .Offset, pMatrix.Columns.Item("colPO").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                        .SetValue("U_InsDate", .Offset, CDate(oRecordSet.Fields.Item("U_InsDate").Value).ToString("yyyyMMdd"))

                        .SetValue("U_Mode", .Offset, oRecordSet.Fields.Item("U_Mode").Value.ToString)
                        .SetValue("U_DspCode", .Offset, oRecordSet.Fields.Item("U_DspCode").Value.ToString)
                        .SetValue("U_Dispatch", .Offset, oRecordSet.Fields.Item("U_Dispatch").Value.ToString)
                        .SetValue("U_EUC", .Offset, oRecordSet.Fields.Item("U_EUC").Value.ToString)
                        .SetValue("U_Attent", .Offset, oRecordSet.Fields.Item("U_Attent").Value.ToString)
                        .SetValue("U_Tel", .Offset, oRecordSet.Fields.Item("U_Tel").Value.ToString)
                        .SetValue("U_Fax", .Offset, oRecordSet.Fields.Item("U_Fax").Value.ToString)
                        .SetValue("U_Email", .Offset, oRecordSet.Fields.Item("U_Email").Value.ToString)
                        .SetValue("U_DspDate", .Offset, CDate(oRecordSet.Fields.Item("U_DspDate").Value).ToString("yyyyMMdd"))
                        .SetValue("U_DspTime", .Offset, oRecordSet.Fields.Item("U_DspTime").Value)
                        .SetValue("U_DspIns", .Offset, oRecordSet.Fields.Item("U_TkrIns").Value.ToString)
                        .SetValue("U_InsRemsk", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                        .SetValue("U_Remarks", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString) 'MSW 14-09-2011 Truck PO
                        .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString) 'MSW 14-09-2011 Truck PO
                        .SetValue("U_PrepBy", .Offset, p_oDICompany.UserName.ToString)
                        .SetValue("U_PODate", .Offset, pMatrix.Columns.Item("colDate").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_MultiJob", .Offset, pMatrix.Columns.Item("colMulti").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_OriginPO", .Offset, pMatrix.Columns.Item("colOrigin").Cells.Item(currentRow).Specific.Value)
                    End With
                    pMatrix.SetLineData(currentRow)

                    oRecordSet.MoveNext()
                Loop
            End If
            PopulateDispatchPurchaseHeader = True
        Catch ex As Exception
            PopulateDispatchPurchaseHeader = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function CreateDspGenPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal cardcode As String, ByVal iCode As String) As Boolean
        Dim sErrDesc As String = ""
        Dim i As Integer = 0
        Dim dblPrice As Double = 0.0
        Dim itemCode As String = String.Empty
        Dim itemDesc As String = String.Empty
        Dim originpdf As String = String.Empty
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection

        CreateDspGenPO = False
        Try

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Not CreatePO(oActiveForm, cardcode, "ed_DICode", "ed_DIDesc", "ed_DQty", "ed_DPrice", "ed_DPONo", "ed_DPO", "ed_DPStus") Then Throw New ArgumentException(sErrDesc)
            oGeneralService = p_oDICompany.GetCompanyService.GetGeneralService("FCPO")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            dblPrice = Convert.ToDouble(oActiveForm.Items.Item("ed_DQty").Specific.Value) * Convert.ToDouble(oActiveForm.Items.Item("ed_DPrice").Specific.Value)
            oGeneralData.SetProperty("U_VCode", oActiveForm.Items.Item("ed_DspCode").Specific.Value)
            oGeneralData.SetProperty("U_VName", oActiveForm.Items.Item("ed_Dspatch").Specific.Value)
            oGeneralData.SetProperty("U_PONo", oActiveForm.Items.Item("ed_DPONo").Specific.Value)
            oGeneralData.SetProperty("U_PO", oActiveForm.Items.Item("ed_DPO").Specific.Value)
            oGeneralData.SetProperty("U_PODate", Today.Date)
            oGeneralData.SetProperty("U_PODay", Today.DayOfWeek.ToString.Substring(0, 3))
            oGeneralData.SetProperty("U_POTime", Convert.ToDateTime(Now.ToString("HH:mm")))
            oGeneralData.SetProperty("U_POITPD", dblPrice)
            oGeneralData.SetProperty("U_POStatus", "Open")
            oGeneralData.SetProperty("U_PORMKS", oActiveForm.Items.Item("ee_DRmsk").Specific.Value)
            oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ee_DIRmsk").Specific.Value)
            oGeneralData.SetProperty("U_TkrIns", oActiveForm.Items.Item("ee_DspIns").Specific.Value)

            oChildren = oGeneralData.Child("OBT_TB09_FFCPOITEM")
            oChild = oChildren.Add
            oChild.SetProperty("U_POINO", oActiveForm.Items.Item("ed_DICode").Specific.Value)
            oChild.SetProperty("U_POIDesc", oActiveForm.Items.Item("ed_DIDesc").Specific.Value)
            oChild.SetProperty("U_POIPrice", dblPrice)
            oChild.SetProperty("U_POIQty", oActiveForm.Items.Item("ed_DQty").Specific.Value)
            oChild.SetProperty("U_POIAmt", dblPrice)
            oChild.SetProperty("U_POIGST", "XI")
            oChild.SetProperty("U_POITot", dblPrice)


            oGeneralService.Add(oGeneralData)

            oRecordSet.DoQuery("SELECT DocEntry,DocNum FROM [@OBT_TB08_FFCPO] Order By DocEntry")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveLast()
                oActiveForm.Items.Item("ed_DPDocNo").Specific.Value = oRecordSet.Fields.Item("DocEntry").Value
            End If
            CreateDspGenPO = True
        Catch ex As Exception
            CreateDspGenPO = False
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Public Sub AddUpdateInstructions(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        rowIndex = pMatrix.GetNextSelectedRow
        'If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
        '    rowIndex = 1
        'End If
        If rowIndex = -1 And pForm.Items.Item("ed_DInsDoc").Specific.Value <> "" Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT LineId FROM [" & DataSource & "] Where LineId=" & pForm.Items.Item("ed_DInsDoc").Specific.Value)
            If oRecordSet.RecordCount > 0 Then
                rowIndex = oRecordSet.Fields.Item("LineId").Value
            End If
        End If


        Try
            If ProcressedState = True Then
                If ObjDBDataSource.GetValue("U_InsDocNo", 0) = vbNullString Then pMatrix.Clear()
                With ObjDBDataSource
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                        .SetValue("U_InsDocNo", .Offset, 1)
                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                        .SetValue("U_InsDocNo", .Offset, pMatrix.Columns.Item("V_1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                    End If
                    .SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_DPDocNo").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PONo", .Offset, pForm.Items.Item("ed_DPONo").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_DPO").Specific.Value) 'New UI PO Serial No
                    .SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_DInDate").Specific.Value)
                    .SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_DInter").Specific.Selected = True, "Internal", "External").ToString)
                    .SetValue("U_DspCode", .Offset, pForm.Items.Item("ed_DspCode").Specific.Value)
                    .SetValue("U_Dispatch", .Offset, pForm.Items.Item("ed_Dspatch").Specific.Value)

                    .SetValue("U_EUC", .Offset, pForm.Items.Item("ed_DEUC").Specific.Value)
                    .SetValue("U_Attent", .Offset, pForm.Items.Item("ed_DAttent").Specific.Value)
                    .SetValue("U_Tel", .Offset, pForm.Items.Item("ed_DspTel").Specific.Value)
                    .SetValue("U_Fax", .Offset, pForm.Items.Item("ed_DFax").Specific.Value)
                    .SetValue("U_Email", .Offset, pForm.Items.Item("ed_DEmail").Specific.Value)
                    .SetValue("U_DspDate", .Offset, pForm.Items.Item("ed_DspDate").Specific.Value)
                    'If Not pForm.Items.Item("ed_TkrTime").Specific.Value.ToString = "" Then
                    '    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(2, 2).ToString())
                    'End If
                    '.SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(3, 2).ToString())
                    .SetValue("U_DspTime", .Offset, pForm.Items.Item("ed_DspTime").Specific.Value)
                    'MSW 14-09-2011 Truck PO
                    If pForm.Items.Item("op_DExter").Specific.Selected = True Then
                        .SetValue("U_Status", .Offset, "Open")
                    Else
                        .SetValue("U_Status", .Offset, "")
                    End If

                    .SetValue("U_DspIns", .Offset, pForm.Items.Item("ee_DspIns").Specific.Value)
                    .SetValue("U_InsRemsk", .Offset, pForm.Items.Item("ee_DIRmsk").Specific.Value)
                    .SetValue("U_Remarks", .Offset, pForm.Items.Item("ee_DRmsk").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PrepBy", .Offset, p_oDICompany.UserName.ToString)
                    .SetValue("U_PODate", .Offset, pForm.Items.Item("ed_DDate").Specific.Value)
                    .SetValue("U_MultiJob", .Offset, pForm.Items.Item("ed_DMulti").Specific.Value)
                    .SetValue("U_OriginPO", .Offset, pForm.Items.Item("ed_DOrigin").Specific.Value)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource

                    .SetValue("LineId", .Offset, pForm.Items.Item("ed_DInsDoc").Specific.Value)
                    .SetValue("U_InsDocNo", .Offset, pForm.Items.Item("ed_DInsDoc").Specific.Value)
                    .SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_DPDocNo").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PONo", .Offset, pForm.Items.Item("ed_DPONo").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_DPO").Specific.Value) 'New UI PO Serial No
                    .SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_DInDate").Specific.Value)
                    .SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_DInter").Specific.Selected = True, "Internal", "External").ToString)
                    .SetValue("U_DspCode", .Offset, pForm.Items.Item("ed_DspCode").Specific.Value)
                    .SetValue("U_Dispatch", .Offset, pForm.Items.Item("ed_Dspatch").Specific.Value)
                    .SetValue("U_EUC", .Offset, pForm.Items.Item("ed_DEUC").Specific.Value)
                    .SetValue("U_Attent", .Offset, pForm.Items.Item("ed_DAttent").Specific.Value)
                    .SetValue("U_Tel", .Offset, pForm.Items.Item("ed_DspTel").Specific.Value)
                    .SetValue("U_Fax", .Offset, pForm.Items.Item("ed_DFax").Specific.Value)
                    .SetValue("U_Email", .Offset, pForm.Items.Item("ed_DEmail").Specific.Value)
                    .SetValue("U_DspDate", .Offset, pForm.Items.Item("ed_DspDate").Specific.Value)
                    '.SetValue("U_TkrTime", .Offset, TruckingTime.ToString())
                    .SetValue("U_DspTime", .Offset, pForm.Items.Item("ed_DspTime").Specific.Value)
                    'If Not pForm.Items.Item("ed_TkrTime").Specific.Value.ToString = "" Then
                    '    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(2, 2).ToString())
                    'End If
                    'MSW 14-09-2011 Truck PO
                    If pForm.Items.Item("op_DExter").Specific.Selected = True Then
                        .SetValue("U_Status", .Offset, pForm.Items.Item("ed_DPStus").Specific.Value)
                    Else
                        .SetValue("U_Status", .Offset, "")
                    End If
                    'MSW 14-09-2011 Truck PO
                    
                    .SetValue("U_DspIns", .Offset, pForm.Items.Item("ee_DspIns").Specific.Value)
                    .SetValue("U_InsRemsk", .Offset, pForm.Items.Item("ee_DIRmsk").Specific.Value)
                    .SetValue("U_Remarks", .Offset, pForm.Items.Item("ee_DRmsk").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PrepBy", .Offset, pForm.Items.Item("ed_DCreate").Specific.Value)
                    .SetValue("U_PODate", .Offset, pForm.Items.Item("ed_DDate").Specific.Value)
                    .SetValue("U_MultiJob", .Offset, pForm.Items.Item("ed_DMulti").Specific.Value)
                    .SetValue("U_OriginPO", .Offset, pForm.Items.Item("ed_DOrigin").Specific.Value)
                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub GetDispatchDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Try
            DocumentNo = pMatrix.Columns.Item("colInsDoc").Cells.Item(Index).Specific.Value
            PurchaseOrderNo = pMatrix.Columns.Item("colPONo").Cells.Item(Index).Specific.Value
            PurchaseDocNo = pMatrix.Columns.Item("colDocNo").Cells.Item(Index).Specific.Value
            POSerialNo = pMatrix.Columns.Item("colPO").Cells.Item(Index).Specific.Value  'New UI
            Mode = pMatrix.Columns.Item("colMode").Cells.Item(Index).Specific.Value
            InstructionDate = pMatrix.Columns.Item("colInsDate").Cells.Item(Index).Specific.Value
            DspCode = pMatrix.Columns.Item("colDspCode").Cells.Item(Index).Specific.Value
            Dispatch = pMatrix.Columns.Item("colDsp").Cells.Item(Index).Specific.Value
            EUC = pMatrix.Columns.Item("colEUC").Cells.Item(Index).Specific.Value
            Attention = pMatrix.Columns.Item("colAttent").Cells.Item(Index).Specific.Value
            Telephone = pMatrix.Columns.Item("colTel").Cells.Item(Index).Specific.Value
            Fax = pMatrix.Columns.Item("colFax").Cells.Item(Index).Specific.Value
            Email = pMatrix.Columns.Item("colEmail").Cells.Item(Index).Specific.Value
            DspDate = pMatrix.Columns.Item("colDspDate").Cells.Item(Index).Specific.Value
            DspTime = pMatrix.Columns.Item("colDspTime").Cells.Item(Index).Specific.Value
            DspIns = pMatrix.Columns.Item("colDspIns").Cells.Item(Index).Specific.Value
            DspIRemarks = pMatrix.Columns.Item("colRemarks").Cells.Item(Index).Specific.Value
            Remarks = pMatrix.Columns.Item("colRmks").Cells.Item(Index).Specific.Value 'MSW 08-09-2011
            DspTime = DspTime.Substring(0, 2).ToString() & ":" & DspTime.Substring(2, 2).ToString()
            POStus = pMatrix.Columns.Item("colPStatus").Cells.Item(Index).Specific.Value 'New UI
            PODate = pMatrix.Columns.Item("colDate").Cells.Item(Index).Specific.Value
            PrepBy = pMatrix.Columns.Item("colPrepBy").Cells.Item(Index).Specific.Value
            DMulti = pMatrix.Columns.Item("colMulti").Cells.Item(Index).Specific.Value
            DOrigin = pMatrix.Columns.Item("colOrigin").Cells.Item(Index).Specific.Value
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SetDispatchDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)
        Try
            pForm.Items.Item("ed_DInsDoc").Specific.Value = DocumentNo
            pForm.Items.Item("ed_DPONo").Specific.Value = PurchaseOrderNo
            pForm.Items.Item("ed_DPDocNo").Specific.Value = PurchaseDocNo
            pForm.Items.Item("ed_DPO").Specific.Value = POSerialNo
            If Mode = "Internal" Then
                pForm.DataSources.UserDataSources.Item("DSINTR").ValueEx = "1"
                pForm.DataSources.UserDataSources.Item("DSEXTR").ValueEx = "2"
            ElseIf Mode = "External" Then
                pForm.DataSources.UserDataSources.Item("DSEXTR").ValueEx = "1"
                pForm.DataSources.UserDataSources.Item("DSINTR").ValueEx = "2"
            End If
            pForm.Items.Item("ed_DInDate").Specific.Value = InstructionDate
            pForm.Items.Item("ed_DspCode").Specific.Value = DspCode
            pForm.Items.Item("ed_Dspatch").Specific.Value = Dispatch

            pForm.Items.Item("ed_DEUC").Specific.Value = EUC
            pForm.Items.Item("ed_DAttent").Specific.Value = Attention
            pForm.Items.Item("ed_DspTel").Specific.Value = Telephone
            pForm.Items.Item("ed_DFax").Specific.Value = Fax
            pForm.Items.Item("ed_DEmail").Specific.Value = Email
            pForm.Items.Item("ed_DspDate").Specific.Value = DspDate
            pForm.Items.Item("ed_DspTime").Specific.Value = DspTime

            pForm.Items.Item("ee_DspIns").Specific.Value = DspIns
            pForm.Items.Item("ee_DIRmsk").Specific.Value = DspIRemarks
            pForm.Items.Item("ee_DRmsk").Specific.Value = Remarks
            pForm.Items.Item("ed_DPStus").Specific.Value = POStus
            pForm.Items.Item("ed_DDate").Specific.Value = PODate
            pForm.Items.Item("ed_DCreate").Specific.Value = PrepBy
            pForm.Items.Item("ed_DMulti").Specific.Value = DMulti
            pForm.Items.Item("ed_DOrigin").Specific.Value = DOrigin
        Catch ex As Exception

        End Try
    End Sub
    Private Sub EnabledDispatchForExternal(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)

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
            pForm.Items.Item("ed_Dspatch").Specific.Active = True
            pForm.Items.Item("ed_DEUC").Enabled = pValue
            pForm.Items.Item("ed_DAttent").Enabled = pValue
            pForm.Items.Item("ed_DspTel").Enabled = pValue
            pForm.Items.Item("ed_DFax").Enabled = pValue
            pForm.Items.Item("ed_DEmail").Enabled = pValue
            pForm.Items.Item("ed_Dspatch").Specific.Active = True

        Catch ex As Exception

        End Try
    End Sub
    Private Function EditPOInEditTab(ByVal ActiveForm As SAPbouiCOM.Form, ByVal PONo As String, ByVal matrixName As String) As Boolean
        EditPOInEditTab = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If matrixName = "mx_TkrList" Then
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_TkrIns = " + FormatString(ActiveForm.Items.Item("ee_TkrIns").Specific.Value) + ",U_TkrTo= " + FormatString(ActiveForm.Items.Item("ee_TkrTo").Specific.Value) + ",U_ColFrm= " + FormatString(ActiveForm.Items.Item("ee_ColFrm").Specific.Value) + ",U_PORMKS= " + FormatString(ActiveForm.Items.Item("ee_Rmsk").Specific.Value) + ",U_POIRMKS= " + FormatString(ActiveForm.Items.Item("ee_InsRmsk").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            ElseIf matrixName = "mx_DspList" Then
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_TkrIns = " + FormatString(ActiveForm.Items.Item("ee_DspIns").Specific.Value) + ",U_PORMKS= " + FormatString(ActiveForm.Items.Item("ee_DRmsk").Specific.Value) + ",U_POIRMKS= " + FormatString(ActiveForm.Items.Item("ee_DIRmsk").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            End If
            EditPOInEditTab = True
        Catch ex As Exception
            EditPOInEditTab = False
        End Try
    End Function

    Private Function EditPOTab(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String) As Boolean
        EditPOTab = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Try
            ObjDbDataSource = pForm.DataSources.DBDataSources.Item(tblName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If pMatrix.RowCount > 0 Then
                For i As Integer = 1 To pMatrix.RowCount
                    If pMatrix.Columns.Item("colPONo").Cells.Item(i).Specific.Value = oRecordSet.Fields.Item("U_PO").Value.ToString Then
                        currentRow = i
                        Exit For
                    End If
                Next
            End If
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                'If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                '    If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNo").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                '        pMatrix.Clear()
                '    End If
                'End If
                Do Until oRecordSet.EoF

                    With ObjDbDataSource
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PODocNo", .Offset, pMatrix.Columns.Item("colPODocNo").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PONo", .Offset, pMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_PODate", .Offset, pMatrix.Columns.Item("colPODate").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_VName", .Offset, pMatrix.Columns.Item("colVendor").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_Desc", .Offset, pMatrix.Columns.Item("colDesc").Cells.Item(currentRow).Specific.Value)
                        .SetValue("U_POStatus", .Offset, oRecordSet.Fields.Item("U_POStatus").Value)
                    End With
                    pMatrix.SetLineData(currentRow)

                    oRecordSet.MoveNext()
                Loop
            End If
            EditPOTab = True
        Catch ex As Exception
            EditPOTab = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

#End Region
    Public Sub LoadMultiJobNormalForm(ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String, ByVal Source As String)

        ' **********************************************************************************
        '   Function    :   LoadButtonForm()
        '   Purpose     :   This function provide to Load and show Button Forms when Button clikede of main form  
        '   Parameters  :   ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String,
        '               :   ByVal KeyName As String 
        '   return      :   No          
        ' **********************************************************************************

        Dim oActiveForm As SAPbouiCOM.Form
        Dim sErrDesc As String = ""
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim sFuncName As String = "SBO_Application_MenuEvent()"
        ' Dim oCombo, oComboMain As SAPbouiCOM.ComboBox
        If Not p_oDICompany.Connected Then
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
            If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        End If
        LoadFromXML(p_oSBOApplication, FormName)
        oActiveForm = p_oSBOApplication.Forms.ActiveForm
        oActiveForm.Freeze(True)
        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

        oActiveForm.EnableMenu("1288", True)
        oActiveForm.EnableMenu("1289", True)
        oActiveForm.EnableMenu("1290", True)
        oActiveForm.EnableMenu("1291", True)
        oActiveForm.EnableMenu("771", False)
        oActiveForm.EnableMenu("774", False)
        oActiveForm.EnableMenu("1284", False)
        oActiveForm.EnableMenu("1286", False)
        oActiveForm.EnableMenu("1283", False)
        oActiveForm.EnableMenu("772", False)
        oActiveForm.EnableMenu("4870", False)

        If AddUserDataSrc(oActiveForm, "FROMDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "TODATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "FILTER", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "CUST", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "MJOBNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "SOURCE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "MAINJOB", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


        oActiveForm.Items.Item("ed_frmDate").Specific.DataBind.SetBound(True, "", "FROMDATE")
        oActiveForm.Items.Item("ed_toDate").Specific.DataBind.SetBound(True, "", "TODATE")
        oActiveForm.Items.Item("cb_filter").Specific.DataBind.SetBound(True, "", "FILTER")
        oActiveForm.Items.Item("ed_Cust").Specific.DataBind.SetBound(True, "", "CUST")
        oActiveForm.Items.Item("ed_Job").Specific.DataBind.SetBound(True, "", "MJOBNO")
        oActiveForm.Items.Item("ed_Source").Specific.DataBind.SetBound(True, "", "SOURCE")
        oActiveForm.Items.Item("ed_MainJob").Specific.DataBind.SetBound(True, "", "MAINJOB")
        oCombo = oActiveForm.Items.Item("cb_filter").Specific
        oCombo.ValidValues.Add("General Cargo", "General Cargo")
        oCombo.ValidValues.Add("Explosive", "Explosive")
        oCombo.ValidValues.Add("Radioactive", "Radioactive")
        oCombo.ValidValues.Add("DG", "DG")
        oCombo.ValidValues.Add("Strategic", "Strategic")
        oActiveForm.Items.Item("ed_toDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oActiveForm.Items.Item("ed_Source").Specific.Value = Source
        oActiveForm.Items.Item("ed_MainJob").Specific.Value = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Freeze(False)


    End Sub
    Public Sub LoadDetachJobNormalForm(ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String, ByVal Source As String, ByVal PO As String)

        ' **********************************************************************************
        '   Function    :   LoadButtonForm()
        '   Purpose     :   This function provide to Load and show Button Forms when Button clikede of main form  
        '   Parameters  :   ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String,
        '               :   ByVal KeyName As String 
        '   return      :   No          
        ' **********************************************************************************

        Dim oActiveForm As SAPbouiCOM.Form
        Dim sErrDesc As String = ""
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim sFuncName As String = "SBO_Application_MenuEvent()"
        ' Dim oCombo, oComboMain As SAPbouiCOM.ComboBox
        If Not p_oDICompany.Connected Then
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
            If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        End If
        LoadFromXML(p_oSBOApplication, FormName)
        oActiveForm = p_oSBOApplication.Forms.ActiveForm
        oActiveForm.Freeze(True)
        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

        oActiveForm.EnableMenu("1288", True)
        oActiveForm.EnableMenu("1289", True)
        oActiveForm.EnableMenu("1290", True)
        oActiveForm.EnableMenu("1291", True)
        oActiveForm.EnableMenu("771", False)
        oActiveForm.EnableMenu("774", False)
        oActiveForm.EnableMenu("1284", False)
        oActiveForm.EnableMenu("1286", False)
        oActiveForm.EnableMenu("1283", False)
        oActiveForm.EnableMenu("772", False)
        oActiveForm.EnableMenu("4870", False)

        oActiveForm.Items.Item("ed_Source").Specific.Value = Source
        oActiveForm.Items.Item("ed_MainJob").Specific.Value = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Items.Item("ed_PO").Specific.Value = parentForm.Items.Item(PO).Specific.Value
        oActiveForm.Freeze(False)


    End Sub

    Public Sub PreviewOutLookMail(ByRef oActiveForm As SAPbouiCOM.Form, ByVal source As String)

        ' **********************************************************************************
        '   Function    :   PreviewOutLookMail()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************


        Dim oOutL As New Microsoft.Office.Interop.Outlook.Application
        Dim oMail As Microsoft.Office.Interop.Outlook.MailItem
        Dim oInsp As Microsoft.Office.Interop.Outlook.Inspector
        Dim oMatrix As SAPbouiCOM.Matrix

        oMail = oOutL.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
        Dim strDate As String
        strDate = IIf(oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString = "", Today.Date.ToString("yyyyMMdd"), oActiveForm.Items.Item("ed_FJbDate").Specific.Value.ToString)
        oMail.Subject = oActiveForm.Items.Item("ed_PO").Specific.Value.ToString
        Dim msg As String = ""
        If source = "Fumigation" Then
            msg = "Hi " + oActiveForm.Items.Item("ed_FCntact").Specific.Value + "," & Chr(13) & _
              oActiveForm.Items.Item("ed_FName").Specific.Value & Chr(13) & Chr(13) & _
                               "As spoken earlier , pls help to arrange fumigation on " & strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4) + " (" + oActiveForm.Items.Item("ed_FJbTime").Specific.Value + ") " & Chr(13) & _
                                "at below address for " & Chr(13) & _
                                 oActiveForm.Items.Item("ed_Item").Specific.Value & ",thanks." & Chr(13) & Chr(13) & _
                                 oActiveForm.Items.Item("ed_Loc").Specific.Value & Chr(13) & _
                                 "Attn:" + oActiveForm.Items.Item("ed_FSIA").Specific.Value & Chr(13) & _
                                 "Tel:" + oActiveForm.Items.Item("ed_SIATel").Specific.Value & Chr(13) & Chr(13) & Chr(13)
        ElseIf source = "Crane" Then
            msg = "Hi " + oActiveForm.Items.Item("ed_FCntact").Specific.Value + "," & Chr(13) & _
                    oActiveForm.Items.Item("ed_FName").Specific.Value & Chr(13) & Chr(13) & _
                    "As spoken we need "
            oMatrix = oActiveForm.Items.Item("mx_CDetail").Specific
            For i As Integer = 1 To oMatrix.RowCount
                msg = msg + oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value() & " " & oMatrix.Columns.Item("colCType").Cells.Item(i).Specific.Value() & Chr(13)
            Next
            msg = msg + "on " & strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4) & _
                " at " & oMatrix.Columns.Item("colHr").Cells.Item(1).Specific.Value() & " " & oActiveForm.Items.Item("ed_Loc").Specific.Value & Chr(13) & _
                "Please indicate Job # " & oActiveForm.Items.Item("ed_JobNo").Specific.Value & " on your billing invoice " & Chr(13) & Chr(13) & Chr(13)
        ElseIf source = "Forklift" Then 'to combine
            msg = "Hi " + oActiveForm.Items.Item("ed_FCntact").Specific.Value + "," & Chr(13) & _
                    oActiveForm.Items.Item("ed_FName").Specific.Value & Chr(13) & Chr(13) & _
                    "Kindly arrange forklift as per follows:-" & Chr(13) & Chr(13)
            oMatrix = oActiveForm.Items.Item("mx_CDetail").Specific
            For i As Integer = 1 To oMatrix.RowCount
                msg = msg + "Tonnage/Type: " + oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value() & " " & oMatrix.Columns.Item("colDesc").Cells.Item(i).Specific.Value() & Chr(13)
            Next
            msg = msg + "Date Required: " & strDate.Substring(6, 2) + "-" + strDate.Substring(4, 2) + "-" + strDate.Substring(0, 4) & Chr(13) & _
                "Time: " & oActiveForm.Items.Item("ed_FJbTime").Specific.Value & Chr(13) & _
                "Location: " + oActiveForm.Items.Item("ed_Loc").Specific.Value & Chr(13) & _
                "Our Job File No " & oActiveForm.Items.Item("ed_JobNo").Specific.Value & Chr(13) & _
                "Contact Person: " & oActiveForm.Items.Item("ed_FSIA").Specific.Value & " " & oActiveForm.Items.Item("ed_SIATel").Specific.Value & Chr(13) & Chr(13) & Chr(13)
        ElseIf source = "Crate" Then 'to combine
            msg = "Hi " + oActiveForm.Items.Item("ed_FCntact").Specific.Value + "," & Chr(13) & _
                    oActiveForm.Items.Item("ed_FName").Specific.Value & Chr(13) & Chr(13) & _
                    "Please help to make " + oActiveForm.Items.Item("ed_Desc").Specific.Value + " and the internal dimension as below:" & Chr(13) & Chr(13)
            oMatrix = oActiveForm.Items.Item("mx_CDetail").Specific
            For i As Integer = 1 To oMatrix.RowCount
                msg = msg + "Internal Dimension: " & Chr(13) & Chr(13) & _
                    oMatrix.Columns.Item("colDimen").Cells.Item(i).Specific.Value() & " - " & oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value().ToString() & " Pcs" & Chr(13)

            Next
        End If

       

        'SigString = "C:\Users\" & Environ("username") & _
        '"\AppData\Roaming\Microsoft\Signatures\Best Regards.txt"

        'If Dir(SigString) <> "" Then
        '    Signature = GetBoiler(SigString)
        'Else
        '    Signature = ""
        'End If
        'oMail.Body = msg & vbNewLine & vbNewLine & Signature
        ''if this line of code is included the defaul signature does not show up when displayed
        'oMail.Body = "My Body"

        oInsp = oMail.GetInspector
        oMail.Display()
        oMail.Body = msg & oMail.Body

    End Sub

    Public Sub SavePOPDFInEditTab(ByRef oActiveForm As SAPbouiCOM.Form, ByVal source As String)

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
        jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Purchase Order.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        'If Not ParentForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
        '    PONo = Convert.ToInt32(ParentForm.Items.Item("ed_PONo").Specific.Value)
        'Else
        '    PONo = DocLastKey
        'End If
        If source = "Trucking" Then
            PONo = Convert.ToInt32(oActiveForm.Items.Item("ed_PONo").Specific.Value)
        ElseIf source = "Dispatch" Then
            PONo = Convert.ToInt32(oActiveForm.Items.Item("ed_DPONo").Specific.Value)
        Else
            PONo = Convert.ToInt32(oActiveForm.Items.Item("ed_PONo").Specific.Value)
        End If

        'pdffilepath = "C:\Users\UNIQUE\Desktop\For NL\Outrider.pdf"

        rptDocument.SetParameterValue("@DocEntry", PONo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, False)
        End If

    End Sub

    Private Sub SaveInsDoc(ByRef ParentForm As SAPbouiCOM.Form)

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


        If ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-LCL") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Air") Or _
             ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Land") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-LCL") Or _
            ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Air") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Land") Or _
            ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Transhipment") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Local") Then

            rptPath = Application.StartupPath.ToString & "\Trucking Instruction.rpt"
        ElseIf ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Import Sea-FCL") Or ParentForm.Items.Item("ed_JType").Specific.Value.ToString.Contains("Export Sea-FCL") Then
            rptPath = Application.StartupPath.ToString & "\Trucking Instruction FCL.rpt"
        End If

        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        InsDoc = Convert.ToInt32(ParentForm.Items.Item("ed_InsDoc").Specific.Value)

        rptDocument.SetParameterValue("@DocEntry", DocNum)
        rptDocument.SetParameterValue("@InsDocNo", InsDoc)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, False)
        End If
    End Sub

    Private Sub SaveDispatchInstruction(ByRef ParentForm As SAPbouiCOM.Form)

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
        pdfFilename = "Dispatch Instruction"
        mainFolder = p_fmsSetting.DocuPath
        jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        rptPath = Application.StartupPath.ToString & "\Dispatch Instruction.rpt"
        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        rptDocument.Load(rptPath)
        rptDocument.Refresh()
        DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        InsDoc = Convert.ToInt32(ParentForm.Items.Item("ed_DInsDoc").Specific.Value)
        rptDocument.SetParameterValue("@DocEntry", DocNum)
        rptDocument.SetParameterValue("@InsDocNo", InsDoc)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, False)
        End If
    End Sub


#Region "=========== FUMIGATION ================="
    Private Sub LoadFumigation(ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String, ByVal KeyName As String)
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim sErrDesc As String = String.Empty
        Dim oCombo As SAPbouiCOM.ComboBox
        MainForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)


        LoadFromXML(p_oSBOApplication, FormName)
        oActiveForm = p_oSBOApplication.Forms.ActiveForm
        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
        oActiveForm.Title = KeyName + " " + MainForm.Title

        oActiveForm.Freeze(True)

        oActiveForm.EnableMenu("1288", True)
        oActiveForm.EnableMenu("1289", True)
        oActiveForm.EnableMenu("1290", True)
        oActiveForm.EnableMenu("1291", True)
        oActiveForm.EnableMenu("1284", False)
        oActiveForm.EnableMenu("1286", False)
        oActiveForm.EnableMenu("1283", False)
        oActiveForm.EnableMenu("771", False)
        oActiveForm.EnableMenu("774", False)

        oActiveForm.AutoManaged = True
        oActiveForm.DataBrowser.BrowseBy = "ed_DocNum"
        If KeyName = "CRANE" Or KeyName = "FORKLIFT" Or KeyName = "CRATE" Then 'BUTTON
            modFumigation.CreateDTDetail(oActiveForm, KeyName) ' to combine
        ElseIf KeyName = "BUNKER" Then
            modFumigation.CreateDTDetailBunker(oActiveForm)
        ElseIf KeyName = "TOLL" Then
            modFumigation.CreateDTDetailToll(oActiveForm)
        End If
        oActiveForm.Items.Item("fo_View").Specific.Select()


        If AddUserDataSrc(oActiveForm, "VCODE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "VNAME", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "SIA", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "SIACODE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "CPERSON", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "VREF", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "SIATEL", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "PO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "POSTATUS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "CREATE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "PODOCNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "PONO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
     
        If AddUserDataSrc(oActiveForm, "REMARK", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        If AddUserDataSrc(oActiveForm, "JDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "JTIME", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "PODATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "JobNo", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "OriginPO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(oActiveForm, "MultiJob", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)


        oActiveForm.Items.Item("ed_FCode").Specific.DataBind.SetBound(True, "", "VCODE")
        oActiveForm.Items.Item("ed_FName").Specific.DataBind.SetBound(True, "", "VNAME")
        oActiveForm.Items.Item("ed_FSIA").Specific.DataBind.SetBound(True, "", "SIA")
        oActiveForm.Items.Item("ed_SIACode").Specific.DataBind.SetBound(True, "", "SIACODE")
        oActiveForm.Items.Item("ed_PO").Specific.DataBind.SetBound(True, "", "PO")
        oActiveForm.Items.Item("ed_FPOStus").Specific.DataBind.SetBound(True, "", "POSTATUS")
        oActiveForm.Items.Item("ed_Create").Specific.DataBind.SetBound(True, "", "CREATE")
        oActiveForm.Items.Item("ed_PODocNo").Specific.DataBind.SetBound(True, "", "PODOCNO")
        oActiveForm.Items.Item("ed_PONo").Specific.DataBind.SetBound(True, "", "PONO")
        oActiveForm.Items.Item("ed_SIATel").Specific.DataBind.SetBound(True, "", "SIATEL")
        oActiveForm.Items.Item("ed_Remark").Specific.DataBind.SetBound(True, "", "REMARK")
        oActiveForm.Items.Item("ed_FVRef").Specific.DataBind.SetBound(True, "", "VREF")
        oActiveForm.Items.Item("ed_FCntact").Specific.DataBind.SetBound(True, "", "CPERSON")
        oActiveForm.Items.Item("ed_FJbDate").Specific.DataBind.SetBound(True, "", "JDATE")
        oActiveForm.Items.Item("ed_FJbTime").Specific.DataBind.SetBound(True, "", "JTIME")
        oActiveForm.Items.Item("ed_FPODate").Specific.DataBind.SetBound(True, "", "PODATE")
        oActiveForm.Items.Item("ed_JDeNo").Specific.DataBind.SetBound(True, "", "JobNo")
        oActiveForm.Items.Item("ed_OriPO").Specific.DataBind.SetBound(True, "", "OriginPO")
        oActiveForm.Items.Item("ed_MultiJb").Specific.DataBind.SetBound(True, "", "MultiJob")

        'to km
        If KeyName = "FUMIGATION" Then
            If AddUserDataSrc(oActiveForm, "ITEM", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Item").Specific.DataBind.SetBound(True, "", "ITEM")
            If AddUserDataSrc(oActiveForm, "LOC", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Loc").Specific.DataBind.SetBound(True, "", "LOC")
            oActiveForm.Items.Item("ed_ICode").Specific.Value = "4809"
            oActiveForm.Items.Item("ed_IDesc").Specific.Value = "Fumigation Charges for wooden boxes"
            oActiveForm.Items.Item("ed_IQty").Specific.Value = "1"
            oActiveForm.Items.Item("ed_IPrice").Specific.Value = "1"
        ElseIf KeyName = "CRANE" Then
            If AddUserDataSrc(oActiveForm, "SIns", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "LOC", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Loc").Specific.DataBind.SetBound(True, "", "LOC")
            oActiveForm.Items.Item("ed_SIns").Specific.DataBind.SetBound(True, "", "SIns")
            oActiveForm.Items.Item("ed_ICode").Specific.Value = "6504"
            oActiveForm.Items.Item("ed_IDesc").Specific.Value = "Crane - 20 Tonne (Min 4 hrs)"
        ElseIf KeyName = "FORKLIFT" Then 'to combine
            If AddUserDataSrc(oActiveForm, "IRemark", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "LOC", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_IRemark").Specific.DataBind.SetBound(True, "", "IRemark")
            oActiveForm.Items.Item("ed_Loc").Specific.DataBind.SetBound(True, "", "LOC")
            oActiveForm.Items.Item("ed_ICode").Specific.Value = "6601"
            oActiveForm.Items.Item("ed_IDesc").Specific.Value = "Forklift Charges"
        ElseIf KeyName = "CRATE" Then 'to combine

            If AddUserDataSrc(oActiveForm, "DESC", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Desc").Specific.DataBind.SetBound(True, "", "DESC")
            oActiveForm.Items.Item("ed_ICode").Specific.Value = "6101"
            oActiveForm.Items.Item("ed_IDesc").Specific.Value = "Supply - Wooden Crate"
        ElseIf KeyName = "OUTRIDER" Then
            If AddUserDataSrc(oActiveForm, "LOCFROM", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "LOCTO", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "IREMARK", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_LocFrom").Specific.DataBind.SetBound(True, "", "LOCFROM")
            oActiveForm.Items.Item("ed_LocTo").Specific.DataBind.SetBound(True, "", "LOCTO")
            oActiveForm.Items.Item("ed_IRemark").Specific.DataBind.SetBound(True, "", "IREMARK")
            oActiveForm.Items.Item("ed_ICode").Specific.Value = "6301"
            oActiveForm.Items.Item("ed_IDesc").Specific.Value = "Outrider Escort -1st hr"
            oActiveForm.Items.Item("ed_IQty").Specific.Value = "1"
            oActiveForm.Items.Item("ed_IPrice").Specific.Value = "1"
        ElseIf KeyName = "BUNKER" Then
            If AddUserDataSrc(oActiveForm, "ACTIVITY", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "STORE1", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "STORE1A", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "STORE2", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "STORE2A", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "CDESC", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "SIns", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If AddUserDataSrc(oActiveForm, "TQTY", sErrDesc, SAPbouiCOM.BoDataType.dt_QUANTITY) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKGS", sErrDesc, SAPbouiCOM.BoDataType.dt_QUANTITY) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TM3", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TNEQ", sErrDesc, SAPbouiCOM.BoDataType.dt_RATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            oActiveForm.Items.Item("cb_Act").Specific.DataBind.SetBound(True, "", "ACTIVITY")
            oActiveForm.Items.Item("chk_1").Specific.DataBind.SetBound(True, "", "STORE1")
            oActiveForm.Items.Item("chk_1a").Specific.DataBind.SetBound(True, "", "STORE1A")
            oActiveForm.Items.Item("chk_2").Specific.DataBind.SetBound(True, "", "STORE2")
            oActiveForm.Items.Item("chk_2a").Specific.DataBind.SetBound(True, "", "STORE2A")
            oActiveForm.Items.Item("ed_CDesc").Specific.DataBind.SetBound(True, "", "CDESC")

            oActiveForm.Items.Item("ed_TQty").Specific.DataBind.SetBound(True, "", "TQTY")
            oActiveForm.Items.Item("ed_TKgs").Specific.DataBind.SetBound(True, "", "TKGS")
            oActiveForm.Items.Item("ed_TM3").Specific.DataBind.SetBound(True, "", "TM3")
            oActiveForm.Items.Item("ed_TNEQ").Specific.DataBind.SetBound(True, "", "TNEQ")
            oActiveForm.Items.Item("ed_ICode").Specific.Value = "7301"
            oActiveForm.Items.Item("ed_IDesc").Specific.Value = "Bunker - Opening (Explosive) (Weekday)"
            oActiveForm.Items.Item("ed_IPrice").Specific.Value = "1"
            oActiveForm.Items.Item("ed_SIns").Specific.DataBind.SetBound(True, "", "SIns")
            oCombo = oActiveForm.Items.Item("cb_Act").Specific
            oCombo.ValidValues.Add("", "")
            oCombo.ValidValues.Add("Delivery", "Delivery")
            oCombo.ValidValues.Add("Collection", "Collection")
            oCombo.ValidValues.Add("Prepare for Shipment", "Prepare for Shipment")
            oCombo.ValidValues.Add("Auditing", "Auditing")
            oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
        ElseIf KeyName = "TOLL" Then
            If AddUserDataSrc(oActiveForm, "LOC", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_Loc").Specific.DataBind.SetBound(True, "", "LOC")
            If AddUserDataSrc(oActiveForm, "IREMARK", sErrDesc, SAPbouiCOM.BoDataType.dt_LONG_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_IRemark").Specific.DataBind.SetBound(True, "", "IREMARK")
        End If
     

        If AddChooseFromList(oActiveForm, "cflFBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddChooseFromList(oActiveForm, "cflFBP2", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        oActiveForm.Items.Item("ed_FCode").Specific.ChooseFromListUID = "cflFBP"
        oActiveForm.Items.Item("ed_FCode").Specific.ChooseFromListAlias = "CardCode"
        oActiveForm.Items.Item("ed_FName").Specific.ChooseFromListUID = "cflFBP2"
        oActiveForm.Items.Item("ed_FName").Specific.ChooseFromListAlias = "CardName"

        AddChooseFromList(oActiveForm, "cflSIA", False, 171)
        AddChooseFromList(oActiveForm, "cflSIACode", False, 171)
        oActiveForm.Items.Item("ed_FSIA").Specific.ChooseFromListUID = "cflSIA"
        oActiveForm.Items.Item("ed_FSIA").Specific.ChooseFromListAlias = "firstName"
        oActiveForm.Items.Item("ed_SIACode").Specific.ChooseFromListUID = "cflSIACode"
        oActiveForm.Items.Item("ed_SIACode").Specific.ChooseFromListAlias = "empID"

        If KeyName <> "TOLL" Then
            AddChooseFromList(oActiveForm, "cflICODE", False, 4)
            AddChooseFromList(oActiveForm, "cflIDESC", False, 4)
            oActiveForm.Items.Item("ed_ICode").Specific.ChooseFromListUID = "cflICODE"
            oActiveForm.Items.Item("ed_ICode").Specific.ChooseFromListAlias = "ItemCode"
            oActiveForm.Items.Item("ed_IDesc").Specific.ChooseFromListUID = "cflIDESC"
            oActiveForm.Items.Item("ed_IDesc").Specific.ChooseFromListAlias = "ItemName"
        End If
        oActiveForm.DataSources.UserDataSources.Item("POSTATUS").ValueEx() = "Open"
        oActiveForm.Items.Item("ed_FPODate").Specific.Value = Today.Date.ToString("yyyyMMdd")

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        jobNo = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Items.Item("ed_JobNo").Specific.Value = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Items.Item("ed_JDeNo").Specific.Value = parentForm.Items.Item("ed_JobNo").Specific.Value
        oActiveForm.Items.Item("ed_DocNum").Specific.Value = parentForm.Items.Item("ed_DocNum").Specific.Value
        oActiveForm.Items.Item("ed_DocID").Specific.Value = GetNewKey(KeyName, oRecordSet)
        oActiveForm.Items.Item("ed_Create").Specific.Value = parentForm.Items.Item("ed_UserID").Specific.Value

        oActiveForm.Freeze(False)




    End Sub

    Private Function PopulatePurchaseHeaderButton(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal TableName As String, ByVal ProcessedState As Boolean, ByRef oPOForm As SAPbouiCOM.Form, ByVal source As String) As Boolean
        PopulatePurchaseHeaderButton = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Try
            ObjDbDataSource = pForm.DataSources.DBDataSources.Item(TableName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If pMatrix.RowCount > 0 Then
                        If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNo").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                            pMatrix.Clear()
                        End If
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
                            .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_Vendor", .Offset, oRecordSet.Fields.Item("U_VCode").Value.ToString)
                            .SetValue("U_Vname", .Offset, oRecordSet.Fields.Item("U_VName").Value.ToString)
                            .SetValue("U_CPerson", .Offset, oRecordSet.Fields.Item("U_CPerson").Value.ToString)
                            .SetValue("U_VRef", .Offset, oRecordSet.Fields.Item("U_VRef").Value.ToString)
                            .SetValue("U_SIA", .Offset, oRecordSet.Fields.Item("U_SInA").Value.ToString)
                            .SetValue("U_JDate", .Offset, CDate(oRecordSet.Fields.Item("U_TDate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_JTime", .Offset, oRecordSet.Fields.Item("U_TTime").Value.ToString)
                            If source = "Fumigation" Or source = "Crane" Then
                                .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            ElseIf source = "Outrider" Then
                                .SetValue("U_LocFrom", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                                .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                            ElseIf source = "Forklift" Then 'to combine
                                .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                                .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                            End If
                            If oPOForm.Title.Contains("Goods Receipt") = True Then 'MSW To Edit New Ticket
                                .SetValue("U_PODate", .Offset, oPOForm.Items.Item("ed_GRDate").Specific.Value)

                            Else
                                .SetValue("U_PODate", .Offset, oPOForm.Items.Item("ed_PODate").Specific.Value)

                            End If
                            .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                            .SetValue("U_Remark", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)

                        End With
                        pMatrix.AddRow()
                    Else
                        With ObjDbDataSource
                            .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(currentRow).Specific.Value)
                            .SetValue("U_PODocNo", .Offset, oRecordSet.Fields.Item("DocEntry").Value.ToString)
                            .SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
                            .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_Vendor", .Offset, oRecordSet.Fields.Item("U_VCode").Value.ToString)
                            .SetValue("U_VRef", .Offset, oRecordSet.Fields.Item("U_VRef").Value.ToString)
                            .SetValue("U_SIA", .Offset, oRecordSet.Fields.Item("U_SInA").Value.ToString)
                            .SetValue("U_JDate", .Offset, CDate(oRecordSet.Fields.Item("U_TDate").Value).ToString("yyyyMMdd"))
                            .SetValue("U_JTime", .Offset, oRecordSet.Fields.Item("U_TTime").Value.ToString)
                            If source = "Fumigation" Or source = "Crane" Then
                                .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            ElseIf source = "Outrider" Then
                                .SetValue("U_LocFrom", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                                .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                            ElseIf source = "Toll" Then
                                .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                                .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                            End If
                            If oPOForm.Title.Contains("Goods Receipt") = True Then 'MSW To Edit New Ticket
                                .SetValue("U_PODate", .Offset, oPOForm.Items.Item("ed_GRDate").Specific.Value)

                            Else
                                .SetValue("U_PODate", .Offset, oPOForm.Items.Item("ed_PODate").Specific.Value)

                            End If
                            .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                            .SetValue("U_Remark", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                        End With
                        pMatrix.SetLineData(currentRow)
                    End If

                    oRecordSet.MoveNext()
                Loop
            End If
            PopulatePurchaseHeaderButton = True
        Catch ex As Exception
            PopulatePurchaseHeaderButton = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function PopulateOtherPurchaseHeaderButton(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal TableName As String, ByVal source As String) As Boolean
        PopulateOtherPurchaseHeaderButton = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Try
            ObjDbDataSource = pForm.DataSources.DBDataSources.Item(TableName)
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(pStrSQL)
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    If pMatrix.RowCount > 0 Then
                        If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNo").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
                            pMatrix.Clear()
                        End If
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
                        If source = "Fumigation" Or source = "Crane" Then
                            .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                        ElseIf source = "Outrider" Then
                            .SetValue("U_LocFrom", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                        ElseIf source = "Forklift" Then 'to combine
                            .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                        ElseIf source = "Toll" Then
                            .SetValue("U_Loc", .Offset, oRecordSet.Fields.Item("U_TPlace").Value.ToString)
                            .SetValue("U_IRemark", .Offset, oRecordSet.Fields.Item("U_POIRMKS").Value.ToString)
                        End If

                        .SetValue("U_PODate", .Offset, CDate(oRecordSet.Fields.Item("U_PODate").Value).ToString("yyyyMMdd"))
                        .SetValue("U_Status", .Offset, oRecordSet.Fields.Item("U_POStatus").Value.ToString)
                        .SetValue("U_Remark", .Offset, oRecordSet.Fields.Item("U_PORMKS").Value.ToString)
                    End With
                    pMatrix.SetLineData(currentRow)


                    oRecordSet.MoveNext()
                Loop
            End If
            PopulateOtherPurchaseHeaderButton = True
        Catch ex As Exception
            PopulateOtherPurchaseHeaderButton = False
            MessageBox.Show(ex.Message)
        End Try
    End Function
#End Region

   
    Private Function CheckPOandVoucherStatus(ByVal oActiveForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   CheckPOandVoucherStatus()
        '   Purpose     :   This function will be providing to check Job has still open PO and Draft Voucher
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByVal oActiveForm As SAPbouiCOM.Form
        '   Return      :   False- FAILURE
        '               :   True - SUCCESS
        ' **********************************************************************************


        CheckPOandVoucherStatus = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            sql = "Select * from [@OBT_TB01_POLIST] Where DocEntry='" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & "' And U_POStatus='Open'"
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                CheckPOandVoucherStatus = True
                Exit Function
                'Else
                '    sql = "Select * from [@OBT_FCL05_EVOUCHER] Where DocEntry='" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & "'"
                '    oRecordSet.DoQuery(sql)
                '    If oRecordSet.RecordCount > 0 Then
                '        CheckPOandVoucherStatus = True
                '        Exit Function
                '    End If
            End If

        Catch ex As Exception
            CheckPOandVoucherStatus = False
            MessageBox.Show(ex.Message)
        End Try

    End Function
    Public Function CreateDGDPDF(ByVal oActiveForm As SAPbouiCOM.Form) As Boolean

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


        Dim sErrDesc As String = ""
        Dim i As Integer = 0
        Dim dblPrice As Double = 0.0
        Dim itemCode As String = String.Empty
        Dim itemDesc As String = String.Empty
        Dim originpdf As String = String.Empty
        Dim reader As iTextSharp.text.pdf.PdfReader
        Dim pdfoutputfile As FileStream
        Dim formfiller As iTextSharp.text.pdf.PdfStamper
        CreateDGDPDF = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            pdfFilename = "DG Note P&M"
            originpdf = "DG Note P&M.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            pdfFilename = "IATA DGD computerized"
            originpdf = "IATA DGD computerized.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            pdfFilename = "IATA DGD manual"
            originpdf = "IATA DGD manual.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            pdfFilename = "IMO DGD 2"
            originpdf = "IMO DGD 2.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()


            pdfFilename = "IMO DGD 1"
            originpdf = "IMO DGD 1.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            pdfFilename = "Multimodal DGD form"
            originpdf = "Multimodal DGD form.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            pdfFilename = "Multimodat DGD form -2nd page"
            originpdf = "Multimodat DGD form -2nd page.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            pdfFilename = "PIL DGD"
            originpdf = "PIL DGD.pdf"
            mainFolder = p_fmsSetting.DocuPath
            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            reader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            pdfoutputfile = New FileStream(pdffilepath, System.IO.FileMode.Create)
            formfiller = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            formfiller.Close()
            reader.Close()

            Process.Start("explorer.exe", IO.Directory.GetParent(pdffilepath).ToString)


            CreateDGDPDF = True
        Catch ex As Exception
            CreateDGDPDF = False
            MessageBox.Show(ex.Message)
        End Try

    End Function


End Module
