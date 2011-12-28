Option Explicit On

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Drawing.Printing
Imports System.Threading

Module modExportSeaLCL

    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Private VocNo, VendorName, PayTo, PayType, BankName, CheqNo, Status, Currency, PaymentDate, GST, PrepBy, ExRate, Remark As String
    Private Total, GSTAmt, SubTotal, vocTotal, gstTotal As Double
    Private ShpInvoice, PO, Details, ShipTo, AWB, Box, Weight, POD, DesCargo As String
    Private lRectCode As Integer
    Private sPicName, sImgePath As String
    Private ActiveForm As SAPbouiCOM.Form
    Private vendorCode As String
    Private DocLastKey As String
    Private currentRow As Integer
    Private ActiveMatrix As String

    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String

    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sql As String = ""
    Dim strAPInvNo, strOutPayNo As String
    Dim DocStatus As String = String.Empty
    Private vComp As SAPbobsCOM.Company
    Dim vedCurCode As String = String.Empty


    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public Function DoExportSeaLCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
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
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oPOForm As SAPbouiCOM.Form = Nothing
        Dim oGRForm As SAPbouiCOM.Form = Nothing
        Dim FunctionName As String = "DoImportSeaLCLFormDataEvent"
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

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", FunctionName)
            Select Case BusinessObjectInfo.FormTypeEx

                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000043", "2000000050"  'MSW to Edit New Ticket
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oPOForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        'If BusinessObjectInfo.FormUID = "PURCHASEORDER" Then
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            'When action of BusinessObjectInfo is nicely done, need to do 2 tasks in action for Purcahse process
                            ' (1). need to show PopulatePurchaseHeader into the matrix of Main Export Form
                            ' (2). need to create PurchaseOrder into OPOR and POR1, related with main PurchaseProcess by using oPurchaseOrder document
                            If AlreadyExist("EXPORTSEALCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                            ElseIf AlreadyExist("EXPORTLAND") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                            End If

                            'oMatrix = oActiveForm.Items.Item("mx_Fumi").Specific
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                        "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If Not CreatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            If BusinessObjectInfo.FormTypeEx = "2000000020" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crate").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB020_CRATE", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000005" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fumi").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB021_FUMIGAT", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB022_BUNKER", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB023_ARMESCORT", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000043" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bok").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB19_BOOKING", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                If Not PopulateTruckingPOToEditTab(oActiveForm, sql, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            End If
                            'If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql) Then Throw New ArgumentException(sErrDesc)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value))
                            jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
                            If BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                CreatePOPDF(oActiveForm, "mx_Armed", "")
                            Else
                                SendAttachFile(oActiveForm, oPOForm)
                            End If

                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                            If AlreadyExist("EXPORTSEALCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                            ElseIf AlreadyExist("EXPORTLAND") Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                            End If
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                       "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If BusinessObjectInfo.FormTypeEx = "2000000020" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crate").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB020_CRATE", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000005" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fumi").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB021_FUMIGAT", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB022_BUNKER", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB023_ARMESCORT", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000043" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bok").Specific
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB19_BOOKING", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                       "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                       "from  [@OBT_TB08_FFCPO] a inner join [@OBT_TB006_ETRUCKING] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB006_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            End If
                            If Not UpdatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)

                        End If
                        'End If
                    End If

                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000044", "2000000051" 'Truck PO
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oGRForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("EXPORTSEALCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRLCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                        ElseIf AlreadyExist("EXPORTLAND") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                        End If
                        'oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            If Not CreateGoodsReceiptPO(oGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)

                            If BusinessObjectInfo.FormTypeEx = "2000000011" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crate").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB020_CRATE", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000007" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fumi").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB021_FUMIGAT", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000016" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB022_BUNKER", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000013" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB023_ARMESCORT", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000044" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bok").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeaderFromMain(oActiveForm, oMatrix, sql, "@OBT_TB19_BOOKING", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000051" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                       "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                       "from  [@OBT_TB08_FFCPO] a inner join [@OBT_TB006_ETRUCKING] b on a.DocEntry=b.U_PODocNo where a.U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB006_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            End If

                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then

                        End If
                    End If

                Case "2000000026", "2000000030", "2000000032", "2000000035", "2000000038", "2000000038", "2000000041"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oPOForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("EXPORTSEALCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRLCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                        ElseIf AlreadyExist("EXPORTLAND") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                        End If

                        Select Case BusinessObjectInfo.FormUID

                            Case "CRANEPURCHASEORDER"
                                oMatrixName = "mx_Crane"
                                TableName = "@OBT_TB33_CRANE"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                            Case "ForkLiftPURCHASEORDER"
                                oMatrixName = "mx_forklif"
                                TableName = "@OBT_TB34_FORKLIFT"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                            Case "CourierPURCHASEORDER"
                                oMatrixName = "mx_Courier"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                TableName = "@OBT_TB35_COURIER"
                            Case "DGLPPURCHASEORDER"
                                oMatrixName = "mx_DGLP"
                                TableName = "@OBT_TB37_DGLP"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                            Case "OutriderPURCHASEORDER"
                                oMatrixName = "mx_Out"
                                TableName = "@OBT_TB38_OUTRIDER"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                            Case "COOPURCHASEORDER"
                                oMatrixName = "mx_COO"
                                TableName = "@OBT_TB39_COO"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                        End Select
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            oMatrix = ExportSeaLCLForm.Items.Item(oMatrixName).Specific
                            Dim sql As String = "select DocEntry,U_PONo,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                               "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If Not CreatePurchaseOrder(oPOForm, "mx_Item") Then
                                BubbleEvent = False
                                Throw New ArgumentException(sErrDesc)
                            End If

                            If Not PopulatePurchaseHeader(ExportSeaLCLForm, oMatrix, sql, TableName, True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value))
                            If oMatrixName = "mx_DGLP" Or oMatrixName = "mx_Out" Or oMatrixName = "mx_COO" Then
                                CreatePOPDF(oPOForm, oMatrixName, "")
                            Else
                                SendAttachFile(ExportSeaLCLForm, oPOForm)
                            End If

                            'SendAttachFile("","")
                        End If
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                            'ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                            oMatrix = ExportSeaLCLForm.Items.Item(oMatrixName).Specific
                            If Not UpdatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                            If Not PopulatePurchaseHeader(ExportSeaLCLForm, oMatrix, sql, TableName, False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            '  ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                        End If
                    End If

                Case "2000000027", "2000000029", "2000000033", "2000000036", "2000000039", "2000000027", "2000000042"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oGRForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        If AlreadyExist("EXPORTSEALCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRLCL") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                        ElseIf AlreadyExist("EXPORTLAND") Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                        End If

                        'oActiveForm = p_oSBOApplication.Forms.ActiveForm
                        Select Case BusinessObjectInfo.FormUID
                            Case "CRANEGOODSRECEIPT"
                                oMatrixName = "mx_Crane"
                                TableName = "@OBT_TB33_CRANE"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                            Case "FORKLIFTGOODSRECEIPT"
                                oMatrixName = "mx_forklif"
                                TableName = "@OBT_TB34_FORKLIFT"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                            Case "COURIERGOODSRECEIPT"
                                oMatrixName = "mx_Courier"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                TableName = "@OBT_TB35_COURIER"
                            Case "DGLPGOODSRECEIPT"
                                oMatrixName = "mx_DGLP"
                                TableName = "@OBT_TB37_DGLP"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                            Case "OUTRIDERGOODSRECEIPT"
                                oMatrixName = "mx_Out"
                                TableName = "@OBT_TB38_OUTRIDER"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                            Case "COOGOODSRECEIPT"
                                oMatrixName = "mx_COO"
                                TableName = "@OBT_TB39_COO"
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                        End Select


                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            oMatrix = ExportSeaLCLForm.Items.Item(oMatrixName).Specific
                            If Not CreateGoodsReceiptPO(oGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())

                            ' ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(BusinessObjectInfo.FormTypeEx.ToString(), 1)
                            If Not PopulatePurchaseHeader(ExportSeaLCLForm, oMatrix, sql, TableName, False, oGRForm) Then Throw New ArgumentException(sErrDesc)

                            'If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            '    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            'End If
                        End If

                    End If

                    
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                        ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                        ExportSeaLCLForm.Close()
                    End If

                Case "SHIPPINGINV"
                    If BusinessObjectInfo.ActionSuccess = True Then
                        Try
                            If AlreadyExist("EXPORTSEALCL") Then
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                            ElseIf AlreadyExist("EXPORTLAND") Then
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                            End If

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
                    
                    If BusinessObjectInfo.ActionSuccess = True Then
                        'ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        If AlreadyExist("EXPORTSEALCL") Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                        ElseIf AlreadyExist("EXPORTAIRLCL") Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                        ElseIf AlreadyExist("EXPORTLAND") Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                        End If
                        oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
                        oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_TB009_EVOUCHER", True)
                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            AddUpdateVoucher(oPayForm, oMatrix, "@OBT_TB009_EVOUCHER", False)
                        End If

                    End If

                Case "EXPORTSEALCL", "EXPORTAIRLCL", "EXPORTLAND"
                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    ExportSeaLCLForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
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
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            Dim JobLastDocEntry As Integer
                            Dim ObjectCode As String = String.Empty

                            sql = "select top 1 Docentry from [@OBT_TB002_EXPORT] order by docentry desc"
                            oRecordSet.DoQuery(sql)
                            Dim FrDocEntry As Integer = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString)
                            Dim NewJobNo As String = GetJobNumber("EX")
                            sql = "select top 1 Docentry from [@OBT_FREIGHTDOCNO] order by docentry desc"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.Fields.Item("Docentry").Value.ToString = "" Then
                                JobLastDocEntry = 1
                            Else
                                JobLastDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString) + 1
                            End If

                            sql = "Update [@OBT_TB002_EXPORT] set U_JbDocNo=" & JobLastDocEntry & ",U_JobNum = '" & NewJobNo & "' Where DocEntry=" & FrDocEntry & ""
                            oRecordSet.DoQuery(sql)
                            p_oSBOApplication.SetStatusBarMessage("Actual Job Number is " & NewJobNo, SAPbouiCOM.BoMessageTime.bmt_Short, False)
                            sql = "Insert Into [@OBT_FREIGHTDOCNO] (DocEntry,DocNum,U_JobNo,U_JobMode,U_JobType,U_JbStus,U_FrDocNo,U_JbDate,U_ObjType,U_CusCode,U_CusName,U_ShpCode,U_ShpName) Values " & _
                                    "(" & JobLastDocEntry & _
                                        "," & JobLastDocEntry & _
                                       "," & IIf(NewJobNo <> "", FormatString(NewJobNo), "NULL") & _
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
                        End If
                    ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & ExportSeaLCLForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & ""
                            oRecordSet.DoQuery(sql)

                        End If
                    End If


                    'MSW New Button at Choose From List 22-03-2011
                Case "VESSEL"
                    Dim vesCode As String = String.Empty
                    Dim voyNo As String = String.Empty
                    Dim oExportSeaLCLForm As SAPbouiCOM.Form = Nothing
                    If AlreadyExist("EXPORTSEALCL") Then
                        oExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                    ElseIf AlreadyExist("EXPORTAIRLCL") Then
                        oExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                    ElseIf AlreadyExist("EXPORTLAND") Then
                        oExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                    End If
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select top 1* from [@OBT_TB018_VESSEL] order by DocEntry desc")
                            If oRecordSet.RecordCount > 0 Then
                                vesCode = oRecordSet.Fields.Item("Name").Value.ToString
                                voyNo = oRecordSet.Fields.Item("U_Voyage").Value.ToString
                            End If
                            oExportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = vesCode
                            oExportSeaLCLForm.Items.Item("ed_Voy").Specific.Value = voyNo

                            'Try
                            '    oImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Active = True
                            'Catch ex As Exception

                            'End Try

                        End If
                    Else
                        If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE) And BusinessObjectInfo.BeforeAction = False Then
                            If BusinessObjectInfo.ActionSuccess = True Then
                                oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                                oExportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = ""
                                oExportSeaLCLForm.Items.Item("ed_Voy").Specific.Value = ""

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
    Dim dtmatrix As SAPbouiCOM.DataTable
    Private gridindex As String

    Public Function DoExportSeaLCLMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
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
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oShpForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oOpt As SAPbouiCOM.OptionBtn = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing
        Dim SqlQuery As String = String.Empty
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim FunctionName As String = "DoImportSeaLCLMenuEvent()"
        Dim oMatrixName As String = ""
        Dim TableName As String = ""

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            Select Case pVal.MenuUID
                Case "mnuExportSeaLCL"
                    If pVal.BeforeAction = False Then
                        LoadExportSeaLCLForm()
                    End If

                Case "1281"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx).Items.Item("ed_JobNo").Enabled = True
                            p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx).Items.Item("ed_JobNo").Specific.Active = True
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ExportSeaLCLForm.Items.Item("ch_POD").Enabled = False

                            End If
                            If AddChooseFromListByOption(p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx), True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
                        ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            If pVal.BeforeAction = True Then
                                BubbleEvent = False
                            End If
                            ''-------------------------For Payment(omm)------------------------------------------'
                            ''If (ImportSeaLCLForm.PaneLevel = 21) Then
                            'oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
                            'RowAddToMatrix(ExportSeaLCLForm, oMatrix)
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
                     p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000026" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000030" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000050" Then

                        ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                        'If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oMatrix = ExportSeaLCLForm.Items.Item("mx_Item").Specific
                        ' Dim lRow As Long
                        If pVal.BeforeAction = True Then
                            If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                BubbleEvent = False
                            End If

                            If BubbleEvent = True Then
                                DeleteMatrixRow(ExportSeaLCLForm, oMatrix, "@OBT_TB09_FFCPOITEM", "LineId")
                                BubbleEvent = False
                                If Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                                CalculateTotalPO(ExportSeaLCLForm, oMatrix)
                            End If
                        End If

                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                        If pVal.BeforeAction = True Then
                            ' BubbleEvent = False
                            'MSW To Edit
                            Dim shpDocEntry As String
                            ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                            oMatrix = ExportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                            If oMatrix.GetNextSelectedRow > 0 Then
                                If oMatrix.Columns.Item("colDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString <> "" Then
                                    shpDocEntry = oMatrix.Columns.Item("colSDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString
                                    DeleteMatrixRow(ExportSeaLCLForm, oMatrix, "@OBT_LCL18_SHPINV", "V_-1")
                                    DeleteUDO(ExportSeaLCLForm, "SHIPPINGINV", shpDocEntry)
                                    ExportSeaLCLForm.Items.Item("1").Click()
                                    If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                End If
                            End If
                            BubbleEvent = False
                            'End MSW To Edit
                        End If
                        
                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Or _
                         p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000007" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000011" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000012" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000013" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000014" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000016" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000044" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000051" Then
                        If pVal.BeforeAction = True Then
                            BubbleEvent = False
                        End If
                    End If

                Case "1282"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            EnabledHeaderControls(ExportSeaLCLForm, False) '25-3-2011
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'EXPORTSEALCL'")
                            If oRecordSet.RecordCount > 0 Then
                                ' ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                                ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                            End If
                            ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("EX")
                            If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Then
                                ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Sea LCL" 'MSW 08-06-2011 for Job Type Table
                            ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Then
                                ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Air" 'MSW 08-06-2011 for Job Type Table
                            End If
                            ExportSeaLCLForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString
                            ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            ExportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                            ExportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            ExportSeaLCLForm.Items.Item("ch_POD").Enabled = False
                        End If
                    End If

                Case "1288", "1289", "1290", "1291"
                    If pVal.BeforeAction = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.Item(p_oSBOApplication.Forms.ActiveForm.TypeEx)
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If
                        End If
                    End If
                Case "EditVoc"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(p_oSBOApplication.Forms.ActiveForm.TypeEx, 1)
                            LoadPaymentVoucher(ExportSeaLCLForm)
                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ExportSeaLCLForm.Items.Item("mx_Voucher").Specific
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
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(p_oSBOApplication.Forms.ActiveForm.TypeEx, 1)
                            LoadShippingInvoice(ExportSeaLCLForm)
                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = ExportSeaLCLForm.Items.Item("mx_ShpInv").Specific
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
                        Select Case ActiveMatrix
                            Case "mx_Crane"
                                RPOsrfname = "CranePurchaseOrder_Form.srf"
                            Case "mx_Armed"
                                RPOsrfname = "ArmedPurchaseOrder.srf"
                            Case "mx_FumiPO"
                                RPOsrfname = "PurchaseOrder.srf"
                            Case "mx_Courier"
                                RPOsrfname = "CourierPurchaseOrder.srf"
                            Case "mx_DGLP"
                                RPOsrfname = "DGLPPurchaseOrder.srf"
                            Case "mx_forklif"
                                RPOsrfname = "ForkliftPurchaseOrder.srf"
                            Case "mx_Out"
                                RPOsrfname = "OutPurchaseOrder.srf"
                            Case "mx_COO"
                                RPOsrfname = "CertificateOfOriginPPurchaseOrder.srf"
                        End Select

                        If RPOsrfname = "CranePurchaseOrder_Form.srf" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                        ElseIf RPOsrfname = "ForkliftPurchaseOrder.srf" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                        ElseIf RPOsrfname = "CourierPurchaseOrder.srf" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                        ElseIf RPOsrfname = "DGLPPurchaseOrder.srf" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                        ElseIf RPOsrfname = "OutPurchaseOrder.srf" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                        ElseIf RPOsrfname = "CertificateOfOriginPPurchaseOrder.srf" Then
                            ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                        Else
                            ' ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000006", 1)
                        End If

                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000025" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000026" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000032" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000035" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000038" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000028" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000031" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000034" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000037" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000040" Then           'Purchase
                            If currentRow > 0 Then
                                oMatrix = ExportSeaLCLForm.Items.Item(RPOmatrixname).Specific
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, RPOsrfname)
                                    CPOForm = p_oSBOApplication.Forms.ActiveForm
                                    CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                    CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                                    CPOForm.Items.Item("1").Click()
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                    p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                End If
                                
                                ' p_oSBOApplication.ActivateMenuItem("1291")
                            Else
                                p_oSBOApplication.MessageBox("Need to select the Row that you want to Edit")
                            End If
                        Else
                            If currentRow > 0 Then
                                If AlreadyExist("EXPORTSEALCL") Then
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                                ElseIf AlreadyExist("EXPORTAIRLCL") Then
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                                ElseIf AlreadyExist("EXPORTLAND") Then
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                                End If
                                oMatrix = ExportSeaLCLForm.Items.Item(RPOmatrixname).Specific
                                If RPOmatrixname = "mx_TkrList" Then
                                    If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                        LoadTruckingPO(ExportSeaLCLForm, RPOsrfname)
                                        ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        oMatrix = ExportSeaLCLForm.Items.Item(RPOmatrixname).Specific
                                        CPOForm = p_oSBOApplication.Forms.ActiveForm
                                        CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        CPOForm.Items.Item("ed_CPOID").Specific.Value = oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString
                                        CPOForm.Items.Item("1").Click()
                                        CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                        p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                    ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                        p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                    End If
                                    'MSW 14-09-2011 Truck PO
                                Else
                                    If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                        LoadAndCreateCPO(ExportSeaLCLForm, RPOsrfname)
                                        ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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

                    End If

                Case "CopyToCGR"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            Select Case ActiveMatrix
                                Case "mx_Crane"
                                    RGRsrfname = "CraneGoodsReceipt_Form.srf"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                                Case "mx_forklif"
                                    RGRsrfname = "ForkliftGoodsReceipt.srf"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                Case "mx_Courier"
                                    RGRsrfname = "CourierGoodsReceipt.srf"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                Case "mx_DGLP"
                                    RGRsrfname = "DGLPGoodsReceipt.srf"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                Case "mx_Out"
                                    RGRsrfname = "OutriderGoodReceipt.srf"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                                Case "mx_COO"
                                    RGRsrfname = "CertificateOfOriginGoodsReceipt.srf"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                            End Select

                            If ActiveMatrix = "mx_Fumi" Or ActiveMatrix = "mx_Crate" Or ActiveMatrix = "mx_Armed" Or ActiveMatrix = "mx_Bunk" Or ActiveMatrix = "mx_Bok" Or ActiveMatrix = "mx_TkrList" Then
                                If currentRow > 0 Then
                                    If AlreadyExist("EXPORTSEALCL") Then
                                        ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                                    ElseIf AlreadyExist("EXPORTAIRLCL") Then
                                        ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                                    ElseIf AlreadyExist("EXPORTLAND") Then
                                        ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                                    End If
                                    oMatrix = ExportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                    'Truck PO
                                    If RPOmatrixname = "mx_TkrList" Then
                                        If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                            LoadAndCreateCGR(ExportSeaLCLForm, RGRsrfname)
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-201
                                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                                            If Not FillDataToGoodsReceipt(ExportSeaLCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                            p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                        End If
                                        'Truck PO
                                    Else
                                        If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                            LoadAndCreateCGR(ExportSeaLCLForm, RGRsrfname)
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                                            If Not FillDataToGoodsReceipt(ExportSeaLCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
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
                                    oMatrix = ExportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                    If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                        LoadAndCreate_GoodReceipt(ExportSeaLCLForm, RGRsrfname)
                                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000027" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000030" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000033" Or _
                                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000036" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000039" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000042" Then

                                            CGRForm = p_oSBOApplication.Forms.ActiveForm
                                            If Not FillDataToGoodsReceipt(ExportSeaLCLForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
                                        End If
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
                            ExportSeaLCLForm.MessageBox("Need to select the Row that you want to copy to Goods Receipt")
                        End If
                    End If

                Case "CancelPO"
                    If pVal.BeforeAction = False Then

                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000025" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000026" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000029" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000032" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000035" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000038" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000028" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000031" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000034" Or _
                            p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000037" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000040" Then           'Purchase
                            If currentRow > 0 Then

                                If ActiveMatrix = "mx_Crane" Then
                                    oMatrixName = "mx_Crane"
                                    TableName = "@OBT_TB33_CRANE"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                                ElseIf ActiveMatrix = "mx_forklif" Then
                                    oMatrixName = "mx_forklif"
                                    TableName = "@OBT_TB34_FORKLIFT"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                ElseIf ActiveMatrix = "mx_Courier" Then
                                    oMatrixName = "mx_Courier"
                                    TableName = "@OBT_TB35_COURIER"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                ElseIf ActiveMatrix = "mx_DGLP" Then
                                    oMatrixName = "mx_DGLP"
                                    TableName = "@OBT_TB37_DGLP"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                ElseIf ActiveMatrix = "mx_Out" Then
                                    oMatrixName = "mx_Out"
                                    TableName = "@OBT_TB38_OUTRIDER"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                                ElseIf ActiveMatrix = "mx_COO" Then
                                    oMatrixName = "mx_COO"
                                    TableName = "@OBT_TB39_COO"
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                                End If

                                oMatrix = ExportSeaLCLForm.Items.Item(oMatrixName).Specific
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    If oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value <> "" Then
                                        If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                        If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                        Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                          "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                        If Not PopulateOtherPurchaseHeader(ExportSeaLCLForm, oMatrix, sql, TableName) Then Throw New ArgumentException(sErrDesc)
                                    End If


                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oMatrix = ExportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                                    ExportSeaLCLForm.Items.Item("1").Click()
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
                                If AlreadyExist("EXPORTSEALCL") Then
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
                                ElseIf AlreadyExist("EXPORTAIRLCL") Then
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
                                ElseIf AlreadyExist("EXPORTLAND") Then
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
                                End If
                                oMatrix = ExportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                If ActiveMatrix = "mx_Fumi" Or ActiveMatrix = "mx_Crate" Or ActiveMatrix = "mx_Armed" Or ActiveMatrix = "mx_Bunk" Or ActiveMatrix = "mx_Bok" Or ActiveMatrix = "mx_TkrList" Then
                                    If ActiveMatrix = "mx_TkrList" Then
                                        If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then

                                            If oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value <> "" Then
                                                If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                                If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                                oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                                        "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                                        "from  [@OBT_TB08_FFCPO] a inner join [@OBT_TB006_ETRUCKING] b on a.DocEntry=b.U_PODocNo where b.U_PODocNo = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                                If Not PopulateTruckPurchaseHeader(ExportSeaLCLForm, oMatrix, sql, "@OBT_TB006_ETRUCKING") Then Throw New ArgumentException(sErrDesc)
                                            End If

                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            oMatrix = ExportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                            ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                                            ExportSeaLCLForm.Items.Item("1").Click()

                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "" Then
                                            p_oSBOApplication.MessageBox("There is no Purchase Order for Internal.")
                                        End If

                                    Else
                                        'Binding To Matrix when updating PO Status to Cancel is finished
                                        If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                            Dim tblName As String = String.Empty
                                            If oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value <> "" Then
                                                If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString())) Then Throw New ArgumentException(sErrDesc)
                                                If Not UpdateForCancelStatus(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value.ToString()) Then Throw New ArgumentException(sErrDesc) 'Update POStatus to Cancel 
                                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                  "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value.ToString())
                                                If ActiveMatrix = "mx_Bok" Then
                                                    tblName = "@OBT_TB19_BOOKING"
                                                ElseIf ActiveMatrix = "mx_Crate" Then
                                                    tblName = "@OBT_TB020_CRATE"
                                                ElseIf ActiveMatrix = "mx_Fumi" Then
                                                    tblName = "@OBT_TB021_FUMIGAT"
                                                ElseIf ActiveMatrix = "mx_Bunk" Then
                                                    tblName = "@OBT_TB022_BUNKER"
                                                ElseIf ActiveMatrix = "mx_Armed" Then
                                                    tblName = "@OBT_TB023_ARMESCORT"
                                                End If
                                                If Not PopulateOtherPurchaseHeaderFromMain(ExportSeaLCLForm, oMatrix, sql, tblName) Then Throw New ArgumentException(sErrDesc)
                                            End If

                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            oMatrix = ExportSeaLCLForm.Items.Item(ActiveMatrix).Specific
                                            ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                                            ExportSeaLCLForm.Items.Item("1").Click()

                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Closed" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Closed.")
                                        ElseIf oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Cancelled" Then
                                            p_oSBOApplication.MessageBox("Purchase Order Status is already Cancelled.")

                                        End If
                                    End If
                                End If
                            Else
                                p_oSBOApplication.MessageBox("Need to select the Row that you want to Cancel")
                            End If
                        End If

                    End If


            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoExportSeaLCLMenuEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaLCLMenuEvent = RTN_ERROR
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
    Private Sub AddGSTComboData(ByVal oColumn As SAPbouiCOM.Column)

        ' **********************************************************************************
        '   Function    :    AddGSTComboData()
        '   Purpose     :    This function will be providing to add GST Amount  for
        '                    Purechase order item value for Purchase Order process                   
        '   Parameters  :    ByVal oColumn As SAPbouiCOM.Column                       
        '   Return      :    No
        '                   
        ' **********************************************************************************
        'Dim oForm As SAPbouiCOM.Form
        'Dim oForm As SAPbouiCOM.Form
        Dim RS As SAPbobsCOM.Recordset
        'oForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL") 'MSW
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
            'NOGST = Convert.ToDouble(oMatrix.Columns.Item("colAmount1").Cells.Item(Row).Specific.Value) 'MSW TO Edit New Ticket
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
        If AlreadyExist("EXPORTSEALCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
        ElseIf AlreadyExist("EXPORTAIRLCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
        ElseIf AlreadyExist("EXPORTLAND") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
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
        '                   of ExportSealCL Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form        '                         
        '   Return      :   No
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
            'If HolidayMarkUp(pForm, pForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, "") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            pForm.Freeze(False)

        Catch ex As Exception

        End Try
    End Sub
    Private Sub GetVoucherDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)

        ' **********************************************************************************
        '   Function    :   GetVoucherDataFromMatrixByIndex()
        '   Purpose     :   This function will be providing to proceed  for to get data from maxtrix 
        '                   of ExportSealCL Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix
        '                   ByVal Index As Integer                     
        '   Return      :   No
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

    Public Function DoExportSeaLCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
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
        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
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
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim BoolResize As Boolean = False
        Dim SqlQuery As String = String.Empty
        Dim FunctionName As String = "DoImportSeaLCLItemEvent()"
        Dim sql As String = ""
        Dim formUid As String = String.Empty
        Dim strDsp As String = String.Empty

        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            If AlreadyExist("EXPORTSEALCL") Then
                formUid = "EXPORTSEALCL"
            ElseIf AlreadyExist("EXPORTSEAFCL") Then
                formUid = "EXPORTSEAFCL"
            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                formUid = "EXPORTAIRLCL"
            ElseIf AlreadyExist("EXPORTLAND") Then
                formUid = "EXPORTLAND"
            End If

            Select Case pVal.FormTypeEx
                Case "SendAlert"
                    Dim mailFrom As String = String.Empty
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                    oActiveForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "bt_Close" Then
                                'oActiveForm.Close()
                                PDFAlert(ExportSeaLCLForm)
                            End If
                            If pVal.ItemUID = "bt_OK" Then
                                PDFAlert(ExportSeaLCLForm)
                                oRecordSet.DoQuery("select E_Mail from OUSR where USER_CODE='" & p_oDICompany.UserName & "'")
                                If oRecordSet.RecordCount > 0 Then
                                    mailFrom = oRecordSet.Fields.Item("E_Mail").Value
                                End If
                                If Not reportuti.SendMailDoc(oActiveForm.Items.Item("ed_To").Specific.Value, mailFrom, "Pre Alert", "Testing Mail", pdffilepath) Then
                                    p_oSBOApplication.MessageBox("Send Message Fail", 0, "OK")
                                Else
                                    p_oSBOApplication.MessageBox("Send Message Successfully", 0, "OK")
                                End If
                            End If

                        End If


                    End If
                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000044", _
                    "2000000027", "2000000029", "2000000033", "2000000036", "2000000039", "2000000042", "2000000051"  'MSW 14-09-2011 Truck PO  ' CGR --> Custom Goods Receipt
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

                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
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
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                                                ElseIf pVal.FormTypeEx = "2000000029" Then 'Forklift
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                                ElseIf pVal.FormTypeEx = "2000000033" Then
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                                ElseIf pVal.FormTypeEx = "2000000036" Then
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                                ElseIf pVal.FormTypeEx = "2000000039" Then
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                                                ElseIf pVal.FormTypeEx = "2000000042" Then
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                                                Else
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                                End If
                                                If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                                End If
                                            Catch ex As Exception

                                            End Try
                                            'If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            '    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            'End If
                                            ExportSeaLCLForm.Items.Item("1").Click()

                                        End If
                                    End If
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_DrBL").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_SpOrder").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_DGD").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                                    ExportSeaLCLForm.Items.Item("bt_Crane").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Foklit").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Courier").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_DGLP").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Outer").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Alert").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Certifi").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Excel").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Label").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_A6Label").Enabled = True
                                End If
                            End If
                        End If

                    End If
                Case "2000000026", "2000000030", "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000029", "2000000032", "2000000035", "2000000038", "2000000041", "2000000043", "2000000050"  ''MSW To Edit New Ticket    ' CPO --> Custom Purchase Order"    ' CPO --> Custom Purchase Order"
                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
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
                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            'when retrieve spefific data for update, add new row in the matrix
                                            If Not AddNewRowPO(CPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                                            If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            End If
                                            CPOForm.Items.Item("ed_Code").Enabled = False
                                            CPOForm.Items.Item("ed_Name").Enabled = False
                                            CPOForm.Items.Item("cb_SInA").Enabled = True
                                            CPOForm.Items.Item("bt_Preview").Visible = True
                                            CPOForm.Items.Item("bt_Resend").Visible = False
                                            CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            CPOForm.Close()
                                            Try
                                                If pVal.FormTypeEx = "2000000026" Then 'Crane
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000025", 1)
                                                ElseIf pVal.FormTypeEx = "2000000030" Then 'Forklift
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000028", 1)
                                                ElseIf pVal.FormTypeEx = "2000000035" Then 'Forklift
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000034", 1)
                                                ElseIf pVal.FormTypeEx = "2000000038" Then 'Forklift
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000037", 1)
                                                ElseIf pVal.FormTypeEx = "2000000041" Then 'Forklift
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000040", 1)
                                                ElseIf pVal.FormTypeEx = "2000000032" Then 'Forklift
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm("2000000031", 1)
                                                Else
                                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                                End If
                                                If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                                End If
                                            Catch ex As Exception

                                            End Try
                                            ExportSeaLCLForm.Items.Item("1").Click()
                                            'MSW 14-09-2011 Truck PO
                                            If pVal.FormTypeEx = "2000000050" Then
                                                EnabledTruckerForExternal(ExportSeaLCLForm, False)
                                                ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                                            End If
                                            'MSW 14-09-2011 Truck PO
                                        End If
                                    End If
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_DrBL").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_SpOrder").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_DGD").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                                    ExportSeaLCLForm.Items.Item("bt_Crane").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Foklit").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Courier").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_DGLP").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Outer").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Alert").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Certifi").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Excel").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_Label").Enabled = True
                                    ExportSeaLCLForm.Items.Item("bt_A6Label").Enabled = True
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If
                            If pVal.ItemUID = "bt_Preview" Then
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                PreviewPO(ExportSeaLCLForm, CPOForm)
                            End If
                        End If
                    End If

                    If pVal.BeforeAction = True Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "1" Then
                                Try
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
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
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
                                    oShpForm.Items.Item("bt_PPView").Visible = True
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Then
                                        If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB03_EXPSHPINV] set " & _
                                            " U_FrDocNo=" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oShpForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            oShpForm.Close()

                                            ExportSeaLCLForm.Items.Item("1").Click()
                                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
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
                                        noofPkg = qty / Convert.ToDouble(oShpForm.Items.Item("ed_PPBNo").Specific.value.ToString)
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


                    End If
                Case "VOUCHER"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
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

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                        If pVal.ItemUID = "1" Then
                            Try
                                ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)
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
                                oPayForm.Items.Item("ed_VedCode").Enabled = False
                                oPayForm.Items.Item("bt_PayView").Visible = True
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
                                PreviewPaymentVoucher(ExportSeaLCLForm, oPayForm)
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
                                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(formUid, 1)

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
                                            'oPayForm.Items.Item("ed_PayTo").Enabled = False

                                        End If
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB031_VHEADER] set U_APInvNo=" & Convert.ToInt32(strAPInvNo) & ",U_OutPayNo=" & Convert.ToInt32(strOutPayNo) & "" & _
                                            " ,U_FrDocNo=" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oPayForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            PreviewPaymentVoucher(ExportSeaLCLForm, oPayForm)
                                            oPayForm.Close()

                                            ExportSeaLCLForm.Items.Item("1").Click()
                                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                End If
                            End If
                        End If

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
                            End If

                            If pVal.ColUID = "colChCode1" Then
                                oChMatrix = oPayForm.Items.Item("mx_ChCode").Specific
                                'MSW to Edit New Ticket
                                Try
                                    oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    ' oChMatrix.Columns.Item("colChCode1").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("U_CCode").Cells.Item(0).Value.ToString
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


                Case "EXPORTSEALCL", "EXPORTSEAFCL", "EXPORTAIRLCL", "EXPORTLAND"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(ExportSeaLCLForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                            LoadHolidayMarkUp(ExportSeaLCLForm)
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_VALIDATE And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                            If pVal.ItemUID = "ed_JobNo" Then
                                ValidateJobNumber(ExportSeaLCLForm, BubbleEvent)
                            End If
                        End If
                    End If
                    'MSW to edit New Ticket 07-09-2011
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                        If pVal.ItemUID = "ed_TotalM3" Or pVal.ItemUID = "ed_TotalWt" Then
                            If Convert.ToDouble(ExportSeaLCLForm.Items.Item("ed_TotalM3").Specific.Value) > Convert.ToDouble(ExportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Value) Then
                                ExportSeaLCLForm.Items.Item("ed_TotChWt").Specific.Value = ExportSeaLCLForm.Items.Item("ed_TotalM3").Specific.Value
                            Else
                                ExportSeaLCLForm.Items.Item("ed_TotChWt").Specific.Value = ExportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Value
                            End If
                            ' ExportSeaLCLForm.Items.Item("ed_TotalWt").Specific.Active = True
                        End If
                    End If
                    'End MSW to edit New Ticket 07-09-2011
                    If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_DrBL").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_SpOrder").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_DGD").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                        ExportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                    End If


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
                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                        If pVal.ItemUID = "ed_FCharge" And pVal.Before_Action = True Then
                            If ExportSeaLCLForm.Items.Item("ed_FCharge").Specific.value <> "" Then
                                Try
                                    If ExportSeaLCLForm.Items.Item("ed_FCharge").Specific.value < 0 Then
                                        ExportSeaLCLForm.Items.Item("ed_FCharge").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_FCharge").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ExportSeaLCLForm.Items.Item("ed_FCharge").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_FCharge").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End Try
                            End If

                        End If
                        If pVal.ItemUID = "ed_Percent" And pVal.Before_Action = True Then
                            If ExportSeaLCLForm.Items.Item("ed_Percent").Specific.value <> "" Then
                                Try
                                    If ExportSeaLCLForm.Items.Item("ed_Percent").Specific.value < 0 Then
                                        ExportSeaLCLForm.Items.Item("ed_Percent").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_Percent").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ExportSeaLCLForm.Items.Item("ed_Percent").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Percent").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End Try
                            End If

                        End If
                        If pVal.ItemUID = "ed_GWt" And pVal.Before_Action = True Then
                            If ExportSeaLCLForm.Items.Item("ed_GWt").Specific.value <> "" Then
                                Try
                                    If ExportSeaLCLForm.Items.Item("ed_GWt").Specific.value < 0 Then
                                        ExportSeaLCLForm.Items.Item("ed_GWt").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_GWt").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                Catch ex As Exception
                                    ExportSeaLCLForm.Items.Item("ed_GWt").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_GWt").Specific.Active = True
                                    p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False

                                End Try
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
                                        If ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                                            ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                        ElseIf ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                                            ExportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                                        Else
                                            ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                        End If
                                        'ExportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
                                        'ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
                                    Case "fo_Trkng"
                                        ExportSeaLCLForm.PaneLevel = 6
                                        ExportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_TkrList"
                                        '#1008 17-09-2011
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
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_Voucher"
                                        '#1008 17-09-2011
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
                                        ExportSeaLCLForm.Freeze(True)
                                        ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                        If ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction" Then

                                            oRecordSet.DoQuery("SELECT Address FROM OCRD WHERE CardCode = '" & ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value.ToString & "'")
                                            If oRecordSet.RecordCount > 0 Then
                                                ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oRecordSet.Fields.Item("Address").Value.ToString
                                            End If
                                            ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRDATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                                            'ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            EnabledTruckerForExternal(ExportSeaLCLForm, True)   'MSW 14-09-2011 Truck PO
                                            ExportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True

                                        End If

                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                        If ExportSeaLCLForm.Items.Item("ed_InsDoc").Specific.Value = "" Then
                                            ExportSeaLCLForm.Items.Item("ed_TkrTime").Specific.Value = Now.ToString("HH:mm")
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
                                        ExportSeaLCLForm.Freeze(False)

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
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_ShpInv"
                                        '#1008 17-09-2011
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_ShpInv").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                                End If

                                            Else
                                                '    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                    Case "fo_ShView"
                                        ExportSeaLCLForm.PaneLevel = 25
                                    Case "fo_BkVsl"
                                        ExportSeaLCLForm.PaneLevel = 26
                                        ExportSeaLCLForm.Items.Item("fo_BView").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_Bok"
                                        '#1008 17-09-2011
                                        oMatrix = ExportSeaLCLForm.Items.Item("mx_Bok").Specific
                                        If oMatrix.RowCount > 1 Then
                                            ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                                        ElseIf oMatrix.RowCount = 1 Then
                                            If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                                If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                    ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                                                End If

                                            Else
                                                'ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                            End If
                                        End If
                                    Case "fo_Crate"
                                        ExportSeaLCLForm.PaneLevel = 27
                                        ExportSeaLCLForm.Items.Item("fo_Cr").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_Crate"
                                        '#1008 17-09-2011
                                    Case "fo_Fumi"
                                        ExportSeaLCLForm.PaneLevel = 28
                                        ExportSeaLCLForm.Items.Item("fo_FV").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_Fumi"
                                        '#1008 17-09-2011
                                    Case "fo_OpBunk"
                                        ExportSeaLCLForm.PaneLevel = 29
                                        ExportSeaLCLForm.Items.Item("fo_Buk").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_Bunk"
                                        '#1008 17-09-2011
                                    Case "fo_ArmEs"
                                        ExportSeaLCLForm.PaneLevel = 30
                                        ExportSeaLCLForm.Items.Item("fo_Arm").Specific.Select()
                                        '#1008 17-09-2011
                                        ExportSeaLCLForm.Settings.Enabled = True
                                        ExportSeaLCLForm.Settings.EnableRowFormat = True
                                        ExportSeaLCLForm.Settings.MatrixUID = "mx_Armed"
                                        '#1008 17-09-2011
                                End Select

                                If pVal.FormTypeEx = "EXPORTSEALCL" Or pVal.FormTypeEx = "EXPORTSEAFCL" Then
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
                                End If



                                'testing

                                If pVal.ItemUID = "bt_Crane" Then
                                    LoadButtonForm(ExportSeaLCLForm, "Crane.srf", "Crane")
                                    oRecordSet.DoQuery("Select * from [@CRANE] Where U_DocNum='" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        CraneForm = p_oSBOApplication.Forms.ActiveForm
                                        CraneForm.EnableMenu("771", False)
                                        CraneForm.EnableMenu("774", False)
                                        CraneForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        CraneForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        CraneForm.Items.Item("1").Click()
                                        CraneForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        CraneForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Foklit" Then
                                    LoadButtonForm(ExportSeaLCLForm, "Forklift.srf", "ForkLift")
                                    oRecordSet.DoQuery("Select * from  [@FORKLIFT] Where U_DocNum='" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        ForkForm = p_oSBOApplication.Forms.ActiveForm
                                        ForkForm.EnableMenu("771", False)
                                        ForkForm.EnableMenu("774", False)
                                        ForkForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        ForkForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        ForkForm.Items.Item("1").Click()
                                        ForkForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        ForkForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Courier" Then
                                    LoadButtonForm(ExportSeaLCLForm, "Courier.srf", "Courier")
                                    oRecordSet.DoQuery("Select * from  [@COURIER] Where U_DocNum='" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        CourForm = p_oSBOApplication.Forms.ActiveForm
                                        CourForm.EnableMenu("771", False)
                                        CourForm.EnableMenu("774", False)
                                        CourForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        CourForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        CourForm.Items.Item("1").Click()
                                        CourForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        CourForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If

                                If pVal.ItemUID = "bt_DGLP" Then
                                    LoadButtonForm(ExportSeaLCLForm, "DGLP.srf", "DGLP")
                                    oRecordSet.DoQuery("Select * from  [@DGLP] Where U_DocNum='" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        DGLPForm = p_oSBOApplication.Forms.ActiveForm
                                        DGLPForm.EnableMenu("771", False)
                                        DGLPForm.EnableMenu("774", False)
                                        DGLPForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        DGLPForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        DGLPForm.Items.Item("1").Click()
                                        DGLPForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        DGLPForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Outer" Then
                                    LoadButtonForm(ExportSeaLCLForm, "Outrider.srf", "Outrider")
                                    oRecordSet.DoQuery("Select * from  [@OUTRIDER] Where U_DocNum='" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        OutForm = p_oSBOApplication.Forms.ActiveForm
                                        OutForm.EnableMenu("771", False)
                                        OutForm.EnableMenu("774", False)
                                        OutForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        OutForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        OutForm.Items.Item("1").Click()
                                        OutForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        OutForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Certifi" Then
                                    LoadButtonForm(ExportSeaLCLForm, "CertificateOfOrigin.srf", "COO")
                                    oRecordSet.DoQuery("Select * from  [@CERTIFICATEOFORIGIN] Where U_DocNum='" & ExportSeaLCLForm.Items.Item("ed_DocNum").Specific.Value & "'")
                                    If oRecordSet.RecordCount > 0 Then
                                        COOForm = p_oSBOApplication.Forms.ActiveForm
                                        COOForm.EnableMenu("771", False)
                                        COOForm.EnableMenu("774", False)
                                        COOForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                        COOForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("U_DocNum").Value
                                        COOForm.Items.Item("1").Click()
                                        COOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        COOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                End If
                                If pVal.ItemUID = "bt_Alert" Then
                                    'LoadFromXML(p_oSBOApplication, "PreAlertTo.srf")
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
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB014_EPLICINFO").RemoveRecord(0)   '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
                                            Dim oUserTable As SAPbobsCOM.UserTable
                                            oUserTable = p_oDICompany.UserTables.Item("OBT_TB014_EPLICINFO")                                '* Change Nyan Lin   "[@OBT_TB014_PLICINFO]"
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
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB012_EPCONT").RemoveRecord(0)    '* Change Nyan Lin   "[@OBT_TB012_PCONTAINE]"
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
                                'MSW to edit 10-09-2011
                                If pVal.ItemUID = "bt_PrntDis" Then
                                    PreviewDispatchInstruction(ExportSeaLCLForm)
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

                                If pVal.ItemUID = "ed_ETDDate" And ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific.String <> String.Empty Then
                                    'If DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific, ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific, ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    Try
                                        If Not DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific, ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific, ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific.Value
                                        ExportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific.Value
                                        ExportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific.Value
                                        'ExportSeaLCLForm.ActiveItem = "ed_CrgDsc"
                                    Catch ex As Exception
                                        'ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific.Value
                                        'ExportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific.Value
                                        'ExportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific.Value
                                        'ExportSeaLCLForm.ActiveItem = "ed_CrgDsc"
                                    End Try
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific, ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific.Value
                                    If Not DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ADate").Specific, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "ed_ETADate" Then
                                    If Not DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "ed_JbDate" And ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspDate" And ExportSeaLCLForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                                    If Not DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                                If pVal.ItemUID = "ed_DspCDte" And ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                                    If Not DateTime(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If

                                If pVal.ItemUID = "bt_Payment" Then
                                    p_oSBOApplication.ActivateMenuItem("2818")
                                End If

                                If pVal.ItemUID = "op_Inter" Then
                                    ExportSeaLCLForm.Freeze(True)
                                    'MSW 14-09-2011 Truck PO
                                    If ExportSeaLCLForm.Items.Item("ed_PONo").Specific.Value <> "" Then
                                        If Not CancelTruckingPurchaseOrder(ExportSeaLCLForm) Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    'MSW 14-09-2011 Truck PO
                                    ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
                                    ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                    'MSW To Edit New Ticket
                                    ExportSeaLCLForm.Items.Item("ed_VehicNo").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_Attent").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = ""
                                    ExportSeaLCLForm.Items.Item("ed_EUC").Specific.Value = ""
                                    EnabledTruckerForExternal(ExportSeaLCLForm, True)
                                    'End MSW To Edit New Ticket
                                    If AddChooseFromListByOption(ExportSeaLCLForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Active = True
                                    ExportSeaLCLForm.Freeze(False)
                                ElseIf pVal.ItemUID = "op_Exter" Then
                                    ExportSeaLCLForm.Freeze(True)
                                    If ExportSeaLCLForm.Items.Item("ed_PONo").Specific.Value = "" Then 'MSW 14-09-2011 Truck PO
                                        ExportSeaLCLForm.Items.Item("ed_Trucker").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_Fax").Specific.Value = ""
                                        ExportSeaLCLForm.Items.Item("ed_Email").Specific.Value = ""
                                        'MSW To Edit New Ticket
                                        ExportSeaLCLForm.Items.Item("ed_EUC").Specific.Value = ""
                                        EnabledTruckerForExternal(ExportSeaLCLForm, False)
                                        'End MSW To Edit New Ticket
                                        ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = True
                                        If AddChooseFromListByOption(ExportSeaLCLForm, False, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                        'ExportSeaLCLForm.Items.Item("ed_PONo").Specific.Active = True
                                    End If
                                    ExportSeaLCLForm.Freeze(False)
                                End If

                                If pVal.ItemUID = "op_DspIntr" Then
                                    If ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                                        strDsp = ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific.Value
                                    End If
                                    ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal"
                                    oCombo = ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If Not ClearComboData(ExportSeaLCLForm, "cb_Dspchr", "@OBT_TB007_EDISPATCH", "U_Dispatch") Then Throw New ArgumentException(sErrDesc)
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
                                    If ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                                        strDsp = ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific.Value
                                    End If
                                    ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External"
                                    oCombo = ExportSeaLCLForm.Items.Item("cb_Dspchr").Specific
                                    If Not ClearComboData(ExportSeaLCLForm, "cb_Dspchr", "@OBT_TB007_EDISPATCH", "U_Dispatch") Then Throw New ArgumentException(sErrDesc)
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

                                    'Purchase Order and Goods Receipt POP UP 

                                    ' ==================================== Creating Custom Purchase Order ==============================
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close"
                                        End If
                                        LoadTruckingPO(ExportSeaLCLForm, "TkrListPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                    ' ==================================== Creating Custom Purchase Order ==============================

                                    'p_oSBOApplication.Menus.Item("2305").Activate()
                                    'p_oSBOApplication.ActivateMenuItem("6913") 'MSW 04-04-2011
                                    'p_oSBOApplication.Menus.Item("6913").Activate()
                                    ''SBO_Application.Menus.Item("6913").Activate()
                                    'Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                                    'UDFAttachForm.Items.Item("U_JobNo").Specific.Value = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    'UDFAttachForm.Items.Item("U_InsDate").Specific.Value = ExportSeaLCLForm.Items.Item("ed_InsDate").Specific.Value
                                End If

                                If pVal.ItemUID = "1" Then
                                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If pVal.ActionSuccess = True Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            ExportSeaLCLForm.Items.Item("ed_Code").Enabled = False 'MSW
                                            ExportSeaLCLForm.Items.Item("ed_Name").Enabled = False
                                            ExportSeaLCLForm.Items.Item("ed_JobNo").Enabled = False
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_DrBL").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_SpOrder").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_DGD").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                                            ExportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                    End If
                                End If
                                'KM to edit
                                If pVal.ItemUID = "bt_PrntIns" Then
                                    PreviewInsDoc(ExportSeaLCLForm)
                                End If
                                If pVal.ItemUID = "bt_A6Label" Then
                                    PreviewA6Label(ExportSeaLCLForm)
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
                                            ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False 'MSW to edit New Ticket 07-09-2011
                                        End If
                                        ClearText(ExportSeaLCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                                        ExportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                        ExportSeaLCLForm.Items.Item("1").Click() 'MSW to edit New Ticket 07-09-2011
                                        ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                                    End If
                                End If

                                If pVal.ItemUID = "bt_DelIns" Then
                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    modTrucking.DeleteByIndex(ExportSeaLCLForm, oMatrix, "@OBT_TB006_ETRUCKING")                             '* Change Nyan Lin   "[@OBT_TB006_TRUCKING]"
                                    ClearText(ExportSeaLCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                                    ExportSeaLCLForm.Items.Item("fo_TkrView").Specific.Select()
                                End If

                                If pVal.ItemUID = "bt_AmdIns" Then

                                    oMatrix = ExportSeaLCLForm.Items.Item("mx_TkrList").Specific
                                    'modTrucking.SetDataToEditTabByIndex(ImportSeaLCLForm)

                                    If (oMatrix.GetNextSelectedRow < 0) Then
                                        p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction", 1, "OK")
                                        'Exit Function
                                    Else
                                        ExportSeaLCLForm.Items.Item("bt_TkrAdd").Specific.Caption = "Update"
                                        modTrucking.GetDataFromMatrixByIndex(ExportSeaLCLForm, oMatrix, oMatrix.GetNextSelectedRow)
                                        modTrucking.SetDataToEditTabByIndex(ExportSeaLCLForm)
                                        'MSW 14-09-2011 Truck PO
                                        If ExportSeaLCLForm.Items.Item("op_Exter").Specific.Selected = True Then
                                            EnabledTruckerForExternal(ExportSeaLCLForm, False)
                                        ElseIf ExportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                            EnabledTruckerForExternal(ExportSeaLCLForm, True)
                                        End If
                                        'MSW 14-09-2011 Truck PO
                                        ExportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                        ExportSeaLCLForm.Items.Item("bt_Tkview").Enabled = True 'MSW to edit New Ticket 07-09-2011
                                        ExportSeaLCLForm.Items.Item("fo_TkrEdit").Specific.Select()
                                    End If

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
                                    ClearText(ExportSeaLCLForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                                    ExportSeaLCLForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                                    ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                                    ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False 'MSW to edit New Ticket 07-09-2011
                                End If

                                '-------------------------Payment Vouncher (OMM)------------------------------------------------------'
                                If pVal.ItemUID = "fo_VoEdit" Then
                                    ExportSeaLCLForm.Freeze(True)
                                    ExportSeaLCLForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    ExportSeaLCLForm.Items.Item("op_Cash").Specific.Selected = True
                                    ExportSeaLCLForm.Items.Item("ed_PJobNo").Specific.Value = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    ExportSeaLCLForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()
                                    'If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
                                End If

                                If pVal.ItemUID = "bt_AmdVoc" Then
                                    'POP UP Payment Voucher
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadPaymentVoucher(ExportSeaLCLForm)
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Voucher.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        'Exit Function
                                    End If

                                End If
                                'Shipping Invoice POP Up
                                If pVal.ItemUID = "bt_ShpInv" Then
                                    'POP UP Shipping Invoice
                                    'If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If formUid = "EXPORTSEALCL" Or formUid = "EXPORTSEAFCL" Then
                                        If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And _
                                        Not ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = "" And _
                                        Not ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value = "" And Not ExportSeaLCLForm.Items.Item("cb_PCode").Specific.Value = "" And _
                                        Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            LoadShippingInvoice(ExportSeaLCLForm)
                                        Else
                                            If ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Customer Code to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                                'ElseIf ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" Then
                                                '    p_oSBOApplication.SetStatusBarMessage("No Shipping Agent Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaLCLForm.Items.Item("cb_PCode").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Port of Loading to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.Value = "" And formUid <> "EXPORTLAND" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Vessel/Voy Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            End If
                                            'Exit Function
                                        End If
                                    ElseIf formUid = "EXPORTAIRLCL" Or formUid = "EXPORTLAND" Or formUid = "EXPORTLAND" Then
                                        If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And _
                                        Not ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value = "" And _
                                         Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            LoadShippingInvoice(ExportSeaLCLForm)
                                        Else
                                            If ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            ElseIf ExportSeaLCLForm.Items.Item("ed_Code").Specific.Value = "" Then
                                                p_oSBOApplication.SetStatusBarMessage("No Customer Code to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                                'ElseIf ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.Value = "" Then
                                                '    p_oSBOApplication.SetStatusBarMessage("No Shipping Agent Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            End If
                                            'Exit Function
                                        End If
                                    End If


                                End If

                                'Purchase Order and Goods Receipt POP UP

                                '---------------------------------- Booking Preview-----------------------------'
                                If pVal.ItemUID = "bt_DrBL" Then
                                    'ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                    'DeletePDF()
                                    'CopyFolder("DraftBL")
                                    pdfFilename = "DraftBL"
                                    originpdf = "DraftBL.pdf"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    PutDBValueToDBAndPreviewForDraftBL("DraftBL", ExportSeaLCLForm, pdffilepath, originpdf)
                                End If
                                If pVal.ItemUID = "bt_DGD" Then
                                    'ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                    'DeletePDF()
                                    'CopyFolder("DGD")
                                    pdfFilename = "DGD"
                                    originpdf = "DGD-Air.pdf"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    PutDBValueToDBAndPreviewForDGD("DGD", ExportSeaLCLForm, pdffilepath, originpdf)
                                End If
                                If pVal.ItemUID = "bt_SpOrder" Then
                                    'ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                                    'DeletePDF()
                                    'CopyFolder("ShippingOrder")
                                    pdfFilename = "ShippingOrder"
                                    originpdf = "ShippingOrder.pdf"
                                    mainFolder = p_fmsSetting.DocuPath
                                    jobNo = ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value
                                    pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
                                    PutDBValueToDBAndPreviewForShippingOrder("ShippingOrder", ExportSeaLCLForm, pdffilepath, originpdf)
                                End If

                                If pVal.ItemUID = "bt_BokPO" Then 'sw
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaLCLForm, "BokPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        'Exit Function
                                    End If
                                End If

                                If pVal.ItemUID = "bt_CPO" Then
                                    ' ==================================== Creating Custom Purchase Order ==============================
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaLCLForm, "FumiPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If
                                    ' ==================================== Creating Custom Purchase Order ==============================
                                End If

                                If pVal.ItemUID = "bt_CrPO" Then 'sw
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaLCLForm, "CratePurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If
                                End If

                                If pVal.ItemUID = "bt_BunkPO" Then 'sw
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        'If Not CreatePOPDF(ImportSeaLCLForm, "mx_Bunk", "B00001") Then Throw New ArgumentException(sErrDesc)
                                        LoadAndCreateCPO(ExportSeaLCLForm, "BunkPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If

                                End If

                                If pVal.ItemUID = "bt_ArmePO" Then 'sw
                                    If Not ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = "" And Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        LoadAndCreateCPO(ExportSeaLCLForm, "ArmedPurchaseOrder.srf")
                                    Else
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Exit Function
                                    End If
                                End If
                                '-----------------------------------------------------------------------------------------------------------'

                                If pVal.ItemUID = "mx_Cont" And ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    If pVal.Row > 0 Then
                                        If (oMatrix.IsRowSelected(pVal.Row)) = True Then
                                            modTrucking.rowIndex = CInt(pVal.Row)
                                            modTrucking.GetDataFromMatrixByIndex(ExportSeaLCLForm, oMatrix, modTrucking.rowIndex)
                                        End If
                                    End If

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
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
                                            If String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific.Value) Then
                                                'ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                                'ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                                'ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific.Value = Now.ToString("HH:mm")
                                                'If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific, ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            End If
                                            If String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value) Then
                                                'ExportSeaLCLForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                                'ExportSeaLCLForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                                'ExportSeaLCLForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                                'If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                            End If
                                            EnabledHeaderControls(ExportSeaLCLForm, True)
                                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                                ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"

                                                ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                                                ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                                                ' ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                                'ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                            End If
                                            'ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)
                                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet.DoQuery("SELECT CardName,VatIdUnCmp FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                            If oRecordSet.RecordCount > 0 Then
                                                ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_IUEN", 0, oRecordSet.Fields.Item("VatIdUnCmp").Value)
                                                ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_IComName", 0, oRecordSet.Fields.Item("CardName").Value)   '* Change Nyan Lin   "[@OBT_LCL06_PMAI]"
                                            End If
                                        End If

                                        If pVal.ItemUID = "ed_ShpAgt" Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)     '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)      '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                            'ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)            '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oRecordSet.DoQuery("SELECT CardName,VatIdUnCmp FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                            If oRecordSet.RecordCount > 0 Then
                                                ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_UEN", 0, oRecordSet.Fields.Item("VatIdUnCmp").Value)
                                                ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value.ToString) '* Change Nyan Lin   "[@OBT_TB005_EPERMIT]"
                                            End If
                                        End If


                                        If pVal.ItemUID = "ed_Trucker" Then
                                            Try
                                                If ExportSeaLCLForm.Items.Item("op_Inter").Specific.Selected = True Then
                                                    'ExportSeaLCLForm.Freeze(True)
                                                    ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRINTR").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & _
                                                                                                                    oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                                    ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRTEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString
                                                    ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRFAX").ValueEx = oDataTable.Columns.Item("fax").Cells.Item(0).Value.ToString
                                                    ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRMAIL").ValueEx = oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString
                                                    ExportSeaLCLForm.DataSources.UserDataSources.Item("TKRATTE").ValueEx = "" '25-3-2011
                                                    'ExportSeaLCLForm.Freeze(False)
                                                End If
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        If pVal.ItemUID = "ed_Vessel" Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)        '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                            ' ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)    '* Change Nyan Lin   "[@OBT_TB002_IMPSEALCL]"
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_VName", 0, ExportSeaLCLForm.Items.Item("ed_Vessel").Specific.String)                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        End If

                                        If pVal.ItemUID = "ed_AirLine" Then
                                            ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_AirLine", 0, oDataTable.Columns.Item("U_AirLine").Cells.Item(0).Value.ToString)
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
                                        oCombo = ExportSeaLCLForm.Items.Item("cb_PCode").Specific
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB002_EXPORT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                             '* Change Nyan Lin   "[@OBT_TB013_PINVDETAI]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_PCode", 0, oCombo.Selected.Value.ToString)                                       '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT").SetValue("U_PName", 0, oCombo.Selected.Description.ToString)                                 '* Change Nyan Lin   "[@OBT_TB010_PMAIN]"
                                    End If
                                    If pVal.ItemUID = "cb_BnkName" Then
                                        oCombo = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim test As String = "select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description
                                        oRecordSet.DoQuery("select GLAccount  from DSC1 A,ODSC B where A.BankCode =B.BankCode and B.BankCode= " & oCombo.Selected.Description)
                                        ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB009_EVOUCHER").SetValue("U_GLAC", 0, oRecordSet.Fields.Item("GLAccount").Value)                            '* Change Nyan Lin   "[@OBT_TB009_VOUCHER]"
                                        'oCombo.Selected.Description
                                    End If

                                    If pVal.ItemUID = "cb_PType" Then
                                        oCombo = ExportSeaLCLForm.Items.Item("cb_PType").Specific
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
                                'If ExportSeaLCLForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Or ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                '    Validateforform(pVal.ItemUID, ExportSeaLCLForm)
                                'End If
                                If BubbleEvent = False Then
                                    Validateforform(pVal.ItemUID, ExportSeaLCLForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

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

                Case "2000000025"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    ' ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_CrPo" Then
                            LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, "CranePurchaseOrder_Form.srf")
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If

                        End If

                    End If
                    'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    '    Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                    '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                    '    Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                    '    If pVal.FormUID = "Crane" Then
                    '        Try
                    '            If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                    '                ExportSeaLCLForm.DataSources.DBDataSources.Item("@CRANE").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)
                    '                ExportSeaLCLForm.DataSources.DBDataSources.Item("@CRANE").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)

                    '            End If
                    '            If pVal.ItemUID = "ed_ShpAgt" Then
                    '                ExportSeaLCLForm.DataSources.DBDataSources.Item("@CRANE").SetValue("U_ShpAgt", 0, oDataTable.GetValue(0, 0).ToString)
                    '            End If
                    '        Catch ex As Exception
                    '        End Try
                    '    End If

                    'End If
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If

                        End If
                    End If


                Case "2000000028"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    ' ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_ForCPO" Then
                            LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, "ForkliftPurchaseOrder.srf")
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If
                        End If
                    End If

                    'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    '    Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                    '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                    '    Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                    '    If pVal.FormUID = "ForkLift" Then
                    '        Try
                    '            If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                    '                ExportSeaLCLForm.DataSources.DBDataSources.Item("@FORKLIFT").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)
                    '                ExportSeaLCLForm.DataSources.DBDataSources.Item("@FORKLIFT").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)
                    '            End If
                    '            If pVal.ItemUID = "ed_ShpAgt" Then
                    '                ExportSeaLCLForm.DataSources.DBDataSources.Item("@FORKLIFT").SetValue("U_ShpAgt", 0, oDataTable.GetValue(0, 0).ToString)
                    '            End If
                    '        Catch ex As Exception
                    '        End Try
                    '    End If

                    'End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If

                        End If
                    End If


                Case "2000000031"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    ' ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_CrPO" Then
                            LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, "CourierPurchaseOrder.srf")
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If

                        End If
                    End If


                Case "2000000034" 'DGLP
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    ' ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_DGLP" Then
                            LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, "DGLPPurchaseOrder.srf")
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If

                        End If
                    End If

                Case "2000000037"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    ' ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_OutPO" Then
                            LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, "OutPurchaseOrder.srf")
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If

                        End If
                    End If
                Case "2000000040" 'COO
                    ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_COPO" Then
                            LoadAndCreate_PurChaseOrder(ExportSeaLCLForm, "CertificateOfOriginPPurchaseOrder.srf")
                            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                            End If

                        End If
                    End If

                    'MSW To Add New Button at ChooseFromList
                    'When User Click New button at ChooseFromList Form ,Setup Form Show to fill data.
                Case "9999"
                    ExportSeaLCLForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                        If pVal.ItemUID = "bt_VNew" Then
                            ExportSeaLCLForm.Close()
                            p_oSBOApplication.ActivateMenuItem("47644")
                        End If
                        If pVal.ItemUID = "bt_VoNew" Then
                            ExportSeaLCLForm.Close()
                            p_oSBOApplication.ActivateMenuItem("47644")
                        End If
                    End If
                    'MSW To Add New Button at ChooseFromList
            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
            DoExportSeaLCLItemEvent = RTN_SUCCESS

        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaLCLItemEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
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
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim bFlag As Boolean = False
        Dim FunctionName As String = "DoImportSeaLCLRightClickEvent"
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim formuid As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)

            '  ExportSeaLCLForm = p_oSBOApplication.Forms.GetForm(eventInfo.FormUID, 1)
            ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
            oMenuItem = p_oSBOApplication.Menus.Item("1280")
            oMenus = oMenuItem.SubMenus
            ExportSeaLCLForm.EnableMenu("772", True)
            ExportSeaLCLForm.EnableMenu("773", True)
            ExportSeaLCLForm.EnableMenu("775", True)
            If (p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND") And (eventInfo.ItemUID = "mx_Fumi" Or eventInfo.ItemUID = "mx_Crate" Or eventInfo.ItemUID = "mx_Armed" Or eventInfo.ItemUID = "mx_Bunk" Or eventInfo.ItemUID = "mx_Bok" Or eventInfo.ItemUID = "mx_TkrList") Then 'Truck PO
                oMatrix = ExportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                If eventInfo.BeforeAction = True Then
                    ExportSeaLCLForm.EnableMenu("772", False)
                    ExportSeaLCLForm.EnableMenu("773", False)
                    ExportSeaLCLForm.EnableMenu("775", False)
                    If eventInfo.ItemUID = "mx_TkrList" Then
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colInsDoc").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
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
                                If oMenus.Exists("EditVoc") Then
                                    oMenus.RemoveEx("EditVoc")
                                End If
                                currentRow = eventInfo.Row
                                ActiveMatrix = eventInfo.ItemUID
                            End If
                        End If

                    Else
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
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
                p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000037" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000040" Then
                If eventInfo.FormUID = "Courier" Or eventInfo.FormUID = "DGLP" Or eventInfo.FormUID = "Crane" Or eventInfo.FormUID = "ForkLift" Or eventInfo.FormUID = "Outrider" Or eventInfo.FormUID = "COO" Then 'Truck PO
                    oMatrix = ExportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                    If eventInfo.BeforeAction = True Then
                        ExportSeaLCLForm.EnableMenu("772", False)
                        ExportSeaLCLForm.EnableMenu("773", False)
                        ExportSeaLCLForm.EnableMenu("775", False)
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
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
            
            If eventInfo.EventType = SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK And eventInfo.ActionSuccess Then
                If eventInfo.FormUID = "OutriderGoodRecepit" Or eventInfo.FormUID = "COOGoodRecepit" Or eventInfo.FormUID = "CRANEGoodRecepit" Or eventInfo.FormUID = "CRANEGoodRecepit" Or eventInfo.FormUID = "DGLPGoodRecepit" Or eventInfo.FormUID = "ForkliftGoodRecepit" Or eventInfo.FormUID = "ForkliftGoodRecepit" Then
                    ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm
                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Item").Specific
                    Select Case ExportSeaLCLForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            RowFunction(ExportSeaLCLForm, "mx_Item", "V_-1", "@OBT_TB13_FFCGRITEM")
                    End Select
                End If
                If eventInfo.FormUID = "CRANEPURCHASEORDER" Or eventInfo.FormUID = "ForkLiftPURCHASEORDER" Or eventInfo.FormUID = "CourierPURCHASEORDER" Or eventInfo.FormUID = "DGLPPURCHASEORDER" Or eventInfo.FormUID = "OutriderPURCHASEORDER" Or eventInfo.FormUID = "COOPURCHASEORDER" Then
                    ExportSeaLCLForm = p_oSBOApplication.Forms.ActiveForm

                    oMatrix = ExportSeaLCLForm.Items.Item("mx_Item").Specific
                    Select Case ExportSeaLCLForm.Mode
                        Case SAPbouiCOM.BoFormMode.fm_ADD_MODE, SAPbouiCOM.BoFormMode.fm_OK_MODE, SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            RowFunction(ExportSeaLCLForm, "mx_Item", "V_-1", "@OBT_TB09_FFCPOITEM")
                    End Select
                End If

            End If
            If p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTSEALCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTAIRLCL" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "EXPORTLAND" Then
                If eventInfo.ItemUID = "mx_ShpInv" Then
                    ExportSeaLCLForm.EnableMenu("772", False)
                    ExportSeaLCLForm.EnableMenu("773", False)
                    ExportSeaLCLForm.EnableMenu("775", False)
                    oMatrix = ExportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                    If eventInfo.BeforeAction = True Then
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
                Else
                    If oMenus.Exists("EditShp") Then
                        p_oSBOApplication.Menus.RemoveEx("EditShp")
                        p_oSBOApplication.Forms.ActiveForm.EnableMenu("1293", False)
                    End If
                End If
                If eventInfo.ItemUID = "mx_Voucher" Then
                    oMatrix = ExportSeaLCLForm.Items.Item(eventInfo.ItemUID).Specific
                    ExportSeaLCLForm.EnableMenu("772", False)
                    ExportSeaLCLForm.EnableMenu("773", False)
                    ExportSeaLCLForm.EnableMenu("775", False)

                    If eventInfo.BeforeAction = True Then
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colVocNo").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If Not oMenus.Exists("EditVoc") Then
                                    p_oSBOApplication.Forms.ActiveForm.EnableMenu("1293", False)
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
            End If


           
        
            DoExportSeaLCLRightClickEvent = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaLCLRightClickEvent = RTN_ERROR
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
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean
        '   Return      :   No

        '*************************************************************



        pForm.Items.Item("ed_ShpAgt").Enabled = pValue
        pForm.Items.Item("ed_OBL").Enabled = pValue
        pForm.Items.Item("ed_HBL").Enabled = pValue
        pForm.Items.Item("ed_Conn").Enabled = pValue
        pForm.Items.Item("ed_Vessel").Enabled = pValue
        pForm.Items.Item("ed_Voy").Enabled = pValue
        pForm.Items.Item("cb_PCode").Enabled = pValue
        pForm.Items.Item("ed_ETDDate").Enabled = pValue
        pForm.Items.Item("ed_ETDHr").Enabled = pValue
        'MSW to edit New Ticket 07-09-2011
        If AlreadyExist("EXPORTAIRLCL") Then
            pForm.Items.Item("ed_ETADate").Enabled = pValue
            pForm.Items.Item("ed_ETAHr").Enabled = pValue
        End If
        'End MSW to edit New Ticket 07-09-2011
        pForm.Items.Item("ed_CrgDsc").Enabled = pValue
        pForm.Items.Item("ed_TotalM3").Enabled = pValue
        pForm.Items.Item("ed_TotalWt").Enabled = pValue
        pForm.Items.Item("ed_NOP").Enabled = pValue
        pForm.Items.Item("cb_PType").Enabled = pValue
        ' pForm.Items.Item("ed_JobNo").Enabled = pValue
        'MSW to Edit #1001
        If pValue = True And pForm.Items.Item("cb_JobType").Specific.Value = "" Then
            pForm.Items.Item("cb_JobType").Enabled = pValue
        End If
        'End MSW to Edit #1001
        pForm.Items.Item("cb_JbStus").Enabled = pValue
        If AlreadyExist("EXPORTAIRLCL") Then
            pForm.Items.Item("ed_AirLine").Enabled = pValue
            pForm.Items.Item("ed_FLNo").Enabled = pValue
            pForm.Items.Item("ed_AAgent").Enabled = pValue
            pForm.Items.Item("ed_PortDis").Enabled = pValue
        End If
        

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

    Private Function ClearComboData(ByRef pForm As SAPbouiCOM.Form, ByVal pComboUID As String) As Long

        ' **********************************************************************************
        '   Function    :   ClearComboData()
        '   Purpose     :   This function will be providing to clear Combo items  data for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pComboUID As String,
        '               :   ByVal DataSource As String, ByVal FieldAlias As String
        '   Return      :   0- FAILURE
        '                   1 - SUCCESS
        '*************************************************************

        Dim sFuncName As String = "ClearComboData()"
        Dim sErrDesc As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If pForm.Items.Item(pComboUID).Specific.ValidValues.Count > 0 Then
                For i As Integer = pForm.Items.Item(pComboUID).Specific.ValidValues.Count - 1 To 0 Step -1
                    pForm.Items.Item(pComboUID).Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
                pForm.DataSources.DBDataSources.Item("@OBT_TB007_EDISPATCH").SetValue("U_Dispatch", 0, vbNullString)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", sFuncName)
            ClearComboData = RTN_SUCCESS
        Catch ex As Exception
            ClearComboData = RTN_ERROR
            Call WriteToLogFile(ex.Message, sFuncName)
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

    Private Function Validateforform(ByVal ItemUID As String, ByVal ExportSeaLCLForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   Validateforform()
        '   Purpose     :   This function will be providing to validate Form   for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  : ByVal ItemUID As String, ByVal ExportSeaLCLForm As SAPbouiCOM.Form
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************]
        Try
            If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_Name").Specific.String) Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            ElseIf AlreadyExist("EXPORTSEALCL") And (ItemUID = "cb_PCode" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("cb_PCode").Specific.Value) Then
                p_oSBOApplication.SetStatusBarMessage("Must Select Port Of Loading", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
                'ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_ShpAgt").Specific.String) Then
                '    p_oSBOApplication.SetStatusBarMessage("Must Choose Shipping Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                '    Return True
                'ElseIf (ItemUID = "ed_JobNo" Or ItemUID = " ") And String.IsNullOrEmpty(ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.String) Then
                '    p_oSBOApplication.SetStatusBarMessage("Must Fill Job No", SAPbouiCOM.BoMessageTime.bmt_Short)
                '    Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
        
    End Function

    Private Sub LoadHolidayMarkUp(ByVal ExportSeaLCLForm As SAPbouiCOM.Form)

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
        If AlreadyExist("EXPORTAIRLCL") Then
            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETADay").Specific, ExportSeaLCLForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        End If
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ETDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ETDDay").Specific, ExportSeaLCLForm.Items.Item("ed_ETDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_DspCDay").Specific, ExportSeaLCLForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_ADay").Specific, ExportSeaLCLForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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

        ' **********************************************************************************
        '   Function    :   PDFAlert()
        '   Purpose     :   This function will be providing to show PDF Alert Function   for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByVal oActiveForm As SAPbouiCOM.Form
        '   Return      :   False- FAILURE
        '               :   True - SUCCESS
        ' **********************************************************************************

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
                MoveWindow(myprocess.MainWindowHandle, p_oSBOApplication.Forms.ActiveForm.Left + p_oSBOApplication.Forms.ActiveForm.Width + 6, p_oSBOApplication.Forms.ActiveForm.Top + 68, 300, 618, True)
            End If
        Catch ex As Exception

        End Try


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
        oPayForm.Items.Item("ed_PayRate").Visible = True
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPayForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("VOUCHER", oRecordSet)
        oPayForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_InvDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        oPayForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()

        'If HolidayMarkUp(oPayForm, oPayForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
        '=============================================================================
        'Function   : DeleteMatrixRow()
        'Purpose    : This function to delete matrixrow of line Id 
        'Parameters : ByRef oActiveForm As SAPbouiCOM.Form,ByRef oMatrix As SAPbouiCOM.Matrix,
        '           : ByVal objDataSource As String, ByVal oColumn As String       
        'Return     : No
        '=============================================================================

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
#Region "PO & GR"
    Private Sub BindingChooseFromList(ByRef ctrlParent As Object, ByVal cflID As String, ByVal ctrlName As String, Optional ByVal fieldName As String = vbNullString)
        '=============================================================================
        'Function   : BindingChooseFromList()
        'Purpose    : This function bind ChooseFromLise and EditText Box
        'Parameters : ByRef ctrlParent As Object
        '             ByVal cflID As String
        '             ByVal ctrlName As String
        '             Optional ByVal fieldName As String = vbNullString
        'Return     :   False   = Failure
        '               True    = Success
        'Author     : @channyeinkyaw
        '=============================================================================
        Dim oEditTextBox As SAPbouiCOM.EditText
        Try
            oEditTextBox = ctrlParent.Items.Item(ctrlName).Specific
            oEditTextBox.ChooseFromListUID = cflID
            oEditTextBox.ChooseFromListAlias = fieldName
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
                CPOForm.Items.Item("bt_Resend").Visible = False
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
    Public Sub LoadFormXML(ByRef pstrFileName As String)
        Dim objXmlDoc As Xml.XmlDocument
        objXmlDoc = New Xml.XmlDocument
        objXmlDoc = New Xml.XmlDocument
        'objXmlDoc.Load(IO.Directory.GetParent((Application.StartupPath)).ToString & "\" & pstrFileName)
        objXmlDoc.Load(Application.StartupPath & "\" & pstrFileName)
        'objXmlDoc.Load(IO.Directory.GetParent(Application.StartupPath).ToString & "/" & pstrFileName)
        p_oSBOApplication.LoadBatchActions(objXmlDoc.InnerXml)
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
    Public Sub PrintDoc(ByVal PathName As String)
        Dim newThread As New Thread(AddressOf ShowPrintDialogBrowser)
        newThread.Priority = ThreadPriority.Highest
        newThread.SetApartmentState(ApartmentState.STA)
        newThread.Start(PathName)
        newThread.Join()
    End Sub
    Public Sub ShowPrintDialogBrowser(ByVal PathName As String)

        Dim MyProcs() As System.Diagnostics.Process
        Dim FileName As String = ""

        Dim PrintFile As New PrintDialog
        PrintFile.UseEXDialog = True
        Try

            PrintFile.AllowSelection = True
            PrintFile.ShowNetwork = True

            MyProcs = Process.GetProcessesByName("SAP Business One")

            If MyProcs.Length = 1 Then

                For i As Integer = 0 To MyProcs.Length - 1

                    Dim MyWindow As New PrinterDialog(MyProcs(i).MainWindowHandle)
                    Dim ret As DialogResult = PrintFile.ShowDialog() '(MyWindow)


                    If ret = DialogResult.OK Then
                        'Dim p As Process = New Process()
                        Dim info As ProcessStartInfo = New ProcessStartInfo()
                        info.FileName = PathName
                        info.Verb = "Print"
                        info.Arguments = PrintFile.PrinterSettings.PrinterName
                        info.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden

                        'p.Start(info)
                        Process.Start(info)
                    End If
                Next

            End If

        Catch ex As Exception

            ' SBO_Application.MessageBox(ex.Message)
            '  FileName = PrintFile.PrinterSettings.PrinterName

        Finally

            ' PrintFile.Dispose()

        End Try

    End Sub
    Public Sub LoadPopUpForm(ByRef sPicName As String)
        LoadFormXML("ShowBigImg.srf")
        ActiveForm = p_oSBOApplication.Forms.ActiveForm
        Dim oPictureBox As SAPbouiCOM.PictureBox = ActiveForm.Items.Item("picBig").Specific
        oPictureBox.Picture = p_fmsSetting.PicturePath & sPicName.Substring(0, 1) & "\" & sPicName & ".jpg"
    End Sub
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
                            If oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value = "None" Then
                                oPurchaseDocument.Lines.VatGroup = "ZI"
                            Else
                                oPurchaseDocument.Lines.VatGroup = oMatrix.Columns.Item("colIGST").Cells.Item(i).Specific.Value
                            End If
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

    Private Sub CalAmtArmetEsc(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim cMatrix As SAPbouiCOM.Matrix
        Dim itemTotal As Double
        cMatrix = oActiveForm.Items.Item("mx_Item").Specific
        cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value = Convert.ToDouble(cMatrix.Columns.Item("colIQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(cMatrix.Columns.Item("colIPrice").Cells.Item(Row).Specific.Value)
        itemTotal += cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value
        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_POITPD", 0, itemTotal)
    End Sub

    Private Sub CalRateArmetEsc(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim cMatrix As SAPbouiCOM.Matrix
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim oEditText As SAPbouiCOM.EditText
        Dim itemTotal As Double
        cMatrix = oActiveForm.Items.Item("mx_Item").Specific
        oRecordset = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordset.DoQuery("select Rate from ovtg where Code='" + cMatrix.Columns.Item("colIGST").Cells.Item(Row).Specific.Value + "'")
        Dim Rate As Double = 0.0
        Dim GSTAMT As Double = 0.0
        Dim total As Double = 0.0
        If oRecordset.RecordCount < 0 Then
            Rate = 0
        Else
            oRecordset.MoveFirst()
            Rate = oRecordset.Fields.Item("Rate").Value
            GSTAMT = Convert.ToDouble(cMatrix.Columns.Item("colIQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(cMatrix.Columns.Item("colIPrice").Cells.Item(Row).Specific.Value)
            cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value = GSTAMT
            GSTAMT = (Convert.ToDouble(GSTAMT) / 100.0) * Rate
            total = cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.Value - GSTAMT
            cMatrix.Columns.Item("colITotal").Editable = True
            cMatrix.Columns.Item("colITotal").Cells.Item(Row).Specific.value = total
            oActiveForm.Items.Item("ed_PONo").Specific.active = True
            cMatrix.Columns.Item("colITotal").Editable = False
            Try

                Dim MCount As Integer = CInt(cMatrix.RowCount.ToString())
                For MStart As Integer = 1 To MCount Step 1
                    itemTotal += Convert.ToDouble(cMatrix.Columns.Item("colITotal").Cells.Item(MStart).Specific.value)
                Next
            Catch ex As Exception

            End Try
            oActiveForm.DataSources.DBDataSources.Item("@OBT_TB08_FFCPO").SetValue("U_POITPD", 0, itemTotal)
            '  oActiveForm.Items.Item("ed_TPDue").Specific.value = itemTotal.ToString()
        End If
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

                oPurchaseDocument.DocDate = CDate(Now.Date) '.ToString("yyyyMMdd"))
                oPurchaseDocument.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO
                oPurchaseDocument.TaxDate = CDate(Now.Date)
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
    Private Sub RowFunction(ByRef oForm As SAPbouiCOM.Form, ByVal mxUID As String, ByVal FieldAlias As String, ByVal TableName As String)
        Dim ObjDbDataSource As SAPbouiCOM.DBDataSource
        Dim oMatrix As SAPbouiCOM.Matrix
        Try
            oMatrix = oForm.Items.Item(mxUID).Specific
            ObjDbDataSource = oForm.DataSources.DBDataSources.Item(TableName) '@OBT_TB09_FFCPOITEM
            Dim MCount As Integer = CInt(oMatrix.RowCount)
            For MStart As Integer = 1 To MCount Step 1
                With ObjDbDataSource
                    '.SetValue(FieldAlias, MStart, MStart.ToString)
                    oMatrix.Columns.Item(FieldAlias.ToString()).Cells.Item(MStart).Specific.Value = MStart.ToString()

                End With
                'Debug.Print(ObjDbDataSource.GetValue(FieldAlias, MStart))
            Next
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
        End Try
    End Sub

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
        If AlreadyExist("EXPORTSEALCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
        ElseIf AlreadyExist("EXPORTSEAFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
        ElseIf AlreadyExist("EXPORTAIRLCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
        ElseIf AlreadyExist("EXPORTLAND") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
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

        If AddChooseFromListByFilter(oShpForm, "cflPart", False, "PART", "U_BPCode", SAPbouiCOM.BoConditionOperation.co_EQUAL, oActiveForm.Items.Item("ed_Code").Specific.Value) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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

                    .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(rowIndex).Specific.Value)
                    .SetValue("U_SerNo", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(rowIndex).Specific.Value)
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
            pForm.Items.Item("ed_DOM").Specific.Value = pMatrix.Columns.Item("colDOM").Cells.Item(Index).Specific.Value
            'pForm.Items.Item("ed_Part").Specific.Value = pMatrix.Columns.Item("colPart").Cells.Item(Index).Specific.Value
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
        ' **********************************************************************************

        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        If AlreadyExist("EXPORTSEALCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
        ElseIf AlreadyExist("EXPORTSEAFCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
        ElseIf AlreadyExist("EXPORTAIRLCL") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
        ElseIf AlreadyExist("EXPORTLAND") Then
            oActiveForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
        End If

        ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item(DataSource)
        '  ObjDBDataSource.Offset = 0

        rowIndex = pMatrix.GetNextSelectedRow

        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
            rowIndex = 1
        End If
        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            If pMatrix.RowCount = 1 And pMatrix.Columns.Item("colDocNum").Cells.Item(pMatrix.RowCount).Specific.Value = "" Then
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

#Region "Validate Job Number"
    Private Sub ValidateJobNumber(ByRef oActiveForm As SAPbouiCOM.Form, ByRef BubbleEvent As Boolean)
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim jobType As String = String.Empty
        Dim strJobNum As String = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        oRecordSet.DoQuery(" SELECT U_JOBMODE FROM [@OBT_TB002_EXPORT] WHERE U_JOBNUM='" & strJobNum & "'")
        If oRecordSet.RecordCount > 0 Then
            jobType = oRecordSet.Fields.Item("U_JOBMODE").Value.ToString.Substring(7, 3)
            If AlreadyExist("EXPORTSEALCL") Then
                If jobType = "Air" Then
                    oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                    p_oSBOApplication.SetStatusBarMessage("Cannot Load Air Job Number in Export Sea LCL.", SAPbouiCOM.BoMessageTime.bmt_Short)
                    BubbleEvent = False
                End If
            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                If jobType = "Sea" Then
                    oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                    p_oSBOApplication.SetStatusBarMessage("Cannot Load Sea Job Number in Export Air LCL.", SAPbouiCOM.BoMessageTime.bmt_Short)
                    BubbleEvent = False
                End If
            End If
        ElseIf AlreadyExist("EXPORTLAND") Then
            If jobType = "Sea" Then
                oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                p_oSBOApplication.SetStatusBarMessage("Cannot Load Sea Job Number in Export Air LCL.", SAPbouiCOM.BoMessageTime.bmt_Short)
                BubbleEvent = False
            End If
        End If



    End Sub
#End Region

#Region "Load_Button_Forms"

    Private Sub LoadButtonForm(ByVal parentForm As SAPbouiCOM.Form, ByVal FormName As String, ByVal KeyName As String)

        Try

      
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
        oActiveForm.EnableMenu("1284", False)
        oActiveForm.EnableMenu("1286", False)
        oActiveForm.EnableMenu("1283", False) 'MSW 01-04-2011
            oActiveForm.EnableMenu("774", False)
            oActiveForm.EnableMenu("771", False)
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


        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Email"

    Dim reportuti As New clsReportUtilities
    Dim pdfFilename As String = ""
    Dim originpdf As String = ""
    Dim mainFolder As String = ""
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
        'pdffilepath = "C:\Users\UNIQUE\Desktop\For NL\Outrider.pdf"

        rptDocument.SetParameterValue("@DocEntry", PONo)
        reportuti.SetDBLogIn(rptDocument)
        If Not pdffilepath = String.Empty Then
            reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument, True)
        End If

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
#End Region

    Public Sub LoadExportSeaLCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)

        '=============================================================================
        'Function   : LoadExportSeaLCLForm()
        'Purpose    : This function will be providing to call and load ExporeSealLcLForm  
        '             and show  ExporeSealLcLForm for process
        'Parameters : Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString
        '             Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        'Return     : No
        '=============================================================================

        Dim ExportSeaLCLForm As SAPbouiCOM.Form = Nothing
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
        Dim FunctionName As String = "DoImportSeaLCLMenuEvent()"
        Dim sErrDesc As String = String.Empty
        Dim jobType As String = String.Empty
        Try
          

            If Title.Substring(7, 3) = "Sea" Then
                If Title.Substring(11, 3) = "LCL" Then
                    LoadFromXML(p_oSBOApplication, "ExportSeaLCLv1.srf")
                ElseIf Title.Substring(11, 3) = "FCL" Then
                    LoadFromXML(p_oSBOApplication, "ExportSeaFCLv1.srf")
                End If

            ElseIf Title.Substring(7, 3) = "Air" Then
                LoadFromXML(p_oSBOApplication, "ExportAirLCLv1.srf")
                'parentForm = "EXPORTAIRLCL"
            ElseIf Title.Substring(7, 4) = "Land" Then
                LoadFromXML(p_oSBOApplication, "ExportLandv1.srf")
                'parentForm = "EXPORTAIRLCL"
            End If

            If AlreadyExist("EXPORTSEALCL") Then
                ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEALCL")
                jobType = "Export Sea LCL"
            ElseIf AlreadyExist("EXPORTSEAFCL") Then
                ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTSEAFCL")
                jobType = "Export Sea FCL"
            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTAIRLCL")
                jobType = "Export Air"
            ElseIf AlreadyExist("EXPORTLAND") Then
                ExportSeaLCLForm = p_oSBOApplication.Forms.Item("EXPORTLAND")
                jobType = "Export Land"
            End If
            ExportSeaLCLForm.Title = Title ' MSW To Edit New Ticket 07-09-2011
            ExportSeaLCLForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
            Try
                ExportSeaLCLForm.EnableMenu("1288", True)
                ExportSeaLCLForm.EnableMenu("1289", True)
                ExportSeaLCLForm.EnableMenu("1290", True)
                ExportSeaLCLForm.EnableMenu("1291", True)
                ExportSeaLCLForm.EnableMenu("1284", False)
                ExportSeaLCLForm.EnableMenu("1286", False)
                ExportSeaLCLForm.EnableMenu("1281", False)
                ExportSeaLCLForm.EnableMenu("1284", False)
                ExportSeaLCLForm.EnableMenu("1286", False)
                ExportSeaLCLForm.EnableMenu("1283", False) 'MSW 01-04-2011
                ExportSeaLCLForm.EnableMenu("1292", False)
                ExportSeaLCLForm.EnableMenu("1293", False)
                ExportSeaLCLForm.EnableMenu("4870", False)
                ExportSeaLCLForm.EnableMenu("771", False)
                ExportSeaLCLForm.EnableMenu("772", True)
                ExportSeaLCLForm.EnableMenu("773", True)
                ExportSeaLCLForm.EnableMenu("774", False)
                ExportSeaLCLForm.EnableMenu("775", True)
                If FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                End If
                ExportSeaLCLForm.DataBrowser.BrowseBy = "ed_DocNum"
                ExportSeaLCLForm.Items.Item("fo_Prmt").Specific.Select()
                ExportSeaLCLForm.Items.Item("bt_Crane").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Foklit").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Courier").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_DGLP").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Outer").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Alert").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Certifi").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Excel").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_Label").Enabled = False
                ExportSeaLCLForm.Items.Item("bt_A6Label").Enabled = False


                ExportSeaLCLForm.Freeze(True)
            Catch ex As Exception
                MessageBox.Show(ex.ToString())
            End Try

            '---------------------------------------------- Vouncher (OMM) 7-Mar-2011-------------------------------------------------------------------------'
            'Dim oColumn As SAPbouiCOM.Column
            'oMatrix = ExportSeaLCLForm.Items.Item("mx_ChCode").Specific
            'ExportSeaLCLForm.DataSources.DataTables.Add("DTCharges")
            'dtmatrix = ExportSeaLCLForm.DataSources.DataTables.Item("DTCharges")
            'dtmatrix.Columns.Add("ChCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
            'dtmatrix.Columns.Add("AcCode", SAPbouiCOM.BoFieldsType.ft_Text)
            'dtmatrix.Columns.Add("Desc", SAPbouiCOM.BoFieldsType.ft_Text)
            'dtmatrix.Columns.Add("Amount", SAPbouiCOM.BoFieldsType.ft_Price)
            'dtmatrix.Columns.Add("GST", SAPbouiCOM.BoFieldsType.ft_Text)
            'dtmatrix.Columns.Add("GSTAmt", SAPbouiCOM.BoFieldsType.ft_Price)
            'dtmatrix.Columns.Add("NoGST", SAPbouiCOM.BoFieldsType.ft_Price)
            'dtmatrix.Columns.Add("SeqNo", SAPbouiCOM.BoFieldsType.ft_Text)
            'dtmatrix.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text) 'To Add Item Code

            'oColumn = oMatrix.Columns.Item("colGST")
            'AddGSTComboData(oColumn)
            'oColumn.DataBind.Bind("DTCharges", "GST")
            'oColumn = oMatrix.Columns.Item("colChCode")
            'oColumn.DataBind.Bind("DTCharges", "ChCode")
            'AddChooseFromList(ExportSeaLCLForm, "CCode", False, "UDOCHCODE")
            '' AddUserDataSrc(oActiveForm, "test", "test", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 0)
            'oColumn.ChooseFromListUID = "CCode"

            'oColumn = oMatrix.Columns.Item("colAcCode")
            'oColumn.DataBind.Bind("DTCharges", "AcCode")
            'oColumn = oMatrix.Columns.Item("colVDesc")
            'oColumn.DataBind.Bind("DTCharges", "Desc")
            'oColumn = oMatrix.Columns.Item("colAmount")
            'oColumn.DataBind.Bind("DTCharges", "Amount")
            'oColumn = oMatrix.Columns.Item("colGSTAmt")
            'oColumn.DataBind.Bind("DTCharges", "GSTAmt")
            'oColumn = oMatrix.Columns.Item("colNoGST")
            'oColumn.DataBind.Bind("DTCharges", "NoGST")
            'oColumn = oMatrix.Columns.Item("V_-1")
            'oColumn.DataBind.Bind("DTCharges", "SeqNo")
            'oColumn = oMatrix.Columns.Item("colICode")    'To Add Item Code
            'oColumn.DataBind.Bind("DTCharges", "ItemCode") 'To Add Item Code

            'dtmatrix.Rows.Add(1)
            'oMatrix.Clear()
            'dtmatrix.SetValue("SeqNo", 0, 1)
            'oMatrix.LoadFromDataSource()
            '-------------------------------------Charge Code--------------------------------------------------------------'
            'Vendor in Voucher
            'If AddChooseFromList(ExportSeaLCLForm, "PAYMENT", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'ExportSeaLCLForm.Items.Item("ed_VedName").Specific.ChooseFromListUID = "PAYMENT"
            'ExportSeaLCLForm.Items.Item("ed_VedName").Specific.ChooseFromListAlias = "CardName"
            ''ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
            ''End

            'If AddUserDataSrc(ExportSeaLCLForm, "CASH", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(ExportSeaLCLForm, "CHEQUE", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'oOpt = ExportSeaLCLForm.Items.Item("op_Cash").Specific
            'oOpt.DataBind.SetBound(True, "", "CASH")
            'oOpt = ExportSeaLCLForm.Items.Item("op_Cheq").Specific
            'oOpt.DataBind.SetBound(True, "", "CHEQUE")
            'oOpt.GroupWith("op_Cash")
            ''OMM voucher
            'oCombo = ExportSeaLCLForm.Items.Item("cb_BnkName").Specific
            'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery("Select BankName From ODSC")
            'If oRecordSet.RecordCount > 0 Then
            '    oRecordSet.MoveFirst()
            '    While oRecordSet.EoF = False
            '        oCombo.ValidValues.Add(oRecordSet.Fields.Item("BankName").Value, "")
            '        oRecordSet.MoveNext()
            '    End While
            'End If
            'oCombo = ExportSeaLCLForm.Items.Item("cb_PayCur").Specific
            'oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRecordSet.DoQuery("SELECT CurrCode FROM OCRN")
            'If oRecordSet.RecordCount > 0 Then
            '    oRecordSet.MoveFirst()
            '    While oRecordSet.EoF = False
            '        oCombo.ValidValues.Add(oRecordSet.Fields.Item("CurrCode").Value, "")
            '        oRecordSet.MoveNext()
            '    End While
            'End If
            '------------------------------------------------------------------------------------------------------------------------------------------'
            ExportSeaLCLForm.Items.Item("bt_GenPO").Enabled = False
            EnabledHeaderControls(ExportSeaLCLForm, False)
            EnabledMaxtix(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("mx_TkrList").Specific, False)
            ExportSeaLCLForm.PaneLevel = 7

            If FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If Not Title = vbNullString Then
                    If Title = "Export Air LCL" Then
                        ExportSeaLCLForm.Title = "Export Air"
                    End If

                End If
                If AlreadyExist("EXPORTSEALCL") Then
                    ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Sea LCL"
                ElseIf AlreadyExist("EXPORTSEAFCL") Then
                    ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Sea FCL"
                ElseIf AlreadyExist("EXPORTAIRLCL") Then
                    ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Air"
                ElseIf AlreadyExist("EXPORTLAND") Then
                    ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export LAND"

                End If
                'ExportSeaLCLForm.Items.Item("ed_JType").Specific.Value = "Export Sea LCL"
                ' ExportSeaLCLForm.DataSources.DBDataSources.Item("@OBT_TB005_EPERMIT ").SetValue("U_TranMode", 0, "Sea") 'MSW LCL Change
                ExportSeaLCLForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                ExportSeaLCLForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                ExportSeaLCLForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                ExportSeaLCLForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString

            End If

            If HolidayMarkUp(ExportSeaLCLForm, ExportSeaLCLForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, ExportSeaLCLForm.Items.Item("ed_JbDay").Specific, ExportSeaLCLForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'MSW to edit New Ticket 07-09-2011
            'If AlreadyExist("EXPORTAIRLCL") Then
            '    If AddChooseFromList(ExportSeaLCLForm, "cflAirLine", False, "AIRLINE") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            '    ExportSeaLCLForm.Items.Item("ed_AirLine").Specific.ChooseFromListUID = "cflAirLine"
            '    ExportSeaLCLForm.Items.Item("ed_AirLine").Specific.ChooseFromListAlias = "U_AirLine"
            'End If



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
            'ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
            'ExportSeaLCLForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"

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
                'ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW For JobType Table 
            End If
            If Not ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("EX")
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
            If AddUserDataSrc(ExportSeaLCLForm, "EUC", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) ' MSW To Edit New Ticket

            'MSW 14-09-2011 Truck PO
            If AddUserDataSrc(ExportSeaLCLForm, "TKRINS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRIRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "TKRRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(ExportSeaLCLForm, "PODOCNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'MSW 14-09-2011 Truck PO

            ExportSeaLCLForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
            ExportSeaLCLForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
            ExportSeaLCLForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
            ExportSeaLCLForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
            ExportSeaLCLForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
            ExportSeaLCLForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")
            ExportSeaLCLForm.Items.Item("ed_EUC").Specific.DataBind.SetBound(True, "", "EUC") ' MSW To Edit New Ticket
            ExportSeaLCLForm.Items.Item("ee_InsRmsk").Specific.DataBind.SetBound(True, "", "TKRIRMK") 'MSW 14-09-2011 Truck PO
            ExportSeaLCLForm.Items.Item("ee_TkrIns").Specific.DataBind.SetBound(True, "", "TKRINS") 'MSW 14-09-2011 Truck PO
            ExportSeaLCLForm.Items.Item("ee_Rmsk").Specific.DataBind.SetBound(True, "", "TKRRMK") 'MSW 14-09-2011 Truck PO
            ExportSeaLCLForm.Items.Item("ed_PODocNo").Specific.DataBind.SetBound(True, "", "PODOCNO") 'MSW 14-09-2011 Truck PO


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
            If ExportSeaLCLForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                Dim tempItem As SAPbouiCOM.Item
                tempItem = ExportSeaLCLForm.Items.Item("ed_JobNo")
                tempItem.Enabled = True
                ExportSeaLCLForm.Items.Item("ed_JobNo").Specific.Value = JobNo
                ExportSeaLCLForm.Items.Item("1").Click()
                ' MSW To Edit New Ticket 07-09-2011
                If AlreadyExist("EXPORTSEALCL") Then
                    ExportSeaLCLForm.Title = "Export Sea-LCL " + ExportSeaLCLForm.Items.Item("cb_JobType").Specific.Value
                ElseIf AlreadyExist("EXPORTSEAFCL") Then
                    ExportSeaLCLForm.Title = "Export Sea-FCL " + ExportSeaLCLForm.Items.Item("cb_JobType").Specific.Value
                ElseIf AlreadyExist("EXPORTAIRLCL") Then
                    ExportSeaLCLForm.Title = "Export Air " + ExportSeaLCLForm.Items.Item("cb_JobType").Specific.Value
                ElseIf AlreadyExist("EXPORTLAND") Then
                    ExportSeaLCLForm.Title = "Export Land " + ExportSeaLCLForm.Items.Item("cb_JobType").Specific.Value
                End If
                ' MSW To Edit New Ticket 07-09-2011
                tempItem.Enabled = False
            End If
            If ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "Internal" Then
                ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
            ElseIf ExportSeaLCLForm.Items.Item("ed_DMode").Specific.Value = "External" Then
                ExportSeaLCLForm.Items.Item("op_DspExtr").Specific.Selected = True
            Else
                ExportSeaLCLForm.Items.Item("op_DspIntr").Specific.Selected = True
            End If
            ExportSeaLCLForm.Freeze(False)

            Select Case ExportSeaLCLForm.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_AddIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = False
                    ExportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = False 'MSW 10-09-2011
                Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    ExportSeaLCLForm.Items.Item("bt_AmdIns").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_AddIns").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_DelIns").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_PrntIns").Enabled = True

                    ExportSeaLCLForm.Items.Item("bt_BokPO").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_DrBL").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_SpOrder").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_DGD").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_CrPO").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_CPO").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_BunkPO").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_ArmePO").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_ShpInv").Enabled = True
                    ExportSeaLCLForm.Items.Item("bt_PrntDis").Enabled = True 'MSW 10-09-2011
            End Select
        Catch ex As Exception

        End Try
    End Sub
    Private Function CreatePOPDF(ByVal oActiveForm As SAPbouiCOM.Form, ByVal matrixName As String, ByVal iCode As String) As Boolean

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
        'Dim docEntry As String = GetNewKey("FCPO", oRecordSet)
        'Dim oGeneralService As SAPbobsCOM.GeneralService
        'Dim oGeneralData As SAPbobsCOM.GeneralData
        'Dim oChild As SAPbobsCOM.GeneralData
        'Dim oChildren As SAPbobsCOM.GeneralDataCollection
        CreatePOPDF = False
        Try
            If matrixName = "mx_Armed" Then
                pdfFilename = "ArmEdit"
                originpdf = "ArmEdit.pdf"
            ElseIf matrixName = "mx_Out" Then
                pdfFilename = "OutRider"
                originpdf = "OutRider (2).pdf"
            ElseIf matrixName = "mx_COO" Then
                pdfFilename = "COO"
                originpdf = "Certificate.pdf"
            ElseIf matrixName = "mx_DGLP" Then
                pdfFilename = "DGLP"
                originpdf = "DGLP_0002.pdf"
            End If
            mainFolder = p_fmsSetting.DocuPath
            'jobNo = jobNo 'oActiveForm.Items.Item("ed_JobNo").Specific.Value (Must be change Job No)
            pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
            Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
            Dim pdfoutputfile As FileStream = New FileStream(pdffilepath, System.IO.FileMode.Create)
            Dim formfiller As iTextSharp.text.pdf.PdfStamper = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
            Dim ac As iTextSharp.text.pdf.AcroFields = formfiller.AcroFields
            If AlreadyExist("EXPORTSEALCL") Then
                oForm = p_oSBOApplication.Forms.GetForm("EXPORTSEALCL", 1)
            ElseIf AlreadyExist("EXPORTSEAFCL") Then
                oForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
            ElseIf AlreadyExist("EXPORTAIRLCL") Then
                oForm = p_oSBOApplication.Forms.GetForm("EXPORTAIRLCL", 1)
            ElseIf AlreadyExist("EXPORTLAND") Then
                oForm = p_oSBOApplication.Forms.GetForm("EXPORTLAND", 1)
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

            ElseIf matrixName = "mx_Out" Then
                ac.SetField("txtCompanyName", "Midwest Freight & Transportation Pte Ltd")
                ac.SetField("txtApplicantName", oActiveForm.Items.Item("cb_SInA").Specific.value)
                ac.SetField("txtCustomerContactNo", oActiveForm.Items.Item("ed_CNo").Specific.value)
                ac.SetField("Serial No", "")
                ac.SetField("TP Permit No", "")
                ac.SetField("txtHeadphone", "")
                ac.SetField("Road Tax expiry date", "")
                ac.SetField("Length", "")
                ac.SetField("Width", "")
                ac.SetField("Height", "")
                ac.SetField("Name of Insurance", "")
                ac.SetField("Insurance Policy No", "")
                ac.SetField("Insurance Policy Expiry Date", "")
                ac.SetField("Number of Units", "")
                ac.SetField("txtEscort From", "")
                ac.SetField("txtEscort To", "")
                ac.SetField("txtDateofEscort", "")
                ac.SetField("txtTimeCommenceEscort", oActiveForm.Items.Item("ed_TTime").Specific.value)
                If oActiveForm.Items.Item("ch_Fax").Specific.checked Then
                    ac.SetField("txtFaxNo", oActiveForm.Items.Item("40").Specific.value)
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
    Public Sub CopyFile(ByVal SourceDirectory As String, ByVal DestinationDirectory As String)

        ' **********************************************************************************
        '   Function    :   CopyFile()
        '   Purpose     :   This function will be providing to copy file to other path of computer
        '               
        '   Parameters  :   ByVal DestinationDirectory As Strig 
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************

        Try

            System.IO.File.Copy(SourceDirectory, DestinationDirectory, True)

        Catch ex As Exception

        End Try
    End Sub
    Public Class PrinterDialog
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle

            Get

                Return _hwnd

            End Get

        End Property
    End Class

#Region "Booking Preview"
    Private Sub DeletePDF()

        ' **********************************************************************************
        '   Function    :   DeletePDF()
        '   Purpose     :   This function will be providing to delete PDF file 
        '   Parameters  :   No
        '   Return      :   No
        '                 
        ' **********************************************************************************


        Dim di As New IO.DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName)
        If di.Exists Then
            Directory.Delete(AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName, True)
        End If

    End Sub
    Public Sub CopyFolder(ByVal foldername As String)

        ' **********************************************************************************
        '   Function    :   CopyFolder()
        '   Purpose     :   This function will be providing to copy folder and file 
        '   Parameters  :   ByVal foldername As String
        '   Return      :   No
        '                 
        ' **********************************************************************************
        Try

            'Create Folder
            Dim di As DirectoryInfo = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName)
            If di.Exists Then
                ' Indicate that it already exists.
                Console.WriteLine("That path exists already.")
                Return
            End If
            ' Try to create the directory.
            di.Create()

            'Copy Original PDF to created folder
            Dim NewPath As String
            For Each NewPath In Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory + "PDF\" + foldername, "*.*", SearchOption.AllDirectories)
                File.Copy(NewPath, NewPath.Replace(AppDomain.CurrentDomain.BaseDirectory + "PDF\" + foldername, AppDomain.CurrentDomain.BaseDirectory + "Booking\" + System.Environment.MachineName))
            Next
        Catch

        End Try
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
                    Dim sql As String = "SELECT dbo.[@OBT_TB002_EXPORT].U_Vessel, dbo.[@OBT_TB002_EXPORT].U_Voyage, dbo.[@OBT_TB002_EXPORT].U_PName, dbo.[@OBT_TB002_EXPORT].U_OceanBL, dbo.[@OBT_TB002_EXPORT].U_HouseBL, dbo.[@OBT_TB002_EXPORT].U_Conn, dbo.[@OBT_TB002_EXPORT].U_TotalWt, dbo.[@OBT_TB002_EXPORT].U_TotalM3,dbo.[@OBT_TB002_EXPORT].U_CrgDsc, dbo.OCRD.CardCode, dbo.OCRD.CardName, dbo.OCRD.Address, dbo.OCRD.Fax, dbo.OCRD.Phone1, dbo.OCRD.Phone2 FROM dbo.[@OBT_TB002_EXPORT] INNER JOIN dbo.OCRD ON dbo.[@OBT_TB002_EXPORT].U_Code = dbo.OCRD.CardCode WHERE dbo.[@OBT_TB002_EXPORT].U_JobNum =" & FormatString(pform.Items.Item("ed_JobNo").Specific.Value)
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
                    Dim sql As String = "SELECT dbo.[@OBT_TB002_EXPORT].U_Vessel, dbo.[@OBT_TB002_EXPORT].U_PName, dbo.OCRD.CardCode, dbo.OCRD.CardName, dbo.OCRD.Address, dbo.OCRD.Phone1,dbo.OCRD.Fax FROM dbo.[@OBT_TB002_EXPORT] INNER JOIN dbo.OCRD ON dbo.[@OBT_TB002_EXPORT].U_Code = dbo.OCRD.CardCode WHERE dbo.[@OBT_TB002_EXPORT].U_JobNum =" & FormatString(pform.Items.Item("ed_JobNo").Specific.Value)
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.RecordCount > 0 Then
                        oRecordSet.MoveFirst()
                        While oRecordSet.EoF = False
                            'addressChangeForm.SetField("ConsigneeName", oRecordSet.Fields.Item("CardName").Value.ToString())
                            'addressChangeForm.SetField("ConsigneeAddress", oRecordSet.Fields.Item("Address").Value.ToString())
                            'addressChangeForm.SetField("Telephone", oRecordSet.Fields.Item("Phone1").Value.ToString())
                            'addressChangeForm.SetField("Fax", oRecordSet.Fields.Item("Fax").Value.ToString())
                            'addressChangeForm.SetField("Vessel", oRecordSet.Fields.Item("U_Vessel").Value.ToString())
                            'addressChangeForm.SetField("PortOfLoading", oRecordSet.Fields.Item("U_PName").Value.ToString())

                            addressChangeForm.SetField("ConsigneeName", oRecordSet.Fields.Item("CardName").Value.ToString())
                            addressChangeForm.SetField("ConsigneeAddress", oRecordSet.Fields.Item("Address").Value.ToString())
                            addressChangeForm.SetField("ConsigneePhone", oRecordSet.Fields.Item("Phone1").Value.ToString())
                            addressChangeForm.SetField("ConsigneeFax", oRecordSet.Fields.Item("Fax").Value.ToString())
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
                    Dim sql As String = "SELECT dbo.[@OBT_TB002_EXPORT].U_Vessel,dbo.[@OBT_TB002_EXPORT].U_Voyage, dbo.[@OBT_TB002_EXPORT].U_PName, dbo.OCRD.CardCode, dbo.OCRD.CardName, dbo.OCRD.Address, dbo.OCRD.Phone1,dbo.OCRD.Fax FROM dbo.[@OBT_TB002_EXPORT] INNER JOIN dbo.OCRD ON dbo.[@OBT_TB002_EXPORT].U_Code = dbo.OCRD.CardCode WHERE dbo.[@OBT_TB002_EXPORT].U_JobNum =" & FormatString(pform.Items.Item("ed_JobNo").Specific.Value)
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
#End Region
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

        Catch ex As Exception

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
    Private Function CancelTruckingPurchaseOrder(ByRef oActiveForm As SAPbouiCOM.Form) As Boolean

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
            If oPurchaseDocument.GetByKey(oActiveForm.Items.Item("ed_PONo").Specific.Value.ToString) Then
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
        rptPath = Application.StartupPath.ToString & "\Trucking Instruction Export.rpt"
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
        If AlreadyExist("EXPORTSEALCL") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Export Sea LCL.rpt"
        ElseIf AlreadyExist("EXPORTAIRLCL") Then
            rptPath = Application.StartupPath.ToString & "\A6 Label Export Air.rpt"
        ElseIf AlreadyExist("EXPORTLAND") Then
            '(rptPath = Application.StartupPath.ToString & "\A6 Label Export Air.rpt")
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
                        .SetValue("U_PONo", .Offset, DocLastKey)
                        '.SetValue("U_PONo", .Offset, oRecordSet.Fields.Item("U_PONo").Value.ToString)
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
                        .SetValue("U_PONo", .Offset, DocLastKey)
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
End Module


