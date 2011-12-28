Option Explicit On

Imports System.Xml
Imports System.Threading
Imports System.Net.Mail
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Runtime.InteropServices
Imports System
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
    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    Public TaskDS As SAPbouiCOM.DBDataSource

    Dim docfolderpath As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oUDSs As SAPbouiCOM.UserDataSources
    Dim BoolResize As Boolean = False
    Dim sql As String = ""
    Dim rowIndex As Integer
    Dim IsAmend As Boolean = False
    'Dim TruckingInstruction As clsTruckingInstruction

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
    Private DocStatus As String

    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String
    Dim vedCurCode As String
    Private Enum ImageType
        TIFF
        BMP
        JPG
        PNG
    End Enum

    Dim dtmatrix As SAPbouiCOM.DataTable
    Public gridindex As String
    Private fname As String

    Public Function DoImportSeaFCLFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long

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

        DoImportSeaFCLFormDataEvent = RTN_ERROR
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim oVesselForm As SAPbouiCOM.Form = Nothing
        Dim oPayForm As SAPbouiCOM.Form = Nothing
        Dim oGRForm As SAPbouiCOM.Form = Nothing
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
        Dim matrixName As String = ""
        Dim sFuncName As String = "DoImportSeaFCLFormDataEvent()"
        Try

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            Select Case BusinessObjectInfo.FormTypeEx

                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000050" 'MSW to Edit New Ticket
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
                                matrixName = "mx_Fork"
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB029_FORKLIFT", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                matrixName = "mx_Armed"
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB030_ARMES", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000010" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                matrixName = "mx_Crane"
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB031_CRANE", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                matrixName = "mx_Bunk"
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB032_BUNKER", True, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                matrixName = "mx_TkrList"
                                If Not PopulateTruckingPOToEditTab(oActiveForm, sql, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            End If
                            'If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql) Then Throw New ArgumentException(sErrDesc)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PONo = " + FormatString(DocLastKey) + " WHERE DocEntry = " + FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value))
                            If matrixName = "mx_Armed" Or matrixName = "mx_TkrList" Then 'MSW 14-09-2011 Truck PO
                                SendAttachFile(oActiveForm, oPOForm)
                            Else
                                CreatePDF(oActiveForm, matrixName)
                            End If

                        End If

                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                       "[@OBT_TB08_FFCPO] where DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                            If BusinessObjectInfo.FormTypeEx = "2000000021" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB029_FORKLIFT", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000009" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB030_ARMES", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000010" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB031_CRANE", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000015" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB032_BUNKER", False, oPOForm) Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000050" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                      "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                      "from  [@OBT_TB08_FFCPO] a inner join [@OBT_TB006_TRUCKING] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oPOForm.Items.Item("ed_CPOID").Specific.Value)
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB006_TRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'MSW 14-09-2011 Truck PO
                            End If
                            If Not UpdatePurchaseOrder(oPOForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                        End If
                        'End If
                    End If

                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000051" 'Truck PO
                    If BusinessObjectInfo.ActionSuccess = True Then
                        oGRForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                        If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                            If Not CreateGoodsReceiptPO(oGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)

                            If BusinessObjectInfo.FormTypeEx = "2000000012" Then
                                oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB029_FORKLIFT", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000013" Then
                                oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB030_ARMES", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000014" Then
                                oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB031_CRANE", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000016" Then
                                oMatrix = oActiveForm.Items.Item("mx_Bunk").Specific
                                Dim sql As String = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
                                                     "[@OBT_TB08_FFCPO] where U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB032_BUNKER", False, oGRForm) Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            ElseIf BusinessObjectInfo.FormTypeEx = "2000000051" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                     "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                     "from  [@OBT_TB08_FFCPO] a inner join [@OBT_TB006_TRUCKING] b on a.DocEntry=b.U_PODocNo where a.U_PONo = " & FormatString(oMatrix.Columns.Item("colPONo").Cells.Item(currentRow).Specific.Value())
                                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB006_TRUCKING") Then Throw New ArgumentException(sErrDesc)
                                'Truck PO
                            End If

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
                
                    'MSW
                Case "IMPORTSEAFCL"
                    oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
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
                            End If
                        End If
                    End If

                    If BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            Dim JobLastDocEntry As Integer
                            Dim ObjectCode As String = String.Empty
                            sql = "select top 1 Docentry from [@OBT_TB002_IMPSEALCL] order by docentry desc"
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

                            sql = "Update [@OBT_TB002_IMPSEALCL] set U_JbDocNo=" & JobLastDocEntry & ",U_JobNum = '" & NewJobNo & "' Where DocEntry=" & FrDocEntry & ""
                            oRecordSet.DoQuery(sql)
                            p_oSBOApplication.SetStatusBarMessage("Actual Job Number is " & NewJobNo, SAPbouiCOM.BoMessageTime.bmt_Short, False)

                            sql = "Insert Into [@OBT_FREIGHTDOCNO] (DocEntry,DocNum,U_JobNo,U_JobMode,U_JobType,U_JbStus,U_FrDocNo,U_JbDate,U_ObjType,U_CusCode,U_CusName,U_ShpCode,U_ShpName) Values " & _
                                "(" & JobLastDocEntry & _
                                    "," & JobLastDocEntry & _
                                   "," & IIf(NewJobNo <> "", FormatString(NewJobNo), "NULL") & _
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
                        End If
                    ElseIf BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            sql = "Update [@OBT_FREIGHTDOCNO] set U_JbStus='" & oActiveForm.Items.Item("cb_JbStus").Specific.Value & "' Where U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & ""
                            oRecordSet.DoQuery(sql)
                        End If
                    End If

                    'End MSW

                    'MSW New Button at Choose From List 22-03-2011
                Case "VESSEL"
                    Dim vesCode As String = String.Empty
                    Dim voyNo As String = String.Empty
                    Dim oImportSeaFCLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.Item("IMPORTSEAFCL")
                    If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) And BusinessObjectInfo.BeforeAction = False Then
                        If BusinessObjectInfo.ActionSuccess = True Then
                            oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select top 1* from [@OBT_TB018_VESSEL] order by DocEntry desc")
                            If oRecordSet.RecordCount > 0 Then
                                vesCode = oRecordSet.Fields.Item("Name").Value.ToString
                                voyNo = oRecordSet.Fields.Item("U_Voyage").Value.ToString
                            End If
                            oImportSeaFCLForm.Items.Item("ed_Vessel").Specific.Value = vesCode
                            oImportSeaFCLForm.Items.Item("ed_Voy").Specific.Value = voyNo

                            'Try
                            '    oImportSeaLCLForm.Items.Item("ed_Vessel").Specific.Active = True
                            'Catch ex As Exception

                            'End Try

                        End If
                    Else
                        If (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE) And BusinessObjectInfo.BeforeAction = False Then
                            If BusinessObjectInfo.ActionSuccess = True Then
                                oActiveForm = p_oSBOApplication.Forms.Item(BusinessObjectInfo.FormUID)
                                oImportSeaFCLForm.Items.Item("ed_Vessel").Specific.Value = ""
                                oImportSeaFCLForm.Items.Item("ed_Voy").Specific.Value = ""

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
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            DoImportSeaFCLFormDataEvent = RTN_SUCCESS
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            DoImportSeaFCLFormDataEvent = RTN_ERROR
        End Try
    End Function

    Public Function DoImportSeaFCLMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   DoImportSeaFCLMenuEvent
        '   Purpose     :   This function will be providing to proceed validating for
        '                   All  Menu Event information
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
        DoImportSeaFCLMenuEvent = RTN_ERROR
        Dim oActiveForm, oPayForm, oShpForm As SAPbouiCOM.Form
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oShpMatrix As SAPbouiCOM.Matrix
        Dim oCheckBox As SAPbouiCOM.CheckBox
        Dim oOpt As SAPbouiCOM.OptionBtn
        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String = "SBO_Application_MenuEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            Select Case pVal.MenuUID
                Case "mnuImportSeaLCL"
                    If pVal.BeforeAction = False Then
                        LoadImportSeaFCLForm()
                    End If
                Case "1292"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                            oActiveForm = p_oSBOApplication.Forms.ActiveForm
                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                                If oMatrix.Columns.Item("colChCode1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    If Not AddNewRow(oActiveForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
                                    ' SBO_Application.MessageBox("Row Count" + oMatrix.RowCount.ToString + " / " + "Visual RowCount" + oMatrix.VisualRowCount.ToString)
                                End If
                            End If
                            
                        ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            If pVal.BeforeAction = True Then
                                BubbleEvent = False
                            End If
                            'oActiveForm = p_oSBOApplication.Forms.ActiveForm
                            ''-------------------------For Payment(omm)------------------------------------------'
                            'If (oActiveForm.PaneLevel = 3) Then
                            '    oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                            '    RowAddToMatrix(oActiveForm, oMatrix)
                            'End If
                            ''----------------------------------------------------------------------------------'
                        End If
                    End If

                Case "1293"
                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                        oActiveForm = p_oSBOApplication.Forms.ActiveForm
                        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
                            ' Dim lRow As Long
                            If pVal.BeforeAction = True Then
                                If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                    BubbleEvent = False
                                End If

                                If BubbleEvent = True Then
                                    DeleteMatrixRow(oActiveForm, oMatrix, "@OBT_TB032_VDETAIL", "V_-1")
                                    BubbleEvent = False
                                    If Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    CalculateTotal(oActiveForm, oMatrix)
                                End If
                            End If
                        Else
                            BubbleEvent = False
                        End If


                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000005" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000020" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000021" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000009" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000010" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000015" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000050" Then
                        oActiveForm = p_oSBOApplication.Forms.ActiveForm
                        'If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        oMatrix = oActiveForm.Items.Item("mx_Item").Specific
                        ' Dim lRow As Long
                        If pVal.BeforeAction = True Then
                            If oMatrix.GetNextSelectedRow = oMatrix.RowCount Then
                                BubbleEvent = False
                            End If

                            If BubbleEvent = True Then
                                DeleteMatrixRow(oActiveForm, oMatrix, "@OBT_TB09_FFCPOITEM", "LineId")
                                BubbleEvent = False
                                If Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                End If
                                CalculateTotalPO(oActiveForm, oMatrix)
                            End If
                        End If
                        'Else
                        'BubbleEvent = False
                        ' End If

                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                        If pVal.BeforeAction = True Then
                            oActiveForm = p_oSBOApplication.Forms.ActiveForm
                            'BubbleEvent = False
                            'MSW To Edit
                            Dim shpDocEntry As String
                            oActiveForm = p_oSBOApplication.Forms.ActiveForm
                            oMatrix = oActiveForm.Items.Item("mx_ShpInv").Specific
                            If oMatrix.GetNextSelectedRow > 0 Then
                                If oMatrix.Columns.Item("colDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString <> "" Then
                                    shpDocEntry = oMatrix.Columns.Item("colSDocNum").Cells.Item(oMatrix.GetNextSelectedRow).Specific.Value.ToString
                                    DeleteMatrixRow(oActiveForm, oMatrix, "@OBT_LCL20_SHPINV", "V_-1")
                                    DeleteUDO(oActiveForm, "SHIPPINGINV", shpDocEntry)
                                    oActiveForm.Items.Item("1").Click()
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    oActiveForm.Items.Item("bt_ShpInv").Enabled = True
                                End If
                            End If
                            BubbleEvent = False
                            'End MSW To Edit
                        End If
                        

                    ElseIf p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000007" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000011" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000012" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000013" Or _
                        p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000014" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000016" Or p_oSBOApplication.Forms.ActiveForm.TypeEx = "2000000051" Then
                        If pVal.BeforeAction = True Then
                            BubbleEvent = False
                        End If
                    End If
                Case "1281"
                    If pVal.BeforeAction = False Then
                        'If SBO_Application.Forms.Item("IMPORTSEAFCL").Selected = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oActiveForm = p_oSBOApplication.Forms.ActiveForm

                            oActiveForm.Items.Item("ed_JobNo").Enabled = True
                            oActiveForm.Items.Item("ed_JobNo").Specific.Active = True
                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oActiveForm.Items.Item("ch_POD").Enabled = False
                                'oActiveForm.Items.Item("ed_Wrhse").Enabled = True
                            End If
                            If AddChooseFromListByOption(p_oSBOApplication.Forms.Item("IMPORTSEAFCL"), True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If
                    End If

                Case "1282"
                    If pVal.BeforeAction = False Then
                        'If SBO_Application.Forms.Item("IMPORTSEAFCL").Selected = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oActiveForm = p_oSBOApplication.Forms.ActiveForm
                            EnabledHeaderControls(oActiveForm, False) 'MSW 26-03-2011
                            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            oRecordSet.DoQuery("select NNM1.NextNumber from ONNM JOIN NNM1 ON (ONNM.ObjectCode = NNM1.ObjectCode And ONNM.DfltSeries = NNM1.Series) where ONNM.ObjectCode = 'IMPORTSEAFCL'")
                            If oRecordSet.RecordCount > 0 Then
                                '  oActiveForm.Items.Item("ed_JobNo").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString
                                oActiveForm.Items.Item("ed_DocNum").Specific.Value = oRecordSet.Fields.Item("NextNumber").Value.ToString 'MSW 01-06-2011  for Job Type Table
                            End If
                            oActiveForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString 'Prep By for Header
                            oActiveForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("IM")
                            oActiveForm.Items.Item("ed_JType").Specific.Value = "Import Sea FCL" 'MSW 30-05-2011 for Job Type Table
                            oActiveForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                            oActiveForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                            oActiveForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
                            If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            oActiveForm.Items.Item("ch_POD").Enabled = False
                        End If
                    End If

                Case "1288", "1289", "1290", "1291"
                    If pVal.BeforeAction = True Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oActiveForm = p_oSBOApplication.Forms.ActiveForm
                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                            End If

                        End If
                    End If

                Case "EditVoc"
                    If pVal.BeforeAction = False Then
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            LoadPaymentVoucher(oActiveForm)
                            oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                            oPayForm = p_oSBOApplication.Forms.ActiveForm
                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            oPayForm.Items.Item("ed_DocNum").Visible = True
                            oPayForm.Items.Item("ed_DocNum").Enabled = True
                            oPayForm.Items.Item("ed_DocNum").Specific.Value = oMatrix.Columns.Item("colVDocNum").Cells.Item(currentRow).Specific.Value.ToString
                            oPayForm.Items.Item("cb_PayCur").Specific.Active = True
                            oPayForm.Items.Item("ed_DocNum").Visible = False
                            oPayForm.Items.Item("ed_DocNum").Enabled = False
                            oPayForm.DataBrowser.BrowseBy = "ed_DocNum"
                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                            oPayForm.Items.Item("ed_VedName").Enabled = False
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
                        If p_oSBOApplication.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            LoadShippingInvoice(oActiveForm)
                            oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                            ' If SBO_Application.Forms.ActiveForm.TypeEx = "IMPORTSEAFCL" Then
                            oMatrix = oActiveForm.Items.Item("mx_ShpInv").Specific
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
                    End If
                Case "EditCPO"
                    If pVal.BeforeAction = False Then
                        If currentRow > 0 Then
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            oMatrix = oActiveForm.Items.Item(ActiveMatrix).Specific
                            If ActiveMatrix = "mx_TkrList" Then
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadTruckingPO(oActiveForm, RPOsrfname)
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oMatrix = oActiveForm.Items.Item(ActiveMatrix).Specific
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
                                    LoadAndCreateCPO(oActiveForm, RPOsrfname)
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            oMatrix = oActiveForm.Items.Item(ActiveMatrix).Specific

                            If ActiveMatrix = "mx_TkrList" Then
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    LoadAndCreateCGR(oActiveForm, RGRsrfname)
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    CGRForm = p_oSBOApplication.Forms.ActiveForm
                                    If Not FillDataToGoodsReceipt(oActiveForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
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
                                    LoadAndCreateCGR(oActiveForm, RGRsrfname)
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    CGRForm = p_oSBOApplication.Forms.ActiveForm
                                    If Not FillDataToGoodsReceipt(oActiveForm, ActiveMatrix, "colPONo", "colDocNo", currentRow, "@OBT_TB12_FFCGR", CGRForm) Then Throw New ArgumentException(sErrDesc)
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
                            oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                            oMatrix = oActiveForm.Items.Item(ActiveMatrix).Specific
                            If ActiveMatrix = "mx_TkrList" Then
                                If oMatrix.Columns.Item("colPStatus").Cells.Item(currentRow).Specific.Value.ToString = "Open" Then
                                    sql = "select a.DocEntry,a.U_PONo,a.U_PORMKS,a.U_POIRMKS,a.U_ColFrm,a.U_TkrIns,a.U_TkrTo,a.U_POStatus,b.U_InsDate,b.U_Mode,b.U_Trucker,b.U_VehNo," & _
                                      "b.U_EUC,b.U_Attent,b.U_Tel,b.U_Fax,b.U_Email,b.U_TkrDate,b.U_TkrTime " & _
                                      "from  [@OBT_TB08_FFCPO] a inner join [@OBT_TB006_TRUCKING] b on a.DocEntry=b.U_PODocNo where a.DocEntry = " & FormatString(oMatrix.Columns.Item("colDocNo").Cells.Item(currentRow).Specific.Value)
                                    If Not CancelPO(oActiveForm, sql, ActiveMatrix) Then Throw New ArgumentException(sErrDesc)
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oActiveForm.Items.Item("1").Click()
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    If Not CancelPO(oActiveForm, sql, ActiveMatrix) Then Throw New ArgumentException(sErrDesc)
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oActiveForm.Items.Item("1").Click()
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                    
            End Select
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            DoImportSeaFCLMenuEvent = RTN_SUCCESS
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            DoImportSeaFCLMenuEvent = RTN_ERROR
        End Try
    End Function

    Public Sub LoadImportSeaFCLForm(Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString, Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE)

        ' **********************************************************************************
        '   Function    :   LoadImportSeaFCLForm()
        '   Purpose     :   This function provide to Load and show ImportSeaFCLForm Forms from  main menu 
        '   Parameters  :   Optional ByVal JobNo As String = vbNullString, Optional ByVal Title As String = vbNullString,
        '               :   Optional ByVal FormMode As SAPbouiCOM.BoFormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE 
        '   return      :   No          
        ' **********************************************************************************
        Dim oActiveForm, oPayForm, oShpForm As SAPbouiCOM.Form
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CGRForm As SAPbouiCOM.Form = Nothing
        Dim oEditText As SAPbouiCOM.EditText
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oCheckBox As SAPbouiCOM.CheckBox
        Dim oOpt As SAPbouiCOM.OptionBtn
        Dim oItem As SAPbouiCOM.Item
        Dim sErrDesc As String = ""
        Dim sFuncName As String = "SBO_Application_MenuEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            If Not LoadFromXML(p_oSBOApplication, "ImportSeaFCLv1.srf") Then Throw New ArgumentException(sErrDesc)

            oActiveForm = p_oSBOApplication.Forms.ActiveForm
            oActiveForm.Title = Title ' MSW To Edit New Ticket 07-09-2011
            'BubbleEvent = False
            '-------------------------------------Charge Code--------------------------------------------------------------'
            If AddChooseFromList(oActiveForm, "Charge", False, "UDOCHCODE") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_ChCode").Specific.ChooseFromListUID = "Charge"
            oActiveForm.Items.Item("ed_ChCode").Specific.ChooseFromListAlias = "U_CName"
            '--------------------------------------------------------------------------------------------------------------'
            'Vendor in Voucher

            'End
            oActiveForm.EnableMenu("1288", True)
            oActiveForm.EnableMenu("1289", True)
            oActiveForm.EnableMenu("1290", True)
            oActiveForm.EnableMenu("1291", True)
            oActiveForm.EnableMenu("1284", False)
            oActiveForm.EnableMenu("1286", False)
            oActiveForm.EnableMenu("1281", True)
            oActiveForm.EnableMenu("1283", False)
            oActiveForm.EnableMenu("1292", False)
            oActiveForm.EnableMenu("1293", False)
            oActiveForm.EnableMenu("4870", False)
            oActiveForm.EnableMenu("771", False)
            oActiveForm.EnableMenu("772", True)
            oActiveForm.EnableMenu("773", True)
            oActiveForm.EnableMenu("775", True)
            oActiveForm.EnableMenu("774", False)
            If FormMode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            End If
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

            If FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                If Not Title = vbNullString Then
                    oActiveForm.Title = Title
                End If
                oActiveForm.Items.Item("ed_PrepBy").Specific.Value = p_oDICompany.UserName.ToString 'Prep By for Header
                oActiveForm.Items.Item("ed_JType").Specific.Value = "Import Sea FCL"
                oActiveForm.Items.Item("ed_JbDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                oActiveForm.Items.Item("ed_JbDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                oActiveForm.Items.Item("ed_JbHr").Specific.Value = Now.ToString("HH:mm")
            End If
            '-----------------DISPATCH Date----------------------------------------------------------------------'
            'oActiveForm.Items.Item("ed_DspDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
            'oActiveForm.Items.Item("ed_DspDay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
            'oActiveForm.Items.Item("ed_DspHr").Specific.Value = Now.ToString("HH:mm")
            'If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
            'oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
            'oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"

            ''MSW to add

            ''If AddChooseFromListCondition(oActiveForm, "CVendor", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S", "U_VType", "Crane") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ''oActiveForm.Items.Item("ed_CVendor").Specific.ChooseFromListUID = "CVendor"
            ''oActiveForm.Items.Item("ed_CVendor").Specific.ChooseFromListAlias = "CardCode"


            ''If AddChooseFromListCondition(oActiveForm, "FVendor", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S", "U_VType", "ForkLift") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ''oActiveForm.Items.Item("ed_FVendor").Specific.ChooseFromListUID = "FVendor"
            ''oActiveForm.Items.Item("ed_FVendor").Specific.ChooseFromListAlias = "CardCode"
            ''End MSW to add

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

            'Get jobNo NO
            If Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                oActiveForm.Items.Item("ed_JobNo").Specific.Value = GetJobNumber("IM")
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
            ' MSW To Edit New Ticket
            If AddUserDataSrc(oActiveForm, "TKRINS", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRIRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "TKRRMK", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddUserDataSrc(oActiveForm, "PODOCNO", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            ' MSW To Edit New Ticket
            If AddUserDataSrc(oActiveForm, "EUC", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc) ' MSW To Edit New Ticket
            oActiveForm.Items.Item("ed_Attent").Specific.DataBind.SetBound(True, "", "TKRATTE")
            oActiveForm.Items.Item("ed_TkrTel").Specific.DataBind.SetBound(True, "", "TKRTEL")
            oActiveForm.Items.Item("ed_Fax").Specific.DataBind.SetBound(True, "", "TKRFAX")
            oActiveForm.Items.Item("ed_Email").Specific.DataBind.SetBound(True, "", "TKRMAIL")
            oActiveForm.Items.Item("ee_ColFrm").Specific.DataBind.SetBound(True, "", "TKRCOL")
            oActiveForm.Items.Item("ee_TkrTo").Specific.DataBind.SetBound(True, "", "TKRTO")
            oActiveForm.Items.Item("ed_EUC").Specific.DataBind.SetBound(True, "", "EUC") ' MSW To Edit New Ticket

            oActiveForm.Items.Item("ee_InsRmsk").Specific.DataBind.SetBound(True, "", "TKRIRMK") ' MSW To Edit New Ticket
            oActiveForm.Items.Item("ee_TkrIns").Specific.DataBind.SetBound(True, "", "TKRINS") ' MSW To Edit New Ticket
            oActiveForm.Items.Item("ee_Rmsk").Specific.DataBind.SetBound(True, "", "TKRRMK") ' MSW To Edit New Ticket
            oActiveForm.Items.Item("ed_PODocNo").Specific.DataBind.SetBound(True, "", "PODOCNO") 'MSW 14-09-2011 Truck PO

            If AddUserDataSrc(oActiveForm, "DSDISP", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("op_DspIntr").Specific.DataBind.SetBound(True, "", "DSDISP")
            oActiveForm.Items.Item("op_DspExtr").Specific.DataBind.SetBound(True, "", "DSDISP")
            oActiveForm.Items.Item("op_DspExtr").Specific.GroupWith("op_DspIntr")

            If AddChooseFromList(oActiveForm, "CFLTKRE", False, 171) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            If AddChooseFromList(oActiveForm, "CFLTKRV", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            ''msw
            'If AddUserDataSrc(oActiveForm, "Permit", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "Voucher", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "Yard", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "Cont", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "Truck", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "Dispatch", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'If AddUserDataSrc(oActiveForm, "Charges", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'oActiveForm.Items.Item("ch_Permit").Specific.DataBind.SetBound(True, "", "Permit")
            'oActiveForm.Items.Item("ch_Voucher").Specific.DataBind.SetBound(True, "", "Voucher")
            'oActiveForm.Items.Item("ch_RYard").Specific.DataBind.SetBound(True, "", "Yard")
            'oActiveForm.Items.Item("ch_Conta").Specific.DataBind.SetBound(True, "", "Cont")
            'oActiveForm.Items.Item("ch_Truck").Specific.DataBind.SetBound(True, "", "Truck")
            'oActiveForm.Items.Item("ch_Disp").Specific.DataBind.SetBound(True, "", "Dispatch")
            'oActiveForm.Items.Item("ch_OCharge").Specific.DataBind.SetBound(True, "", "Charges")


            ''oActiveForm.DataSources.UserDataSources.Item("Permit").Value = "Y"
            ''oActiveForm.DataSources.UserDataSources.Item("Voucher").Value = "Y"

            'oActiveForm.Items.Item("ch_Permit").Specific.Checked = True
            'oActiveForm.Items.Item("ch_Voucher").Specific.Checked = True
            'oActiveForm.Items.Item("ch_RYard").Specific.Checked = True
            'oActiveForm.Items.Item("ch_Conta").Specific.Checked = True
            'oActiveForm.Items.Item("ch_Truck").Specific.Checked = True
            'oActiveForm.Items.Item("ch_Disp").Specific.Checked = True
            'oActiveForm.Items.Item("ch_OCharge").Specific.Checked = True
            ' ---------------------------10-1-2011-------------------------------------
            '----------Recordset for Binding colCType of Matrix (mx_Cont)-------------
            '--------------------------SYMA & OMM-------------------------------------
            oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
            Dim oColum As SAPbouiCOM.Column
            oColum = oMatrix.Columns.Item("colCType1")
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
            oMatrix.Columns.Item("colCSeqNo1").Cells.Item(1).Specific.Value = 1

            'If AddUserDataSrc(oActiveForm, "Code", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException
            'oMatrix = oActiveForm.Items.Item("mx_ChCode").Specific
            'oColum = oMatrix.Columns.Item("colChCode")
            'oColum.DataBind.SetBound(True, "", "Code")
            'oColum.ChooseFromListUID = "CFLTKRE"
            'oColum.ChooseFromListAlias = "Code"
            'oMatrix.LoadFromDataSource()
            oActiveForm.EnableMenu("1283", False) 'MSW 01-04-2011
            oActiveForm.EnableMenu("5907", True)
            '------------------------------For License Info SYMA & OMM (13/Jan/2011)---------------'
            oMatrix = oActiveForm.Items.Item("mx_License").Specific
            oMatrix.AddRow()
            oMatrix.Columns.Item("colLicNo").Cells.Item(1).Specific.Value = 1
            '-------------------------------------------------------------------------------------'
            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                Dim tempItem As SAPbouiCOM.Item
                tempItem = oActiveForm.Items.Item("ed_JobNo")
                tempItem.Enabled = True
                oActiveForm.Items.Item("ed_JobNo").Specific.Value = JobNo
                oActiveForm.Items.Item("1").Click()
                oActiveForm.Title = "Import Sea-FCL " + oActiveForm.Items.Item("cb_JobType").Specific.Value ' MSW To Edit New Ticket 07-09-2011
                tempItem.Enabled = False
            End If

            oActiveForm.Freeze(False)
            Select Case oActiveForm.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE Or SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    oActiveForm.Items.Item("bt_AmdIns").Enabled = False
                    oActiveForm.Items.Item("bt_AddIns").Enabled = False
                    oActiveForm.Items.Item("bt_DelIns").Enabled = False
                    oActiveForm.Items.Item("bt_PrntIns").Enabled = False
                    oActiveForm.Items.Item("bt_PrtDisp").Enabled = False
                    oActiveForm.Items.Item("bt_A6Label").Enabled = False

                Case SAPbouiCOM.BoFormMode.fm_OK_MODE Or SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                    oActiveForm.Items.Item("bt_AddIns").Enabled = True
                    oActiveForm.Items.Item("bt_DelIns").Enabled = True
                    oActiveForm.Items.Item("bt_PrntIns").Enabled = True
                    oActiveForm.Items.Item("bt_PrtDisp").Enabled = True
                    oActiveForm.Items.Item("bt_A6Label").Enabled = True

                    oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                    oActiveForm.Items.Item("bt_CranePO").Enabled = True
                    oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                    oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                    oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                    oActiveForm.Items.Item("bt_ShpInv").Enabled = True
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Function DoImportSeaFCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   DoImportSeaFCLItemEvent
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
        'Dim oShpMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oCombo As SAPbouiCOM.ComboBox


        Dim sFuncName As String = "SBO_Application_ItemEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            Select Case pVal.FormTypeEx
                Case "2000000007", "2000000011", "2000000012", "2000000013", "2000000014", "2000000016", "2000000051"  'MSW 14-09-2011 Truck PO       ' CGR --> Custom Goods Receipt

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
                            If pVal.ColUID = "colIGST" Then
                                CalRatePO(CGRForm, pVal.Row)
                            End If

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
                                'CGRForm.Items.Item("ed_TPlace").Specific.Active = True
                            End If
                            'MSW 
                            If pVal.ItemUID = "1" Then
                                oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                ' When click the Add button in AddMode of Custom Purchase Order form
                                ' need to also trigger the item pressed event of Main Export Form according by Customize Biz Logic
                                If pVal.Action_Success = True Then
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                                        CGRForm = p_oSBOApplication.Forms.ActiveForm
                                        CGRForm.Close()
                                        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If
                                        oActiveForm.Items.Item("1").Click()
                                    End If
                                    oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                                    oActiveForm.Items.Item("bt_CranePO").Enabled = True
                                    oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                                    oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                                    oActiveForm.Items.Item("bt_PrtDisp").Enabled = True
                                End If
                            End If


                            ''If pVal.ItemUID = "1" Then
                            ''    If pVal.BeforeAction = True Then
                            ''        If p_oSBOApplication.Forms.ActiveForm.TypeEx = pVal.FormTypeEx Then
                            ''            'CGRForm = p_oSBOApplication.Forms.ActiveForm
                            ''            'If CGRForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            ''            '    If Not CreateGoodsReceiptPO(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
                            ''            'End If
                            ''        End If
                            ''    End If
                            ''End If
                        End If

                    End If

                Case "2000000005", "2000000020", "2000000021", "2000000009", "2000000010", "2000000015", "2000000050"  ''MSW To Edit New Ticket       ' CPO --> Custom Purchase Order"
                    CPOForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)
                    Try
                        CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                    If pVal.BeforeAction = False Then
                        If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                            CPOForm.Items.Item("ed_Code").Enabled = False
                            CPOForm.Items.Item("ed_Name").Enabled = False
                            CPOForm.Items.Item("cb_SInA").Enabled = True
                        End If
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
                                ' CPOForm.Items.Item("ed_TPlace").Specific.Active = True
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
                                            If CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                                CPOForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                                CPOForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If
                                            oActiveForm.Items.Item("1").Click()
                                            'MSW 14-09-2011 Truck PO
                                            If pVal.FormTypeEx = "2000000050" Then
                                                EnabledTruckerForExternal(oActiveForm, False)
                                                oActiveForm.Items.Item("bt_PrntIns").Enabled = False
                                            End If
                                            'MSW 14-09-2011 Truck PO
                                        End If
                                    End If
                                    oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                                    oActiveForm.Items.Item("bt_CranePO").Enabled = True
                                    oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                                    oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                                    oActiveForm.Items.Item("bt_PrtDisp").Enabled = True 'MSW 10-09-2011
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
                    Try
                        oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                        ' oShpMatrix = oShpForm.Items.Item("mx_ShipInv").Specific
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
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                                    oShpForm.Items.Item("bt_PPView").Visible = True
                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "SHIPPINGINV" Then
                                
                                        If oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            oShpForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oShpForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB03_EXPSHPINV] set " & _
                                            " U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oShpForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            oShpForm.Close()

                                            oActiveForm.Items.Item("1").Click()
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                                oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    oActiveForm.Items.Item("bt_ShpInv").Enabled = True
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
                                'If Not reportuti.SendMailDoc("maysandar@sap-infotech.com,kyawzin@sap-infotech.com", "tuntunaung@sap-infotech.com", "ShippingInvoice", "Testing Mail", pdffilepath) Then
                                '    p_oSBOApplication.MessageBox("Send Message Fail", 0, "OK")
                                'Else
                                '    p_oSBOApplication.MessageBox("Send Message Successfully", 0, "OK")
                                'End If

                                'End Save Report To specific JobFile as PDF File


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
                                        'oShpForm.Items.Item("bt_Add").Specific.Caption = "ADD"
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
                        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
                    oPayForm = p_oSBOApplication.Forms.GetForm("VOUCHER", 1)
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

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.BeforeAction = False Then
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
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                            If pVal.ItemUID = "1" Then
                                Try
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
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

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
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
                                    oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)

                                    If p_oSBOApplication.Forms.ActiveForm.TypeEx = "VOUCHER" Then
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            ' If Not AddNewRow(oPayForm, "mx_ChCode") Then Throw New ArgumentException(sErrDesc)
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
                                        End If
                                        If oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                        oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            sql = "Update [@OBT_TB031_VHEADER] set U_APInvNo=" & Convert.ToInt32(strAPInvNo) & ",U_OutPayNo=" & Convert.ToInt32(strOutPayNo) & "" & _
                                            " ,U_FrDocNo=" & oActiveForm.Items.Item("ed_DocNum").Specific.Value & " Where DocEntry = " & oPayForm.Items.Item("ed_DocNum").Specific.Value & ""
                                            oRecordSet.DoQuery(sql)
                                            PreviewPaymentVoucher(oActiveForm, oPayForm)
                                            oPayForm.Close()

                                            oActiveForm.Items.Item("1").Click()
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            End If

                                        ElseIf oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oPayForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        oPayForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                            'SBO_Application.ActivateMenuItem("1291")
                                            oPayForm.Items.Item("ed_VedCode").Enabled = False
                                            oPayForm.Items.Item("ed_VedName").Enabled = False
                                            oPayForm.Items.Item("bt_PayView").Visible = True
                                            ' oPayForm.Close()
                                        End If
                                    End If
                                    oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                End If
                            End If
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.Before_Action = False Then
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
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT And pVal.Before_Action = False Then
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
                        'MSW TO Edit New Ticket

                        'MSW TO Edit New Ticket
                        If pVal.ItemUID = "cb_PayFor" Then
                            If oPayForm.Items.Item("cb_PayFor").Specific.Value.ToString.Trim = "Define New" Then
                                oCombo = oPayForm.Items.Item("cb_PayFor").Specific
                                oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                p_oSBOApplication.ActivateMenuItem("47617")
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
                Case "frmSend"
                        Try
                            oActiveForm = p_oSBOApplication.Forms.Item(pVal.FormUID)

                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST And pVal.BeforeAction = False Then
                                Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                                Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                                If pVal.ItemUID = "ed_ToName" Then
                                    oActiveForm.DataSources.DBDataSources.Item("OHEM").SetValue("lastName", 0, oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString)
                                    oActiveForm.DataSources.DBDataSources.Item("OHEM").SetValue("jobTitle", 0, oDataTable.Columns.Item("jobTitle").Cells.Item(0).Value.ToString)
                                    oActiveForm.DataSources.DBDataSources.Item("OHEM").SetValue("mobile", 0, oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString)
                                    oActiveForm.DataSources.DBDataSources.Item("OHEM").SetValue("email", 0, oDataTable.Columns.Item("email").Cells.Item(0).Value.ToString)
                                End If
                            End If
                            If pVal.ItemUID = "bt_Attach" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then
                                Try
                                    Dim newThread As New Thread(New ThreadStart(AddressOf ThreadMethodEmail))
                                    newThread.SetApartmentState(ApartmentState.STA)

                                    newThread.Start()
                                    newThread.Join()
                                    oActiveForm.Items.Item("ed_Attach").Specific.String = fname
                                Catch ex As Exception
                                End Try
                            End If
                            If pVal.ItemUID = "bt_Send" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = False Then
                                Try
                                    Dim msg As New MailMessage(oActiveForm.Items.Item("ed_From").Specific.String, oActiveForm.Items.Item("ed_Email").Specific.String, oActiveForm.Items.Item("ed_Subject").Specific.String, oActiveForm.Items.Item("ed_Body").Specific.String)
                                    Dim smtpser As New SmtpClient("MIDSERVER", 25)
                                    If fname <> "" Then

                                        Dim attach As Attachment = New Attachment(fname)
                                        msg.Attachments.Add(attach)
                                    End If

                                    smtpser.Send(msg)

                                Catch ex As Exception

                                End Try
                                ' Thread.Sleep(60000)
                                ClearText(oActiveForm, "ed_From", "ed_ToName", "ed_Posit", "ed_Mobile", "ed_Email", "ed_Subject", "ed_Body", "ed_Attach")
                                oActiveForm.Items.Item("ed_From").Specific.Active = True
                                p_oSBOApplication.MessageBox("Send Message Successfully", 0, "OK")
                            End If
                        Catch ex As Exception

                        End Try

                Case "IMPORTSEAFCL"
                    oActiveForm = p_oSBOApplication.Forms.Item(pVal.FormUID)
                    Try
                        oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                        Dim index As Integer = 6
                        Dim folder As String
                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = False Then
                            LoadHolidayMarkUp(oActiveForm)
                        End If
                    End If
                   

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                            If Not RemoveFromAppList(oActiveForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                        End If
                        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                            oActiveForm.Items.Item("bt_CranePO").Enabled = True
                            oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                            oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                            oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                        oActiveForm.Items.Item("bt_ShpInv").Enabled = True
                        oActiveForm.Items.Item("bt_PrtDisp").Enabled = True 'MSW 10-09-2011
                        End If


                        If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And pVal.BeforeAction = True And pVal.InnerEvent = False Then
                            If pVal.ItemUID = "ed_FCharge" And pVal.Before_Action = True Then
                                If oActiveForm.Items.Item("ed_FCharge").Specific.value <> "" Then
                                    Try
                                        If oActiveForm.Items.Item("ed_FCharge").Specific.value < 0 Then
                                            oActiveForm.Items.Item("ed_FCharge").Specific.Value = ""
                                            oActiveForm.Items.Item("ed_FCharge").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oActiveForm.Items.Item("ed_FCharge").Specific.Value = ""
                                        oActiveForm.Items.Item("ed_FCharge").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End Try
                                End If


                            End If
                            If pVal.ItemUID = "ed_Percent" And pVal.Before_Action = True Then
                                If oActiveForm.Items.Item("ed_Percent").Specific.value <> "" Then
                                    Try
                                        If oActiveForm.Items.Item("ed_Percent").Specific.value < 0 Then
                                            oActiveForm.Items.Item("ed_Percent").Specific.Value = ""
                                            oActiveForm.Items.Item("ed_Percent").Specific.Active = True
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oActiveForm.Items.Item("ed_Percent").Specific.Value = ""
                                        oActiveForm.Items.Item("ed_Percent").Specific.Active = True
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End Try
                                End If

                            End If
                            If pVal.ItemUID = "ed_GWt" And pVal.Before_Action = True Then
                                If oActiveForm.Items.Item("ed_GWt").Specific.value <> "" Then
                                    Try
                                        If oActiveForm.Items.Item("ed_GWt").Specific.value < 0 Then
                                            oActiveForm.Items.Item("ed_GWt").Specific.Active = True
                                            oActiveForm.Items.Item("ed_GWt").Specific.Value = ""
                                            p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                            BubbleEvent = False
                                        End If
                                    Catch ex As Exception
                                        oActiveForm.Items.Item("ed_GWt").Specific.Active = True
                                        oActiveForm.Items.Item("ed_GWt").Specific.Value = ""
                                        p_oSBOApplication.SetStatusBarMessage("Please fill Positive numeric", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False

                                    End Try
                                End If

                            End If
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
                            'oActiveForm.Freeze(True)
                            '    If pVal.ItemUID = "bt_Reset" Then
                            '        If oActiveForm.Items.Item("ch_Permit").Specific.Checked = True Then
                            '            index = index + 100
                            '            oActiveForm.Items.Item("fo_Prmt").Visible = True
                            '            'oActiveForm.Items.Item("fo_Prmt").Enabled = True
                            '            folder = "fo_Prmt"
                            '        Else
                            '            oActiveForm.Items.Item("fo_Prmt").Visible = False
                            '            'oActiveForm.Items.Item("fo_Prmt").Enabled = False
                            '        End If
                            '        If oActiveForm.Items.Item("ch_Voucher").Specific.Checked = True Then
                            '            oActiveForm.Items.Item("fo_Vchr").Visible = True
                            '            If folder Is Nothing Then
                            '                folder = "fo_Vchr"
                            '            End If
                            '            oActiveForm.Items.Item("fo_Vchr").Left = index
                            '            index = index + 100
                            '        Else
                            '            oActiveForm.Items.Item("fo_Vchr").Visible = False
                            '        End If
                            '        If oActiveForm.Items.Item("ch_RYard").Specific.Checked = True Then
                            '            oActiveForm.Items.Item("fo_Yard").Visible = True
                            '            If folder Is Nothing Then
                            '                folder = "fo_Yard"
                            '            End If
                            '            oActiveForm.Items.Item("fo_Yard").Left = index
                            '            index = index + 100
                            '        Else
                            '            oActiveForm.Items.Item("fo_Yard").Visible = False
                            '        End If
                            '        If oActiveForm.Items.Item("ch_Conta").Specific.Checked = True Then
                            '            oActiveForm.Items.Item("fo_Cont").Visible = True
                            '            If folder Is Nothing Then
                            '                folder = "fo_Cont"
                            '            End If
                            '            oActiveForm.Items.Item("fo_Cont").Left = index
                            '            index = index + 100
                            '        Else
                            '            oActiveForm.Items.Item("fo_Cont").Visible = False
                            '        End If
                            '        If oActiveForm.Items.Item("ch_Truck").Specific.Checked = True Then
                            '            oActiveForm.Items.Item("fo_Trkng").Visible = True
                            '            If folder Is Nothing Then
                            '                folder = "fo_Trkng"
                            '            End If
                            '            oActiveForm.Items.Item("fo_Trkng").Left = index
                            '            index = index + 100
                            '        Else
                            '            oActiveForm.Items.Item("fo_Trkng").Visible = False
                            '        End If
                            '        If oActiveForm.Items.Item("ch_Disp").Specific.Checked = True Then
                            '            oActiveForm.Items.Item("fo_Dsptch").Visible = True
                            '            If folder Is Nothing Then
                            '                folder = "fo_Dsptch"
                            '            End If
                            '            oActiveForm.Items.Item("fo_Dsptch").Left = index
                            '            index = index + 100
                            '        Else
                            '            oActiveForm.Items.Item("fo_Dsptch").Visible = False
                            '        End If
                            '        If oActiveForm.Items.Item("ch_OCharge").Specific.Checked = vbTrue Then
                            '            oActiveForm.Items.Item("fo_Charge").Visible = True
                            '            If folder Is Nothing Then
                            '                folder = "fo_Charge"
                            '            End If
                            '            oActiveForm.Items.Item("fo_Charge").Left = index
                            '            index = index + 100
                            '        Else
                            '            oActiveForm.Items.Item("fo_Charge").Visible = False
                            '        End If
                            '        If folder Is Nothing Then
                            '            oActiveForm.PaneLevel = 0
                            '            oActiveForm.Items.Item("rt_Outer").Visible = False
                            '        Else
                            '            oActiveForm.Items.Item("rt_Outer").Visible = True
                            '            oActiveForm.Items.Item(folder).Specific.Select()
                            '            If folder = "fo_Prmt" Then
                            '                oActiveForm.PaneLevel = 8
                            '            End If
                            '        End If
                            '    End If
                        '    oActiveForm.Freeze(False)
                 
                        End If



                        If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            EnabledTrucker(oActiveForm, False)
                        End If

                        If Not pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE And pVal.Before_Action = False Then
                            Try
                                oActiveForm.Items.Item("ed_Code").Enabled = False
                                oActiveForm.Items.Item("ed_Name").Enabled = False
                                'oActiveForm.Items.Item("ed_JobNo").Enabled = False 'MSW 30-05-2011
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
                            '    If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            'End If
                            If pVal.ItemUID = "ed_InsDate" And oActiveForm.Items.Item("ed_InsDate").Specific.String <> String.Empty Then
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_InsDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            If pVal.ItemUID = "ed_TkrDate" And oActiveForm.Items.Item("ed_TkrDate").Specific.String <> String.Empty Then
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_TkrDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            'MSW
                        'If pVal.ItemUID = "ed_PosDate" And oActiveForm.Items.Item("ed_PosDate").Specific.String <> String.Empty Then
                        '    If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        'End If
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
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_TkrList"
                                '#1008 17-09-2011
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
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_ConTab"
                                '#1008 17-09-2011
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
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_Voucher"
                                '#1008 17-09-2011
                                    oMatrix = oActiveForm.Items.Item("mx_Voucher").Specific
                                    If oMatrix.RowCount > 1 Then
                                        oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                        ' oActiveForm.Items.Item("bt_PrPDF").Enabled = True 'MSW To Add 18-03-2011
                                        ' oActiveForm.Items.Item("bt_Email").Enabled = True 'MSW To Add 18-03-2011
                                    ElseIf oMatrix.RowCount = 1 Then
                                        If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                            End If

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
                                oActiveForm.Freeze(True)
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
                                    'oActiveForm.Items.Item("ed_InsDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    oActiveForm.DataSources.UserDataSources.Item("TKRDATE").ValueEx = Today.Date.ToString("yyyyMMdd")
                                    EnabledTruckerForExternal(oActiveForm, True) 'MSW 14-09-2011 Truck PO
                                    oActiveForm.Items.Item("op_Inter").Specific.Selected = True 'MSW 01-04-2011
                                End If
                                    oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
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
                                oActiveForm.Freeze(False)
                                Case ("fo_Charge")
                                    oActiveForm.PaneLevel = 19
                                oActiveForm.Items.Item("fo_ChView").Specific.Select()
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_Charge"
                                '#1008 17-09-2011
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
                                oActiveForm.Items.Item("bt_PrtDisp").Enabled = False
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
                                        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRecordSet.DoQuery("select count(*)  from [@OBT_TB024_HCHARGES] where U_ChSeqNo=  '" + oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + "'")
                                        If oRecordSet.RecordCount > 0 And IsAmend = False Then
                                            oActiveForm.Items.Item("ed_CSeqNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                        End If
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
                                            If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oActiveForm.Items.Item("bt_ShpInv").Enabled = True
                                            End If

                                        Else
                                            '    ImportSeaLCLForm.Items.Item("bt_AmdVoc").Enabled = False
                                        End If
                                    End If
                                Case "fo_ShView"
                                    oActiveForm.PaneLevel = 25


                                Case "fo_OpBunk"
                                    oActiveForm.PaneLevel = 29
                                oActiveForm.Items.Item("fo_BkView").Specific.Select()
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_Bunk"
                                '#1008 17-09-2011
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                                    End If

                                Case "fo_ArmEs"
                                    oActiveForm.PaneLevel = 30
                                oActiveForm.Items.Item("fo_ArView").Specific.Select()
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_Armed"
                                '#1008 17-09-2011
                                    oMatrix = oActiveForm.Items.Item("mx_Armed").Specific
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                                    End If

                                Case "fo_Crane"
                                    oActiveForm.PaneLevel = 31
                                oActiveForm.Items.Item("fo_CrView").Specific.Select()
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_Crane"
                                '#1008 17-09-2011
                                    oMatrix = oActiveForm.Items.Item("mx_Crane").Specific
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oActiveForm.Items.Item("bt_CranePO").Enabled = True
                                    End If

                                Case "fo_Fork"
                                    oActiveForm.PaneLevel = 32
                                oActiveForm.Items.Item("fo_FkView").Specific.Select()
                                '#1008 17-09-2011
                                oActiveForm.Settings.Enabled = True
                                oActiveForm.Settings.EnableRowFormat = True
                                oActiveForm.Settings.MatrixUID = "mx_Fork"
                                '#1008 17-09-2011
                                    oMatrix = oActiveForm.Items.Item("mx_Fork").Specific
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Or oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                                    End If

                                Case "fo_FkView"
                                    oActiveForm.PaneLevel = 32
                                Case "fo_ArView"
                                    oActiveForm.PaneLevel = 30
                                Case "fo_CrView"
                                    oActiveForm.PaneLevel = 31
                                Case "fo_BkView"
                                    oActiveForm.PaneLevel = 29

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
                                        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB014_PLICINFO").RemoveRecord(0)
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
                            oMatrix.Columns.Item("colCSeqNo1").Cells.Item(oMatrix.RowCount).Specific.Value = oMatrix.RowCount.ToString
                            End If

                            If pVal.ItemUID = "bt_DelCon" Then
                                oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
                                Dim lRow As Long
                                lRow = oMatrix.GetNextSelectedRow
                                If lRow > -1 Then
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        oActiveForm.DataSources.DBDataSources.Item("@OBT_TB012_PCONTAINE").RemoveRecord(0)
                                        Dim oUserTable As SAPbobsCOM.UserTable
                                        oUserTable = p_oDICompany.UserTables.Item("OBT_TB012_PCONTAINE")
                                        oUserTable.GetByKey(lRow)
                                        oUserTable.Remove()
                                        oUserTable = Nothing
                                    End If
                                    oMatrix.DeleteRow(lRow)
                                SetMatrixSeqNo(oMatrix, "colCSeqNo1")
                                    If lRow = 1 And oMatrix.RowCount = 0 Then
                                        oMatrix.FlushToDataSource()
                                        oMatrix.AddRow()
                                    oMatrix.Columns.Item("colCSeqNo1").Cells.Item(1).Specific.Value = 1
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
                                'If DateTime(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                Try
                                    If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                    oActiveForm.Items.Item("ed_ADate").Specific.Value = oActiveForm.Items.Item("ed_ETADate").Specific.Value
                                    oActiveForm.Items.Item("ed_ADay").Specific.Value = oActiveForm.Items.Item("ed_ETADay").Specific.Value
                                    oActiveForm.Items.Item("ed_ATime").Specific.Value = oActiveForm.Items.Item("ed_ETAHr").Specific.Value
                                Catch ex As Exception
                                    'oActiveForm.Items.Item("ed_ADate").Specific.Value = oActiveForm.Items.Item("ed_ETADate").Specific.Value
                                    'oActiveForm.Items.Item("ed_ADay").Specific.Value = oActiveForm.Items.Item("ed_ETADay").Specific.Value
                                    'oActiveForm.Items.Item("ed_ATime").Specific.Value = oActiveForm.Items.Item("ed_ETAHr").Specific.Value

                                End Try
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                oActiveForm.Items.Item("ed_ADate").Specific.Value = oActiveForm.Items.Item("ed_ETADate").Specific.Value
                                If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If

                            If pVal.ItemUID = "ed_ConLast" And oActiveForm.Items.Item("ed_ConLast").Specific.String <> String.Empty Then
                                If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTim").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTim").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If

                            If pVal.ItemUID = "ed_CunDate" And oActiveForm.Items.Item("ed_CunDate").Specific.String <> String.Empty Then
                                If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_CunDate").Specific, oActiveForm.Items.Item("ed_CunDay").Specific, oActiveForm.Items.Item("ed_CunTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CunDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CunDay").Specific, oActiveForm.Items.Item("ed_CunTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            If pVal.ItemUID = "ed_ConDate" And oActiveForm.Items.Item("ed_ConDate").Specific.String <> String.Empty Then
                                If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_ConDate").Specific, oActiveForm.Items.Item("ed_ConDay").Specific, oActiveForm.Items.Item("ed_ConTime").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConDay").Specific, oActiveForm.Items.Item("ed_ConTime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If

                            If pVal.ItemUID = "ed_JbDate" And oActiveForm.Items.Item("ed_JbDate").Specific.String <> String.Empty Then
                                If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                            'If pVal.ItemUID = "ed_DspDate" And oActiveForm.Items.Item("ed_DspDate").Specific.String <> String.Empty Then
                            '    If DateTime(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            '    If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            'End If
                            If pVal.ItemUID = "ed_DspCDte" And oActiveForm.Items.Item("ed_DspCDte").Specific.String <> String.Empty Then
                                If Not DateTime(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) Then Throw New ArgumentException(sErrDesc)
                                If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If

                            If pVal.ItemUID = "bt_Payment" Then
                                p_oSBOApplication.ActivateMenuItem("2818")
                            End If

                        If pVal.ItemUID = "op_Inter" Then
                            oActiveForm.Freeze(True)
                            'MSW 14-09-2011 Truck PO
                            If oActiveForm.Items.Item("ed_PONo").Specific.Value <> "" Then
                                If Not CancelTruckingPurchaseOrder(Convert.ToInt32(oActiveForm.Items.Item("ed_PONo").Specific.Value)) Then Throw New ArgumentException(sErrDesc)
                            End If
                            'MSW 14-09-2011 Truck PO
                            oActiveForm.Items.Item("bt_GenPO").Enabled = False
                            oActiveForm.Items.Item("ed_Trucker").Specific.Value = ""
                            oActiveForm.Items.Item("ed_TkrTel").Specific.Value = ""
                            oActiveForm.Items.Item("ed_Fax").Specific.Value = ""
                            oActiveForm.Items.Item("ed_Email").Specific.Value = ""
                            oActiveForm.Items.Item("ed_PODocNo").Specific.Value = ""
                            'MSW To Edit New Ticket
                            oActiveForm.Items.Item("ed_VehicNo").Specific.Value = ""
                            oActiveForm.Items.Item("ed_Attent").Specific.Value = ""
                            oActiveForm.Items.Item("ed_PONo").Specific.Value = ""
                            oActiveForm.Items.Item("ed_EUC").Specific.Value = ""
                            EnabledTruckerForExternal(oActiveForm, True)
                            'End MSW To Edit New Ticket
                            If AddChooseFromListByOption(oActiveForm, True, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            oActiveForm.Freeze(False)
                        ElseIf pVal.ItemUID = "op_Exter" Then
                            oActiveForm.Freeze(True)
                            If oActiveForm.Items.Item("ed_PONo").Specific.Value = "" Then 'MSW 14-09-2011 Truck PO
                                oActiveForm.Items.Item("ed_Trucker").Specific.Value = ""
                                oActiveForm.Items.Item("ed_TkrTel").Specific.Value = ""
                                oActiveForm.Items.Item("ed_Fax").Specific.Value = ""
                                oActiveForm.Items.Item("ed_Email").Specific.Value = ""
                                'MSW To Edit New Ticket
                                oActiveForm.Items.Item("ed_VehicNo").Specific.Value = ""
                                oActiveForm.Items.Item("ed_Attent").Specific.Value = ""
                                oActiveForm.Items.Item("ed_PONo").Specific.Value = ""
                                oActiveForm.Items.Item("ed_EUC").Specific.Value = ""
                                EnabledTruckerForExternal(oActiveForm, False)
                                'End MSW To Edit New Ticket
                                oActiveForm.Items.Item("bt_GenPO").Enabled = True
                                If AddChooseFromListByOption(oActiveForm, False, "ed_Trucker", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            End If
                           
                            oActiveForm.Freeze(False)
                        End If

                            If pVal.ItemUID = "op_DspIntr" Then
                                oCombo = oActiveForm.Items.Item("cb_Dspchr").Specific
                                If Not ClearComboData(oActiveForm, "cb_Dspchr", "@OBT_TB007_DISPATCH", "U_Dispatch") Then Throw New ArgumentException(sErrDesc)
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

                                If Not ClearComboData(oActiveForm, "cb_Dspchr", "@OBT_TB007_DISPATCH", "U_Dispatch") Then Throw New ArgumentException(sErrDesc)
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

                            'Purchase Order and Goods Receipt POP UP

                            ' ==================================== Creating Custom Purchase Order ==============================
                            If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close"
                                End If
                                LoadTruckingPO(oActiveForm, "TkrListPurchaseOrder.srf")
                            Else
                                p_oSBOApplication.SetStatusBarMessage("No Job Number to create Purchase Order.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                Exit Function
                            End If

                            ' ==================================== Creating Custom Purchase Order ==============================
                            'p_oSBOApplication.Menus.Item("2305").Activate()
                            'p_oSBOApplication.ActivateMenuItem("6913") 'MSW 04-04-2011
                            'p_oSBOApplication.Menus.Item("6913").Activate()

                            'Dim UDFAttachForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.GetForm("-142", 1)
                            'UDFAttachForm.Items.Item("U_JobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
                            'UDFAttachForm.Items.Item("U_InsDate").Specific.Value = oActiveForm.Items.Item("ed_InsDate").Specific.Value
                        End If

                        If pVal.ItemUID = "1" Then
                            If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                If pVal.ActionSuccess = True Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    oActiveForm.Items.Item("ed_Code").Enabled = False 'MSW
                                    oActiveForm.Items.Item("ed_Name").Enabled = False
                                    p_oSBOApplication.ActivateMenuItem("1291")
                                    oActiveForm.Items.Item("bt_ForkPO").Enabled = True
                                    oActiveForm.Items.Item("bt_CranePO").Enabled = True
                                    oActiveForm.Items.Item("bt_ArmePO").Enabled = True
                                    oActiveForm.Items.Item("bt_BunkPO").Enabled = True
                                    oActiveForm.Items.Item("bt_AmdVoc").Enabled = True
                                    oActiveForm.Items.Item("bt_ShpInv").Enabled = True
                                    oActiveForm.Items.Item("bt_PrtDisp").Enabled = True 'MSW 10-09-2011
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011

                                End If

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
                                    oActiveForm.Items.Item("bt_PrntIns").Enabled = False 'MSW to edit New Ticket 07-09-2011
                                    End If
                                ClearText(oActiveForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk", "ee_ColFrm") 'MSW New Ticket 07-09-2011
                                oActiveForm.Items.Item("fo_TkrView").Specific.Select()
                                oActiveForm.Items.Item("1").Click() 'MSW to edit New Ticket 07-09-2011
                                    oActiveForm.Items.Item("bt_AmdIns").Enabled = True
                                End If
                            End If

                            If pVal.ItemUID = "bt_DelIns" Then
                                oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                                modTrucking.DeleteByIndex(oActiveForm, oMatrix, "@OBT_TB006_TRUCKING")
                            ClearText(oActiveForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk", "ee_ColFrm") 'MSW New Ticket 07-09-2011
                                oActiveForm.Items.Item("fo_TkrView").Specific.Select()
                            End If

                        'KM to edit
                        If pVal.ItemUID = "bt_PrntIns" Then
                            PreviewInsDoc(oActiveForm)
                        End If


                        If pVal.ItemUID = "bt_A6Label" Then
                            PreviewA6Label(oActiveForm)
                        End If

                        If pVal.ItemUID = "bt_AmdIns" Then
                            oMatrix = oActiveForm.Items.Item("mx_TkrList").Specific
                            If (oMatrix.GetNextSelectedRow < 0) Then
                                p_oSBOApplication.MessageBox("Please Select One Row To Edit Trucking Instruction", 1, "OK")
                                Exit Function
                            Else
                                oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Update Trucking Instruction"
                                modTrucking.GetDataFromMatrixByIndex(oActiveForm, oMatrix, oMatrix.GetNextSelectedRow)
                                modTrucking.SetDataToEditTabByIndex(oActiveForm)
                                oActiveForm.Items.Item("fo_TkrEdit").Specific.Select()
                                'MSW 14-09-2011 Truck PO
                                If oActiveForm.Items.Item("op_Exter").Specific.Selected = True Then
                                    EnabledTruckerForExternal(oActiveForm, False)
                                ElseIf oActiveForm.Items.Item("op_Inter").Specific.Selected = True Then
                                    EnabledTruckerForExternal(oActiveForm, True)
                                End If
                                oActiveForm.Items.Item("bt_DelIns").Enabled = True 'MSW
                                oActiveForm.Items.Item("bt_PrntIns").Enabled = True 'MSW to edit New Ticket 07-09-2011
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
                            ClearText(oActiveForm, "ed_InsDoc", "ed_PODocNo", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk", "ee_ColFrm") 'MSW New Ticket 07-09-2011
                                oActiveForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                            oActiveForm.Items.Item("bt_DelIns").Enabled = False 'MSW
                            oActiveForm.Items.Item("bt_PrntIns").Enabled = False 'MSW to edit New Ticket 07-09-2011
                            End If
                        If pVal.ItemUID = "fo_DiscView" Then
                            oMatrix = oActiveForm.Items.Item("mx_DispTab").Specific
                            If oMatrix.RowCount > 1 Then
                                oActiveForm.Items.Item("bt_AmdDisp").Enabled = True
                            ElseIf oMatrix.RowCount = 1 Then
                                If oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    oActiveForm.Items.Item("bt_AmdDisp").Enabled = True
                                Else
                                    oActiveForm.Items.Item("bt_AmdDisp").Enabled = False
                                End If
                            End If
                            'ClearText(oActiveForm, "ed_InsDoc", "ed_PONo", "ed_InsDate", "ed_Trucker", "ed_VehicNo", "ed_EUC", "ed_Attent", "ed_TkrTel", "ed_Fax", "ed_Email", "ed_TkrDate", "ed_TkrTime", "ee_TkrIns", "ee_InsRmsk", "ee_Rmsk") 'MSW New Ticket 07-09-2011
                            oActiveForm.Items.Item("bt_AddDisp").Specific.Caption = "Add Dispatch"
                            oActiveForm.Items.Item("bt_DelDisp").Enabled = False 'MSW
                            oActiveForm.Items.Item("bt_PrtDisp").Enabled = False 'MSW to edit New Ticket 07-09-2011
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
                                'ClearText(oActiveForm, "ed_VedName", "ed_VedCode", "ed_PayRate", "ed_Cheque", "ed_VocNo", "ed_PosDate", "ed_VRemark", "ed_VPrep", "ed_SubTot", "ed_GSTAmt", "ed_Total")

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

                            If pVal.ItemUID = "bt_AmdVoc" Then
                                'POP UP Payment Voucher
                                If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Items.Item("ed_YAddr").Specific.Value = "" And _
                                    Not oActiveForm.Items.Item("ed_Yard").Specific.Value = "" And _
                                     Not oActiveForm.Items.Item("ed_Vessel").Specific.Value = "" And _
                                      Not oActiveForm.Items.Item("ed_Code").Specific.Value = "" And Not oActiveForm.Items.Item("cb_PCode").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    LoadShippingInvoice(oActiveForm)
                                Else
                                    If oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("No Job Number to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    ElseIf oActiveForm.Items.Item("ed_Code").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("No Customer Code to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    'ElseIf oActiveForm.Items.Item("ed_ShpAgt").Specific.Value = "" Then
                                    '    p_oSBOApplication.SetStatusBarMessage("No Shipping Agent Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    ElseIf oActiveForm.Items.Item("cb_PCode").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("No Port of Loading to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    ElseIf oActiveForm.Items.Item("ed_Vessel").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("No Vessel/Voy Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    ElseIf oActiveForm.Items.Item("ed_Yard").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("No Return Yard Name to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    ElseIf oActiveForm.Items.Item("ed_YAddr").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("Please go to Warehouse Tab and  fill Return Yard Address to create Shipping Invoice.", SAPbouiCOM.BoMessageTime.bmt_Short)

                                    End If
                                End If
                            End If
                            'Purchase Order and Goods Receipt POP UP

                            If pVal.ItemUID = "bt_ForkPO" Then 'sw
                                If Not oActiveForm.Items.Item("ed_JobNo").Specific.Value = "" And Not oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    'If oActiveForm.Items.Item("ed_FVendor").Specific.Value = "TO001" Then
                                    '    If Not CreatePOPDF(oActiveForm, "mx_Fork", "TO001", "F00001") Then Throw New ArgumentException(sErrDesc)
                                    'Else
                                    '    LoadAndCreateCPO(oActiveForm, "ForkPurchaseOrder.srf")
                                    'End If
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
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    'If oActiveForm.Items.Item("ed_CVendor").Specific.Value = "TO001" Then
                                    '    If Not CreatePOPDF(oActiveForm, "mx_Crane", "TO001", "CR0001") Then Throw New ArgumentException(sErrDesc)
                                    'Else
                                    '    LoadAndCreateCPO(oActiveForm, "CranePurchaseOrder.srf")
                                    'End If
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
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If
                                    ''If Not CreatePOPDF(oActiveForm, "mx_Bunk", "TO001", "B00001") Then Throw New ArgumentException(sErrDesc)
                                    'If oActiveForm.Items.Item("ed_V").Specific.Value = "TO001" Then
                                    '    If Not CreatePOPDF(oActiveForm, "mx_Bunk") Then Throw New ArgumentException(sErrDesc)
                                    'Else
                                    '    LoadAndCreateCPO(oActiveForm, "BunkPurchaseOrder.srf")
                                    'End If
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
                                Dim dtPayment As New DataTable
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

                                Dim attach As Attachment = New Attachment(adjunto)
                                msg.Attachments.Add(attach)
                                smtpser.Send(msg)
                                rptDocument.Close()
                                'System.Diagnostics.Process.Start("D:\VoucherItem.pdf")
                            End If
                            If pVal.ItemUID = "bt_PrPDF" Then
                                Dim rptDocument As ReportDocument = New ReportDocument()
                                Dim CrxReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument
                                Dim dtPayment As New DataTable
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
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery("select count(*)  from [@OBT_TB020_HCONTAINE] where U_ConSeqNo=  '" + oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + "'")
                                    If oRecordSet.RecordCount > 0 And IsAmend = False Then
                                        oActiveForm.Items.Item("ed_ConNo").Specific.Value = oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value + 1
                                    End If
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
                            IsAmend = False
                        End If
                        If pVal.ItemUID = "fo_ChView" Then
                            oMatrix = oActiveForm.Items.Item("mx_Charge").Specific
                            If oMatrix.RowCount > 1 Then
                                oActiveForm.Items.Item("bt_AmdCh").Enabled = True
                            ElseIf oMatrix.RowCount = 1 Then
                                If oMatrix.Columns.Item("V_-1").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    oActiveForm.Items.Item("bt_AmdCh").Enabled = True
                                Else
                                    oActiveForm.Items.Item("bt_AmdCh").Enabled = False
                                End If
                            End If
                            ClearText(oActiveForm, "ed_CSeqNo", "ed_ChCode", "ed_Remarks")
                        
                            oActiveForm.Items.Item("bt_AddCh").Specific.Caption = "Add Charges"
                            oActiveForm.Items.Item("bt_DelCh").Enabled = False 'MSW
                            IsAmend = False
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
                            IsAmend = True
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
                                oActiveForm.Items.Item("bt_PrtDisp").Enabled = False 'MSW to edit New Ticket 07-09-2011
                                End If

                                oActiveForm.Items.Item("op_DspIntr").Specific.Selected = True
                                ClearText(oActiveForm, "ee_Instru", "ed_DspCDte", "ed_DspCDay", "ed_DspCHr")
                                oActiveForm.Items.Item("bt_AmdDisp").Enabled = True
                            oActiveForm.Items.Item("fo_DisView").Specific.Select()
                            oActiveForm.Items.Item("1").Click() 'MSW to edit New Ticket 07-09-2011

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
                                oActiveForm.Items.Item("bt_PrtDisp").Enabled = True 'MSW to edit New Ticket 07-09-2011
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
                                IsAmend = True
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
                            oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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
                                    oActiveForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
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

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.ItemUID = "mx_Cont" And pVal.ColUID = "colCSize1" And pVal.Before_Action = True And pVal.Row <> 0 Then
                        oMatrix = oActiveForm.Items.Item("mx_Cont").Specific
                        Dim oColCombo As SAPbouiCOM.Column
                        Dim omatCombo As SAPbouiCOM.ComboBox
                        oColCombo = oMatrix.Columns.Item("colCSize1")
                        omatCombo = oColCombo.Cells.Item(pVal.Row).Specific
                        If Not String.IsNullOrEmpty(oMatrix.Columns.Item("colCType1").Cells.Item(pVal.Row).Specific.Value) Then
                            Dim type As String = oMatrix.Columns.Item("colCType1").Cells.Item(pVal.Row).Specific.Selected.Value
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
                                ''If pVal.ItemUID = "ed_CVendor" Then
                                ''    'oActiveForm.DataSources.UserDataSources.Item("CVendor").Value = oDataTable.GetValue(0, 0).ToString
                                ''    oActiveForm.Items.Item("ed_CVendor").Specific.Value = oDataTable.GetValue(0, 0).ToString
                                ''End If
                                ''If pVal.ItemUID = "ed_FVendor" Then
                                ''    ' oActiveForm.DataSources.UserDataSources.Item("FVendor").Value = oDataTable.GetValue(0, 0).ToString
                                ''    oActiveForm.Items.Item("ed_FVendor").Specific.Value = oDataTable.GetValue(0, 0).ToString
                                ''End If
                                If pVal.ItemUID = "ed_Code" Or pVal.ItemUID = "ed_Name" Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Code", 0, oDataTable.GetValue(0, 0).ToString)
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Name", 0, oDataTable.GetValue(1, 0).ToString)
                                    oActiveForm.DataSources.UserDataSources.Item("TKRTO").ValueEx = oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString
                                    If (oActiveForm.Items.Item("ed_ETADate").Specific.Value = "") Then
                                    'oActiveForm.Items.Item("ed_ETADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    'oActiveForm.Items.Item("ed_ETADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                    'oActiveForm.Items.Item("ed_ETAHr").Specific.Value = Now.ToString("HH:mm")
                                    'If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    'oActiveForm.Items.Item("ed_ADate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                    'oActiveForm.Items.Item("ed_ADay").Specific.Value = Today.DayOfWeek.ToString.Substring(0, 3)
                                    'oActiveForm.Items.Item("ed_ATime").Specific.Value = Now.ToString("HH:mm")
                                    'If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                    End If
                                    EnabledHeaderControls(oActiveForm, True)
                                    If oActiveForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        oActiveForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "cflBP3"
                                        oActiveForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                                        oActiveForm.Items.Item("ed_Yard").Specific.ChooseFromListUID = "YARD"
                                        oActiveForm.Items.Item("ed_Yard").Specific.ChooseFromListAlias = "Code"
                                        oActiveForm.Items.Item("ed_Vessel").Specific.ChooseFromListUID = "DSVES01"
                                        oActiveForm.Items.Item("ed_Vessel").Specific.ChooseFromListAlias = "Name"
                                    ' oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListUID = "DSVES02"
                                    'oActiveForm.Items.Item("ed_Voy").Specific.ChooseFromListAlias = "U_Voyage"
                                    End If
                                ' oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_IUEN", 0, oDataTable.GetValue(0, 0).ToString)
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("SELECT CardName,VatIdUnCmp FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                If oRecordSet.RecordCount > 0 Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_IUEN", 0, oRecordSet.Fields.Item("VatIdUnCmp").Value)
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_IComName", 0, oRecordSet.Fields.Item("CardName").Value)
                                End If
                                End If
                                If pVal.ItemUID = "ed_ShpAgt" Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_ShpAgt", 0, oDataTable.GetValue(1, 0).ToString)
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_VCode", 0, oDataTable.GetValue(0, 0).ToString)
                                'oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_UEN", 0, oDataTable.GetValue(0, 0).ToString)
                                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRecordSet.DoQuery("SELECT CardName,VatIdUnCmp FROM OCRD WHERE CmpPrivate = 'C' And CardName = '" & oDataTable.GetValue(1, 0).ToString & "'")
                                If oRecordSet.RecordCount > 0 Then
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_UEN", 0, oRecordSet.Fields.Item("VatIdUnCmp").Value)
                                    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB005_PERMIT").SetValue("U_ComName", 0, oRecordSet.Fields.Item("CardName").Value)
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

                            If pVal.ItemUID = "ed_Vessel" Then
                                oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Vessel", 0, oDataTable.Columns.Item("Name").Cells.Item(0).Value.ToString)
                                'oActiveForm.DataSources.DBDataSources.Item("@OBT_TB002_IMPSEALCL").SetValue("U_Voyage", 0, oDataTable.Columns.Item("U_Voyage").Cells.Item(0).Value.ToString)
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
                                'If pVal.ItemUID = "ed_VedName" Then
                                '    ObjDBDataSource = oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC") 'MSW To Add 18-3-2011
                                '    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_BPName", ObjDBDataSource.Offset, oDataTable.GetValue(1, 0).ToString)  'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*
                                '    oActiveForm.DataSources.DBDataSources.Item("@OBT_TB028_VOUC").SetValue("U_PayToAdd", ObjDBDataSource.Offset, oDataTable.Columns.Item("Address").Cells.Item(0).Value.ToString() & "," & vbNewLine & oDataTable.Columns.Item("Country").Cells.Item(0).Value.ToString() _
                                '                                                                           & "-" & oDataTable.Columns.Item("ZipCode").Cells.Item(0).Value.ToString()) 'MSW To Add 18-3-2011  *ObjDBDataSource.Offset*

                                '    vendorCode = oDataTable.GetValue(0, 0).ToString 'MSW To Add Draft
                                'End If
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

                Case "426"
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction Then
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
                Case Else



            End Select
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            DoImportSeaFCLItemEvent = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            ShowErr(sErrDesc)
            MessageBox.Show(exc.Message)
            DoImportSeaFCLItemEvent = RTN_ERROR
        End Try
    End Function


    Public Function DoImportSeaFCLRightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        DoImportSeaFCLRightClickEvent = RTN_ERROR
        Dim sFuncName As String = "SBO_Application_RightClickEvent()"
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim ImportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If eventInfo.FormUID = "IMPORTSEAFCL" Then
                oMenuItem = p_oSBOApplication.Menus.Item("1280")
                oMenus = oMenuItem.SubMenus
                ImportSeaFCLForm = p_oSBOApplication.Forms.GetForm(eventInfo.FormUID, 1)
                ImportSeaFCLForm.EnableMenu("772", True)
                ImportSeaFCLForm.EnableMenu("773", True)
                ImportSeaFCLForm.EnableMenu("775", True)
                If eventInfo.ItemUID = "mx_Fumi" Or eventInfo.ItemUID = "mx_Crate" Or eventInfo.ItemUID = "mx_Fork" Or eventInfo.ItemUID = "mx_Armed" Or eventInfo.ItemUID = "mx_Crane" Or eventInfo.ItemUID = "mx_Bunk" Or eventInfo.ItemUID = "mx_TkrList" Then 'Truck PO
                    If eventInfo.BeforeAction = True Then
                        ImportSeaFCLForm.EnableMenu("772", False)
                        ImportSeaFCLForm.EnableMenu("773", False)
                        ImportSeaFCLForm.EnableMenu("775", False)
                        oMatrix = ImportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
                        If eventInfo.ItemUID = "mx_TkrList" Then
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
                            If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                                oMatrix = ImportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
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
                    If eventInfo.BeforeAction = True Then
                        ImportSeaFCLForm.EnableMenu("772", False)
                        ImportSeaFCLForm.EnableMenu("773", False)
                        ImportSeaFCLForm.EnableMenu("775", False)
                        oMatrix = ImportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
                        If eventInfo.Row > 0 And oMatrix.RowCount <> 0 Then
                            If oMatrix.RowCount >= 1 And oMatrix.Columns.Item("colDocNum").Cells.Item(eventInfo.Row).Specific.Value <> "" Then
                                If Not oMenus.Exists("EditShp") Then
                                    oMenus.Add("EditShp", "Edit Shipping Invoice", SAPbouiCOM.BoMenuType.mt_STRING, 8)
                                    RPOmatrixname = eventInfo.ItemUID
                                    RPOsrfname = "ShipInvoice.srf"
                                End If
                                currentRow = eventInfo.Row
                                p_oSBOApplication.Forms.ActiveForm.EnableMenu("1293", True) 'MSW
                            End If
                        End If
                    Else
                        If oMenus.Exists("EditShp") Then
                            p_oSBOApplication.Menus.RemoveEx("EditShp")
                            p_oSBOApplication.Forms.ActiveForm.EnableMenu("1293", False) 'MSW
                        End If
                    End If
                End If
                If eventInfo.ItemUID = "mx_Voucher" Then
                    If eventInfo.BeforeAction = True Then
                        ImportSeaFCLForm.EnableMenu("772", False)
                        ImportSeaFCLForm.EnableMenu("773", False)
                        ImportSeaFCLForm.EnableMenu("775", False)
                        oMatrix = ImportSeaFCLForm.Items.Item(eventInfo.ItemUID).Specific
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

                If eventInfo.ItemUID = "mx_ConTab" Or eventInfo.ItemUID = "mx_Charge" Or eventInfo.ItemUID = "mx_ShpInv" Or eventInfo.ItemUID = "mx_DispTab" Then
                    ImportSeaFCLForm.EnableMenu("772", False)
                    ImportSeaFCLForm.EnableMenu("773", False)
                    ImportSeaFCLForm.EnableMenu("775", False)
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            DoImportSeaFCLRightClickEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoImportSeaFCLRightClickEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
        Finally
            GC.Collect()            'Forces garbage collection of all generations.
        End Try
    End Function

#Region "MSW Voucher POP Up"
    Private Function AddNewRow(ByRef oActiveForm As SAPbouiCOM.Form, ByVal MatrixUID As String) As Boolean
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

#Region "Validate Job Number"
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
#End Region

#Region "SHIPPING INVOICE POP UP"
    Private Sub LoadShippingInvoice(ByRef oActiveForm As SAPbouiCOM.Form)
        Dim oShpForm As SAPbouiCOM.Form
        Dim sErrDesc As String = String.Empty
        Dim oMatrix As SAPbouiCOM.Matrix
        LoadFromXML(p_oSBOApplication, "ShipInvoice.srf")
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
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

        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        ' oShpForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()

        If HolidayMarkUp(oShpForm, oShpForm.Items.Item("ed_ShDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        ' oShpForm.DataSources.DBDataSources.Item("@OBT_TB017_EVHEADER").SetValue("U_DocNo", 0, oActiveForm.Items.Item("ed_JobNo").Specific.Value)


        'oShpForm.Freeze(True)
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

    Public Sub AddUpdateShippingInv(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)

        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)

        'rowIndex = pMatrix.GetNextSelectedRow

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

    Public Sub SetShipInvDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
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

    Public Sub AddUpdateShippingMatrix(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        Dim oActiveForm As SAPbouiCOM.Form
        oActiveForm = p_oSBOApplication.Forms.GetForm("IMPORTSEAFCL", 1)
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
    Public Sub LoadAndCreateCPO(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)
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

    Public Sub LoadAndCreateCGR(ByRef ParentForm As SAPbouiCOM.Form, ByRef srfName As String)
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

            'If AddUserDataSrc(CGRForm, "Email", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            'CGRForm.Items.Item("39").Specific.DataBind.SetBound(True, "", "Email")

            If Not AddNewRowGR(CGRForm, "mx_Item") Then Throw New ArgumentException(sErrDesc)
            CGRForm.DataSources.DBDataSources.Item("@OBT_TB12_FFCGR").SetValue("U_EXPNUM", 0, ParentForm.Items.Item("ed_DocNum").Specific.Value)
            'MSW
            CGRForm.Items.Item("ed_Code").Enabled = False
            CGRForm.Items.Item("ed_Name").Enabled = False
            CGRForm.Items.Item("ed_VRef").Specific.Active = True
            CGRForm.Freeze(False)
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
                    Dim tsql As String = "SELECT U_SInA,U_TPlace,U_TDate,U_TDay,U_TTime,U_PORMKS,U_POIRMKS,U_ColFrm,U_TkrIns,U_TkrTo,U_CNo,U_Dest,U_LocWork,U_POITPD FROM [@OBT_TB08_FFCPO] WHERE DocEntry = " + FormatString(oMatrix.Columns.Item(SourceColName2).Cells.Item(ActiveRow).Specific.Value)
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

    Private Function PopulatePurchaseHeader(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean, ByRef oPOForm As SAPbouiCOM.Form) As Boolean
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

    Public Sub RowFunction(ByRef oForm As SAPbouiCOM.Form, ByVal mxUID As String, ByVal FieldAlias As String, ByVal TableName As String)
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
        '   Function    :   CreatePDF()
        '   Purpose     :   This function provide to create pdf document  when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form, ByVal matrixName As String
        '   return      :   No          
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
            '  ConvertFunction(oActiveForm, pdffilepath)
        End If

        rptDocument.Close()
    End Sub
    

    

    Private Function CreatePurchaseOrderPDF(ByRef oActiveForm As SAPbouiCOM.Form, ByVal CardCode As String, ByVal iCode As String) As Boolean
        ' **********************************************************************************
        '   Function    :   CreatePurchaseOrderPDF()
        '   Purpose     :   This function provide to create pdf document for purchase order  when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef oActiveForm As SAPbouiCOM.Form, ByVal CardCode As String, ByVal iCode As String
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************
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
    Private Function PopulatePurchaseHeaderPDF(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean, ByRef oPOForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   PopulatePurchaseHeaderPDF
        '   Purpose     :   This function will be providing to create  purchase
        '                   order PDF Form.
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal pStrSQL As String, ByVal tblName As String, ByVal ProcessedState As Boolean
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
                            If oPOForm.Title = "Goods Receipt PO" Then
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
                            If oPOForm.Title = "Goods Receipt PO" Then
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
            PopulatePurchaseHeaderPDF = True
        Catch ex As Exception
            PopulatePurchaseHeaderPDF = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    'Private Function CreatePOPDF(ByVal oActiveForm As SAPbouiCOM.Form, ByVal matrixName As String, ByVal cardcode As String, ByVal iCode As String) As Boolean

    ' **********************************************************************************
    '   Function    :   CreatePOPDF
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
    '    Dim sErrDesc As String = ""
    '    Dim i As Integer = 0
    '    Dim dblPrice As Double = 0.0
    '    Dim itemCode As String = String.Empty
    '    Dim itemDesc As String = String.Empty
    '    Dim oMatrix As SAPbouiCOM.Matrix
    '    Dim originpdf As String = String.Empty
    '    Dim docEntry As String = GetNewKey("FCPO", oRecordSet)
    '    Dim oGeneralService As SAPbobsCOM.GeneralService
    '    Dim oGeneralData As SAPbobsCOM.GeneralData
    '    Dim oChild As SAPbobsCOM.GeneralData
    '    Dim oChildren As SAPbobsCOM.GeneralDataCollection


    '    CreatePOPDF = False
    '    Try
    '        oMatrix = oActiveForm.Items.Item(matrixName).Specific

    '        If Not CreatePurchaseOrderPDF(oActiveForm, cardcode, iCode) Then Throw New ArgumentException(sErrDesc)
    '        oRecordSet.DoQuery("Select ITM1.Price,OITM.ItemCode,OITM.ItemName from OITM Inner Join ITM1 on OITM.itemcode=ITM1.itemcode Where OITM.itemcode= '" & iCode & "' and ITM1.pricelist=1")
    '        If oRecordSet.RecordCount > 0 Then
    '            dblPrice = oRecordSet.Fields.Item("Price").Value
    '            itemCode = oRecordSet.Fields.Item("ItemCode").Value
    '            itemDesc = oRecordSet.Fields.Item("ItemName").Value
    '        End If
    '        oRecordSet.DoQuery("Select CardName From OCRD Where CardCode='" & cardcode & "'")

    '        oGeneralService = p_oDICompany.GetCompanyService.GetGeneralService("FCPO")

    '        oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
    '        oGeneralData.SetProperty("U_VCode", cardcode)
    '        oGeneralData.SetProperty("U_VName", oRecordSet.Fields.Item("CardName").Value)
    '        oGeneralData.SetProperty("U_PONo", DocLastKey)
    '        oGeneralData.SetProperty("U_PODate", Today.Date.ToString("yyyyMMdd"))
    '        oGeneralData.SetProperty("U_PODay", Today.DayOfWeek.ToString.Substring(0, 3))
    '        ' oGeneralData.SetProperty("U_POTime", Now.ToString("HH:mm tt"))
    '        oGeneralData.SetProperty("U_POITPD", dblPrice)
    '        oGeneralData.SetProperty("U_POStatus", "Open")


    '        'sql = "Insert Into [@OBT_TB08_FFCPO] (DocEntry,DocNum,U_VCode,U_VName,U_PONo,U_PODate,U_PODay,U_POTime,U_POITPD,U_POStatus) Values " & _
    '        '           "(" & Convert.ToInt32(docEntry) & _
    '        '            "," & Convert.ToInt32(docEntry) & _
    '        '            "," & IIf(oActiveForm.Items.Item("ed_V").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_V").Specific.Value), "NULL") & _
    '        '            "," & IIf(oActiveForm.Items.Item("ed_ShpAgt").Specific.Value <> "", FormatString(oActiveForm.Items.Item("ed_ShpAgt").Specific.Value), "NULL") & _
    '        '            "," & DocLastKey & _
    '        '            ",'" & Today.Date.ToString("yyyyMMdd") & _
    '        '            "','" & Today.DayOfWeek.ToString.Substring(0, 3) & _
    '        '             "','" & Now.ToString("HHmm") & _
    '        '             "'," & dblPrice & _
    '        '             ",'Open')"

    '        'oRecordSet.DoQuery(sql)
    '        'oRecordSet.DoQuery("UPDATE NNM1 SET NEXTNUMBER=NEXTNUMBER+1 WHERE OBJECTCODE='FCPO'")


    '        oChildren = oGeneralData.Child("OBT_TB09_FFCPOITEM")
    '        oChild = oChildren.Add
    '        oChild.SetProperty("U_POINO", itemCode)
    '        oChild.SetProperty("U_POIDesc", itemDesc)
    '        oChild.SetProperty("U_POIPrice", dblPrice)
    '        oChild.SetProperty("U_POIQty", 1)
    '        oChild.SetProperty("U_POIAmt", dblPrice)
    '        oChild.SetProperty("U_POIGST", "ZI")
    '        oChild.SetProperty("U_POITot", dblPrice)

    '        oGeneralService.Add(oGeneralData)
    '        'If oMatrix.RowCount > 0 Then
    '        '    If oMatrix.RowCount = 1 And oMatrix.Columns.Item("colDocNo").Cells.Item(oMatrix.RowCount).Specific.Value = "" Then
    '        '        oMatrix.Clear()
    '        '    End If

    '        '    If oMatrix.RowCount = 0 Then
    '        '        i = 1
    '        '    Else
    '        '        i = oMatrix.RowCount + 1
    '        '    End If
    '        'End If



    '        'sql = "Insert Into [@OBT_TB09_FFCPOITEM] (DocEntry,LineId,U_POINO,U_POIDesc,U_POIPrice,U_POIQty,U_POIAmt,U_POIGST,U_POITot) Values " & _
    '        '     "(" & Convert.ToInt32(docEntry) & _
    '        '      "," & 1 & _
    '        '      ",'" & itemCode & _
    '        '      "','" & itemDesc & _
    '        '      "'," & dblPrice & _
    '        '      "," & 1 & _
    '        '      "," & dblPrice & _
    '        '       ",'ZI'" & _
    '        '       "," & dblPrice & ")"
    '        'oRecordSet.DoQuery(sql)

    '        sql = "select DocEntry,U_PONo,U_PODate,U_VCode,U_TPlace,U_PODate,U_POTime,U_POStatus,U_PORMKS from " & _
    '                               "[@OBT_TB08_FFCPO] where DocEntry = " & docEntry

    '        If matrixName = "mx_Bunk" Then
    '            If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB032_BUNKER", True) Then Throw New ArgumentException(sErrDesc)
    '            pdfFilename = "Bunker"
    '            originpdf = "BunkTOPS.pdf"
    '        ElseIf matrixName = "mx_Fork" Then
    '            If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB029_FORKLIFT", True) Then Throw New ArgumentException(sErrDesc)
    '            pdfFilename = "ForkLift"
    '            originpdf = "F&CSOPS.pdf"
    '        ElseIf matrixName = "mx_Crane" Then
    '            If Not PopulatePurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB031_CRANE", True) Then Throw New ArgumentException(sErrDesc)
    '            pdfFilename = "Crane"
    '            originpdf = "F&CSOPS.pdf"
    '        End If

    '        oActiveForm.Items.Item("1").Click()

    '        mainFolder = p_fmsSetting.DocuPath
    '        jobNo = oActiveForm.Items.Item("ed_JobNo").Specific.Value
    '        pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
    '        Dim reader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(Application.StartupPath.ToString & "\" & originpdf)
    '        Dim pdfoutputfile As FileStream = New FileStream(pdffilepath, System.IO.FileMode.Create)
    '        Dim formfiller As iTextSharp.text.pdf.PdfStamper = New iTextSharp.text.pdf.PdfStamper(reader, pdfoutputfile)
    '        Dim ac As iTextSharp.text.pdf.AcroFields = formfiller.AcroFields
    '        'ac.SetField("txtDocEntry", oActiveForm.Items.Item("ed_DocNum").Specific.Value)
    '        'ac.SetField("txtLineId", i.ToString)
    '        formfiller.Close()
    '        reader.Close()
    '        Process.Start(pdffilepath)
    '        CreatePOPDF = True
    '    Catch ex As Exception
    '        CreatePOPDF = False
    '        MessageBox.Show(ex.Message)
    '    End Try

    'End Function
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
    End Sub
#End Region

#Region "---------- 'MSW - Voucher Tab View List & Edit"

    Private Sub AddUpdateVoucher(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)

        ' **********************************************************************************
        '   Function    :   AddUpdateVoucher()
        '   Purpose     :   This function will be providing to Add and Update to Voucher form  
        '                   of ImportSeaFCL Form
        '               
        '   Parameters  :  ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :  ByVal DataSource As String, ByVal ProcressedState As Boolean          
        '   Return      :  No
        ' **********************************************************************************
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
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
                    ' .SetValue("U_PayToAdd", .Offset, pForm.Items.Item("ed_PayTo").Specific.Value)
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
                    .SetValue("U_VDocNum", .Offset, pForm.Items.Item("ed_DocNum").Specific.Value) 'MSW Voucher POP Up

                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    'Public Sub SetVoucherDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)
    '    Dim sErrDesc As String = String.Empty
    '    Try
    '        pForm.Freeze(True)
    '        Dim oComboBank As SAPbouiCOM.ComboBox
    '        Dim oComboCurr As SAPbouiCOM.ComboBox
    '        Dim oComboGST As SAPbouiCOM.ComboBox
    '        oComboBank = pForm.Items.Item("cb_BnkName").Specific
    '        oComboCurr = pForm.Items.Item("cb_PayCur").Specific
    '        oComboGST = pForm.Items.Item("cb_GST").Specific

    '        oComboBank.Select(BankName, SAPbouiCOM.BoSearchKey.psk_ByValue)
    '        oComboCurr.Select(Currency, SAPbouiCOM.BoSearchKey.psk_ByValue)
    '        oComboGST.Select(GST, SAPbouiCOM.BoSearchKey.psk_ByValue)


    '        pForm.Items.Item("ed_PayTo").Specific.Value = PayTo
    '        pForm.Items.Item("ed_PayRate").Specific.Value = PayType

    '        If PayType = "Cash" Then
    '            pForm.DataSources.UserDataSources.Item("CASH").ValueEx = "1"
    '            pForm.DataSources.UserDataSources.Item("CHEQUE").ValueEx = "2"
    '        ElseIf PayType = "Cheque" Then
    '            pForm.DataSources.UserDataSources.Item("CHEQUE").ValueEx = "1"
    '            pForm.DataSources.UserDataSources.Item("CASH").ValueEx = "2"
    '        End If

    '        pForm.Items.Item("ed_VocNo").Specific.Value = VocNo
    '        pForm.Items.Item("ed_PosDate").Specific.Value = PaymentDate
    '        pForm.Items.Item("ed_PJobNo").Specific.Value = pForm.Items.Item("ed_JobNo").Specific.Value
    '        pForm.Items.Item("ed_VRemark").Specific.Value = Remark
    '        pForm.Items.Item("ed_VPrep").Specific.Value = PrepBy
    '        pForm.Items.Item("ed_SubTot").Specific.Value = SubTotal
    '        pForm.Items.Item("ed_GSTAmt").Specific.Value = GSTAmt
    '        pForm.Items.Item("ed_Total").Specific.Value = Total
    '        pForm.Items.Item("ed_Cheque").Specific.Value = CheqNo
    '        pForm.Items.Item("ed_PayRate").Specific.Value = ExRate
    '        pForm.Items.Item("ed_VedName").Specific.Value = VendorName
    '        If HolidaysMarkUp(pForm, pForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, "") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

    '        pForm.Freeze(False)

    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub GetVoucherDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
    '    Try
    '        VocNo = pMatrix.Columns.Item("colVocNo").Cells.Item(Index).Specific.Value
    '        VendorName = pMatrix.Columns.Item("colVedName").Cells.Item(Index).Specific.Value
    '        PayTo = pMatrix.Columns.Item("colPayTo").Cells.Item(Index).Specific.Value
    '        PayType = pMatrix.Columns.Item("colPayType").Cells.Item(Index).Specific.Value
    '        BankName = pMatrix.Columns.Item("colBnkName").Cells.Item(Index).Specific.Value
    '        CheqNo = pMatrix.Columns.Item("colCheqNo").Cells.Item(Index).Specific.Value
    '        Status = pMatrix.Columns.Item("colStatus").Cells.Item(Index).Specific.Value
    '        Currency = pMatrix.Columns.Item("colCurCode").Cells.Item(Index).Specific.Value
    '        PaymentDate = pMatrix.Columns.Item("colPosDate").Cells.Item(Index).Specific.Value
    '        GST = pMatrix.Columns.Item("colGST").Cells.Item(Index).Specific.Value
    '        Total = Convert.ToDouble(pMatrix.Columns.Item("colTotal").Cells.Item(Index).Specific.Value)
    '        PrepBy = pMatrix.Columns.Item("colPrepBy").Cells.Item(Index).Specific.Value

    '        ExRate = pMatrix.Columns.Item("colExRate").Cells.Item(Index).Specific.Value
    '        GSTAmt = pMatrix.Columns.Item("colGSTAmt").Cells.Item(Index).Specific.Value
    '        SubTotal = pMatrix.Columns.Item("colSubTot").Cells.Item(Index).Specific.Value
    '        Remark = pMatrix.Columns.Item("colRemark").Cells.Item(Index).Specific.Value
    '    Catch ex As Exception

    '    End Try
    'End Sub


#End Region

#Region "---------- 'MSW - Purchase Voucher Save To Draft 09-03-2010"
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
        Dim oForm As SAPbouiCOM.Form
        Dim RS As SAPbobsCOM.Recordset
        oForm = p_oSBOApplication.Forms.Item("IMPORTSEAFCL")
        RS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery("select Code,Name from ovtg Where Category = 'I'")
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
        '   Parameters  :   ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer        '                          
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
        'Function   : CalRateArmetEsc()
        'Purpose    : This function provide for calculate  Rate amount for purchae order itmes
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

        '====CalCulate Total======'
        CalculateTotal(oActiveForm, oMatrix)
    End Sub
    '----------------------------------------------------------------------------------'
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
                'oEditText.ChooseFromListUID = ""
                oEditText.ChooseFromListUID = "CFLTKRV"
                'oEditText.ChooseFromListAlias = "CardName"
            End If
            AddChooseFromListByOption = RTN_SUCCESS
        Catch ex As Exception

        End Try
    End Function

    Private Sub LoadPaymentVoucher(ByRef oActiveForm As SAPbouiCOM.Form)

        '=============================================================================
        'Function   : LoadPaymentVoucher()
        'Purpose    : This function to provide for load payment voucher form in ImportSeaFCL 
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
        oPayForm.Items.Item("ed_PayRate").Visible = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oPayForm.Items.Item("ed_DocNum").Specific.Value = GetNewKey("VOUCHER", oRecordSet)
        oPayForm.Items.Item("ed_PosDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_InvDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oPayForm.Items.Item("ed_PJobNo").Specific.Value = oActiveForm.Items.Item("ed_JobNo").Specific.Value
        'oPayForm.Items.Item("ed_FrDocNo").Specific.Value = oActiveForm.Items.Item("ed_DocNum").Specific.Value
        oPayForm.Items.Item("ed_VPrep").Specific.Value = p_oDICompany.UserName.ToString()
        ' If HolidayMarkUp(oActiveForm, oPayForm.Items.Item("ed_PosDate").Specific, p_oCompDef.HolidaysName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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

    Private Sub EnabledDispatchMatrix(ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, ByVal pValue As Boolean)

        ' **********************************************************************************
        '   Function    :   EnabledDispatchMatrix()
        '   Purpose     :   This function will be providing to enable dispatch matrix  in
        '                   in ImportSeaFCL form
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByRef pMatrix As SAPbouiCOM.Matrix, 
        '                   ByVal pValue As Boolea        
        '   Return      :   No
        '                  
        ' ***********************************************************************************
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

        End Try
    End Sub

    Private Sub ThreadMethodEmail()
        'Try
        '    Dim dlg As New OpenFileDialog()
        '    dlg.InitialDirectory = "\\MIDSERVER2\CertConfig"
        '    dlg.Filter = "All files (*.*)|*.*"
        '    dlg.ShowDialog()
        '    fname = dlg.FileName
        '    If (System.IO.Path.GetExtension(fname) = ".exe") Then
        '        MessageBox.Show("Exe file may not be able to send.", "Warning")
        '        fname = ""
        '    End If
        '    ' MessageBox.Show(dlg.FileName)
        'Catch ex As Exception

        'End Try
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

    Private Sub Start(ByRef pform As SAPbouiCOM.Form)

        '=============================================================================
        'Function   : Start()
        'Purpose    : This function to provide for To process with pdf form in ImportSeaFCL
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

        End Try
    End Sub

    Private Sub SetMatrixSeqNo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal ColName As String)

        '=============================================================================
        'Function   : SetMatrixSeqNo()
        'Purpose    : This function to provide to add Matrix Sequence No in ImportSeaFCL form

        'Parameters : ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '           : ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer
        'Return     : No
        '             
        '==========================================

        For i As Integer = 1 To oMatrix.RowCount
            oMatrix.Columns.Item(ColName).Cells.Item(i).Specific.Value = i
        Next
    End Sub

    Private Sub AddUpdateDisp(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer)

        '=============================================================================
        'Function   : AddUpdateDisp()
        'Purpose    : This function to provide to add Update Dispatch  form in ImportSeaFCL

        'Parameters : ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '           : ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal Index As Integer
        'Return     : No
        '             
        '==========================================

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

        End Try
    End Sub

    Private Sub SetDispatchDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)

        ' **********************************************************************************
        '   Function    :   SetDispatchDataToEditTabByIndex()
        '   Purpose     :   This function will be providing to Dispatch Data to edit by index  for
        '                   ImportSeaFCL Form
        '               
        '   Parameters  :   ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix,
        '               :   ByVal Index As Integer
        '   Return      :   No
        '               
        '                   
        ' **********************************************************************************


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

    Private Function Validateforform(ByVal ItemUID As String, ByVal oActiveForm As SAPbouiCOM.Form) As Boolean

        ' **********************************************************************************
        '   Function    :   Validateforform()
        '   Purpose     :   This function will be providing to validate form  for
        '                   ImportSeaFCL Form
        '               
        '   Parameters  :   ByVal ItemUID As String, ByVal oActiveForm As SAPbouiCOM.Form
        '               
        '   Return      :   False -Fail
        '               :   True  -Success
        '                   
        ' **********************************************************************************
        Try
            If (ItemUID = "ed_Name" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_Name").Specific.String) Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Customer Code", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
            ElseIf (ItemUID = "cb_PCode" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("cb_PCode").Specific.Value) Then
                p_oSBOApplication.SetStatusBarMessage("Must Select Port Of Loading", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
                'ElseIf (ItemUID = "ed_ShpAgt" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_ShpAgt").Specific.String) Then
                '    p_oSBOApplication.SetStatusBarMessage("Must Choose Shipping Agent", SAPbouiCOM.BoMessageTime.bmt_Short)
                '    Return True
            ElseIf (ItemUID = "ed_Yard" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_Yard").Specific.String) Then
                p_oSBOApplication.SetStatusBarMessage("Must Choose Return Yard", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
                'ElseIf (ItemUID = "ed_LCDDate" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_LCDDate").Specific.String) Then
                '    p_oSBOApplication.SetStatusBarMessage("Must Choose Last Clearance Date", SAPbouiCOM.BoMessageTime.bmt_Short)
                '    Return True
            ElseIf (ItemUID = "ed_YAddr" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_YAddr").Specific.String) Then
                p_oSBOApplication.SetStatusBarMessage("Must Fill Return Yard Address", SAPbouiCOM.BoMessageTime.bmt_Short)
                Return True
                'ElseIf (ItemUID = "ed_JobNo" Or ItemUID = " ") And String.IsNullOrEmpty(oActiveForm.Items.Item("ed_JobNo").Specific.String) Then
                '    p_oSBOApplication.SetStatusBarMessage("Must Fill Job No", SAPbouiCOM.BoMessageTime.bmt_Short)
                '    Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
        
    End Function

    Private Sub LoadHolidayMarkUp(ByVal oActiveForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   LoadHolidayMarkUp()
        '   Purpose     :   This function will be providing to load Holiday Markup form for
        '                   ExporeSeaLcl Form
        '               
        '   Parameters  :   ByVal ExportSeaLCLForm As SAPbouiCOM.Form
        '               
        '   Return      :   No
        '                   
        ' **********************************************************************************
        Dim sErrDesc As String = String.Empty
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ETADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ETADay").Specific, oActiveForm.Items.Item("ed_ETAHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ConLast").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ConLDay").Specific, oActiveForm.Items.Item("ed_ConLTim").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        'If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_LCDDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_LCDDay").Specific, oActiveForm.Items.Item("ed_LCDHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        'If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CnDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CnDay").Specific, oActiveForm.Items.Item("ed_CnHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        'If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_CgDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_CgDay").Specific, oActiveForm.Items.Item("ed_CgHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_JbDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_JbDay").Specific, oActiveForm.Items.Item("ed_JbHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        'If HolidaysMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspDate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspDay").Specific, oActiveForm.Items.Item("ed_DspHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_DspCDte").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_DspCDay").Specific, oActiveForm.Items.Item("ed_DspCHr").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If HolidayMarkUp(oActiveForm, oActiveForm.Items.Item("ed_ADate").Specific, p_oCompDef.HolidaysName, sErrDesc, oActiveForm.Items.Item("ed_ADay").Specific, oActiveForm.Items.Item("ed_ATime").Specific) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
    End Sub

    Private Sub EnabledHeaderControls(ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean)


        ' **********************************************************************************
        '   Function    :   EnabledHeaderControls()
        '   Purpose     :   This function will be providing to enable header control in 
        '                   ImportSeaFCL()
        '                   
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pValue As Boolean
        '   Return      :   No

        '*************************************************************

        pForm.Items.Item("ed_ShpAgt").Enabled = pValue
        pForm.Items.Item("ed_OBL").Enabled = pValue
        pForm.Items.Item("ed_HBL").Enabled = pValue
        'pForm.Items.Item("ed_Conn").Enabled = pValue
        pForm.Items.Item("ed_Vessel").Enabled = pValue
        pForm.Items.Item("ed_Voy").Enabled = pValue
        pForm.Items.Item("cb_PCode").Enabled = pValue
        pForm.Items.Item("ed_ETADate").Enabled = pValue
        pForm.Items.Item("ed_ETAHr").Enabled = pValue
        'pForm.Items.Item("ed_TotChWt").Enabled = pValue ' MSW To Edit New Ticket 07-09-2011
        'pForm.Items.Item("ed_Cont20").Enabled = pValue
        'pForm.Items.Item("ed_Cont40").Enabled = pValue
        'pForm.Items.Item("ed_Cont45").Enabled = pValue

        pForm.Items.Item("ed_ConDate").Enabled = pValue
        pForm.Items.Item("ed_ConTime").Enabled = pValue
        pForm.Items.Item("ed_ConLast").Enabled = pValue
        pForm.Items.Item("ed_ConLTim").Enabled = pValue
        pForm.Items.Item("ed_CrgDsc").Enabled = pValue 'MSW New Ticket 07-09-2011
        'pForm.Items.Item("ed_TotalM3").Enabled = pValue
        'pForm.Items.Item("ed_TotalWt").Enabled = pValue
        'pForm.Items.Item("ed_NOP").Enabled = pValue
        'pForm.Items.Item("cb_PType").Enabled = pValue

        'MSW to Edit #1001
        If pValue = True And pForm.Items.Item("cb_JobType").Specific.Value = "" Then
            pForm.Items.Item("cb_JobType").Enabled = pValue
        End If
        'End MSW to Edit #1001
        pForm.Items.Item("cb_JbStus").Enabled = pValue
        pForm.Items.Item("ed_Yard").Enabled = pValue
        'pForm.Items.Item("ed_JobNo").Enabled = pValue
        'pForm.Items.Item("ed_LCDDate").Enabled = pValue
        'pForm.Items.Item("ed_LCDHr").Enabled = pValue

        'pForm.Items.Item("ch_CnUn").Enabled = pValue
        'pForm.Items.Item("ch_CgCl").Enabled = pValue
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
    'MSW To Edit
    Private Sub CalculateNoOfBoxes(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix)
        If pMatrix.RowCount > 1 Then
            For i As Integer = Convert.ToInt32(pForm.Items.Item("ed_ItemNo").Specific.Value) To pMatrix.RowCount - 1
                pMatrix.Columns.Item("colBox").Cells.Item(i + 1).Specific.Value() = pMatrix.Columns.Item("colBoxLast").Cells.Item(i).Specific.Value() + 1
                pMatrix.Columns.Item("colBoxLast").Cells.Item(i + 1).Specific.Value() = (Convert.ToDouble(pMatrix.Columns.Item("colBox").Cells.Item(i + 1).Specific.Value()) + Convert.ToDouble(pMatrix.Columns.Item("colTNBox").Cells.Item(i + 1).Specific.Value())) - 1
            Next
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
                If Not PopulateTruckPurchaseHeader(oActiveForm, oMatrix, sql, "@OBT_TB006_TRUCKING") Then Throw New ArgumentException(sErrDesc)
            Else
                If matrixName = "mx_Fork" Then
                    dbSource = "@OBT_TB029_FORKLIFT"
                ElseIf matrixName = "mx_Armed" Then
                    dbSource = "@OBT_TB030_ARMES"
                ElseIf matrixName = "mx_Crane" Then
                    dbSource = "@OBT_TB031_CRANE"
                ElseIf matrixName = "mx_Bunk" Then
                    dbSource = "@OBT_TB032_BUNKER"
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

    'MSW to edit New Ticket 08-09-2011
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
        rptPath = Application.StartupPath.ToString & "\Trucking Instruction ImportFCL.rpt"
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
        rptPath = Application.StartupPath.ToString & "\A6 Label Import Sea FCL.rpt"
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

    Private Sub PreviewDispIns(ByRef ParentForm As SAPbouiCOM.Form)

        ' **********************************************************************************
        '   Function    :   PreviewPO()
        '   Purpose     :   This function provide to view of purchase order form when purchase  
        '                   order items save to database
        '   Parameters  :   ByRef ParentForm As SAPbouiCOM.Form,ByRef oActiveForm As SAPbouiCOM.Form
        '   return      :   No          
        ' **********************************************************************************
        'p_oSBOApplication.MessageBox("Dispatch Instruction")
        'Dim DocNum As Integer
        'Dim InsDoc As Integer
        'Dim oMatrix As SAPbouiCOM.Matrix
        'rptDocument = New ReportDocument

        'pdfFilename = "Dispatch Instruction"
        'mainFolder = p_fmsSetting.DocuPath
        'jobNo = ParentForm.Items.Item("ed_JobNo").Specific.Value
        'rptPath = Application.StartupPath.ToString & "\Dispatch Instruction ImportFCL.rpt"
        'pdffilepath = reportuti.CreateJobFolderFile(mainFolder, jobNo, pdfFilename)
        'rptDocument.Load(rptPath)
        'rptDocument.Refresh()
        'oMatrix = ParentForm.Items.Item("mx_DispTab").Specific

        'DocNum = Convert.ToInt32(ParentForm.Items.Item("ed_DocNum").Specific.Value)
        'InsDoc = Convert.ToInt32(oMatrix.GetNextSelectedRow)

        'rptDocument.SetParameterValue("@DocEntry", DocNum)
        'rptDocument.SetParameterValue("@InsDocNo", InsDoc)
        'reportuti.SetDBLogIn(rptDocument)
        'If Not pdffilepath = String.Empty Then
        '    reportuti.ExportCRToPDF(pdffilepath, clsReportUtilities.ExportFileType.PDFFILE, rptDocument,true)
        'End If
    End Sub
End Module
