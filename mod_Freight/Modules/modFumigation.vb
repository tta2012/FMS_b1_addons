Option Explicit On

Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Drawing.Printing
Imports System.Threading

Module modFumigation
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String
    Private currentRow As Integer
    Private ActiveMatrix As String
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sql As String = ""
    Private ObjMatrix As SAPbouiCOM.Matrix
    Dim MJobForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
    Dim dtmatrix As SAPbouiCOM.DataTable
    Dim rowIndex As Integer
    Dim selectedRow As Integer

    Public Function DoExportSeaFCLItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Long
        Dim FumigationForm As SAPbouiCOM.Form = Nothing
        Dim OutriderForm As SAPbouiCOM.Form = Nothing
        Dim ExportSeaFCLForm As SAPbouiCOM.Form = Nothing
        Dim FunctionName As String = "DoExportSeaFCLItemEvent()"
        Dim oCombo As SAPbouiCOM.ComboBox = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
        Dim oMatrix As SAPbouiCOM.Matrix = Nothing
        Dim poMatrix As SAPbouiCOM.Matrix = Nothing
        Dim CPOForm As SAPbouiCOM.Form = Nothing
        Dim CPOMatrix As SAPbouiCOM.Matrix = Nothing
        Dim oComboPO As SAPbouiCOM.ComboBox = Nothing
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim tableName As String = ""
        Dim itemtableName As String = ""
        Dim detailtableName As String = ""
        Dim source As String = ""
        Dim matrixName As String = ""

        ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
        Try
            Select Case pVal.FormTypeEx
                Case "FUMIGATION", "OUTRIDER"

                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    FumigationForm = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
                    If pVal.FormTypeEx = "FUMIGATION" Then
                        tableName = "@OBT_TBL01_FUMIGAT"
                        itemtableName = "@OBT_TBL02_ITEM"
                        source = "Fumigation"
                        matrixName = "mx_Fumi"
                    ElseIf pVal.FormTypeEx = "OUTRIDER" Then
                        tableName = "@OBT_TBL03_OUTRIDER"
                        itemtableName = "@OBT_TBL04_ITEM"
                        source = "Outrider"
                        matrixName = "mx_Outer"
                    End If
                    Try
                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(FumigationForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If
                    If pVal.Before_Action = False Then
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "fo_Edit"
                                        FumigationForm.PaneLevel = 1
                                        FumigationForm.Items.Item("fo_Edit").Specific.Select()
                                        If FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add" Then
                                            FumigationForm.Items.Item("bt_FumiNew").Click()
                                        End If

                                    Case "fo_View"
                                        FumigationForm.PaneLevel = 2
                                        FumigationForm.Items.Item("fo_View").Specific.Select()
                                        If pVal.FormTypeEx = "FUMIGATION" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Item", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                        ElseIf pVal.FormTypeEx = "OUTRIDER" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_LocFrom", "ed_LocTo", "ed_Remark", "ed_IRemark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                        End If

                                        FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add"

                                    Case "fo_Set"
                                        FumigationForm.PaneLevel = 3
                                        FumigationForm.Items.Item("fo_Set").Specific.Select()
                                End Select
                        End Select

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Try
                                If pVal.ItemUID = "ed_FSIA" Then

                                    FumigationForm.DataSources.UserDataSources.Item("SIA").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                    FumigationForm.DataSources.UserDataSources.Item("SIACODE").ValueEx = oDataTable.GetValue(0, 0).ToString
                                    FumigationForm.DataSources.UserDataSources.Item("SIATEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString()
                                End If
                                If pVal.ItemUID = "ed_FCode" Or pVal.ItemUID = "ed_FName" Then

                                    FumigationForm.DataSources.UserDataSources.Item("VCODE").ValueEx = oDataTable.GetValue(0, 0).ToString
                                    FumigationForm.DataSources.UserDataSources.Item("VNAME").ValueEx = oDataTable.GetValue(1, 0).ToString
                                    FumigationForm.DataSources.UserDataSources.Item("CPERSON").ValueEx = oDataTable.Columns.Item("CntctPrsn").Cells.Item(0).Value.ToString()
                                End If
                                If pVal.ItemUID = "ed_ICode" Or pVal.ItemUID = "ed_IDesc" Then

                                    FumigationForm.DataSources.DBDataSources.Item(itemtableName).SetValue("U_ICode", 0, oDataTable.GetValue(0, 0).ToString)
                                    FumigationForm.DataSources.DBDataSources.Item(itemtableName).SetValue("U_IDes", 0, oDataTable.GetValue(1, 0).ToString)

                                End If


                            Catch ex As Exception

                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            Try

                                If pVal.ItemUID = "cb_FCntact" Then
                                    Dim oComboContact As SAPbouiCOM.ComboBox = FumigationForm.Items.Item("cb_FCntact").Specific
                                    FumigationForm.DataSources.DBDataSources.Item(tableName).SetValue("U_CPerson", 0, oComboContact.Selected.Description.ToString)
                                    FumigationForm.Items.Item("ed_FVRef").Specific.Active = True
                                End If



                            Catch ex As Exception
                                MessageBox.Show(ex.ToString())
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "bt_CrPO" Then
                                Dim VCode As String = FumigationForm.Items.Item("ed_FCode").Specific.Value
                                Dim ICode As String = FumigationForm.Items.Item("ed_ICode").Specific.Value
                                If FumigationForm.Items.Item("ed_FCode").Specific.Value = "" Then
                                    p_oSBOApplication.SetStatusBarMessage("There is no Vendor to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                ElseIf FumigationForm.Items.Item("ed_ICode").Specific.Value = "" Then
                                    p_oSBOApplication.SetStatusBarMessage("There is no Item to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                ElseIf Convert.ToDouble(FumigationForm.Items.Item("ed_IQty").Specific.Value) = 0.0 Or Convert.ToDouble(FumigationForm.Items.Item("ed_IPrice").Specific.Value) = 0.0 Then
                                    p_oSBOApplication.SetStatusBarMessage("Document Total is zero", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End If
                                If BubbleEvent = True Then
                                    If Not CreateGenPO(FumigationForm, VCode, ICode, source) Then Throw New ArgumentException(sErrDesc)
                                    FumigationForm.Items.Item("bt_CrPO").Enabled = False
                                    FumigationForm.Items.Item("bt_FumiAdd").Click()
                                    SavePOPDFInEditTab(FumigationForm, source)

                                End If
                            End If
                            If pVal.ItemUID = "bt_FumiNew" Then
                                If pVal.FormTypeEx = "FUMIGATION" Then
                                    ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Item", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                ElseIf pVal.FormTypeEx = "OUTRIDER" Then
                                    ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_LocFrom", "ed_LocTo", "ed_Remark", "ed_IRemark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                End If
                                rowIndex = -1
                                FumigationForm.Items.Item("ed_FPODate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add"
                                FumigationForm.Items.Item("bt_CrPO").Enabled = True
                                FumigationForm.Items.Item("ed_FCode").Enabled = True
                                FumigationForm.Items.Item("ed_FName").Enabled = True
                            End If


                            If pVal.ItemUID = "bt_FumiAdd" Then
                                If String.IsNullOrEmpty(FumigationForm.Items.Item("ed_FCode").Specific.String) Or String.IsNullOrEmpty(FumigationForm.Items.Item("ed_FName").Specific.String) Then
                                    p_oSBOApplication.SetStatusBarMessage("Must Fill Vendor", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                ElseIf String.IsNullOrEmpty(FumigationForm.Items.Item("ed_PO").Specific.String) Then
                                    p_oSBOApplication.SetStatusBarMessage("Must Fill PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                Else

                                    poMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                    If pVal.FormTypeEx = "FUMIGATION" Then
                                        oMatrix = FumigationForm.Items.Item("mx_Fumi").Specific
                                    ElseIf pVal.FormTypeEx = "OUTRIDER" Then
                                        oMatrix = FumigationForm.Items.Item("mx_Outer").Specific
                                    End If

                                    If FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add" Then

                                        AddUpdateFumigation(FumigationForm, oMatrix, tableName, True, source)
                                        SaveToPOTab(ExportSeaFCLForm, poMatrix, True, FumigationForm.Items.Item("ed_PO").Specific.Value, FumigationForm.Items.Item("ed_PODocNo").Specific.Value, FumigationForm.Items.Item("ed_FName").Specific.Value, FumigationForm.Items.Item("ed_FPODate").Specific.Value, source, FumigationForm.Items.Item("ed_FPOStus").Specific.Value)
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        If source = "Outrider" Then ' to combine
                                            CreatePOPDF(FumigationForm, "mx_Outer", "")
                                        End If
                                    Else

                                        AddUpdateFumigation(FumigationForm, oMatrix, tableName, False, source)
                                        If Not CreateFFCPO(FumigationForm, False, source) Then Throw New ArgumentException(sErrDesc)
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    If FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    FumigationForm.Items.Item("1").Click()
                                    FumigationForm.Items.Item("2").Specific.Caption = "Close"
                                    FumigationForm.Items.Item("bt_CrPO").Enabled = False
                                    ExportSeaFCLForm.Items.Item("1").Click()
                                    FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Update"

                                    If FumigationForm.Items.Item("ed_MultiJb").Specific.Value = "Y" Then
                                        FumigationForm.Items.Item("bt_DJob").Enabled = True
                                    Else
                                        FumigationForm.Items.Item("bt_DJob").Enabled = False
                                    End If

                                    End If
                            End If
                            If pVal.ItemUID = "bt_IView" Then
                                If source = "Fumigation" Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                                    PreviewPO(ExportSeaFCLForm, FumigationForm)
                                End If
                                If source = "Outrider" Then
                                    CreatePOPDF(FumigationForm, "mx_Outer", "")
                                End If
                            End If
                            If pVal.ItemUID = "bt_PreView" Then
                                If selectedRow > 0 Then
                                    If source = "Fumigation" Then
                                        oMatrix = FumigationForm.Items.Item(matrixName).Specific
                                        PreviewPOInViewList(ExportSeaFCLForm, selectedRow, oMatrix)
                                    ElseIf source = "Outrider" Then
                                        CreatePOPDF(FumigationForm, matrixName, "View", selectedRow)
                                    End If
                                Else
                                    p_oSBOApplication.MessageBox("Need to select one Row to Preview.")

                                End If
                            End If
                            If pVal.ItemUID = "bt_MJob" Then
                                modExportSeaFCL.LoadMultiJobNormalForm(FumigationForm, "MultiJobNormal.srf", source)
                                oActiveForm = p_oSBOApplication.Forms.ActiveForm
                                modMultiJobForNormal.LoadMultiJForm()
                            End If


                            If pVal.ItemUID = "bt_DJob" Then
                                LoadDetachJobNormalForm(FumigationForm, "DetachJobNormal.srf", source, "ed_PO")
                                oActiveForm = p_oSBOApplication.Forms.ActiveForm
                                modDetachJobNormal.LoadDetachJForm()
                            End If

                            If pVal.ItemUID = "bt_Email" Then
                                modExportSeaFCL.PreviewOutLookMail(FumigationForm, source)
                            End If

                            If pVal.ItemUID = "1" Then
                                If pVal.Action_Success = True Then
                                    If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        'FumigationForm = p_oSBOApplication.Forms.GetForm(pVal.FormTypeEx, 1)

                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                        p_oSBOApplication.ActivateMenuItem("1291")
                                        FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                    End If



                                End If
                            End If
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                            If pVal.ItemUID = "mx_Fumi" Or pVal.ItemUID = "mx_Outer" Then
                                oMatrix = FumigationForm.Items.Item(pVal.ItemUID).Specific
                                If pVal.Row > 0 Then
                                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                    oMatrix.SelectRow(pVal.Row, True, False)
                                    selectedRow = pVal.Row
                                End If
                            End If
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK And pVal.BeforeAction = False Then
                            If pVal.FormTypeEx = "FUMIGATION" Then
                                oMatrix = FumigationForm.Items.Item("mx_Fumi").Specific
                            ElseIf pVal.FormTypeEx = "OUTRIDER" Then
                                oMatrix = FumigationForm.Items.Item("mx_Outer").Specific
                            End If
                            If oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value = "Closed" Or oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value = "Cancelled" Then
                                p_oSBOApplication.SetStatusBarMessage("This PO is already " & oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value & ".", SAPbouiCOM.BoMessageTime.bmt_Short)
                            Else

                                If pVal.ItemUID = "mx_Fumi" Or pVal.ItemUID = "mx_Outer" Then
                                    If pVal.Row > 0 Then
                                        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        oMatrix.SelectRow(pVal.Row, True, False)

                                        GetDataFromMatrixByIndex(FumigationForm, oMatrix, pVal.Row, source)

                                        FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Update"
                                        FumigationForm.Items.Item("fo_Edit").Specific.Select()
                                        FumigationForm.Items.Item("bt_CrPO").Enabled = False
                                        FumigationForm.Items.Item("ed_FCode").Enabled = False
                                        FumigationForm.Items.Item("ed_FName").Enabled = False
                                    End If
                                End If
                                If FumigationForm.Items.Item("ed_OriPO").Specific.Value = "Y" Then
                                    FumigationForm.Items.Item("bt_DJob").Enabled = True
                                Else
                                    FumigationForm.Items.Item("bt_DJob").Enabled = False
                                End If
                            End If

                        End If
                    End If

                Case "CRANE", "FORKLIFT", "CRATE", "BUNKER", "TOLL"

                    oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                    FumigationForm = p_oSBOApplication.Forms.Item(pVal.FormTypeEx)
                    If pVal.FormTypeEx = "CRANE" Then
                        source = "Crane"
                        tableName = "@OBT_TB33_CRANE"
                        itemtableName = "@OBT_TB01_CRANEITEM"
                        detailtableName = "@OBT_TB01_CRNDETAIL"
                        matrixName = "mx_Crane"
                    ElseIf pVal.FormTypeEx = "FORKLIFT" Then 'to combine
                        source = "Forklift"
                        tableName = "@OBT_TBL05_FORKLIFT"
                        matrixName = "mx_Fork"
                        detailtableName = "@OBT_TBL07_FORDETAIL"
                    ElseIf pVal.FormTypeEx = "CRATE" Then 'to combine
                        source = "Crate"
                        tableName = "@OBT_TBL08_CRATE"
                        matrixName = "mx_Crate"
                        detailtableName = "@OBT_TBL10_CRADETAIL"
                    ElseIf pVal.FormTypeEx = "BUNKER" Then
                        source = "Bunker"
                        tableName = "@OBT_TB01_BUNKER"
                        itemtableName = "@OBT_TB01_BUNKITEM"
                        detailtableName = "@OBT_TB01_BNKDETAIL"
                        matrixName = "mx_Bunk"
                    ElseIf pVal.FormTypeEx = "TOLL" Then
                        source = "Toll"
                        tableName = "@OBT_TB01_TOLL"
                        itemtableName = "@OBT_TB01_TOLLITEM"
                        detailtableName = "@OBT_TB01_TOLLDETAIL"
                        matrixName = "mx_Toll"
                    End If
                    Try
                        FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                    Catch ex As Exception

                    End Try
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_CLOSE And pVal.BeforeAction = False Then
                        If Not RemoveFromAppList(FumigationForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                    End If
                    If pVal.Before_Action = False Then
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                Select Case pVal.ItemUID
                                    Case "fo_Edit"
                                        FumigationForm.PaneLevel = 1
                                        FumigationForm.Items.Item("fo_Edit").Specific.Select()
                                        If FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add" Then
                                            FumigationForm.Items.Item("bt_FumiNew").Click()
                                        End If
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    Case "fo_View"
                                        FumigationForm.PaneLevel = 2
                                        FumigationForm.Items.Item("fo_View").Specific.Select()
                                        If source = "Crane" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb", "ed_SIns")
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            If dtmatrix.Rows.Count <> 0 Then
                                                For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                                    dtmatrix.Rows.Remove(i)
                                                Next
                                            End If
                                            oMatrix.Clear()
                                            RowAddToMatrix(FumigationForm, oMatrix, source)
                                        ElseIf source = "Forklift" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Remark", "ed_IRemark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            If dtmatrix.Rows.Count <> 0 Then
                                                For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                                    dtmatrix.Rows.Remove(i)
                                                Next
                                            End If
                                            oMatrix.Clear()
                                            RowAddToMatrix(FumigationForm, oMatrix, source)
                                        ElseIf source = "Crate" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Desc", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            If dtmatrix.Rows.Count <> 0 Then
                                                For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                                    dtmatrix.Rows.Remove(i)
                                                Next
                                            End If
                                            oMatrix.Clear()
                                            RowAddToMatrix(FumigationForm, oMatrix, source)
                                        ElseIf source = "Bunker" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Remark", "ed_SIns", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb", "ed_CDesc", "ed_SIns", "ed_TQty", "ed_TKgs", "ed_TM3", "ed_TNEQ")
                                            FumigationForm.Items.Item("chk_1").Specific.Checked = False
                                            FumigationForm.Items.Item("chk_1a").Specific.Checked = False
                                            FumigationForm.Items.Item("chk_2").Specific.Checked = False
                                            FumigationForm.Items.Item("chk_2a").Specific.Checked = False
                                            oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtBDetail")
                                            If dtmatrix.Rows.Count <> 0 Then
                                                For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                                    dtmatrix.Rows.Remove(i)
                                                Next
                                            End If
                                            oMatrix.Clear()
                                            RowAddToMatrixBunker(FumigationForm, oMatrix)
                                        ElseIf source = "Toll" Then
                                            ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb", "ed_IRemark")
                                            oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtTDetail")
                                            If dtmatrix.Rows.Count <> 0 Then
                                                For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                                    dtmatrix.Rows.Remove(i)
                                                Next
                                            End If
                                            oMatrix.Clear()
                                            RowAddToMatrixToll(FumigationForm, oMatrix)

                                        End If

                                        FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add"


                                    Case "fo_Set"
                                        FumigationForm.PaneLevel = 3
                                        FumigationForm.Items.Item("fo_Set").Specific.Select()
                                End Select
                        End Select
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS And pVal.Before_Action = False Then
                            If pVal.ItemUID = "mx_CDetail" And pVal.ColUID = "colCType" Then
                                oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                If oMatrix.Columns.Item("colCType").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    RowAddToMatrix(FumigationForm, oMatrix, source)
                                End If
                            End If
                            If pVal.ItemUID = "mx_CDetail" And pVal.ColUID = "colTon" Then 'to combine
                                oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                If oMatrix.Columns.Item("colTon").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    RowAddToMatrix(FumigationForm, oMatrix, source) ' to combine
                                End If

                            End If
                            If pVal.ItemUID = "mx_CDetail" And pVal.ColUID = "colDimen" Then 'to combine
                                oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                If oMatrix.Columns.Item("colDimen").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    RowAddToMatrix(FumigationForm, oMatrix, source) ' to combine
                                End If

                            End If
                            'Bunker
                            If pVal.ItemUID = "mx_BDetail" And pVal.ColUID = "colPermit" Then
                                oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                If oMatrix.Columns.Item("colPermit").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    RowAddToMatrixBunker(FumigationForm, oMatrix)
                                End If
                            End If
                            If pVal.ItemUID = "mx_TDetail" And pVal.ColUID = "colICode" Then
                                oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                If oMatrix.Columns.Item("colICode").Cells.Item(oMatrix.RowCount).Specific.Value <> "" Then
                                    RowAddToMatrixToll(FumigationForm, oMatrix)
                                End If

                            End If
                            If pVal.ItemUID = "mx_TDetail" And (pVal.ColUID = "colQty" Or pVal.ColUID = "colPrice") Then
                                oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                CalTotalToll(FumigationForm, pVal.Row)
                            End If
                            If pVal.ItemUID = "mx_BDetail" And (pVal.ColUID = "colQty" Or pVal.ColUID = "colKgs" Or pVal.ColUID = "colM3" Or pVal.ColUID = "colNEQ") Then
                                oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                CalTotalBunker(FumigationForm, pVal.Row)
                            End If
                            If pVal.ItemUID = "mx_CDetail" And (pVal.ColUID = "colQty" Or pVal.ColUID = "colPrice") Then
                                oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                If Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(pVal.Row).Specific.Value) <> 0 And Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(pVal.Row).Specific.Value) <> 0 Then
                                    CalTotal(FumigationForm, pVal.Row, source)
                                End If
                            End If
                        End If
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            Try
                                If pVal.ItemUID = "mx_TDetail" And pVal.ColUID = "colICode" Then
                                    oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                    Try
                                        oMatrix.Columns.Item("colICode").Cells.Item(pVal.Row).Specific.Value = oDataTable.GetValue(0, 0).ToString
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        oMatrix.Columns.Item("colIDesc").Cells.Item(pVal.Row).Specific.Value = oDataTable.Columns.Item("ItemName").Cells.Item(0).Value.ToString
                                    Catch ex As Exception
                                    End Try
                                End If
                                If pVal.ItemUID = "ed_FSIA" Then

                                    FumigationForm.DataSources.UserDataSources.Item("SIA").ValueEx = oDataTable.Columns.Item("firstName").Cells.Item(0).Value.ToString & ", " & oDataTable.Columns.Item("lastName").Cells.Item(0).Value.ToString
                                    FumigationForm.DataSources.UserDataSources.Item("SIACODE").ValueEx = oDataTable.GetValue(0, 0).ToString
                                    FumigationForm.DataSources.UserDataSources.Item("SIATEL").ValueEx = oDataTable.Columns.Item("mobile").Cells.Item(0).Value.ToString()
                                End If
                                If pVal.ItemUID = "ed_FCode" Or pVal.ItemUID = "ed_FName" Then

                                    FumigationForm.DataSources.UserDataSources.Item("VCODE").ValueEx = oDataTable.GetValue(0, 0).ToString
                                    FumigationForm.DataSources.UserDataSources.Item("VNAME").ValueEx = oDataTable.GetValue(1, 0).ToString
                                    FumigationForm.DataSources.UserDataSources.Item("CPERSON").ValueEx = oDataTable.Columns.Item("CntctPrsn").Cells.Item(0).Value.ToString()
                                End If
                            Catch ex As Exception

                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_COMBO_SELECT Then
                            Try

                                If pVal.ItemUID = "cb_FCntact" Then
                                    Dim oComboContact As SAPbouiCOM.ComboBox = FumigationForm.Items.Item("cb_FCntact").Specific
                                    FumigationForm.DataSources.DBDataSources.Item(tableName).SetValue("U_CPerson", 0, oComboContact.Selected.Description.ToString)
                                    FumigationForm.Items.Item("ed_FVRef").Specific.Active = True
                                End If



                            Catch ex As Exception
                                MessageBox.Show(ex.ToString())
                            End Try
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "bt_CrPO" Then
                                If source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                                    oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                ElseIf source = "Bunker" Then
                                    oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                ElseIf source = "Toll" Then
                                    oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                End If

                                If CheckQtyValue(oMatrix, source) = True Then
                                    p_oSBOApplication.SetStatusBarMessage("Document Total is Zero.", SAPbouiCOM.BoMessageTime.bmt_Short)
                                    BubbleEvent = False
                                End If
                                If BubbleEvent = True Then
                                    If source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                                        oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                        If oMatrix.RowCount > 0 Then
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            If dtmatrix.Rows.Count > 0 Then
                                                dtmatrix.Rows.Remove(oMatrix.RowCount - 1)
                                                oMatrix.LoadFromDataSource()
                                            End If
                                        End If
                                    ElseIf source = "Bunker" Then
                                        oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                        If oMatrix.RowCount > 0 Then
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtBDetail")
                                            If dtmatrix.Rows.Count > 0 Then
                                                dtmatrix.Rows.Remove(oMatrix.RowCount - 1)
                                                oMatrix.LoadFromDataSource()
                                            End If
                                        End If
                                    ElseIf source = "Toll" Then
                                        oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                        If oMatrix.RowCount > 0 Then
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtTDetail")
                                            If dtmatrix.Rows.Count > 0 Then
                                                dtmatrix.Rows.Remove(oMatrix.RowCount - 1)
                                                oMatrix.LoadFromDataSource()
                                            End If
                                        End If
                                    End If

                                    Dim VCode As String = FumigationForm.Items.Item("ed_FCode").Specific.Value
                                    Dim ICode As String = FumigationForm.Items.Item("ed_ICode").Specific.Value
                                    If FumigationForm.Items.Item("ed_FCode").Specific.Value = "" Then
                                        p_oSBOApplication.SetStatusBarMessage("There is no Vendor to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf FumigationForm.Items.Item("ed_ICode").Specific.Value = "" And source <> "Toll" Then
                                        p_oSBOApplication.SetStatusBarMessage("There is no Item to create PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    End If
                                    If BubbleEvent = True Then
                                        If Not CreateGenPO(FumigationForm, VCode, ICode, source) Then Throw New ArgumentException(sErrDesc)
                                        FumigationForm.Items.Item("bt_CrPO").Enabled = False
                                        FumigationForm.Items.Item("bt_FumiAdd").Click()
                                        If source = "CRANE" Then 'to combine
                                            modExportSeaFCL.SavePOPDFInEditTab(FumigationForm, source)

                                        ElseIf source = "Bunker" Then
                                            modExportSeaFCL.CreatePOPDF(FumigationForm, "mx_Bunk", "Edit")
                                        ElseIf source = "Toll" Then
                                            modExportSeaFCL.CreatePOPDF(FumigationForm, "mx_Toll", "Edit")
                                        End If

                                    End If
                                    End If

                                End If
                                If pVal.ItemUID = "bt_FumiNew" Then
                                    FumigationForm.Freeze(True)
                                    If source = "Crane" Then
                                        ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb", "ed_SIns")
                                        'button
                                        oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                        dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                        If dtmatrix.Rows.Count <> 0 Then
                                            For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                                dtmatrix.Rows.Remove(i)
                                            Next
                                        End If
                                        oMatrix.Clear()
                                    RowAddToMatrix(FumigationForm, oMatrix, source)
                                ElseIf source = "Forklift" Then
                                    ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Remark", "ed_IRemark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                    'button
                                    oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                    dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                    If dtmatrix.Rows.Count <> 0 Then
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                    End If
                                    oMatrix.Clear()
                                    RowAddToMatrix(FumigationForm, oMatrix, source)
                                ElseIf source = "Crate" Then
                                    ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Desc", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb")
                                    'button
                                    oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                    dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                    If dtmatrix.Rows.Count <> 0 Then
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                    End If
                                    oMatrix.Clear()
                                    RowAddToMatrix(FumigationForm, oMatrix, source)
                                ElseIf source = "Bunker" Then 'Bunker
                                    ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb", "ed_CDesc", "ed_SIns", "ed_TQty", "ed_TKgs", "ed_TM3", "ed_TNEQ")
                                    FumigationForm.Items.Item("chk_1").Specific.Checked = False
                                    FumigationForm.Items.Item("chk_1a").Specific.Checked = False
                                    FumigationForm.Items.Item("chk_2").Specific.Checked = False
                                    FumigationForm.Items.Item("chk_2a").Specific.Checked = False
                                    'button
                                    oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                    dtmatrix = FumigationForm.DataSources.DataTables.Item("dtBDetail")
                                    If dtmatrix.Rows.Count <> 0 Then
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                    End If
                                    oMatrix.Clear()
                                    RowAddToMatrixBunker(FumigationForm, oMatrix)
                                ElseIf source = "Toll" Then 'Bunker
                                    ClearText(FumigationForm, "ed_FCode", "ed_FName", "ed_FSIA", "ed_PO", "ed_SIACode", "ed_PODocNo", "ed_PONo", "ed_SIATel", "ed_Loc", "ed_Remark", "ed_FVRef", "ed_FCntact", "ed_FJbDate", "ed_FJbTime", "ed_OriPO", "ed_MultiJb", "ed_IRemark")

                                    'button
                                    oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                    dtmatrix = FumigationForm.DataSources.DataTables.Item("dtTDetail")
                                    If dtmatrix.Rows.Count <> 0 Then
                                        For i As Integer = dtmatrix.Rows.Count - 1 To 0 Step -1
                                            dtmatrix.Rows.Remove(i)
                                        Next
                                    End If
                                    oMatrix.Clear()
                                    RowAddToMatrixToll(FumigationForm, oMatrix)
                                End If
                                rowIndex = -1
                                FumigationForm.Items.Item("ed_FPODate").Specific.Value = Today.Date.ToString("yyyyMMdd")
                                FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add"

                                FumigationForm.Items.Item("bt_CrPO").Enabled = True
                                FumigationForm.Items.Item("ed_FCode").Enabled = True
                                FumigationForm.Items.Item("ed_FName").Enabled = True
                                FumigationForm.Items.Item("ed_FCode").Specific.Active = True 'button
                                FumigationForm.Freeze(False)
                            End If


                                If pVal.ItemUID = "bt_FumiAdd" Then
                                    Dim DocEntry As Integer
                                    If String.IsNullOrEmpty(FumigationForm.Items.Item("ed_FCode").Specific.String) Or String.IsNullOrEmpty(FumigationForm.Items.Item("ed_FName").Specific.String) Then
                                        p_oSBOApplication.SetStatusBarMessage("Must Fill Vendor", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    ElseIf String.IsNullOrEmpty(FumigationForm.Items.Item("ed_PO").Specific.String) Then
                                        p_oSBOApplication.SetStatusBarMessage("Must Fill PO", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        BubbleEvent = False
                                    Else
                                        oMatrix = FumigationForm.Items.Item(matrixName).Specific
                                        poMatrix = ExportSeaFCLForm.Items.Item("mx_PO").Specific
                                        If FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Add" Then
                                            AddUpdateFumigation(FumigationForm, oMatrix, tableName, True, source)
                                            SaveToPOTab(ExportSeaFCLForm, poMatrix, True, FumigationForm.Items.Item("ed_PO").Specific.Value, FumigationForm.Items.Item("ed_PODocNo").Specific.Value, FumigationForm.Items.Item("ed_FName").Specific.Value, FumigationForm.Items.Item("ed_FPODate").Specific.Value, source, FumigationForm.Items.Item("ed_FPOStus").Specific.Value)

                                            If source = "Crane" Then
                                                oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                                oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [" & detailtableName & "]")
                                                If oRecordSet.RecordCount > 0 Then
                                                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                                End If
                                                If oMatrix.RowCount > 0 Then  'button
                                                    For i As Integer = 1 To oMatrix.RowCount
                                                        sql = "Insert Into [" & detailtableName & "] (DocEntry,LineID,U_CType,U_Ton,U_Desc,U_Qty,U_UOM,U_Price,U_Hrs,U_Total,U_Remark,U_PO,U_DocN) Values " & _
                                                               "(" & DocEntry & _
                                                                "," & i & _
                                                                "," & IIf(oMatrix.Columns.Item("colCType").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colCType").Cells.Item(i).Specific.Value), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value), "Null") & _
                                                                "," & IIf(oMatrix.Columns.Item("colDesc").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colDesc").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colUOM").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colUOM").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colHr").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colHr").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colRmk").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colRmk").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(FumigationForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                                                "," & IIf(FumigationForm.Items.Item("ed_DocNum").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_DocNum").Specific.Value), "NULL") & ")"
                                                        oRecordSet.DoQuery(sql)

                                                    Next
                                            End If
                                        ElseIf source = "Forklift" Then
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [" & detailtableName & "]")
                                            If oRecordSet.RecordCount > 0 Then
                                                DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                            End If
                                            If oMatrix.RowCount > 0 Then  'button
                                                For i As Integer = 1 To oMatrix.RowCount
                                                   sql = "Insert Into [@OBT_TBL07_FORDETAIL] (DocEntry,LineID,U_Ton,U_Desc,U_Qty,U_UOM,U_Price,U_Total,U_Remark,U_PO,U_DocN) Values " & _
                                                                "(" & DocEntry & _
                                                                 "," & i & _
                                                                 "," & IIf(oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colTon").Cells.Item(i).Specific.Value), "Null") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colDesc").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colDesc").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colUOM").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colUOM").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colRmk").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colRmk").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(FumigationForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                                                 "," & IIf(FumigationForm.Items.Item("ed_DocNum").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_DocNum").Specific.Value), "NULL") & ")"
                                                    oRecordSet.DoQuery(sql)

                                                Next
                                            End If
                                        ElseIf source = "Crate" Then
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [" & detailtableName & "]")
                                            If oRecordSet.RecordCount > 0 Then
                                                DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                            End If
                                            If oMatrix.RowCount > 0 Then  'button
                                                For i As Integer = 1 To oMatrix.RowCount
                                                    sql = "Insert Into [@OBT_TBL10_CRADETAIL] (DocEntry,LineID,U_Dimen,U_Type,U_Qty,U_Price,U_Total,U_PO,U_DocN) Values " & _
                                                               "(" & DocEntry & _
                                                                "," & i & _
                                                                "," & IIf(oMatrix.Columns.Item("colDimen").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colDimen").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                 "," & IIf(oMatrix.Columns.Item("colType").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colType").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(FumigationForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                                                "," & IIf(FumigationForm.Items.Item("ed_DocNum").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_DocNum").Specific.Value), "NULL") & ")"
                                                    oRecordSet.DoQuery(sql)

                                                Next
                                            End If
                                        ElseIf source = "Bunker" Then
                                            oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                            oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [" & detailtableName & "]")
                                            If oRecordSet.RecordCount > 0 Then
                                                DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                            End If
                                            If oMatrix.RowCount > 0 Then  'button
                                                For i As Integer = 1 To oMatrix.RowCount
                                                    sql = "Insert Into [" & detailtableName & "] (DocEntry,LineID,U_Permit,U_JobNo,U_Client,U_Qty,U_UOM,U_Kgs,U_M3,U_NEQ,U_Stus,U_PO,U_DocN) Values " & _
                                                           "(" & DocEntry & _
                                                            "," & i & _
                                                            "," & IIf(oMatrix.Columns.Item("colPermit").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colPermit").Cells.Item(i).Specific.Value), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value), "Null") & _
                                                            "," & IIf(oMatrix.Columns.Item("colClient").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colClient").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colUOM").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colUOM").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colKgs").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colKgs").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colM3").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colM3").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colNEQ").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colNEQ").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(oMatrix.Columns.Item("colStus").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colStus").Cells.Item(i).Specific.Value()), "NULL") & _
                                                            "," & IIf(FumigationForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                                            "," & IIf(FumigationForm.Items.Item("ed_DocNum").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_DocNum").Specific.Value), "NULL") & ")"
                                                    oRecordSet.DoQuery(sql)

                                                Next
                                            End If
                                            ElseIf source = "Toll" Then
                                                oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                                oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [" & detailtableName & "]")
                                                If oRecordSet.RecordCount > 0 Then
                                                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                                                End If
                                                If oMatrix.RowCount > 0 Then  'button
                                                    For i As Integer = 1 To oMatrix.RowCount
                                                        sql = "Insert Into [" & detailtableName & "] (DocEntry,LineID,U_ICode,U_IDesc,U_Qty,U_Price,U_Total,U_PO,U_DocN) Values " & _
                                                               "(" & DocEntry & _
                                                                "," & i & _
                                                                "," & IIf(oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value), "Null") & _
                                                                "," & IIf(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value() <> "", FormatString(oMatrix.Columns.Item("colTotal").Cells.Item(i).Specific.Value()), "NULL") & _
                                                                "," & IIf(FumigationForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                                                "," & IIf(FumigationForm.Items.Item("ed_DocNum").Specific.Value <> "", FormatString(FumigationForm.Items.Item("ed_DocNum").Specific.Value), "NULL") & ")"
                                                        oRecordSet.DoQuery(sql)

                                                    Next
                                                End If
                                            End If

                                        Else
                                            AddUpdateFumigation(FumigationForm, oMatrix, tableName, False, source)
                                            ' If Not CreateFFCPO(FumigationForm, False, "Crane") Then Throw New ArgumentException(sErrDesc)
                                            EditPOInEditTab(FumigationForm, FumigationForm.Items.Item("ed_PONo").Specific.Value, matrixName) '
                                            ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                        If FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                        FumigationForm.Items.Item("1").Click()
                                        FumigationForm.Items.Item("2").Specific.Caption = "Close"
                                        FumigationForm.Items.Item("bt_CrPO").Enabled = False
                                        ExportSeaFCLForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ExportSeaFCLForm.Items.Item("1").Click()
                                        FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Update"

                                    End If
                                End If
                                If pVal.ItemUID = "bt_PreView" Then
                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                                    If selectedRow > 0 Then
                                    If source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                                        oMatrix = FumigationForm.Items.Item(matrixName).Specific
                                        PreviewPOInViewList(ExportSeaFCLForm, selectedRow, oMatrix)
                                    ElseIf source = "Bunker" Or source = "Toll" Then
                                        CreatePOPDF(FumigationForm, matrixName, "View", selectedRow)
                                    End If

                                    Else
                                        p_oSBOApplication.MessageBox("Need to select one Row to Preview.")
                                    End If
                                End If
                                If pVal.ItemUID = "bt_IView" Then

                                    ExportSeaFCLForm = p_oSBOApplication.Forms.GetForm("EXPORTSEAFCL", 1)
                                    If source = "Crane" Then
                                        PreviewPO(ExportSeaFCLForm, FumigationForm)
                                    ElseIf source = "Bunker" Then
                                        CreatePOPDF(FumigationForm, "mx_Bunk", "Edit")
                                    ElseIf source = "Toll" Then
                                        CreatePOPDF(FumigationForm, "mx_Toll", "Edit")
                                    End If

                                End If
                                '''' TO Do
                                If pVal.ItemUID = "bt_MJob" Then
                                    modExportSeaFCL.LoadMultiJobNormalForm(FumigationForm, "MultiJobNormal.srf", source)
                                    oActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    modMultiJobForNormal.LoadMultiJForm()
                                End If
                                If pVal.ItemUID = "bt_DJob" Then
                                    LoadDetachJobNormalForm(FumigationForm, "DetachJobNormal.srf", source, "ed_PO")
                                    oActiveForm = p_oSBOApplication.Forms.ActiveForm
                                    modDetachJobNormal.LoadDetachJForm()
                                End If
                                If pVal.ItemUID = "bt_Email" Then
                                    modExportSeaFCL.PreviewOutLookMail(FumigationForm, source)
                                End If

                                If pVal.ItemUID = "1" Then
                                    If pVal.Action_Success = True Then
                                        If pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        If source = "Crane" Then
                                            FumigationForm = p_oSBOApplication.Forms.GetForm("CRANE", 1)
                                        ElseIf source = "Forklift" Then 'to combine
                                            FumigationForm = p_oSBOApplication.Forms.GetForm("FORKLIFT", 1)
                                        ElseIf source = "Crate" Then 'to combine
                                            FumigationForm = p_oSBOApplication.Forms.GetForm("CRATE", 1)
                                        ElseIf source = "Bunker" Then
                                            FumigationForm = p_oSBOApplication.Forms.GetForm("BUNKER", 1)
                                        ElseIf source = "Toll" Then
                                            FumigationForm = p_oSBOApplication.Forms.GetForm("TOLL", 1)
                                        End If
                                            FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
                                            p_oSBOApplication.ActivateMenuItem("1291")
                                            FumigationForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            FumigationForm.Items.Item("2").Specific.Caption = "Close" 'MSW 09-09-2011
                                        End If



                                    End If
                                End If
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = False Then
                            If pVal.ItemUID = "mx_Crane" Or pVal.ItemUID = "mx_Bunk" Or pVal.ItemUID = "mx_Fork" Or pVal.ItemUID = "mx_Crate" Or pVal.ItemUID = "mx_Toll" Then
                                oMatrix = FumigationForm.Items.Item(pVal.ItemUID).Specific
                                If pVal.Row > 0 Then
                                    oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                    oMatrix.SelectRow(pVal.Row, True, False)
                                    selectedRow = pVal.Row
                                End If
                            End If
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK And pVal.BeforeAction = False Then
                                If pVal.ItemUID = matrixName Then
                                    If pVal.Row > 0 Then
                                        oMatrix = FumigationForm.Items.Item(matrixName).Specific
                                        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                                        oMatrix.SelectRow(pVal.Row, True, False)

                                        If oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value() = "Closed" Or oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value() = "Cancelled" Then
                                            p_oSBOApplication.SetStatusBarMessage("This PO is already " & oMatrix.Columns.Item("colPStatus").Cells.Item(pVal.Row).Specific.Value & ".", SAPbouiCOM.BoMessageTime.bmt_Short)
                                        Else
                                            GetDataFromMatrixByIndex(FumigationForm, oMatrix, pVal.Row, source)
                                            'To Add Item Code in Select Statement

                                        If source = "Crane" Then
                                            sql = "Select LineId as LineId,U_CType as CType,U_Ton as Ton,U_Desc as [Desc],U_Qty as Qty,U_UOM As UOM,U_Price As Price,U_Hrs as Hrs,U_Total as Total,U_Remark as Remark From [" & detailtableName & "] " & _
                                            " Where U_DocN = '" & FumigationForm.Items.Item("ed_DocNum").Specific.Value & "' And U_PO = '" & oMatrix.Columns.Item("colPO").Cells.Item(pVal.Row).Specific.Value() & "' Order By LineId"

                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            dtmatrix.ExecuteQuery(sql)
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            oMatrix.LoadFromDataSource()
                                        ElseIf source = "Forklift" Then 'to combine
                                            sql = "Select LineId as LineId,U_Ton as Ton,U_Desc as [Desc],U_Qty as Qty,U_UOM As UOM,U_Price As Price,U_Total as Total,U_Remark as Remark From [@OBT_TBL07_FORDETAIL] " & _
                                                    " Where U_DocN = '" & FumigationForm.Items.Item("ed_DocNum").Specific.Value & "' And U_PO = '" & oMatrix.Columns.Item("colPO").Cells.Item(pVal.Row).Specific.Value & "' Order By LineId"
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            dtmatrix.ExecuteQuery(sql)
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            oMatrix.LoadFromDataSource()
                                        ElseIf source = "Crate" Then 'to combine
                                            sql = "Select LineId as LineId,U_Dimen as [Dimension],U_Type as [Type],U_Qty as Qty,U_Price As Price,U_Total as [Total] From [@OBT_TBL10_CRADETAIL] " & _
                                                    " Where U_DocN = '" & FumigationForm.Items.Item("ed_DocNum").Specific.Value & "' And U_PO = '" & oMatrix.Columns.Item("colPO").Cells.Item(pVal.Row).Specific.Value & "' Order By LineId"
                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtCDetail")
                                            dtmatrix.ExecuteQuery(sql)
                                            oMatrix = FumigationForm.Items.Item("mx_CDetail").Specific
                                            oMatrix.LoadFromDataSource()
                                        ElseIf source = "Bunker" Then
                                            sql = "Select LineId as LineId,U_Permit as Permit,U_JobNo as JobNo,U_Client as Client,U_Qty as Qty,U_UOM As UOM,U_Kgs As Kgs,U_M3 as M3,U_NEQ as NEQ,U_Stus as Stus From [" & detailtableName & "] " & _
                                            " Where U_DocN = '" & FumigationForm.Items.Item("ed_DocNum").Specific.Value & "' And U_PO = '" & oMatrix.Columns.Item("colPO").Cells.Item(pVal.Row).Specific.Value() & "' Order By LineId"

                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtBDetail")
                                            dtmatrix.ExecuteQuery(sql)
                                            oMatrix = FumigationForm.Items.Item("mx_BDetail").Specific
                                            oMatrix.LoadFromDataSource()
                                        ElseIf source = "Toll" Then
                                            sql = "Select LineId as LineId,U_ICode as ICode,U_IDesc as IDesc,U_Qty as Qty,U_Price as Price,U_Total as Total From [" & detailtableName & "] " & _
                                            " Where U_DocN = '" & FumigationForm.Items.Item("ed_DocNum").Specific.Value & "' And U_PO = '" & oMatrix.Columns.Item("colPO").Cells.Item(pVal.Row).Specific.Value() & "' Order By LineId"

                                            dtmatrix = FumigationForm.DataSources.DataTables.Item("dtTDetail")
                                            dtmatrix.ExecuteQuery(sql)
                                            oMatrix = FumigationForm.Items.Item("mx_TDetail").Specific
                                            oMatrix.LoadFromDataSource()
                                        End If


                                            FumigationForm.Items.Item("bt_FumiAdd").Specific.Caption = "Update"
                                            FumigationForm.Items.Item("fo_Edit").Specific.Select()
                                            FumigationForm.Items.Item("bt_CrPO").Enabled = False
                                            FumigationForm.Items.Item("ed_FCode").Enabled = False
                                            FumigationForm.Items.Item("ed_FName").Enabled = False
                                            If FumigationForm.Items.Item("ed_OriPO").Specific.Value = "Y" Then
                                                FumigationForm.Items.Item("bt_DJob").Enabled = True
                                            Else
                                                FumigationForm.Items.Item("bt_DJob").Enabled = False
                                            End If
                                        End If

                                    End If

                                End If
                            End If
                        End If

            End Select
            DoExportSeaFCLItemEvent = RTN_SUCCESS
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            DoExportSeaFCLItemEvent = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", FunctionName)
            WriteToLogFile(Err.Description, FunctionName)
        Finally
            GC.Collect()
        End Try
      

      
    End Function

    Public Sub GetDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer, ByVal source As String)
        Dim oCombo As SAPbouiCOM.ComboBox
        Try
            pForm.Items.Item("ed_FCode").Specific.Value = pMatrix.Columns.Item("colVCode").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FName").Specific.Value = pMatrix.Columns.Item("colVendor").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FCntact").Specific.Value = pMatrix.Columns.Item("colCPerson").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FVRef").Specific.Value = pMatrix.Columns.Item("colVRef").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_SIACode").Specific.Value = pMatrix.Columns.Item("colSIACode").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_JDeNo").Specific.Value = pMatrix.Columns.Item("colJobNo").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FSIA").Specific.Value = pMatrix.Columns.Item("colSIA").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_SIATel").Specific.Value = pMatrix.Columns.Item("colTelNo").Cells.Item(Index).Specific.Value
            If source = "Fumigation" Then
                pForm.Items.Item("ed_Loc").Specific.Value = pMatrix.Columns.Item("colLoc").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_Item").Specific.Value = pMatrix.Columns.Item("colItem").Cells.Item(Index).Specific.Value
            ElseIf source = "Outrider" Then
                pForm.Items.Item("ed_LocFrom").Specific.Value = pMatrix.Columns.Item("colLocFrom").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_LocTo").Specific.Value = pMatrix.Columns.Item("colLocTo").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_IRemark").Specific.Value = pMatrix.Columns.Item("colIRemark").Cells.Item(Index).Specific.Value
            ElseIf source = "Crane" Then
                pForm.Items.Item("ed_Loc").Specific.Value = pMatrix.Columns.Item("colLoc").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_SIns").Specific.Value = pMatrix.Columns.Item("colSIns").Cells.Item(Index).Specific.Value
            ElseIf source = "Forklift" Then 'to combine
                pForm.Items.Item("ed_Loc").Specific.Value = pMatrix.Columns.Item("colLoc").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_IRemark").Specific.Value = pMatrix.Columns.Item("colIRemark").Cells.Item(Index).Specific.Value
            ElseIf source = "Crate" Then 'to combine
                pForm.Items.Item("ed_Desc").Specific.Value = pMatrix.Columns.Item("colDesc").Cells.Item(Index).Specific.Value
            ElseIf source = "Bunker" Then
                oCombo = pForm.Items.Item("cb_Act").Specific
                oCombo.Select(pMatrix.Columns.Item("colActive").Cells.Item(Index).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
                pForm.Items.Item("ed_CDesc").Specific.Value = pMatrix.Columns.Item("colCDesc").Cells.Item(Index).Specific.Value
                If pMatrix.Columns.Item("colS1").Cells.Item(Index).Specific.Value <> "" Then
                    pForm.DataSources.UserDataSources.Item("STORE1").ValueEx = "Y"
                Else
                    pForm.DataSources.UserDataSources.Item("STORE1").ValueEx = ""
                End If
                If pMatrix.Columns.Item("colS1a").Cells.Item(Index).Specific.Value <> "" Then
                    pForm.DataSources.UserDataSources.Item("STORE1a").ValueEx = "Y"
                Else
                    pForm.DataSources.UserDataSources.Item("STORE1a").ValueEx = ""
                End If
                If pMatrix.Columns.Item("colS2").Cells.Item(Index).Specific.Value <> "" Then
                    pForm.DataSources.UserDataSources.Item("STORE2").ValueEx = "Y"
                Else
                    pForm.DataSources.UserDataSources.Item("STORE2").ValueEx = ""
                End If
                If pMatrix.Columns.Item("colS2a").Cells.Item(Index).Specific.Value <> "" Then
                    pForm.DataSources.UserDataSources.Item("STORE2a").ValueEx = "Y"
                Else
                    pForm.DataSources.UserDataSources.Item("STORE2a").ValueEx = ""
                End If
                pForm.Items.Item("ed_TQty").Specific.Value = pMatrix.Columns.Item("colTQty").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_TKgs").Specific.Value = pMatrix.Columns.Item("colTKgs").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_TM3").Specific.Value = pMatrix.Columns.Item("colTM3").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_TNEQ").Specific.Value = pMatrix.Columns.Item("colTNEQ").Cells.Item(Index).Specific.Value
            ElseIf source = "Toll" Then
               pForm.Items.Item("ed_Loc").Specific.Value = pMatrix.Columns.Item("colLoc").Cells.Item(Index).Specific.Value
                pForm.Items.Item("ed_IRemark").Specific.Value = pMatrix.Columns.Item("colIRemark").Cells.Item(Index).Specific.Value
            End If
            pForm.Items.Item("ed_FJbDate").Specific.Value = pMatrix.Columns.Item("colDate").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FJbTime").Specific.Value = pMatrix.Columns.Item("colTime").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PO").Specific.Value = pMatrix.Columns.Item("colPO").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FPOStus").Specific.Value = pMatrix.Columns.Item("colPStatus").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Create").Specific.Value = pMatrix.Columns.Item("colPrepare").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_FPODate").Specific.Value = pMatrix.Columns.Item("colPODate").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PODocNo").Specific.Value = pMatrix.Columns.Item("colDocNo").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_PONo").Specific.Value = pMatrix.Columns.Item("colPONo").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_Remark").Specific.Value = pMatrix.Columns.Item("colRemark").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_OriPO").Specific.Value = pMatrix.Columns.Item("colOriPO").Cells.Item(Index).Specific.Value
            pForm.Items.Item("ed_MultiJb").Specific.Value = pMatrix.Columns.Item("colMultiJb").Cells.Item(Index).Specific.Value
        Catch ex As Exception

        End Try
    End Sub

    Public Sub AddUpdateFumigation(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean, ByVal source As String)

        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        rowIndex = pMatrix.GetNextSelectedRow
        'If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
        '    rowIndex = 1
        'End If
        If rowIndex = -1 And pForm.Items.Item("ed_PO").Specific.Value <> "" Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT LineId FROM [" & DataSource & "] Where U_PO=" & pForm.Items.Item("ed_PO").Specific.Value)
            If oRecordSet.RecordCount > 0 Then
                rowIndex = oRecordSet.Fields.Item("LineId").Value
            End If
        End If
        Try
            If ProcressedState = True Then

                With ObjDBDataSource

                    If ObjDBDataSource.GetValue("U_PO", 0) = vbNullString Then pMatrix.Clear()
                    If pMatrix.RowCount = 0 Then
                        .SetValue("LineId", .Offset, 1)
                    Else
                        .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(pMatrix.RowCount).Specific.Value + 1)
                    End If
                    .SetValue("U_Vendor", .Offset, pForm.Items.Item("ed_FCode").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_SIA", .Offset, pForm.Items.Item("ed_FSIA").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value)
                    .SetValue("U_Status", .Offset, pForm.Items.Item("ed_FPOStus").Specific.Value)
                    .SetValue("U_CreateBy", .Offset, pForm.Items.Item("ed_Create").Specific.Value)
                    .SetValue("U_JobNo", .Offset, pForm.Items.Item("ed_JDeNo").Specific.Value)
                    .SetValue("U_SIACode", .Offset, pForm.Items.Item("ed_SIACode").Specific.Value)
                    .SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_PODocNo").Specific.Value)
                    .SetValue("U_PONo", .Offset, pForm.Items.Item("ed_PONo").Specific.Value)
                    .SetValue("U_Vname", .Offset, pForm.Items.Item("ed_FName").Specific.Value)
                    .SetValue("U_TelNo", .Offset, pForm.Items.Item("ed_SIATel").Specific.Value)
                    If source = "Fumigation" Then 'button
                        .SetValue("U_IFumi", .Offset, pForm.Items.Item("ed_Item").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    ElseIf source = "Outrider" Then
                        .SetValue("U_IRemark", .Offset, pForm.Items.Item("ed_IRemark").Specific.Value)
                        .SetValue("U_LocFrom", .Offset, pForm.Items.Item("ed_LocFrom").Specific.Value)
                        .SetValue("U_LocTo", .Offset, pForm.Items.Item("ed_LocTo").Specific.Value)
                    ElseIf source = "Crane" Then
                        .SetValue("U_SIns", .Offset, pForm.Items.Item("ed_SIns").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    ElseIf source = "Forklift" Then ' to combine
                        .SetValue("U_IRemark", .Offset, pForm.Items.Item("ed_IRemark").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    ElseIf source = "Crate" Then ' to combine
                        .SetValue("U_Desc", .Offset, pForm.Items.Item("ed_Desc").Specific.Value)
                    ElseIf source = "Bunker" Then
                        .SetValue("U_SIns", .Offset, pForm.Items.Item("ed_SIns").Specific.Value)
                        .SetValue("U_CDesc", .Offset, pForm.Items.Item("ed_CDesc").Specific.Value)
                        .SetValue("U_Active", .Offset, pForm.Items.Item("cb_Act").Specific.Value.ToString.Trim)
                        .SetValue("U_TQty", .Offset, pForm.Items.Item("ed_TQty").Specific.Value)
                        .SetValue("U_TKgs", .Offset, pForm.Items.Item("ed_TKgs").Specific.Value)
                        .SetValue("U_TM3", .Offset, pForm.Items.Item("ed_TM3").Specific.Value)
                        .SetValue("U_TNEQ", .Offset, pForm.Items.Item("ed_TNEQ").Specific.Value)
                        If pForm.Items.Item("chk_1").Specific.Checked = True Then
                            .SetValue("U_Store1", .Offset, "Y")
                        Else
                            .SetValue("U_Store1", .Offset, "")
                        End If
                        If pForm.Items.Item("chk_1a").Specific.Checked = True Then
                            .SetValue("U_Store1a", .Offset, "Y")
                        Else
                            .SetValue("U_Store1a", .Offset, "")
                        End If
                        If pForm.Items.Item("chk_2").Specific.Checked = True Then
                            .SetValue("U_Store2", .Offset, "Y")
                        Else
                            .SetValue("U_Store2", .Offset, "")
                        End If
                        If pForm.Items.Item("chk_2a").Specific.Checked = True Then
                            .SetValue("U_Store2a", .Offset, "Y")
                        Else
                            .SetValue("U_Store2a", .Offset, "")
                        End If
                    ElseIf source = "Toll" Then
                        .SetValue("U_IRemark", .Offset, pForm.Items.Item("ed_IRemark").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    End If
                    .SetValue("U_Remark", .Offset, pForm.Items.Item("ed_Remark").Specific.Value)
                    .SetValue("U_VRef", .Offset, pForm.Items.Item("ed_FVRef").Specific.Value)
                    .SetValue("U_CPerson", .Offset, pForm.Items.Item("ed_FCntact").Specific.Value)
                    .SetValue("U_JDate", .Offset, pForm.Items.Item("ed_FJbDate").Specific.Value)
                    .SetValue("U_JTime", .Offset, pForm.Items.Item("ed_FJbTime").Specific.Value)
                    .SetValue("U_PODate", .Offset, pForm.Items.Item("ed_FPODate").Specific.Value)
                    .SetValue("U_OriginPO", .Offset, pForm.Items.Item("ed_OriPO").Specific.Value)
                    .SetValue("U_MultiJob", .Offset, pForm.Items.Item("ed_MultiJb").Specific.Value)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource

                    .SetValue("LineId", .Offset, pMatrix.Columns.Item("V_-1").Cells.Item(rowIndex).Specific.Value)
                    .SetValue("U_Vendor", .Offset, pForm.Items.Item("ed_FCode").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_SIA", .Offset, pForm.Items.Item("ed_FSIA").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value)
                    .SetValue("U_Status", .Offset, pForm.Items.Item("ed_FPOStus").Specific.Value)
                    .SetValue("U_CreateBy", .Offset, pForm.Items.Item("ed_Create").Specific.Value)
                    .SetValue("U_JobNo", .Offset, pForm.Items.Item("ed_JDeNo").Specific.Value)
                    .SetValue("U_SIACode", .Offset, pForm.Items.Item("ed_SIACode").Specific.Value)
                    .SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_PODocNo").Specific.Value)
                    .SetValue("U_PONo", .Offset, pForm.Items.Item("ed_PONo").Specific.Value)
                    .SetValue("U_Vname", .Offset, pForm.Items.Item("ed_FName").Specific.Value)
                    .SetValue("U_TelNo", .Offset, pForm.Items.Item("ed_SIATel").Specific.Value)
                    If source = "Fumigation" Then 'button
                        .SetValue("U_IFumi", .Offset, pForm.Items.Item("ed_Item").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    ElseIf source = "Outrider" Then
                        .SetValue("U_IRemark", .Offset, pForm.Items.Item("ed_IRemark").Specific.Value)
                        .SetValue("U_LocFrom", .Offset, pForm.Items.Item("ed_LocFrom").Specific.Value)
                        .SetValue("U_LocTo", .Offset, pForm.Items.Item("ed_LocTo").Specific.Value)
                    ElseIf source = "Crane" Then
                        .SetValue("U_SIns", .Offset, pForm.Items.Item("ed_SIns").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    ElseIf source = "Forklift" Then ' to combine
                        .SetValue("U_IRemark", .Offset, pForm.Items.Item("ed_IRemark").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    ElseIf source = "Crate" Then ' to combine
                        .SetValue("U_Desc", .Offset, pForm.Items.Item("ed_Desc").Specific.Value)
                    ElseIf source = "Bunker" Then
                        .SetValue("U_SIns", .Offset, pForm.Items.Item("ed_SIns").Specific.Value)
                        .SetValue("U_CDesc", .Offset, pForm.Items.Item("ed_CDesc").Specific.Value)
                        .SetValue("U_Active", .Offset, pForm.Items.Item("cb_Act").Specific.Value.ToString.Trim)
                        .SetValue("U_TQty", .Offset, pForm.Items.Item("ed_TQty").Specific.Value)
                        .SetValue("U_TKgs", .Offset, pForm.Items.Item("ed_TKgs").Specific.Value)
                        .SetValue("U_TM3", .Offset, pForm.Items.Item("ed_TM3").Specific.Value)
                        .SetValue("U_TNEQ", .Offset, pForm.Items.Item("ed_TNEQ").Specific.Value)
                        If pForm.Items.Item("chk_1").Specific.Checked = True Then
                            .SetValue("U_Store1", .Offset, "Y")
                        Else
                            .SetValue("U_Store1", .Offset, "")
                        End If
                        If pForm.Items.Item("chk_1a").Specific.Checked = True Then
                            .SetValue("U_Store1a", .Offset, "Y")
                        Else
                            .SetValue("U_Store1a", .Offset, "")
                        End If
                        If pForm.Items.Item("chk_2").Specific.Checked = True Then
                            .SetValue("U_Store2", .Offset, "Y")
                        Else
                            .SetValue("U_Store2", .Offset, "")
                        End If
                        If pForm.Items.Item("chk_2a").Specific.Checked = True Then
                            .SetValue("U_Store2a", .Offset, "Y")
                        Else
                            .SetValue("U_Store2a", .Offset, "")
                        End If
                    ElseIf source = "Toll" Then
                        .SetValue("U_IRemark", .Offset, pForm.Items.Item("ed_IRemark").Specific.Value)
                        .SetValue("U_Loc", .Offset, pForm.Items.Item("ed_Loc").Specific.Value)
                    End If
                    .SetValue("U_Remark", .Offset, pForm.Items.Item("ed_Remark").Specific.Value)
                    .SetValue("U_VRef", .Offset, pForm.Items.Item("ed_FVRef").Specific.Value)
                    .SetValue("U_CPerson", .Offset, pForm.Items.Item("ed_FCntact").Specific.Value)
                    .SetValue("U_JDate", .Offset, pForm.Items.Item("ed_FJbDate").Specific.Value)
                    .SetValue("U_JTime", .Offset, pForm.Items.Item("ed_FJbTime").Specific.Value)
                    .SetValue("U_PODate", .Offset, pForm.Items.Item("ed_FPODate").Specific.Value)
                    .SetValue("U_OriginPO", .Offset, pForm.Items.Item("ed_OriPO").Specific.Value)
                    .SetValue("U_MultiJob", .Offset, pForm.Items.Item("ed_MultiJb").Specific.Value)
                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try

    End Sub

#Region "Create PO"
    Public Function CreateGenPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal VCode As String, ByVal iCode As String, ByVal source As String) As Boolean
        Dim sErrDesc As String = ""
        CreateGenPO = False
        Try


            If Not CreatePO(oActiveForm, VCode, iCode, source) Then Throw New ArgumentException(sErrDesc)
            If Not CreateFFCPO(oActiveForm, True, source) Then Throw New ArgumentException(sErrDesc)
            oActiveForm.Items.Item("ed_OriPO").Specific.Value = "Y"

            CreateGenPO = True
        Catch ex As Exception
            CreateGenPO = False
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Public Function CreatePO(ByRef oActiveForm As SAPbouiCOM.Form, ByVal CardCode As String, ByVal iCode As String, ByVal source As String) As Boolean

        Dim oPurchaseDocument As SAPbobsCOM.Documents
        Dim oBusinessPartner As SAPbobsCOM.BusinessPartners
        Dim oRecordset As SAPbobsCOM.Recordset
        Dim dblPrice As Double = 0.0
        Dim sErrDesc As String = vbNullString
        Dim oMatrix As SAPbouiCOM.Matrix
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
                    If source = "Fumigation" Or source = "Outrider" Then
                        oPurchaseDocument.Lines.ItemCode = oActiveForm.Items.Item("ed_ICode").Specific.Value
                        oPurchaseDocument.Lines.ItemDescription = oActiveForm.Items.Item("ed_IDesc").Specific.Value
                        oPurchaseDocument.Lines.Quantity = oActiveForm.Items.Item("ed_IQty").Specific.Value
                        oPurchaseDocument.Lines.UnitPrice = oActiveForm.Items.Item("ed_IPrice").Specific.Value
                        oPurchaseDocument.Lines.VatGroup = "XI"
                        oPurchaseDocument.Lines.Add()
                    ElseIf source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                        oMatrix = oActiveForm.Items.Item("mx_CDetail").Specific
                        If oMatrix.RowCount > 0 Then
                            For i As Integer = 1 To oMatrix.RowCount
                                oPurchaseDocument.Lines.ItemCode = oActiveForm.Items.Item("ed_ICode").Specific.Value
                                oPurchaseDocument.Lines.ItemDescription = oActiveForm.Items.Item("ed_IDesc").Specific.Value
                                oPurchaseDocument.Lines.Quantity = oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.UnitPrice = oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.VatGroup = "XI"
                                oPurchaseDocument.Lines.Add()
                            Next
                        End If
                    ElseIf source = "Bunker" Then
                        oMatrix = oActiveForm.Items.Item("mx_BDetail").Specific
                        If oMatrix.RowCount > 0 Then
                            For i As Integer = 1 To oMatrix.RowCount
                                oPurchaseDocument.Lines.ItemCode = oActiveForm.Items.Item("ed_ICode").Specific.Value
                                oPurchaseDocument.Lines.ItemDescription = oActiveForm.Items.Item("ed_IDesc").Specific.Value
                                oPurchaseDocument.Lines.Quantity = oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.UnitPrice = oActiveForm.Items.Item("ed_IPrice").Specific.Value
                                oPurchaseDocument.Lines.VatGroup = "XI"
                                oPurchaseDocument.Lines.Add()
                            Next
                        End If
                    ElseIf source = "Toll" Then
                        oMatrix = oActiveForm.Items.Item("mx_TDetail").Specific
                        If oMatrix.RowCount > 0 Then
                            For i As Integer = 1 To oMatrix.RowCount
                                oPurchaseDocument.Lines.ItemCode = oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.ItemDescription = oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.Quantity = oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.UnitPrice = oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value
                                oPurchaseDocument.Lines.VatGroup = "XI"
                                oPurchaseDocument.Lines.Add()
                            Next
                        End If
                    End If
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
                    oActiveForm.Items.Item("ed_PONo").Specific.Value = oRecordset.Fields.Item("DocEntry").Value
                    oActiveForm.Items.Item("ed_PO").Specific.Value = oRecordset.Fields.Item("DocNum").Value
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

    Private Function CreateFFCPO(ByVal oActiveForm As SAPbouiCOM.Form, ByVal ProcessState As Boolean, ByVal source As String) As Boolean
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim dblPrice As Double
        Dim oMatrix As SAPbouiCOM.Matrix
        CreateFFCPO = False
        Try
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGeneralService = p_oDICompany.GetCompanyService.GetGeneralService("FCPO")

            If ProcessState = True Then
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                ' dblPrice = Convert.ToDouble(oActiveForm.Items.Item("ed_IQty").Specific.Value) * Convert.ToDouble(oActiveForm.Items.Item("ed_IPrice").Specific.Value)
                oGeneralData.SetProperty("U_VCode", oActiveForm.Items.Item("ed_FCode").Specific.Value)
                oGeneralData.SetProperty("U_VName", oActiveForm.Items.Item("ed_FName").Specific.Value)
                oGeneralData.SetProperty("U_CPerson", oActiveForm.Items.Item("ed_FCntact").Specific.Value)
                oGeneralData.SetProperty("U_VRef", oActiveForm.Items.Item("ed_FVRef").Specific.Value)
                oGeneralData.SetProperty("U_SInA", oActiveForm.Items.Item("ed_FSIA").Specific.Value)
                oGeneralData.SetProperty("U_TDate", Today.Date)
                oGeneralData.SetProperty("U_TTime", Convert.ToDateTime(Now.ToString("hh:mm")))
                oGeneralData.SetProperty("U_TDay", Today.DayOfWeek.ToString.Substring(0, 3))
                'oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_Loc").Specific.Value)

                If source = "Fumigation" Or source = "Crane" Then
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_Loc").Specific.Value)
                ElseIf source = "Outrider" Then
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_LocFrom").Specific.Value)
                    oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ed_IRemark").Specific.Value)
                ElseIf source = "Forklift" Then 'to combine
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_Loc").Specific.Value)
                    oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ed_IRemark").Specific.Value)
                ElseIf source = "Toll" Then
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_Loc").Specific.Value)
                    oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ed_IRemark").Specific.Value)
                End If

                oGeneralData.SetProperty("U_PONo", oActiveForm.Items.Item("ed_PONo").Specific.Value)
                oGeneralData.SetProperty("U_PODate", Today.Date)
                oGeneralData.SetProperty("U_PORMKS", oActiveForm.Items.Item("ed_Remark").Specific.Value)
                oGeneralData.SetProperty("U_POITPD", dblPrice)
                oGeneralData.SetProperty("U_POStatus", oActiveForm.Items.Item("ed_FPOStus").Specific.Value)
                oGeneralData.SetProperty("U_PO", oActiveForm.Items.Item("ed_PO").Specific.Value)

                oChildren = oGeneralData.Child("OBT_TB09_FFCPOITEM")
                If source = "Fumigation" Or source = "Outrider" Then
                    dblPrice = Convert.ToDouble(oActiveForm.Items.Item("ed_IQty").Specific.Value) * Convert.ToDouble(oActiveForm.Items.Item("ed_IPrice").Specific.Value)
                    oChild = oChildren.Add
                    oChild.SetProperty("U_POINO", oActiveForm.Items.Item("ed_ICode").Specific.Value)
                    oChild.SetProperty("U_POIDesc", oActiveForm.Items.Item("ed_IDesc").Specific.Value)
                    oChild.SetProperty("U_POIPrice", oActiveForm.Items.Item("ed_IPrice").Specific.Value)
                    oChild.SetProperty("U_POIQty", oActiveForm.Items.Item("ed_IQty").Specific.Value)
                    oChild.SetProperty("U_POIAmt", dblPrice)
                    oChild.SetProperty("U_POIGST", "XI")
                    oChild.SetProperty("U_POITot", dblPrice)
                ElseIf source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                    oMatrix = oActiveForm.Items.Item("mx_CDetail").Specific
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            oChild = oChildren.Add
                            dblPrice = Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) * Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POINO", oActiveForm.Items.Item("ed_ICode").Specific.Value)
                            oChild.SetProperty("U_POIDesc", oActiveForm.Items.Item("ed_IDesc").Specific.Value)
                            oChild.SetProperty("U_POIPrice", oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIQty", oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIAmt", dblPrice)
                            oChild.SetProperty("U_POIGST", "XI")
                            oChild.SetProperty("U_POITot", dblPrice)
                        Next
                    End If
                ElseIf source = "Bunker" Then
                    oMatrix = oActiveForm.Items.Item("mx_BDetail").Specific
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            oChild = oChildren.Add
                            dblPrice = Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) * Convert.ToDouble(oActiveForm.Items.Item("ed_IPrice").Specific.Value)
                            oChild.SetProperty("U_POINO", oActiveForm.Items.Item("ed_ICode").Specific.Value)
                            oChild.SetProperty("U_POIDesc", oActiveForm.Items.Item("ed_IDesc").Specific.Value)
                            oChild.SetProperty("U_POIPrice", oActiveForm.Items.Item("ed_IPrice").Specific.Value)
                            oChild.SetProperty("U_POIQty", oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIAmt", dblPrice)
                            oChild.SetProperty("U_POIGST", "XI")
                            oChild.SetProperty("U_POITot", dblPrice)
                        Next
                    End If
                ElseIf source = "Toll" Then
                    oMatrix = oActiveForm.Items.Item("mx_TDetail").Specific
                    If oMatrix.RowCount > 0 Then
                        For i As Integer = 1 To oMatrix.RowCount
                            oChild = oChildren.Add
                            dblPrice = Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value) * Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POINO", oMatrix.Columns.Item("colICode").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIDesc", oMatrix.Columns.Item("colIDesc").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIPrice", oMatrix.Columns.Item("colPrice").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIQty", oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value)
                            oChild.SetProperty("U_POIAmt", dblPrice)
                            oChild.SetProperty("U_POIGST", "XI")
                            oChild.SetProperty("U_POITot", dblPrice)
                        Next
                    End If
                End If


                oGeneralService.Add(oGeneralData)
            ElseIf ProcessState = False Then
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", oActiveForm.Items.Item("ed_PODocNo").Specific.Value)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)

                Dim IDesc As String = Replace(oActiveForm.Items.Item("ed_IDesc").Specific.Value, "'", "''")

                dblPrice = Convert.ToDouble(oActiveForm.Items.Item("ed_IQty").Specific.Value) * Convert.ToDouble(oActiveForm.Items.Item("ed_IPrice").Specific.Value)
                oGeneralData.SetProperty("U_VCode", oActiveForm.Items.Item("ed_FCode").Specific.Value)
                oGeneralData.SetProperty("U_VName", oActiveForm.Items.Item("ed_FName").Specific.Value)
                oGeneralData.SetProperty("U_CPerson", oActiveForm.Items.Item("ed_FCntact").Specific.Value)
                oGeneralData.SetProperty("U_VRef", oActiveForm.Items.Item("ed_FVRef").Specific.Value)
                oGeneralData.SetProperty("U_SInA", oActiveForm.Items.Item("ed_FSIA").Specific.Value)
                If oActiveForm.Title.Contains("FUMIGATION") = True Then
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_Loc").Specific.Value)
                ElseIf oActiveForm.Title.Contains("OUTRIDER") = True Then
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_LocFrom").Specific.Value)
                    oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ed_IRemark").Specific.Value)
                ElseIf source = "Forklift" Then 'to combine
                    oGeneralData.SetProperty("U_TPlace", oActiveForm.Items.Item("ed_Loc").Specific.Value)
                    oGeneralData.SetProperty("U_POIRMKS", oActiveForm.Items.Item("ed_IRemark").Specific.Value)
                End If
                oGeneralData.SetProperty("U_PONo", oActiveForm.Items.Item("ed_PONo").Specific.Value)
                oGeneralData.SetProperty("U_PORMKS", oActiveForm.Items.Item("ed_Remark").Specific.Value)
                oGeneralData.SetProperty("U_POITPD", dblPrice)
                oGeneralData.SetProperty("U_POStatus", oActiveForm.Items.Item("ed_FPOStus").Specific.Value)
                oGeneralData.SetProperty("U_PO", oActiveForm.Items.Item("ed_PO").Specific.Value)
                oGeneralService.Update(oGeneralData)
            End If



            oRecordSet.DoQuery("SELECT DocEntry,DocNum FROM [@OBT_TB08_FFCPO] Order By DocEntry")
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveLast()
                oActiveForm.Items.Item("ed_PODocNo").Specific.Value = oRecordSet.Fields.Item("DocEntry").Value
            End If

            CreateFFCPO = True
        Catch ex As Exception
            CreateFFCPO = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

#Region "TableForCrane"
    Public Sub CreateDTDetail(ByVal oActiveForm As SAPbouiCOM.Form, ByVal KeyName As String)

        ObjMatrix = oActiveForm.Items.Item("mx_CDetail").Specific

        If KeyName = "CRATE" Then 'to combine
            CreateTblStructureForCrate(oActiveForm, dtmatrix, "dtCDetail")
        Else
            CreateTblStructure(oActiveForm, dtmatrix, "dtCDetail", KeyName) ' to combine
        End If
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtCDetail")
        dtmatrix.Rows.Add(1)
        dtmatrix.SetValue("LineId", 0, 1)
        ObjMatrix.LoadFromDataSource()
    End Sub
    Private Sub CreateTblStructureForCrate(ByVal oActiveForm As SAPbouiCOM.Form, ByVal dtmatrix As SAPbouiCOM.DataTable, ByVal tblname As String) 'to combine
        Dim oColumn As SAPbouiCOM.Column
        oActiveForm.DataSources.DataTables.Add(tblname)
        dtmatrix = oActiveForm.DataSources.DataTables.Item(tblname)

        dtmatrix.Columns.Add("LineId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Dimension", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Type", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Rate)
        dtmatrix.Columns.Add("Price", SAPbouiCOM.BoFieldsType.ft_Price)
        dtmatrix.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Rate)

        oColumn = ObjMatrix.Columns.Item("V_-1")
        oColumn.DataBind.Bind(tblname, "LineId")


        oColumn = ObjMatrix.Columns.Item("colDimen")
        oColumn.DataBind.Bind(tblname, "Dimension")
        oColumn = ObjMatrix.Columns.Item("colType")
        oColumn.DataBind.Bind(tblname, "Type")
        oColumn = ObjMatrix.Columns.Item("colQty")
        oColumn.DataBind.Bind(tblname, "Qty")
        oColumn = ObjMatrix.Columns.Item("colPrice")
        oColumn.DataBind.Bind(tblname, "Price")
        oColumn = ObjMatrix.Columns.Item("colTotal")
        oColumn.DataBind.Bind(tblname, "Total")


    End Sub
    Public Sub CreateTblStructure(ByVal oActiveForm As SAPbouiCOM.Form, ByVal dtmatrix As SAPbouiCOM.DataTable, ByVal tblname As String, ByVal KeyName As String)
        Dim oColumn As SAPbouiCOM.Column
        oActiveForm.DataSources.DataTables.Add(tblname)
        dtmatrix = oActiveForm.DataSources.DataTables.Item(tblname)

        dtmatrix.Columns.Add("LineId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Ton", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Desc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Rate)
        dtmatrix.Columns.Add("UOM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Price", SAPbouiCOM.BoFieldsType.ft_Price)
        If KeyName = "CRANE" Then  ' to combine
            dtmatrix.Columns.Add("CType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
            dtmatrix.Columns.Add("Hrs", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
            oColumn = ObjMatrix.Columns.Item("colCType")
            oColumn.DataBind.Bind(tblname, "CType")
            oColumn = ObjMatrix.Columns.Item("colHr")
            oColumn.DataBind.Bind(tblname, "Hrs")
        End If
        dtmatrix.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Rate)
        dtmatrix.Columns.Add("Remark", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)



        oColumn = ObjMatrix.Columns.Item("V_-1")
        oColumn.DataBind.Bind(tblname, "LineId")
     
        oColumn = ObjMatrix.Columns.Item("colTon")
        oColumn.DataBind.Bind(tblname, "Ton")
        oColumn = ObjMatrix.Columns.Item("colDesc")
        oColumn.DataBind.Bind(tblname, "Desc")
        oColumn = ObjMatrix.Columns.Item("colQty")
        oColumn.DataBind.Bind(tblname, "Qty")
        oColumn = ObjMatrix.Columns.Item("colUOM")
        oColumn.DataBind.Bind(tblname, "UOM")
        oColumn = ObjMatrix.Columns.Item("colPrice")
        oColumn.DataBind.Bind(tblname, "Price")
        oColumn = ObjMatrix.Columns.Item("colTotal")
        oColumn.DataBind.Bind(tblname, "Total")
        oColumn = ObjMatrix.Columns.Item("colRmk")
        oColumn.DataBind.Bind(tblname, "Remark")

    End Sub

    Public Sub RowAddToMatrix(ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal source As String)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtCDetail")
        AddDataToDataTable(oActiveForm, oMatrix, source)
        dtmatrix.Rows.Add(1)
        dtmatrix.SetValue("LineId", dtmatrix.Rows.Count - 1, dtmatrix.Rows.Count)
        oMatrix.Clear()
        oMatrix.LoadFromDataSource()
    End Sub
    Public Sub AddDataToDataTable(ByVal oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal source As String)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtCDetail")
        Dim i As Integer = 0
        If dtmatrix.Rows.Count > 0 Then
            If source = "Crane" Then
                For i = 0 To dtmatrix.Rows.Count - 1
                    dtmatrix.SetValue("CType", i, oMatrix.Columns.Item("colCType").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Ton", i, oMatrix.Columns.Item("colTon").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Desc", i, oMatrix.Columns.Item("colDesc").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Qty", i, oMatrix.Columns.Item("colQty").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("UOM", i, oMatrix.Columns.Item("colUOM").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Price", i, oMatrix.Columns.Item("colPrice").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Hrs", i, oMatrix.Columns.Item("colHr").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Total", i, oMatrix.Columns.Item("colTotal").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Remark", i, oMatrix.Columns.Item("colRmk").Cells.Item(i + 1).Specific.Value)
                Next
            ElseIf source = "Forklift" Then 'to combine
                For i = 0 To dtmatrix.Rows.Count - 1
                    dtmatrix.SetValue("Ton", i, oMatrix.Columns.Item("colTon").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Desc", i, oMatrix.Columns.Item("colDesc").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Qty", i, oMatrix.Columns.Item("colQty").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("UOM", i, oMatrix.Columns.Item("colUOM").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Price", i, oMatrix.Columns.Item("colPrice").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Total", i, oMatrix.Columns.Item("colTotal").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Remark", i, oMatrix.Columns.Item("colRmk").Cells.Item(i + 1).Specific.Value)
                Next
            ElseIf source = "Crate" Then 'to combine
                For i = 0 To dtmatrix.Rows.Count - 1

                    dtmatrix.SetValue("Dimension", i, oMatrix.Columns.Item("colDimen").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Type", i, oMatrix.Columns.Item("colType").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Qty", i, oMatrix.Columns.Item("colQty").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Price", i, oMatrix.Columns.Item("colPrice").Cells.Item(i + 1).Specific.Value)
                    dtmatrix.SetValue("Total", i, oMatrix.Columns.Item("colTotal").Cells.Item(i + 1).Specific.Value)

                Next
            End If
        End If



    End Sub

    Public Sub CalTotal(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer, ByVal source As String)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim RS As SAPbobsCOM.Recordset
        RS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oMatrix = oActiveForm.Items.Item("mx_CDetail").Specific

        Dim Qty As Double = 0.0
        Dim Price As Double = 0.0
        Dim Total As Double = 0.0

        Total = Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(Row).Specific.Value)

        oMatrix.Columns.Item("colTotal").Editable = True
        oMatrix.Columns.Item("colTotal").Cells.Item(Row).Specific.Value = Convert.ToString(Total)
        If source = "Crane" Then 'to combine
            oMatrix.Columns.Item("colHr").Cells.Item(Row).Click()
        ElseIf source = "Forklift" Then 'to combine
            oMatrix.Columns.Item("colRmk").Cells.Item(Row).Click()
        ElseIf source = "Crate" Then 'to combine
            oMatrix.Columns.Item("colDimen").Cells.Item(Row).Click()
        End If

        oMatrix.Columns.Item("colTotal").Editable = False

        If source = "Crane" Then 'to combine
            dtmatrix.SetValue("CType", Row - 1, oMatrix.Columns.Item("colCType").Cells.Item(Row).Specific.Value)
            dtmatrix.SetValue("Hrs", Row - 1, oMatrix.Columns.Item("colHr").Cells.Item(Row).Specific.Value)

        End If
        If source = "Crane" Or source = "Forklift" Then 'to combine
            dtmatrix.SetValue("Ton", Row - 1, oMatrix.Columns.Item("colTon").Cells.Item(Row).Specific.Value)
            dtmatrix.SetValue("Desc", Row - 1, oMatrix.Columns.Item("colDesc").Cells.Item(Row).Specific.Value)
            dtmatrix.SetValue("UOM", Row - 1, oMatrix.Columns.Item("colUOM").Cells.Item(Row).Specific.Value)
        End If
        If source = "Crate" Then 'to combine
            dtmatrix.SetValue("Type", Row - 1, oMatrix.Columns.Item("colType").Cells.Item(Row).Specific.Value)
        End If

        dtmatrix.SetValue("Qty", Row - 1, oMatrix.Columns.Item("colQty").Cells.Item(Row).Specific.Value)
        dtmatrix.SetValue("Price", Row - 1, oMatrix.Columns.Item("colPrice").Cells.Item(Row).Specific.Value)
        dtmatrix.SetValue("Total", Row - 1, Total)

        oMatrix.LoadFromDataSource()
    End Sub
#End Region

    Private Function EditPOInEditTab(ByVal ActiveForm As SAPbouiCOM.Form, ByVal PONo As String, ByVal matrixName As String) As Boolean
        EditPOInEditTab = False
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            If matrixName = "mx_TkrList" Then
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_TkrIns = " + FormatString(ActiveForm.Items.Item("ee_TkrIns").Specific.Value) + ",U_TkrTo= " + FormatString(ActiveForm.Items.Item("ee_TkrTo").Specific.Value) + ",U_ColFrm= " + FormatString(ActiveForm.Items.Item("ee_ColFrm").Specific.Value) + ",U_PORMKS= " + FormatString(ActiveForm.Items.Item("ee_Rmsk").Specific.Value) + ",U_POIRMKS= " + FormatString(ActiveForm.Items.Item("ee_InsRmsk").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            ElseIf matrixName = "mx_DspList" Then
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_TkrIns = " + FormatString(ActiveForm.Items.Item("ee_DspIns").Specific.Value) + ",U_PORMKS= " + FormatString(ActiveForm.Items.Item("ee_DRmsk").Specific.Value) + ",U_POIRMKS= " + FormatString(ActiveForm.Items.Item("ee_DIRmsk").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            ElseIf matrixName = "mx_Crane" Then
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_SIns= " + FormatString(ActiveForm.Items.Item("ed_SIns").Specific.Value) + ",U_PORMKS= " + FormatString(ActiveForm.Items.Item("ed_Remark").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            ElseIf matrixName = "mx_Fork" Then 'to combine
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PORMKS= " + FormatString(ActiveForm.Items.Item("ed_Remark").Specific.Value) + ",U_POIRMKS= " + FormatString(ActiveForm.Items.Item("ed_IRemark").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            ElseIf matrixName = "mx_Crate" Then 'to combine
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PORMKS= " + FormatString(ActiveForm.Items.Item("ed_Remark").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            ElseIf matrixName = "mx_Toll" Then
                oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_POIRMKS= " + FormatString(ActiveForm.Items.Item("ee_InsRmsk").Specific.Value) + ",U_PORMKS= " + FormatString(ActiveForm.Items.Item("ed_Remark").Specific.Value) + " WHERE U_PONo = " + FormatString(PONo))
            End If
            EditPOInEditTab = True
        Catch ex As Exception
            EditPOInEditTab = False
        End Try
    End Function

#Region "TableForBunker"
    Public Sub CreateDTDetailBunker(ByVal oActiveForm As SAPbouiCOM.Form)
        ObjMatrix = oActiveForm.Items.Item("mx_BDetail").Specific
        CreateTblStructureBunker(oActiveForm, dtmatrix, "dtBDetail")
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtBDetail")
        dtmatrix.Rows.Add(1)
        dtmatrix.SetValue("LineId", 0, 1)
        ObjMatrix.LoadFromDataSource()
    End Sub

    Public Sub CreateTblStructureBunker(ByVal oActiveForm As SAPbouiCOM.Form, ByVal dtmatrix As SAPbouiCOM.DataTable, ByVal tblname As String)
        Dim oColumn As SAPbouiCOM.Column
        oActiveForm.DataSources.DataTables.Add(tblname)
        dtmatrix = oActiveForm.DataSources.DataTables.Item(tblname)

        dtmatrix.Columns.Add("LineId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Permit", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("JobNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Client", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity)
        dtmatrix.Columns.Add("UOM", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Kgs", SAPbouiCOM.BoFieldsType.ft_Price)
        dtmatrix.Columns.Add("M3", SAPbouiCOM.BoFieldsType.ft_Rate)
        dtmatrix.Columns.Add("NEQ", SAPbouiCOM.BoFieldsType.ft_Rate)
        dtmatrix.Columns.Add("Stus", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)



        oColumn = ObjMatrix.Columns.Item("V_-1")
        oColumn.DataBind.Bind(tblname, "LineId")
        oColumn = ObjMatrix.Columns.Item("colPermit")
        oColumn.DataBind.Bind(tblname, "Permit")
        oColumn = ObjMatrix.Columns.Item("colJobNo")
        oColumn.DataBind.Bind(tblname, "JobNo")
        oColumn = ObjMatrix.Columns.Item("colClient")
        oColumn.DataBind.Bind(tblname, "Client")


        oColumn = ObjMatrix.Columns.Item("colQty")
        oColumn.DataBind.Bind(tblname, "Qty")
        oColumn = ObjMatrix.Columns.Item("colUOM")
        oColumn.DataBind.Bind(tblname, "UOM")
        oColumn = ObjMatrix.Columns.Item("colKgs")
        oColumn.DataBind.Bind(tblname, "Kgs")
        oColumn = ObjMatrix.Columns.Item("colM3")
        oColumn.DataBind.Bind(tblname, "M3")
        oColumn = ObjMatrix.Columns.Item("colNEQ")
        oColumn.DataBind.Bind(tblname, "NEQ")
        oColumn = ObjMatrix.Columns.Item("colStus")
        oColumn.DataBind.Bind(tblname, "Stus")

    End Sub

    Public Sub RowAddToMatrixBunker(ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtBDetail")
        AddDataToDataTableBunker(oActiveForm, oMatrix)
        dtmatrix.Rows.Add(1)
        dtmatrix.SetValue("LineId", dtmatrix.Rows.Count - 1, dtmatrix.Rows.Count)
        oMatrix.Clear()
        oMatrix.LoadFromDataSource()
    End Sub
    Public Sub AddDataToDataTableBunker(ByVal oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtBDetail")
        Dim i As Integer = 0
        If dtmatrix.Rows.Count > 0 Then
            For i = 0 To dtmatrix.Rows.Count - 1
                dtmatrix.SetValue(1, i, oMatrix.Columns.Item("colPermit").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(2, i, oMatrix.Columns.Item("colJobNo").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(3, i, oMatrix.Columns.Item("colClient").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(4, i, oMatrix.Columns.Item("colQty").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(5, i, oMatrix.Columns.Item("colUOM").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(6, i, oMatrix.Columns.Item("colKgs").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(7, i, oMatrix.Columns.Item("colM3").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(8, i, oMatrix.Columns.Item("colNEQ").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(9, i, oMatrix.Columns.Item("colStus").Cells.Item(i + 1).Specific.Value)
            Next
        End If
    End Sub

    Public Sub CalTotalBunker(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim oMatrix As SAPbouiCOM.Matrix
        oActiveForm.Freeze(True)
        oMatrix = oActiveForm.Items.Item("mx_BDetail").Specific
        Dim TQty As Double = 0.0
        Dim TKgs As Double = 0.0
        Dim TM3 As Double = 0.0
        Dim TNEQ As Double = 0.0
        For i As Integer = 1 To oMatrix.RowCount
            TQty = TQty + Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(i).Specific.Value)
            TKgs = TKgs + Convert.ToDouble(oMatrix.Columns.Item("colKgs").Cells.Item(i).Specific.Value)
            TM3 = TM3 + Convert.ToDouble(oMatrix.Columns.Item("colM3").Cells.Item(i).Specific.Value)
            TNEQ = TNEQ + Convert.ToDouble(oMatrix.Columns.Item("colNEQ").Cells.Item(i).Specific.Value)
        Next
        oActiveForm.Items.Item("ed_TQty").Specific.Value = TQty
        oActiveForm.Items.Item("ed_TKgs").Specific.Value = TKgs
        oActiveForm.Items.Item("ed_TM3").Specific.Value = TM3
        oActiveForm.Items.Item("ed_TNEQ").Specific.Value = TNEQ
        AddDataToDataTableBunker(oActiveForm, oMatrix)
        oMatrix.LoadFromDataSource()
        oActiveForm.Freeze(False)
    End Sub
#End Region

#Region "TableForToll"
    Public Sub CreateDTDetailToll(ByVal oActiveForm As SAPbouiCOM.Form)
        Dim oColumn As SAPbouiCOM.Column
        ObjMatrix = oActiveForm.Items.Item("mx_TDetail").Specific
        CreateTblStructureToll(oActiveForm, dtmatrix, "dtTDetail")
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtTDetail")
        dtmatrix.Rows.Add(1)
        dtmatrix.SetValue("LineId", 0, 1)
        oColumn = ObjMatrix.Columns.Item("colICode")
        AddChooseFromList(oActiveForm, "ICode", False, "4")
        oColumn.ChooseFromListUID = "ICode"
        oColumn.ChooseFromListAlias = "ItemCode" 'MSW To Edit New Ticket
        ObjMatrix.LoadFromDataSource()
    End Sub

    Public Sub CreateTblStructureToll(ByVal oActiveForm As SAPbouiCOM.Form, ByVal dtmatrix As SAPbouiCOM.DataTable, ByVal tblname As String)
        Dim oColumn As SAPbouiCOM.Column
        oActiveForm.DataSources.DataTables.Add(tblname)
        dtmatrix = oActiveForm.DataSources.DataTables.Item(tblname)

        dtmatrix.Columns.Add("LineId", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("ICode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("IDesc", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity)
        dtmatrix.Columns.Add("Price", SAPbouiCOM.BoFieldsType.ft_Rate)
        dtmatrix.Columns.Add("Total", SAPbouiCOM.BoFieldsType.ft_Rate)




        oColumn = ObjMatrix.Columns.Item("V_-1")
        oColumn.DataBind.Bind(tblname, "LineId")
        oColumn = ObjMatrix.Columns.Item("colICode")
        oColumn.DataBind.Bind(tblname, "ICode")
        oColumn = ObjMatrix.Columns.Item("colIDesc")
        oColumn.DataBind.Bind(tblname, "IDesc")
        oColumn = ObjMatrix.Columns.Item("colQty")
        oColumn.DataBind.Bind(tblname, "Qty")
        oColumn = ObjMatrix.Columns.Item("colPrice")
        oColumn.DataBind.Bind(tblname, "Price")
        oColumn = ObjMatrix.Columns.Item("colTotal")
        oColumn.DataBind.Bind(tblname, "Total")


    End Sub

    Public Sub RowAddToMatrixToll(ByRef oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtTDetail")
        AddDataToDataTableToll(oActiveForm, oMatrix)
        dtmatrix.Rows.Add(1)
        dtmatrix.SetValue("LineId", dtmatrix.Rows.Count - 1, dtmatrix.Rows.Count)
        oMatrix.Clear()
        oMatrix.LoadFromDataSource()
    End Sub
    Public Sub AddDataToDataTableToll(ByVal oActiveForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix)
        dtmatrix = oActiveForm.DataSources.DataTables.Item("dtTDetail")
        Dim i As Integer = 0
        If dtmatrix.Rows.Count > 0 Then
            For i = 0 To dtmatrix.Rows.Count - 1
                dtmatrix.SetValue(1, i, oMatrix.Columns.Item("colICode").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(2, i, oMatrix.Columns.Item("colIDesc").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(3, i, oMatrix.Columns.Item("colQty").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(4, i, oMatrix.Columns.Item("colPrice").Cells.Item(i + 1).Specific.Value)
                dtmatrix.SetValue(5, i, oMatrix.Columns.Item("colTotal").Cells.Item(i + 1).Specific.Value)
            Next
        End If
    End Sub

    Public Sub CalTotalToll(ByVal oActiveForm As SAPbouiCOM.Form, ByVal Row As Integer)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim Total As Double
        oActiveForm.Freeze(True)
        oMatrix = oActiveForm.Items.Item("mx_TDetail").Specific
      

        Total = Convert.ToDouble(oMatrix.Columns.Item("colQty").Cells.Item(Row).Specific.Value) * Convert.ToDouble(oMatrix.Columns.Item("colPrice").Cells.Item(Row).Specific.Value)

        oMatrix.Columns.Item("colTotal").Editable = True
        oMatrix.Columns.Item("colTotal").Cells.Item(Row).Specific.Value = Convert.ToString(Total)
        oMatrix.Columns.Item("colIDesc").Cells.Item(Row).Click()
        oMatrix.Columns.Item("colTotal").Editable = False
        
        AddDataToDataTableToll(oActiveForm, oMatrix)
        oMatrix.LoadFromDataSource()
        oActiveForm.Freeze(False)
    End Sub
#End Region
    
End Module



