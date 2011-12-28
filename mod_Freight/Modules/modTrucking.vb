Option Explicit On

Imports SAPbouiCOM

Module modTrucking
    Private ObjForm As SAPbouiCOM.Form
    Private ObjItem As SAPbouiCOM.Item
    Private ObjMatrix As SAPbouiCOM.Matrix
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Private LineId, DocumentNo, PurchaseOrderNo, PurchaseDocNo, POSerialNo, POStus, PrepBy, TMulti, TOrigin, PODate, InstructionDate, Mode, TkrCode, Trucker, VehicleNo, EUC, Attention, Telephone, Fax, Email, TruckingDate, TruckingTime, CollectFrom, TruckTo, TruckingInstruction, Remarks, newRemarks, PreparedBy As String
    Public rowIndex As Integer

    Public Sub AddUpdateInstructions(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String, ByVal ProcressedState As Boolean)
        ObjDBDataSource = pForm.DataSources.DBDataSources.Item(DataSource)
        rowIndex = pMatrix.GetNextSelectedRow

        'If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And (pMatrix.GetNextSelectedRow = -1 Or pMatrix.GetNextSelectedRow = 0) Then
        '    rowIndex = 1
        'End If

        If rowIndex = -1 And pForm.Items.Item("ed_InsDoc").Specific.Value <> "" Then
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("SELECT LineId FROM [" & DataSource & "] Where LineId=" & pForm.Items.Item("ed_InsDoc").Specific.Value)
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
                    .SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_PODocNo").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PONo", .Offset, pForm.Items.Item("ed_PONo").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value) 'New UI PO Serial No
                    .SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_InsDate").Specific.Value)
                    .SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_Inter").Specific.Selected = True, "Internal", "External").ToString)
                    .SetValue("U_TkrCode", .Offset, pForm.Items.Item("ed_TkrCode").Specific.Value)
                    .SetValue("U_Trucker", .Offset, pForm.Items.Item("ed_Trucker").Specific.Value)
                    .SetValue("U_VehNo", .Offset, pForm.Items.Item("ed_VehicNo").Specific.Value)
                    .SetValue("U_EUC", .Offset, pForm.Items.Item("ed_EUC").Specific.Value)
                    .SetValue("U_Attent", .Offset, pForm.Items.Item("ed_Attent").Specific.Value)
                    .SetValue("U_Tel", .Offset, pForm.Items.Item("ed_TkrTel").Specific.Value)
                    .SetValue("U_Fax", .Offset, pForm.Items.Item("ed_Fax").Specific.Value)
                    .SetValue("U_Email", .Offset, pForm.Items.Item("ed_Email").Specific.Value)
                    .SetValue("U_TkrDate", .Offset, pForm.Items.Item("ed_TkrDate").Specific.Value)
                    'If Not pForm.Items.Item("ed_TkrTime").Specific.Value.ToString = "" Then
                    '    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(2, 2).ToString())
                    'End If
                    '.SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(3, 2).ToString())
                    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value)
                    'MSW 14-09-2011 Truck PO
                    If pForm.Items.Item("op_Exter").Specific.Selected = True Then
                        .SetValue("U_Status", .Offset, "Open")
                    Else
                        .SetValue("U_Status", .Offset, "")
                    End If

                    'MSW 14-09-2011 Truck PO
                    .SetValue("U_ColFrm", .Offset, pForm.Items.Item("ee_ColFrm").Specific.Value)
                    .SetValue("U_TkrTo", .Offset, pForm.Items.Item("ee_TkrTo").Specific.Value)
                    .SetValue("U_TkrIns", .Offset, pForm.Items.Item("ee_TkrIns").Specific.Value)
                    .SetValue("U_InsRemsk", .Offset, pForm.Items.Item("ee_InsRmsk").Specific.Value)
                    .SetValue("U_Remarks", .Offset, pForm.Items.Item("ee_Rmsk").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PrepBy", .Offset, p_oDICompany.UserName.ToString)
                    .SetValue("U_PODate", .Offset, pForm.Items.Item("ed_Date").Specific.Value)
                    .SetValue("U_MultiJob", .Offset, pForm.Items.Item("ed_TMulti").Specific.Value)
                    .SetValue("U_OriginPO", .Offset, pForm.Items.Item("ed_TOrigin").Specific.Value)
                    pMatrix.AddRow()
                End With
            Else
                With ObjDBDataSource
                    
                    .SetValue("LineId", .Offset, pForm.Items.Item("ed_InsDoc").Specific.Value)
                    .SetValue("U_InsDocNo", .Offset, pForm.Items.Item("ed_InsDoc").Specific.Value)
                    .SetValue("U_PODocNo", .Offset, pForm.Items.Item("ed_PODocNo").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PONo", .Offset, pForm.Items.Item("ed_PONo").Specific.Value)
                    .SetValue("U_PO", .Offset, pForm.Items.Item("ed_PO").Specific.Value) 'New UI PO Serial No
                    .SetValue("U_InsDate", .Offset, pForm.Items.Item("ed_InsDate").Specific.Value)
                    .SetValue("U_Mode", .Offset, IIf(pForm.Items.Item("op_Inter").Specific.Selected = True, "Internal", "External").ToString)
                    .SetValue("U_TkrCode", .Offset, pForm.Items.Item("ed_TkrCode").Specific.Value)
                    .SetValue("U_Trucker", .Offset, pForm.Items.Item("ed_Trucker").Specific.Value)
                    .SetValue("U_VehNo", .Offset, pForm.Items.Item("ed_VehicNo").Specific.Value)
                    .SetValue("U_EUC", .Offset, pForm.Items.Item("ed_EUC").Specific.Value)
                    .SetValue("U_Attent", .Offset, pForm.Items.Item("ed_Attent").Specific.Value)
                    .SetValue("U_Tel", .Offset, pForm.Items.Item("ed_TkrTel").Specific.Value)
                    .SetValue("U_Fax", .Offset, pForm.Items.Item("ed_Fax").Specific.Value)
                    .SetValue("U_Email", .Offset, pForm.Items.Item("ed_Email").Specific.Value)
                    .SetValue("U_TkrDate", .Offset, pForm.Items.Item("ed_TkrDate").Specific.Value)
                    '.SetValue("U_TkrTime", .Offset, TruckingTime.ToString())
                    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value)
                    'If Not pForm.Items.Item("ed_TkrTime").Specific.Value.ToString = "" Then
                    '    .SetValue("U_TkrTime", .Offset, pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(0, 2).ToString() & pForm.Items.Item("ed_TkrTime").Specific.Value.ToString.Substring(2, 2).ToString())
                    'End If
                    'MSW 14-09-2011 Truck PO
                    If pForm.Items.Item("op_Exter").Specific.Selected = True Then
                        .SetValue("U_Status", .Offset, pForm.Items.Item("ed_PStus").Specific.Value)
                    Else
                        .SetValue("U_Status", .Offset, "")
                    End If
                    'MSW 14-09-2011 Truck PO
                    .SetValue("U_ColFrm", .Offset, pForm.Items.Item("ee_ColFrm").Specific.Value)
                    .SetValue("U_TkrTo", .Offset, pForm.Items.Item("ee_TkrTo").Specific.Value)
                    .SetValue("U_TkrIns", .Offset, pForm.Items.Item("ee_TkrIns").Specific.Value)
                    .SetValue("U_InsRemsk", .Offset, pForm.Items.Item("ee_InsRmsk").Specific.Value)
                    .SetValue("U_Remarks", .Offset, pForm.Items.Item("ee_Rmsk").Specific.Value) 'MSW 14-09-2011 Truck PO
                    .SetValue("U_PrepBy", .Offset, pForm.Items.Item("ed_Created").Specific.Value)
                    .SetValue("U_PODate", .Offset, pForm.Items.Item("ed_Date").Specific.Value)
                    .SetValue("U_MultiJob", .Offset, pForm.Items.Item("ed_TMulti").Specific.Value)
                    .SetValue("U_OriginPO", .Offset, pForm.Items.Item("ed_TOrigin").Specific.Value)
                    pMatrix.SetLineData(rowIndex)
                End With
            End If
        Catch ex As Exception

        End Try

    End Sub

    Public Sub GetDataFromMatrixByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal Index As Integer)
        Try
            DocumentNo = pMatrix.Columns.Item("colInsDoc").Cells.Item(Index).Specific.Value
            PurchaseOrderNo = pMatrix.Columns.Item("colPONo").Cells.Item(Index).Specific.Value
            PurchaseDocNo = pMatrix.Columns.Item("colDocNo").Cells.Item(Index).Specific.Value
            POSerialNo = pMatrix.Columns.Item("colPO").Cells.Item(Index).Specific.Value  'New UI
            Mode = pMatrix.Columns.Item("colMode").Cells.Item(Index).Specific.Value
            InstructionDate = pMatrix.Columns.Item("colInsDate").Cells.Item(Index).Specific.Value
            TkrCode = pMatrix.Columns.Item("colTkrCode").Cells.Item(Index).Specific.Value
            Trucker = pMatrix.Columns.Item("colTrucker").Cells.Item(Index).Specific.Value
            VehicleNo = pMatrix.Columns.Item("colVehNo").Cells.Item(Index).Specific.Value
            EUC = pMatrix.Columns.Item("colEUC").Cells.Item(Index).Specific.Value
            Attention = pMatrix.Columns.Item("colAttent").Cells.Item(Index).Specific.Value
            Telephone = pMatrix.Columns.Item("colTel").Cells.Item(Index).Specific.Value
            Fax = pMatrix.Columns.Item("colFax").Cells.Item(Index).Specific.Value
            Email = pMatrix.Columns.Item("colEmail").Cells.Item(Index).Specific.Value
            TruckingDate = pMatrix.Columns.Item("colTkrDate").Cells.Item(Index).Specific.Value
            TruckingTime = pMatrix.Columns.Item("colTkrTime").Cells.Item(Index).Specific.Value
            CollectFrom = pMatrix.Columns.Item("colColFrom").Cells.Item(Index).Specific.Value
            TruckTo = pMatrix.Columns.Item("colTkrTo").Cells.Item(Index).Specific.Value
            TruckingInstruction = pMatrix.Columns.Item("colTkrIns").Cells.Item(Index).Specific.Value
            Remarks = pMatrix.Columns.Item("colRemarks").Cells.Item(Index).Specific.Value
            newRemarks = pMatrix.Columns.Item("colRmks").Cells.Item(Index).Specific.Value 'MSW 08-09-2011
            TruckingTime = TruckingTime.Substring(0, 2).ToString() & ":" & TruckingTime.Substring(2, 2).ToString()
            POStus = pMatrix.Columns.Item("colPStatus").Cells.Item(Index).Specific.Value 'New UI
            PODate = pMatrix.Columns.Item("colDate").Cells.Item(Index).Specific.Value
            PrepBy = pMatrix.Columns.Item("colPrepBy").Cells.Item(Index).Specific.Value
            TMulti = pMatrix.Columns.Item("colMulti").Cells.Item(Index).Specific.Value
            TOrigin = pMatrix.Columns.Item("colOrigin").Cells.Item(Index).Specific.Value
        Catch ex As Exception

        End Try
    End Sub

    Public Sub SetDataToEditTabByIndex(ByVal pForm As SAPbouiCOM.Form)
        Try
            pForm.Items.Item("ed_InsDoc").Specific.Value = DocumentNo
            pForm.Items.Item("ed_PONo").Specific.Value = PurchaseOrderNo
            pForm.Items.Item("ed_PODocNo").Specific.Value = PurchaseDocNo
            pForm.Items.Item("ed_PO").Specific.Value = POSerialNo
            If Mode = "Internal" Then
                pForm.DataSources.UserDataSources.Item("TKINTR").ValueEx = "1"
                pForm.DataSources.UserDataSources.Item("TKEXTR").ValueEx = "2"
            ElseIf Mode = "External" Then
                pForm.DataSources.UserDataSources.Item("TKEXTR").ValueEx = "1"
                pForm.DataSources.UserDataSources.Item("TKINTR").ValueEx = "2"
            End If
            pForm.Items.Item("ed_InsDate").Specific.Value = InstructionDate
            pForm.Items.Item("ed_TkrCode").Specific.Value = TkrCode
            pForm.Items.Item("ed_Trucker").Specific.Value = Trucker
            pForm.Items.Item("ed_VehicNo").Specific.Value = VehicleNo
            pForm.Items.Item("ed_EUC").Specific.Value = EUC
            pForm.Items.Item("ed_Attent").Specific.Value = Attention
            pForm.Items.Item("ed_TkrTel").Specific.Value = Telephone
            pForm.Items.Item("ed_Fax").Specific.Value = Fax
            pForm.Items.Item("ed_Email").Specific.Value = Email
            pForm.Items.Item("ed_TkrDate").Specific.Value = TruckingDate
            pForm.Items.Item("ed_TkrTime").Specific.Value = TruckingTime
            pForm.Items.Item("ee_ColFrm").Specific.Value = CollectFrom
            pForm.Items.Item("ee_TkrTo").Specific.Value = TruckTo
            pForm.Items.Item("ee_TkrIns").Specific.Value = TruckingInstruction
            pForm.Items.Item("ee_InsRmsk").Specific.Value = Remarks
            pForm.Items.Item("ee_Rmsk").Specific.Value = newRemarks 'MSW 08-09-2011
            pForm.Items.Item("ed_PStus").Specific.Value = POStus 'MSW 08-09-2011
            pForm.Items.Item("ed_Date").Specific.Value = PODate
            pForm.Items.Item("ed_Created").Specific.Value = PrepBy
            pForm.Items.Item("ed_TMulti").Specific.Value = TMulti
            pForm.Items.Item("ed_TOrigin").Specific.Value = TOrigin
        Catch ex As Exception

        End Try
    End Sub

    Public Sub DeleteByIndex(ByVal pForm As SAPbouiCOM.Form, ByVal pMatrix As SAPbouiCOM.Matrix, ByVal DataSource As String)
        Try
            If pMatrix.IsRowSelected(rowIndex) = True Then
                Try
                    If (pForm.Mode <> BoFormMode.fm_ADD_MODE) Then
                        pForm.DataSources.DBDataSources.Item(DataSource).RemoveRecord(rowIndex - 1)
                    End If
                Catch ex As Exception

                End Try

                pMatrix.DeleteRow(rowIndex)
                If pForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE And pMatrix.RowCount = 0 Then
                    pMatrix.FlushToDataSource()
                    pMatrix.AddRow(1)
                End If
                pForm.Items.Item("bt_AddIns").Specific.Caption = "Add Trucking Instruction"
                pForm.Items.Item("bt_DelIns").Enabled = False
            End If
        Catch ex As Exception

        End Try
    End Sub
End Module
