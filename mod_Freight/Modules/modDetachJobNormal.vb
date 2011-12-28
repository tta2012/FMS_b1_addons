Option Explicit On
Imports SAPbouiCOM
Module modDetachJobNormal
    Private ObjForm As SAPbouiCOM.Form
    Private ObjItem As SAPbouiCOM.Item
    Private ObjMatrix As SAPbouiCOM.Matrix
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim MJobForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
    Dim dtmatrix As SAPbouiCOM.DataTable
    Dim sErrDesc As String = String.Empty
    Dim oCombo As SAPbouiCOM.ComboBox


    Public Sub LoadDetachJForm()
        CreateDTMJ()
    End Sub

    Public Sub CreateDTMJ()
        CreateDTSelect()
    End Sub

    Public Sub CreateDTSelect()

        ObjMatrix = MJobForm.Items.Item("mx_Select").Specific

        CreateTblStructure(dtmatrix, "dtSelect")
        dtmatrix = MJobForm.DataSources.DataTables.Item("dtSelect")
        dtmatrix.Rows.Add(1)
        Dim sql As String = "select U_JobNum from [@OBT_TB01_MULTIPO] Where U_JobNum <>  '" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' And U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
        dtmatrix.ExecuteQuery(sql)
        ObjMatrix.LoadFromDataSource()
    End Sub

    Public Sub CreateTblStructure(ByVal dtmatrix As SAPbouiCOM.DataTable, ByVal tblname As String)
        Dim oColumn As SAPbouiCOM.Column
        MJobForm.DataSources.DataTables.Add(tblname)
        dtmatrix = MJobForm.DataSources.DataTables.Item(tblname)


        dtmatrix.Columns.Add("SJobNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Schk", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)

        oColumn = ObjMatrix.Columns.Item("colJobNo")
        oColumn.DataBind.Bind(tblname, "SJobNo")
        oColumn = ObjMatrix.Columns.Item("colChk")
        oColumn.DataBind.Bind(tblname, "Schk")

    End Sub

    Public Function DetachMultiJob(ByVal pForm As SAPbouiCOM.Form, ByVal rmkName As String, ByVal multibox As String) As Boolean
        DetachMultiJob = False
        Dim omatrix As SAPbouiCOM.Matrix = MJobForm.Items.Item("mx_Select").Specific
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim sql As String = ""
        Dim count As Integer = 0
        Dim jobDocEntry As Integer
        Dim jobDocEntryButton As Integer  'Fumigation

        Dim remark As String = ""

        Dim tblHeader As String = ""
        Dim tblDetail As String = ""
        Dim tblItem As String = ""
        Dim tblDetail2 As String = ""

        If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Then
            tblDetail = "OBT_FCL03_ETRUCKING"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
            tblDetail = "OBT_FCL04_EDISPATCH"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Fumigation" Then
            tblHeader = "FUMIGATION"
            tblDetail = "OBT_TBL01_FUMIGAT"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Outrider" Then  '15/12/2011
            tblHeader = "OUTRIDER"
            tblDetail = "OBT_TBL03_OUTRIDER"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Crane" Then
            tblHeader = "CRANE"
            tblDetail = "OBT_TB33_CRANE"
            tblDetail2 = "OBT_TB01_CRNDETAIL"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Forklift" Then 'to combine
            tblHeader = "FORKLIFT"
            tblDetail = "OBT_TBL05_FORKLIFT"
            tblDetail2 = "OBT_TBL07_FORDETAIL" 'to combine
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Crate" Then 'to combine
            tblHeader = "CRATE"
            tblDetail = "OBT_TBL08_CRATE"
            tblDetail2 = "OBT_TBL10_CRADETAIL" 'to combine
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Bunker" Then
            tblHeader = "BUNKER"
            tblDetail = "OBT_TB01_BUNKER"
            tblDetail2 = "OBT_TB01_BNKDETAIL"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Toll" Then
            tblHeader = "TOLL"
            tblDetail = "OBT_TB01_TOLL"
            tblDetail2 = "OBT_TB01_TOLLDETAIL"
        End If

        Try
            If omatrix.RowCount > 0 Then

                For i As Integer = 1 To omatrix.RowCount

                    If omatrix.Columns.Item("colChk").Cells.Item(i).Specific.Checked = True Then
                        sql = "Delete from [@OBT_TB01_MULTIPO]  Where U_JobNum='" & omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' And U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                        oRecordSet.DoQuery(sql)
                        sql = "select DocEntry from [@OBT_FCL01_EXPORT]  Where U_JobNum='" & omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "'"
                        oRecordSet.DoQuery(sql)
                        jobDocEntry = oRecordSet.Fields.Item("DocEntry").Value

                        If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
                            sql = "Delete from [@" & tblDetail & "] Where DocEntry='" & jobDocEntry & "' And U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                            oRecordSet.DoQuery(sql)
                            sql = "Delete from [@OBT_TB01_POLIST]  Where DocEntry='" & jobDocEntry & "' And U_PONo='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                            oRecordSet.DoQuery(sql)
                            sql = "Select * from [@" & tblDetail & "]  Where DocEntry='" & jobDocEntry & "'"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount = 0 Then
                                sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId) Values ('" & jobDocEntry & "','1')"
                                oRecordSet.DoQuery(sql)
                            End If
                            sql = "Select * from[@OBT_TB01_POLIST]  Where DocEntry='" & jobDocEntry & "'"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount = 0 Then
                                sql = "Insert Into [@OBT_TB01_POLIST] (DocEntry,LineId) Values ('" & jobDocEntry & "','1')"
                                oRecordSet.DoQuery(sql)
                            End If
                            If remark = "" Then
                                remark = Replace(pForm.Items.Item(rmkName).Specific.Value.ToString, omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value, "")
                            Else
                                remark = Replace(remark, omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value, "")
                            End If

                        Else
                            sql = "select DocEntry  from [@" & tblHeader & "]  Where U_JobNo='" & omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "'"  'Fumigation 
                            oRecordSet.DoQuery(sql)
                            jobDocEntryButton = oRecordSet.Fields.Item("DocEntry").Value

                            sql = "Delete from [@" & tblDetail & "]  Where DocEntry='" & jobDocEntryButton & "' And U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                            oRecordSet.DoQuery(sql)
                            sql = "Delete from [@OBT_TB01_POLIST]  Where DocEntry='" & jobDocEntry & "' And U_PONo='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                            oRecordSet.DoQuery(sql)
                            sql = "Select * from [@" & tblDetail & "]  Where DocEntry='" & jobDocEntryButton & "'"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount = 0 Then
                                sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId) Values ('" & jobDocEntryButton & "','1')"
                                oRecordSet.DoQuery(sql)
                            End If
                            sql = "Select * from[@OBT_TB01_POLIST]  Where DocEntry='" & jobDocEntry & "'"
                            oRecordSet.DoQuery(sql)
                            If oRecordSet.RecordCount = 0 Then
                                sql = "Insert Into [@OBT_TB01_POLIST] (DocEntry,LineId) Values ('" & jobDocEntry & "','1')"
                                oRecordSet.DoQuery(sql)
                            End If
                            If tblDetail2 <> "" Then
                                sql = "Delete from [@" & tblDetail2 & "]  Where U_DocN='" & jobDocEntry & "' And U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                                oRecordSet.DoQuery(sql)
                            End If
                            If remark = "" Then
                                remark = Replace(pForm.Items.Item(rmkName).Specific.Value.ToString, omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value, "")
                            Else
                                remark = Replace(remark, omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value, "")
                            End If
                        End If
                    End If
                Next
                If remark <> "" Then
                    remark = Replace(remark, ",,", ",")
                    If Right(remark, 1) = "," Then
                        remark = Left(remark, remark.Length - 1)
                    End If
                    oRecordSet.DoQuery("UPDATE OPOR SET Comments = '" + remark + "' WHERE DocNum = '" & MJobForm.Items.Item("ed_PO").Specific.Value & "'")
                    oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO]  SET U_PORMKS  = '" + remark + "' WHERE U_PO = '" & MJobForm.Items.Item("ed_PO").Specific.Value & "'")
                    pForm.Items.Item(rmkName).Specific.Value = remark
                End If
                sql = "Select * from [@OBT_TB01_MULTIPO] where U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                oRecordSet.DoQuery(sql)
                If oRecordSet.RecordCount = 1 Then
                    sql = "Delete from [@OBT_TB01_MULTIPO] where U_PO='" & MJobForm.Items.Item("ed_PO").Specific.Value & "'"
                    oRecordSet.DoQuery(sql)
                    pForm.Items.Item(multibox).Specific.Value = ""
                End If
            End If
            DetachMultiJob = True


        Catch ex As Exception
            DetachMultiJob = False
            MessageBox.Show(ex.Message)
        End Try


    End Function
    Public Function UpdateRemarkForPOTable(ByVal PONo As String, ByVal poRemark As String) As Boolean
        UpdateRemarkForPOTable = False
        Try

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("UPDATE [@OBT_TB08_FFCPO] SET U_PORMKS = '" + poRemark + "' WHERE U_PONo = " + FormatString(PONo))
            UpdateRemarkForPOTable = True
        Catch ex As Exception
            UpdateRemarkForPOTable = False
        End Try
    End Function
    Public Function UpdateB1RemarkForMultiPO(ByVal pForm As SAPbouiCOM.Form, ByVal poRemark As String, ByVal poBox As String) As Boolean
        UpdateB1RemarkForMultiPO = False
        Try

            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery("UPDATE OPOR SET Comments = '" + poRemark + "' WHERE DocEntry = " + FormatString(pForm.Items.Item(poBox).Specific.Value))
            UpdateB1RemarkForMultiPO = True
        Catch ex As Exception
            UpdateB1RemarkForMultiPO = False
        End Try
    End Function

    

End Module


