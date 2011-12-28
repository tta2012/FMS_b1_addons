Option Explicit On
Imports SAPbouiCOM
Module modMultiJobForNormal
    Private ObjForm As SAPbouiCOM.Form
    Private ObjItem As SAPbouiCOM.Item
    Private ObjMatrix As SAPbouiCOM.Matrix
    Private ObjDBDataSource As SAPbouiCOM.DBDataSource
    Dim oRecordSet As SAPbobsCOM.Recordset = Nothing
    Dim MJobForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
    Dim dtmatrix As SAPbouiCOM.DataTable
    Dim sErrDesc As String = String.Empty
    Dim oCombo As SAPbouiCOM.ComboBox

    Dim oGeneralService As SAPbobsCOM.GeneralService  'Fumigation
    Dim oGeneralData As SAPbobsCOM.GeneralData
    Dim oChild As SAPbobsCOM.GeneralData
    Dim oChildren As SAPbobsCOM.GeneralDataCollection


    Public Sub LoadMultiJForm()
        'CreateUserDS()
        Bind_CFLBP()
        CreateDTMJ()
    End Sub

    Public Sub Bind_CFLBP()
        If AddChooseFromList(MJobForm, "cflMultiBP", False, 2, "CardType", SAPbouiCOM.BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        MJobForm.Items.Item("ed_Cust").Specific.ChooseFromListUID = "cflMultiBP"
        MJobForm.Items.Item("ed_Cust").Specific.ChooseFromListAlias = "CardName"

    End Sub

    Public Sub CreateUserDS()
        If AddUserDataSrc(MJobForm, "FilterBy", sErrDesc, SAPbouiCOM.BoDataType.dt_SHORT_TEXT) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(MJobForm, "TODATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
        If AddUserDataSrc(MJobForm, "FDATE", sErrDesc, SAPbouiCOM.BoDataType.dt_DATE) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

        MJobForm.Items.Item("ed_frmDate").Specific.DataBind.SetBound(True, "", "FDATE")
        MJobForm.Items.Item("ed_toDate").Specific.DataBind.SetBound(True, "", "TODATE")
        MJobForm.Items.Item("cb_filter").Specific.DataBind.SetBound(True, "", "FilterBy")

        MJobForm.Items.Item("ed_toDate").Specific.Value = Today.Date.ToString("yyyyMMdd")
        oCombo = MJobForm.Items.Item("cb_filter").Specific
        oCombo.ValidValues.Add("Explosive", "Explosive")
        oCombo.ValidValues.Add("RadioActive", "RadioActive'")
        oCombo.ValidValues.Add("DG", "DG")
    End Sub

    Public Sub CreateDTMJ()
        CreateDTSelect()
        CreateDTAdd()
    End Sub

    Public Sub CreateDTSelect()

        ObjMatrix = MJobForm.Items.Item("mx_Select").Specific

        CreateTblStructure(dtmatrix, "dtSelect")
        dtmatrix = MJobForm.DataSources.DataTables.Item("dtSelect")
        dtmatrix.Rows.Add(1)
    End Sub

    Public Sub CreateTblStructure(ByVal dtmatrix As SAPbouiCOM.DataTable, ByVal tblname As String)
        Dim oColumn As SAPbouiCOM.Column
        MJobForm.DataSources.DataTables.Add(tblname)
        dtmatrix = MJobForm.DataSources.DataTables.Item(tblname)

       
        dtmatrix.Columns.Add("SJobNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("SType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("SCust", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
        dtmatrix.Columns.Add("Schk", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)

      
        oColumn = ObjMatrix.Columns.Item("colJobNo")
        oColumn.DataBind.Bind(tblname, "SJobNo")
        oColumn = ObjMatrix.Columns.Item("colType")
        oColumn.DataBind.Bind(tblname, "SType")
        oColumn = ObjMatrix.Columns.Item("colCust")
        oColumn.DataBind.Bind(tblname, "SCust")
        oColumn = ObjMatrix.Columns.Item("colChk")
        oColumn.DataBind.Bind(tblname, "Schk")

    End Sub

    Public Sub CreateDTAdd()
        ObjMatrix = MJobForm.Items.Item("mx_Add").Specific
        CreateTblStructure(dtmatrix, "dtAdd")
  

    End Sub

    Public Function DateTimeMJ(ByVal DGLP As SAPbouiCOM.Form, ByRef CtrlYear As SAPbouiCOM.EditText) As Boolean

        DateTimeMJ = False
        Dim SelectDate As String = CtrlYear.Value.ToString
        Dim Year, Month, Day, sErrDesc, ShowDay As String
        Dim sFuncName As String = String.Empty
        Dim [Datetiem1] As DateTime
        Try
            sFuncName = "DateTime()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            Year = SelectDate.Substring(0, 4).ToString()
            Month = SelectDate.Substring(4, 2).ToString()
            Day = SelectDate.Substring(6, 2).ToString()
            [Datetiem1] = New DateTime(Year, Month, Day)
            ShowDay = Datetiem1.ToString("ddd")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", sFuncName)
            DateTimeMJ = True
        Catch ex As Exception
            sErrDesc = ex.Message
            ShowErr(sErrDesc)
            DateTimeMJ = RTN_ERROR
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with ERROR", sFuncName)
            DateTimeMJ = False
        End Try
    End Function

    Public Sub Bind_MxSelect()
        oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ObjMatrix = MJobForm.Items.Item("mx_Select").Specific
        Dim sql As String = String.Empty
        Dim strFilter As String = ""
        Dim fromDate As String = ""
        Dim toDate As String = ""
        Dim customer As String = ""
        Dim jobNo As String = ""
        Dim tbl1 As String = ""
        Dim tbl2 As String = ""
        Dim tbl3 As String = ""
        If MJobForm.Items.Item("ed_frmDate").Specific.Value <> "" Then
            fromDate = MJobForm.Items.Item("ed_frmDate").Specific.Value.ToString.Substring(0, 4) + "-" + MJobForm.Items.Item("ed_frmDate").Specific.Value.ToString.Substring(4, 2) + "-" + MJobForm.Items.Item("ed_frmDate").Specific.Value.ToString.Substring(6, 2)
        End If
        If MJobForm.Items.Item("ed_toDate").Specific.Value <> "" Then
            toDate = MJobForm.Items.Item("ed_toDate").Specific.Value.ToString.Substring(0, 4) + "-" + MJobForm.Items.Item("ed_toDate").Specific.Value.ToString.Substring(4, 2) + "-" + MJobForm.Items.Item("ed_toDate").Specific.Value.ToString.Substring(6, 2)
        End If
        customer = MJobForm.Items.Item("ed_Cust").Specific.Value
        jobNo = MJobForm.Items.Item("ed_Job").Specific.Value
        If MJobForm.Items.Item("cb_filter").Specific.Value.ToString().Trim() = "General Cargo" Then
            strFilter = "GEN"
        ElseIf MJobForm.Items.Item("cb_filter").Specific.Value.ToString().Trim() = "Explosive" Then
            strFilter = "EXP"
        ElseIf MJobForm.Items.Item("cb_filter").Specific.Value.ToString().Trim() = "RadioActive" Then
            strFilter = "RA"
        ElseIf MJobForm.Items.Item("cb_filter").Specific.Value.ToString().Trim() = "DG" Then
            strFilter = "DG"
        ElseIf MJobForm.Items.Item("cb_filter").Specific.Value.ToString().Trim() = "Strategic" Then
            strFilter = "STG"
        ElseIf MJobForm.Items.Item("cb_filter").Specific.Value.ToString().Trim() = "" Then
            strFilter = ""
        End If
        tbl1 = "@OBT_FCL01_EXPORT"
        If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Then
            tbl2 = "@OBT_FCL03_ETRUCKING"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
            tbl2 = "@OBT_FCL04_EDISPATCH"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Fumigation" Then
            tbl2 = "@FUMIGATION"
            tbl3 = "@OBT_TBL01_FUMIGAT"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Outrider" Then
            tbl2 = "@OUTRIDER"
            tbl3 = "@OBT_TBL03_OUTRIDER"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Crane" Then
            tbl2 = "@CRANE"
            tbl3 = "@OBT_TB33_CRANE"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Forklift" Then
            tbl2 = "@FORKLIFT"
            tbl3 = "@OBT_TBL05_FORKLIFT"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Crate" Then 'to combine
            tbl2 = "@CRATE"
            tbl3 = "@OBT_TBL08_CRATE"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Bunker" Then
            tbl2 = "@BUNKER"
            tbl3 = "@OBT_TB01_BUNKER"
        ElseIf MJobForm.Items.Item("ed_Source").Specific.Value = "Toll" Then
            tbl2 = "@TOLL"
            tbl3 = "@OBT_TB01_TOLL"
        End If


        If strFilter <> "" And customer <> "" And jobNo <> "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_Name ='" & customer & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else 'Button 
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                  " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c on b.DocEntry=c.DocEntry  " & _
                  " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_Name ='" & customer & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If


        ElseIf strFilter = "" And customer = "" And jobNo = "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' and a.U_JobNum<>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b on a.DocEntry= b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry = c.DocEntry  " & _
                    " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' and a.U_JobNum<>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer = "" And jobNo = "" And fromDate = "" And toDate = "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"


            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                     " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                     " where a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer = "" And jobNo = "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' and a.U_JobNum<>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry " & _
                    " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' and a.U_JobNum<>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"


            End If

        ElseIf strFilter <> "" And customer = "" And jobNo = "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    "FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    "where  a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & toDate & "' and '" & toDate & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    "FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                    "where  a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & toDate & "' and '" & toDate & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"


            End If

        ElseIf strFilter <> "" And customer = "" And jobNo = "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter <> "" And customer <> "" And jobNo = "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & toDate & "' and '" & toDate & "' And a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                 " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & toDate & "' and '" & toDate & "' And a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter <> "" And customer <> "" And jobNo = "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                  " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter <> "" And customer = "" And jobNo <> "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & toDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                  " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & toDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter <> "" And customer = "" And jobNo <> "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                  " where a.U_CrgType='" & strFilter & "' And a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer <> "" And jobNo = "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                 " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer <> "" And jobNo = "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                 " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' and a.U_JobNum <>'" & MJobForm.Items.Item("ed_MainJob").Specific.Value & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer <> "" And jobNo <> "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                     " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                 " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer <> "" And jobNo <> "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then 'Tab
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' a.U_Name ='" & customer & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer = "" And jobNo <> "" And fromDate = "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                 " where a.U_JobDate between '" & toDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        ElseIf strFilter = "" And customer = "" And jobNo <> "" And fromDate <> "" And toDate <> "" Then
            If MJobForm.Items.Item("ed_Source").Specific.Value = "Trucking" Or MJobForm.Items.Item("ed_Source").Specific.Value = "Dispatch" Then
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a INNER JOIN [" & tbl2 & "] b ON b.DocEntry = a.DocEntry  " & _
                    " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (b.U_MultiJob<>'Y' or b.U_MultiJob is null) ORDER BY a.U_JobNum"

            Else
                sql = "SELECT Distinct a.U_JobNum,a.U_JobMode,a.U_Name" & _
                    " FROM [" & tbl1 & "] a LEFT JOIN [" & tbl2 & "] b ON a.DocEntry = b.U_DocNum LEFT JOIN [" & tbl3 & "] c ON b.DocEntry=c.DocEntry  " & _
                 " where a.U_JobDate between '" & fromDate & "' and '" & toDate & "' And a.U_JobNum='" & jobNo & "' and (c.U_MultiJob<>'Y' or c.U_MultiJob is null) ORDER BY a.U_JobNum"

            End If

        End If

        dtmatrix = MJobForm.DataSources.DataTables.Item("dtSelect")
        dtmatrix.ExecuteQuery(sql)
        ObjMatrix.LoadFromDataSource()

    End Sub
    Public Sub Bind_MxAdd()
        Dim omatrix As SAPbouiCOM.Matrix = MJobForm.Items.Item("mx_Select").Specific
        ObjMatrix = MJobForm.Items.Item("mx_Add").Specific
        dtmatrix = MJobForm.DataSources.DataTables.Item("dtAdd")
        Dim count As Integer = 0
        If omatrix.RowCount > 0 Then

            For i As Integer = 1 To omatrix.RowCount
                If omatrix.Columns.Item("colChk").Cells.Item(i).Specific.Checked = True Then

                    dtmatrix.Rows.Add(1)

                    dtmatrix.SetValue(0, count, omatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value)
                    dtmatrix.SetValue(1, count, omatrix.Columns.Item("colType").Cells.Item(i).Specific.Value)
                    dtmatrix.SetValue(2, count, omatrix.Columns.Item("colCust").Cells.Item(i).Specific.Value)
                    count = count + 1

                End If
            Next
            ObjMatrix.Clear()
            For i As Integer = 0 To dtmatrix.Rows.Count - 1

                If dtmatrix.Columns.Item("SJobNo").Cells.Item(i).Value.ToString() <> "" Then
                    ObjMatrix.AddRow()
                    ObjMatrix.Columns.Item("colJobNo").Cells.Item(i + 1).Specific.Value = dtmatrix.Columns.Item("SJobNo").Cells.Item(i).Value.ToString()
                    ObjMatrix.Columns.Item("colType").Cells.Item(i + 1).Specific.Value = dtmatrix.Columns.Item("SType").Cells.Item(i).Value.ToString()
                    ObjMatrix.Columns.Item("colCust").Cells.Item(i + 1).Specific.Value = dtmatrix.Columns.Item("SCust").Cells.Item(i).Value.ToString()
                End If
            Next

        End If
    End Sub
    Public Function SavetoMultiPOTable(ByVal pForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal PO As String, ByVal PODocNo As String) As Boolean
        SavetoMultiPOTable = False
        Dim lastDocEntry As Integer
        Dim sql As String = ""
        Try
            ' p_oDICompany.StartTransaction()
            For i As Integer = 1 To oMatrix.RowCount
                sql = "select top 1 Docentry from [@OBT_TB01_MULTIPO] order by docentry desc"
                oRecordSet.DoQuery(sql)
                If oRecordSet.Fields.Item("Docentry").Value.ToString = "" Then
                    lastDocEntry = 1
                Else
                    lastDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("Docentry").Value.ToString) + 1
                End If
                sql = "Select U_PO from [@OBT_TB01_MULTIPO] where U_JobNum='" & pForm.Items.Item("ed_JobNo").Specific.Value & "' and U_PO='" & pForm.Items.Item(PO).Specific.Value & "' "
                oRecordSet.DoQuery(sql)
                If oRecordSet.RecordCount = 0 Then
                    sql = "Insert Into [@OBT_TB01_MULTIPO] (DocEntry,DocNum,U_PO,U_PODocNo,U_JobNum) Values " & _
                    "(" & lastDocEntry & _
                        "," & lastDocEntry & _
                       "," & IIf(pForm.Items.Item(PO).Specific.Value <> "", FormatString(pForm.Items.Item(PO).Specific.Value), "NULL") & _
                        "," & IIf(pForm.Items.Item(PODocNo).Specific.Value <> "", FormatString(pForm.Items.Item(PODocNo).Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_JobNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JobNo").Specific.Value), "Null") & ")"
                    oRecordSet.DoQuery(sql)
                    lastDocEntry = lastDocEntry + 1
                End If

                sql = "Insert Into [@OBT_TB01_MULTIPO] (DocEntry,DocNum,U_PO,U_PODocNo,U_JobNum) Values " & _
                    "(" & lastDocEntry & _
                        "," & lastDocEntry & _
                       "," & IIf(pForm.Items.Item(PO).Specific.Value <> "", FormatString(pForm.Items.Item(PO).Specific.Value), "NULL") & _
                        "," & IIf(pForm.Items.Item(PODocNo).Specific.Value <> "", FormatString(pForm.Items.Item(PODocNo).Specific.Value), "NULL") & _
                    "," & IIf(oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value <> "", FormatString(oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value), "Null") & ")"
                oRecordSet.DoQuery(sql)

            Next
            'p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            SavetoMultiPOTable = True

        Catch ex As Exception
            SavetoMultiPOTable = False
            ' p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function LineAddtoTable(ByVal pForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal source As String) As Boolean
        LineAddtoTable = False
        Dim headerDocEntry As Integer
        Dim headerDocNum As Integer
        Dim lastLineDocEntry As Integer
        Dim InsDocNo As Integer
        Dim poListLine As Integer
        Dim sql As String = ""
        Try
            ' p_oDICompany.StartTransaction()
            For i As Integer = 1 To oMatrix.RowCount
                If source = "Trucking" Then
                    sql = "select top 1 LineId,U_InsDocNo,a.DocEntry from [@OBT_FCL01_EXPORT] a inner join [@OBT_FCL03_ETRUCKING] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.Fields.Item("U_InsDocNo").Value.ToString = "" Then
                        lastLineDocEntry = 1
                        InsDocNo = 1
                    Else
                        lastLineDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("LineId").Value.ToString) + 1
                        InsDocNo = Convert.ToInt32(oRecordSet.Fields.Item("U_InsDocNo").Value.ToString) + 1
                    End If
                    headerDocEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    'POList Table
                    sql = "select top 1 LineId,a.DocEntry,b.U_PONo from [@OBT_FCL01_EXPORT] a inner join [@OBT_TB01_POLIST] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.Fields.Item("U_PONo").Value.ToString = "" Then
                        poListLine = 1

                        sql = "Update [@OBT_TB01_POLIST]  set U_PONo='" & pForm.Items.Item("ed_PO").Specific.Value & "',U_VName='" & pForm.Items.Item("ed_Trucker").Specific.Value & "'" & _
                            ",U_PODate='" & pForm.Items.Item("ed_Date").Specific.Value & "',U_Desc='Trucking',U_POStatus='" & pForm.Items.Item("ed_PStus").Specific.Value & "'" & _
                            ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "' Where DocEntry='" & headerDocEntry & "' and LineId='1' "
                        oRecordSet.DoQuery(sql)
                    Else
                        poListLine = Convert.ToInt32(oRecordSet.Fields.Item("LineId").Value.ToString) + 1
                        sql = "Insert Into [@OBT_TB01_POLIST] (DocEntry,LineId,U_PONo,U_VName,U_PODate,U_Desc,U_POStatus,U_PODocNo) Values " & _
                           "(" & headerDocEntry & _
                          "," & poListLine & _
                          "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                          "," & IIf(pForm.Items.Item("ed_Trucker").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Trucker").Specific.Value), "NULL") & _
                          ",'" & pForm.Items.Item("ed_Date").Specific.Value & "'" & _
                         "," & FormatString(source) & _
                          "," & IIf(pForm.Items.Item("ed_PStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PStus").Specific.Value), "NULL") & _
                           "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & ")"
                        oRecordSet.DoQuery(sql)
                    End If

                    'Setting Table
                    sql = "select top 1 LineId,a.DocEntry,b.U_ICode from [@OBT_FCL01_EXPORT] a inner join [@OBT_TB01_TRKPOSET] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.Fields.Item("U_ICode").Value.ToString = "" Then
                        Dim IDesc As String = Replace(pForm.Items.Item("ed_POIDesc").Specific.Value, "'", "''")
                        sql = "Update [@OBT_TB01_TRKPOSET] set U_ICode='" & pForm.Items.Item("ed_POICode").Specific.Value & "', U_IDesc='" & IDesc & _
                            "', U_IQty= " & Convert.ToDouble(pForm.Items.Item("ed_POQty").Specific.Value) & ",U_IPrice=" & Convert.ToDouble(pForm.Items.Item("ed_POPrice").Specific.Value) & " Where DocEntry='" & headerDocEntry & "' and LineId='1'"
                        oRecordSet.DoQuery(sql)
                    End If

                    'Trucking Table

                    If lastLineDocEntry = 1 Then
                        sql = "Update [@OBT_FCL03_ETRUCKING]  set U_InsDocNo='" & InsDocNo & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_InsDate='" & pForm.Items.Item("ed_InsDate").Specific.Value & "'" & _
                            ",U_Mode='External',U_Trucker='" & pForm.Items.Item("ed_Trucker").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_Date").Specific.Value & "',U_PrepBy='" & p_oDICompany.UserName.ToString & "',U_Status='" & pForm.Items.Item("ed_PStus").Specific.Value & "'" & _
                        ",U_VehNo='" & pForm.Items.Item("ed_VehicNo").Specific.Value & "',U_EUC='" & pForm.Items.Item("ed_EUC").Specific.Value & "',U_Attent='" & pForm.Items.Item("ed_Attent").Specific.Value & "',U_Tel='" & pForm.Items.Item("ed_TkrTel").Specific.Value & "'" & _
                        ",U_Fax='" & pForm.Items.Item("ed_Fax").Specific.Value & "',U_Email='" & pForm.Items.Item("ed_Email").Specific.Value & "',U_TkrDate='" & pForm.Items.Item("ed_TkrDate").Specific.Value & "',U_TkrTime='" & pForm.Items.Item("ed_TkrTime").Specific.Value & "'" & _
                        ",U_ColFrm='" & pForm.Items.Item("ee_ColFrm").Specific.Value & "',U_TkrTo='" & pForm.Items.Item("ee_TkrTo").Specific.Value & "',U_TkrIns='" & pForm.Items.Item("ee_TkrIns").Specific.Value & "',U_InsRemsk='" & pForm.Items.Item("ee_InsRmsk").Specific.Value & "'" & _
                        ",U_Remarks='" & pForm.Items.Item("ee_Rmsk").Specific.Value & "'" & _
                       ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "',U_MultiJob='Y',U_TkrCode='" & pForm.Items.Item("ed_TkrCode").Specific.Value & "' Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    ElseIf lastLineDocEntry > 1 Then
                        sql = "Insert Into [@OBT_FCL03_ETRUCKING] (DocEntry,LineId,U_InsDocNo,U_PONo,U_Mode,U_Trucker,U_TkrCode,U_PODocNo,U_PO,U_VehNo,U_EUC,U_Attent,U_Tel,U_Fax,U_Email,U_TkrDate,U_TkrTime,U_ColFrm," & _
                            "U_TkrTo,U_TkrIns,U_InsRemsk,U_Remarks,U_PODate,U_PrepBy,U_Status, U_MultiJob,U_InsDate) Values " & _
                        "(" & headerDocEntry & _
                            "," & lastLineDocEntry & _
                           "," & FormatString(InsDocNo) & _
                            "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                            ",'External'" & _
                            "," & IIf(pForm.Items.Item("ed_Trucker").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Trucker").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_TkrCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TkrCode").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_VehicNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_VehicNo").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_EUC").Specific.Value <> "", FormatString(pForm.Items.Item("ed_EUC").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_Attent").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Attent").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_TkrTel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TkrTel").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_Fax").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Fax").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_Email").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Email").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_TkrDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TkrDate").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_TkrTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TkrTime").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_ColFrm").Specific.Value <> "", FormatString(pForm.Items.Item("ee_ColFrm").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_TkrTo").Specific.Value <> "", FormatString(pForm.Items.Item("ee_TkrTo").Specific.Value), "NULL") & _
                             "," & IIf(pForm.Items.Item("ee_TkrIns").Specific.Value <> "", FormatString(pForm.Items.Item("ee_TkrIns").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_InsRmsk").Specific.Value <> "", FormatString(pForm.Items.Item("ee_InsRmsk").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_Rmsk").Specific.Value <> "", FormatString(pForm.Items.Item("ee_Rmsk").Specific.Value), "NULL") & _
                        "," & IIf(pForm.Items.Item("ed_Date").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Date").Specific.Value), "NULL") & _
                             "," & IIf(p_oDICompany.UserName.ToString <> "", FormatString(p_oDICompany.UserName.ToString), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_PStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PStus").Specific.Value), "NULL") & _
                              ",'Y'" & _
                        "," & IIf(pForm.Items.Item("ed_InsDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_InsDate").Specific.Value), "NULL") & ")"
                    End If
                    oRecordSet.DoQuery(sql)


                ElseIf source = "Dispatch" Then
                    sql = "select top 1 LineId,U_InsDocNo,a.DocEntry from [@OBT_FCL01_EXPORT] a inner join [@OBT_FCL04_EDISPATCH] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.Fields.Item("U_InsDocNo").Value.ToString = "" Then
                        lastLineDocEntry = 1
                        InsDocNo = 1
                    Else
                        lastLineDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("LineId").Value.ToString) + 1
                        InsDocNo = Convert.ToInt32(oRecordSet.Fields.Item("U_InsDocNo").Value.ToString) + 1
                    End If
                    headerDocEntry = oRecordSet.Fields.Item("DocEntry").Value.ToString
                    'POList Table
                    sql = "select top 1 LineId,a.DocEntry,b.U_PONo from [@OBT_FCL01_EXPORT] a inner join [@OBT_TB01_POLIST] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.Fields.Item("U_PONo").Value.ToString = "" Then
                        poListLine = 1

                        sql = "Update [@OBT_TB01_POLIST]  set U_PONo='" & pForm.Items.Item("ed_DPO").Specific.Value & "',U_VName='" & pForm.Items.Item("ed_Dspatch").Specific.Value & "'" & _
                            ",U_PODate='" & pForm.Items.Item("ed_DDate").Specific.Value & "',U_Desc='Dispatch',U_POStatus='" & pForm.Items.Item("ed_DPStus").Specific.Value & "'" & _
                            ",U_PODocNo='" & pForm.Items.Item("ed_DPDocNo").Specific.Value & "' Where DocEntry='" & headerDocEntry & "' and LineId='1' "
                        oRecordSet.DoQuery(sql)
                    Else
                        poListLine = Convert.ToInt32(oRecordSet.Fields.Item("LineId").Value.ToString) + 1
                        sql = "Insert Into [@OBT_TB01_POLIST] (DocEntry,LineId,U_PONo,U_VName,U_PODate,U_Desc,U_POStatus,U_PODocNo) Values " & _
                           "(" & headerDocEntry & _
                          "," & poListLine & _
                          "," & IIf(pForm.Items.Item("ed_DPO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPO").Specific.Value), "NULL") & _
                          "," & IIf(pForm.Items.Item("ed_Dspatch").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Dspatch").Specific.Value), "NULL") & _
                          ",'" & pForm.Items.Item("ed_DDate").Specific.Value & "'" & _
                         "," & FormatString(source) & _
                          "," & IIf(pForm.Items.Item("ed_DPStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPStus").Specific.Value), "NULL") & _
                           "," & IIf(pForm.Items.Item("ed_DPDocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPDocNo").Specific.Value), "NULL") & ")"
                        oRecordSet.DoQuery(sql)
                    End If

                    'Setting Table
                    sql = "select top 1 LineId,a.DocEntry,b.U_ICode from [@OBT_FCL01_EXPORT] a inner join [@OBT_TB01_DSPPOSET] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                    oRecordSet.DoQuery(sql)
                    If oRecordSet.Fields.Item("U_ICode").Value.ToString = "" Then
                        sql = "Update [@OBT_TB01_DSPPOSET] set U_ICode='" & pForm.Items.Item("ed_DICode").Specific.Value & "', U_IDesc='" & pForm.Items.Item("ed_DIDesc").Specific.Value & _
                            "', U_IQty= " & Convert.ToDouble(pForm.Items.Item("ed_DQty").Specific.Value) & ",U_IPrice=" & Convert.ToDouble(pForm.Items.Item("ed_DPrice").Specific.Value) & " Where DocEntry='" & headerDocEntry & "' and LineId='1'"
                        oRecordSet.DoQuery(sql)
                    End If

                    'Trucking Table

                    If lastLineDocEntry = 1 Then
                        sql = "Update [@OBT_FCL04_EDISPATCH]  set U_InsDocNo='" & InsDocNo & "',U_PONo='" & pForm.Items.Item("ed_DPONo").Specific.Value & "',U_InsDate='" & pForm.Items.Item("ed_DInDate").Specific.Value & "'" & _
                            ",U_Mode='External',U_Dispatch='" & pForm.Items.Item("ed_Dspatch").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_DDate").Specific.Value & "',U_PrepBy='" & p_oDICompany.UserName.ToString & "',U_Status='" & pForm.Items.Item("ed_DPStus").Specific.Value & "'" & _
                        ",U_EUC='" & pForm.Items.Item("ed_DEUC").Specific.Value & "',U_Attent='" & pForm.Items.Item("ed_DAttent").Specific.Value & "',U_Tel='" & pForm.Items.Item("ed_DspTel").Specific.Value & "'" & _
                        ",U_Fax='" & pForm.Items.Item("ed_DFax").Specific.Value & "',U_Email='" & pForm.Items.Item("ed_DEmail").Specific.Value & "',U_DspDate='" & pForm.Items.Item("ed_DspDate").Specific.Value & "',U_DspTime='" & pForm.Items.Item("ed_DspTime").Specific.Value & "'" & _
                        ",U_DspIns='" & pForm.Items.Item("ee_DspIns").Specific.Value & "',U_InsRemsk='" & pForm.Items.Item("ee_DIRmsk").Specific.Value & "'" & _
                        ",U_Remarks='" & pForm.Items.Item("ee_DRmsk").Specific.Value & "'" & _
                       ",U_PODocNo='" & pForm.Items.Item("ed_DPDocNo").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_DPO").Specific.Value & "',U_MultiJob='Y',U_DspCode='" & pForm.Items.Item("ed_DspCode").Specific.Value & "' Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    ElseIf lastLineDocEntry > 1 Then
                        sql = "Insert Into [@OBT_FCL04_EDISPATCH] (DocEntry,LineId,U_InsDocNo,U_PONo,U_Mode,U_Dispatch,U_DspCode,U_PODocNo,U_PO,U_EUC,U_Attent,U_Tel,U_Fax,U_Email,U_DspDate,U_DspTime," & _
                            "U_DspIns,U_InsRemsk,U_Remarks,U_PODate,U_PrepBy,U_Status, U_MultiJob,U_InsDate) Values " & _
                        "(" & headerDocEntry & _
                            "," & lastLineDocEntry & _
                           "," & FormatString(InsDocNo) & _
                            "," & IIf(pForm.Items.Item("ed_DPONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPONo").Specific.Value), "NULL") & _
                            ",'External'" & _
                            "," & IIf(pForm.Items.Item("ed_Dspatch").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Dspatch").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DspCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DspCode").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DPDocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPDocNo").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DPO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPO").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DEUC").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DEUC").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DAttent").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DAttent").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DspTel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DspTel").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DFax").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DFax").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DEmail").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DEmail").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DspDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DspDate").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DspTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DspTime").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_DspIns").Specific.Value <> "", FormatString(pForm.Items.Item("ee_DspIns").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_DIRmsk").Specific.Value <> "", FormatString(pForm.Items.Item("ee_DIRmsk").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ee_DRmsk").Specific.Value <> "", FormatString(pForm.Items.Item("ee_DRmsk").Specific.Value), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DDate").Specific.Value), "NULL") & _
                             "," & IIf(p_oDICompany.UserName.ToString <> "", FormatString(p_oDICompany.UserName.ToString), "NULL") & _
                            "," & IIf(pForm.Items.Item("ed_DPStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DPStus").Specific.Value), "NULL") & _
                              ",'Y'" & _
                        "," & IIf(pForm.Items.Item("ed_DInDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_DInDate").Specific.Value), "NULL") & ")"
                    End If
                    oRecordSet.DoQuery(sql)

                    'ElseIf source = "Fumigation" Then
                    '    SaveToObject(pForm, oMatrix, source, i)
                    'ElseIf source = "Crane" Then
                    '    SaveToObject(pForm, oMatrix, source, i)
                    'ElseIf source = "Outrider" Then
                    '    SaveToObject(pForm, oMatrix, source, i)
                Else
                    SaveToObject(pForm, oMatrix, source, i)
                End If

            Next
            'p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            LineAddtoTable = True

        Catch ex As Exception
            ' p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            LineAddtoTable = False
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Public Function UpdateRemarkForMultiJob(ByVal pForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal rmkName As String) As Boolean
        UpdateRemarkForMultiJob = False
        ' Dim lastDocEntry As Integer
        Dim sql As String = ""
        Try

            For i As Integer = 1 To oMatrix.RowCount
                If pForm.Items.Item(rmkName).Specific.Value.ToString.Contains(pForm.Items.Item("ed_JobNo").Specific.Value) Then
                    pForm.Items.Item(rmkName).Specific.Value = pForm.Items.Item(rmkName).Specific.Value & "," & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value
                Else
                    pForm.Items.Item(rmkName).Specific.Value = pForm.Items.Item(rmkName).Specific.Value & Chr(13) & pForm.Items.Item("ed_JobNo").Specific.Value & "," & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value
                End If
            Next

            UpdateRemarkForMultiJob = True

        Catch ex As Exception

            p_oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
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

    Public Function UpdateMultiJobPOStatus(ByVal pForm As SAPbouiCOM.Form, ByVal poStatus As String, ByVal PO As String, ByVal dbsource As String) As Boolean
        UpdateMultiJobPOStatus = False
        Dim oMultijobRec As SAPbobsCOM.Recordset = Nothing
        Dim sql As String = ""
        Dim docEntry As Integer
        Try
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMultijobRec = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sql = "Select U_PO,U_JobNum from [@OBT_TB01_MULTIPO] Where U_PO='" & PO & "'"
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False
                    sql = "Select DocEntry from [@OBT_FCL01_EXPORT] WHERE U_JOBNUM='" & oRecordSet.Fields.Item("U_JobNum").Value & "'"
                    oMultijobRec.DoQuery(sql)
                    docEntry = oMultijobRec.Fields.Item("DocEntry").Value
                    sql = "Update " & dbsource & " set U_Status='" & poStatus & "' Where DocEntry=" & docEntry & " And U_PO='" & oRecordSet.Fields.Item("U_PO").Value & "'"
                    oMultijobRec.DoQuery(sql)
                    sql = "Update [@OBT_TB01_POLIST]  set U_POStatus='" & poStatus & "' Where DocEntry=" & docEntry & " And U_PONo='" & oRecordSet.Fields.Item("U_PO").Value & "'"
                    oMultijobRec.DoQuery(sql)
                    oRecordSet.MoveNext()
                End While

            End If

            UpdateMultiJobPOStatus = True
        Catch ex As Exception
            UpdateMultiJobPOStatus = False
        End Try
    End Function

    Public Function UpdateMultiJobPOStatusButton(ByVal pForm As SAPbouiCOM.Form, ByVal poStatus As String, ByVal PO As String, ByVal dbsource As String, ByVal tblHeader As String) As Boolean
        UpdateMultiJobPOStatusButton = False
        Dim oMultijobRec As SAPbobsCOM.Recordset = Nothing
        Dim sql As String = ""
        Dim docEntry As Integer
        Dim docNum As Integer
        Try
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMultijobRec = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            sql = "Select U_PO,U_JobNum from [@OBT_TB01_MULTIPO] Where U_PO='" & PO & "'"
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then

                oRecordSet.MoveFirst()
                While oRecordSet.EoF = False

                    sql = "Select DocEntry,U_DocNum from [" & tblHeader & "] WHERE U_JobNo='" & oRecordSet.Fields.Item("U_JobNum").Value & "'"
                    oMultijobRec.DoQuery(sql)
                    docEntry = oMultijobRec.Fields.Item("DocEntry").Value
                    docNum = oMultijobRec.Fields.Item("U_DocNum").Value

                    sql = "Update [" & dbsource & "] set U_Status='" & poStatus & "' Where DocEntry=" & docEntry & " And U_PO='" & oRecordSet.Fields.Item("U_PO").Value & "'"
                    oMultijobRec.DoQuery(sql)
                    sql = "Update [@OBT_TB01_POLIST]  set U_POStatus='" & poStatus & "' Where DocEntry=" & docNum & " And U_PONo='" & oRecordSet.Fields.Item("U_PO").Value & "'"
                    oMultijobRec.DoQuery(sql)
                    oRecordSet.MoveNext()
                End While
            End If

            UpdateMultiJobPOStatusButton = True
        Catch ex As Exception
            UpdateMultiJobPOStatusButton = False
        End Try
    End Function
    Public Function SaveToObject(ByVal pForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal source As String, ByVal i As Integer) As Boolean
        SaveToObject = False
        Dim sql As String = ""
        Dim lastLineDocEntry As Integer
        Dim headerDocNum As Integer
        Dim headerDocEntry As Integer
        Dim poListLine As Integer
        Dim tblHeader As String = ""
        Dim tblDetail As String = ""
        Dim tblItem As String = ""
        Dim objName As String = ""
        Dim tblDetail2 As String = ""
        Dim updateDocEntry As Integer
        Dim store1 As String = ""
        Dim store1a As String = ""
        Dim store2 As String = ""
        Dim store2a As String = ""

        Dim DocEntry As Integer
        Dim DMatrix As SAPbouiCOM.Matrix
        If source = "Fumigation" Then
            objName = "FUMIGATION"
            tblHeader = "FUMIGATION"
            tblDetail = "OBT_TBL01_FUMIGAT"
            tblItem = "OBT_TBL02_ITEM"
        ElseIf source = "Crane" Then
            objName = "CRANE"
            tblHeader = "CRANE"
            tblDetail = "OBT_TB33_CRANE"
            tblItem = "OBT_TB01_CRANEITEM"
            tblDetail2 = "OBT_TB01_CRNDETAIL"
        ElseIf source = "Outrider" Then '15/12/2011
            objName = "OUTRIDER"
            tblHeader = "OUTRIDER"
            tblDetail = "OBT_TBL03_OUTRIDER"
            tblItem = "OBT_TBL04_ITEM"
        ElseIf source = "Forklift" Then 'to combine
            objName = "FORKLIFT"
            tblHeader = "FORKLIFT"
            tblDetail = "OBT_TBL05_FORKLIFT"
            tblItem = "OBT_TBL06_FORKITEM"
            tblDetail2 = "OBT_TBL07_FORDETAIL" 'to combine
        ElseIf source = "Crate" Then 'to combine
            objName = "CRATE"
            tblHeader = "CRATE"
            tblDetail = "OBT_TBL08_CRATE"
            tblItem = "OBT_TBL09_CRAITEM"
            tblDetail2 = "OBT_TBL10_CRADETAIL" 'to combine
        ElseIf source = "Bunker" Then '15/12/2011
            objName = "BUNKER"
            tblHeader = "BUNKER"
            tblDetail = "OBT_TB01_BUNKER"
            tblItem = "OBT_TB01_BUNKITEM"
            tblDetail2 = "OBT_TB01_BNKDETAIL"
        ElseIf source = "Toll" Then '15/12/2011
            objName = "TOLL"
            tblHeader = "TOLL"
            tblDetail = "OBT_TB01_TOLL"
            tblItem = "OBT_TB01_TOLLITEM"
            tblDetail2 = "OBT_TB01_TOLLDETAIL"

        End If
        Try
            sql = "select  top 1 LineId,a.U_DocNum as DocNum,a.DocEntry as DocEntry,U_Vendor as Vendor from [@" & tblHeader & "] a inner join [@" & tblDetail & "]  b on a.DocEntry=b.DocEntry  where a.U_JobNo='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount = 0 Then
                lastLineDocEntry = 1
            ElseIf oRecordSet.Fields.Item("Vendor").Value.ToString = "" Then
                updateDocEntry = 1
                lastLineDocEntry = 1
                headerDocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value.ToString)
                headerDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value.ToString)
            Else
                lastLineDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("LineId").Value.ToString) + 1
                headerDocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value.ToString)
                headerDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value.ToString)
            End If



            If updateDocEntry <> 1 And lastLineDocEntry = 1 Then  ' Add New @Fumigation [Header] and @OBT_TBL01_FUMIGAT,@OBT_TBL02_ITEM [Child] tables
                sql = "select DocNum  from [@OBT_FCL01_EXPORT] where U_JobNum  ='" + oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value + "'"
                oRecordSet.DoQuery(sql)

                oGeneralService = p_oDICompany.GetCompanyService.GetGeneralService(objName)
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                oGeneralData.SetProperty("U_DocNum", oRecordSet.Fields.Item("DocNum").Value.ToString)
                oGeneralData.SetProperty("U_JobNo", oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value)

                oChildren = oGeneralData.Child(tblDetail)
                oChild = oChildren.Add
                oChild.SetProperty("U_Vendor", pForm.Items.Item("ed_FCode").Specific.Value)
                oChild.SetProperty("U_SIA", pForm.Items.Item("ed_FSIA").Specific.Value)
                oChild.SetProperty("U_PO", pForm.Items.Item("ed_PO").Specific.Value)
                oChild.SetProperty("U_Status", pForm.Items.Item("ed_FPOStus").Specific.Value)
                oChild.SetProperty("U_CreateBy", pForm.Items.Item("ed_Create").Specific.Value)
                oChild.SetProperty("U_JobNo", oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value)
                oChild.SetProperty("U_SIACode", pForm.Items.Item("ed_SIACode").Specific.Value)
                oChild.SetProperty("U_PODocNo", pForm.Items.Item("ed_PODocNo").Specific.Value)
                oChild.SetProperty("U_PONo", pForm.Items.Item("ed_PONo").Specific.Value)
                oChild.SetProperty("U_Vname", pForm.Items.Item("ed_FName").Specific.Value)
                oChild.SetProperty("U_TelNo", pForm.Items.Item("ed_SIATel").Specific.Value)
                If source = "Fumigation" Then 'button
                    oChild.SetProperty("U_IFumi", pForm.Items.Item("ed_Item").Specific.Value)
                    oChild.SetProperty("U_Loc", pForm.Items.Item("ed_Loc").Specific.Value)
                ElseIf source = "Crane" Then
                    oChild.SetProperty("U_SIns", pForm.Items.Item("ed_SIns").Specific.Value)
                    oChild.SetProperty("U_Loc", pForm.Items.Item("ed_Loc").Specific.Value)
                ElseIf source = "Outrider" Then
                    oChild.SetProperty("U_LocFrom", pForm.Items.Item("ed_LocFrom").Specific.Value)
                    oChild.SetProperty("U_LocTo", pForm.Items.Item("ed_LocTo").Specific.Value)
                    oChild.SetProperty("U_IRemark", pForm.Items.Item("ed_IRemark").Specific.Value)
                ElseIf source = "Forklift" Then 'to combine
                    oChild.SetProperty("U_Loc", pForm.Items.Item("ed_Loc").Specific.Value)
                    oChild.SetProperty("U_IRemark", pForm.Items.Item("ed_IRemark").Specific.Value)
                ElseIf source = "Crate" Then 'to combine
                    oChild.SetProperty("U_Desc", pForm.Items.Item("ed_Desc").Specific.Value)
                ElseIf source = "Bunker" Then
                    oChild.SetProperty("U_SIns", pForm.Items.Item("ed_SIns").Specific.Value)
                    oChild.SetProperty("U_CDesc", pForm.Items.Item("ed_CDesc").Specific.Value)

                    oChild.SetProperty("U_Active", pForm.Items.Item("cb_Act").Specific.Value.ToString.Trim)
                    oChild.SetProperty("U_TQty", pForm.Items.Item("ed_TQty").Specific.Value)
                    oChild.SetProperty("U_TKgs", pForm.Items.Item("ed_TKgs").Specific.Value)
                    oChild.SetProperty("U_TM3", pForm.Items.Item("ed_TM3").Specific.Value)
                    oChild.SetProperty("U_TNEQ", pForm.Items.Item("ed_TNEQ").Specific.Value)
                    If pForm.Items.Item("chk_1").Specific.Checked = True Then
                        oChild.SetProperty("U_Store1", "Y")
                    ElseIf pForm.Items.Item("chk_1a").Specific.Checked = True Then
                        oChild.SetProperty("U_Store1a", "Y")
                    ElseIf pForm.Items.Item("chk_2").Specific.Checked = True Then
                        oChild.SetProperty("U_Store2", "Y")
                    ElseIf pForm.Items.Item("chk_2a").Specific.Checked = True Then
                        oChild.SetProperty("U_Store2a", "Y")
                    End If
                ElseIf source = "Toll" Then
                    oChild.SetProperty("U_Loc", pForm.Items.Item("ed_Loc").Specific.Value)
                    oChild.SetProperty("U_IRemark", pForm.Items.Item("ed_IRemark").Specific.Value)
                End If

                oChild.SetProperty("U_Remark", pForm.Items.Item("ed_Remark").Specific.Value)
                oChild.SetProperty("U_VRef", pForm.Items.Item("ed_FVRef").Specific.Value)
                oChild.SetProperty("U_CPerson", pForm.Items.Item("ed_FCntact").Specific.Value)
                oChild.SetProperty("U_JDate", Today.Date)
                oChild.SetProperty("U_JTime", pForm.Items.Item("ed_FJbTime").Specific.Value)
                oChild.SetProperty("U_PODate", Today.Date)
                oChild.SetProperty("U_MultiJob", "Y")

                oChildren = oGeneralData.Child(tblItem)
                oChild = oChildren.Add
                If source <> "Toll" Then
                    oChild.SetProperty("U_ICode", pForm.Items.Item("ed_ICode").Specific.Value)
                    oChild.SetProperty("U_IDes", pForm.Items.Item("ed_IDesc").Specific.Value)
                Else
                    oChild.SetProperty("U_SCost", pForm.Items.Item("cb_SCost").Specific.Value)
                End If
                If source = "Fumigation" Or source = "Outrider" Then
                    oChild.SetProperty("U_IQty", pForm.Items.Item("ed_IQty").Specific.Value)
                    oChild.SetProperty("U_IPrice", pForm.Items.Item("ed_IPrice").Specific.Value)
                ElseIf source = "Bunker" Then
                    oChild.SetProperty("U_IPrice", pForm.Items.Item("ed_IPrice").Specific.Value)
                End If

                oGeneralService.Add(oGeneralData)

                sql = "select  top 1 LineId,a.U_DocNum as DocNum,a.DocEntry as DocEntry from [@" & tblHeader & "] a inner join [@" & tblDetail & "]  b on a.DocEntry=b.DocEntry  where a.U_JobNo='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                oRecordSet.DoQuery(sql)
                headerDocNum = Convert.ToInt32(oRecordSet.Fields.Item("DocNum").Value.ToString)
                headerDocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value.ToString)
            ElseIf updateDocEntry = 1 Then
                If source = "Bunker" Then
                    If pForm.Items.Item("chk_1").Specific.Checked = True Then
                        store1 = "Y"
                    ElseIf pForm.Items.Item("chk_1a").Specific.Checked = True Then
                        store1a = "Y"
                    ElseIf pForm.Items.Item("chk_2").Specific.Checked = True Then
                        store2 = "Y"
                    ElseIf pForm.Items.Item("chk_2a").Specific.Checked = True Then
                        store2a = "Y"
                    End If

                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_SIns='" & pForm.Items.Item("ed_SIns").Specific.Value & "',U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_CDesc='" & pForm.Items.Item("ed_CDesc").Specific.Value & "'" & _
                           ",U_Active='" & pForm.Items.Item("cb_Act").Specific.Value.ToString.Trim & "',U_TQty='" & pForm.Items.Item("ed_TQty").Specific.Value & "',U_TKgs='" & pForm.Items.Item("ed_TKgs").Specific.Value & "'" & _
                          ",U_TM3='" & pForm.Items.Item("ed_TM3").Specific.Value & "',U_TNEQ='" & pForm.Items.Item("ed_TNEQ").Specific.Value & "',U_Store1='" & store1 & "'" & _
                          ",U_Store1a='" & store1a & "',U_Store2='" & store2 & "',U_Store2a='" & store2a & "'" & _
                          ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                ElseIf source = "Crane" Then
                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_SIns='" & pForm.Items.Item("ed_SIns").Specific.Value & "',U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_Loc='" & pForm.Items.Item("ed_Loc").Specific.Value & "'" & _
                           ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                ElseIf source = "Fumigation" Then
                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_Loc='" & pForm.Items.Item("ed_Loc").Specific.Value & "'" & _
                           ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                ElseIf source = "Outrider" Then 'to combine
                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_IRemark='" & pForm.Items.Item("ed_IRemark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_LocFrom='" & pForm.Items.Item("ed_LocFrom").Specific.Value & "',U_LocTo='" & pForm.Items.Item("ed_LocTo").Specific.Value & "'" & _
                           ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                ElseIf source = "Forklift" Then 'to combine
                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_IRemark='" & pForm.Items.Item("ed_IRemark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_Loc='" & pForm.Items.Item("ed_Loc").Specific.Value & "'" & _
                           ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                ElseIf source = "Crate" Then 'to combine
                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_Desc='" & pForm.Items.Item("ed_Desc").Specific.Value & "'" & _
                           ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                ElseIf source = "Toll" Then
                    sql = "Update [@" & tblDetail & "]  set U_Vendor='" & pForm.Items.Item("ed_FCode").Specific.Value & "',U_SIA='" & pForm.Items.Item("ed_FSIA").Specific.Value & "',U_PO='" & pForm.Items.Item("ed_PO").Specific.Value & "'" & _
                               ",U_Status='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "',U_CreateBy='" & pForm.Items.Item("ed_Create").Specific.Value & "',U_JobNo='" & pForm.Items.Item("ed_JDeNo").Specific.Value & "',U_SIACode='" & pForm.Items.Item("ed_SIACode").Specific.Value & "'" & _
                           ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "',U_PONo='" & pForm.Items.Item("ed_PONo").Specific.Value & "',U_Vname='" & pForm.Items.Item("ed_FName").Specific.Value & "',U_TelNo='" & pForm.Items.Item("ed_SIATel").Specific.Value & "'" & _
                           ",U_IRemark='" & pForm.Items.Item("ed_IRemark").Specific.Value & "',U_Remark='" & pForm.Items.Item("ed_Remark").Specific.Value & "',U_VRef='" & pForm.Items.Item("ed_FVRef").Specific.Value & "',U_CPerson='" & pForm.Items.Item("ed_FCntact").Specific.Value & "'" & _
                           ",U_JDate='" & pForm.Items.Item("ed_FJbDate").Specific.Value & "',U_JTime='" & pForm.Items.Item("ed_FJbTime").Specific.Value & "',U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_Loc='" & pForm.Items.Item("ed_Loc").Specific.Value & "'" & _
                           ",U_MultiJob='Y'" & _
                          " Where DocEntry='" & headerDocEntry & "' and LineId='" & lastLineDocEntry & "' "
                    oRecordSet.DoQuery(sql)
                End If

            ElseIf lastLineDocEntry > 1 Then
                If source = "Fumigation" Then
                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_IFumi,U_Remark,U_VRef,U_CPerson,U_JDate," & _
                    "U_JTime,U_PODate,U_Loc,U_MultiJob) Values " & _
                "(" & headerDocEntry & _
                    "," & lastLineDocEntry & _
                    "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_Item").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Item").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                     "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
                "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_Loc").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Loc").Specific.Value), "NULL") & _
                    ",'Y')"
                ElseIf source = "Crane" Then
                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_SIns,U_Remark,U_VRef,U_CPerson,U_JDate," & _
                   "U_JTime,U_PODate,U_Loc,U_MultiJob) Values " & _
               "(" & headerDocEntry & _
                   "," & lastLineDocEntry & _
                   "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIns").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIns").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
               "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Loc").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Loc").Specific.Value), "NULL") & _
                   ",'Y')"

                ElseIf source = "Outrider" Then
                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_Remark,U_IRemark,U_VRef,U_CPerson,U_JDate," & _
                    "U_JTime,U_PODate,U_LocFrom,U_LocTo,U_MultiJob) Values " & _
                "(" & headerDocEntry & _
                    "," & lastLineDocEntry & _
                    "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_IRemark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_IRemark").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                     "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
                "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_LocFrom").Specific.Value <> "", FormatString(pForm.Items.Item("ed_LocTo").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_LocTo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_LocTo").Specific.Value), "NULL") & _
                    ",'Y')"
                ElseIf source = "Forklift" Then  'to combine
                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_Remark,U_IRemark,U_VRef,U_CPerson,U_JDate," & _
                   "U_JTime,U_PODate,U_Loc,U_MultiJob) Values " & _
               "(" & headerDocEntry & _
                   "," & lastLineDocEntry & _
                   "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_IRemark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_IRemark").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
               "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Loc").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Loc").Specific.Value), "NULL") & _
                   ",'Y')"
                ElseIf source = "Crate" Then  'to combine
                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_Remark,U_VRef,U_CPerson,U_JDate," & _
                   "U_JTime,U_PODate,U_Desc,U_MultiJob) Values " & _
               "(" & headerDocEntry & _
                   "," & lastLineDocEntry & _
                   "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
               "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Desc").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Desc").Specific.Value), "NULL") & _
                   ",'Y')"
                ElseIf source = "Bunker" Then

                    If pForm.Items.Item("chk_1").Specific.Checked = True Then
                        store1 = "Y"
                    ElseIf pForm.Items.Item("chk_1a").Specific.Checked = True Then
                        store1a = "Y"
                    ElseIf pForm.Items.Item("chk_2").Specific.Checked = True Then
                        store2 = "Y"
                    ElseIf pForm.Items.Item("chk_2a").Specific.Checked = True Then
                        store2a = "Y"
                    End If

                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_SIns,U_Remark,U_VRef,U_CPerson,U_JDate," & _
                   "U_JTime,U_PODate,U_CDesc,U_Active,U_TQty,U_TKgs,U_TM3,U_TNEQ,U_Store1,U_Store1a,U_Store2,U_Store2a,U_MultiJob) Values " & _
               "(" & headerDocEntry & _
                   "," & lastLineDocEntry & _
                   "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIns").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIns").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_CDesc").Specific.Value <> "", FormatString(pForm.Items.Item("ed_CDesc").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("cb_Act").Specific.Value.ToString.Trim <> "", FormatString(pForm.Items.Item("cb_Act").Specific.Value.ToString.Trim), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_TQty").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TQty").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_TKgs").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TKgs").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_TM3").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TM3").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_TNEQ").Specific.Value <> "", FormatString(pForm.Items.Item("ed_TNEQ").Specific.Value), "NULL") & _
                   "," & IIf(store1 <> "", FormatString(store1), "NULL") & _
                   "," & IIf(store1a <> "", FormatString(store1a), "NULL") & _
                   "," & IIf(store2 <> "", FormatString(store2), "NULL") & _
                   "," & IIf(store2a <> "", FormatString(store2a), "NULL") & _
                   ",'Y')"

                ElseIf source = "Toll" Then
                    sql = "Insert Into [@" & tblDetail & "] (DocEntry,LineId,U_Vendor,U_SIA,U_PO,U_Status,U_CreateBy,U_JobNo,U_SIACode,U_PODocNo,U_PONo,U_Vname,U_TelNo,U_IRemark,U_Remark,U_VRef,U_CPerson,U_JDate," & _
                   "U_JTime,U_PODate,U_Loc,U_MultiJob) Values " & _
               "(" & headerDocEntry & _
                   "," & lastLineDocEntry & _
                   "," & IIf(pForm.Items.Item("ed_FCode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FSIA").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FSIA").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Create").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Create").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_JDeNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_JDeNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIACode").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIACode").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PONo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PONo").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_SIATel").Specific.Value <> "", FormatString(pForm.Items.Item("ed_SIATel").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_IRemark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_IRemark").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Remark").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Remark").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FVRef").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FVRef").Specific.Value), "NULL") & _
                    "," & IIf(pForm.Items.Item("ed_FCntact").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FCntact").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbDate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbDate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_FJbTime").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FJbTime").Specific.Value), "NULL") & _
               "," & IIf(pForm.Items.Item("ed_FPODate").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPODate").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_Loc").Specific.Value <> "", FormatString(pForm.Items.Item("ed_Loc").Specific.Value), "NULL") & _
                   ",'Y')"

                End If
            End If
            oRecordSet.DoQuery(sql)
            'POList Table
            sql = "select top 1 LineId,a.DocEntry,b.U_PONo from [@OBT_FCL01_EXPORT] a inner join [@OBT_TB01_POLIST] b on a.DocEntry=b.DocEntry where a.U_JobNum='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
            oRecordSet.DoQuery(sql)
            If oRecordSet.Fields.Item("U_PONo").Value.ToString = "" Then
                poListLine = 1

                sql = "Update [@OBT_TB01_POLIST]  set U_PONo='" & pForm.Items.Item("ed_PO").Specific.Value & "',U_VName='" & pForm.Items.Item("ed_FName").Specific.Value & "'" & _
                   ",U_PODate='" & pForm.Items.Item("ed_FPODate").Specific.Value & "',U_Desc='" & source & "',U_POStatus='" & pForm.Items.Item("ed_FPOStus").Specific.Value & "'" & _
                   ",U_PODocNo='" & pForm.Items.Item("ed_PODocNo").Specific.Value & "' Where DocEntry='" & headerDocNum & "' and LineId='1' "
                oRecordSet.DoQuery(sql)
            Else
                poListLine = Convert.ToInt32(oRecordSet.Fields.Item("LineId").Value.ToString) + 1
                sql = "Insert Into [@OBT_TB01_POLIST] (DocEntry,LineId,U_PONo,U_VName,U_PODate,U_Desc,U_POStatus,U_PODocNo) Values " & _
                   "(" & headerDocNum & _
                  "," & poListLine & _
                  "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                  "," & IIf(pForm.Items.Item("ed_FName").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FName").Specific.Value), "NULL") & _
                  ",'" & pForm.Items.Item("ed_FPODate").Specific.Value & "'" & _
                 "," & FormatString(source) & _
                  "," & IIf(pForm.Items.Item("ed_FPOStus").Specific.Value <> "", FormatString(pForm.Items.Item("ed_FPOStus").Specific.Value), "NULL") & _
                   "," & IIf(pForm.Items.Item("ed_PODocNo").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PODocNo").Specific.Value), "NULL") & ")"
                oRecordSet.DoQuery(sql)
            End If

            'Setting Table
            If source <> "Toll" Then
                sql = "select top 1 LineId,a.DocEntry,b.U_ICode from [@" & tblHeader & "] a inner join [@" & tblItem & "] b on a.DocEntry=b.DocEntry where a.U_JobNo='" & oMatrix.Columns.Item("colJobNo").Cells.Item(i).Specific.Value & "' order by LineId desc  "
                oRecordSet.DoQuery(sql)
                If oRecordSet.Fields.Item("U_ICode").Value.ToString = "" Then
                    If source = "Fumigation" Or source = "Outrider" Then
                        sql = "Update [@" & tblItem & "] set U_ICode='" & pForm.Items.Item("ed_ICode").Specific.Value & "', U_IDes='" & pForm.Items.Item("ed_IDesc").Specific.Value & _
                        "', U_IQty= " & Convert.ToDouble(pForm.Items.Item("ed_IQty").Specific.Value) & ",U_IPrice=" & Convert.ToDouble(pForm.Items.Item("ed_IPrice").Specific.Value) & " Where DocEntry='" & headerDocEntry & "' and LineId='1'"
                    ElseIf source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                        sql = "Update [@" & tblItem & "] set U_ICode='" & pForm.Items.Item("ed_ICode").Specific.Value & "', U_IDes='" & pForm.Items.Item("ed_IDesc").Specific.Value & " Where DocEntry='" & headerDocEntry & "' and LineId='1'"
                    ElseIf source = "Bunker" Then
                        sql = "Update [@" & tblItem & "] set U_ICode='" & pForm.Items.Item("ed_ICode").Specific.Value & "', U_IDes='" & pForm.Items.Item("ed_IDesc").Specific.Value & _
                            ",U_IPrice=" & Convert.ToDouble(pForm.Items.Item("ed_IPrice").Specific.Value) & " Where DocEntry='" & headerDocEntry & "' and LineId='1'"
                    End If
                    oRecordSet.DoQuery(sql)
                End If
            End If


            If source = "Crane" Or source = "Forklift" Or source = "Crate" Then
                oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [@" & tblDetail2 & "]")
                If oRecordSet.RecordCount > 0 Then
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                End If
                DMatrix = pForm.Items.Item("mx_CDetail").Specific
                Dim j As Integer
                If DMatrix.RowCount > 0 Then
                    If source = "Crane" Then
                        For j = 1 To DMatrix.RowCount
                            sql = "Insert Into " + "[@" + tblDetail2 + "]" + " (DocEntry,LineID,U_CType,U_Ton,U_Desc,U_Qty,U_UOM,U_Price,U_Hrs,U_Total,U_Remark,U_PO,U_DocN) Values " & _
                                   "(" & DocEntry & _
                                    "," & j & _
                                    "," & IIf(DMatrix.Columns.Item("colCType").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colCType").Cells.Item(j).Specific.Value), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colTon").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colTon").Cells.Item(j).Specific.Value), "Null") & _
                                    "," & IIf(DMatrix.Columns.Item("colDesc").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colDesc").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colHr").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colHr").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colRmk").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colRmk").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                    "," & IIf(headerDocNum.ToString <> "", FormatString(headerDocNum.ToString), "NULL") & ")"
                            oRecordSet.DoQuery(sql)

                        Next
                    ElseIf source = "Forklift" Then 'to combine
                        For j = 1 To DMatrix.RowCount
                            sql = "Insert Into " + "[@" + tblDetail2 + "]" + " (DocEntry,LineID,U_Ton,U_Desc,U_Qty,U_UOM,U_Price,U_Total,U_Remark,U_PO,U_DocN) Values " & _
                                   "(" & DocEntry & _
                                    "," & j & _
                                    "," & IIf(DMatrix.Columns.Item("colTon").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colTon").Cells.Item(j).Specific.Value), "Null") & _
                                    "," & IIf(DMatrix.Columns.Item("colDesc").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colDesc").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colRmk").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colRmk").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                    "," & IIf(headerDocNum.ToString <> "", FormatString(headerDocNum.ToString), "NULL") & ")"
                            oRecordSet.DoQuery(sql)

                        Next
                    ElseIf source = "Crate" Then 'to combine
                        For j = 1 To DMatrix.RowCount
                            sql = "Insert Into " + "[@" + tblDetail2 + "]" + " (DocEntry,LineID,U_Dimen,U_Type,U_Qty,U_Price,U_Total,U_PO,U_DocN) Values " & _
                                   "(" & DocEntry & _
                                    "," & j & _
                                    "," & IIf(DMatrix.Columns.Item("colDimen").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colDimen").Cells.Item(j).Specific.Value()), "NULL") & _
                                     "," & IIf(DMatrix.Columns.Item("colType").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colType").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value()), "NULL") & _
                                    "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                    "," & IIf(headerDocNum.ToString <> "", FormatString(headerDocNum.ToString), "NULL") & ")"
                            oRecordSet.DoQuery(sql)

                        Next
                    End If

                End If
            ElseIf source = "Bunker" Then
                oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [@" & tblDetail2 & "]")
                If oRecordSet.RecordCount > 0 Then
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                End If
                DMatrix = pForm.Items.Item("mx_BDetail").Specific
                Dim j As Integer
                If DMatrix.RowCount > 0 Then
                    For j = 1 To DMatrix.RowCount
                        sql = "Insert Into [@" & tblDetail2 & "] (DocEntry,LineID,U_Permit,U_JobNo,U_Client,U_Qty,U_UOM,U_Kgs,U_M3,U_NEQ,U_Stus,U_PO,U_DocN) Values " & _
                               "(" & DocEntry & _
                                "," & j & _
                                "," & IIf(DMatrix.Columns.Item("colPermit").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colPermit").Cells.Item(j).Specific.Value), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colJobNo").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colJobNo").Cells.Item(j).Specific.Value), "Null") & _
                                "," & IIf(DMatrix.Columns.Item("colClient").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colClient").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colUOM").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colKgs").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colKgs").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colM3").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colM3").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colNEQ").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colNEQ").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colStus").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colStus").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                "," & IIf(headerDocNum.ToString <> "", FormatString(headerDocNum.ToString), "NULL") & ")"
                        oRecordSet.DoQuery(sql)
                    Next
                End If
            ElseIf source = "Toll" Then
                oRecordSet.DoQuery("Select isnull(Max(DocEntry),0)+1 As DocEntry From [@" & tblDetail2 & "]")
                If oRecordSet.RecordCount > 0 Then
                    DocEntry = Convert.ToInt32(oRecordSet.Fields.Item("DocEntry").Value)
                End If
                DMatrix = pForm.Items.Item("mx_TDetail").Specific
                Dim j As Integer
                If DMatrix.RowCount > 0 Then
                    For j = 1 To DMatrix.RowCount
                        sql = "Insert Into [@" & tblDetail2 & "] (DocEntry,LineID,U_ICode,U_IDesc,U_Qty,U_Price,U_Total,U_PO,U_DocN) Values " & _
                               "(" & DocEntry & _
                                "," & j & _
                                "," & IIf(DMatrix.Columns.Item("colICode").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colICode").Cells.Item(j).Specific.Value), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colIDesc").Cells.Item(j).Specific.Value <> "", FormatString(DMatrix.Columns.Item("colIDesc").Cells.Item(j).Specific.Value), "Null") & _
                                "," & IIf(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colQty").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colPrice").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value() <> "", FormatString(DMatrix.Columns.Item("colTotal").Cells.Item(j).Specific.Value()), "NULL") & _
                                "," & IIf(pForm.Items.Item("ed_PO").Specific.Value <> "", FormatString(pForm.Items.Item("ed_PO").Specific.Value), "NULL") & _
                                "," & IIf(headerDocNum.ToString <> "", FormatString(headerDocNum.ToString), "NULL") & ")"
                        oRecordSet.DoQuery(sql)
                    Next
                End If
            End If

            SaveToObject = True
        Catch ex As Exception
            SaveToObject = False
        End Try
    End Function

End Module

