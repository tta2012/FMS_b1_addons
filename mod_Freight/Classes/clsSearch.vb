Option Explicit On

Imports SAPbobsCOM
Imports SAPbouiCOM

Public Class clsSearch
    Private Const SQLOR As String = "OR"
    Private Const SQLAND As String = "AND"
    Private oActiveForm As Form = Nothing
    Private oSearchList As SAPbouiCOM.DataTable = Nothing
    Private oGrid As Grid = Nothing
    Private JobNo, FromDate, ToDate, CustomerName, VendorName As EditText
    Private Status As ComboBox

    Private Enum SQLCombineCondition
        None = 0
        JobNoAndCustomer = 1 'JC
        JobNoAndVendor = 2 'JV
        JobNoAndStatus = 3 'JS
        JobNoAndDate = 4 'JD
        CustomerAndVendor = 5 'CV
        CustomerAndStatus = 6  'CS
        CustomerAndDate = 7 'CD
        VendorAndStatus = 8  'VS
        VendorAndDate = 9 'VD
        StatusAndDate = 10

        JobNoAndCustomerAndVendor = 11 'JCV
        JobNoAndCustomerAndStatus = 12 'JCS
        JobNoAndVendorAndStatus = 13 'JVS
        JobNoAndCustomerAndDate = 14 'JCD
        JobNoAndVendorAndDate = 15 'JVD
        JobNoAndStatusAndDate = 16 'JSD
        CustomerAndVendorAndStatus = 17 'CVS
        CustomerAndVendorAndDate = 18 'CVD
        CustomerAndStatusAndDate = 19 'CSD
        VendorAndStatusAndDate = 20 'VSD

        JobNoAndCustomerAndVendorAndStatus = 21 'JCVS
        JobNoAndCustomerAndVendorAndDate = 22 ' JCVD
        JobNoAndCustomerAndStatusAndDate = 23 'JCSD
        JobNoAndVendorAndStatusAndDate = 24 'JVSD
        CustomerAndVendorAndStatusAndDate = 25 'CVSD

        JobNo = 26
        CustomerName = 27
        VendorName = 28
        Status = 29
        _Date = 30
        All = 31 'All
    End Enum

    Public ReadOnly Property SearchList As SAPbouiCOM.DataTable
        Get
            Return oSearchList
        End Get
    End Property

    Public Sub New(ByVal oForm As SAPbouiCOM.Form)
        oActiveForm = oForm
    End Sub

    Public Sub New(ByVal oForm As SAPbouiCOM.Form, ByVal DataTableUID As String, ByVal GridUID As String, _
                   ByVal JobNoUID As String, ByVal FromDateUID As String, ByVal ToDateUID As String, _
                   ByVal CustNameUID As String, ByVal VenNameUID As String, ByVal StatusUID As String)
        oActiveForm = oForm
        oSearchList = oActiveForm.DataSources.DataTables.Item(DataTableUID)
        oGrid = oActiveForm.Items.Item(GridUID).Specific
        JobNo = oActiveForm.Items.Item(JobNoUID).Specific
        FromDate = oActiveForm.Items.Item(FromDateUID).Specific
        ToDate = oActiveForm.Items.Item(ToDateUID).Specific
        CustomerName = oActiveForm.Items.Item(CustNameUID).Specific
        VendorName = oActiveForm.Items.Item(VenNameUID).Specific
        Status = oActiveForm.Items.Item(StatusUID).Specific
    End Sub

    Private Function SearchFilter() As String
        SearchFilter = vbNullString
        Try
            If JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                SearchFilter = ConditionString(SQLCombineCondition.All)
            Else
                If JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                    SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomerAndVendorAndStatus)  'JCVS
                ElseIf JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                    SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomerAndVendorAndDate)    'JCVD
                ElseIf JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                    SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomerAndStatusAndDate)    'JCSD
                ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                    SearchFilter = ConditionString(SQLCombineCondition.JobNoAndVendorAndStatusAndDate)      'JVSD
                ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                    SearchFilter = ConditionString(SQLCombineCondition.CustomerAndVendorAndStatusAndDate)   'CVSD
                Else
                    If JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                        SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomerAndVendor)       'JCV
                    ElseIf JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                        SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomerAndStatus)       'JCS
                    ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                        SearchFilter = ConditionString(SQLCombineCondition.JobNoAndVendorAndStatus)         'JVS
                    ElseIf JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                        SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomerAndDate)         'JCD
                    ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                        SearchFilter = ConditionString(SQLCombineCondition.JobNoAndVendorAndDate)           'JVD
                    ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                        SearchFilter = ConditionString(SQLCombineCondition.JobNoAndStatusAndDate)           'JSD
                    ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                        SearchFilter = ConditionString(SQLCombineCondition.CustomerAndVendorAndStatus)      'CVS
                    ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                        SearchFilter = ConditionString(SQLCombineCondition.CustomerAndVendorAndDate)        'CVD
                    ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                        SearchFilter = ConditionString(SQLCombineCondition.CustomerAndStatusAndDate)        'CSD
                    ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                        SearchFilter = ConditionString(SQLCombineCondition.VendorAndStatusAndDate)          'VSD
                    Else
                        If JobNo.Value <> vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                            SearchFilter = ConditionString(SQLCombineCondition.JobNoAndCustomer)
                        ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                            SearchFilter = ConditionString(SQLCombineCondition.JobNoAndVendor)
                        ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                            SearchFilter = ConditionString(SQLCombineCondition.JobNoAndStatus)
                        ElseIf JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                            SearchFilter = ConditionString(SQLCombineCondition.JobNoAndDate)
                        ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                            SearchFilter = ConditionString(SQLCombineCondition.CustomerAndVendor)
                        ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                            SearchFilter = ConditionString(SQLCombineCondition.CustomerAndStatus)
                        ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                            SearchFilter = ConditionString(SQLCombineCondition.CustomerAndDate)
                        ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                            SearchFilter = ConditionString(SQLCombineCondition.VendorAndStatus)
                        ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                            SearchFilter = ConditionString(SQLCombineCondition.VendorAndDate)
                        ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                            SearchFilter = ConditionString(SQLCombineCondition.StatusAndDate)
                        Else
                            If JobNo.Value <> vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                                SearchFilter = ConditionString(SQLCombineCondition.JobNo)
                            ElseIf JobNo.Value = vbNullString And CustomerName.Value <> vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                                SearchFilter = ConditionString(SQLCombineCondition.CustomerName)
                            ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value <> vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                                SearchFilter = ConditionString(SQLCombineCondition.VendorName)
                            ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value <> vbNullString And ValidDatePeriod(FromDate, ToDate) = False Then
                                SearchFilter = ConditionString(SQLCombineCondition.Status)
                            ElseIf JobNo.Value = vbNullString And CustomerName.Value = vbNullString And VendorName.Value = vbNullString And Status.Value = vbNullString And ValidDatePeriod(FromDate, ToDate) = True Then
                                SearchFilter = ConditionString(SQLCombineCondition._Date)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            SearchFilter = vbNullString
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Sub SearchByOptions()
        Dim strSQL, strFilters As String
        Try
            strFilters = SearchFilter()
            If strFilters <> "" Then
                strSQL = "SELECT U_JobNo As [Job No],U_JbDate As [Job Date],U_JbStus As [Status],U_CusName As [Customer],U_ShpName As [Shipping Agent],U_ObjType As [ObjectName],U_JobType As [JobType] FROM [@OBT_FREIGHTDOCNO] WHERE " + strFilters
                p_oSBOApplication.Forms.ActiveForm.Freeze(True)
                oSearchList.ExecuteQuery(strSQL)
                oGrid.SelectionMode = BoMatrixSelect.ms_Single
                oGrid.Columns.Item("Job No").Editable = False
                oGrid.Columns.Item("Job No").BackColor = 16645629
                oGrid.Columns.Item("Job Date").Editable = False
                oGrid.Columns.Item("Job Date").BackColor = 16645629
                oGrid.Columns.Item("Status").Editable = False
                oGrid.Columns.Item("Status").BackColor = 16645629
                oGrid.Columns.Item("Customer").Editable = False
                oGrid.Columns.Item("Customer").BackColor = 16645629
                oGrid.Columns.Item("Shipping Agent").Editable = False
                oGrid.Columns.Item("Shipping Agent").BackColor = 16645629
                oGrid.Columns.Item("ObjectName").Visible = False
                oGrid.Columns.Item("JobType").Visible = False
                oGrid.AutoResizeColumns()
                p_oSBOApplication.Forms.ActiveForm.Freeze(False)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Function ConditionString(ByVal Condition As SQLCombineCondition) As String
        ConditionString = vbNullString
        Dim ByJobNo As String = "(U_JobNo =" + FormatString(JobNo.Value) + ")"
        Dim ByDate As String = "(U_JbDate BETWEEN " + FormatString(FromDate.Value) + "AND" + FormatString(ToDate.Value) + ")"
        Dim ByCustomer As String = "(U_CusName =" + FormatString(CustomerName.Value) + ")"
        Dim ByVendor As String = "(U_ShpName =" + FormatString(VendorName.Value) + ")"
        Dim ByStatus As String = "(U_JbStus =" + FormatString(Status.Value).Trim + ")"
        Try
            Select Case Condition
                Case SQLCombineCondition.JobNoAndCustomer
                    ConditionString = ByJobNo + SQLAND + ByCustomer
                Case SQLCombineCondition.JobNoAndVendor
                    ConditionString = ByJobNo + SQLAND + ByVendor
                Case SQLCombineCondition.JobNoAndStatus
                    ConditionString = ByJobNo + SQLAND + ByStatus
                Case SQLCombineCondition.JobNoAndDate
                    ConditionString = ByJobNo + SQLAND + ByDate
                Case SQLCombineCondition.CustomerAndVendor
                    ConditionString = ByCustomer + SQLAND + ByVendor
                Case SQLCombineCondition.CustomerAndStatus
                    ConditionString = ByCustomer + SQLAND + ByStatus
                Case SQLCombineCondition.CustomerAndDate
                    ConditionString = ByCustomer + SQLAND + ByDate
                Case SQLCombineCondition.VendorAndStatus
                    ConditionString = ByVendor + SQLAND + ByStatus
                Case SQLCombineCondition.VendorAndDate
                    ConditionString = ByVendor + SQLAND + ByDate
                Case SQLCombineCondition.StatusAndDate
                    ConditionString = ByStatus + SQLAND + ByDate
                    ' ========================================
                Case SQLCombineCondition.JobNoAndCustomerAndVendor
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByVendor
                Case SQLCombineCondition.JobNoAndCustomerAndStatus
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByStatus
                Case SQLCombineCondition.JobNoAndVendorAndStatus
                    ConditionString = ByJobNo + SQLAND + ByVendor + SQLAND + ByStatus
                Case SQLCombineCondition.JobNoAndCustomerAndDate
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByDate
                Case SQLCombineCondition.JobNoAndVendorAndDate
                    ConditionString = ByJobNo + SQLAND + ByVendor + SQLAND + ByDate
                Case SQLCombineCondition.JobNoAndStatusAndDate
                    ConditionString = ByJobNo + SQLAND + ByStatus + SQLAND + ByDate
                Case SQLCombineCondition.CustomerAndVendorAndStatus
                    ConditionString = ByCustomer + SQLAND + ByVendor + SQLAND + ByStatus
                Case SQLCombineCondition.CustomerAndVendorAndDate
                    ConditionString = ByCustomer + SQLAND + ByVendor + SQLAND + ByDate
                Case SQLCombineCondition.CustomerAndStatusAndDate
                    ConditionString = ByCustomer + SQLAND + ByStatus + SQLAND + ByDate
                Case SQLCombineCondition.VendorAndStatusAndDate
                    ConditionString = ByVendor + SQLAND + ByStatus + SQLAND + ByDate
                    ' ===================================
                Case SQLCombineCondition.JobNoAndCustomerAndVendorAndStatus
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByVendor + SQLAND + ByStatus
                Case SQLCombineCondition.JobNoAndCustomerAndVendorAndDate
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByVendor + SQLAND + ByDate
                Case SQLCombineCondition.JobNoAndCustomerAndStatusAndDate
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByStatus + SQLAND + ByDate
                Case SQLCombineCondition.JobNoAndVendorAndStatusAndDate
                    ConditionString = ByJobNo + SQLAND + ByVendor + SQLAND + ByStatus + SQLAND + ByDate
                Case SQLCombineCondition.CustomerAndVendorAndStatusAndDate
                    ConditionString = ByCustomer + SQLAND + ByVendor + SQLAND + ByStatus + SQLAND + ByDate
                Case SQLCombineCondition.All
                    ConditionString = ByJobNo + SQLAND + ByCustomer + SQLAND + ByVendor + SQLAND + ByStatus + SQLAND + ByDate
                Case SQLCombineCondition.JobNo
                    ConditionString = ByJobNo
                Case SQLCombineCondition.CustomerName
                    ConditionString = ByCustomer
                Case SQLCombineCondition.VendorName
                    ConditionString = ByVendor
                Case SQLCombineCondition.Status
                    ConditionString = ByStatus
                Case SQLCombineCondition._Date
                    ConditionString = ByDate
            End Select
        Catch ex As Exception
            ConditionString = vbNullString
        End Try
    End Function

    Private Function ValidDatePeriod(ByRef oFromDate As SAPbouiCOM.EditText, ByRef oToDate As SAPbouiCOM.EditText) As Boolean
        ValidDatePeriod = False
        If oFromDate.Value <> "" And oToDate.Value <> "" Then
            ValidDatePeriod = True
        Else
            If oFromDate.Value = "" Then
                ValidDatePeriod = False
                'p_oSBOApplication.MessageBox("Need to fill FromDate!")
            ElseIf oToDate.Value = "" Then
                ValidDatePeriod = False
                'p_oSBOApplication.MessageBox("Need to fill ToDate")
            Else
                ValidDatePeriod = False
                'p_oSBOApplication.MessageBox("Need to fill Dates")
            End If
        End If
    End Function
End Class
