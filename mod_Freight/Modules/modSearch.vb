Option Explicit On

Imports System.Xml
Imports SAPbobsCOM
Imports SAPbouiCOM

Module modSearch
    Private dtSearhList As DataTable

    Public Function DoSearchFormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Boolean
        ' **********************************************************************************
        '   Function    :   DoSearchFormDataEvent
        '   Purpose     :   This function will be providing to proceed validating and form level data processing
        '                   SearchForm Form Data Event event    
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.BusinessObjectInfo
        '                       pVal =  set the SAP UI Menu Event Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   FALSE - FAILURE
        '                   TRUE - SUCCESS
        ' **********************************************************************************
        DoSearchFormDataEvent = False
        Try
            DoSearchFormDataEvent = True
        Catch ex As Exception

            DoSearchFormDataEvent = False
        End Try
    End Function

    Public Function DoSearchFormMenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Boolean
        ' **********************************************************************************
        '   Function    :   ImportSeaLCLMenuEvent
        '   Purpose     :   This function will be providing to handle all Menu Event of SearchForm
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.MenuEvent
        '                       pVal =  set the SAP UI Menu Event Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   FALSE - FAILURE
        '                   TRUE - SUCCESS
        ' ***********************************************************************************
        DoSearchFormMenuEvent = False
        Dim oSearchForm As SAPbouiCOM.Form = Nothing
        Dim oComboBox As SAPbouiCOM.ComboBox = Nothing
        'Dim oColumns As SAPbouiCOM.DataColumns
        Dim oGrid As SAPbouiCOM.Grid
        Dim FunctionName As String = "DoSearchFormMenuEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            Select Case pVal.MenuUID
                Case "mnuSearchForm"
                    If pVal.BeforeAction = False Then
                        If Not AlreadyExist("SEARCHFORM") Then
                            If Not LoadFromXML(p_oSBOApplication, "SearchForm.srf") Then Throw New ArgumentException(sErrDesc)
                            oSearchForm = p_oSBOApplication.Forms.Item("SEARCHFORM")
                            oSearchForm.Freeze(True)
                            oSearchForm.DataSources.DataTables.Add("SEARCHLIST")
                            dtSearhList = oSearchForm.DataSources.DataTables.Item("SEARCHLIST")
                            dtSearhList.Columns.Add("Job No", BoFieldsType.ft_AlphaNumeric)
                            dtSearhList.Columns.Add("Date", BoFieldsType.ft_Date)
                            dtSearhList.Columns.Add("Customer", BoFieldsType.ft_AlphaNumeric)
                            dtSearhList.Columns.Add("Shipping Agent", BoFieldsType.ft_AlphaNumeric)
                            oGrid = oSearchForm.Items.Item("gd_JobList").Specific
                            oGrid.DataTable = dtSearhList

                            ' --------------------------- Adding DataSource --------------------------- 
                            AddUserDataSrc(oSearchForm, "JOBNO", sErrDesc, BoDataType.dt_SHORT_TEXT, 15)
                            AddUserDataSrc(oSearchForm, "FROMDATE", sErrDesc, BoDataType.dt_DATE)
                            AddUserDataSrc(oSearchForm, "TODATE", sErrDesc, BoDataType.dt_DATE)
                            AddUserDataSrc(oSearchForm, "CUSTOMER", sErrDesc, BoDataType.dt_SHORT_TEXT, 100)
                            AddUserDataSrc(oSearchForm, "VENDOR", sErrDesc, BoDataType.dt_SHORT_TEXT, 100)
                            AddUserDataSrc(oSearchForm, "STATUS", sErrDesc, BoDataType.dt_SHORT_TEXT, 11)
                            ' --------------------------- Adding DataSource ---------------------------- 

                            'ed_JobNo,ed_FrDate,ed_ToDate,ed_Cust,ed_ShpAgt,cb_Status
                            oSearchForm.Items.Item("ed_JobNo").Specific.DataBind.SetBound(True, "", "JOBNO")
                            oSearchForm.Items.Item("ed_FrDate").Specific.DataBind.SetBound(True, "", "FROMDATE")
                            oSearchForm.Items.Item("ed_ToDate").Specific.DataBind.SetBound(True, "", "TODATE")
                            oSearchForm.Items.Item("ed_Cust").Specific.DataBind.SetBound(True, "", "CUSTOMER")
                            oSearchForm.Items.Item("ed_ShpAgt").Specific.DataBind.SetBound(True, "", "VENDOR")
                            oSearchForm.Items.Item("cb_Status").Specific.DataBind.SetBound(True, "", "STATUS")
                            oComboBox = oSearchForm.Items.Item("cb_Status").Specific
                            oComboBox.ValidValues.Add(vbNullString, "")
                            oComboBox.ValidValues.Add("Open", "Opened")
                            oComboBox.ValidValues.Add("Closed", "Closed")
                            oComboBox.ValidValues.Add("Cancelled", "Cancelled")
                            'need to add default value
                            oComboBox.Select("Open")

                            If AddChooseFromList(oSearchForm, "CFLCUST", False, 2, "CardType", BoConditionOperation.co_EQUAL, "C") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            oSearchForm.Items.Item("ed_Cust").Specific.ChooseFromListUID = "CFLCUST"
                            oSearchForm.Items.Item("ed_Cust").Specific.ChooseFromListAlias = "CardName"
                            If AddChooseFromList(oSearchForm, "CFLVEN", False, 2, "CardType", BoConditionOperation.co_EQUAL, "S") <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                            oSearchForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListUID = "CFLVEN"
                            oSearchForm.Items.Item("ed_ShpAgt").Specific.ChooseFromListAlias = "CardName"
                            oSearchForm.Freeze(False)
                        Else
                            p_oSBOApplication.Forms.Item("SEARCHFORM").Select()
                        End If
                    End If
            End Select

            DoSearchFormMenuEvent = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", FunctionName)
        Catch ex As Exception
            DoSearchFormMenuEvent = False
        End Try
    End Function

    Public Function DoSearchFormItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Boolean
        ' **********************************************************************************
        '   Function    :   DoImportSeaLCLItemEvent
        '   Purpose     :   This function will be providing to handle all item event of SearchForm
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.ItemEvent
        '                       pVal =  set the SAP UI Menu Event Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   FALSE - FAILURE
        '                   TRUE - SUCCESS
        ' ***********************************************************************************
        DoSearchFormItemEvent = False
        Dim oSearchForm As SAPbouiCOM.Form = Nothing
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oJobList As SAPbouiCOM.Grid
        'Dim oComboBox As SAPbouiCOM.ComboBox
        Try
            Select Case pVal.FormTypeEx
                Case "2000000200"
                    oSearchForm = p_oSBOApplication.Forms.Item("SEARCHFORM")
                    oRecordSet = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                    If pVal.BeforeAction = False Then
                        If pVal.EventType = BoEventTypes.et_FORM_CLOSE Then
                            If Not RemoveFromAppList(oSearchForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.EventType = BoEventTypes.et_ITEM_PRESSED Then
                            If pVal.ItemUID = "bt_Close" Then
                                oSearchForm.Close()
                            End If

                            If pVal.ItemUID = "bt_Search" Then
                                Dim oSearch As clsSearch
                                oSearch = New clsSearch(oSearchForm, "SEARCHLIST", "gd_JobList", "ed_JobNo", "ed_FrDate", "ed_ToDate", "ed_Cust", "ed_ShpAgt", "cb_Status")
                                oSearch.SearchByOptions()
                            End If

                            If pVal.ItemUID = "bt_NewJob" Then
                                If Not AlreadyExist("IGNITER") Then
                                    If Not modCreateNewJob.LoadAndCreateJobForm() Then Throw New ArgumentException(sErrDesc)
                                Else
                                    p_oSBOApplication.Forms.Item("IGNITER").Select()
                                End If
                            End If

                            If pVal.ItemUID = "gd_JobList" And (pVal.ColUID = "Job No" Or pVal.ColUID = "Job Date" Or _
                                                                pVal.ColUID = "Status" Or pVal.ColUID = "Customer" Or _
                                                                pVal.ColUID = "Shipping Agent") Then
                                oJobList = oSearchForm.Items.Item("gd_JobList").Specific
                                oJobList.Rows.SelectedRows.Add(pVal.Row)
                                'p_oSBOApplication.MessageBox(oJobList.Rows.SelectedRows.Item(pVal.Row, BoOrderType.ot_RowOrder).ToString)
                            End If
                        End If

                        If pVal.EventType = BoEventTypes.et_DOUBLE_CLICK Then
                            If pVal.ItemUID = "gd_JobList" And (pVal.ColUID = "Job No" Or pVal.ColUID = "Job Date" Or _
                                                                pVal.ColUID = "Status" Or pVal.ColUID = "Customer" Or _
                                                                pVal.ColUID = "Shipping Agent") Then
                                oJobList = oSearchForm.Items.Item("gd_JobList").Specific
                                Dim JType As String = ""
                                If Left(oJobList.DataTable.GetValue("JobType", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), 6) = "Import" Then
                                    JType = "Import"
                                ElseIf Left(oJobList.DataTable.GetValue("JobType", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), 6) = "Export" Then
                                    JType = "Export"
                                ElseIf Left(oJobList.DataTable.GetValue("JobType", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), 5) = "Local" Then
                                    JType = "Local"
                                ElseIf Left(oJobList.DataTable.GetValue("JobType", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), 12) = "Transhipment" Then
                                    JType = "Transhipment"
                                End If
                                If Not AlreadyExist("EXPORTSEAFCL") Then
                                    modExportSeaFCL.LoadExportSeaFCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                                                         JType, _
                                                                    oJobList.DataTable.GetValue("JobType", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                                                     SAPbouiCOM.BoFormMode.fm_FIND_MODE)


                                Else
                                    p_oSBOApplication.Forms.Item("EXPORTSEAFCL").Select()
                                End If

                                'If oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "IMPORTSEALCL" Then
                                '    If Not AlreadyExist("IMPORTSEALCL") Then
                                '        modImportSeaLCL.LoadImportSeaLCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                             "Import Sea LCL", _
                                '                                             SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '    Else
                                '        p_oSBOApplication.Forms.Item("IMPORTSEALCL").Select()
                                '    End If
                                'ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "IMPORTSEAFCL" Then
                                '    If Not AlreadyExist("IMPORTSEAFCL") Then
                                '        modImportSeaFCL.LoadImportSeaFCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                              "Import Sea FCL", _
                                '                                             SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '    Else
                                '        p_oSBOApplication.Forms.Item("IMPORTSEAFCL").Select()
                                '    End If
                                'ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "IMPORTAIR" Then
                                '    If Not AlreadyExist("IMPORTAIR") Then
                                '        modImportSeaLCL.LoadImportSeaLCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                              "Import Air", _
                                '                                             SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '    Else
                                '        p_oSBOApplication.Forms.Item("IMPORTAIR").Select()
                                '    End If
                                'ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "IMPORTLAND" Then
                                '    If Not AlreadyExist("IMPORTLAND") Then
                                '        modImportSeaLCL.LoadImportSeaLCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                              "Import Land", _
                                '                                             SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '    Else
                                '        p_oSBOApplication.Forms.Item("IMPORTLAND").Select()
                                '    End If
                                'ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "EXPORTSEALCL" Then
                                '    If Not AlreadyExist("EXPORTSEALCL") Then
                                '        If Not AlreadyExist("EXPORTAIRLCL") Then
                                '            modExportSeaLCL.LoadExportSeaLCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                             "Export Sea LCL", _
                                '                                             SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '        Else
                                '            p_oSBOApplication.Forms.Item("EXPORTAIRLCL").Select()
                                '        End If

                                '    Else
                                '        p_oSBOApplication.Forms.Item("EXPORTSEALCL").Select()
                                '    End If
                                'ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "EXPORTSEAFCL" Then

                                '    If Not AlreadyExist("EXPORTSEAFCL") Then
                                '        modExportSeaFCL.LoadExportSeaFCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                        oJobList.DataTable.GetValue("JobMode", oJobList.Rows.SelectedRows.Item(6, BoOrderType.ot_RowOrder)), _
                                '                                         SAPbouiCOM.BoFormMode.fm_FIND_MODE)


                                '    Else
                                '        p_oSBOApplication.Forms.Item("EXPORTSEAFCL").Select()
                                '    End If

                                '    'For ExportSeaFCL
                                'ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "EXPORTAIRLCL" Then
                                '    If Not AlreadyExist("EXPORTSEALCL") Then
                                '        If Not AlreadyExist("EXPORTAIRLCL") Then
                                '            modExportSeaLCL.LoadExportSeaLCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                             "Export Air LCL", _
                                '                                             SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '        Else
                                '            p_oSBOApplication.Forms.Item("EXPORTAIRLCL").Select()
                                '        End If

                                '    Else
                                '        p_oSBOApplication.Forms.Item("EXPORTSEALCL").Select()
                                '    End If
                                '    'For ExportSeaFCL
                                '    ElseIf oJobList.DataTable.GetValue("ObjectName", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)) = "EXPORTLAND" Then
                                '        If Not AlreadyExist("EXPORTLAND") Then
                                '            If Not AlreadyExist("EXPORTLAND") Then
                                '                modExportSeaLCL.LoadExportSeaLCLForm(oJobList.DataTable.GetValue("Job No", oJobList.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder)), _
                                '                                                 "Export Land LCL", _
                                '                                                 SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                '            Else
                                '                p_oSBOApplication.Forms.Item("EXPORTLAND").Select()
                                '            End If

                                '        Else
                                '            p_oSBOApplication.Forms.Item("EXPORTLAND").Select()
                                '        End If

                                '    End If
                            End If
                        End If

                        If pVal.EventType = BoEventTypes.et_CHOOSE_FROM_LIST Then
                            Dim CFLForm As SAPbouiCOM.Form = p_oSBOApplication.Forms.ActiveForm
                            Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent = pVal
                            Dim oDataTable As SAPbouiCOM.DataTable = oCFLEvento.SelectedObjects
                            If pVal.ItemUID = "ed_Cust" Then
                                oSearchForm.DataSources.UserDataSources.Item("CUSTOMER").ValueEx = oDataTable.GetValue(1, 0).ToString
                            End If
                            If pVal.ItemUID = "ed_ShpAgt" Then
                                oSearchForm.DataSources.UserDataSources.Item("VENDOR").ValueEx = oDataTable.GetValue(1, 0).ToString
                            End If
                        End If
                    End If

            End Select

            DoSearchFormItemEvent = True
        Catch ex As Exception
            DoSearchFormItemEvent = False
        End Try
    End Function

    Public Function DoSearchFormRightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Boolean
        ' **********************************************************************************
        '   Function    :   DoImportSeaLCLItemEvent
        '   Purpose     :   This function is provider for SearchForm's  Right Click Event
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
        DoSearchFormRightClickEvent = False
        Try

            DoSearchFormRightClickEvent = True
        Catch ex As Exception
            DoSearchFormRightClickEvent = False
        End Try
    End Function
End Module
