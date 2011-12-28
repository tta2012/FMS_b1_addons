Option Explicit On 

'This module will have common functions

Module modCommon

    Public Function ConnectDICompSSO(ByRef objCompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ConnectDICompSSO
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for add folder like tab .
        '               
        '   Parameters  : ByRef objCompany As SAPbobsCOM.Company
        '                       objCompany =  set the SAP company object  to be returned to calling function                                        
        '               :  ByVal sErrDesc As String
        '                    sErrDesc= to show error message in sap appliction 
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sCookie As String
        Dim sConnStr As String
        Dim lRetval As Long
        Dim iErrCode As Int32
        Dim sFuncName As String = "ConnectDICompSSO()"
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim str As String

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            objCompany = New SAPbobsCOM.Company
            sCookie = objCompany.GetContextCookie
            sConnStr = p_oUICompany.GetConnectionContext(sCookie)
            lRetval = objCompany.SetSboLoginContext(sConnStr)

            If Not lRetval = 0 Then
                Throw New ArgumentException("SetSboLoginContext of Single SignOn Failed.")
            End If

            lRetval = objCompany.Connect
            If lRetval <> 0 Then
                objCompany.GetLastError(iErrCode, sErrDesc)
                Throw New ArgumentException("Connect of Single SignOn failed : " & sErrDesc)
            Else
                oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRecordSet.DoQuery("select CODE,NAME,U_DOCUPATH,U_SERVER,U_DBNAME,U_UID,U_PWD,U_PicPath from [@FMSSETTING]")
                If oRecordSet.RecordCount > 0 Then
                    p_fmsSetting.DocuPath = oRecordSet.Fields.Item("U_DOCUPATH").Value.ToString
                    p_fmsSetting.ServerName = oRecordSet.Fields.Item("U_SERVER").Value.ToString
                    p_fmsSetting.DBName = oRecordSet.Fields.Item("U_DBNAME").Value.ToString
                    p_fmsSetting.UserID = oRecordSet.Fields.Item("U_UID").Value.ToString
                    p_fmsSetting.Password = oRecordSet.Fields.Item("U_PWD").Value.ToString
                    p_fmsSetting.PicturePath = oRecordSet.Fields.Item("U_PicPath").Value.ToString
                End If
                'str = p_fmsSetting.UserID & "," & p_fmsSetting.Password & "," & p_fmsSetting.DBName & "," & p_fmsSetting.ServerName
                'p_oSBOApplication.MessageBox(str)
            End If
            ConnectDICompSSO = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with SUCCESS", sFuncName)
        Catch exc As Exception
            sErrDesc = exc.Message
            ConnectDICompSSO = RTN_ERROR
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        End Try

    End Function

    Public Sub ShowErr(ByVal sErrMsg As String)

        ' **********************************************************************************
        '   Function    :   ShowErr()
        '   Purpose     :   This function will be providing to show error message for
        '                   ImportSeall Form .
        '               
        '   Parameters  :   ByVal sErrMsg As String
        '                      
        '   Return      :   No
        ' **********************************************************************************
        Try
            If sErrMsg <> "" Then
                If Not p_oSBOApplication Is Nothing Then
                    If p_iErrDispMethod = ERR_DISPLAY_STATUS Then
                        p_oSBOApplication.SetStatusBarMessage("Error : " & sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short)
                    ElseIf p_iErrDispMethod = ERR_DISPLAY_DIALOGUE Then
                        p_oSBOApplication.MessageBox("Error : " & sErrMsg)
                    End If
                End If
            End If
        Catch exc As Exception
            WriteToLogFile(exc.Message, "ShowErr()")
        End Try
    End Sub

    Public Function AddButton(ByRef oForm As SAPbouiCOM.Form, _
                              ByVal sButtonID As String, _
                              ByVal sCaption As String, _
                              ByVal sItemNo As String, _
                              ByVal iSpacing As Integer, _
                              ByVal iWidth As Integer, _
                              ByVal blnVisable As Boolean, _
                              ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   AddButton
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for button control .
        '               
        '   Parameters  :  ByRef oForm As SAPbouiCOM.Form
        '                       oForm =  set the SAP form  to be returned to calling function                                               
        '               :  ByVal sButtonID As String
        '                       sButtonID = to use button id in sap appliction to  be returned function.
        '               :  ByVal sCaption As String
        '                       sCaption=button caption to be returned to calling function 
        '               :  ByRef sErrDesc As String) As Long
        '                       sErrDesc= to show error message in sap application.
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oItems As SAPbouiCOM.Items
        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String = "AddButton()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oItems = oForm.Items
            oItem = oItems.Add(sButtonID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)

            oItem.Specific.Caption = sCaption
            oItem.Visible = blnVisable
            oItem.Left = oItems.Item(sItemNo).Left + oItems.Item(sItemNo).Width + iSpacing
            oItem.Height = oItems.Item(sItemNo).Height
            oItem.Top = oItems.Item(sItemNo).Top
            oItem.Width = iWidth

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            AddButton = RTN_SUCCESS
        Catch exc As Exception
            AddButton = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oItems = Nothing
            oItem = Nothing
        End Try
    End Function

    Public Function AddFolder(ByRef oForm As SAPbouiCOM.Form, ByVal sItemUID As String, _
                              ByVal sItemCaption As String, ByVal sBeforeItemUID As String, _
                              ByVal sUDS As String, ByVal bVisible As Boolean, ByRef sErrDesc As String, _
                              Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0) As Long
        ' **********************************************************************************
        '   Function    :   AddFolder
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for add folder like tab .
        '               
        '   Parameters  :  ByRef oForm As SAPbouiCOM.Form
        '                       oForm =  set the SAP form  to be returned to calling function                                               
        '               :   ByVal sItemUID As String
        '                       sItemUID = to use folder id in sap appliction.
        '                    ByVal sItemCaption As String
        '                    sItemCaption=folder caption to be returned to calling function        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************


        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String
        Try
            sFuncName = "AddFolder()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oForm.DataSources.UserDataSources.Add(sUDS, SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)

            oItem = oForm.Items.Add(sItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.AffectsFormMode = False
            oItem.Specific.Caption = sItemCaption
            oItem.Specific.DataBind.SetBound(True, "", sUDS)
            oItem.Specific.GroupWith(sBeforeItemUID)        'Group with existing folders
            oItem.Left = oForm.Items.Item(sBeforeItemUID).Left + 1
            oItem.Visible = bVisible
            'Channyeinkyaw
            oItem.FromPane = iFromPane
            oItem.ToPane = iToPane

            AddFolder = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddFolder = RTN_ERROR
            sErrDesc = exc.Message
            sFuncName = "AddFolder()"
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oItem = Nothing
        End Try
    End Function

    Public Function AddingFolderWithPos(ByRef oForm As SAPbouiCOM.Form, ByVal sItemUID As String, _
                              ByVal sItemCaption As String, ByVal sBeforeItemUID As String, _
                              ByVal sUDS As String, ByVal bVisible As Boolean, ByRef sErrDesc As String, _
                              Optional ByVal lLeft As Int16 = 0, Optional ByVal lTop As Int16 = 0, _
                              Optional ByVal lHeight As Int16 = 0, Optional ByVal lWidth As Int16 = 0, _
                              Optional ByVal iFromPane As Int16 = 0, Optional ByVal iToPane As Int16 = 0) As Long


        ' **********************************************************************************
        '   Function    :   AddingFolderWithPos
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for add folder like tab .
        '               
        '   Parameters  :  ByRef oForm As SAPbouiCOM.Form
        '                       oForm =  set the SAP form  to be returned to calling function                                               
        '               :   ByVal sItemUID As String
        '                       sItemUID = to use folder id in sap appliction.
        '                    ByVal sItemCaption As String
        '                    sItemCaption=folder caption to be returned to calling function        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String
        Try
            sFuncName = "AddingFolder()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oForm.DataSources.UserDataSources.Add(sUDS, SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)

            oItem = oForm.Items.Add(sItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.AffectsFormMode = False
            oItem.Specific.Caption = sItemCaption
            oItem.Specific.DataBind.SetBound(True, "", sUDS)
            oItem.Specific.GroupWith(sBeforeItemUID)        'Group with existing folders
            oItem.Left = oForm.Items.Item(sBeforeItemUID).Left + 1
            oItem.Visible = bVisible
            oItem.FromPane = iFromPane
            oItem.ToPane = iToPane

            AddingFolderWithPos = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddingFolderWithPos = RTN_ERROR
            sErrDesc = exc.Message
            sFuncName = "AddingFolder()"
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oItem = Nothing
        End Try
    End Function

    Public Function AddingFolder(ByRef oForm As SAPbouiCOM.Form, ByVal sItemUID As String, _
                              ByVal sItemCaption As String, ByVal sBeforeItemUID As String, _
                              ByVal sUDS As String, ByVal bVisible As Boolean, ByRef sErrDesc As String, _
                              Optional ByVal iFromPane As Int16 = 0, Optional ByVal iToPane As Int16 = 0) As Long

        ' **********************************************************************************
        '   Function    :   AddingFolder
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for add folder like tab .
        '               
        '   Parameters  :  ByRef oForm As SAPbouiCOM.Form
        '                       oForm =  set the SAP form  to be returned to calling function                                               
        '               :   ByVal sItemUID As String
        '                       sItemUID = to use folder id in sap appliction.
        '                    ByVal sItemCaption As String
        '                    sItemCaption=folder caption to be returned to calling function        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String
        Try
            sFuncName = "AddingFolder()"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oForm.DataSources.UserDataSources.Add(sUDS, SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)

            oItem = oForm.Items.Add(sItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.AffectsFormMode = False
            oItem.Specific.Caption = sItemCaption
            oItem.Specific.DataBind.SetBound(True, "", sUDS)
            oItem.Specific.GroupWith(sBeforeItemUID)        'Group with existing folders
            oItem.Left = oForm.Items.Item(sBeforeItemUID).Left + 1
            oItem.Visible = bVisible
            oItem.FromPane = iFromPane
            oItem.ToPane = iToPane
            AddingFolder = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddingFolder = RTN_ERROR
            sErrDesc = exc.Message
            sFuncName = "AddingFolder()"
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oItem = Nothing
        End Try
    End Function

    Public Function AddMatrixRow(ByRef oMatrix As SAPbouiCOM.Matrix, _
                                 ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   AddMatrixRow
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for add Maxtrix row to be returned function.
        '   Parameters  :   ByRef oMatrix As SAPbouiCOM.Matrix
        '                       oMatrix =  set the SAP Maxtrix to be returned to calling function                                               
        '               :   ByRef sErrDesc As String
        '                       sErrDesc = To show error message in sap form data.        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oItems As SAPbouiCOM.Items
        Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String = "AddMatrixRow()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oMatrix.RowCount = 0 Then
                oMatrix.AddRow()
            Else
                If Not oMatrix.Columns.Item("UOMFr").Cells.Item(oMatrix.RowCount).Specific.Selected Is Nothing Then
                    If oMatrix.Columns.Item("UOMFr").Cells.Item(oMatrix.RowCount).Specific.Selected.Value <> "" Then
                        oMatrix.AddRow()
                    End If
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            AddMatrixRow = RTN_SUCCESS
        Catch exc As Exception
            AddMatrixRow = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oItems = Nothing
            oItem = Nothing
        End Try
    End Function

    Public Function AddDBDataSrc(ByRef oForm As SAPbouiCOM.Form, ByVal sTableName As String, _
                                 ByVal sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   AddDBDataSrc
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for Database Datasource .
        '               
        '   Parameters  :    ByRef oForm As SAPbouiCOM.Form
        '                       oForm =  set the SAP UI form to be returned to calling function
        '                    ByVal sTableName As String
        '                       sTableName = set SAP table Name for DBDatasource to be returned to calling function                          
        '               :    ByRef sErrDesc As String
        '                       sErrDesc = To show error message in sap form data.        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim lCount As Long = 0
        Dim sFuncName As String = "AddDBDataSrc()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oForm.DataSources.DBDataSources.Add(sTableName)

            AddDBDataSrc = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddDBDataSrc = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function AddUserDataSrc(ByRef oForm As SAPbouiCOM.Form, ByVal sDSUID As String, _
                                   ByRef sErrDesc As String, ByVal oDataType As SAPbouiCOM.BoDataType, _
                                   Optional ByVal lLen As Long = 0) As Long

        ' **********************************************************************************
        '   Function    :   AddUserDataSrc
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for  use User Data Source .
        '               
        '   Parameters  :    ByVal sDSUID As String
        '                       sDSUID =  set the SAP User Data Source Id be returned to calling function
        '                     ByVal oDataType As SAPbouiCOM.BoDataType
        '                       oDataType = set SAP User Data Source for Data Type to be returned to calling function
        '                     Optional ByVal lLen As Long  
        '                     lLen=To use option
        '                , ByRef sErrDesc As String
        '                       sErrDesc = To show error message in sap form data.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************


        Dim sFuncName As String = "AddUserDataSrc()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If lLen = 0 Then
                oForm.DataSources.UserDataSources.Add(sDSUID, oDataType)
            Else
                oForm.DataSources.UserDataSources.Add(sDSUID, oDataType, lLen)
            End If

            AddUserDataSrc = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddUserDataSrc = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function AddItem(ByRef oForm As SAPbouiCOM.Form, ByVal vItemUID As String, _
                             ByVal oItemType As SAPbouiCOM.BoFormItemTypes, ByRef sErrDesc As String, _
                             Optional ByVal sCaption As String = "", Optional ByVal iPos As Integer = 0, _
                             Optional ByVal sPosItemUID As String = "", Optional ByVal lSpace As Long = 5, _
                             Optional ByVal lLeft As Long = 0, Optional ByVal lTop As Long = 0, _
                             Optional ByVal lHeight As Long = 0, Optional ByVal lWidth As Long = 0, _
                             Optional ByVal lFromPane As Long = 0, Optional ByVal lToPane As Long = 0, _
                             Optional ByVal sBindTbl As String = "", Optional ByVal sAlias As String = "", _
                             Optional ByVal bDisplayDesc As Boolean = False) As Long

        ' **********************************************************************************
        '   Function    :   AddItem
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for  use add item .
        '               
        '   Parameters  :   BByRef oForm As SAPbouiCOM.Form, ByVal vItemUID As String,
        '                   ByVal oItemType As SAPbouiCOM.BoFormItemTypes, ByRef sErrDesc As String, 
        '                   Optional ByVal sPosItemUID As String = "", Optional ByVal lSpace As Long = 5, 
        '                   Optional ByVal lLeft As Long = 0, Optional ByVal lTop As Long = 0, 
        '                   Optional ByVal lHeight As Long = 0, Optional ByVal lWidth As Long = 0, 
        '                   Optional ByVal lFromPane As Long = 0, Optional ByVal lToPane As Long = 0,
        '                   Optional ByVal sBindTbl As String = "", Optional ByVal sAlias As String = "", 
        '                   Optional ByVal bDisplayDesc As Boolean = False.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oItem As SAPbouiCOM.Item
        Dim oPosItem As SAPbouiCOM.Item
        Dim sFuncName As String = "AddItem()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oItem = oForm.Items.Add(vItemUID, oItemType)

            If Trim(sPosItemUID) <> "" Then
                oPosItem = oForm.Items.Item(sPosItemUID)
                oItem.Height = oPosItem.Height
                oItem.Width = oPosItem.Width

                Select Case iPos
                    Case 1      'Left of sPosItemUID
                        oItem.Left = oPosItem.Left - lSpace
                        oItem.Top = oPosItem.Top
                    Case 2      '2=Right of sPosItemUID
                        oItem.Left = oPosItem.Left + oPosItem.Width + lSpace
                        oItem.Top = oPosItem.Top
                    Case 3      '3=Top of sPosItemUID
                        oItem.Left = oPosItem.Left
                        oItem.Top = oPosItem.Top - lSpace
                    Case 4
                        oItem.Left = oPosItem.Left + lSpace
                        oItem.Top = oPosItem.Top + lSpace
                    Case Else   'Below sPosItemUID
                        oItem.Left = oPosItem.Left
                        oItem.Top = oPosItem.Top + oPosItem.Height + lSpace
                End Select
            End If

            If lTop <> 0 Then oItem.Top = lTop
            If lLeft <> 0 Then oItem.Left = lLeft
            If lHeight <> 0 Then oItem.Height = lHeight
            If lWidth <> 0 Then oItem.Width = lWidth

            If Trim(sBindTbl) <> "" Or Trim(sAlias) <> "" Then oItem.Specific.DataBind.SetBound(True, sBindTbl, sAlias)

            oItem.FromPane = lFromPane
            oItem.ToPane = lToPane
            oItem.DisplayDesc = bDisplayDesc

            If Trim(sCaption) <> "" Then oItem.Specific.Caption = sCaption
            AddItem = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddItem = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oItem = Nothing
            oPosItem = Nothing
        End Try
    End Function

    Public Function AddMatrixCol(ByRef oForm As SAPbouiCOM.Form, ByVal sMtxUID As String, _
                                 ByVal sColUID As String, ByVal oColType As SAPbouiCOM.BoFormItemTypes, _
                                 ByVal sCaption As String, ByRef sErrDesc As String, _
                                 Optional ByVal lWidth As Long = 20, Optional ByVal bEditable As Boolean = True, _
                                 Optional ByVal bVisible As Boolean = True, Optional ByVal sTblName As String = "", _
                                 Optional ByVal sAlias As String = "", Optional ByVal oLinkedObj As SAPbouiCOM.BoLinkedObject = 0, _
                                 Optional ByVal oDataType As SAPbouiCOM.BoDataType = 0, Optional ByVal lDataLen As Long = 0, _
                                 Optional ByVal sValOn As String = "", Optional ByVal sValOff As String = "") As Long

        ' **********************************************************************************
        '   Function    :   AddMatrixCol
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for  use add matrix columns  item .
        '               
        '   Parameters  :   BByRef oForm As SAPbouiCOM.Form, ByVal vItemUID As String,
        '                   ByVal sColUID As String, ByVal oColType As SAPbouiCOM.BoFormItemTypes, _
        '                   ByVal sCaption As String, ByRef sErrDesc As String, 
        '                   Optional ByVal lWidth As Long = 20, Optional ByVal bEditable As Boolean = True, 
        '                   Optional ByVal bVisible As Boolean = True, Optional ByVal sTblName As String = "", 
        '                   Optional ByVal sAlias As String = "", Optional ByVal oLinkedObj As SAPbouiCOM.BoLinkedObject = 0, 
        '                   Optional ByVal oDataType As SAPbouiCOM.BoDataType = 0, Optional ByVal lDataLen As Long = 0, 
        '                   Optional ByVal sValOn As String = "", Optional ByVal sValOff As String
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oCol As SAPbouiCOM.Column
        Dim sFuncName As String = "AddMatrixCol()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            AddMatrixCol = False

            oMatrix = oForm.Items.Item(sMtxUID).Specific

            oCol = oMatrix.Columns.Add(sColUID, oColType)
            oCol.TitleObject.Caption = sCaption
            oCol.Width = lWidth
            oCol.Editable = bEditable
            oCol.Visible = bVisible

            If Trim(sTblName) <> "" Or Trim(sAlias) <> "" Then
                If Trim(sTblName) <> "" Then
                    If AddDBDataSrc(oForm, sTblName, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Else
                    If AddUserDataSrc(oForm, sAlias, sErrDesc, oDataType, lDataLen) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If
            End If

            If Trim(sTblName) <> "" Or Trim(sAlias) <> "" Then oCol.DataBind.SetBound(True, sTblName, sAlias)
            If oColType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
                If sValOn <> "" And sValOff <> "" Then
                    oCol.ValOn = sValOn
                    oCol.ValOff = sValOff
                End If
            End If
            If oLinkedObj <> 0 Then oCol.ExtendedObject.LinkedObject = oLinkedObj

            AddMatrixCol = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AddMatrixCol = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oCol = Nothing
            oMatrix = Nothing
        End Try

    End Function

    Public Function ActivateFolder(ByRef oForm As SAPbouiCOM.Form, ByVal vPaneLevel As Long, _
                                   ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   ActivateFolder
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeal Form for  use to active tab item .
        '               
        '   Parameters  :  ByRef oForm As SAPbouiCOM.Form, ByVal vPaneLevel As Long
        '                  ByRef sErrDesc As String       
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "ActivateFolder()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oForm.PaneLevel = vPaneLevel

            ActivateFolder = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            ActivateFolder = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function LoadMatrixCombo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal sColUID As String, _
                                    ByVal oDBDS As SAPbouiCOM.DBDataSource, ByVal bDisplayDesc As Boolean, _
                                    ByVal sValue As String, ByVal sDesc As String, _
                                    ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   LoadMatrixCombo()
        '   Purpose     :   This function will be providing to load matrix combo items
        '                   ImportSeal Form.
        '               
        '   Parameters  :  ByRef oMatrix As SAPbouiCOM.Matrix, ByVal sColUID As String,
        '                  ByVal oDBDS As SAPbouiCOM.DBDataSource, ByVal bDisplayDesc As Boolean,  
        '                  ByVal sValue As String, ByVal sDesc As String
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim oCol As SAPbouiCOM.Column
        Dim lCount As Long
        Dim sFuncName As String = "LoadMatrixCombo()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oCol = oMatrix.Columns.Item(sColUID)
            oCol.DisplayDesc = bDisplayDesc
            oDBDS.Query()

            If oCol.ValidValues.Count = 0 Then
                oCol.ValidValues.Add("", "")
                For lCount = 0 To oDBDS.Size - 1
                    oCol.ValidValues.Add(Trim(oDBDS.GetValue(sValue, lCount)), Trim(oDBDS.GetValue(sDesc, lCount)))
                Next lCount
            End If

            'oCol.ValidValues.Add("", "")
            LoadMatrixCombo = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            LoadMatrixCombo = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oCol = Nothing
        End Try

    End Function

    Public Function AttachCurrency(ByVal dAmount As Double, ByRef sCurAmt As String, ByVal iDec As Int16, ByVal sCur As String, ByVal sCurPos As Char, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   AttachCurrency()
        '   Purpose     :   This function will be providing to attach currency for
        '                   ImportSeal Form.
        '               
        '   Parameters  :  ByVal dAmount As Double, ByRef sCurAmt As String, ByVal iDec As Int16,
        '                  ByVal sCur As String, ByVal sCurPos As Char, ByRef sErrDesc As String  
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "AttachCurrency()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sCurAmt = ""

            If sCur = "" Then
                sCurAmt = FormatNumber(dAmount, iDec)
            Else
                sCur = UCase(sCur)
                If Trim(sCurPos) = "L" Then
                    sCurAmt = Trim(sCur) & " " & FormatNumber(dAmount, iDec)
                Else
                    sCurAmt = FormatNumber(dAmount, iDec) & " " & Trim(sCur)
                End If
            End If

            AttachCurrency = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            AttachCurrency = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Function

    Public Function CheckCurrFld(ByVal sPrice As String, ByRef sCur As String, ByRef sErrDesc As String) As String
        ' **********************************************************************************
        '   Function    :   CheckCurrFld()
        '   Purpose     :   This function will be providing to attach currency for
        '                   ImportSeal Form.
        '               
        '   Parameters  :  ByVal sPrice As String, ByRef sCur As String, ByRef sErrDesc As String
        '                   
        '               
        '   Return      :   string - FAILURE
        '                   string - SUCCESS
        ' **********************************************************************************

        Dim iCnt As Integer
        Dim iLimit As Integer
        Dim sFuncName As String = "CheckCurrFld()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sCur = ""
            iLimit = Len(sPrice)

            For iCnt = 1 To iLimit
                Select Case CStr(Mid(sPrice, iCnt, 1))
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", " "
                    Case "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T" _
                         , "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n" _
                         , "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
                        sCur = sCur + UCase(CStr(Mid(sPrice, iCnt, 1)))
                End Select
            Next

            CheckCurrFld = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            CheckCurrFld = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function CheckValFld(ByVal sPrice As String, ByRef sValue As String, ByRef sErrDesc As String) As String
        ' **********************************************************************************
        '   Function    :   CheckValFld()
        '   Purpose     :   This function will be providing to attach currency for
        '                   ImportSeal Form.
        '               
        '   Parameters  :  ByVal sPrice As String, ByRef sValue As String, ByRef sErrDesc As String
        '                 
        '               
        '   Return      :   string - FAILURE
        '                   string - SUCCESS
        ' **********************************************************************************

        Dim iCnt As Integer
        Dim iLimit As Integer
        Dim sFuncName As String = "CheckValFld()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            sValue = ""
            iLimit = Len(sPrice)

            For iCnt = 1 To iLimit
                Select Case CStr(Mid(sPrice, iCnt, 1))
                    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
                        sValue = sValue + CStr(Mid(sPrice, iCnt, 1))
                    Case Else

                End Select
            Next

            CheckValFld = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            CheckValFld = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function GetSystemDefaultData(ByRef gCompDef As CompanyDefault, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   GetSystemDefaultData()
        '   Purpose     :   This function will be providing to get system default data for
        '                   ImportSeal Form.
        '               
        '   Parameters  :  ByRef gCompDef As CompanyDefault, ByRef sErrDesc As String
        '                 
        '               
        '   Return      :   string - FAILURE
        '                   string - SUCCESS
        ' **********************************************************************************
        Dim oRS As SAPbobsCOM.Recordset
        Dim sSql As String
        Dim sFuncName As String = "GetSystemDefaultData()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            gCompDef.CurrencyPosition = "R"
            gCompDef.LocalCurrency = ""
            gCompDef.SystemCurrency = ""
            gCompDef.CompanyName = ""
            gCompDef.DBName = ""
            gCompDef.HolidaysName = ""

            oRS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSql = "SELECT top 1 * FROM OADM " '& _
            '"WHERE FinancYear <= GetDate() " ' & _
            '"ORDER BY FinancYear DESC "
            oRS.DoQuery(sSql)

            If oRS.RecordCount > 0 Then
                gCompDef.DBName = p_oDICompany.CompanyDB
                If Trim(oRS.Fields.Item("CurOnRight").Value) = "N" Then
                    gCompDef.CurrencyPosition = "L"
                Else
                    gCompDef.CurrencyPosition = "R"
                End If
                gCompDef.LocalCurrency = oRS.Fields.Item("MainCurncy").Value & vbNullString
                gCompDef.SystemCurrency = oRS.Fields.Item("SysCurrncy").Value & vbNullString
                gCompDef.CompanyName = oRS.Fields.Item("CompnyName").Value & vbNullString
                gCompDef.iPriceDecimal = oRS.Fields.Item("PriceDec").Value & vbNullString
                gCompDef.iQtyDecimal = oRS.Fields.Item("QtyDec").Value & vbNullString
                gCompDef.HolidaysName = oRS.Fields.Item("HldCode").Value & vbNullString
            End If

            GetSystemDefaultData = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            GetSystemDefaultData = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        Finally
            oRS = Nothing
        End Try

    End Function

    Public Function DisplayStatus(ByVal oFrmParent As SAPbouiCOM.Form, ByVal sMsg As String, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   DisplayStatus()
        '   Purpose     :   This function will be providing to show status of display  for
        '                   ImportSeal Form.
        '               
        '   Parameters  :  ByVal oFrmParent As SAPbouiCOM.Form, ByVal sMsg As String, ByRef sErrDesc As String
        '                 
        '               
        '   Return      :   string - FAILURE
        '                   string - SUCCESS
        ' **********************************************************************************

        Dim oForm As SAPbouiCOM.Form
        Dim oItem As SAPbouiCOM.Item
        Dim oTxt As SAPbouiCOM.StaticText
        Dim creationPackage As SAPbouiCOM.FormCreationParams
        Dim iCount As Integer
        Dim sFuncName As String = "DisplayStatus"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            'Check whether the form exists.If exists then close the form
            For iCount = 0 To p_oSBOApplication.Forms.Count - 1
                oForm = p_oSBOApplication.Forms.Item(iCount)
                If oForm.UniqueID = "dStatus1" Then
                    oForm.Close()
                    oForm = Nothing
                    Exit For
                End If
            Next iCount
            'Add Form
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating form Display Status", sFuncName)
            creationPackage = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            creationPackage.UniqueID = "dStatus1"
            creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle
            creationPackage.FormType = "OB1_dStatus"
            oForm = p_oSBOApplication.Forms.AddEx(creationPackage)
            With oForm
                .AutoManaged = False
                .Width = 300
                .Height = 100
                If oFrmParent Is Nothing Then
                    .Left = ((Screen.PrimaryScreen.WorkingArea.Width) - oForm.Width) / 2
                    .Top = ((Screen.PrimaryScreen.WorkingArea.Height) - oForm.Height) / 3
                Else
                    .Left = ((oFrmParent.Left * 2) + oFrmParent.Width - oForm.Width) / 2
                    .Top = ((oFrmParent.Top * 2) + oFrmParent.Height - oForm.Height) / 2
                End If

                .Visible = True
            End With

            'Add Label
            oItem = oForm.Items.Add("3", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Top = 40
            oItem.Left = 40
            oItem.Width = 250
            oTxt = oItem.Specific
            oTxt.Caption = sMsg

            DisplayStatus = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with SUCCESS", sFuncName)

        Catch ex As Exception
            DisplayStatus = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        Finally
            creationPackage = Nothing
            oForm = Nothing
            oItem = Nothing
            oTxt = Nothing
        End Try
    End Function

    Public Function EndStatus(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function    :   EndStatus()
        '   Purpose     :   This function will be providing to show status of end  for
        '                   ImportSeal Form.
        '               
        '   Parameters  :   ByRef sErrDesc As String
        '                 
        '               
        '   Return      :   string - FAILURE
        '                   string - SUCCESS
        ' **********************************************************************************


        Dim oForm As SAPbouiCOM.Form
        Dim iCount As Integer
        Dim sFuncName As String = "EndStatus()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)
            'Check whether the form is exist. If exist then close the form
            For iCount = 0 To p_oSBOApplication.Forms.Count - 1
                oForm = p_oSBOApplication.Forms.Item(iCount)
                If oForm.UniqueID = "dStatus1" Then
                    oForm.Close()
                    Exit For
                End If
            Next iCount

            EndStatus = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with SUCCESS", sFuncName)

        Catch ex As Exception
            EndStatus = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
        Finally
            oForm = Nothing
        End Try
    End Function

    Public Function StartTransaction(ByRef oDICompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   StartTransaction
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeal Form for company connection   in Editext
        '               
        '   Parameters  :   ByRef oDICompany As SAPbobsCOM.Company
        '                       oDICompany =  set the SAP DI Company  Object to be returned to calling function
        '                    ByVal pUniqueID  As string 
        '                       pUniqueID = set SAP UI ItemID Object to be returned to calling function
        '                , ByRef sErrDesc As String
        '                       sErrDesc = To show error message in sap form data.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "StartTransaction()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oDICompany.InTransaction Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Found hanging transaction.Rolling it back.", sFuncName)
                oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If

            oDICompany.StartTransaction()

            StartTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            StartTransaction = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function RollBackTransaction(ByRef oDICompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   RollBackTransaction
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeal Form for company connection   in Editext
        '               
        '   Parameters  :   ByRef oDICompany As SAPbobsCOM.Company
        '                       oDICompany =  set the SAP DI Company  Object to be returned to calling function
        '                    ByVal pUniqueID  As string 
        '                       pUniqueID = set SAP UI ItemID Object to be returned to calling function
        '                , ByRef sErrDesc As String
        '                       sErrDesc = To show error message in sap form data.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "RollBackTransaction()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oDICompany.InTransaction Then
                oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for rollback", sFuncName)
            End If

            RollBackTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            RollBackTransaction = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try

    End Function

    Public Function CommitTransaction(ByRef oDICompany As SAPbobsCOM.Company, ByRef sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   CommitTransaction
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for company connection   in Editext
        '               
        '   Parameters  :   ByRef oDICompany As SAPbobsCOM.Company
        '                       oDICompany =  set the SAP DI Company  Object to be returned to calling function
        '                    ByVal pUniqueID  As string 
        '                       pUniqueID = set SAP UI ItemID Object to be returned to calling function
        '                , ByRef sErrDesc As String
        '                       sErrDesc = To show error message in sap form data.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "CommitTransaction()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oDICompany.InTransaction Then
                oDICompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No active transaction found for commit", sFuncName)
            End If

            CommitTransaction = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            CommitTransaction = RTN_ERROR
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Function

    Public Function LoadFromXML(ByVal objTemp As Object, ByVal pstrFileName As String) As Boolean

        '=============================================================================
        'Function   : LoadFromXML()
        'Purpose    : This function will be providing to call and load SAP Forms to process 
        '             with main  ExporeSealLcLForm 
        'Parameters : ByVal objTemp As Object,
        '             ByVal pstrFileName As String
        'Return     : False - Fail
        '           : True  - Suceess 
        '=============================================================================

        Dim objXmlDoc As Xml.XmlDocument
        LoadFromXML = False
        Try
            objXmlDoc = New Xml.XmlDocument
            'objXmlDoc.Load(IO.Directory.GetParent((Application.StartupPath)).ToString & "\" & pstrFileName)
            objXmlDoc.Load(Application.StartupPath & "\" & pstrFileName)
            objTemp.LoadBatchActions(objXmlDoc.InnerXml)
            If Not AddToAppList(objXmlDoc.SelectSingleNode("/Application/forms/action/form/@uid").Value) Then Throw New ArgumentException()
            LoadFromXML = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            LoadFromXML = False
        End Try
    End Function

    Public Function RemoveFromAppList(ByVal FormUID As String) As Boolean

        ' **********************************************************************************
        '   Function    :   RemoveFromAppList()
        '   Purpose     :   This function will be providing to remove   for
        '                    Form  object name
        '               
        '   Parameters  :  ByVal FormUID As String
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************
        RemoveFromAppList = False
        Try

            For Each objName As String In AppObjString
                If objName = FormUID Then
                    AppObjString.Remove(objName)
                    Exit For
                End If
            Next
            RemoveFromAppList = True
        Catch ex As Exception
            RemoveFromAppList = False
        End Try
    End Function

    Public Function AddToAppList(ByVal FormUID As String) As Boolean

        ' **********************************************************************************
        '   Function    :   AddToAppList()
        '   Purpose     :   This function will be providing to add   for
        '                    Form  object name
        '               
        '   Parameters  :  ByVal FormUID As String
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************
        AddToAppList = False
        Try
            AppObjString.Add(FormUID)
            AddToAppList = True
        Catch ex As Exception
            AddToAppList = False
        End Try
    End Function

    Public Function AlreadyExist(ByVal FormUID As String) As Boolean

        ' **********************************************************************************
        '   Function    :   AlreadyExist()
        '   Purpose     :   This function will be providing to proceed  for
        '                   to check  ImportSeall Form and ExporeSeaLcl Form
        '               
        '   Parameters  :  ByVal FormUID As String
        '               
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        ' **********************************************************************************
        AlreadyExist = False
        Try
            For Each objName As String In AppObjString
                If objName = FormUID Then
                    AlreadyExist = True
                End If
            Next
        Catch ex As Exception
            AlreadyExist = False
        End Try
    End Function

    Public Function AddChooseFromList(ByRef pForm As SAPbouiCOM.Form, ByVal pUniqueID As String, ByVal pMultiSelection As Boolean, _
                                  ByVal pObjectType As Integer, ByVal pAlias As String, ByRef pOperation As SAPbouiCOM.BoConditionOperation, _
                                  ByVal pCondValue As String) As Int16

        ' **********************************************************************************
        '   Function    :   AddChooseFromList
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for AddChooseFromList  in Editext
        '               
        '   Parameters  :   ByRef pForm As As SAPbouiCOM.Form
        '                       pForm =  set the SAP UI Form  Object to be returned to calling function
        '                    ByVal pUniqueID  As string 
        '                       pUniqueID = set SAP UI ItemID Object to be returned to calling function
        '                  ByVal pMultiSelection As Boolean
        '                       pMultiSelection = To Select ChoosseFromList data.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        AddChooseFromList = RTN_ERROR
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", "AddChooseFromList()")
            oCFLs = pForm.ChooseFromLists
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.UniqueID = pUniqueID
            oCFLCreationParams.MultiSelection = pMultiSelection
            oCFLCreationParams.ObjectType = pObjectType
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = pAlias
            oCon.Operation = pOperation
            oCon.CondVal = pCondValue
            oCFL.SetConditions(oCons)
            AddChooseFromList = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", "AddChooseFromList()")
        Catch ex As Exception
            ShowErr(ex.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function With Error" & ex.Message, "AddChooseFromList()")
        End Try
    End Function

    Public Function AddChooseFromList(ByRef pForm As SAPbouiCOM.Form, ByVal pUniqueID As String, _
                                        ByVal pMultiSelection As Boolean, ByVal pObjectType As String) As Int16
        ' **********************************************************************************
        '   Function    :   AddChooseFromList
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for AddChooseFromList  in Editext
        '               
        '   Parameters  :   ByRef pForm AsAs SAPbouiCOM.Form
        '                       pForm =  set the SAP UI Form  Object to be returned to calling function
        '                    ByVal pUniqueID  As string 
        '                       pUniqueID = set SAP UI ItemID Object to be returned to calling function
        '                    ByVal pMultiSelection As Boolean
        '                       pMultiSelection = To Select ChoosseFromList data.        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************


        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        AddChooseFromList = RTN_ERROR
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", "AddChooseFromList()")
            oCFLs = pForm.ChooseFromLists
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.UniqueID = pUniqueID
            oCFLCreationParams.MultiSelection = pMultiSelection
            oCFLCreationParams.ObjectType = pObjectType
            oCFL = oCFLs.Add(oCFLCreationParams)
            AddChooseFromList = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", "AddChooseFromList()")
        Catch ex As Exception
            ShowErr(ex.Message)
            AddChooseFromList = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function With Error" & ex.Message, "AddChooseFromList()")
        End Try
    End Function

    'Public Sub ClearEditBoxValues(ByVal pFormUID As SAPbouiCOM.Form, ByVal p_oItemIDs() As String)
    '    Dim oEditBox As SAPbouiCOM.EditText
    '    Dim oItemID As String
    '    Try
    '        For Each oItemID In p_oItemIDs
    '            oEditBox = pFormUID.Items.Item(oItemID).Specific
    '            'oEditBox.Value.
    '        Next
    '    Catch ex As Exception
    '    End Try
    'End Sub

    Public Function FormatString(ByVal pStrValue As String) As String
        FormatString = "'" + pStrValue + "'"
    End Function

    Public Function HolidayMarkUp(ByRef pForm As SAPbouiCOM.Form, ByRef DateBox As SAPbouiCOM.EditText, _
                                   ByVal HolidaySetName As String, ByVal sErrDesc As String, _
                                   ByRef DayBox As SAPbouiCOM.EditText, ByRef HrBox As SAPbouiCOM.EditText) As Long

        ' **********************************************************************************
        '   Function    :   HolidayMarkUp()
        '   Purpose     :   This function will be providing to calculate  for
        '                   ImportSeall Form Date Time Picker of holiday  and Date  information
        '               
        '   Parameters  :  ByRef pForm As SAPbouiCOM.Form, ByRef DateBox As SAPbouiCOM.EditText,
        '                  ByVal HolidaySetName As String, ByVal sErrDesc As String
        '                  sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "HolidayMarkUp()"
        Dim ObjRecordSet As SAPbobsCOM.Recordset
        Dim SqlQuery As String = String.Empty
        Dim StartDate, EndDate As Date
        Dim TotalDay As Long = 0
        Dim OperandDate As String = String.Empty
        Dim OperandDay As Integer = 0

        Try
            If Not String.IsNullOrEmpty(DateBox.String) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                ObjRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                SqlQuery = "SELECT StrDate,EndDate FROM HLD1 INNER JOIN OHLD ON HLD1.HldCode = OHLD.HldCode " & _
                           "WHERE OHLD.HldCode = '" & HolidaySetName & "' AND " & _
                           "MONTH(StrDate) = '" & CInt(DateBox.Value.ToString.Substring(4, 2)) & "'"
                ObjRecordSet.DoQuery(SqlQuery)
                OperandDate = DateBox.Value.ToString
                If ObjRecordSet.RecordCount > 0 Then
                    ObjRecordSet.MoveFirst()
                    Do Until ObjRecordSet.EoF
                        StartDate = CDate(ObjRecordSet.Fields.Item("StrDate").Value)
                        EndDate = CDate(ObjRecordSet.Fields.Item("EndDate").Value)
                        TotalDay = DateDiff(DateInterval.Day, StartDate, EndDate)
                        If StartDate.ToString("yyyyMMdd") = EndDate.ToString("yyyyMMdd") Then
                            If StartDate.ToString("dd") = DateBox.Value.ToString.Substring(6, 2) Then
                                DateBox.TextStyle = 1
                                DateBox.ForeColor = 255
                                DayBox.TextStyle = 1
                                DayBox.ForeColor = 255
                                HrBox.TextStyle = 1
                                HrBox.ForeColor = 255
                            Else
                                DateBox.TextStyle = 0
                                DateBox.ForeColor = 0
                                DayBox.TextStyle = 0
                                DayBox.ForeColor = 0
                                HrBox.TextStyle = 0
                                HrBox.ForeColor = 0
                            End If
                        ElseIf TotalDay <> 0 Then
                            For Count As Integer = 0 To TotalDay
                                OperandDay = CInt(StartDate.ToString("dd")) + Count
                                If CStr(OperandDay) = DateBox.Value.ToString.Substring(6, 2) Then
                                    DateBox.TextStyle = 1
                                    DateBox.ForeColor = 255
                                    DayBox.TextStyle = 1
                                    DayBox.ForeColor = 255
                                    HrBox.TextStyle = 1
                                    HrBox.ForeColor = 255
                                    Exit For
                                Else
                                    If Count = TotalDay Then
                                        DateBox.TextStyle = 0
                                        DateBox.ForeColor = 0
                                        DayBox.TextStyle = 0
                                        DayBox.ForeColor = 0
                                        HrBox.TextStyle = 0
                                        HrBox.ForeColor = 0
                                    End If
                                End If
                            Next Count
                        End If
                        ObjRecordSet.MoveNext()
                    Loop
                Else
                    DateBox.TextStyle = 0
                    DateBox.ForeColor = 0
                    DayBox.TextStyle = 0
                    DayBox.ForeColor = 0
                    HrBox.TextStyle = 0
                    HrBox.ForeColor = 0
                End If
            End If
            HolidayMarkUp = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", sFuncName)
        Catch ex As Exception
            HolidayMarkUp = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Function

    Public Function HolidayMarkUpWithoutDay(ByRef pForm As SAPbouiCOM.Form, ByRef DateBox As SAPbouiCOM.EditText, _
                                 ByVal HolidaySetName As String, ByVal sErrDesc As String, _
                                  ByRef HrBox As SAPbouiCOM.EditText) As Long

        ' **********************************************************************************
        '   Function    :   HolidayMarkUp()
        '   Purpose     :   This function will be providing to calculate  for
        '                   ImportSeall Form Date Time Picker of holiday  and Date  information
        '               
        '   Parameters  :  ByRef pForm As SAPbouiCOM.Form, ByRef DateBox As SAPbouiCOM.EditText,
        '                  ByVal HolidaySetName As String, ByVal sErrDesc As String
        '                  sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "HolidayMarkUp()"
        Dim ObjRecordSet As SAPbobsCOM.Recordset
        Dim SqlQuery As String = String.Empty
        Dim StartDate, EndDate As Date
        Dim TotalDay As Long = 0
        Dim OperandDate As String = String.Empty
        Dim OperandDay As Integer = 0

        Try
            If Not String.IsNullOrEmpty(DateBox.String) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                ObjRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                SqlQuery = "SELECT StrDate,EndDate FROM HLD1 INNER JOIN OHLD ON HLD1.HldCode = OHLD.HldCode " & _
                           "WHERE OHLD.HldCode = '" & HolidaySetName & "' AND " & _
                           "MONTH(StrDate) = '" & CInt(DateBox.Value.ToString.Substring(4, 2)) & "'"
                ObjRecordSet.DoQuery(SqlQuery)
                OperandDate = DateBox.Value.ToString
                If ObjRecordSet.RecordCount > 0 Then
                    ObjRecordSet.MoveFirst()
                    Do Until ObjRecordSet.EoF
                        StartDate = CDate(ObjRecordSet.Fields.Item("StrDate").Value)
                        EndDate = CDate(ObjRecordSet.Fields.Item("EndDate").Value)
                        TotalDay = DateDiff(DateInterval.Day, StartDate, EndDate)
                        If StartDate.ToString("yyyyMMdd") = EndDate.ToString("yyyyMMdd") Then
                            If StartDate.ToString("dd") = DateBox.Value.ToString.Substring(6, 2) Then
                                DateBox.TextStyle = 1
                                DateBox.ForeColor = 255

                                HrBox.TextStyle = 1
                                HrBox.ForeColor = 255
                            Else
                                DateBox.TextStyle = 0
                                DateBox.ForeColor = 0
                          
                                HrBox.TextStyle = 0
                                HrBox.ForeColor = 0
                            End If
                        ElseIf TotalDay <> 0 Then
                            For Count As Integer = 0 To TotalDay
                                OperandDay = CInt(StartDate.ToString("dd")) + Count
                                If CStr(OperandDay) = DateBox.Value.ToString.Substring(6, 2) Then
                                    DateBox.TextStyle = 1
                                    DateBox.ForeColor = 255
                                
                                    HrBox.TextStyle = 1
                                    HrBox.ForeColor = 255
                                    Exit For
                                Else
                                    If Count = TotalDay Then
                                        DateBox.TextStyle = 0
                                        DateBox.ForeColor = 0
                              
                                        HrBox.TextStyle = 0
                                        HrBox.ForeColor = 0
                                    End If
                                End If
                            Next Count
                        End If
                        ObjRecordSet.MoveNext()
                    Loop
                Else
                    DateBox.TextStyle = 0
                    DateBox.ForeColor = 0
                  
                    HrBox.TextStyle = 0
                    HrBox.ForeColor = 0
                End If
            End If
            HolidayMarkUpWithoutDay = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", sFuncName)
        Catch ex As Exception
            HolidayMarkUpWithoutDay = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Function


    Public Function HolidayMarkUp(ByRef pForm As SAPbouiCOM.Form, ByRef DateBox As SAPbouiCOM.EditText, _
                                   ByVal HolidaySetName As String, ByVal sErrDesc As String) As Long

        ' **********************************************************************************
        '   Function    :   HolidayMarkUp()
        '   Purpose     :   This function will be providing to calculate  for
        '                   ImportSeall Form Date Time Picker of holiday  and Date  information
        '               
        '   Parameters  :  ByRef pForm As SAPbouiCOM.Form, ByRef DateBox As SAPbouiCOM.EditText,
        '                  ByVal HolidaySetName As String, ByVal sErrDesc As String
        '                  sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim sFuncName As String = "HolidayMarkUp()"
        Dim ObjRecordSet As SAPbobsCOM.Recordset
        Dim SqlQuery As String = String.Empty
        Dim StartDate, EndDate As Date
        Dim TotalDay As Long = 0
        Dim OperandDate As String = String.Empty
        Dim OperandDay As Integer = 0
        Try
            If Not String.IsNullOrEmpty(DateBox.String) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
                ObjRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                SqlQuery = "SELECT StrDate,EndDate FROM HLD1 INNER JOIN OHLD ON HLD1.HldCode = OHLD.HldCode " & _
                           "WHERE OHLD.HldCode = '" & HolidaySetName & "' AND " & _
                           "MONTH(StrDate) = '" & CInt(DateBox.Value.ToString.Substring(4, 2)) & "'"
                ObjRecordSet.DoQuery(SqlQuery)
                OperandDate = DateBox.Value.ToString
                If ObjRecordSet.RecordCount > 0 Then
                    ObjRecordSet.MoveFirst()
                    Do Until ObjRecordSet.EoF
                        StartDate = CDate(ObjRecordSet.Fields.Item("StrDate").Value)
                        EndDate = CDate(ObjRecordSet.Fields.Item("EndDate").Value)
                        TotalDay = DateDiff(DateInterval.Day, StartDate, EndDate)
                        If StartDate.ToString("yyyyMMdd") = EndDate.ToString("yyyyMMdd") Then
                            If StartDate.ToString("dd") = DateBox.Value.ToString.Substring(6, 2) Then
                                DateBox.TextStyle = 1
                                DateBox.ForeColor = 255
                            Else
                                DateBox.TextStyle = 0
                                DateBox.ForeColor = 0
                            End If
                        ElseIf TotalDay <> 0 Then
                            For Count As Integer = 0 To TotalDay
                                OperandDay = CInt(StartDate.ToString("dd")) + Count
                                If CStr(OperandDay) = DateBox.Value.ToString.Substring(6, 2) Then
                                    DateBox.TextStyle = 1
                                    DateBox.ForeColor = 255
                                    Exit For
                                Else
                                    If Count = TotalDay Then
                                        DateBox.TextStyle = 0
                                        DateBox.ForeColor = 0
                                    End If
                                End If
                            Next Count
                        End If
                        ObjRecordSet.MoveNext()
                    Loop
                Else
                    'MSW
                    DateBox.TextStyle = 0
                    DateBox.ForeColor = 0
                    'MSW
                End If
            End If
            HolidayMarkUp = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", sFuncName)
        Catch ex As Exception
            HolidayMarkUp = RTN_ERROR
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
        End Try
    End Function

    Public Function DateTime(ByVal ImportSeaLCLForm As SAPbouiCOM.Form, ByRef CtrlYear As SAPbouiCOM.EditText, ByRef CtrlDay As SAPbouiCOM.EditText, ByRef CtrlTime As SAPbouiCOM.EditText) As Boolean
        ' **********************************************************************************
        '   Function    :   DateTime
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form Date Time Picker and Date Event information
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.BusinessObjectInfo
        '                       ImportSeaLCLForm =  set the SAP UI Form  Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        DateTime = False
        Dim SelectDate As String = CtrlYear.Value.ToString
        Dim Year, Month, Day, Time, sErrDesc, ShowDay As String
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
            Time = Now.ToString("HH:mm")
            'If Convert.ToInt32(Left(Time, 2)) > 12 Then
            '    Time = Time + "PM"
            'Else
            '    Time = Time + "AM"
            'End If
            CtrlDay.Value = ShowDay.ToString
            CtrlTime.Value = Now.ToString("HH:mm")
            'CtrlTime.Value = Time
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", sFuncName)
            DateTime = True
        Catch ex As Exception
            sErrDesc = ex.Message
            ShowErr(sErrDesc)
            DateTime = RTN_ERROR
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with ERROR", sFuncName)
            DateTime = False
        End Try
    End Function

    Public Function DateTimeWithoutDay(ByVal ImportSeaLCLForm As SAPbouiCOM.Form, ByRef CtrlYear As SAPbouiCOM.EditText, ByRef CtrlTime As SAPbouiCOM.EditText) As Boolean
        ' **********************************************************************************
        '   Function    :   DateTime
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form Date Time Picker and Date Event information
        '               
        '   Parameters  :   ByRef pVal As SAPbouiCOM.BusinessObjectInfo
        '                       ImportSeaLCLForm =  set the SAP UI Form  Object to be returned to calling function
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set SAP UI Bubble Event Object to be returned to calling function
        '                   ByRef sErrDesc AS String 
        '                       sErrDesc = Error Description to be returned to calling function
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        DateTimeWithoutDay = False
        Dim SelectDate As String = CtrlYear.Value.ToString
        Dim Year, Month, Day, Time, sErrDesc, ShowDay As String
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
            'Time = Now.ToString("HH:mm")
            'If Convert.ToInt32(Left(Time, 2)) > 12 Then
            '    Time = Time + "PM"
            'Else
            '    Time = Time + "AM"
            'End If
            'CtrlTime.Value = Now.ToString("HH:mm")
            'CtrlTime.Value = Time
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", sFuncName)
            DateTimeWithoutDay = True
        Catch ex As Exception
            sErrDesc = ex.Message
            ShowErr(sErrDesc)
            DateTimeWithoutDay = RTN_ERROR
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with ERROR", sFuncName)
            DateTimeWithoutDay = False
        End Try
    End Function


    Public Sub BindingChooseFromList(ByRef ctrlParent As Object, ByVal cflID As String, ByVal ctrlName As String, Optional ByVal fieldName As String = vbNullString)
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

    Public Sub AddGSTComboData(ByVal oColumn As SAPbouiCOM.Column)
        Dim RS As SAPbobsCOM.Recordset
        RS = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RS.DoQuery("select Code,Name from ovtg Where Category='I'")
        RS.MoveFirst()
        For j As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
            oColumn.ValidValues.Remove(j, SAPbouiCOM.BoSearchKey.psk_Index)
        Next
        oColumn.ValidValues.Add("None", "None")
        While RS.EoF = False
            oColumn.ValidValues.Add(RS.Fields.Item("Code").Value, RS.Fields.Item("Name").Value)
            RS.MoveNext()
        End While
    End Sub

    Public Function ClearComboData(ByRef pForm As SAPbouiCOM.Form, ByVal pComboUID As String, ByVal DBDataSource As String, ByVal UDF As String) As Boolean


        ' **********************************************************************************
        '   Function    :   ClearComboData()
        '   Purpose     :   This function will be providing to clear Combo items  data for
        '                   ExporeSeaLcl Form
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pComboUID As String,
        '               :   ByVal DataSource As String, ByVal FieldAlias As String
        '   Return      :   Fase - FAILURE
        '                   True - SUCCESS
        '*************************************************************
        ClearComboData = False
        Dim sFuncName As String = "ClearComboData()"
        Dim sErrDesc As String = String.Empty
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            If pForm.Items.Item(pComboUID).Specific.ValidValues.Count > 0 Then
                For i As Integer = pForm.Items.Item(pComboUID).Specific.ValidValues.Count - 1 To 0 Step -1
                    pForm.Items.Item(pComboUID).Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
                pForm.DataSources.DBDataSources.Item(DBDataSource).SetValue(UDF, 0, vbNullString)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", sFuncName)
            ClearComboData = True
        Catch ex As Exception
            Call WriteToLogFile(ex.Message, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete with ERROR", sFuncName)
            ClearComboData = False
        End Try
    End Function

    Public Function AddChooseFromListCondition(ByRef pForm As SAPbouiCOM.Form, ByVal pUniqueID As String, ByVal pMultiSelection As Boolean, _
                              ByVal pObjectType As Integer, ByVal pAlias As String, ByRef pOperation As SAPbouiCOM.BoConditionOperation, _
                              ByVal pCondValue As String, ByVal pAlias2 As String, ByVal pConValue2 As String) As Int16


        ' **********************************************************************************
        '   Function    :   AddChooseFromListCondition
        '   Purpose     :   This function will be providing to add choose fromlist conndition 
        '                   ImportSeal Form for AddChooseFromList  in Editext
        '               
        '   Parameters  :   ByRef pForm As SAPbouiCOM.Form, ByVal pUniqueID As String, ByVal pMultiSelection As Boolean,
        '                   ByVal pObjectType As Integer, ByVal pAlias As String, ByRef pOperation As SAPbouiCOM.BoConditionOperation,
        '                   ByVal pCondValue As String, ByVal pAlias2 As String, ByVal pConValue2 As String        '                
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection

        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        AddChooseFromListCondition = RTN_ERROR
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", "AddChooseFromList()")
            oCFLs = pForm.ChooseFromLists
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.UniqueID = pUniqueID
            oCFLCreationParams.MultiSelection = pMultiSelection
            oCFLCreationParams.ObjectType = pObjectType
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = pAlias
            oCon.Operation = pOperation
            oCon.CondVal = pCondValue
            oCFL.SetConditions(oCons)

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
            oCon = oCons.Add()
            oCon.Alias = pAlias2
            oCon.Operation = pOperation
            oCon.CondVal = pConValue2
            oCFL.SetConditions(oCons)

            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
            oCon = oCons.Add()
            oCon.Alias = pAlias2
            oCon.Operation = pOperation
            oCon.CondVal = "All"
            oCFL.SetConditions(oCons)

            AddChooseFromListCondition = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", "AddChooseFromList()")
        Catch ex As Exception
            ShowErr(ex.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function With Error" & ex.Message, "AddChooseFromList()")
        End Try
    End Function

    Public Function GetJobNumber(ByVal prefix As String) As String


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
        Dim oRecordSet As SAPbobsCOM.Recordset
        GetJobNumber = vbNullString
        Try
            Dim jobSrNo As Integer = 0
            Dim postFix As String = String.Empty
            Dim strJobNo As String = String.Empty
            oRecordSet = p_oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If Now.Month = 1 And Now.Day = 1 Then
                jobSrNo = 1
            Else
                oRecordSet.DoQuery("select U_JobNo from [@OBT_FREIGHTDOCNO] Order by docentry asc")
                If oRecordSet.RecordCount > 0 Then
                    oRecordSet.MoveLast()
                    jobSrNo = Convert.ToInt32(oRecordSet.Fields.Item("U_JobNo").Value.ToString.Substring(2, 5)) + 1
                Else
                    jobSrNo = 1
                End If
            End If

            For i = 1 To 5 - jobSrNo.ToString.Length
                postFix = postFix + "0"
            Next
            GetJobNumber = Now.ToString("yyyy").Substring(2, 2) + postFix + jobSrNo.ToString + "-" + prefix
        Catch ex As Exception
            GetJobNumber = vbNullString
            MessageBox.Show(ex.Message)
        End Try
    End Function
   

    Public Sub AddNewBtToCFLFrm(ByRef CFLForm As SAPbouiCOM.Form, ByVal btName As String)

        ' **********************************************************************************
        '   Procedure   :  AddNewBtToCFLFrm
        '   Purpose     :   This function will be providing to Add New Button 
        '                   in Choose From List Form so that user can add new data at once.
        '               
        '   Parameters  :   ByVal CFLForm As SAPbouiCOM.Form 
        '                       Pass the Current Active Form                            
        '                   ByVal btName As String
        '                       Pass the specific button Name                        
        '                       

        '               
        ' **********************************************************************************
        Try
            Dim oItem As SAPbouiCOM.Item
            CFLForm.Freeze(True)
            oItem = CFLForm.Items.Add(btName, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            CFLForm.Width = 455
            CFLForm.Height = 310
            oItem.Top = CFLForm.Items.Item("1").Top
            oItem.Left = 145
            oItem.Width = 65
            oItem.Height = 19
            oItem.FromPane = 0
            oItem.LinkTo = "1"
            oItem.Specific.Caption = "New"
            oItem.Visible = True
            CFLForm.Freeze(False)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        
    End Sub

    Public Sub SetMatrixSeqNo(ByRef oMatrix As SAPbouiCOM.Matrix, ByVal ColName As String)

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
    
    Public Sub DeleteUDO(ByRef oActiveForm As SAPbouiCOM.Form, ByVal UODName As String, ByVal selectedRow As String)

        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        oGeneralService = p_oDICompany.GetCompanyService.GetGeneralService(UODName)

        'Delete UDO record
        oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
        oGeneralParams.SetProperty("DocEntry", selectedRow)
        oGeneralService.Delete(oGeneralParams)

    End Sub
    Public Function AddChooseFromListByFilter(ByRef pForm As SAPbouiCOM.Form, ByVal pUniqueID As String, ByVal pMultiSelection As Boolean, _
                                  ByVal pObjectType As String, ByVal pAlias As String, ByRef pOperation As SAPbouiCOM.BoConditionOperation, _
                                  ByVal pCondValue As String) As Int16

        ' **********************************************************************************
        '   Function    :   AddChooseFromList
        '   Purpose     :   This function will be providing to proceed  for
        '                   ImportSeall Form for AddChooseFromList  in Editext
        '               
        '   Parameters  :   ByRef pForm As As SAPbouiCOM.Form
        '                       pForm =  set the SAP UI Form  Object to be returned to calling function
        '                    ByVal pUniqueID  As string 
        '                       pUniqueID = set SAP UI ItemID Object to be returned to calling function
        '                  ByVal pMultiSelection As Boolean
        '                       pMultiSelection = To Select ChoosseFromList data.
        '               
        '   Return      :   0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        AddChooseFromListByFilter = RTN_ERROR
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", "AddChooseFromList()")
            oCFLs = pForm.ChooseFromLists
            oCFLCreationParams = p_oSBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            oCFLCreationParams.UniqueID = pUniqueID
            oCFLCreationParams.MultiSelection = pMultiSelection
            oCFLCreationParams.ObjectType = pObjectType
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = pAlias
            oCon.Operation = pOperation
            oCon.CondVal = pCondValue
            oCFL.SetConditions(oCons)
            AddChooseFromListByFilter = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function Successfully", "AddChooseFromList()")
        Catch ex As Exception
            ShowErr(ex.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Complete Function With Error" & ex.Message, "AddChooseFromList()")
        End Try
    End Function
End Module

