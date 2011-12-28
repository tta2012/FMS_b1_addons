Option Explicit On

Imports SAPbobsCOM
Imports SAPbouiCOM

Module modCreateNewJob
    Public Function DoCreateNewJobItemEvent(ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean, ByRef sErrDesc As String) As Boolean
        DoCreateNewJobItemEvent = False
        Dim oNewJobForm As SAPbouiCOM.Form = Nothing
        Dim oRecordset As SAPbobsCOM.Recordset = Nothing
        Dim FunctionName As String = "DoCreateNewJobItemEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", FunctionName)
            Select Case pVal.FormTypeEx
                Case "2000000201"
                    oNewJobForm = p_oSBOApplication.Forms.Item("IGNITER")
                    oRecordset = p_oDICompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    If pVal.BeforeAction = False Then
                        If pVal.EventType = BoEventTypes.et_FORM_CLOSE Then
                            If Not RemoveFromAppList(oNewJobForm.UniqueID) Then Throw New ArgumentException(sErrDesc)
                        End If

                        If pVal.EventType = BoEventTypes.et_ITEM_PRESSED Then
                            'MSW To Edit #1001
                            If pVal.ItemUID = "op_Air" Or pVal.ItemUID = "op_Sea" Or pVal.ItemUID = "op_Land" Then
                                If oNewJobForm.Items.Item("op_Sea").Specific.Selected = False Then
                                    oNewJobForm.Items.Item("op_LCL").Visible = False
                                    oNewJobForm.Items.Item("op_FCL").Visible = False
                                    oNewJobForm.Items.Item("op_OPL").Visible = False
                                Else
                                    oNewJobForm.Items.Item("op_LCL").Visible = True
                                    oNewJobForm.Items.Item("op_FCL").Visible = True
                                    oNewJobForm.Items.Item("op_OPL").Visible = True
                                End If
                            End If

                            'Google Doc
                            If pVal.ItemUID = "op_Local" Then
                                oNewJobForm.Items.Item("op_Air").Enabled = False
                                oNewJobForm.Items.Item("op_Sea").Enabled = False
                            End If
                            If pVal.ItemUID = "op_Tshp" Then
                                oNewJobForm.Items.Item("op_Air").Enabled = False
                                oNewJobForm.Items.Item("op_Sea").Enabled = False
                                oNewJobForm.Items.Item("op_Land").Enabled = False
                            End If
                            If pVal.ItemUID = "op_Export" Then
                                oNewJobForm.Items.Item("op_Air").Enabled = True
                                oNewJobForm.Items.Item("op_Sea").Enabled = True
                                oNewJobForm.Items.Item("op_Land").Enabled = True
                            End If
                            If pVal.ItemUID = "op_Import" Then
                                oNewJobForm.Items.Item("op_Air").Enabled = True
                                oNewJobForm.Items.Item("op_Sea").Enabled = True
                                oNewJobForm.Items.Item("op_Land").Enabled = True
                            End If
                           
                            'MSW To Edit
                            If pVal.ItemUID = "bt_Close" Then oNewJobForm.Close()

                            If pVal.ItemUID = "bt_Create" Then
                                Dim JobType As String = String.Empty
                                Dim JobMode As String = String.Empty
                                Dim JobClass As String = String.Empty
                                Dim TypeLCL As Boolean = IIf((oNewJobForm.Items.Item("op_LCL").Specific.Selected = True Or oNewJobForm.Items.Item("op_FCL").Specific.Selected = False), True, False)
                                Dim TypeFCL As Boolean = IIf((oNewJobForm.Items.Item("op_FCL").Specific.Selected = True Or oNewJobForm.Items.Item("op_LCL").Specific.Selected = False), True, False)
                                'MSW To Edit #1001
                                If oNewJobForm.Items.Item("op_Export").Specific.Selected = True Then
                                    JobType = "Export"
                                ElseIf oNewJobForm.Items.Item("op_Import").Specific.Selected = True Then
                                    JobType = "Import"
                                ElseIf oNewJobForm.Items.Item("op_Local").Specific.Selected = True Then
                                    JobType = "Local"
                                ElseIf oNewJobForm.Items.Item("op_Tshp").Specific.Selected = True Then
                                    JobType = "Transhipment"

                                End If

                                If oNewJobForm.Items.Item("op_Air").Specific.Selected = True Then
                                    JobMode = "Air"
                                ElseIf oNewJobForm.Items.Item("op_Sea").Specific.Selected = True Then
                                    JobMode = "Sea"
                                ElseIf oNewJobForm.Items.Item("op_Land").Specific.Selected = True Then
                                    JobMode = "Land"
                                End If

                                If oNewJobForm.Items.Item("op_Gen").Specific.Selected = True Then
                                    JobClass = "GEN"
                                ElseIf oNewJobForm.Items.Item("op_DG1").Specific.Selected = True Then
                                    JobClass = "EXP"
                                ElseIf oNewJobForm.Items.Item("op_DG7").Specific.Selected = True Then
                                    JobClass = "RA"
                                ElseIf oNewJobForm.Items.Item("op_DG").Specific.Selected = True Then
                                    JobClass = "DG"
                                ElseIf oNewJobForm.Items.Item("op_Other").Specific.Selected = True Then
                                    JobClass = "STG"
                                End If
                                'End MSW To Edit #1001
                                If (Left(JobType, 6) = "Export" Or Left(JobType, 6) = "Import") And (JobType = vbNullString Or JobMode = vbNullString Or JobClass = vbNullString) Then
                                    p_oSBOApplication.MessageBox("Need to select options to create new job!")
                                Else
                                    If JobMode = "Sea" And oNewJobForm.Items.Item("op_LCL").Specific.Selected = False And oNewJobForm.Items.Item("op_FCL").Specific.Selected = False Then
                                        p_oSBOApplication.MessageBox("Need to select job mode to create new job!")
                                    Else
                                        'If AlreadyExist("IMPORTSEALCL") Or AlreadyExist("IMPORTSEAFCL") Or AlreadyExist("EXPORTSEALCL") Or AlreadyExist("EXPORTSEAFCL") Or AlreadyExist("EXPORTAIRLCL") Then
                                        '    p_oSBOApplication.MessageBox("Close The Form First.")
                                        'Else
                                        '    If Not LoadSpecificJobForm(JobType, JobMode, JobClass, TypeLCL) Then Throw New ArgumentException(sErrDesc)
                                        'End If
                                        If Not LoadSpecificJobForm(JobType, JobMode, JobClass, TypeLCL) Then Throw New ArgumentException(sErrDesc)
                                    End If


                                End If
                            End If
                        End If
                    End If
            End Select
            DoCreateNewJobItemEvent = True
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed Function", FunctionName)
        Catch ex As Exception
            DoCreateNewJobItemEvent = False
        End Try
    End Function

    Public Function LoadAndCreateJobForm() As Boolean
        LoadAndCreateJobForm = False
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim oComboBox As SAPbouiCOM.ComboBox
        Dim oOpt As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = vbNullString
        Try
            If Not LoadFromXML(p_oSBOApplication, "Igniter.srf") Then Throw New ArgumentException(sErrDesc)
            oActiveForm = p_oSBOApplication.Forms.Item("IGNITER")
            oActiveForm.Freeze(True)
            AddUserDataSrc(oActiveForm, "Export", sErrDesc, BoDataType.dt_SHORT_TEXT, 6)
            AddUserDataSrc(oActiveForm, "Import", sErrDesc, BoDataType.dt_SHORT_TEXT, 6)
            AddUserDataSrc(oActiveForm, "Local", sErrDesc, BoDataType.dt_SHORT_TEXT, 5)
            AddUserDataSrc(oActiveForm, "Tranship", sErrDesc, BoDataType.dt_LONG_TEXT, 12)
            oOpt = oActiveForm.Items.Item("op_Export").Specific
            oOpt.DataBind.SetBound(True, "", "Export")
            oOpt = oActiveForm.Items.Item("op_Import").Specific
            oOpt.DataBind.SetBound(True, "", "Import")
            oOpt.GroupWith("op_Export")
            oOpt = oActiveForm.Items.Item("op_Local").Specific
            oOpt.DataBind.SetBound(True, "", "Local")
            oOpt.GroupWith("op_Export")
            oOpt = oActiveForm.Items.Item("op_Tshp").Specific
            oOpt.DataBind.SetBound(True, "", "Tranship")
            oOpt.GroupWith("op_Export")


            AddUserDataSrc(oActiveForm, "Air", sErrDesc, BoDataType.dt_SHORT_TEXT, 3)
            AddUserDataSrc(oActiveForm, "Sea", sErrDesc, BoDataType.dt_SHORT_TEXT, 3)
            AddUserDataSrc(oActiveForm, "Land", sErrDesc, BoDataType.dt_SHORT_TEXT, 4)
            oOpt = oActiveForm.Items.Item("op_Air").Specific
            oOpt.DataBind.SetBound(True, "", "Air")
            oOpt = oActiveForm.Items.Item("op_Sea").Specific
            oOpt.DataBind.SetBound(True, "", "Sea")
            oOpt.GroupWith("op_Air")
            oOpt = oActiveForm.Items.Item("op_Land").Specific
            oOpt.DataBind.SetBound(True, "", "Land")
            oOpt.GroupWith("op_Air")


            AddUserDataSrc(oActiveForm, "Gen", sErrDesc, BoDataType.dt_SHORT_TEXT, 3)
            AddUserDataSrc(oActiveForm, "DG1", sErrDesc, BoDataType.dt_SHORT_TEXT, 3)
            AddUserDataSrc(oActiveForm, "DG7", sErrDesc, BoDataType.dt_SHORT_TEXT, 3)
            AddUserDataSrc(oActiveForm, "DG", sErrDesc, BoDataType.dt_SHORT_TEXT, 2)
            AddUserDataSrc(oActiveForm, "Other", sErrDesc, BoDataType.dt_SHORT_TEXT, 5)
            oOpt = oActiveForm.Items.Item("op_Gen").Specific
            oOpt.DataBind.SetBound(True, "", "Gen")
            oOpt = oActiveForm.Items.Item("op_DG1").Specific
            oOpt.DataBind.SetBound(True, "", "DG1")
            oOpt.GroupWith("op_Gen")
            oOpt = oActiveForm.Items.Item("op_DG7").Specific
            oOpt.DataBind.SetBound(True, "", "DG7")
            oOpt.GroupWith("op_Gen")
            oOpt = oActiveForm.Items.Item("op_DG").Specific
            oOpt.DataBind.SetBound(True, "", "DG")
            oOpt.GroupWith("op_Gen")
            oOpt = oActiveForm.Items.Item("op_Other").Specific
            oOpt.DataBind.SetBound(True, "", "Other")
            oOpt.GroupWith("op_Gen")



            AddUserDataSrc(oActiveForm, "OPTLCL", sErrDesc, BoDataType.dt_SHORT_TEXT)
            AddUserDataSrc(oActiveForm, "OPTFCL", sErrDesc, BoDataType.dt_SHORT_TEXT)
            AddUserDataSrc(oActiveForm, "OPTOPL", sErrDesc, BoDataType.dt_SHORT_TEXT)
            oOpt = oActiveForm.Items.Item("op_LCL").Specific
            oOpt.DataBind.SetBound(True, "", "OPTLCL")
            oOpt = oActiveForm.Items.Item("op_FCL").Specific
            oOpt.DataBind.SetBound(True, "", "OPTFCL")
            oOpt.GroupWith("op_LCL")
            oOpt = oActiveForm.Items.Item("op_OPL").Specific
            oOpt.DataBind.SetBound(True, "", "OPTOPL")
            oOpt.GroupWith("op_LCL")
            oActiveForm.Freeze(False)

            'oComboBox = oActiveForm.Items.Item("cb_JbType").Specific
            'oComboBox.ValidValues.Add("Import", "Import")
            'oComboBox.ValidValues.Add("Export", "Export")
            'oComboBox = oActiveForm.Items.Item("cb_JbMode").Specific
            'oComboBox.ValidValues.Add("Sea", "Sea")
            'oComboBox.ValidValues.Add("Air", "Air")
            'oComboBox.ValidValues.Add("Land", "Land")
            'oComboBox = oActiveForm.Items.Item("cb_JbClass").Specific
            'oComboBox.ValidValues.Add("GEN", "General Cargo")   'NonDG
            'oComboBox.ValidValues.Add("DG1", "Dangerous Goods")
            'oComboBox.ValidValues.Add("DG7", "Strategic Goods")
            'oActiveForm.Items.Item("op_LCL").Specific.Selected = True
            LoadAndCreateJobForm = True
        Catch ex As Exception
            LoadAndCreateJobForm = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function LoadSpecificJobForm(ByVal pJobType As String, ByVal pJobMode As String, ByVal pJobClass As String, ByVal pTypeLCL As Boolean) As Boolean
        Dim parentform As String = String.Empty
        LoadSpecificJobForm = False
        Dim oActiveForm As SAPbouiCOM.Form = Nothing
        Dim title As String = ""
        Dim jobMode As String = ""

        Try
            'MSW To Edit #1010
            'Select Case pJobClass
            '    Case "GEN"
            '        pJobClass = "GEN"
            '    Case "DG1"
            '        pJobClass = "DG1"
            '    Case "DG7"
            '        pJobClass = "DG7"
            '    Case "DG"
            '        pJobClass = "DG"
            '    Case "Other"
            '        pJobClass = ""
            'End Select
            'End MSW To Edit #1010
            If pJobType = "Import" And Left(pJobMode, 3) = "Sea" And pTypeLCL = True Then
                title = "Import " + Left(pJobMode, 3) + "-LCL " + pJobClass.Trim
                jobMode = "LCL"

            ElseIf pJobType = "Import" And Left(pJobMode, 3) = "Sea" And pTypeLCL = False Then
                title = "Import " + Left(pJobMode, 3) + "-FCL " + pJobClass.Trim
                jobMode = "FCL"
            ElseIf pJobType = "Import" And Left(pJobMode, 3) = "Air" Then
                title = "Import " + Left(pJobMode, 3) + " " + pJobClass.Trim
            ElseIf pJobType = "Import" And Left(pJobMode, 4) = "Land" Then
                title = "Import " + Left(pJobMode, 4) + " " + pJobClass.Trim
            ElseIf pJobType = "Export" And Left(pJobMode, 4) = "Sea" And pTypeLCL = True Then
                title = "Export " + Left(pJobMode, 3) + "-LCL " + pJobClass.Trim
                jobMode = "LCL"
            ElseIf pJobType = "Export" And Left(pJobMode, 4) = "Sea" And pTypeLCL = False Then
                title = "Export " + Left(pJobMode, 3) + "-FCL " + pJobClass.Trim
                jobMode = "FCL"
            ElseIf pJobType = "Export" And Left(pJobMode, 4) = "Air" Then
                title = "Export " + Left(pJobMode, 3) + " " + pJobClass.Trim
            ElseIf pJobType = "Export" And Left(pJobMode, 4) = "Land" Then
                title = "Export " + Left(pJobMode, 4) + " " + pJobClass.Trim
            ElseIf pJobType = "Local" Then
                title = "Local " + pJobClass.Trim
            ElseIf pJobType = "Transhipment" Then
                title = "Transhipment " + pJobClass.Trim
            End If
            modExportSeaFCL.LoadExportSeaFCLForm(vbNullString, pJobType, title)
            oActiveForm = p_oSBOApplication.Forms.Item("EXPORTSEAFCL")
            oActiveForm.Items.Item("ed_JMode").Specific.Value = jobMode
            oActiveForm.Items.Item("ed_CrgType").Specific.Value = pJobClass.Trim

            LoadSpecificJobForm = True
        Catch ex As Exception
            LoadSpecificJobForm = False
            MessageBox.Show(ex.Message)
        End Try
    End Function

End Module
