Option Explicit On

Imports System
Imports System.Collections.Generic

Module modMain

    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16

    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0

    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    Public Const ERR_DISPLAY_STATUS As Int16 = 1
    Public Const ERR_DISPLAY_DIALOGUE As Int16 = 2

    Public Structure FMSSetting
        Public DocuPath As String
        Public ServerName As String
        Public DBName As String
        Public UserID As String
        Public Password As String
        Public PicturePath As String
    End Structure

    Public Structure CompanyDefault
        Public DBName As String
        Public CompanyName As String
        Public LocalCurrency As String
        Public SystemCurrency As String
        Public CurrencyPosition As String
        Public iPriceDecimal As Int16
        Public iQtyDecimal As Int16
        Public HolidaysName As String
    End Structure

    Public AppObjString As List(Of String)
    Public p_vDBList(,) As String

    Public p_oApps As SAPbouiCOM.SboGuiApi
    Public p_oEventHandler As clsEventHandler
    Public p_oSBOApplication As SAPbouiCOM.Application
    Public p_oDICompany As SAPbobsCOM.Company
    Public p_oUICompany As SAPbouiCOM.Company
    Public p_ObjTargetCmp As SAPbobsCOM.Company
    Public p_oCompDef As CompanyDefault
    Public p_fmsSetting As FMSSetting

    Public p_TargetCompUserN As String
    Public p_TargetCompUserP As String

    Sub main(ByVal Args() As String)
        Dim sErrDesc As String = ""
        Try
            p_iDebugMode = DEBUG_ON
            p_iErrDispMethod = ERR_DISPLAY_STATUS

            AppObjString = New List(Of String)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Addon startup function", "Main()")
            p_oApps = New SAPbouiCOM.SboGuiApi
            p_oApps.Connect(Args(0))

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Creating Event handler class", "Main()")
            p_oEventHandler = New clsEventHandler

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing public SBO Application object", "Main()")
            p_oSBOApplication = p_oEventHandler.SBO_Application

            p_oDICompany = New SAPbobsCOM.Company
            p_ObjTargetCmp = New SAPbobsCOM.Company

            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", "Main()")
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling SetApplication()", "Main()")

            If p_oEventHandler.SetApplication(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Addon started successfully", "Main()")

            Do While GetMessage(Msg1, 0&, 0&, 0&) > 0
                TranslateMessage(Msg1)
                DispatchMessage(Msg1)
                System.Windows.Forms.Application.DoEvents()
            Loop

        Catch exp As Exception
            Call WriteToLogFile(exp.Message, "Main()")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", "Main()")
        Finally

        End Try
    End Sub
End Module
