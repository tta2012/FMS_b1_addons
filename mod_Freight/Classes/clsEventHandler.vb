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


Public Class clsEventHandler
    <DllImport("User32.dll", ExactSpelling:=False, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
    Public Shared Function MoveWindow(ByVal hwnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function
    Public WithEvents SBO_Application As SAPbouiCOM.Application ' holds connection with SBO
    Public TaskDS As SAPbouiCOM.DBDataSource

    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim oUDSs As SAPbouiCOM.UserDataSources
    Dim sql As String = String.Empty

    Dim docfolderpath As String
    Dim BoolResize As Boolean = False
    Dim rowIndex As Integer

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

    Dim RPOmatrixname As String
    Dim RPOsrfname As String
    Dim RGRmatrixname As String
    Dim RGRsrfname As String
    Private Enum ImageType
        TIFF
        BMP
        JPG
        PNG
    End Enum

    Public Sub New()
        Dim sErrDesc As String = String.Empty
        Dim sFuncName As String = "Class_Initialize()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO Application handle", sFuncName)
            SBO_Application = p_oApps.GetApplication
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Retriving SBO application company handle", sFuncName)
            p_oUICompany = SBO_Application.Company

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initialize Trucking Instruction", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up event filters", sFuncName)
            If SetFilters(sErrDesc) <> RTN_SUCCESS Then
                Throw New ArgumentException(sErrDesc)
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
        Catch exc As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
            Call WriteToLogFile(exc.Message, sFuncName)
        End Try
    End Sub

    Public Function SetApplication(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetApplication()
        '   Purpose    :    This function will be calling to initialize the default settings
        '                   such as Retrieving the Company Default settings, Creating Menus, and
        '                   Initialize the Event Filters
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim sFuncName As String = "SetApplication()"
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        Dim sPath As String
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemDefaultData()", sFuncName)
            GetSystemDefaultData(p_oCompDef, sErrDesc)


            oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oMenuItem = SBO_Application.Menus.Item("43520")

            '=========== Start Creating Menus =============
            'TODO to create function in the Common Function module files seperately with this class
            sPath = Application.StartupPath.ToString
            sPath = sPath.Remove(sPath.Length - 3, 3)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oCreationPackage.UniqueID = "mnuFMS"
            oCreationPackage.String = "Freight Management" 'To Change
            oCreationPackage.Image = Application.StartupPath.ToString & "\fms.png"
            oCreationPackage.Enabled = True
            oCreationPackage.Position = 14
            oMenus = oMenuItem.SubMenus

            If Not oMenus.Exists("mnuFMS") Then
                oMenus.AddEx(oCreationPackage)
                oMenuItem = SBO_Application.Menus.Item("mnuFMS")
                oMenus = oMenuItem.SubMenus

                ''Create a sub menu
                'oCreationPackage.UniqueID = "mnuImportSeaLCL"
                'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                'oCreationPackage.String = "Import Sea LCL"
                'oMenus.AddEx(oCreationPackage)

                'oCreationPackage.UniqueID = "mnuImportSeaFCL"
                'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                'oCreationPackage.String = "Import Sea FCL"
                'oMenus.AddEx(oCreationPackage)

                oCreationPackage.UniqueID = "mnuSearchForm"
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oCreationPackage.String = "Find Job"
                oMenus.AddEx(oCreationPackage)
            End If
            '=========== End Creating Menus =============

            SetApplication = RTN_SUCCESS
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function completed with SUCCESS", sFuncName)
        Catch exc As Exception
            SetApplication = RTN_ERROR
            sErrDesc = exc.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
            Call WriteToLogFile(exc.Message, sFuncName)
        End Try
    End Function

    Private Function SetFilters(ByRef sErrDesc As String) As Long
        ' **********************************************************************************
        '   Function   :    SetFilters()
        '   Purpose    :    This function will be gathering to declare the event filter 
        '                   before starting the AddOn Application
        '               
        '   Parameters :    ByRef sErrDesc AS string
        '                       sErrDesc = Error Description to be returned to calling function
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        ' **********************************************************************************
        Dim oFilters As SAPbouiCOM.EventFilters
        Dim oFilter As SAPbouiCOM.EventFilter
        Dim sFuncName As String = "SetFilters()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Initializing EventFilters object", "SetFilters()")
            'Create a new EventFilters object
            oFilters = New SAPbouiCOM.EventFilters

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_MENU_CLICK filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
            oFilter.AddEx("mnuImportSeaLCL")
            oFilter.AddEx("mnuImportSeaFCL")
            oFilter.AddEx("2000000200")



            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_ITEM_PRESSED filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("SendAlert")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("9999")
            oFilter.AddEx("VESSEL")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("MULTIJOBN")
            oFilter.AddEx("DETACHJOBN")
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_MENU_CLICK filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_CHOOSE_FROM_LIST filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("MULTIJOBN")
            oFilter.AddEx("DETACHJOBN")
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_CLICK filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("9999")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("MULTIJOBN")
            oFilter.AddEx("DETACHJOBN")
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_LOST_FOCUS filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_GOT_FOCUS filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_COMBO_SELECT filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("MULTIJOBN")
            oFilter.AddEx("DETACHJOBN")
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_DATA_LOAD filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_DATA_ADD filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("VESSEL")
            oFilter.AddEx("CHARGES") 'MSW to Edit New Ticket
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_DATA_UPDATE filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("134")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("VESSEL")
            oFilter.AddEx("CHARGES") 'MSW to Edit New Ticket
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_DATA_DELETE filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
            oFilter.AddEx("VESSEL")
            oFilter.AddEx("CHARGES") 'MSW to Edit New Ticket

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_DOUBLE_CLICK filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            oFilter.AddEx("2000000200")
            oFilter.AddEx("9999")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_KEY_DOWN filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("9999")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_VALIDATE filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_CLOSE filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_UNLOAD filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("142")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_LOAD filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTAIR")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Setting up et_FORM_LOAD filter", "SetFilters()")
            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            oFilter.AddEx("IMPORTSEALCL")
            oFilter.AddEx("IMPORTSEAFCL")
            oFilter.AddEx("EXPORTSEAFCL")
            oFilter.AddEx("EXPORTSEALCL")
            oFilter.AddEx("EXPORTLAND")
            oFilter.AddEx("IMPORTLAND")
            oFilter.AddEx("IMPORTAIR")
            oFilter.AddEx("FUMIGATION")
            oFilter.AddEx("CRANE")
            oFilter.AddEx("OUTRIDER")
            oFilter.AddEx("BUNKER")
            oFilter.AddEx("FORKLIFT")
            oFilter.AddEx("TOLL")
            oFilter.AddEx("CRATE")
            oFilter.AddEx("LOCAL")
            oFilter.AddEx("Image")
            oFilter.AddEx("ShowBigImg")
            oFilter.AddEx("EXPORTAIRLCL")
            oFilter.AddEx("142")
            oFilter.AddEx("134")
            oFilter.AddEx("VOUCHER")
            oFilter.AddEx("2000000200")
            oFilter.AddEx("2000000201")
            oFilter.AddEx("SHIPPINGINV")
            oFilter.AddEx("2000000005")
            oFilter.AddEx("2000000007")
            oFilter.AddEx("2000000009")
            oFilter.AddEx("2000000010")
            oFilter.AddEx("2000000011")
            oFilter.AddEx("2000000012")
            oFilter.AddEx("2000000013")
            oFilter.AddEx("2000000014")
            oFilter.AddEx("2000000015")
            oFilter.AddEx("2000000016")
            oFilter.AddEx("2000000020")
            oFilter.AddEx("2000000021")
            oFilter.AddEx("2000000025")
            oFilter.AddEx("2000000026")
            oFilter.AddEx("2000000027")
            oFilter.AddEx("2000000028")
            oFilter.AddEx("2000000029")
            oFilter.AddEx("2000000030")
            oFilter.AddEx("2000000031")
            oFilter.AddEx("2000000032")
            oFilter.AddEx("2000000033")
            oFilter.AddEx("2000000034")
            oFilter.AddEx("2000000035")
            oFilter.AddEx("2000000036")
            oFilter.AddEx("2000000037")
            oFilter.AddEx("2000000038")
            oFilter.AddEx("2000000039")
            oFilter.AddEx("2000000040")
            oFilter.AddEx("2000000041")
            oFilter.AddEx("2000000042")
            oFilter.AddEx("2000000043")
            oFilter.AddEx("2000000044")
            oFilter.AddEx("Excel")
            oFilter.AddEx("2000000050") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000051") 'MSW 14-09-2011 Truck PO
            oFilter.AddEx("2000000052") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000053") 'MSW 14-09-2011 Dispatch PO
            oFilter.AddEx("2000000060")
            oFilter.AddEx("2000000061")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adding filters", "SetFilters()")
            SBO_Application.SetFilter(oFilters)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with SUCCESS", sFuncName)
            SetFilters = RTN_SUCCESS
        Catch exc As Exception
            sErrDesc = exc.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", sFuncName)
            SetFilters = RTN_ERROR
        Finally
            oFilter = Nothing
            GC.Collect()        'Forces garbage collection of all generations.
        End Try
    End Function

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_AppEvent()
        '   Purpose    :    This function will be handling the SAP Application Event
        '               
        '   Parameters :    ByVal EventType As SAPbouiCOM.BoAppEventTypes
        '                       EventType = set the SAP UI Application Eveny Object        
        ' **********************************************************************************
        Dim sErrDesc As String
        Dim FunctionName As String = "SBO_Application_AppEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", FunctionName)
            If EventType = SAPbouiCOM.BoAppEventTypes.aet_ShutDown Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SBO shut down event is fired.Shutting down the addon", FunctionName)
                End 'shut down
            ElseIf EventType = SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SBO company changed event is fired.Shutting down the addon", FunctionName)
                End
            ElseIf EventType = SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("SBO shut down event is fired.Shutting down the addon", FunctionName)
                End 'shut down
            End If
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with SUCCESS", FunctionName)
        Catch exc As Exception
            sErrDesc = exc.Message
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", FunctionName)
            Call WriteToLogFile(sErrDesc, FunctionName)
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_FormDataEvent()
        '   Purpose    :    This function will be handling the SAP FormData Event
        '               
        '   Parameters :    ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo
        '                       BusinessObjectInfo = set the SAP UI BusinessObjectInfo Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Dim sErrDesc As String = String.Empty
        Dim FunctionName As String = "SBO_Application_FormDataEvent"
        Try
            Select Case BusinessObjectInfo.FormTypeEx
                Case "142"

                    If AlreadyExist("IMPORTSEALCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLFormDataEvent()", FunctionName)
                        If modImportSeaLCL.DoImportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If AlreadyExist("IMPORTAIR") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLFormDataEvent()", FunctionName)
                        If modImportSeaLCL.DoImportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If AlreadyExist("IMPORTLAND") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLFormDataEvent()", FunctionName)
                        If modImportSeaLCL.DoImportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If AlreadyExist("IMPORTSEAFCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLFormDataEvent()", FunctionName)
                        If modImportSeaFCL.DoImportSeaFCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If AlreadyExist("EXPORTSEALCL") Or AlreadyExist("EXPORTAIRLCL") Or AlreadyExist("EXPORTLAND") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLFormDataEvent()", FunctionName)
                        If modExportSeaLCL.DoExportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If AlreadyExist("EXPORTSEAFCL") Or AlreadyExist("EXPORTAIRFCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaFCLFormDataEvent()", FunctionName)
                        If modExportSeaFCL.DoExportSeaFCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                Case "LCLVOUCHER"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLFormDataEvent()", FunctionName)
                    If modImportSeaLCL.DoImportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "IMPORTSEALCL", "IMPORTAIR", "IMPORTLAND", "LOCAL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLFormDataEvent()", FunctionName)
                    If modImportSeaLCL.DoImportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "IMPORTSEAFCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLFormDataEvent()", FunctionName)
                    If modImportSeaFCL.DoImportSeaFCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "EXPORTSEALCL", "EXPORTAIRLCL", "EXPORTLAND"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLFormDataEvent()", FunctionName)
                    If modExportSeaLCL.DoExportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "EXPORTSEAFCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaFCLFormDataEvent()", FunctionName)
                    If modExportSeaFCL.DoExportSeaFCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "2000000200"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoSearchFormDataEvent()", FunctionName)
                    If Not modSearch.DoSearchFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) Then Throw New ArgumentException(sErrDesc)

                Case "SHIPPINGINV", "VOUCHER", "2000000005", "2000000007", "2000000020", "2000000021", "2000000011", "2000000012", "2000000009", _
                    "2000000010", "2000000013", "2000000014", "2000000015", "2000000016", "2000000026", "2000000025", "2000000027", "2000000028", _
                    "2000000029", "2000000030", "2000000031", "2000000032", "2000000033", "2000000034", "2000000035", "2000000036", "2000000037", _
                    "2000000038", "2000000039", "2000000040", "2000000041", "2000000042", "2000000043", "2000000044", "Image", "VESSEL", "CHARGES", "2000000050", "2000000051", "2000000052", "2000000053", "134", "2000000060", "2000000061"  'MSW to Edit New Ticket

                    If AlreadyExist("IMPORTSEALCL") Or AlreadyExist("IMPORTAIR") Or AlreadyExist("IMPORTLAND") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLFormDataEvent()", FunctionName)
                        If modImportSeaLCL.DoImportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If AlreadyExist("IMPORTSEAFCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLFormDataEvent()", FunctionName)
                        If modImportSeaFCL.DoImportSeaFCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If AlreadyExist("EXPORTSEALCL") Or AlreadyExist("EXPORTAIRLCL") Or AlreadyExist("EXPORTLAND") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLFormDataEvent()", FunctionName)
                        If modExportSeaLCL.DoExportSeaLCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If AlreadyExist("EXPORTSEAFCL") Or AlreadyExist("EXPORTAIRFCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaFCLFormDataEvent()", FunctionName)
                        If modExportSeaFCL.DoExportSeaFCLFormDataEvent(BusinessObjectInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If


            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", FunctionName)
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, FunctionName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed function with ERROR", FunctionName)
        Finally
            GC.Collect()    'Forces garbage collection of all generations.
        End Try
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_MenuEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByRef pVal As SAPbouiCOM.MenuEvent
        '                       pVal = set the SAP UI MenuEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False        
        ' **********************************************************************************

        'Dim ImportSeaLCLForm As SAPbouiCOM.Form
        'Dim oEditText As SAPbouiCOM.EditText
        'Dim oCombo As SAPbouiCOM.ComboBox
        'Dim oMatrix As SAPbouiCOM.Matrix
        'Dim oCheckBox As SAPbouiCOM.CheckBox
        'Dim oOpt As SAPbouiCOM.OptionBtn
        Dim sErrDesc As String = ""
        'Dim oItem As SAPbouiCOM.Item
        Dim sFuncName As String = "SBO_Application_MenuEvent()"

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            Select Case pVal.MenuUID
                Case "mnuImportSeaLCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLMenuEvent()", sFuncName)
                    If modImportSeaLCL.DoImportSeaLCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "mnuImportSeaFCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLMenuEvent()", sFuncName)
                    If modImportSeaFCL.DoImportSeaFCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "mnuExportSeaLCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLMenuEvent()", sFuncName)
                    If modExportSeaLCL.DoExportSeaLCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "mnuSearchForm"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoSearchFormMenuEvent()", sFuncName)
                    If Not modSearch.DoSearchFormMenuEvent(pVal, BubbleEvent, sErrDesc) Then Throw New ArgumentException(sErrDesc)

                Case "1281", "1282", "1288", "1289", "1290", "1291", "1292", "1293", "EditVoc", "EditShp", "EditCPO", "CopyToCGR", "CancelPO"

                    If AlreadyExist("IMPORTSEALCL") Or AlreadyExist("IMPORTAIR") Or AlreadyExist("IMPORTLAND") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLMenuEvent()", sFuncName)
                        If modImportSeaLCL.DoImportSeaLCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If AlreadyExist("IMPORTSEAFCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLMenuEvent()", sFuncName)
                        If modImportSeaFCL.DoImportSeaFCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    If AlreadyExist("EXPORTSEALCL") Or AlreadyExist("EXPORTAIRLCL") Or AlreadyExist("EXPORTLAND") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLMenuEvent()", sFuncName)
                        If modExportSeaLCL.DoExportSeaLCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If AlreadyExist("EXPORTSEAFCL") Or AlreadyExist("EXPORTAIRFCL") Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLMenuEvent()", sFuncName)
                        If modExportSeaFCL.DoExportSeaFCLMenuEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch exc As Exception
            BubbleEvent = False
            ShowErr(exc.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
        Finally
            GC.Collect() 'Forces garbage collection of all generations.
        End Try
    End Sub

    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_ItemEvent()
        '   Purpose    :    This function will be handling the SAP Menu Event
        '               
        '   Parameters :    ByVal FormUID As String
        '                       FormUID = set the FormUID
        '                   ByRef pVal As SAPbouiCOM.ItemEvent
        '                       pVal = set the SAP UI ItemEvent Object
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False
        '
        ' **********************************************************************************

        Dim sErrDesc As String = ""
        Dim sFuncName As String = "SBO_Aspplication_ItemEvent()"
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If Not p_oDICompany.Connected Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectDICompSSO()", sFuncName)
                If ConnectDICompSSO(p_oDICompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
            End If

            Select Case pVal.FormTypeEx
                Case "IMPORTSEALCL", "IMPORTLAND", "IMPORTAIR", "LOCAL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLItemEvent()", sFuncName)
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If modImportSeaLCL.DoImportSeaLCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                Case "IMPORTSEAFCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLItemEvent()", sFuncName)
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If modImportSeaFCL.DoImportSeaFCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                Case "LCLVOUCHER"
                    'From ImportSeaLCLForm
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLItemEvent()", sFuncName)
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If modImportSeaLCL.DoImportSeaLCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                Case "2000000200"
                    'From SearchForm
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoSearchFormItemEvent()", sFuncName)
                    If Not modSearch.DoSearchFormItemEvent(pVal, BubbleEvent, sErrDesc) Then Throw New ArgumentException(sErrDesc)

                Case "2000000201"
                    'From IgniterForm
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoCreateNewJobItemEvent()", sFuncName)
                    If Not modCreateNewJob.DoCreateNewJobItemEvent(pVal, BubbleEvent, sErrDesc) Then Throw New ArgumentException(sErrDesc)

                Case "SHIPPINGINV", "VOUCHER", "2000000005", "2000000007", "2000000009", "2000000010", "2000000011", "2000000012", "2000000013", _
                    "2000000014", "2000000015", "2000000016", "2000000020", "2000000021", _
                     "2000000025", "2000000026", "2000000027", "2000000028", "2000000029", _
                        "2000000030", "2000000031", "2000000032", "2000000033", "2000000034", _
                         "2000000035", "2000000036", "2000000037", "2000000038", "2000000039", "2000000040", "2000000041", "2000000042", "2000000043", "2000000044",
                         "Image", "ShowBigImg", "Excel", "9999", "2000000050", "2000000051", "2000000052", "2000000053", "MULTIJOBN", "DETACHJOBN", "2000000060", "2000000061"
                    'From ImportSeaLCLForm

                    If (AlreadyExist("IMPORTSEALCL") Or AlreadyExist("IMPORTLAND") Or AlreadyExist("IMPORTAIR")) And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLItemEvent()", sFuncName)
                        If modImportSeaLCL.DoImportSeaLCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    'From ImportSeaFCLForm
                    If AlreadyExist("IMPORTSEAFCL") And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLItemEvent()", sFuncName)
                        If modImportSeaFCL.DoImportSeaFCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                    'From ExportSeaLCLForm
                    If (AlreadyExist("EXPORTSEALCL") Or AlreadyExist("EXPORTAIRLCL") Or AlreadyExist("EXPORTLAND")) And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLItemEvent()", sFuncName)
                        If modExportSeaLCL.DoExportSeaLCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If (AlreadyExist("EXPORTSEAFCL") Or AlreadyExist("EXPORTAIRFCL")) And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaFCLItemEvent()", sFuncName)
                        If modExportSeaFCL.DoExportSeaFCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                Case "EXPORTSEALCL", "EXPORTAIRLCL", "EXPORTLAND", "SendAlert"
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaLCLItemEvent()", sFuncName)
                        If modExportSeaLCL.DoExportSeaLCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                Case "EXPORTSEAFCL", "EXPORTAIRFCL"
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then
                        'If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaFCLItemEvent()", sFuncName)
                        If modExportSeaFCL.DoExportSeaFCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                Case "FUMIGATION", "OUTRIDER", "CRANE", "BUNKER", "FORKLIFT", "TOLL", "CRATE"
                    If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.EventType <> SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD Then

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoExportSeaFCLItemEvent()", sFuncName)
                        If modFumigation.DoExportSeaFCLItemEvent(pVal, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

            End Select
        Catch exc As Exception
            BubbleEvent = False
            ShowErr(exc.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.RightClickEvent
        ' **********************************************************************************
        '   Function   :    SBO_Application_ItemEvent()
        '   Purpose    :    This function will be handling the SAP Right Click Event
        '               
        '   Parameters :    ByRef eventInfo As SAPbouiCOM.ContextMenuInfo
        '                       eventInfo = set the ContextMenuInfo event
        '                   ByRef BubbleEvent As Boolean
        '                       BubbleEvent = set the True/False   
        '
        ' **********************************************************************************
        Dim sErrDesc As String = ""
        Dim sFuncName As String = "SBO_Application_RightClickEvent"
        Dim ImportSeaLCLForm As SAPbouiCOM.Form = Nothing
        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            ImportSeaLCLForm = p_oSBOApplication.Forms.Item(eventInfo.FormUID)
            Select Case ImportSeaLCLForm.TypeEx
                Case "IMPORTSEAFCL"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaFCLRightClickEvent()", sFuncName)
                    If modImportSeaFCL.DoImportSeaFCLRightClickEvent(eventInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "IMPORTSEALCL", "IMPORTAIR", "IMPORTLAND"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLRightClickEvent()", sFuncName)
                    If modImportSeaLCL.DoImportSeaLCLRightClickEvent(eventInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                Case "EXPORTSEALCL", "EXPORTSEAFCL", "EXPORTAIRLCL", "EXPORTLAND", "2000000028", "2000000025",
                    "2000000031", "2000000034", "2000000037", "2000000040", "FUMIGATION", "OUTRIDER", "CRANE", "BUNKER", "FORKLIFT", "TOLL", "CRATE"
                    If (AlreadyExist("EXPORTSEALCL") Or AlreadyExist("EXPORTAIRLCL") Or AlreadyExist("EXPORTLAND")) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ()", sFuncName)
                        If modExportSeaLCL.DoExportSeaLCLRightClickEvent(eventInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If
                    If (AlreadyExist("EXPORTSEAFCL") Or AlreadyExist("EXPORTAIRFCL")) Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ()", sFuncName)
                        If modExportSeaFCL.DoExportSeaFCLRightClickEvent(eventInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                    End If

                Case "VOUCHER"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling DoImportSeaLCLRightClickEvent()", sFuncName)
                    If modImportSeaLCL.DoImportSeaLCLRightClickEvent(eventInfo, BubbleEvent, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            End Select

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            BubbleEvent = False
            ShowErr(ex.Message)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            WriteToLogFile(Err.Description, sFuncName)
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
