Option Explicit On

Imports System.Drawing.Printing
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Configuration
Imports System.IO
Imports System.Net.Mail

Public Class clsReportUtilities
    Public Enum ExportFileType
        PDFFILE = 0
        WORDFILE = 1
    End Enum
    
#Region "CreateJobFolderFile"

#End Region
    ''' <summary>
    ''' Create Job Folder And File Path To Save the Job File.
    ''' </summary>
    ''' <param name="mainFolder">Set the main folder where you want to save your job folder.</param>
    ''' <param name="jobNo">Set jobno to create as folder name.</param>
    ''' <param name="pdfFileName">Set your job pdf File Name without extension.</param>
    ''' <returns> Returns pdfFileName to export PDF. </returns>
    ''' <remarks></remarks>
    Public Function CreateJobFolderFile(ByVal mainFolder As String, ByVal jobNo As String, ByVal pdfFileName As String) As String
        Dim jobDir As String = mainFolder & jobNo
        Dim pdfpath As String = String.Empty
        Try
            Dim di As DirectoryInfo = New DirectoryInfo(jobDir)
            If Not di.Exists Then
                di.Create()
            End If

            Dim str() As String = System.IO.Directory.GetFiles(jobDir, "*.pdf*", IO.SearchOption.AllDirectories)
            Dim strfile As New ArrayList
            Dim strfilename As String

            Dim filecount As String
            If Not str.Length = 0 Then
                For i = 0 To str.Length - 1
                    If Path.GetFileNameWithoutExtension(str(i)).Length > pdfFileName.Length Then
                        If Path.GetFileNameWithoutExtension(str(i)).Substring(0, pdfFileName.Length) = pdfFileName Then
                            strfile.Add(str(i))
                        End If
                    End If
                Next
                If Not strfile.Count = 0 Then
                    Dim str1 As Integer = Convert.ToInt32(Right(Path.GetFileNameWithoutExtension(strfile.Item(strfile.Count - 1)), 2)) + 1
                    If str1.ToString.Length = 1 Then
                        filecount = "0" & str1.ToString
                    Else
                        filecount = str1.ToString
                    End If
                    strfilename = jobDir + "\" + pdfFileName + " - " + filecount + ".pdf"
                    pdfpath = strfilename
                End If
            End If
            If pdfpath = String.Empty Then
                pdfpath = jobDir + "\" + pdfFileName + " - 01.pdf"
            End If
            Return pdfpath
        Catch ex As Exception
            Return pdfpath
        End Try
    End Function

#Region "ExportCRToPDF"
    ''' <summary>
    ''' Export the Job File CrystalReport to PDF 
    ''' </summary>
    ''' <param name="filePath">Set specific filepath where you want to save.</param>
    ''' <param name="fileType">File Type which you want to export.</param>
    ''' <param name="rptDocument">Report Document Parameter to export the file.</param>
    ''' <remarks></remarks>
    Public Sub ExportCRToPDF(ByVal filePath As String, ByVal fileType As ExportFileType, ByVal rptDocument As ReportDocument, ByVal showPDF As Boolean)
        Dim crExportOptions As New CrystalDecisions.Shared.ExportOptions
        Dim crDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
        crDiskFileDestinationOptions.DiskFileName = filePath
        crDiskFileDestinationOptions.Clone()
        crExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
        Select Case fileType
            Case ExportFileType.PDFFILE
                crExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
            Case ExportFileType.WORDFILE
                crExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.WordForWindows
        End Select
        crExportOptions.ExportDestinationOptions = crDiskFileDestinationOptions
        SetDBLogIn(rptDocument)
        rptDocument.Export(crExportOptions)
        'rptDocument.Close()
        'to km
        If showPDF = True Then
            Process.Start(filePath)
        End If

    End Sub
#End Region

    Public Sub SetDBLogIn(ByRef rpt As ReportDocument)
        rpt.SetDatabaseLogon(p_fmsSetting.UserID, p_fmsSetting.Password, p_fmsSetting.ServerName, p_fmsSetting.DBName)
        'Dim str As String
        'str = p_fmsSetting.UserID & "," & p_fmsSetting.Password & "," & p_fmsSetting.ServerName & "," & p_fmsSetting.DBName
        'p_oSBOApplication.MessageBox(str)
    End Sub

    Public Sub PrintDoc(ByRef rpt As ReportDocument)
        Dim newThread As New Thread(AddressOf ShowPrint)
        newThread.Priority = ThreadPriority.Highest
        newThread.SetApartmentState(ApartmentState.STA)
        newThread.Start(rpt)
        newThread.Join()
    End Sub
    
    Private Shared Sub ShowPrint(ByVal rpt As Object)
        Try
            Dim printdl As New PrintDialog
            Dim pdoc As New PrintDocument

            'Dim strfile As String = rpt
            'pdoc.DocumentName = strfile
            'pdoc.PrinterSettings.PrinterName = "DocuWorks Printer"
            'printdl.Document = pdoc
            'printdl.PrintToFile = True
            ''pdoc.PrinterSettings.
            'pdoc.PrinterSettings.PrintFileName = "D:\Test1.xdw"
            'pdoc.Print()
            printdl.AllowPrintToFile = True
            printdl.UseEXDialog = True
            printdl.PrintToFile = True
            ' Dim dresult As DialogResult = printdl.ShowDialog()
            '  If dresult = DialogResult.OK Then

            Dim printerName As String = printdl.PrinterSettings.PrinterName
            rpt.PrintOptions.PrinterName = "DocuWorks Printer"
            printdl.PrinterSettings.PrintFileName = "D:\TT.xdw"
            rpt.PrintToPrinter(printdl.PrinterSettings.Copies, False, 0, 0)

            ' pdoc.Print()
            '  End If
        Catch ex As Exception

        End Try
        'MessageBox.Show(dlg.FileName)
    End Sub

    Public Function SendMailDoc(ByVal SendTo As String, ByVal mailFrom As String, ByVal Subject As String, ByVal Body As String, Optional ByVal attachmail As String = "") As Boolean
        SendMailDoc = False
        Try
            Dim msg As New MailMessage(mailFrom, SendTo, Subject, Body)
            Dim smtpser As New SmtpClient("midserver", 25)
            If attachmail <> "" Then

                Dim attach As Attachment = New Attachment(attachmail)
                msg.Attachments.Add(attach)
            End If
            smtpser.Send(msg)
            SendMailDoc = True
        Catch ex As Exception
            SendMailDoc = False
        End Try

    End Function

End Class
