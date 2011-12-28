Option Explicit On 

Imports System.IO

Module modLogging
    '***************************************
    'Name       :   modLogging
    'Descrption :   Contains function for logging errors and Application related information
    'Author     :   Dev (Satish Dwivedi)
    'Created    :   09 March 2004
    '***************************************
    Private Const MAXFILESIZE_IN_MB As Int16 = 5 '(2 MB)
    Private Const LOG_FILE_ERROR As String = "ErrorLog"
    Private Const LOG_FILE_ERROR_ARCH As String = "ErrorLog_"
    Private Const LOG_FILE_DEBUG As String = "DebugLog"
    Private Const LOG_FILE_DEBUG_ARCH As String = "DebugLog_"
    Private Const FILE_SIZE_CHECK_ENABLE As Int16 = 1
    Private Const FILE_SIZE_CHECK_DISABLE As Int16 = 0

    Public Function WriteToLogFile(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long
        'Function   :   WriteToLogFile()
        'Purpose    :   This function checks if given input file name exists or not
        'Parameters :   ByVal strErrText As String
        '                   strErrText = Text to be written to the log
        '               ByVal intLogType As Integer
        '                   intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '               ByVal strSourceName As String
        '                   strSourceName = Function name calling this function
        '               Optional ByVal intCheckFileForDelete As Integer
        '                   intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   DEV (SATISH DWEVEDI)
        'Date       :   09 March 2004

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String
        Dim strArchFileName As String
        Dim strTempString As String
        Dim lngFileSizeInMB As Double

        Try
            strTempString = Space(40 - Len(strSourceName))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText
            strFileName = System.Windows.Forms.Application.StartupPath & "\" & LOG_FILE_ERROR & ".log"
            strArchFileName = System.Windows.Forms.Application.StartupPath & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "YYMMDDHHMMSS") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function WriteToLogFile_Debug(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long
        'Function   :   WriteToLogFile_Debug()
        'Purpose    :   This function checks if given input file name exists or not
        'Parameters :   ByVal strErrText As String
        '                   strErrText = Text to be written to the log
        '               ByVal intLogType As Integer
        '                   intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '               ByVal strSourceName As String
        '                   strSourceName = Function name calling this function
        '               Optional ByVal intCheckFileForDelete As Integer
        '                   intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        'Return     :   0 - FAILURE
        '               1 - SUCCESS
        'Author     :   DEV (SATISH DWEVEDI)
        'Date       :   09 March 2004
        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String
        Dim strArchFileName As String
        Dim strTempString As String
        Dim lngFileSizeInMB As Double

        Try
            strTempString = Space(40 - Len(strSourceName))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText
            strFileName = System.Windows.Forms.Application.StartupPath & "\" & LOG_FILE_DEBUG & ".log"
            strArchFileName = System.Windows.Forms.Application.StartupPath & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "YYMMDDHHMMSS") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile_Debug = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile_Debug = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

End Module
