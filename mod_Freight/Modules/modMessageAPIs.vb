Option Explicit On 

Module modMessageAPIs

    'Public Enum POINTAPI
    '    x
    '    Y
    'End Enum
    Public Structure POINT
        Public x As Integer
        Public y As Integer
    End Structure

    'Public Enum Msg
    '    hwnd
    '    message
    '    wParam
    '    lParam
    '    Time
    '    pt
    'End Enum
    Public Structure MSG
        Public hwnd As Integer
        Public message As Integer
        Public wParam As Integer
        Public lParam As Integer
        Public itime As Integer
        Public pt_x As Integer
        Public pt_y As Integer
    End Structure


    '// Retrieves messages sent to the calling thread's message queue
    'Public Declare Function GetMessage Lib "user32" _
    '    Alias "GetMessageA" _
    '     (ByVal lpMsg As MSG, _
    '      ByVal hwnd As Long, _
    '      ByVal wMsgFilterMin As Long, _
    '      ByVal wMsgFilterMax As Long) As Long
    Declare Function GetMessage Lib "user32" _
        Alias "GetMessageA" _
        (ByRef lpMsg As MSG, _
        ByVal hwnd As Integer, _
        ByVal wMsgFilterMin As Integer, _
        ByVal wMsgFilterMax As Integer) As Integer


    '// Translates virtual-key messages into character messages
    'Public Declare Function TranslateMessage Lib "user32" _
    '    (ByVal lpMsg As MSG) As Long
    Declare Function TranslateMessage Lib "user32" _
        (ByRef lpMsg As MSG) As Integer

    '// Forwards the message on to the window represented by the
    '// hWnd member of the Msg structure
    'Public Declare Function DispatchMessage Lib "user32" _
    '    Alias "DispatchMessageA" _
    '     (ByVal lpMsg As MSG) As Long
    Declare Function DispatchMessage Lib "user32" _
        Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Integer

    Public Msg1 As New MSG

    'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


End Module