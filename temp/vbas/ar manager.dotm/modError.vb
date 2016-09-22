Attribute VB_Name = "modError"
'===================================================================================================================================
' Module:       modError
' Purpose:      Handles error reporting and event logging.
'
' Author:       Peter Hewett - Inner Word Limited (innerword@xnet.co.nz)
' Copyright:    Ministry of Social Development (MSD) ©2016 All rights reserved.
' Contact       Inner Word Limited
' details:      134 Kahu Road
'               Paremata
'               Porirua City
'               5024
'               T: +64 4 233 2124
'               M: +64 21 213 5063
'               E: innerword@xnet.co.nz
'
' History:      12/05/16    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit


'=======================================================================================================================
' Proc:         ErrorReporter
' Purpose:      Error reporting procedure.
' Note:         Call this procedure to both display and log an error.
'               The error information is expected in the Err object.
'
' On Entry:     procName            The name of the procedure that triggered the error.
'               optionalText        Optional text displayed after the error message.
'=======================================================================================================================
Public Sub ErrorReporter(ByVal procName As String, _
                         Optional ByVal optionalText As String)

    ' Now tell the user about the error that has just occurred
    ShowError procName, mgrTitle, optionalText

    ' Log the error
    EventLogError procName, mgrTitle, optionalText

    ' Tidy up...
    Err.Clear
End Sub ' ErrorReporter

Public Sub XMLErrorReporter(ByVal theParseError As MSXML2.IXMLDOMParseError, _
                            ByVal ErrorMessage As String, _
                            ByVal errorProc As String)
    Dim errorText As String
    Dim title     As String

    With theParseError
        errorText = "Error (" & .ErrorCode & "): " & .reason & vbCr & _
                    "Line #: " & .Line & vbCr & _
                    "Line position: " & .linepos & vbCr & _
                    "Position in file: " & .filepos & vbCr & _
                    "Source text: " & .srcText & vbCr & _
                    "Document URL: " & .Url

        If LenB(ErrorMessage) > 0 Then
            errorText = ErrorMessage & vbCr & errorText
        End If
    End With

    ' Make sure the log object exists
    If g_eventLog Is Nothing Then
        SetupEventLog
    End If

    ' Log the error
    EventLog errorText, errorProc

    ' Title for the MsgBox to provide some context for the user
    title = mgrTitle & " - Error"

    ' Now tell the user about the xml error that just occured
    errorText = "Procedure: " & errorProc & vbCr & errorText
    MsgBox errorText, vbCritical Or vbOKOnly, title
End Sub ' XMLErrorReporter

Public Sub EventLog(ByVal logText As String, _
                    Optional ByVal procedureName As String)
    If g_eventLog Is Nothing Then
        SetupEventLog
    End If

    If LenB(procedureName) > 0 Then
        logText = procedureName & " - " & logText
    End If

    If Not g_eventLog Is Nothing Then
        g_eventLog.Output logText
    End If
End Sub ' EventLog

Public Sub EventLog2(ByVal logText As String, _
                     ByVal addinName As String, _
                     Optional ByVal procedureName As String)
    If g_eventLog Is Nothing Then
        SetupEventLog
    End If

    If LenB(procedureName) > 0 Then
        logText = addinName & "." & procedureName & " - " & logText
    Else
        logText = addinName & " - " & logText
    End If

    If Not g_eventLog Is Nothing Then
        g_eventLog.Output logText
    End If
End Sub ' EventLog2

Private Sub SetupEventLog()
    Const c_proc As String = "modError.SetupEventLog"

    On Error GoTo Do_Error

    ' Create the event log file
    If Not g_configuration Is Nothing Then
        Set g_eventLog = New LogIt
        With g_eventLog
            .AutoPad = True
            .TimeStamp = False
            .path = g_configuration.EventLogPath
            .File = g_configuration.EventLogFile
            .enabled = g_configuration.EventLogging
        End With
    Else
        Debug.Print c_proc & " g_configuration is Nothing "
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ShowError c_proc, mgrTitle
    Resume Do_Exit
End Sub ' SetupEventLog

'=======================================================================================================================
' Proc:         ShowError
' Purpose:      Displays an error message to the user.
' Note:         Call this procedure to display an error.
'               Does not clear the Err object.
'
' On Entry:     procName            The name of the procedure that triggered the error.
'               addinName           The name of the addin that raised the error.
'               optionalText        Optional text displayed after the error message.
'=======================================================================================================================
Public Sub ShowError(ByVal procName As String, _
                     ByVal addinName As String, _
                     Optional ByVal optionalText As String)
    Dim title As String

    ' Title for the MsgBox to provide some context for the user
    title = addinName & " - Error"

    ' When the optional text is present display it on its own line
    If LenB(optionalText) > 0 Then
        optionalText = optionalText & vbCrLf
    End If

    MsgBox "Error: (" & Err.Number & ") " & Err.Description & vbCrLf & optionalText & vbCrLf & _
           "Procedure: " & procName, vbCritical Or vbOKOnly, title
End Sub ' ShowError

'=======================================================================================================================
' Procedure:    EventLogError
' Purpose:      Writes an error entry to the event log.
' Notes:        Sets up the event log if necessary.
'
' On Entry:     procName            The name of the procedure that threw the error.
'               addinName           The name of the addin that raised the error.
'               optionalText        Optional text displayed after the error message.
'=======================================================================================================================
Public Sub EventLogError(ByVal procName As String, _
                         ByVal addinName As String, _
                         Optional ByVal optionalText As String)
    Const c_proc As String = "modError.EventLogError"

    Dim errorText As String

    ' Keep it simple, just the procedure name, the error number and its description.
    ' Do this before instantiating our own error handler or the Err object will be reset.
    errorText = addinName & "." & procName & " - Error: (" & Err.Number & ") " & Err.Description

    On Error GoTo Do_Error

    ' Modify the optional text so we can append it to the error message
    If LenB(optionalText) > 0 Then
        optionalText = ". " & optionalText
    End If

    ' Make sure the log object exists
    If g_eventLog Is Nothing Then
        SetupEventLog
    End If

    ' Now write the log entry
    g_eventLog.Output errorText & optionalText

Do_Exit:
    Exit Sub

Do_Error:
    ShowError c_proc, mgrTitle
    Resume Do_Exit
End Sub ' EventLogError
