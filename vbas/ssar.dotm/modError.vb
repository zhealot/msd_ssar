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
' History:      12/06/16    1.  Created.
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
    ShowError procName, ssarTitle, optionalText

    ' Log the error
    EventLogError procName, ssarTitle, optionalText

    ' Tidy up...
    Err.Clear
End Sub ' ErrorReporter

'===================================================================================================================================
' Procedure:    EventLog
' Purpose:      Writes an entry to the event.
' Notes:        This procedure overrides (by scope definition) the EventLog procedure in the 'AR Manager' addin.
' Date:         23/07/2016
'
' On Entry:     logText             The text to write to the log file.
'               procedureName       The procedure name to be displayed in the log file.
'===================================================================================================================================
Public Sub EventLog(ByVal logText As String, _
                    Optional procedureName As String)
    EventLog2 logText, ssarAddinName, procedureName
End Sub ' EventLog

