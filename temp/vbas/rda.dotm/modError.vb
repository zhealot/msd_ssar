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
    ShowError procName, rdaTitle, optionalText

    ' Log the error
    EventLogError procName, rdaTitle, optionalText

    ' Tidy up...
    Err.Clear
End Sub ' ErrorReporter

