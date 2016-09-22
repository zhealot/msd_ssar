Attribute VB_Name = "AutoExec"
'===================================================================================================================================
' Module:       AutoExec
' Purpose:      The only purpose of this module is to house the AutoExec procedure (aka Main).
' Note:         After 'cleaning' this project Word generally cannot find this procedure so it may require manually exporting and
'               importing and even being manually edited before being recognised by Word.
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
'               30/05/16    1.  NEW     Updated to store manager version number.
'===================================================================================================================================
Option Explicit


Public Sub Main()
    Const c_proc As String = "AutoExec.Main"

    On Error GoTo Do_Error

    ' Store the version number of this addin
    g_addinVersionManager = ThisDocument.Variables(mgrDVVersion).Value

    Debug.Print "AR Manager." & c_proc & " [version: " & g_addinVersionManager & "]"

    ' The first thing we need to do is load and initialise the Configuration object.
    ' Until we load fetchdoc.infopathxml, which in turn initialises g_rootData,
    ' the Configuration object is only partially functional!
    Set g_configuration = New Configuration
    g_configuration.Initialise

    Debug.Print "Config file loaded"

    ' We can't log anything until the config file is loaded as that has the event log path and name info in it
    Set g_eventLog = Nothing
    EventLog vbCr & c_proc

    ' Wire up the Application Event handler so that we can monitor what files are being opened.
    ' When we intercept an '*.rep' file being opened we use this as the trigger for creating an Assessment Report from it.
    Set g_wordEvents = New ApplicationEvents

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Main

