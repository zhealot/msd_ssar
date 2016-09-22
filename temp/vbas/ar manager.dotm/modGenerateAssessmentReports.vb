Attribute VB_Name = "modGenerateAssessmentReports"
'===================================================================================================================================
' Module:       modGenerateAssessmentReports
' Purpose:      This is where the assessment report generation kicks off.
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
' History:      30/05/16    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module

'=======================================================================================================================
' Procedure:    ?
' Purpose:      ?.
'
' On Entry:     x                   .
' On Exit:      x.
'=======================================================================================================================
Public Sub dummy()
    Const c_proc As String = "xxx.Dummy"

    On Error GoTo Do_Error


Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Dummy

'=======================================================================================================================
' Procedure:    GenerateAssessmentReport
' Purpose:      Generates the Manifest 2, Manifest 3 and SSAR assessment reports.
'
' On Entry:     theInputFileFullPath    The path to the rep file that the user tried to open using Word.
'=======================================================================================================================
Public Sub GenerateAssessmentReport(ByVal theInputFileFullPath As String)
    Const c_proc As String = "modGenerateAssessmentReports.StartRDAProcessingHere"

    Dim assessmentWindow    As Word.Window
    Dim infoText            As String
    Dim instructionFile     As String
    Dim mainStory           As Word.Range
    Dim startTime           As Variant
    Dim theFSO              As Scripting.FileSystemObject
    Dim theTemplate         As String

    On Error GoTo Do_Error

    ' This is a "just in case" statement that cleans up any hidden documents left from a previous execution
    CloseHiddenDocuments

    ' Always clear this flag in case it remains set at the end of producing a previous Assessment Report
    g_hasBeenSubmitted = False

    ' This is a chicken and egg situation. We need to load the configuration file before loading the rep file, even though
    ' some of the configuration data is dependent upon which manifest the rep file is using. So certain functionality of the
    ' Configuration object is not fully available until the rep file has been loaded (and the Rootdata object initialised).
    If g_configuration Is Nothing Then
        Err.Raise mgrErrNoUnexpectedCondition
    End If

    ' If we are in development mode we can override the default input file by displaying a dialog thats allows us to choose another file
    If g_configuration.PromptForFile Then
        theInputFileFullPath = PromptForFile(theInputFileFullPath)
    End If

    ' Save the input file path and name as we will need it if we need to reload the input file
    g_configuration.InputXMLFileFullPath = theInputFileFullPath

    ' Reset the temporary file names or else the same names get reused and since they
    ' contain information from the Assessment Report xml this can be misleading
    g_configuration.ResetGeneratedFileNames

    ' Generate the assessment report file name, needs to be done just once so the name does not change (it contains a timestamp)
    g_configuration.GenerateAssessmentReportFileName

    startTime = Now

    ' Load the rda xml document
    EventLog2 "Loading rda input file: " & theInputFileFullPath, mgrAddinName, c_proc
    If LoadAssessmentReport(theInputFileFullPath) Then

        ' Load the instruction file that contains the instructions for building the Assessment report
        instructionFile = g_configuration.CurrentInstructionsFileFullName
        EventLog2 "Loading instruction file: " & instructionFile, mgrAddinName, c_proc
        If LoadInstructionData(instructionFile) Then

            ' See if the Action file should be validated (useful after any changes and during testing).
            ' Note: This is not the standard DOM Document validation. These are specific validation tests embedded in the instruction
            '       file that allow assumptions about the xml data in the rep file to be proved or disproved. Used in development.
            If g_configuration.ShouldValidateInstructionsFile Then

                ' Run the validation tests
                If Not g_configuration.PerformActionsFileValidation Then

                    ' One or more of the tests failed - so we stop here
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If

        ' Create new (invisible - as it is much faster to create the document this way) Word Document
        EventLog2 "Creating Assessment Report document", mgrAddinName, c_proc
        theTemplate = g_configuration.CurrentTemplateFullName
        Set g_assessmentReport = Documents.Add(theTemplate, False, wdNewBlankDocument, False)

        EventLog2 "Adding document variables to the Assessment Report document", mgrAddinName, c_proc
        AddDocumentVariables theInputFileFullPath, instructionFile, theTemplate

        ' Fill in the document
        EventLog2 "Next step is BuildDocument", mgrAddinName, c_proc
        If g_rootData.ManifestVersion <= 3 Then
            BuildRDAAssessmentReport
        Else
            BuildSSARAssessmentReport
        End If

        If g_configuration.CloseWordHTMLDocument Then
            If Not g_htmlWordDocument Is Nothing Then
                g_htmlWordDocument.Close wdDoNotSaveChanges
            End If

            ' Delete the temporary file used to supply RichText to the Assessment Report
            If g_configuration.WordHTMLTextFileDelete Then
                Set theFSO = New Scripting.FileSystemObject
                theFSO.DeleteFile (g_configuration.WordHTMLTextFileFullName)
                Set theFSO = Nothing
            End If
        Else
            If Not g_htmlWordDocument Is Nothing Then
                g_htmlWordDocument.ActiveWindow.Visible = True
            End If
        End If

        ' Get a reference to the assessment Window
        EventLog2 "Setting assessment window reference", mgrAddinName, c_proc
        Set assessmentWindow = g_assessmentReport.ActiveWindow

        ' For some bizare reason what works for Manifest 2 and 3 does not work for SSAR assessment reports
        If g_rootData.ManifestVersion > 3 Then

            EventLog2 "Starting SSAR post generation Window tidy up", mgrAddinName, c_proc

            ' For some reason the order of the following code has a huge impact on whether it actually
            ' works and can cause Word to crash! So consider yourself warned if you mess with this!!
    
            ' Make sure the Assessment Report is not minimised
            Application.WindowState = wdWindowStateNormal

            ' Scroll to the first editable area or the top of the document if there are no editable areas
            ScrollToTopOfDocument

            ' Make sure the document is displayed in Print Layout view
            assessmentWindow.View = wdPrintView

            ' Display the Word document we've been building
            assessmentWindow.Activate
            assessmentWindow.Visible = True
        Else

            EventLog2 "Starting Manifest 2/3 post generation Window tidy up", mgrAddinName, c_proc

            ' Set a reference to the documents main story, then collapse it so that we are at the start of the document
            Set mainStory = g_assessmentReport.Content
            mainStory.Collapse wdCollapseStart

            ' Select the main story reference which will cause the document to exit the Header and return to the main document body
            EventLog2 "Step 1", mgrAddinName, c_proc
            mainStory.Select

            EventLog2 "Step 2", mgrAddinName, c_proc
            If assessmentWindow.View.SplitSpecial = wdPaneNone Then
                EventLog2 "Step 3 - True", mgrAddinName, c_proc
                assessmentWindow.ActivePane.View.Type = wdPrintView
            Else
                EventLog2 "Step 3 - False", mgrAddinName, c_proc
                assessmentWindow.View.Type = wdPrintView
            End If

            EventLog2 "Step 4", mgrAddinName, c_proc
            assessmentWindow.Activate
            EventLog2 "Step 5", mgrAddinName, c_proc
            assessmentWindow.Visible = True
        End If

        ' Always try to eliminate any hidden documents as they can prevent this addin from
        ' reinitialising because Word will not actually close with hidden documents open
        If g_configuration.IsDevelopment Then
            CloseHiddenDocuments
        End If
    End If

    infoText = "Elapsed time (in seconds): " & DateDiff("s", startTime, Now)
    Debug.Print infoText
    Debug.Print g_configuration.AssessmentReportPassword

    EventLog2 infoText, mgrAddinName, c_proc

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' GenerateAssessmentReport
