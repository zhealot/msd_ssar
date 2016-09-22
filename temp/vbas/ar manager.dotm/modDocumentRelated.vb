Attribute VB_Name = "modDocumentRelated"
'===================================================================================================================================
' Module:       modDocumentRelated
' Purpose:      Contains generalised document centric code used by the RDA and SSAR addins.
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


'===================================================================================================================================
' Procedure:    AddDocumentVariables
' Purpose:      Adds document variables to an assessment report so that in the event of a problem we can see what
'               components were used to create it.
'
' On Entry:     dataFile                The name of the input file used to create the assessment report.
'               instructionFile         The name of the instruction file used to create the assessment report.
'               theTemplate             The template name used to create the assessment report.
'               g_addinVersionManager   The version number of the Manager addin (ar manager).
'               g_addinVersionRDA       The version number of the RDA addin used to create the assessment report.
'                                       This document variable will not exist if this is a SSAR assessment report.
'               g_addinVersionSSAR      The version number of the SSAR addin used to create the assessment report.
'                                       This document variable will not exist if this is a Manifest 2/3 assessment report.
'===================================================================================================================================
Public Sub AddDocumentVariables(ByVal dataFile As String, _
                                ByVal instructionFile As String, _
                                ByVal theTemplate As String)
    Const c_proc As String = "modDocumentRelated.AddDocumentVariables"

    On Error GoTo Do_Error

    ' We don't need to add the Assessment Report DOcument Vriable as that is defined as part of the template
    With g_assessmentReport
        .Variables.Add mgrDVDataFile, dataFile
        .Variables.Add mgrDVInstructionFile, instructionFile
        .Variables.Add mgrDVTemplate, theTemplate
        .Variables.Add mgrDVManagerAddinVersion, g_addinVersionManager

        ' Add just the document variable appropriate to the assessment report being produced
        If g_rootData.ManifestVersion <= 3 Then
            .Variables.Add mgrDVRDAAddinVersion, g_addinVersionRDA
        Else
            .Variables.Add mgrDVSSARAddinVersion, g_addinVersionSSAR
        End If

        ' This one is essential as it indicates that the current document is an assessment report
        .Variables.Add mgrDVAssessmentReport, CStr(True)
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' AddDocumentVariables

'=======================================================================================================================
' Procedure:    AddWatermark
' Purpose:      Adds the appropriate Watermark to the Assessment Report.
' Note:         Thanks microsoft another FWB. There have been numerous problems with this code not working when
'               launching Word using a 'rep' file. If Word is already running - no problem, but if the OS has to launch
'               Word some documents (not all) cause Word to hang.
'
' On Entry:     g_assessmentReport  Is the document to add the watermark to.
'=======================================================================================================================
Public Sub AddWatermark()
    Const c_proc As String = "modDocumentRelated.AddWatermark"

    Dim bbName                  As String
    Dim doDocumentProtection    As DocumentProtection
    Dim headerItem              As Word.HeaderFooter
    Dim mainStory               As Word.Range
    Dim target                  As Word.Range
    Dim useTemplate             As Word.Template

    On Error GoTo Do_Error

    EventLog "AR Manager." & c_proc

    ' Unprotect the Assessment Report (if it is protected) so that we can add the watermark.
    ' On termination the instantiated class object will reprotect the document for us.
    Set doDocumentProtection = New DocumentProtection
    doDocumentProtection.DisableProtection True

    ' Select the watermark appropriate to the report type
    Select Case g_rootData.ReportVersion
    Case mgrVersionTypeDraft
        bbName = mgrBBWatermarkDraft

    Case mgrVersionTypeFinal
        bbName = mgrBBWatermarkFinal

    Case mgrVersionTypeManagers
        bbName = mgrBBWatermarkManagers

    End Select

    ' This is the template that contains all Building Blocks we insert into the Assessment Report
    Set useTemplate = g_assessmentReport.AttachedTemplate

    ' Locate the headers to insert the watermark in
    For Each headerItem In g_assessmentReport.Sections(1).Headers
        If headerItem.Exists Then

            ' Because there have been so many FWB problems, log what is being done in case there is an error
            Select Case headerItem.index
            Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                EventLog "Updating primary header", c_proc
            Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                EventLog "Updating first page header", c_proc
            Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                EventLog "Updating even page header", c_proc
            End Select

            Set target = headerItem.Range
            target.Collapse wdCollapseEnd

            ' Without this line of code Word occasionally goes into an infinite when inserting the Builiding Block. This problem only
            ' occurs if Word if not running when the 'rep' file is opened. If Word is already running you never encounter the problem.
            ' Activate the Header being updated
            target.Select

            ' Insert the selected Building Block that contains the required watermark
            EventLog "Inserting Building Block: " & bbName, c_proc
            useTemplate.BuildingBlockEntries(bbName).Insert target, True
        End If
    Next

' Select the main story to exit from the header
Set mainStory = g_assessmentReport.Content
mainStory.Collapse wdCollapseStart
mainStory.Select

    EventLog "Setting view type to Print View", c_proc
    g_assessmentReport.ActiveWindow.ActivePane.View.Type = wdPrintView

Do_Exit:

    ' Reprotect the Assessment Report
    EventLog "Completed", c_proc
    Set doDocumentProtection = Nothing

    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' AddWatermark

'=======================================================================================================================
' Procedure:    CloseHiddenDocuments
' Purpose:      Closes any hidden documents that could be left if there was an error producing the previous assessment
'               report.
'=======================================================================================================================
Public Sub CloseHiddenDocuments()
    Const c_proc As String = "modDocumentRelated.CloseHiddenDocuments"

    Dim index       As Long
    Dim theDocument As Word.Document

    On Error GoTo Do_Error

    EventLog c_proc

    ' Iterate all documents, if we find a hidden document then close it
    For index = Documents.Count To 1 Step -1
        Set theDocument = Documents(index)

        ' We are only interested if the document is hidden
        If Not theDocument.ActiveWindow.Visible Then

            ' Close it and discard any changes as it wont contain anything we need
            theDocument.Close wdDoNotSaveChanges
            index = Documents.Count
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CloseHiddenDocuments

'=======================================================================================================================
' Procedure:    HasWatermark
' Purpose:      Determines whether the Assessment Report has a Watermark.
'
' Returns:      True if the Assessment Report has a Watermark.
'=======================================================================================================================
Public Function HasWatermark() As Boolean
    If IsAssessmentReport Then
        HasWatermark = (g_assessmentReport.Sections(1).Headers(wdHeaderFooterPrimary).Range.ShapeRange.Count > 0)
    End If
End Function ' HasWatermark

'=======================================================================================================================
' Procedure:    RemoveWatermark
' Purpose:      Removes the Watermark from the Assessment Report.
'=======================================================================================================================
Public Sub RemoveWatermark()
    Const c_proc As String = "modDocumentRelated.RemoveWatermark"

    Dim doDocumentProtection    As DocumentProtection
    Dim headerItem              As Word.HeaderFooter
    Dim target                  As Word.Range
    Dim theWatermark            As Word.ShapeRange

    On Error GoTo Do_Error

    ' Make sure there is a watermark before we try to delete it
    If HasWatermark Then

        ' Unprotect the Assessment Report (if it is protected) so that we can delete the watermark.
        ' On termination the instantiated class object will reprotect the document for us.
        Set doDocumentProtection = New DocumentProtection
        doDocumentProtection.DisableProtection True

        ' Remove all watermarks from the document
        For Each headerItem In g_assessmentReport.Sections(1).Headers
            If headerItem.Exists Then
                Set target = headerItem.Range

                ' Get a reference to the watermark
                Set theWatermark = target.ShapeRange

                ' Now delete the watermark
                theWatermark.Delete
            End If
        Next
    End If

Do_Exit:

    ' Reprotect the Assessment Report
    Set doDocumentProtection = Nothing

    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RemoveWatermark

'=======================================================================================================================
' Procedure:    ScrollToTopOfDocument
' Purpose:      Scrolls the Assessment Report to the first editable area or the top of the document if there are no
'               editable areas.
' Notes:        There are no editable area when the 'openView' attribute is Read Only View or Print View.
'
' On Entry:     g_assessmentReport  Is the document to scroll.
'=======================================================================================================================
Public Sub ScrollToTopOfDocument()
    Const c_proc As String = "modDocumentRelated.ScrollToTopOfDocument"

    Dim theRange    As Word.Range

    On Error GoTo Do_Error

     ' If the Assessment Report is Writeable it must have editable areas, so scroll to the
     ' first editable area and position the cursor at the start of the editable area
    If g_rootData.IsWritable Then

        ' Set the Range object to the Assessment Reports entire content
        Set theRange = g_assessmentReport.Content

        ' Collapse the Range object to the start of theAssessment Report
        theRange.Collapse wdCollapseStart

        ' We have to use the Selection object for this as the GoToEditableRange method is not part of the Range object.
        ' Position the Selection object at the start of the Assessment Report.
        theRange.Select

        ' Now use the Selection objects GoToEditableRange method to locate the first editable area (the first area with an Editor)
        Set theRange = Selection.GoToEditableRange(wdEditorEveryone)

        ' Make sure we actually located an editable area
        If Not theRange Is Nothing Then

            ' Now position the active Selection at the start of the editable area
            theRange.Collapse wdCollapseStart
            theRange.Select
        End If
    Else

        ' Scroll to the top of the Assessment Report as there are no editable areas
        g_assessmentReport.ActiveWindow.ScrollIntoView g_assessmentReport.Content
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ScrollToTopOfDocument
