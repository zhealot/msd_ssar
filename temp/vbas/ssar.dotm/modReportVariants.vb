Attribute VB_Name = "modReportVariants"
'===================================================================================================================================
' Module:       modReportVariants
' Purpose:      Contains code to produce the Full report and Summary Report variants of an assessment report from a Base Report.
' Note 1:       Base Report is the name given to the default assessment report that is produced.
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
' History:      19/08/16    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module


'===================================================================================================================================
' Procedure:    GenerateFullReport
' Purpose:      Creates a Full Report variant of an assessment report from the Base Report.
' Note 1:       The Base Report should still be useable after creating a Full Report.
' Date:         19/08/16    Created.
'===================================================================================================================================
Public Sub GenerateFullReport()
    Const c_proc As String = "modReportVariants.GenerateFullReport"

    Dim fullReportActions   As Actions
    Dim fullReportFullName  As String

    On Error GoTo Do_Error

    ' Get the path and name for saving the Full Report document
    fullReportFullName = g_configuration.FullReportFileFullName

    ' The Actions block that contains the instructions necessary to Strip out the unwanted text blocks from the Full Report
    Set fullReportActions = g_instructions.FullReport

    ' Ceate the Full Report using the current assessment report (Base Report)
    GenerateReportVariant fullReportFullName, fullReportActions

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' GenerateFullReport

'===================================================================================================================================
' Procedure:    GenerateSummaryReport
' Purpose:      Creates a Summary Report variant of an assessment report from the Base Report.
' Note 1:       The Base Report should still be useable after creating a Summary Report.
' Date:         20/08/16    Created.
'===================================================================================================================================
Public Sub GenerateSummaryReport()
    Const c_proc As String = "modReportVariants.GenerateSummaryReport"

    Dim summaryReportActions    As Actions
    Dim summaryReportFullName   As String

    On Error GoTo Do_Error

    ' Get the path and name for saving the Full Report document
    summaryReportFullName = g_configuration.SummaryReportFileFullName

    ' The Actions block that contains the instructions necessary to Strip out the unwanted text blocks from the Full Report
    Set summaryReportActions = g_instructions.SummaryReport

    ' Ceate the Full Report using the current assessment report (Base Report)
    GenerateReportVariant summaryReportFullName, summaryReportActions

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' GenerateSummaryReport

'===================================================================================================================================
' Procedure:    GenerateReportVariant
' Purpose:      Generates Full Report and Summary Report assessment report variants.
' Note 1:       A lot of care is taken so that the data structures used by the standard assessment report (Base Report) are still
'               useable and the Base Report can continue to be used and edited by the user.
'
' Date:         20/08/16    Created.
'
' On Entry:     variantReportFullName   The file name to use when saving the variant report.
'               variantActions          The Actions list that strips out the unwanted text blocks.
'===================================================================================================================================
Public Sub GenerateReportVariant(ByVal variantReportFullName As String, _
                                 ByVal variantActions As Actions)
    Const c_proc As String = "modReportVariants.GenerateReportVariant"

    Dim assessmentReportFullName    As String
    Dim doDocumentProtection        As DocumentProtection
    Dim noScreenUpdating            As ScreenUpdates

    On Error GoTo Do_Error

    ' Prevent screen updates as it's messy and confusing for the user to look at and it slows things down
    Set noScreenUpdating = New ScreenUpdates

    ' Perform a Full Refresh so that the variant Report being generated is up to date
    If g_rootData.IsWritable Then
        SSARFullRefresh
    End If

    ' Add this document variable to the assessment report so the the AR Manager addin does not reset the User Interface and all
    ' of its related data. This way after creating a variant Report the assessment report can be reopened and the user can continue
    ' editing it or even Submit it to RDA.
    g_assessmentReport.Variables.Add mgrDVPreserveUI, CStr(True)

    ' Get the path and name to use for saving and/or reopening the Assessment Report
    assessmentReportFullName = g_configuration.AssessmentReportFileFullName

    ' If the assessment has not yet been saved, then we need to save it, because we in turn need to create a second document (the
    ' variant Report) and this can only be done if the first document (the asessment report) has already been saved.
    If LenB(g_assessmentReport.Path) = 0 Then

        ' Now save the assessment report
        g_assessmentReport.SaveAs2 assessmentReportFullName, wdFormatDocumentDefault, True, vbNullString, False
    Else

        ' The assessment report has previously been saved, so just save it again
        g_assessmentReport.Save
    End If

    ' Now that we know the assessment report has been saved, we can do another SaveAs and know that we are creating a second document.
    ' There's some subtle stuff that occurs here, with g_assessmentReport being used to create the variant Report without the internal
    ' data structures being changed. This is turn allows us to reopen the assessment report (Base Report) and assign it to
    ' g_assessmentReport and still have the User Interface work!

    ' Now save another copy of the assessment report that will become the variant Report. This of course closes the Base assessment
    ' report. Once the variant Report has been created we will reopen the Base assessment report.
    g_assessmentReport.SaveAs2 variantReportFullName, wdFormatDocumentDefault, True, vbNullString, False

    ' Unprotect the assessment report (if it is protected) so that we can delete the necessary text blocks from it.
    ' On termination the instantiated class object will reprotect the document for us.
    Set doDocumentProtection = NewDocumentProtection
    doDocumentProtection.DisableProtection True

    ' Strip out the unwanted text blocks from the copy of the Base Report so that it becomes the variant Report
    variantActions.BuildAssessmentReport

    ' Strip all Editors from the variant Report so that it is not editable
    If g_rootData.IsWritable Then
        DeleteAllEditors
    End If

    ' Delete the document variable that indicates that a document is an assessment report
    g_assessmentReport.Variables(mgrDVAssessmentReport).Delete

    ' Reprotect the Assessment Report
    Set doDocumentProtection = Nothing

    ' Save the variant Report so that the user is not prompted to save changes when they close it (just makes things a little tidier)
    g_assessmentReport.Save

    ' Clear the undo buffer
    g_assessmentReport.UndoClear

    ' Reopen the original assessment report (Base Report)
    Set g_assessmentReport = Documents.Open(assessmentReportFullName, False, False, False, vbNullString)

    ' Restore and repaint the screen
    Set noScreenUpdating = Nothing

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' GenerateReportVariant

'===================================================================================================================================
' Procedure:    DeleteAllEditors
' Purpose:      Removes all Editors from a variant report (Full Report or Summary Report) so that it is not editable.
' Note 1:       Make sure that the Editors are not deleted from our editable bookmarks object.
' Date:         19/08/16    Created.
'===================================================================================================================================
Private Sub DeleteAllEditors()
    Const c_proc As String = "modReportVariants.DeleteAllEditors"

    Dim dummy           As Word.Range
    Dim editorRange     As Word.Range

    On Error GoTo Do_Error

    ' Set the starting point for the search for editable ranges to the start of the document
    Set dummy = g_assessmentReport.Content
    dummy.Collapse wdCollapseStart
    dummy.Select

    ' You can't iterate the Editors in any simple fashion, so we do it this way
    Do

        ' The ONLY way to find Editable ranges is to use the Selection object
        Selection.GoToEditableRange (wdEditorEveryone)
        Set editorRange = Selection.Range

        ' Delete the current Editor as it includes the paragraph mark
        editorRange.Editors(1).Delete
    Loop

Do_Exit:
    Exit Sub

Do_Error:

    ' If the above code throws error 5941 we deleted all of the Editors
    If Err.Number = mgrErrNoRequestedCollectionMemberDoesNotExist Then
        Resume Do_Exit
    End If

    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DeleteAllEditors
