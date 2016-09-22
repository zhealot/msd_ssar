Attribute VB_Name = "modHTMLToXHTML"
'===================================================================================================================================
' Module:       modHTMLToXML
' Purpose:      This module takes all of the editable RichText text entered into the Assessment Report and creates a new Word
'               document from it. This new document is then saved as a "Filtered HTML" file, the "Filtered HTML" file is then
'               parsed into ersatz XHTML. The XHTML is in turn used to update the Assessment Report xml that will be returned to
'               the Remedy/RDA webservice.
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
' History:      11/06/16    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit
Option Private Module


'=======================================================================================================================
' Procedure:    UpdateAssessmentReportXMLUsingRichText
' Purpose:      Updates the assessment report xml DOMDocument using the data the user has entered into the editable
'               areas of the assessment report document.
' Notes:        .
'=======================================================================================================================
Public Sub UpdateAssessmentReportXMLUsingRichText()
    Const c_proc            As String = "modHTMLToXHTML.UpdateAssessmentReportXMLUsingRichText"
    Const c_wingdingsTick   As Long = -3844

    Dim allActions      As Actions
    Dim defaultFontName As String
    Dim paraIndex       As Word.Paragraph
    Dim theBookmark     As Word.Bookmark
    Dim theFSO          As Scripting.FileSystemObject
    Dim wingdingsTick   As String

    On Error GoTo Do_Error

    If Not g_rootData.IsWritable Then
        Exit Sub
    End If

    ' Check that all Bookmarks defined Ranges match their corresponding Editor Range.
    ' It is possible for these to get out of sync if the user cuts and pastes one entire input area
    ' into another entire input area. This destroys the Bookmark of the area being pasted into.
    RebuildBookmarksBeforeEdit

    ' Create a new document to hold the RichText from the Assessment Report
    Set g_xhtmlWordDoc = Documents.Add(g_configuration.CurrentTemplateFullName, False, wdNewBlankDocument, False)

    ' Delete it's entire contents (we just want a blank document that uses the same styles)
' #TODO#    We may also need to deleted the two headers in case the logos end up in the Filtered HTML file
    g_xhtmlWordDoc.Content.Delete

    ' The 'action' list used to build the Assessment Report is now used to create the XHTML file
    Set allActions = g_instructions.Actions

    ' Initialise the counters collection
    Set g_counters = New Counters

    ' Carry out all actions in the Actions list (this builds the HTML file we parse into XHTML)
    allActions.HTMLForXMLUpdate

    ' Delete all bookmarks as they generate extra unwanted elements in the HTML document that we will need to parse
    For Each theBookmark In g_xhtmlWordDoc.bookmarks
        theBookmark.Delete
    Next

    ' Nasty workaround for Tick characters (Wingdings character set) in the
    ' document, without this the ticks get converted to a 'u' with an umlaut.
    wingdingsTick = ChrW$(c_wingdingsTick)
    defaultFontName = g_assessmentReport.Styles(wdStyleNormal).Font.Name

    ' Loop through all paragraphs, reset to default font all those that contain another font unless the text contains a WingDings tick character
    For Each paraIndex In g_xhtmlWordDoc.Paragraphs
        If InStr(paraIndex.Range.Text, wingdingsTick) = 0 Then
            paraIndex.Range.Font.Name = defaultFontName
        End If
    Next

    ' Save the xhtml Word document as a filtered HTML file
    g_xhtmlWordDoc.SaveAs2 g_configuration.WordXHTMLTextFileFullName, WdSaveFormat.wdFormatFilteredHTML, _
                           False, vbNullString, False

    ' Fix for problem with being prompted to save changes to the documents template
    g_xhtmlWordDoc.AttachedTemplate.Saved = True

    ' Now close the document so that it can be opened again as a text (html) file
    g_xhtmlWordDoc.Close

    ' Now we take the xhtml document and parse it as html into ersatz xhtml and then use it to update the Assessment Report xml
    ParseHTMLToXML g_configuration.WordXHTMLTextFileFullName

    ' See if the xhtml file should be deleted
    If g_configuration.WordXHTMLTextFileDelete Then
        Set theFSO = New Scripting.FileSystemObject
        theFSO.DeleteFile (g_configuration.WordXHTMLTextFileFullName)
        Set theFSO = Nothing
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' UpdateAssessmentReportXMLUsingRichText

Public Sub RebuildBookmarksBeforeEdit(Optional ByVal suppressScreenUpdates As Boolean)
    Const c_proc As String = "modHTMLToXHTML.RebuildBookmarksBeforeEdit"

    Dim originalSelection   As Word.Range
    Dim noScreenUpdating    As ScreenUpdates

    On Error GoTo Do_Error

    EventLog "Global bookmark count: " & g_bookmarkCount & ", Actual bookmark count: " & g_assessmentReport.bookmarks.Count, c_proc

    ' #DEBUG# Remove before production deployment - verifies Exception Table integrity
    VerifyAllExceptionsTables

    ' Prevent screen updates as it's messy and confusing for the user to look at and it slows things down
    Set noScreenUpdating = New ScreenUpdates

    ' Save the current selection as the RebuildBookmarks code uses it
    Set originalSelection = Selection.Range.Duplicate

    ' Rebuild the Bookmarks
    RebuildBookmarks

    ' Restore the selection object
    originalSelection.Select

    ' Update the global Bookmark count
    g_bookmarkCount = g_assessmentReport.bookmarks.Count

    ' Restore and repaint the screen
    Set noScreenUpdating = Nothing

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RebuildBookmarksBeforeEdit

'=======================================================================================================================
' Procedure:    RebuildBookmarks
' Purpose:      Checks and rebuilds all Bookmarks for editable areas (Ranges that have an Editor specified).
' Notes:        This needs to be done because the when a user pastes data into an editable area, depending on exactly
'               where the insertion point is, the Bookmark can be deleted or the text pasted after the Bookmark.
'=======================================================================================================================
Public Sub RebuildBookmarks()
    Const c_proc As String = "modHTMLToXHTML.RebuildBookmarks"

    Dim bookmarkName    As String
    Dim bookmarkRange   As Word.Range
    Dim cellRange       As Word.Range
    Dim columnNumber    As Long
    Dim doDocumentProtection    As DocumentProtection
    Dim dummy           As Word.Range
    Dim editorRange     As Word.Range
    Dim fixBookmark     As Boolean
    Dim index           As Long
    Dim info            As String
    Dim rowNumber       As Long

    On Error GoTo Do_Error

    If Not g_editableBookmarks Is Nothing Then

        ' Unprotect the Assessment Report (if it is protected) so that we can refresh the necessary areas.
        ' On termination the instantiated class object will reprotect the document for us.
        Set doDocumentProtection = NewDocumentProtection
        doDocumentProtection.DisableProtection True

        ' Iterate the Editors Dictionary object checking that:
        ' 1. The Bookmark (Dictionary key) actually exists
        ' 2. That the Bookmarks Range matches the Editors Range
        ' If the expected Bookmark does not exist it is created. Likewise if there is
        ' a mismatch in Bookmark Range vs Editor Range the Bookmark will be recreated.
        For index = 1 To g_editableBookmarks.bookmarkCount

            ' Get the key (Bookmark name) and value (an Editor object)
            bookmarkName = g_editableBookmarks.Bookmark(index)

            ' Set the starting point for the search for editable ranges to the start of the document
            If index = 1 Then
                Set dummy = g_assessmentReport.Content
                dummy.Collapse wdCollapseStart
                dummy.Select
            End If

            ' The ONLY way to find Editable ranges is to use the Selection object
            Selection.GoToEditableRange (wdEditorEveryone)
            Set editorRange = Selection.Range

            With g_assessmentReport.bookmarks

                ' Check to see if the bookmark exists
                If .Exists(bookmarkName) Then

                    ' Get the Bookmark Range just the once
                    Set bookmarkRange = .Item(bookmarkName).Range

                    ' Check whether it contains a Content Control, ignore the range check
                    ' if it does as it should not be possible to mess up the bookmark
                    If bookmarkRange.ContentControls.Count = 0 Then

                        ' Now see if it matches the Editor Range
                        If bookmarkRange.Start <> editorRange.Start Or bookmarkRange.End <> editorRange.End Then
                            info = "Difference detected (" & index & "). BM_S,_E: " & _
                                   bookmarkRange.Start & ", " & bookmarkRange.End & " ER_S,_E: " & editorRange.Start & ", " & editorRange.End
                            fixBookmark = True
                        End If
                    End If
                Else
                    info = "Undefined bookmark (" & index & ")"
                    fixBookmark = True
                End If

                ' See if the Bookmark needs recreating or repairing (the same process applies to both)
                If fixBookmark Then

                    ' If the Editor is within a table then we repair the Bookmark differently.
                    ' Information(wdStartOfRangeRowNumber) returns -1 if the Range is not in a table.
                    rowNumber = editorRange.Information(wdStartOfRangeRowNumber)
                    If rowNumber = -1 Then

                        ' See if the last character of the Range is a pargraph mark.
                        ' If it is don't include it as part of the editable range.
                        If Right$(editorRange.Text, 1) = vbCr Then

                            ' Delete the current Editor as it includes the paragraph mark
                            editorRange.Editors(1).Delete
                        End If
                    Else
                        ' This code assumes that the editable range mapped by the Editor is the last text in the Cell.

                        ' Recreate the editor so that its start position remain the same, but its end
                        ' position corresponds to the Range of the Table Cell it is in, but excludes
                        ' the end of cell marker as we do not want that as part of the Editor range.
                        columnNumber = editorRange.Information(wdStartOfRangeColumnNumber)
                        Set cellRange = editorRange.Tables(1).Cell(rowNumber, columnNumber).Range

                        ' Exclude the end of cell marker from the Range
                        editorRange.End = cellRange.End - 1
                    End If

                    ' Recreate the Editor object
                    editorRange.Editors.Add wdEditorEveryone

                    info = info & ", repairing bookmark: " & bookmarkName
                    Debug.Print info
                    EventLog info, c_proc

                    .Add bookmarkName, editorRange

                    fixBookmark = False
                End If
            End With
        Next

        ' Clear the undo buffer to prevent pottential problems
        g_assessmentReport.UndoClear

        EventLog "Editable bookmark count: " & g_editableBookmarks.bookmarkCount, c_proc
    Else
        Err.Raise mgrErrNoEditorsDictionaryUndefined, c_proc, mgrErrTextEditorsDictionaryUndefined
    End If

Do_Exit:

    ' These actions should happen even if there was an error
    Set doDocumentProtection = Nothing

    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RebuildBookmarks

'=======================================================================================================================
' Procedure:    StripBullets
' Purpose:      Strips Bullets from editable areas of the Assessment Report and replaces them with a unique character
'               sequence, so that we can recognise which paragraphs require Bullets when Remedy/RDA sends the data back.
' Notes:        We have to do this as Word does not generate the correct html tags (<ul><li></li></ul>) for bullets.
'               Instead it converts the bullet to a symbol character and pads the text with non-breaking spaces.
'
' On Entry:     targetArea          The area to search for Bulleted paragraphs.
'=======================================================================================================================
Public Sub StripBullets(ByVal targetArea As Word.Range)
    Const c_proc As String = "modUpdateAssessmentReportXML.StripBullets"

    Dim bulletCount         As Long
    Dim indexParagraph      As Word.Paragraph
    Dim paragraphListFormat As Word.ListFormat
    Dim paragraphRange      As Word.Range

    On Error GoTo Do_Error

    ' Check the entire area for Bullets before resorting to the much slower paragraph by paragraph search
    If targetArea.ListFormat.CountNumberedItems > 0 Then

        ' Now iterate each paragraph in the editable area looking for bulleted paragraphs
        For Each indexParagraph In targetArea.Paragraphs

            ' We are only interested if the paragraph has a bullet point
            Set paragraphRange = indexParagraph.Range
            Set paragraphListFormat = paragraphRange.ListFormat
            If paragraphListFormat.CountNumberedItems > 0 Then
                bulletCount = bulletCount + 1

                ' Strip the bullet
                paragraphListFormat.RemoveNumbers wdNumberParagraph

                ' Add the special preface text that allows us to recognise where a bullet should be
                paragraphRange.InsertBefore mgrBulletRecognitionSequence
            End If
        Next

        ' Make sure not too many undo operation accumulate on the undo stack
        g_xhtmlWordDoc.UndoClear
    End If

'''Debug.Print "Bullet count " & CStr(bulletCount)

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' StripBullets
