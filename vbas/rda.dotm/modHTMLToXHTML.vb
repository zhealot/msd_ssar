Attribute VB_Name = "modHTMLToXHTML"
'===================================================================================================================================
' Module:       modHTMLToXML
' Purpose:      This module takes all of the editable RichText text entered into the Assessment Report and creates a new Word
'               document from it. This new document is then saved as a "Filtered HTML" file, the "Filtered HTML" file is then used
'               as the data source for updating the Assessment Report xml that will be returned to the Remedy/RDA webservice.
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
' History:      09/12/15    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module

Private m_xhtmlWordDoc As Word.Document


'=======================================================================================================================
' Procedure:    UpdateAssessmentReportXMLUsingRichText
' Purpose:      .
' Notes:        .
'=======================================================================================================================
Public Sub UpdateAssessmentReportXMLUsingRichText()
    Const c_proc            As String = "modHTMLToXHTML.UpdateAssessmentReportXMLUsingRichText"
    Const c_wingdingsTick    As Long = -3844

    Dim allActions      As VBA.Collection
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
    ' into another entire input area. This destroy the Bookmark of the area being pasted into.
    RebuildBookmarksBeforeEdit

    ' Create a new document to hold the RichText from the Assessment Report
    Set m_xhtmlWordDoc = Documents.Add(g_configuration.CurrentTemplateFullName, False, wdNewBlankDocument, False)

    ' Delete it's entire contents (we just want a blank document that uses the same styles)
    m_xhtmlWordDoc.Content.Delete

    ' The 'action' list used to build the Assessment Report is now used to create the XHTML file
    Set allActions = g_instructions.Actions.actionList

    ' Initialise the counters collection
    Set g_counters = New Counters

    ' Parse out each "action" (which results in each of the actions in the list being carried out)
    XhtmlActions allActions

    ' Delete all bookmarks as they generate extra unwanted elements we will need to parse
    For Each theBookmark In m_xhtmlWordDoc.bookmarks
        theBookmark.Delete
    Next

    ' Force all text entered into the Assessment Report to be the default Font (whatever the Normal Style is).
    ' So if the user has formatted text using another font face this overrides it.
'''    m_xhtmlWordDoc.Content.Font.Name = g_assessmentReport.Styles(wdStyleNormal).Font.Name

    ' Nasty workaround for Tick characters (Wingdings character set) in the
    ' document, without this the ticks get converted to a 'u' with an umlaut.
    wingdingsTick = ChrW$(c_wingdingsTick)
    defaultFontName = g_assessmentReport.Styles(wdStyleNormal).Font.Name

    ' Loop through all paragraphs, reset to default font all those that contain another font unless the text contains a WingDings tick character
    For Each paraIndex In m_xhtmlWordDoc.Paragraphs
        If InStr(paraIndex.Range.Text, wingdingsTick) = 0 Then
            paraIndex.Range.Font.Name = defaultFontName
        End If
    Next

    ' Save the xhtml Word document as a filtered HTML file
    m_xhtmlWordDoc.SaveAs2 g_configuration.WordXHTMLTextFileFullName, WdSaveFormat.wdFormatFilteredHTML, _
                           False, vbNullString, False

    ' Fix for problem with being prompted to save changes to the documents template
    m_xhtmlWordDoc.AttachedTemplate.Saved = True

    ' Now close the document so that it can be opened again as a text (html) file
    m_xhtmlWordDoc.Close

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

    Dim originalSelection As Word.Range

    On Error GoTo Do_Error

    EventLog "Global bookmark count: " & g_bookmarkCount & ", Actual bookmark count: " & g_assessmentReport.bookmarks.Count, c_proc

    ' Prevent screen updates as it's messy and confusing for the user to look at and it slows things down
    If Not suppressScreenUpdates Then
        Application.ScreenUpdating = False
    End If

    ' Save the current selection as the RebuildBookmarks code uses it
    Set originalSelection = Selection.Range.Duplicate

    ' Rebuild the Bookmarks
    RebuildBookmarks

    ' Restore the selection object
    originalSelection.Select

    ' Update the global Bookmark count
    g_bookmarkCount = g_assessmentReport.bookmarks.Count

    ' Restore and repaint the screen
    If Not suppressScreenUpdates Then
        Application.ScreenUpdating = True
        Application.ScreenRefresh
    End If

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
'               where the insertion point is, the Bookmark can be deted or the text pasted after the Bookmark.
'=======================================================================================================================
Public Sub RebuildBookmarks()
    Const c_proc As String = "modHTMLToXHTML.RebuildBookmarks"

    Dim bookmarkName  As String
    Dim bookmarkRange As Word.Range
    Dim dummy         As Word.Range
    Dim editorRange   As Word.Range
    Dim fixBookmark   As Boolean
    Dim index         As Long
    Dim info          As String

    On Error GoTo Do_Error

    If Not g_editableBookmarks Is Nothing Then
        
        ' Iterate the Editors Dictionary object checking that:
        ' 1. The Bookmark (Dictionary key) actually exists
        ' 2. That the Bookmarks Range matches the Editors Range
        ' If the expected Bookmark does not exist it is created. Likewise if there is
        ' a mismatch in Bookmark Range vs Editor Range the Bookmark will be recreated.
        For index = 1 To g_editableBookmarks.BookmarkCount

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

                    ' Now see if it matches the Editor Range
                    If bookmarkRange.Start <> editorRange.Start Or bookmarkRange.End <> editorRange.End Then
                        info = "Difference detected"
                        fixBookmark = True
                    End If
                Else
                    info = "Undefined bookmark"
                    fixBookmark = True
                End If

                ' See if the Bookmark need recreating or repairing (the same process applies to both)
                If fixBookmark Then

                    ' See if the last character of the Range is a pargraph mark.
                    ' If it is don't include it as part of the editable range.
                    If Right$(editorRange.Text, 1) = vbCr Then

                        ' Delete the current Editor as it includes the paragraph mark
                        editorRange.Editors(1).Delete

                        ' Exclude the final paragraph mark from the Range
                        editorRange.End = editorRange.End - 1

                        ' Recreate the Editor object
                        editorRange.Editors.Add wdEditorEveryone
                    End If

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
    Else
        Err.Raise mgrErrNoEditorsDictionaryUndefined, c_proc, mgrErrTextEditorsDictionaryUndefined
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RebuildBookmarks

'=======================================================================================================================
' Procedure:    XhtmlActions
' Purpose:      Main processing loop for creating the XHTML Word document
' Notes:        This procedure is indirectly recursive..
'
' On Entry:     allActions          A Collection object that contains the list of 'actions' to be carried out.
'=======================================================================================================================
Public Sub XhtmlActions(ByVal allActions As VBA.Collection)
    Const c_proc As String = "modHTMLToXHTML.XhtmlActions"

    Dim errorText As String
    Dim theAction As Object

    On Error GoTo Do_Error

    EventLog c_proc

    ' Check that the collection exists before using it
    If allActions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all ActionSetup, ActionAdd and ActionInsert objects in the collection
    For Each theAction In allActions

        ' There are four types of objects in this collection so choose the method appropriate to the object
        Select Case TypeName(theAction)
        Case rdaTypeActionAdd
            XhtmlActionAdd theAction

        Case rdaTypeActionInsert
            XhtmlActionInsert theAction

        Case rdaTypeActionLink, rdaTypeActionRename, rdaTypeActionSetup
            ' We do not need to do anything for these actions

        Case Else
            errorText = Replace$(mgrErrTextUnknownActionVerbType, mgrP1, TypeName(theAction))
            Err.Raise mgrErrNoUnknownActionVerbType, c_proc, errorText
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' XhtmlActions

'=======================================================================================================================
' Procedure:    XhtmlActionAdd
' Purpose:      Processes the Add action when creating the XHTML data source.
' Notes:        This procedure is indirectly recursive..
'
' On Entry:     info                The ActionAdd object to be actioned.
'=======================================================================================================================
Private Sub XhtmlActionAdd(ByVal info As ActionAdd)
    Const c_proc As String = "modHTMLToXHTML.XhtmlActionAdd"

    Dim index       As Long
    Dim theQuery    As String
    Dim subNodes    As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(info.Test)

    ' Use the ActionAdd objects 'test' string as an xpath query to retrieve the matching node
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            XhtmlActions info.SubActions
        Next
    End If

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' XhtmlActionAdd

'=======================================================================================================================
' Procedure:    XhtmlActionInsert
' Purpose:      Processes the Insert action when creating the XHTML data source.
' Notes:        Copies editable areas from the Assessment Report to the word html data source document.
'
' On Entry:     info                The ActionInsert object to be actioned.
'=======================================================================================================================
Public Sub XhtmlActionInsert(ByVal info As ActionInsert)
    Const c_proc As String = "modCreateDocument.XhtmlActionInsert"

    Dim dataNode   As MSXML2.IXMLDOMNode
    Dim theQuery   As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' We only need to update the RichText for fields that are marked as editable, otherwise they are unchanged
    With info
        If .Editable Then

            ' Perform the appropriate update type
            Select Case .DataFormat
            Case rdaDataFormatRichText, rdaDataFormatMultiline

                ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
                theQuery = g_counters.UpdatePredicates(.DataSource)

                ' Get the data to update the bookmark with
                Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

                ' Check that we actually retrieved a node
                If Not dataNode Is Nothing Then

                    ' Copy the RichText from the Assessment Report to the Word XHTML document
                    CopyRichTextBlock info
                End If

            Case rdaDataFormatText, rdaDataFormatLong, rdaDataFormatDateLong, rdaDataFormatDateShort, rdaDataFormatTick
                ' There is no need to do anything for these 'actions'

            End Select
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' XhtmlActionInsert

'=======================================================================================================================
' Procedure:    CopyRichTextBlock
' Purpose:      Copies an editable block of text from the Assessment Report to the html file used to generate xhtml.
' Notes:        It is quite possible that if this Insert action is part of an Add action block and the Add action has
'               a 'deleteIfNull' attribute that the whole block has not been generated. As a consequence of that the
'               bookmark for this ActionInsert may not exist.
'
' On Entry:     info                The ActionInsert object containing the source data bookmark name and xpath query.
'=======================================================================================================================
Private Sub CopyRichTextBlock(ByVal info As ActionInsert)
    Const c_proc As String = "modHTMLToXHTML.CopyRichTextBlock"

    Dim queryKey       As String
    Dim Source         As Word.Range
    Dim sourceBookmark As String
    Dim target         As Word.Range

    On Error GoTo Do_Error

    ' Use the Pattern bookmark name if it is specified
    With info
        If LenB(.BookmarkPattern) > 0 Then
            sourceBookmark = ReplacePatternData(.BookmarkPattern, .PatternData)
        Else
            sourceBookmark = .Bookmark
        End If
        sourceBookmark = g_counters.UpdatePredicates(sourceBookmark)

        ' Create the queryKey (the xml node that will be updated using the following rich text plus a new paragraph.
        ' Wrap the query key in an leadin character block and a query end character block so that
        ' we can find the start and end of a complete entry (the query and the RichText block).
        queryKey = mgrHTMLBookmarkedBlockLeadIn & g_counters.UpdatePredicates(.DataSource) & mgrHTMLBookmarkNameEnd & vbCr
    End With

    ' Make sure the source Bookmark actually exists, if it is in a nested Add block it may not
    If g_assessmentReport.bookmarks.Exists(sourceBookmark) Then

        ' Get a reference to the source Range
        Set Source = g_assessmentReport.bookmarks(sourceBookmark).Range

        ' Set a Range object to the xhtml Word document so that we can add text to it
        Set target = m_xhtmlWordDoc.Content

        ' Make sure we add all new text at the very end of the document
        target.Collapse wdCollapseEnd

        ' If the xhtml Word document already contains text add a new paragraph to hold
        ' the query string so that it not contiguous to text already present
        If m_xhtmlWordDoc.Content.End > 1 Then
            target.InsertParagraph
        End If

        ' Add the queryKey (the xml node that will be updated using the following rich text
        target.InsertAfter queryKey
        target.Collapse wdCollapseEnd

        ' Check to see if the source range is using the Default Text (which we do not want to update the xml with)
        If info.HasDefaultText And info.DefaultText = Source.Text Then

            ' The source document is using the Default Text (which means that the user has not updated it).
            ' Since we do not want to propagate the Default Text to the xml replace the Default Text with a null string.
            target.InsertAfter vbNullString
        Else

            ' Make sure there is some text to copy or we get error 4605 "This method or property is not available because no text is selected."
            If LenB(Source.Text) > 0 Then

                ' Copy the source RichText data
                Source.Copy

                ' Paste the RichText into the xhtml Word document
                target.Paste

                ' Now fix up any Bullets in the text we have just pasted into the html document.
                ' We need to replace all paragraphs with Bullet points with a special character sequence as Word does not generate
                ' the correct html '<ul><li>info</li></ul>' sequence. It just converts the bullet to a symbol character and pads
                ' the text with non-breaking space characters. Generating the correct html in the parser is too complex so we
                ' compromise by stripping the bullets from each bulleted paragraph and adding a special character sequence to
                ' enable us to recognise where bullets should be when the text comes back in from Remedy/RDA.
                StripBullets target
            End If
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CopyRichTextBlock

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
        m_xhtmlWordDoc.UndoClear
    End If

'''Debug.Print "Bullet count " & CStr(bulletCount)

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' StripBullets
