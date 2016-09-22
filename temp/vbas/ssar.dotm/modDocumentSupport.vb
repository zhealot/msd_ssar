Attribute VB_Name = "modDocumentSupport"
'===================================================================================================================================
' Module:       modDocumentSupport
' Purpose:      Contains support code used for building and submitting an Assessment Report.
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
' History:      03/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module

' EC = Excluded Content
' This is content that should not be present in an assessment report
Private Const mc_ECInlineShapes         As String = "Inline Shapes"
Private Const mc_ECShapeRange           As String = "Shape Range"
Private Const mc_ECFrames               As String = "Frames"
Private Const mc_ECMaths                As String = "Maths"
Private Const mc_ECRevisions            As String = "Revisions"
Private Const mc_ECSectionBreaks        As String = "Section Breaks"

Private Const mc_TickCharacterFont      As String = "Wingdings"
Private Const mc_TickCharacterNumber    As Long = -3844


'=======================================================================================================================
' Procedure:    PreSubmitVetting
' Purpose:      Check that the document does not contain content we do not want.
' Notes:        We can't stop the user adding things to the document that we don't want, so we check the Assessment
'               Report before submitting it to the Remedy/RDA webservice.
'
' Returns:      True if no invalid content was located.
'=======================================================================================================================
Public Function PreSubmitVetting() As Boolean
    Const c_proc As String = "modDocumentSupport.PreSubmitVetting"

    Dim errorFound  As Boolean

    On Error GoTo Do_Error

    With g_assessmentReport.Content
        If .InlineShapes.Count > 0 Then
            .InlineShapes(1).Select
            DisplayContentWarning mc_ECInlineShapes, c_proc, errorFound
        End If
        If .ShapeRange.Count > 0 Then
            .ShapeRange(1).Select
            DisplayContentWarning mc_ECShapeRange, c_proc, errorFound
        End If
        If .Frames.Count > 0 Then
            .Frames(1).Select
            DisplayContentWarning mc_ECFrames, c_proc, errorFound
        End If
        If .OMaths.Count > 0 Then
            .OMaths(1).Range.Select
            DisplayContentWarning mc_ECMaths, c_proc, errorFound
        End If
        If .Revisions.Count > 0 Then
            .Revisions(1).Range.Select
            DisplayContentWarning mc_ECRevisions, c_proc, errorFound
        End If
        If .Sections.Count > 1 Then
            DisplayContentWarning mc_ECSectionBreaks, c_proc, errorFound
        End If
    End With

    ' Return whether we encountered problems or not
    PreSubmitVetting = Not errorFound

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' PreSubmitVetting

Private Sub DisplayContentWarning(ByVal errorType As String, _
                                  ByVal caller As String, _
                                  ByRef errorFound As Boolean)
    Dim errorText   As String
                                
    ' This is more serious than a warning but we don't want to actually Raise an error for this
    errorText = Replace$(mgrErrTextInvalidDocumentContent, mgrP1, errorType)
    EventLog errorText, caller
    MsgBox errorText, vbOKOnly, ssarTitle

    errorFound = True
End Sub ' DisplayContentWarning

'===================================================================================================================================
' Procedure:    ApplyColourMap
' Purpose:      Applies the Colour Map (foreground and background colours) based on the value of either the node in the Custom XML
'               Part Content Control data store or the value of the text in the passed in range object.
' Date:         20/06/16    Created.
'
' On Entry:     theRange            The range to which the Colour Map should be applied, this range also supplied the Colour Map
'                                   lookup value if ccDataNode is a null string.
'               ccDataNode          The xml node in the Custom XML Part Content Control data store, that supplies the value to be
'                                   used as the Colour Map lookup key.
'===================================================================================================================================
Public Sub ApplyColourMap(ByVal theRange As Word.Range, _
                          Optional ByVal ccDataNode As String)
    Const c_proc As String = "modDocumentSupport.ApplyColourMap"

    Dim backgroundColour    As Long
    Dim colourLookupValue   As String
    Dim textColour          As Long

    On Error GoTo Do_Error

    ' If there is no value for CCDataNode use the value of the passed in range object
    If LenB(ccDataNode) > 0 Then

        ' Replace any predicate placeholders with their real values
        ccDataNode = g_counters.UpdatePredicates(ccDataNode)

        ' Get the Content Controls current value and use that to look up the required text and background colours
        colourLookupValue = g_ccXMLDataStore.Value(ccDataNode)
    Else

        ' If the passed in range contains a Content Control, use the Content Controls range instead as by convention we place a
        ' space character either side of a DropDown Content Control. These spaces will be returned as part of the Bookmarks range
        ' causing the Colour Map lookup to fail.
        If theRange.ContentControls.Count > 0 Then
            Set theRange = theRange.ContentControls(1).Range
        End If

        ' Use the value of the text specified by the range object (which in turn should be a value from rda)
        colourLookupValue = theRange.Text
    End If

    ' Make sure that the Content Controls currently displayed value exists in the colour lookup tables
    If g_colourMapForeground.Exists(colourLookupValue) Then
        textColour = g_colourMapForeground(colourLookupValue)
        backgroundColour = g_colourMapBackground(colourLookupValue)
    Else

        ' The lookup value is not in the table so use Words default values
        textColour = wdColorAutomatic
        backgroundColour = wdColorAutomatic
    End If

    ' Set the font colour
    theRange.Font.Color = textColour

    ' We need to set the background colour differently if the range is in a table
    If theRange.Tables.Count > 0 Then

        ' This sets the backround colour of the cell
        theRange.Shading.BackgroundPatternColor = backgroundColour
    Else

        ' This sets the background colour of the paragraph
        theRange.ParagraphFormat.Shading.BackgroundPatternColor = backgroundColour
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ApplyColourMap

'=======================================================================================================================
' Procedure:    AddStringVariantToRTDictionary
' Purpose:      Adds a string (as a variant) to the Rich Text Dictionary.
' Note:         We convert the data from a string to a variant so that it can be added to the dictionary.
'
' On Entry:     theKey              The key used to add the string (variant) to the Dictionary.
'               theString           The data to be added to the Dictionary.
'=======================================================================================================================
Public Sub AddStringVariantToRTDictionary(ByVal theKey As String, _
                                           ByVal theString As Variant)
    '  Add the string (as a variant) to the Rich Text Dictionary object
    g_richTextData.Add theKey, theString
End Sub ' AddStringVariantToRTDictionary

'=======================================================================================================================
' Procedure:    DeleteUnusedTextBlock
' Purpose:      Deletes the bookmark and and text it marks.
'
' On Entry:     bookmarkToDelete    The name of the bookmark whose range is to be deleted.
'=======================================================================================================================
Public Sub DeleteUnusedTextBlock(ByVal bookmarkToDelete As String)
    Const c_proc As String = "modDocumentSupport.DeleteUnusedTextBlock"

    On Error GoTo Do_Error

    If LenB(bookmarkToDelete) Then

        ' Replace any numeric placeholders in the bookmark name with their real values
        bookmarkToDelete = g_counters.UpdatePredicates(bookmarkToDelete)

        ' Delete the bookmark and any text and nested bookmarks it may contain
        g_assessmentReport.bookmarks(bookmarkToDelete).Range.Delete
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DeleteUnusedTextBlock

'=======================================================================================================================
' Procedure:    InsertBuildingBlock
' Purpose:      Inserts the specified Building Block into the Assessment Report document.
'
' On Entry:     bbName              The name of the Building Block to insert.
'               bbWhere             Where relative to the Range specified by bookmarkName the Building Block will be
'                                   inserted.
'               bookmarkName        The name of the Bookmark that identifies the Range where the Building block will be
'                                   inserted.
'               BookmarkPattern     A parameterised bookmark name used to generate a unique name for the Range where the
'                                   Building Block was inserted.
'               insertPosition      ???.
'               bookmarkExtend      Whether the Range specified by bookmarkName should be extended to include the
'                                   Building Block being inserted.
'=======================================================================================================================
Public Sub InsertBuildingBlock(ByVal bbName As String, _
                               ByVal bbWhere As ssarWhereType, _
                               ByVal bookmarkName As String, _
                               ByVal bookmarkPattern As String, _
                               ByVal insertPosition As Long, _
                               ByVal bookmarkExtend As Boolean)
    Const c_proc As String = "modDocumentSupport.InsertBuildingBlock"

    Dim bookmarkError       As Boolean
    Dim buildingBlockError  As Boolean
    Dim errorText           As String
    Dim originalBMEnd       As Long
    Dim originalBMStart     As Long
    Dim patternRange        As Word.Range
    Dim target              As Word.Range
    Dim targetCopy          As Word.Range
    Dim useTemplate         As Word.Template
    Dim whereInserted       As Word.Range

    On Error GoTo Do_Error

    ' Add the choosen Building Block to the document
    If LenB(bbName) > 0 Then

        ' This is the template that contains all Building Blocks we insert in the Assessment Report
        Set useTemplate = g_assessmentReport.AttachedTemplate

        ' Replace any numeric placeholders in the Bookmark name with their corresponding values
        bookmarkName = g_counters.UpdatePredicates(bookmarkName)

        ' Location to insert the Building Block
        bookmarkError = True
        Set target = g_assessmentReport.bookmarks(bookmarkName).Range
        bookmarkError = False
        originalBMStart = target.Start
        originalBMEnd = target.End

        ' Check to see if we need to insert a Paragraph or NewLine before we insert the Building Block
        Select Case bbWhere
        Case ssarWhereTypeAfterLastParagraph

            ' Start by collapsing the range to the end and then adding either a new paragraph or new line
            target.MoveEndUntil vbCr, 1

            ' Bookmark protection code: This prevents bookmarks in following paragraphs from having their start position extended
            ' to include the Building Block we are inserting. We add a paragraph at the end of the target range and then insert
            ' the Building Block before the added paragraph mark. After we add the Building block we remove the unwanted paragraph.
            If insertPosition > 1 Then
                target.End = target.End - 1
                target.InsertParagraphAfter
            End If

            target.Collapse wdCollapseEnd
            target.Select

        Case ssarWhereTypeAtEndOfRange

            target.Collapse wdCollapseEnd

        Case ssarWhereTypeReplaceRange
            
            ' Leave the target range unchanged as we want to replace it
        End Select

        EventLog "Inserting Building Block: " & bbName

        ' Preserve the original bookmark pattern Range End position if possible
        If LenB(bookmarkPattern) > 0 Then
    
            ' Generate the real bookmark name from the pattern
            bookmarkPattern = g_counters.UpdatePredicates(bookmarkPattern)

            ' If the bookmark exists, create a Range object for it, so that we can recreate the bookmark
            ' since updating the bookmark text will destroy or alter the pattern bookmark
            If g_assessmentReport.bookmarks.Exists(bookmarkPattern) Then
                Set patternRange = g_assessmentReport.bookmarks(bookmarkPattern).Range.Duplicate
            End If
        End If

        ' Insert the specified Building Block from the Assessment Report template
        buildingBlockError = True
        Set whereInserted = useTemplate.BuildingBlockEntries(bbName).Insert(target, True)
        buildingBlockError = False

        ' Bookmark protection code: Remove the paragraph mark we added above as it has served its purpose and is no longer required
        If bbWhere = ssarWhereTypeAfterLastParagraph And insertPosition > 1 Then
            Set target = whereInserted
            target.MoveEndWhile vbCr, 1
            target.Start = target.End - 1
            target.Delete
        End If

        ' Recreate the bookmark name as a collapsed range at the end of the inserted block of text.
        ' If Sections are inserted contiguously this bookmark is used for the next insertion point.
        ' If the inserted Section is part of an outer Section, the bookmark will be replaced the
        ' next time another outer Section is inserted as the bookmark will be defined in the outer
        ' Section.
        If bbWhere <> ssarWhereTypeReplaceRange Then
            Set target = g_assessmentReport.bookmarks(bookmarkName).Range
            target.End = whereInserted.End

            ' The choice is to either extend the bookmark or to redefine it to what it originally was
            If bookmarkExtend Then
                target.Start = originalBMStart
                g_assessmentReport.bookmarks.Add bookmarkName, target
            Else

                Set targetCopy = target.Duplicate
                targetCopy.Start = originalBMStart
                targetCopy.End = originalBMEnd
                g_assessmentReport.bookmarks.Add bookmarkName, targetCopy
            End If
        End If

        ' See if a second bookmark should be create using a name derived from the bookmarkPattern
        If LenB(bookmarkPattern) > 0 Then

            ' Recreate the pattern bookmark based on whether it existed when this procedure was called.
            ' Preserve the original bookmark pattern Range End position if possible.
            If patternRange Is Nothing Then

                ' It did not exist before this procedure was called
                g_assessmentReport.bookmarks.Add bookmarkPattern, whereInserted
            Else

                ' It did exist, so recreate it with the expanded range
                g_assessmentReport.bookmarks.Add bookmarkPattern, patternRange
            End If
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    If bookmarkError Then
        If Err.Number = mgrErrNoRequestedCollectionMemberDoesNotExist Then
            Err.Description = Replace$(mgrErrTextBookmarkDoesNotExist, mgrP1, bookmarkName)
        End If
    ElseIf buildingBlockError Then
        If Err.Number = mgrErrNoRequestedCollectionMemberDoesNotExist Then
            Err.Description = Replace$(mgrErrTextBuildingBlockDoesNotExist, mgrP1, bbName)
        End If
    End If

    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' InsertBuildingBlock

Public Function InsertBuildingBlock2(ByVal bbName As String, _
                                     ByVal whereToInsert As Word.Range) As Word.Range
    Const c_proc As String = "modDocumentSupport.InsertBuildingBlock2"

    Dim useTemplate         As Word.Template

    On Error GoTo Do_Error

    ' This is the template that contains all Building Blocks we insert in the Assessment Report
    Set useTemplate = g_assessmentReport.AttachedTemplate

    ' Insert the Building Block
    Set InsertBuildingBlock2 = useTemplate.BuildingBlockEntries(bbName).Insert(whereToInsert, True)

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' InsertBuildingBlock2

'=======================================================================================================================
' Procedure:    FillBookmarkRichText
' Purpose:      Fills the specified bookmark with rich text from the rich text data source.
'
' On Entry:     info                The ActionInsert object for which RIchText will be inserted.
'=======================================================================================================================
Public Function FillBookmarkRichText(ByVal info As ActionInsert) As Word.Range
    Const c_proc As String = "modDocumentSupport.FillBookmarkRichText"

    Dim keyBookmark As String
    Dim rawName     As String
    Dim targetArea  As Word.Range
    Dim theBookmark As String
    Dim theHTMLText As Word.Range
    Dim theText     As String

    On Error GoTo Do_Error

    ' The key to retrieving the RichText is the g_richTextData dictionary object. Retrieval of RichText is optimised by storing
    ' RichText that contains only plain text in a Variant, obviating writing to the html file and the copy/paste operation.

    ' Generate the bookmark we need to retrieve either the Range object (which we will use to retrieve text from the Word html
    ' document) or the Variant which contains the actual text we will use
    rawName = ReplacePatternData(info.bookmarkPattern, info.PatternData)
    keyBookmark = g_counters.UpdatePredicates(rawName)

    ' This is the name of the bookmark to be updated in the Assessment Report
    theBookmark = info.Bookmark

    ' Now either copy/paste the data or just update the bookmark with text
    Select Case TypeName(g_richTextData.Item(keyBookmark))
    Case "Range"

        ' Get the the Range object from the dictionary object
        Set theHTMLText = g_richTextData.Item(keyBookmark)

        ' Copy the block marked by the Range object
        theHTMLText.Copy

        ' Set the range object to the area we will paste the data into
        Set targetArea = g_assessmentReport.bookmarks(theBookmark).Range

        ' Paste the block of text into the rda Word document. This preserves and formatting from the original object.
        targetArea.Paste
        
        ' Fix up bullet points in the last paragraph.
        ' When RichText is copied and pasted into the Assessment Report, the final paragraph mark is not copied from the
        ' html file. So if the final paragraph is bulleted we miss this, because we did not copy the paragraph mark. The
        ' reason the paragraph mark is not copied is to preserve the layout of the Assessment Report as copying it would
        ' override any formatting assoiacated with the paragraph being pasted into in the Assessment Report.
        
        ' #TRY# This code predates the stripping of bullets and replacing them with the mgrBulletRecognitionSequence
        ' #TRY# character sequence at the beginging of the paragraph. So we should be able to skip this.
'''        FixupBullets theHTMLText, targetArea, keyBookmark

        ' Create a new bookmark using the keyBookmark name as this will be requried to
        ' retrieve theRichText when it is passed back to the Remedy/RDA webservice
        g_assessmentReport.bookmarks.Add keyBookmark, targetArea

        ' Recreate the bookmark because the Paste action will have removed it and it may possibly be needed
        g_assessmentReport.bookmarks.Add theBookmark, targetArea

        ' Make a best effort endeavour to sort out over width table (tables that extend past the right margin)
        TableFixerUpper targetArea

        RestoreBullets targetArea

        ' Return Range object for the area we pasted the RichText into
        Set FillBookmarkRichText = targetArea

    Case "String", "Empty"

        ' Get the value that we will use to update the bookmark
        theText = g_richTextData.Item(keyBookmark)

        ' If the current value is a null string use the Default Text, the worse
        ' we can do is overwrite one null string with another null string
        If Len(theText) = 0 Then
            theText = info.DefaultText
        End If

        ' Update the bookmark in the Assessment Report using the string data in the Variant stored in the dictionary object
        Set FillBookmarkRichText = UpdateBookmark(theBookmark, keyBookmark, info.PatternData, info.DeleteIfNull, theText, True)
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' FillBookmarkRichText

'=======================================================================================================================
' Procedure:    ReplacePatternData
' Purpose:      Adds data retrieved from the assessment report xml by an xpath query as the replacement data in the
'               passed in inputString.
' Notes:        The location where the replacement data should be inserted is defined by ssarPDP1.
'
' On Entry:     inputString         The string containing the bookmark name and pattern placeholder text to be updated.
'               xpathQuery          The xpath query supplying the data for the replaceble parameter.
'=======================================================================================================================
Public Function ReplacePatternData(ByVal inputString As String, _
                                   ByVal xpathQuery As String) As String
    Const c_proc As String = "modDocumentSupport.ReplacePatternData"

    Dim dataNode    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Assign inputString as the default return value as this code more often than not exits without doing anything
    ReplacePatternData = inputString

    If LenB(xpathQuery) > 0 Then
        If LenB(inputString) > 0 Then

            ' Get the data to update the bookmark with
            xpathQuery = g_counters.UpdatePredicates(xpathQuery)
            Set dataNode = g_xmlDocument.SelectSingleNode(xpathQuery)

            ' Use a node from the assessment report with a unique value to generate a unique but relevant bookmark name
            If Not dataNode Is Nothing Then
                ReplacePatternData = Replace$(inputString, ssarPDP1, dataNode.Text)
            End If
        End If
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ReplacePatternData

'=======================================================================================================================
' Procedure:    SetAsEditableRange
' Purpose:      Sets an area as being editable, only input areas are marked as editable.
' Notes 1:      This is necessary as the Assessment Report is Read Only protected (wdAllowOnlyReading). If we don't do
'               this the user is unable to edit the Assessment Report.
' Notes 2:      We also add the Editor object that allows the Range to be edited to a Collection object so that we can
'               recreate the Bookmark associated with the editable range. If the user copies and pastes an entire input
'               area into another complete input area it destroys the Bookmark of the area being pasted into. This is
'               contrary to what Word normal does, where you would expect to see the source Bookmark removed and
'               recreated at the paste location.
'
' On Entry:     targetArea          The area that should be authorised as being editable.
'               bookmarkName        The name of the Bookmark to associate with the editable area.
'=======================================================================================================================
Public Sub SetAsEditableRange(ByVal targetArea As Word.Range, _
                              ByVal bookmarkName As String)
    Const c_proc As String = "modDocumentSupport.SetAsEditableRange"

    On Error GoTo Do_Error

    ' Set the Range as being editable by everyone
    If Not targetArea Is Nothing Then

        targetArea.Editors.Add wdEditorEveryone

        ' Add the Bookmark name to the Collection object in the order the Bookmark names occur.
        ' We reuse the Bookmark name as the key so that we can insert new items relative to the key.
        g_editableBookmarks.BookmarkAdd bookmarkName
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' SetAsEditableRange

''''=======================================================================================================================
'''' Procedure:    FixupBullets
'''' Purpose:      .
'''' Notes:        .
''''
'''' On Entry:     sourceHTMLText      .
''''               targetArea          .
''''=======================================================================================================================
'''Private Sub FixupBullets(ByVal sourceHTMLText As Word.Range, _
'''                         ByVal targetArea As Word.Range, _
'''                         ByVal bm As String)
'''    Const c_proc As String = "modDocumentSupport.FixupBullets"
'''
'''    Dim finalParagraphSource    As Word.Range
'''    Dim finalParagraphTarget    As Word.Range
'''    Dim paragraphCount          As Long
'''    Dim listTemplateSource      As Word.ListTemplate
'''
'''    On Error GoTo Do_Error
'''
'''    ' Get the final paragraph of the html derived text
'''    paragraphCount = sourceHTMLText.Paragraphs.Count
'''    Set finalParagraphSource = sourceHTMLText.Paragraphs(paragraphCount).Range
'''
'''    ' See if the final paragraph of the source html derived text has a bullet
'''    If finalParagraphSource.ListFormat.CountNumberedItems > 0 Then
'''
'''        ' Get the last paragraph of the target text
'''        paragraphCount = targetArea.Paragraphs.Count
'''        Set finalParagraphTarget = targetArea.Paragraphs(paragraphCount).Range
'''
'''        ' Get the source final paragraphs ListTemplate object (which contains the Bullet info)
'''        Set listTemplateSource = finalParagraphSource.ListFormat.ListTemplate
'''
'''        ' Now apply the ListTemplate to the final target paragraph to apply the html sources Bullet
'''        finalParagraphTarget.ListFormat.ApplyListTemplate listTemplateSource, True, wdListApplyToThisPointForward
'''    End If
'''
'''Do_Exit:
'''    Exit Sub
'''
'''Do_Error:
'''    ErrorReporter c_proc
'''    Resume Do_Exit
'''End Sub ' FixupBullets

'===================================================================================================================================
' Procedure:    RestoreBullets
' Purpose:      When building an assessment report we restore bullets to paragraphs that are supposed to have bullets.
'
' Note 1:       This explanation starts at the end and works backwards:
'               When we return data to the remedy RDA system, we have to save the editable areas of the assessment report to another
'               Word document, which in turn is saved as a Filtered HTML file. However, word does not generate the correct HTML
'               tags. Instead of generating <ul><li>abc</li><li>def</li></ul> structures, Word just adds a character (for the
'               bullet) followed by a number of spaces for padding. This makes the bullet impossible to reconstitute. So we strip
'               the bullets from bulleted paragraphs and add a unique character sequence to the start of the paragraph before
'               returning the data to RDA. This code looks for the special character sequence at the start of a paragraph. If it is
'               found, the special character sequence is stripped out and a default bullet is applied to that paragraph,
'               reconstituting a bullet (but it may not be the original bullet character selected by the user as we use default
'               bullets).
'
' Date:         26/08/16    Rewrite to use styles where possible.
'
' On Entry:     areaOfInterest      A Range in the assessment report to replace character placeholder sequences with real bullets.
'===================================================================================================================================
Public Sub RestoreBullets(ByVal areaOfInterest As Word.Range)
    Const c_proc As String = "modDocumentSupport.RestoreBullets"

    Dim finalParagraph      As Word.Range
    Dim index               As Long
    Dim indexParagraph      As Word.Paragraph
    Dim bulletPrefixLength  As Long
    Dim prefixText          As String
    Dim paragraphCount      As Long
    Dim paragraphRange      As Word.Range
    Dim useStyle            As String

    On Error GoTo Do_Error

    ' Short circuit the paragraph by paragraph search by searching the entire area for a Bullet sequence
    If InStr(areaOfInterest.Text, mgrBulletRecognitionSequence) > 0 Then

        ' Now iterate each paragraph in the editable area looking for bulleted paragraphs
        bulletPrefixLength = Len(mgrBulletRecognitionSequence)

        ' If the final paragraph is already bulleted (the paragraph is already bulleted in the template),
        ' get its Style for use with other bulleted paragraphs
        paragraphCount = areaOfInterest.Paragraphs.Count
        Set finalParagraph = areaOfInterest.Paragraphs(paragraphCount).Range
        If finalParagraph.ListParagraphs.Count > 0 Then
            useStyle = finalParagraph.Style
        End If

        For index = 1 To paragraphCount
            Set paragraphRange = areaOfInterest.Paragraphs(index).Range

            ' We are only interested if the paragraph has a bullet point
            prefixText = Left$(paragraphRange.Text, bulletPrefixLength)
            If prefixText = mgrBulletRecognitionSequence Then

                ' Strip the special character sequence used to indicate that this paragraph should have a bullet
                paragraphRange.End = paragraphRange.Start + bulletPrefixLength
                paragraphRange.Delete

'''                ' See if we should use Style based bullets or the default bullet
'''                If LenB(useStyle) > 0 Then

'''                    ' This paragraph is part of a list, where the final paragraph has been inserted
'''                    ' into a paragraph that is formatted with a Style that uses a bullet
'''                    paragraphRange.Style = g_assessmentReport.Styles(useStyle)
'''                Else

                    ' Now apply the default bullet
                    paragraphRange.ListFormat.ApplyBulletDefault
'''                End If
            End If
        Next

        ' Make sure not too many undo operation accumulate on the undo stack
        g_assessmentReport.UndoClear
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RestoreBullets

'=======================================================================================================================
' Procedure:    TableFixerUpper
' Purpose:      This code does its best to detect and correct tables that extend past the right margin.
' Notes:        The calling code expects no changes to the Table content only the Table meta data, so the passed in
'               Range object (Start and End) is not altered.
'
' On Entry:     areaOfInterest      The Range indicating the area to be searched for Tables.
'=======================================================================================================================
Private Sub TableFixerUpper(ByVal areaOfInterest As Word.Range)
    Const c_proc As String = "modDocumentSupport.TableFixerUpper"

    Dim childTable  As Word.Table
    Dim mainTable   As Word.Table
    Dim thePS       As Word.PageSetup
    Dim usableWidth As Single

    On Error GoTo Do_Error

' #TRY# See if tables still need adjusting, this may lead to problems with tables the users insert
' #TRY# rather than tables that are actually part of the template???
Exit Sub

    With areaOfInterest

        ' There must be at least one Table in the Range or there is nothing to do
        If .Tables.Count > 0 Then

            ' Iterate all Tables defined eith the specified Range
            For Each mainTable In .Tables

                ' Different actions are required based on the tables PreferredWidthType
                Select Case mainTable.PreferredWidthType
                Case wdPreferredWidthAuto
                Case wdPreferredWidthPercent

                    ' Setting a tables PreferredWidth to more than 100% forces the
                    ' table past the right margin (which we definite don't want)
                    If mainTable.PreferredWidth > 100 Then
                        Debug.Print "Correcting PreferredWidth percentage"
                        mainTable.PreferredWidth = 100
                    End If
                Case wdPreferredWidthPoints

                    ' TODO: Unsure at this time if nested tables need to be handled differently than unnested tables
                    If mainTable.NestingLevel > 1 Then
                        
                    Else

                        ' Get a reference to the documents PageSetup object to we can determine the usable page width
                        Set thePS = .Document.PageSetup
                        usableWidth = thePS.PageWidth - thePS.LeftMargin - thePS.RightMargin

                        ' Rather than set the PreferredWidth to the usable width set AutoFitBehavior
                        ' to wdAutoFitContent as this generally yields better results
                        If mainTable.PreferredWidth > usableWidth Then
                            Debug.Print "Correcting PreferredWidthPoints"
                            mainTable.AutoFitBehavior wdAutoFitContent
                        End If
                    End If
                End Select

                ' Now deal to any child (nested) Tables
                For Each childTable In mainTable.Tables
                    TableFixerUpper childTable.Range
                Next
            Next
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' TableFixerUpper

Public Sub UpdateAllRefFields()
    Const c_proc As String = "modDocumentSupport.UpdateAllRefFields"

    Dim arField  As Word.Field

    On Error GoTo Do_Error

    EventLog c_proc

    ' Iterate all fields in the Assessment Report, we check the type so we only update Ref Fields
    For Each arField In g_assessmentReport.Content.Fields
        If arField.Type = wdFieldRef Then

            ' We do not need to unprotect the document for this to work
            With arField
                .Locked = False
                .Update
                .Locked = True

                ' FWB: Delete the unwanted Editor object added by Word to the Field when it was updated!
                If .result.Editors.Count > 0 Then
                    .result.Editors(1).Delete
                End If
            End With
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' UpdateAllRefFields

'=======================================================================================================================
' Procedure:    UpdateBookmark
' Purpose:      Updated the specified bookmark with the passed in value.
' Notes:        The bookmark name is recreated (as updating a bookmark deletes the bookmark) and creates a secondary
'               bookmark as well is required.
'
' On Entry:     bookmarkName        The bookmark to update.
'               patternBookmark     Secondary bookmark name assigned to the original 'bookmarkName' location.
'               patternQuery        .
'               deletableBookmark   A bookmark name that species a range that should be deleted if 'newValue' is a null
'                                   string.
'               newValue            The value to update the location specied by BookmarkName with.
' Returns:      The Range object of the updated bookmark.
'=======================================================================================================================
Public Function UpdateBookmark(ByVal bookmarkName As String, _
                               ByVal patternBookmark As String, _
                               ByVal patternQuery As String, _
                               ByVal deletableBookmark As String, _
                               ByVal newValue As String, _
                               Optional ByVal outputNullString As Boolean) As Word.Range
    Const c_proc As String = "modDocumentSupport.UpdateBookmark"

    Dim doUpdate    As Boolean
    Dim rawName     As String
    Dim targetArea  As Word.Range

    On Error GoTo Do_Error

    ' Replace any numeric parameters in the pattern Bookmark with their corresponding values
    patternBookmark = ReplacePatternData(patternBookmark, patternQuery)
    patternBookmark = g_counters.UpdatePredicates(patternBookmark)

    ' Get the bookmarks Range object so that we can recreate the bookmark
    With g_assessmentReport.bookmarks
        If .Exists(bookmarkName) Then
            Set targetArea = .Item(bookmarkName).Range
        Else

            ' The bookmark does not exist so there is nothing more to do
            Exit Function
        End If

        ' Check to see that there is data to update the specified bookmark with
        doUpdate = outputNullString Or LenB(newValue) > 0
        If doUpdate Then

            ' Replace the bookmarked text (which also deletes the bookmark)
            targetArea.Text = newValue

            ' Recreate the bookmark
            Set UpdateBookmark = .Add(bookmarkName, targetArea).Range

            ' If a patternBookmark name has been supplied, create a second bookmark using that name
            If LenB(patternBookmark) Then
                .Add patternBookmark, targetArea
            End If
        Else

            ' There is no text, so check to see if there is a bookmark that should be deleted. The
            ' bookmark may in turn contain other bookmarks. This provides a mechanism for deleting
            ' corresponding boiler plate text when there is no data for associated bookmarks.
            If LenB(deletableBookmark) Then

                ' Delete the bookmark and any text and nested bookmarks it may contain
                .Item(deletableBookmark).Range.Delete
            Else

                ' Create the secondary 'pattern' bookmark as it is used to return RichText data to the Remedy/RDA webservice
                If LenB(patternBookmark) Then
                    .Add patternBookmark, targetArea
                End If
            End If
        End If
    End With

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' UpdateBookmark
