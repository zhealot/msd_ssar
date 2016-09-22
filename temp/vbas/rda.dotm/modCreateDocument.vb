Attribute VB_Name = "modCreateDocument"
'===================================================================================================================================
' Module:       modCreateDocument
' Purpose:      Creates the actual Assessment Report.
'
'               This module performs the following activities:
'               1. Flushes the current Fluent UI custom RDA Tab setup.
'               2. Loads and parses the 'information.xml' file (specified by the config file) into an internal data model
'                  accessed through the global variable g_instructions.
'               3. Sets up the counters object used for numeric parameter replacement in xpath queries and Bookmark names.
'               4. Creates the HTML data source text file by parsing the 'fetchdoc.infopathxml' file for RichText:
'                   a.  To minimise the data written to and copy/pasted from the HTML file an intermediary Dictionary object is used
'                       as an index. RichText data that contains no formatting is stored in the dictionary as a String assigned to a
'                       Variant. True RichText is stored in the dictionary as a Range object that references the HTML file.
'                   b.  Once the HTML text file has been created it is opened as a Word document. It is then parsed and for each
'                       block of RichText a Range object is created and added to the dictionary.
'               5.  The order in which the rep input file is parsed is determined by the list of 'actions' contained in
'                   the 'information.xml' file. In reality we actually use the g_instructions.Actions.actionList object.
'               6.  The 'action' list is made up of four possible 'actions':
'                   a)  Add     -   Adds one of two Building Blocks to the Assessment Report at the specified Bookmark location.
'                                   A Bookmarked area can be deleted if there is no data for the Building Block being added.
'                   b)  Insert  -   Inserts data into the Assessment Report from either the xml loaded from the
'                                   'fetchdoc.infopathxml' file or the HTML document if the dataFormat is RichText or Multiline.
'                   c)  Rename  -   Renames a Bookmark.
'                   d)  Setup   -   Adds a Building Block to the Assessment Report if it does not already exist.
'               7.  Once the 'instructions' file is fully parsed and the HTML data source created and opened as a Word document, the
'                   construction of the Assessment Report starts. The 'actions' list is iterated, Adding Building Blocks, updating
'                   Bookmarks and Renaming Bookmarks as necessary.
'               8.  The document is password protected and editable ranges enabled, so that the user can only edit text in those
'                   areas marked as being editable.
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
' History:      30/05/16    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit
Option Private Module

' EC = Excluded Content
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
    Const c_proc As String = "modCreateDocument.PreSubmitVetting"

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
    MsgBox errorText, vbOKOnly, rdaTitle

    errorFound = True
End Sub ' DisplayContentWarning

'=======================================================================================================================
' Procedure:    RemoveWatermark
' Purpose:      Removes the Watermark from the Assessment Report.
'=======================================================================================================================
Public Sub RemoveWatermark()
    Const c_proc As String = "modCreateDocument.RemoveWatermark"

    Dim doDocumentProtection As DocumentProtection
    Dim theWatermark         As Word.ShapeRange

    On Error GoTo Do_Error

    ' Make sure there is a watermark before we try to delete it
    If HasWatermark Then

        ' Unprotect the Assessment Report (if it is protected) so that we can delete the watermark.
        ' On termination the instantiated class object will reprotect the document for us.
        Set doDocumentProtection = NewDocumentProtection
        doDocumentProtection.DisableProtection True

        ' Get a reference to the watermark
        Set theWatermark = g_assessmentReport.Sections(1).Headers(wdHeaderFooterPrimary).Range.ShapeRange

        ' Now delete the watermark
        theWatermark.Delete
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
' Procedure:    BARActions
' Purpose:      Main processing loop for Building the Assessment Report (BAR).
' Notes:        This procedure is indirectly recursive.
'
' On Entry:     allActions          A Collection object that contains the list of 'actions' to be carried out.
'=======================================================================================================================
Public Sub BARActions(ByVal allActions As VBA.Collection)
    Const c_proc As String = "modCreateDocument.BARActions"

    Dim errorText As String
    Dim theAction As Object

    On Error GoTo Do_Error

    ' Check that the collection exists before using it
    If allActions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all ActionSetup, ActionAdd and ActionInsert objects in the collection
    For Each theAction In allActions

        ' There are four types of objects in this collection so choose the method appropriate to the object
        Select Case TypeName(theAction)
        Case rdaTypeActionAdd
            BARActionAdd theAction

        Case rdaTypeActionInsert
            BARActionInsert theAction

        Case rdaTypeActionLink
            BARActionLink theAction

        Case rdaTypeActionRename
            BARActionRename theAction

        Case rdaTypeActionSetup
            BARActionSetup theAction

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
End Sub ' BARActions

'=======================================================================================================================
' Procedure:    BARActionAdd
' Purpose:      Inserts Building Blocks while Building the Assessment Report (BAR).
' Notes:        This procedure is indirectly recursive..
'
' On Entry:     info                The ActionAdd object to be actioned.
'=======================================================================================================================
Public Sub BARActionAdd(ByVal info As ActionAdd)
    Const c_proc As String = "modCreateDocument.BARActionAdd"

    Dim bbName         As String
    Dim bbWhere        As rdaWhere
    Dim bookmarkExtend As Boolean
    Dim bookmarkName   As String
    Dim index          As Long
    Dim theQuery       As String
    Dim subNode        As MSXML2.IXMLDOMNode
    Dim subNodes       As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(info.Test)

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' Use the ActionAdd objects 'test' string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' The section only needs to be inserted if it contains data. So we follow
            ' all paths from this node to see if any of them contain data
            Set subNode = subNodes(index - 1)
            If DoesNodePathHaveData(subNode) Then

                ' If this Section uses more than one Building Block, choose the Building Block
                ' based on this sections counter. This allows a different Building Block to be
                ' used the first time a Building Block is added to the document.
                If index = 1 Then
                    bbName = info.BuildingBlock
                    bbWhere = info.Where
                    bookmarkName = info.Bookmark
                    bookmarkExtend = info.ExtendBookmark
                Else
                    bbName = info.BuildingBlockN
                    bbWhere = info.WhereN
                    bookmarkName = info.BookmarkN
                    bookmarkExtend = info.ExtendBookmarkN
                End If

                ' Insert the selected Building Block
                InsertBuildingBlock bbName, bbWhere, bookmarkName, info.Pattern, index, bookmarkExtend

                ' Perform any nested actions in the current ActionAdd object - but not when we delete the block!
                BARActions info.SubActions
            Else

                ' Delete the unused text block if it is present
                DeleteUnusedTextBlock info.DeleteIfNullBookmark
            End If
        Next
    Else

        ' No nodes were returned by the query, so check to see if there is a bookmark that should be
        ' deleted. The bookmark may in turn contain other bookmarks. This provides a mechanism for
        ' deleting corresponding boiler plate text when there is no data for associated bookmarks.

        ' Delete the bookmarked block of text
        DeleteUnusedTextBlock info.DeleteIfNullBookmark
    End If

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BARActionAdd

'=======================================================================================================================
' Procedure:    BARActionInsert
' Purpose:      Inserts data into bookmarks while Building the Assessment Report (BAR).
'
' On Entry:     info                The ActionInsert object to be actioned.
'=======================================================================================================================
Public Sub BARActionInsert(ByVal info As ActionInsert)
    Const c_proc As String = "modCreateDocument.BARActionInsert"

    Dim canEdit       As Boolean
    Dim dataNode      As MSXML2.IXMLDOMNode
    Dim doInsert      As Boolean
    Dim rawName       As String
    Dim targetArea    As Word.Range
    Dim theQuery      As String
    Dim theDataFormat As rdaDataFormat
    Dim theText       As String
    Dim useBookmark   As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
    theQuery = g_counters.UpdatePredicates(info.DataSource)

    ' Get the data to update the bookmark with
    Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

    ' Get this information just the once
    theDataFormat = info.DataFormat
    canEdit = info.Editable

    If dataNode Is Nothing Then
        If canEdit Then
            theText = info.DefaultText
            doInsert = True
        End If
    Else

        ' Get the dataNodes text which will be used to update the bookmark (as long as the DataFormat is not RichText or Multiline)
        If theDataFormat <> rdaDataFormatMultiline And theDataFormat <> rdaDataFormatRichText Then
            theText = dataNode.Text
        End If
        doInsert = True
    End If

    ' Check that we actually retrieved a node
    If doInsert Then
        With info

            ' Perform the appropriate update type
            Select Case theDataFormat
            Case rdaDataFormatText, rdaDataFormatLong

                EventLog "Updating (Long) bookmark: " & .Bookmark

                ' Replace the bookmarked text with the dataNodes value
                Set targetArea = UpdateBookmark(.Bookmark, .BookmarkPattern, .PatternData, .DeleteIfNull, theText, canEdit)

            Case rdaDataFormatRichText, rdaDataFormatMultiline

                EventLog "Updating (RichText) bookmark: " & .Bookmark

                ' Copy the RichText (xhtml) from the Word HTML document and paste it into the assessment report
                Set targetArea = FillBookmarkRichText(info)

            Case rdaDataFormatDateLong

                EventLog "Updating (DateLong) bookmark: " & .Bookmark

                ' Replace the bookmarked text with the dataNodes value
                Set targetArea = UpdateBookmark(.Bookmark, .BookmarkPattern, .PatternData, .DeleteIfNull, _
                                                Format$(theText, mgrDateFormatLong), canEdit)

            Case rdaDataFormatDateShort

                EventLog "Updating (DateShort) bookmark: " & .Bookmark

                ' Replace the bookmarked text with the dataNodes value
                Set targetArea = UpdateBookmark(.Bookmark, .BookmarkPattern, .PatternData, .DeleteIfNull, _
                                                Format$(theText, mgrDateFormatShort), canEdit)

            Case rdaDataFormatTick
            
                EventLog "Updating (Tick) bookmark: " & .Bookmark

                ' Replace the bookmarked text with either a Tick symbol or a null string.
                ' DeleteIfNull is not supported for this data format.
                Set targetArea = TickBookmark(.Bookmark, .BookmarkPattern, .PatternData, CBool(theText))

            End Select

            ' If the range is editable then set it as editable
            If canEdit Then

                ' By preference use the BookmarkPattern which is the secondary Bookmark
                ' name and used where Bookmark names containing a value are incremented
                If LenB(.BookmarkPattern) > 0 Then

                    ' Replace any pattern data parameters (data obtained from the actual Assessment Report) and
                    ' numeric parameters in the pattern Bookmark with their corresponding values
                    rawName = ReplacePatternData(.BookmarkPattern, .PatternData)
                    useBookmark = g_counters.UpdatePredicates(rawName)
                Else
                    useBookmark = .Bookmark
                End If

                ' Furthermore, it must be "Write View" for it to be editable.
                ' The user must not be able to edit the Assessment Report if it is "Read View" or "Print View".
                If g_rootData.IsWritable Then
                    SetAsEditableRange targetArea, useBookmark
                End If
            End If
        End With
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BARActionInsert

'=======================================================================================================================
' Procedure:    BARActionLink
' Purpose:      Sets up the Link Action while Building the Assessment Report (BAR)..
' Notes 1:      The Link Action create a Ref field that references the actual data at a bookmarked location.
' Notes 2:      However, there is a FWB we have to contend with, because the data source (the area referred to by the
'               Bookmark in the Ref Field) has an Editor (it's an editable area of a protected document) for some
'               bizzare reason Word make the Ref Field editable as well!!
' Note 3:       Because we do not want the Ref Fields being editable we have to update them in a very specific manner.
' Note 4:       At the time we insert the Ref the data source Bookmark may not exist, causing the Ref Field to display
'               an error until it is updated when the Assessment Report build process is complete.
'
' On Entry:     info                The ActionLink object to be actioned.
'=======================================================================================================================
Public Sub BARActionLink(ByVal info As ActionLink)
    Const c_proc As String = "modCreateDocument.BARActionLink"

    Dim dataSourceBM    As String
    Dim newBM           As String
    Dim newField        As Word.Field
    Dim rawName         As String
    Dim target          As Word.Range
    Dim targetBM        As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    With g_assessmentReport.bookmarks

        ' Make sure the target bookmark (the location where the Ref Field will be inserted) exists
        targetBM = info.Bookmark
        If .Exists(targetBM) Then

            ' Create the name for pattern Bookmark (used to give a repeating bookmark a unique name)
            newBM = g_counters.UpdatePredicates(info.BookmarkPattern)

            ' Create the name of the data source Bookmark
            rawName = ReplacePatternData(info.Source, info.PatternData)
            dataSourceBM = g_counters.UpdatePredicates(rawName)

            ' Get the target bookmarks Range object so that we can use it to create the new bookmark
            Set target = .Item(targetBM).Range

            ' Insert the Ref Field
'''            Set newField = g_assessmentReport.Content.Fields.Add(target, wdFieldRef, dataSourceBM, True)
            Set newField = g_assessmentReport.Content.Fields.Add(target, wdFieldRef, dataSourceBM, False)

            ' Use the unique new Bookmark name to create a Bookmark for the Ref Field so that we can find it again
            .Add newBM, newField.result

            ' FWB: Delete the unwanted Editor object added by Word to the Field when it was updated!
            If newField.result.Editors.Count > 0 Then
                newField.result.Editors(1).Delete
            End If
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BARActionLink

'=======================================================================================================================
' Procedure:    BARActionRename
' Purpose:      Renames a bookmark.
' Notes:        This is generally used to create a unique bookmark name where one does not currently exist or a unique
'               bookmark name is required by a subsequent action block.
'
' On Entry:     info                The ActionRename object to be actioned.
'=======================================================================================================================
Public Sub BARActionRename(ByVal info As ActionRename)
    Const c_proc As String = "modCreateDocument.BARActionRename"

    Dim bookmarkName    As String
    Dim newName         As String
    Dim target          As Word.Range

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    With g_assessmentReport.bookmarks

        ' Update any numeric parameters in the Old Bookmark name with their corresponding values
        bookmarkName = g_counters.UpdatePredicates(info.OldBookmarkName)

        ' Make sure the old bookmark exists before trying to rename it
        If .Exists(bookmarkName) Then

            ' Generate the name for the new bookmark
            newName = g_counters.UpdatePredicates(info.NewBookmarkName)

            ' Get the old bookmarks Range object so that we can use it to create the new bookmark
            Set target = .Item(bookmarkName).Range

            ' Create a bookmark using the new name
            .Add newName, target

            ' Delete the old bookmark name
            .Item(bookmarkName).Delete
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BARActionRename

'=======================================================================================================================
' Procedure:    BARActionSetup
' Purpose:      Currently just a null procedure as nothing needs to be done.
' Notes:        The Setup Action is used by the userinterface.
'
' On Entry:     info                The ActionSetup object to be actioned.
'=======================================================================================================================
Private Sub BARActionSetup(ByVal info As ActionSetup)
    Const c_proc As String = "modCreateDocument.BARActionSetup"

    On Error GoTo Do_Error
    

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BARActionSetup

'=======================================================================================================================
' Procedure:    RichTextActions
' Purpose:      Main processing loop for creating the Rich Text data source used to build the Assessment Report.
' Notes:        This procedure is indirectly recursive.
'
' On Entry:     allActions          A Collection object that contains the list of 'actions' to be carried out.
'=======================================================================================================================
Public Sub RichTextActions(ByVal allActions As VBA.Collection)
    Const c_proc As String = "modCreateDocument.RichTextActions"

    Dim errorText As String
    Dim theAction As Object

    On Error GoTo Do_Error

    If allActions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all ActionAdd and ActionInsert objects in the collection
    For Each theAction In allActions

        ' There are four object types in this collection so choose the method appropriate to the object
        Select Case TypeName(theAction)
        Case rdaTypeActionAdd
            RichTextAdd theAction

        Case rdaTypeActionInsert
            RichTextInsert theAction

        Case rdaTypeActionLink, rdaTypeActionRename, rdaTypeActionSetup
            ' We do not need to do anything for these actions

        Case Else
            errorText = mgrErrTextUnknownActionVerbType & TypeName(theAction)
            Err.Raise mgrErrNoUnknownActionVerbType, c_proc, errorText
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RichTextActions

'=======================================================================================================================
' Procedure:    RichTextAdd
' Purpose:      .
' Notes:        .
'
' On Entry:     info                The ActionAdd object to be actioned.
'=======================================================================================================================
Private Sub RichTextAdd(ByVal info As ActionAdd)
    Const c_proc As String = "modCreateDocument.RichTextAdd"

    Dim index    As Long
    Dim theQuery As String
    Dim subNodes As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(info.Test)

    ' Use the ActionAdd objects Test string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            RichTextActions info.SubActions
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
End Sub ' RichTextAdd

'=======================================================================================================================
' Procedure:    RichTextInsert
' Purpose:      Creates the Dictionary object and html file, which jointly are the source of all RichText data consumed
'               by the Assessment Report.
'
' On Entry:     info                The ActionInsert object to be actioned.
'=======================================================================================================================
Public Sub RichTextInsert(ByVal info As ActionInsert)
    Const c_proc As String = "modCreateDocument.RichTextInsert"

    Dim bookmarkName  As String
    Dim childData     As String
    Dim childNode     As MSXML2.IXMLDOMNode
    Dim dataNode      As MSXML2.IXMLDOMNode
    Dim divNode       As MSXML2.IXMLDOMNode
    Dim index         As Long
    Dim plainText       As Boolean
    Dim pNode         As MSXML2.IXMLDOMNode
    Dim rawName       As String
    Dim theQuery      As String

    ' Storage of RichText is optimised by storing RichText that contains only plain text in a Variant which
    ' is added to the dictionary. This means that we do not have to write the to thye html file or retrieve
    ' it from the html file, thus improving the overall speed of the Assessment Report generation.

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' We only need to process the query if the current ActionInsert object has a MultiLine or RichText data format
    If info.DataFormat = rdaDataFormatMultiline Or info.DataFormat = rdaDataFormatRichText Then

        ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
        theQuery = g_counters.UpdatePredicates(info.DataSource)

        ' Get the RichText node, it may contain rich text or plain text
        Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

        ' Create a unique bookmark name for this data node
        rawName = ReplacePatternData(info.BookmarkPattern, info.PatternData)
        bookmarkName = g_counters.UpdatePredicates(rawName)

        ' Check that we actually retrieved a node
        If Not dataNode Is Nothing Then


            If dataNode.ChildNodes.Length = 0 Then
                plainText = True
            ElseIf dataNode.FirstChild.NodeType = NODE_TEXT Then
                plainText = True
            Else
                plainText = False
            End If

            ' If there is a 'div' node then retrieve the nodes xml
            If plainText Then

                ' Add the variant that contains the string to the dictionary object using generated bookmark name as the key
                AddStringVariantToRTDictionary bookmarkName, dataNode.Text
            Else

                ' We don't want to take the dataNodes xml as it will contain the tags (node name) for the node.
                ' So we need to take the data from the child nodes. If there is more than one child node we need
                ' to concatenate the xml data into a single string before it can be used.
                ' The xml (really xhtml) is then written to the html text file.
                If dataNode.ChildNodes.Length > 1 Then

                    ' Concatenate the data from multiple child nodes into a single string
                    For Each childNode In dataNode.ChildNodes

                        ' Check for a non default namespace (aka xmlns='http://www.w3.org/1999/xhtml' with no data.
                        If LenB(childNode.NamespaceURI) > 0 And LenB(childNode.Text) > 0 Then
                            childData = childData & childNode.XML
                        End If
                    Next

                    ' If there is no data add a null string to the Rich Text Dictionary object
                    If LenB(childData) > 0 Then

                        ' Add the xhtml/html to the html text file
                        g_htmlTextDocument.AddText bookmarkName, childData
                    Else

                        ' Add the variant containing a null string to the dictionary object using generated bookmark name as the key
                        AddStringVariantToRTDictionary bookmarkName, dataNode.Text
                    End If
                Else

                    Set childNode = dataNode.FirstChild

                    ' Do not add any null strings to the html text file or Word will error later when trying to copy a null string
                    If LenB(childNode.Text) > 0 Then

                        ' Add the xhtml/html to the html text file
                        g_htmlTextDocument.AddText bookmarkName, dataNode.FirstChild.XML
                    Else

                        ' Add the variant containing a null string to the dictionary object using generated bookmark name as the key
                        AddStringVariantToRTDictionary bookmarkName, dataNode.Text
                    End If
                End If
            End If
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RichTextInsert

'=======================================================================================================================
' Procedure:    AddStringVariantToRTDictionary
' Purpose:      Adds a string (as a variant) to the Rich Text Dictionary.
' Note:         We convert the data from a string to a variant so that it can be added to the dictionary.
'
' On Entry:     theKey              The key used to add the string (variant) to the Dictionary.
'               theString           The data to be added to the Dictionary.
'=======================================================================================================================
Private Sub AddStringVariantToRTDictionary(ByVal theKey As String, _
                                           ByVal theString As Variant)
    '  Add the string (as a variant) to the Rich Text Dictionary object
    g_richTextData.Add theKey, theString
End Sub ' AddStringVariantToRTDictionary

'=======================================================================================================================
' Procedure:    DeleteUnusedTextBlock
' Purpose:      .
' Notes:        .
'
' On Entry:     x                   .
'=======================================================================================================================
Private Sub DeleteUnusedTextBlock(ByVal bookmarkToDelete As String)
    Const c_proc As String = "modCreateDocument.DeleteUnusedTextBlock"

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
' Notes:        .
'
' On Entry:     bbName              The name of the Building Block to insert.
'               bbWhere             Where relative to the Range specified by bookmarkName the Building Block will be
'                                   inserted.
'               bookmarkName        The name of the Bookmark that identifies the Range where the Buildin block will be
'                                   inserted.
'               BookmarkPattern     A parameterised bookmark name used to generate a unique name for the Range where the
'                                   Building Block was inserted.
'               insertPosition      ???.
'               bookmarkExtend      Whether the Range specified by bookmarkName should be extended to include the
'                                   Building Block being inserted.
'=======================================================================================================================
Public Sub InsertBuildingBlock(ByVal bbName As String, _
                               ByVal bbWhere As rdaWhere, _
                               ByVal bookmarkName As String, _
                               ByVal BookmarkPattern As String, _
                               ByVal insertPosition As Long, _
                               ByVal bookmarkExtend As Boolean)
    Const c_proc As String = "modCreateDocument.InsertBuildingBlock"

    Dim originalBMEnd   As Long
    Dim originalBMStart As Long
    Dim target          As Word.Range
    Dim targetCopy      As Word.Range
    Dim useTemplate     As Word.Template
    Dim whereInserted   As Word.Range

    On Error GoTo Do_Error

    ' Add the choosen Building Block to the document
    If LenB(bbName) > 0 Then

        ' This is the template that contains all Building Blocks we insert in the Assessment Report
        Set useTemplate = g_assessmentReport.AttachedTemplate

        ' Replace any numeric placeholders in the Bookmark name with their corresponding values
        bookmarkName = g_counters.UpdatePredicates(bookmarkName)

        ' Location to insert the Building Block
        Set target = g_assessmentReport.bookmarks(bookmarkName).Range
        originalBMStart = target.Start
        originalBMEnd = target.End

        ' Check to see if we need to insert a Paragraph or NewLine before we insert the Building Block
        Select Case bbWhere
        Case rdaWhereAfterLastParagraph

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

        Case rdaWhereAtEndOfRange

            target.Collapse wdCollapseEnd

        Case rdaWhereReplaceRange
            
            ' Leave the target range unchanged as we want to replace it
        End Select

        EventLog "Inserting Building Block: " & bbName

        ' Insert the specified Building Block from the Assessment Report template
        Set whereInserted = useTemplate.BuildingBlockEntries(bbName).Insert(target, True)

        ' Bookmark protection code: Remove the paragraph mark we added above as it has served its purpose and is no longer required
        If bbWhere = rdaWhereAfterLastParagraph And insertPosition > 1 Then
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
        If bbWhere <> rdaWhereReplaceRange Then
            Set target = g_assessmentReport.bookmarks(bookmarkName).Range
            target.End = whereInserted.End

            ' The choice is to either extend the bookmark or to redefine it to what it originally was
            If bookmarkExtend Then
                g_assessmentReport.bookmarks.Add bookmarkName, target
            Else

                Set targetCopy = target.Duplicate
                targetCopy.Start = originalBMStart
                targetCopy.End = originalBMEnd
                g_assessmentReport.bookmarks.Add bookmarkName, targetCopy
            End If
        End If

        ' See if a second bookmark should be create using a name derived from the bookmarkPattern
        If LenB(BookmarkPattern) > 0 Then

            ' Generate the real bookmark name from the pattern
            BookmarkPattern = g_counters.UpdatePredicates(BookmarkPattern)

            g_assessmentReport.bookmarks.Add BookmarkPattern, whereInserted
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' InsertBuildingBlock

'=======================================================================================================================
' Procedure:    FillBookmarkRichText
' Purpose:      .
' Notes:        .
'
' On Entry:     info                The ActionInsert object for which RIchText will be inserted.
'=======================================================================================================================
Private Function FillBookmarkRichText(ByVal info As ActionInsert) As Word.Range
    Const c_proc As String = "modCreateDocument.FillBookmarkRichText"

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
    rawName = ReplacePatternData(info.BookmarkPattern, info.PatternData)
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
        ' override any formatting assosicated with the paragraph being pasted into in the Assessment Report.
        FixupBullets theHTMLText, targetArea, keyBookmark

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
' Notes:        The location where the replacement data should be inserted is defined by rdaPDP1.
'
' On Entry:     inputString         The string containing the bookmark name and pattern placeholder text to be updated.
'               xpathQuery          The xpath query supplying the data for the replaceble parameter.
'=======================================================================================================================
Public Function ReplacePatternData(ByVal inputString As String, _
                                   ByVal xpathQuery As String) As String
    Const c_proc As String = "modCreateDocument.ReplacePatternData"

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
                ReplacePatternData = Replace$(inputString, rdaPDP1, dataNode.Text)
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
' Notes 1:      This is necessary as the Assessment Report is Read Only protected (wdAllowOnlyReading). if we don't do
'               this the user is unable to edit the Assessment Report.
' Notes 2:      We also add the Editor object that allows the Range to be edited to a Dictionary object so that we can
'               recreate the Bookmark associated with the editable range. If the user copies and pastes an entire input
'               area into another complete input area it destroys the Bookmark of the area being pasted into. This is
'               contrary to what Word normal does, where you would expect to see the source Bookmark removed and
'               recreated at the paste location.
'
' On Entry:     targetArea          The area that should be authorised as being editable.
'               bookmarkName        The name of the Bookmark to associate with the editable area.
'=======================================================================================================================
Private Sub SetAsEditableRange(ByVal targetArea As Word.Range, _
                               ByVal bookmarkName As String)
    Const c_proc As String = "modCreateDocument.SetAsEditableRange"

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

'=======================================================================================================================
' Procedure:    FixupBullets
' Purpose:      .
' Notes:        .
'
' On Entry:     sourceHTMLText      .
'               targetArea          .
'=======================================================================================================================
Private Sub FixupBullets(ByVal sourceHTMLText As Word.Range, _
                         ByVal targetArea As Word.Range, _
                         ByVal bm As String)
    Const c_proc As String = "modCreateDocument.FixupBullets"

    Dim finalParagraphSource    As Word.Range
    Dim finalParagraphTarget    As Word.Range
    Dim paragraphCount          As Long
    Dim listTemplateSource      As Word.ListTemplate

    On Error GoTo Do_Error

    ' Get the final paragraph of the html derived text
    paragraphCount = sourceHTMLText.Paragraphs.Count
    Set finalParagraphSource = sourceHTMLText.Paragraphs(paragraphCount).Range

    ' See if the final paragraph of the source html derived text has a bullet
    If finalParagraphSource.ListFormat.CountNumberedItems > 0 Then

        ' Get the last paragraph of the target text
        paragraphCount = targetArea.Paragraphs.Count
        Set finalParagraphTarget = targetArea.Paragraphs(paragraphCount).Range

        ' Get the source final paragraphs ListTemplate object (which contains the Bullet info)
        Set listTemplateSource = finalParagraphSource.ListFormat.ListTemplate

        ' Now apply the ListTemplate to the final target paragraph to apply the html sources Bullet
        finalParagraphTarget.ListFormat.ApplyListTemplate listTemplateSource, True, wdListApplyToThisPointForward
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' FixupBullets

Public Sub RestoreBullets(ByVal areaOfInterest As Word.Range)
    Const c_proc As String = "modCreateDocument.RestoreBullets"

    Dim indexParagraph      As Word.Paragraph
    Dim bulletPrefixLength  As Long
    Dim prefixText          As String
    Dim paragraphRange      As Word.Range

    On Error GoTo Do_Error

    ' Short circuit the paragraph by paragraph search by searching the entire area for a Bullet sequence
    If InStr(areaOfInterest.Text, mgrBulletRecognitionSequence) > 0 Then

        ' Now iterate each paragraph in the editable area looking for bulleted paragraphs
        bulletPrefixLength = Len(mgrBulletRecognitionSequence)
        For Each indexParagraph In areaOfInterest.Paragraphs

            ' We are only interested if the paragraph has a bullet point
            Set paragraphRange = indexParagraph.Range
            prefixText = Left$(paragraphRange.Text, bulletPrefixLength)
            If prefixText = mgrBulletRecognitionSequence Then

                ' Strip the special character sequence used to indicate that this paragraph should have a bullet
                paragraphRange.End = paragraphRange.Start + bulletPrefixLength
                paragraphRange.Delete

                ' Now apply the default bullet
                paragraphRange.ListFormat.ApplyBulletDefault
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
    Const c_proc As String = "modCreateDocument.TableFixerUpper"

    Dim childTable  As Word.Table
    Dim mainTable   As Word.Table
    Dim thePS       As Word.PageSetup
    Dim usableWidth As Single

    On Error GoTo Do_Error

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

Private Function TickBookmark(ByVal bookmarkName As String, _
                              ByVal patternBookmark As String, _
                              ByVal patternQuery As String, _
                              ByVal tickValue As Boolean) As Word.Range
    Const c_proc As String = "modCreateDocument.TickBookmark"

    Dim targetArea As Word.Range

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

        ' Insert a tick or delete the bookmark contents, both delete the actual bookmark which will require recreating
        If tickValue Then
            targetArea.InsertSymbol mc_TickCharacterNumber, mc_TickCharacterFont, True

            ' Modify the target area to include the character we have just inserted
            targetArea.MoveEnd wdCharacter, 1
        Else
            targetArea.Text = vbNullString
        End If

        ' Recreate the bookmark
        Set TickBookmark = .Add(bookmarkName, targetArea).Range

        ' If a patternBookmark name has been supplied, create a second bookmark using that name
        If LenB(patternBookmark) Then
            .Add patternBookmark, targetArea
        End If
    End With

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' TickBookmark

Public Sub UpdateAllRefFields()
    Const c_proc As String = "modCreateDocument.UpdateAllRefFields"

    Dim arField  As Word.Field

    On Error GoTo Do_Error

    ' Iterate all fields in the Assessment Report, we check the type just in case the user has inserted a field
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
' Purpose:      .
' Notes:        .
'
' On Entry:     bookmarkName        The bookmark to update.
'               patternBookmark     Secondary bookmark name assigned to the original 'bookmarkName' location.
'               deletableBookmark   A bookmark name that species a range that should be deleted if 'newValue' is a null
'                                   string.
'               newValue            The value to update the location specied by BookmarkName with.
' Returns:      The Range object of the updated bookmark.
'=======================================================================================================================
Private Function UpdateBookmark(ByVal bookmarkName As String, _
                                ByVal patternBookmark As String, _
                                ByVal patternQuery As String, _
                                ByVal deletableBookmark As String, _
                                ByVal newValue As String, _
                                Optional ByVal outputNullString As Boolean) As Word.Range
    Const c_proc As String = "modCreateDocument.UpdateBookmark"

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
