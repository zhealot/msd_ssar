VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ActionInsert
' Purpose:      Data class that models the 'insert' action node of the 'instructions.xml' file.
' Note 1:       The Insert Action inserts data into a bookmark in the assessment report. The bookmark can be editable or
'               non-editable.
'
' Note 2:       The following attributes are available for use with this Action:
'               bookmark [Mandatory]
'                   The name of a Bookmark to update with the results of the 'data' xpath query.
'               data [Mandatory]
'                   An xpath query run against the rda xml, whose results yield the value to update 'bookmark' with.
'               dataFormat [Optional]
'                   The format of the data inserted into the bookmarked area (specified by the 'bookmark' attribute), values are:
'                       Text        [The default if the attribute is omitted]
'                       MultiLine
'                       RichText
'                       DateShort
'                       DateLong
'                       Long
'               defaultText [Unused]
'                   Partially implemented to allow a default value to be displayed if the data returned by the 'data' attributes
'                   xpath query was null.
'               deleteIfNull [Optional]
'                   The name of a Bookmark whose contents will be deleted if the 'data' attributes xpath query returns no nodes.
'               editable [Optional]
'                   Whether the Bookmarked area of the document should be editable (True/False).
'                   [Default = False]
'               pattern [Optional]
'                   A Bookmark with a unique name, which duplicates the Range of the Bookmark specified by the 'bookmark'
'                   attribute. This is used when there are repeating blocks of text with an Action Add block. The Bookmark name
'                   specified by the 'bookmark' attribute is generally reused, buy using this mechanism we end up with a Bookmark
'                   with a unique name.
'               patternData [Optional]
'                   An xpath query that yields a value to use to generate a unique bookmark named based on the contents of an
'                   assessment report node.
'               coloured [Optional]
'                   Colours the text and backround based on the value of the text using the values defined in the foreground and
'                   background Colour Maps.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
'
' Note 3:       Insert nodes with a dataFormat of 'Multiline' or 'RichText' must have a 'pattern' attribute value as this value is
'               used as a compact bookmark name.
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
' History:      01/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBookmark         As String = "bookmark"
Private Const mc_ANBreak            As String = "break"
Private Const mc_ANColoured         As String = "coloured"
Private Const mc_ANData             As String = "data"
Private Const mc_ANDataFormat       As String = "dataFormat"
Private Const mc_ANDefaultText      As String = "defaultText"
Private Const mc_ANDeleteIfNull     As String = "deleteIfNull"
Private Const mc_ANEditable         As String = "editable"
Private Const mc_ANPattern          As String = "pattern"
Private Const mc_ANPatternData      As String = "patternData"       ' Used for the xpath query


' DFV = DataFormat node Values
Private Const mc_DFVDateLong        As String = "DateLong"
Private Const mc_DFVDateShort       As String = "DateShort"
Private Const mc_DFVDateShortYear   As String = "DateShortYear"
Private Const mc_DFVLong            As String = "Long"
Private Const mc_DFVMultiline       As String = "Multiline"
Private Const mc_DFVRichText        As String = "RichText"
Private Const mc_DFVText            As String = "Text"


Private m_bookmark          As String
Private m_bookmarkPattern   As String
Private m_break             As Boolean          ' Used to cause the code to execute a Stop instruction
Private m_coloured          As Boolean          ' When True, the text and background will be coloured based on the actual text
Private m_defaultText       As String
Private m_deleteIfNull      As String           ' Delete this bookmark if the node/nodeList returned by the xpath query contains no data
Private m_editable          As Boolean          ' If true the bookmark is added to the editable bookmarks collection
Private m_dataFormat        As ssarDataFormat
Private m_dataSource        As String           ' The xpath query that supplies the data
Private m_patternData       As String           ' The xpath query for the pattern parameter replacement value


'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses out the Action Insert xml node.
'
' On Entry:     actionNode          The Action Insert xml node to be parsed.
'               addPatternBookmarks  Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionInsert.IAction_Parse"

    Dim errorText       As String
    Dim theDataFormat   As String

    On Error GoTo Do_Error

    ' Default values, these values prevail when there is no xml data present for a particular attribute (the xml data is optional).
    ' All strings start out with their inherent value of vbNullString.
    m_editable = False
    m_dataFormat = ssarDataFormatText
    m_break = False

    ' Parse all attributes (each entry is a pair of values - the Attribute Name and Variable, to contain the Attribute value)
    ParseAttributes actionNode, mc_ANBookmark, m_bookmark, mc_ANData, m_dataSource, mc_ANEditable, m_editable, mc_ANDataFormat, theDataFormat, _
                                mc_ANDeleteIfNull, m_deleteIfNull, mc_ANPattern, m_bookmarkPattern, mc_ANPatternData, m_patternData, _
                                mc_ANColoured, m_coloured, mc_ANDefaultText, m_defaultText, mc_ANBreak, m_break

    ' Parse out the 'dataSource' nodes value
    Select Case theDataFormat
    Case mc_DFVText
        m_dataFormat = ssarDataFormatText

    Case mc_DFVRichText
        m_dataFormat = ssarDataFormatRichText

    Case mc_DFVMultiline
        m_dataFormat = ssarDataFormatMultiline

    Case mc_DFVDateLong
        m_dataFormat = ssarDataFormatDateLong

    Case mc_DFVDateShort
        m_dataFormat = ssarDataFormatDateShort

    Case mc_DFVDateShortYear
        m_dataFormat = ssarDataFormatDateShortYear

    Case mc_DFVLong
        m_dataFormat = ssarDataFormatLong

    Case Else
        errorText = Replace$(mgrErrTextInvalidDataFormatNodeValue, mgrP1, theDataFormat)
        Err.Raise mgrErrNoInvalidDataFormatNodeValue, c_proc, errorText

    End Select

' #TRY#  See if we can get by without defining a pattern bookmark by reusing the bookmark name
' For Text and Rich Text reuse the bookmark name as the bookmark patter name, when nobookmark patter name is present
If (m_dataFormat = ssarDataFormatRichText Or m_dataFormat = ssarDataFormatText) And LenB(m_bookmarkPattern) = 0 Then
    m_bookmarkPattern = m_bookmark
End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'=======================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Inserts data into bookmarks while Building the Assessment Report.
'=======================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionInsert.IAction_BuildAssessmentReport"

    Dim dataNode        As MSXML2.IXMLDOMNode
    Dim doInsert        As Boolean
    Dim rawName         As String
    Dim targetArea      As Word.Range
    Dim theQuery        As String
    Dim theText         As String
    Dim useBookmark     As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
    theQuery = g_counters.UpdatePredicates(m_dataSource)

    ' Get the data to update the bookmark with
    Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

    If dataNode Is Nothing Then
        If m_editable Then
            theText = m_defaultText
            doInsert = True
        End If
    Else

        ' Get the dataNodes text which will be used to update the bookmark (as long as the DataFormat is not RichText or Multiline)
        If m_dataFormat <> ssarDataFormatMultiline And m_dataFormat <> ssarDataFormatRichText Then
            theText = dataNode.Text
        End If
        doInsert = True
    End If

    ' Check that we actually retrieved a node
    If doInsert Then

        ' Perform the appropriate update type
        Select Case m_dataFormat
        Case ssarDataFormatText
            EventLog "Updating (Text) bookmark: " & m_bookmark, c_proc

            ' Replace the bookmarked text with the dataNodes value
            Set targetArea = UpdateBookmark(m_bookmark, m_bookmarkPattern, m_patternData, m_deleteIfNull, theText, m_editable)

        Case ssarDataFormatRichText
            EventLog "Updating (RichText) bookmark: " & m_bookmark, c_proc

            ' Copy the RichText (xhtml) from the Word HTML document and paste it into the assessment report
            Set targetArea = FillBookmarkRichText(Me)

        Case ssarDataFormatMultiline
            EventLog "Updating (MultiLine) bookmark: " & m_bookmark, c_proc

            ' Copy the RichText (xhtml) from the Word HTML document and paste it into the assessment report
            Set targetArea = FillBookmarkRichText(Me)

        Case ssarDataFormatDateLong
            EventLog "Updating (DateLong) bookmark: " & m_bookmark, c_proc

            ' Replace the bookmarked text with the dataNodes value
            Set targetArea = UpdateBookmark(m_bookmark, m_bookmarkPattern, m_patternData, m_deleteIfNull, _
                                            Format$(theText, mgrDateFormatLong), m_editable)

        Case ssarDataFormatDateShort
            EventLog "Updating (DateShort) bookmark: " & m_bookmark, c_proc

            ' Replace the bookmarked text with the dataNodes value
            Set targetArea = UpdateBookmark(m_bookmark, m_bookmarkPattern, m_patternData, m_deleteIfNull, _
                                            Format$(theText, mgrDateFormatShort), m_editable)

        Case ssarDataFormatDateShortYear
            EventLog "Updating (DateShortYear) bookmark: " & m_bookmark, c_proc

            ' Replace the bookmarked text with the dataNodes value
            Set targetArea = UpdateBookmark(m_bookmark, m_bookmarkPattern, m_patternData, m_deleteIfNull, _
                                            Format$(theText, mgrDateFormatShortYear), m_editable)
        Case ssarDataFormatLong
            EventLog "Updating (Long) bookmark: " & m_bookmark, c_proc

            ' Replace the bookmarked text with the dataNodes value
            Set targetArea = UpdateBookmark(m_bookmark, m_bookmarkPattern, m_patternData, m_deleteIfNull, theText, m_editable)
        End Select

        ' If the range is editable then set it as editable
        If m_editable Then

            ' By preference use the BookmarkPattern which is the secondary Bookmark
            ' name and used where Bookmark names containing a value are incremented
            If LenB(m_bookmarkPattern) > 0 Then

                ' Replace any pattern data parameters (data obtained from the actual Assessment Report) and
                ' numeric parameters in the pattern Bookmark with their corresponding values
                rawName = ReplacePatternData(m_bookmarkPattern, m_patternData)
                useBookmark = g_counters.UpdatePredicates(rawName)
            Else
                useBookmark = m_bookmark
            End If

            ' Furthermore, it must be "Write View" for it to be editable.
            ' The user must not be able to edit the Assessment Report if it is "Read View" or "Print View".
            If g_rootData.IsWritable Then
                SetAsEditableRange targetArea, useBookmark
            End If
        Else

            ' Make sure a value is assigned to useBookmark as it is required if we need to apply a Colour Map
            useBookmark = m_bookmarkPattern
        End If

        ' Apply text and background colours based on the selected drop down value
        If m_coloured Then
            ApplyColourMap g_assessmentReport.bookmarks(useBookmark).Range
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

'=======================================================================================================================
' Procedure:    IAction_ConstructRichText
' Purpose:      Creates the Dictionary object and html file, which jointly are the source of all RichText data consumed
'               by the Assessment Report.
'=======================================================================================================================
Private Sub IAction_ConstructRichText()
    Const c_proc As String = "ActionInsert.IAction_ConstructRichText"

    Dim bookmarkName    As String
    Dim childData       As String
    Dim childNode       As MSXML2.IXMLDOMNode
    Dim dataNode        As MSXML2.IXMLDOMNode
    Dim plainText       As Boolean
    Dim rawName         As String
    Dim theQuery        As String

    ' Storage of RichText is optimised by storing RichText that contains only plain text in a Variant which
    ' is added to the dictionary. This means that we do not have to write the to thye html file or retrieve
    ' it from the html file, thus improving the overall speed of the Assessment Report generation.

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' We only need to process the query if the current ActionInsert object has a MultiLine or RichText data format
    If m_dataFormat = ssarDataFormatMultiline Or m_dataFormat = ssarDataFormatRichText Then

        ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
        theQuery = g_counters.UpdatePredicates(m_dataSource)

        ' Get the RichText node, it may contain rich text or plain text
        Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

        ' Create a unique bookmark name for this data node
        rawName = ReplacePatternData(m_bookmarkPattern, m_patternData)
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
End Sub ' IAction_ConstructRichText

'===================================================================================================================================
' Procedure:    IAction_HTMLForXMLUpdate
' Purpose:      Create the HTML that will be converted into XHTML and then used to update the assessment report xml.
' Note 1:       The word document we are adding the text from the editable areas to, must exist when this procedure is first called.
' Date:         07/08/16    Created.
'===================================================================================================================================
Private Sub IAction_HTMLForXMLUpdate()
    Const c_proc As String = "ActionInsert.IAction_HTMLForXMLUpdate"

    Dim dataNode   As MSXML2.IXMLDOMNode
    Dim theQuery   As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' We only need to update the RichText for fields that are marked as editable, otherwise they are unchanged
    If m_editable Then

        ' Perform the appropriate update type
        Select Case m_dataFormat
        Case ssarDataFormatRichText, ssarDataFormatMultiline

            ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
            theQuery = g_counters.UpdatePredicates(m_dataSource)

            ' Get the data to update the bookmark with
            Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

            ' Check that we actually retrieved a node
            If Not dataNode Is Nothing Then

                ' Copy the RichText from the Assessment Report to the Word XHTML document
                CopyRichTextBlock
            End If

        Case ssarDataFormatText
            UpdatePlainText

        Case ssarDataFormatLong, ssarDataFormatDateLong, ssarDataFormatDateShort, ssarDataFormatDateShortYear
            ' There is no need to do anything for these 'actions'
        End Select
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub '  Sub IAction_HTMLForXMLUpdate

Private Sub IAction_UpdateContentControlXML()
    ' Definition required - But nothing to do
End Sub ' IAction_UpdateContentControlXML

'===================================================================================================================================
' Procedure:    IAction_UpdateDateXML
' Purpose:      If the area is both editable and a date, then validate the date and update the corresponding assessment report xml.
' Date:         07/08/16    Created.
'===================================================================================================================================
Private Sub IAction_UpdateDateXML()
    Const c_proc As String = "ActionInsert.IAction_UpdateDateXML"

    Dim dataBookmarkName As String
    Dim dataNode         As MSXML2.IXMLDOMNode
    Dim dateInputArea    As Word.Range
    Dim dummy            As Date
    Dim thedate          As String
    Dim theQuery         As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' We only need to validate editable input areas
    If m_editable Then

        ' Perform the appropriate update type
        Select Case m_dataFormat
        Case ssarDataFormatDateLong, ssarDataFormatDateShort, ssarDataFormatDateShortYear

            ' Replace the predicate place holder (if present) with the predicate
            ' value to build the Bookmark name of an editable date area
            dataBookmarkName = ReplacePatternData(m_bookmarkPattern, m_patternData)
            dataBookmarkName = g_counters.UpdatePredicates(dataBookmarkName)

            ' See if the Bookmark actually exists
            If g_assessmentReport.bookmarks.Exists(dataBookmarkName) Then

                ' Retrieve the bookmark contents
                Set dateInputArea = g_assessmentReport.bookmarks(dataBookmarkName).Range
                thedate = dateInputArea.Text

                ' Null strings and valid dates are legal values
                If LenB(thedate) > 0 Then

                    ' If the text successfully converts to a date then it's valid
                    On Error Resume Next
                    dummy = CDate(thedate)
                    If Err.Number = mgrErrNoTypeMismatch Then

                        ' Set the font colour of the date input area to make it stand
                        ' out so that the user knows it contains an invalid date
                        dateInputArea.Font.ColorIndex = wdRed
                        Err.Clear

                        ' Bit of a nasty way to do this (global flag), but indicate that there has been a date that is not valid
                        g_dateValidationError = True
                    ElseIf Err.Number <> 0 Then
                        Err.Raise Err.Number, c_proc
                    Else

                        ' Get the Bookmarks corresponding Assessment Report xml data node so that we can update it
                        theQuery = g_counters.UpdatePredicates(m_dataSource)
                        Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

                        ' Now update the corresponding Assessment Report xml
                        If Not dataNode Is Nothing Then
                            dataNode.Text = thedate
                        End If

                        ' Clear the background colour just in case we previously set it to highlight an error
                        dateInputArea.Font.ColorIndex = wdAuto
                    End If
                End If
            End If

        Case Else
            ' There is no need to do anything for these data types

        End Select
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Sub IAction_UpdateDateXML

'=======================================================================================================================
' Procedure:    CopyRichTextBlock
' Purpose:      Copies an editable block of text from the Assessment Report to the html file used to generate xhtml.
' Notes:        It is quite possible that if this Insert action is part of an Add action block and the Add action has
'               a 'deleteIfNull' attribute that the whole block has not been generated. As a consequence of that the
'               bookmark for this ActionInsert may not exist.
'=======================================================================================================================
Private Sub CopyRichTextBlock()
    Const c_proc As String = "modHTMLToXHTML.CopyRichTextBlock"

    Dim queryKey        As String
    Dim sourceArea      As Word.Range
    Dim sourceBookmark  As String
    Dim target          As Word.Range

    On Error GoTo Do_Error

    ' Use the Pattern bookmark name if it is specified
    If LenB(m_bookmarkPattern) > 0 Then
        sourceBookmark = ReplacePatternData(m_bookmarkPattern, m_patternData)
    Else
        sourceBookmark = m_bookmark
    End If
    sourceBookmark = g_counters.UpdatePredicates(sourceBookmark)

    ' Create the queryKey (the xml node that will be updated using the following rich text plus a new paragraph.
    ' Wrap the query key in an leadin character block and a query end character block so that
    ' we can find the start and end of a complete entry (the query and the RichText block).
    queryKey = mgrHTMLBookmarkedBlockLeadIn & g_counters.UpdatePredicates(m_dataSource) & mgrHTMLBookmarkNameEnd & vbCr

    ' Make sure the source Bookmark actually exists, if it is in a nested Add block it may not
    If g_assessmentReport.bookmarks.Exists(sourceBookmark) Then

        ' Get a reference to the source Range
        Set sourceArea = g_assessmentReport.bookmarks(sourceBookmark).Range

        ' Set a Range object to the xhtml Word document so that we can add text to it
        Set target = g_xhtmlWordDoc.Content

        ' Make sure we add all new text at the very end of the document
        target.Collapse wdCollapseEnd

        ' If the xhtml Word document already contains text add a new paragraph to hold
        ' the query string so that it not contiguous to text already present
        If g_xhtmlWordDoc.Content.End > 1 Then
            target.InsertParagraph
        End If

        ' Add the queryKey (the xml node that will be updated using the following rich text
        target.InsertAfter queryKey
        target.Collapse wdCollapseEnd

        ' Check to see if the source range is using the Default Text (which we do not want to update the xml with)
        If HasDefaultText And m_defaultText = sourceArea.Text Then

            ' The source document is using the Default Text (which means that the user has not updated it).
            ' Since we do not want to propagate the Default Text to the xml replace the Default Text with a null string.
            target.InsertAfter vbNullString
        Else

            ' Make sure there is some text to copy or we get error 4605 "This method or property is not available because no text is selected."
            If LenB(sourceArea.Text) > 0 Then

                ' Copy the source RichText data
                sourceArea.Copy

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

'===================================================================================================================================
' Procedure:    UpdatePlainText
' Purpose:      Updates the assessment reports xml with plain text.
' Note 1:       This updates the xml when the class instance is defined as both Text and Editable.
' Note 2:       Because we cannot prevent a user from formatting the text within Word, it gets lost when we just return the text.
' Date:         17/08/16    Created
'===================================================================================================================================
Private Sub UpdatePlainText()
    Const c_proc As String = "ActionInsert.UpdatePlainText"

    Dim dataInputArea   As Word.Range
    Dim dataNode        As MSXML2.IXMLDOMNode
    Dim sourceBookmark  As String
    Dim theQuery        As String

    On Error GoTo Do_Error

    ' Use the Pattern bookmark name if it is specified
    If LenB(m_bookmarkPattern) > 0 Then
        sourceBookmark = ReplacePatternData(m_bookmarkPattern, m_patternData)
    Else
        sourceBookmark = m_bookmark
    End If
    sourceBookmark = g_counters.UpdatePredicates(sourceBookmark)

    ' Retrieve the bookmark contents
    Set dataInputArea = g_assessmentReport.bookmarks(sourceBookmark).Range

    ' Get the Bookmarks corresponding Assessment Report xml data node so that we can update it
    theQuery = g_counters.UpdatePredicates(m_dataSource)
    Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

    ' Now update the corresponding Assessment Report xml with unformatted text!
    If Not dataNode Is Nothing Then
        dataNode.Text = dataInputArea.Text
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' UpdatePlainText

Friend Property Get Bookmark() As String
    Bookmark = m_bookmark
End Property ' Get Bookmark

Friend Property Get bookmarkPattern() As String
    bookmarkPattern = m_bookmarkPattern
End Property 'Get BookmarkPattern

Friend Property Get Coloured() As Boolean
    Coloured = m_coloured
End Property ' Coloured

Friend Property Get DefaultText() As String
    DefaultText = m_defaultText
End Property ' Get DefaultText

Friend Property Get DeleteIfNull() As String
    DeleteIfNull = m_deleteIfNull
End Property ' Get DeleteIfNull

Friend Property Get DataFormat() As ssarDataFormat
    DataFormat = m_dataFormat
End Property ' Get DataFormat

Friend Property Get DataSource() As String
    DataSource = m_dataSource
End Property ' Get DataSource

Friend Property Get Editable() As Boolean
    Editable = m_editable
End Property ' Get Editable

Friend Property Get HasDefaultText() As Boolean
    HasDefaultText = (Len(m_defaultText) > 0)
End Property ' HasDefaultText

Friend Property Get PatternData() As String
    PatternData = m_patternData
End Property ' PatternData

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeInsert
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
