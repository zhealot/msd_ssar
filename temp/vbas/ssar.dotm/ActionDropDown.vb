VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionDropDown
' Purpose:      Data class that models the 'dropDown' action node of the 'instructions.xml' file.
'
' Note 1:       Implements a DropDown Content Control.
'
' Note 2:       This is a variant of ActionInsert, but it is different enough to warrant its own class.
'
' Note 3:       The following attributes are available for use with this Action:
'               bookmark [Mandatory]
'                   The name of a Bookmark that contains the Content Control to be mapped to the Custom XML Part data source node.
'               coloured [Optional]
'                   True if the text of the selected DropDown Content Control entry should be matched against the Foreground and
'                   Background colour tables and the resulting lookup value used to set the text colour and background colour.
'               ccDataNode [Mandatory]
'                   The name of an xml node added to the Custom XML Part mapped Content Control data store root node.
'                   Critical: This name must be unique amongst all mapped nodes in the mapped Content Control data store.
'               data [Mandatory]
'                   An xpath query run against the rda xml, whose results are copied to and from the mapped Custom XML Part data
'                   source node.
'               pattern [Optional]
'                   A Bookmark with a unique name, which duplicates the Range of the Bookmark specified by the 'bookmark'
'                   attribute. This is used when there are repeating Content Controls with an Action Add block. The Bookmark name
'                   specified by the 'bookmark' attribute is generally reused, buy using this mechanism we end up with a Bookmark
'                   with a unique name.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
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
' History:      18/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBookmark     As String = "bookmark"          ' The bookmark range that contacins the content control
Private Const mc_ANBreak        As String = "break"             ' Break if value is True
Private Const mc_ANCCDataNode   As String = "ccDataNode"        ' The Content Control mapped data store node (just the node name not the full xpath)
Private Const mc_ANColoured     As String = "coloured"          ' Coloured text and background
Private Const mc_ANData         As String = "data"              ' The rda xpath query to the data
Private Const mc_ANPattern      As String = "pattern"           ' A unique name for the Bookmark when there are repeating Content Controls


Private m_bookmark          As String                       ' The name of the bookmark that contains the content control
Private m_bookmarkPattern   As String                       ' A unique name for the bookmark when there are repeating Content Controls
Private m_break             As Boolean                      ' Used to cause the code to execute a Stop instruction
Private m_ccDataNode        As String                       ' The node in the custom xml data store part that stores the data for the mapped Content Control
Private m_coloured          As Boolean                      ' When True, the text and background will be coloured based on the selected text

Private m_rdaData           As String                       ' The xpath query to the rda xml data


'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the DropDown action instruction.
' Date:         18/06/16    Created.
'
' On Entry:     actionNode          The xml node containing the drop down instruction to parse.
'               addPatternBookmarks  Unused.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionDropDown.IAction_Parse"

    Dim bookmarkCount   As Long
    Dim ccDataNodeCount As Long
    Dim colouredCount   As Long
    Dim patternCount    As Long
    Dim rdaDataCount    As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out the 'dropDown' nodes attributes
    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANBookmark
            bookmarkCount = bookmarkCount + 1
            m_bookmark = theAttribute.Text

        Case mc_ANData
            rdaDataCount = rdaDataCount + 1
            m_rdaData = theAttribute.Text

        Case mc_ANCCDataNode
            ccDataNodeCount = ccDataNodeCount + 1
            m_ccDataNode = theAttribute.Text

        Case mc_ANPattern
            patternCount = patternCount + 1
            m_bookmarkPattern = theAttribute.Text

        Case mc_ANColoured
            colouredCount = colouredCount + 1
            m_coloured = CBool(theAttribute.Text)

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure these attributes are present
    If bookmarkCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    ElseIf rdaDataCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANData, c_proc
    ElseIf ccDataNodeCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If bookmarkCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    ElseIf rdaDataCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANData, c_proc
    ElseIf ccDataNodeCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    ElseIf patternCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANPattern, c_proc
    ElseIf colouredCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANColoured, c_proc
    End If

    ' Make sure the attributes contain some data
    If LenB(m_bookmark) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANBookmark, m_bookmark, c_proc
    ElseIf LenB(m_ccDataNode) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANCCDataNode, m_bookmark, c_proc
    ElseIf LenB(m_rdaData) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANData, m_rdaData, c_proc
    ElseIf patternCount > 0 And LenB(m_bookmarkPattern) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANPattern, m_bookmarkPattern, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Creates the Custom XML Part data store node for the Content Control and then maps the Content Control to it.
' Date:         20/06/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionDropDown.IAction_BuildAssessmentReport"

    Dim bookmarkName        As String
    Dim bookmarkPatternName As String
    Dim bookmarkRange       As Word.Range
    Dim errorText           As String
    Dim theContentControl   As Word.ContentControl

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Get the name of the bookmark that contains the Content Control
    bookmarkName = m_bookmark

    ' Locate the Content Control
    With g_assessmentReport.bookmarks
        If .Exists(bookmarkName) Then
            Set bookmarkRange = .Item(bookmarkName).Range
            If bookmarkRange.ContentControls.Count > 0 Then
                Set theContentControl = bookmarkRange.ContentControls(1)
            Else
                errorText = Replace$(mgrErrTextContentControlMissingFromBookmark, mgrP1, bookmarkName)
                Err.Raise mgrErrNoContentControlMissingFromBookmark, c_proc, errorText
            End If
        Else
            BookmarkDoesNotExistError bookmarkName, c_proc
        End If
    End With

    ' Create the Custom XML Part data store node and then map the Content Control to it
    g_ccXMLDataStore.MapDropDown theContentControl, m_rdaData, m_ccDataNode

    ' If necessary Create a unique bookmark name as the default bookmark name may get reused elsewhere
    If LenB(m_bookmarkPattern) > 0 Then

        ' Replace any counter placeholder text with the appropriate counter value
        bookmarkPatternName = g_counters.UpdatePredicates(m_bookmarkPattern)

        ' Create the unique bookmark name
        g_assessmentReport.bookmarks.Add bookmarkPatternName, bookmarkRange
    End If

    ' Apply text and background colours based on the selected drop down value
    If m_coloured Then
        ApplyColourMap bookmarkRange, m_ccDataNode
    End If

    ' The assessment report must be "Write View" for it to be editable.
    ' The user must not be able to edit the Assessment Report if it is "Read View" or "Print View".
    If g_rootData.IsWritable Then

        ' Note: The bookmark range and the content control range are NOT congruent. Word does not allow you to create a
        ' bookmark that exactly matches the content control range. So the bookmarks for all Content Controls in the
        ' assessment report template start one character before the content control and end one character after the
        ' content control. This also impacts the code that compares content controls to their bookmarked range.
        If LenB(bookmarkPatternName) > 0 Then

            ' Add the Content Control range to the editable range dictionary object using the bookmark pattern name as the key
            SetAsEditableRange theContentControl.Range, bookmarkPatternName
        Else

            ' Add the Content Control range to the editable range dictionary object using the bookmark name as the key
            SetAsEditableRange theContentControl.Range, bookmarkName
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

Private Sub IAction_ConstructRichText()
    ' Definition required - But nothing to do
End Sub ' IAction_ConstructRichText

Private Sub IAction_HTMLForXMLUpdate()
    ' Definition required - But nothing to do
End Sub '  Sub IAction_HTMLForXMLUpdate

'===================================================================================================================================
' Procedure:    IAction_UpdateContentControlXML
' Purpose:      Updates the assessment report XML using the value of the Content Control.
' Note 1:       Using the Content Control directly (rather than the assessment report document) saves us having to mess around and
'               sort out the padding we add to the Content Control bookmark.
' Date:         16/08/16    Created.
'===================================================================================================================================
Private Sub IAction_UpdateContentControlXML()
    Const c_proc As String = "ActionDropDown.IAction_UpdateContentControlXML"

    On Error GoTo Do_Error

    ' Update the assessment report xml from the Custom XML Part mapped Content Control data store node
    g_ccXMLDataStore.UpdateRDAXML m_rdaData, m_ccDataNode

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_UpdateContentControlXML

Private Sub IAction_UpdateDateXML()
    ' Definition required - But nothing to do
End Sub ' Sub IAction_UpdateDateXML

Friend Property Get Bookmark() As String
    Bookmark = m_bookmark
End Property ' Get Bookmark

Friend Property Get bookmarkPattern() As String
    bookmarkPattern = m_bookmarkPattern
End Property 'Get BookmarkPattern

Friend Property Get ccDataNode() As String
    ccDataNode = m_ccDataNode
End Property ' Get CCDataNode

Friend Property Get Coloured() As Boolean
    Coloured = m_coloured
End Property ' Coloured

Friend Property Get RDAData() As String
    RDAData = m_rdaData
End Property ' Get RDAData

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeDropDown
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break

