VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionDelete
' Purpose:      Data class that models the 'delete' action node of the 'instructions.xml' file.
'
' Note 1:       Delete a specified bookmark and the text and tables, or just tables or just the paragraph mark immediately before
'               the bookmark start position.
'
' Note 2:       Special provision is made to ensure that any tables within 'bookmark' range are deleted and not just the table
'               contents.
'
' Note 3:       The following attributes are available for use with this Action:
'               bookmark [Mandatory]
'                   The name of a Bookmark that will be deleted including any bookmark content.
'               preserveEditableBookmarks [Optional]
'                   True = Prevents any editable bookmarks in the range being deleted from being removed from the Editable Bookmarks
'                   collection. This is critical when creating report variants Full Report and Summary Report. Without this being
'                   specified the Base assessment report will become unusable as our interal data structures will not match the
'                   assessment report.
'               what [Optional]
'                   What should be deleted, permissible values are:
'                       'PMBefore', 'Table' and 'All'. All is the default value.
'                       'PMBefore' = Paragraph Mark Before the specified bookmark.
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
' History:      25/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction

' AN = Attribute Name
Private Const mc_ANBookmark                     As String = "bookmark"                      ' The bookmark name (and range) to be deleted
Private Const mc_ANBreak                        As String = "break"                         ' Break if value is True
Private Const mc_ANPreserveEditableBookmarks    As String = "preserveEditableBookmarks"     ' Do not remove Editable Bookmarks from the editable bookmark collection
Private Const mc_ANWhat                         As String = "what"                          ' Exactly what should be deleted

' AWV = Attribute What Value
Private Const mc_AWVAll                         As String = "All"
Private Const mc_AWVPMBefore                    As String = "PMBefore"
Private Const mc_AWVTable                       As String = "Table"


Private m_break                     As Boolean                          ' Used to cause the code to execute a Stop instruction
Private m_bookmark                  As String                           ' The bookmark name (and range) to be deleted
Private m_preserveEditableBookmarks As Boolean                          ' True = Prevents the Editable Bookmark from being removed from the editable bookmarks collection
Private m_what                      As ssarWhatType                     ' Exactly what to delete from the assessment report


'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Delete action instruction.
' Date:         25/06/16    Created.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionDelete.IAction_Parse"

    Dim bookmarkCount   As Long
    Dim preserveCount   As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode
    Dim whatCount       As Long

    On Error GoTo Do_Error

    ' Default values if attributes are not present
    m_preserveEditableBookmarks = False
    m_what = ssarWhatTypeAll

    ' Parse out the 'delete' nodes attributes
    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANBookmark
            bookmarkCount = bookmarkCount + 1
            m_bookmark = theAttribute.Text

        Case mc_ANPreserveEditableBookmarks
            preserveCount = preserveCount + 1
            m_preserveEditableBookmarks = CBool(theAttribute.Text)

        Case mc_ANWhat
            whatCount = whatCount + 1

            ' Parse out the what attribute
            m_what = ParseAttributeWhat(actionNode, theAttribute)

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure the Mandatory attributes are present
    If bookmarkCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If bookmarkCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    ElseIf preserveCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANPreserveEditableBookmarks, c_proc
    ElseIf whatCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANWhat, c_proc
    End If

    ' Make sure these attributes contain some data
    If LenB(m_bookmark) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANBookmark, m_bookmark, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

Private Function ParseAttributeWhat(ByRef actionNode As MSXML2.IXMLDOMNode, _
                                    ByRef theAttribute As MSXML2.IXMLDOMNode) As ssarWhatType
    Const c_proc As String = "ActionDelete.ParseAttributeWhat"

    On Error GoTo Do_Error

    Select Case theAttribute.Text
    Case mc_AWVAll
        ParseAttributeWhat = ssarWhatTypeAll

    Case mc_AWVTable
        ParseAttributeWhat = ssarWhatTypeTable

    Case mc_AWVPMBefore
        ParseAttributeWhat = ssarWhatTypePMBefore

    Case Else
        InvalidAttributeWhatValueError actionNode.nodeName, theAttribute.BaseName, theAttribute.Text, c_proc
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ParseAttributeWhat

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Deletes the specified bookmark and any tables and text it contains while building as assessment report.
' Date:         25/06/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionDelete.IAction_BuildAssessmentReport"

    Dim bookmarkRange   As Word.Range
    Dim deletionRange   As Word.Range
    Dim index           As Long
    Dim realBoookmark   As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Replace any predicate placeholder text with the predicate value
    realBoookmark = g_counters.UpdatePredicates(m_bookmark)

    With g_assessmentReport.bookmarks
        If .Exists(realBoookmark) Then

            ' Get the bookmarks range, so that it can be deleted
            Set bookmarkRange = .Item(realBoookmark).Range

            Select Case m_what
            Case ssarWhatTypeAll

                ' Remove any editable bookmarks from the editable bookmarks object before the range is deleted (it is too late after!).
                If Not m_preserveEditableBookmarks Then
                    g_editableBookmarks.DeleteInRange bookmarkRange
                End If

                ' Delete all tables present in the range
                For index = 1 To bookmarkRange.Tables.Count
                    bookmarkRange.Tables(1).Delete
                Next

                ' If the bookmark range is collapsed (just an insertion point) deleting the insertion point will delete
                ' the character to the right of it, which we do not want! So check that the Range is not collapsed.
                If bookmarkRange.Start <> bookmarkRange.End Then

                    ' Delete any other bookmark content and the bookmark itself
                    bookmarkRange.Delete
                End If
            Case ssarWhatTypeTable

                ' Remove any editable bookmarks from the editable bookmarks object before the range is deleted (it is too late after!)
                If Not m_preserveEditableBookmarks Then
                    g_editableBookmarks.DeleteInRange bookmarkRange
                End If

                ' Delete all tables present in the range
                For index = 1 To bookmarkRange.Tables.Count
                    bookmarkRange.Tables(1).Delete
                Next
            Case ssarWhatTypePMBefore

                ' Delete the paragraph mark (not the paragraph contents) immediately before the start of the bookmarked range
                bookmarkRange.Collapse wdCollapseStart
                bookmarkRange.MoveStart wdCharacter, -1
                bookmarkRange.Delete
            Case Else
                Err.Raise mgrErrNoUnexpectedCondition, c_proc, mgrErrTextUnexpectedCondition
            End Select
        End If
    End With

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

Private Sub IAction_UpdateContentControlXML()
    ' Definition required - But nothing to do
End Sub ' IAction_UpdateContentControlXML

Private Sub IAction_UpdateDateXML()
    ' Definition required - But nothing to do
End Sub ' Sub IAction_UpdateDateXML

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeDelete
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break

Friend Property Get Bookmark() As String
    Bookmark = m_bookmark
End Property ' Get Bookmark

Friend Property Get PreserveEditableBookmarks() As Boolean
    PreserveEditableBookmarks = m_preserveEditableBookmarks
End Property ' Get preserveEditableBookmarks

Friend Property Get What() As ssarWhatType
    What = m_what
End Property ' Get What
