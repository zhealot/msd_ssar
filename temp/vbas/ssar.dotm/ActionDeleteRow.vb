VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionDeleteRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionDeleteRow
' Purpose:      Data class that models the 'deleteRow' action node-tree of the 'instructions.xml' file.
'
' Note 1:       Deletes a specified row from the table in the specified bookmark.
'
' Note 2:       The following attributes are available for use with this Action:
'               tableBookmark   [Mandatory]
'                   The name of a bookmark that wholely or partially contains the table the row will be deleted from.
'               row [Mandatory]
'                   The table row which should be deleted.
'                   This can be either a numeric value or the name of a named counter.
'               break [Optional]
'                   True, then code breaks before execution of the 'add' instruction.
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
' History:      25/07/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBreak                    As String = "break"
Private Const mc_ANRow                      As String = "row"                       ' The number of the row in table to be deleted
Private Const mc_ANTableBookmark            As String = "tableBookmark"             ' The name of a bookmark that wholely or partially contains the table the row will be deleted from


Private m_break                 As Boolean              ' Used to cause the code to execute a Stop instruction
Private m_row                   As Long
Private m_rowCounter            As String               ' The name of the named counter
Private m_tableBookmark         As String


'===================================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the DeleteRow action instruction.
' Date:         25/07/16    Created.
'
' On Entry:     actionNode          The xml node containing the deleteRow instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'===================================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionDeleteRow.IAction_Parse"

    Dim rowCount            As Long
    Dim tableBookmarkCount  As Long
    Dim theAttribute        As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out the 'addRow' nodes attributes
    For Each theAttribute In actionNode.Attributes
    
        Select Case theAttribute.BaseName
        Case mc_ANTableBookmark
            tableBookmarkCount = tableBookmarkCount + 1
            m_tableBookmark = theAttribute.Text

        Case mc_ANRow
            rowCount = rowCount + 1

            ' If value is non numeric then it must be a named counter
            If IsNumeric(theAttribute.Text) Then
                m_row = CLng(theAttribute.Text)
            Else
                m_rowCounter = Trim$(theAttribute.Text)
                If LenB(m_rowCounter) = 0 Then
                    InvalidAttributeValueError actionNode.nodeName, mc_ANRow, m_rowCounter, c_proc
                End If
            End If

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure the Mandatory attributes are present
    If tableBookmarkCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANTableBookmark, c_proc
    ElseIf rowCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANRow, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If tableBookmarkCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANTableBookmark, c_proc
    ElseIf rowCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANRow, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Deletes the specified row from the specified Standards Exceptions table.
' Notes:        The actual standard whose Exception table is updated is implicitly specified by the predicate index of g_counters.
' Date:         25/07/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionDeleteRow.IAction_BuildAssessmentReport"

    Dim bookmarkRange       As Word.Range
    Dim deleteRow           As Long
    Dim deletionRange       As Word.Range
    Dim tableBookmark       As String
    Dim theTable            As Word.Table

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Get the name of the bookmark that contains the table we want to insert a row into
    tableBookmark = m_tableBookmark

    ' Replace any numeric placeholders with their real values
    tableBookmark = g_counters.UpdatePredicates(tableBookmark)

    ' See if the bookmark exists
    If g_assessmentReport.bookmarks.Exists(tableBookmark) Then

        ' Get a reference to the bookmark range
        Set bookmarkRange = g_assessmentReport.bookmarks(tableBookmark).Range

        ' Now see if it contains a table
        If bookmarkRange.Tables.Count > 0 Then
            Set theTable = bookmarkRange.Tables(1)

        Else
            BookmarkIsExpectedToContainATableError tableBookmark, c_proc
        End If
    Else
        BookmarkDoesNotExistError tableBookmark, c_proc
    End If

    ' The Row (to be deleted) is either a numeric value or a named counter
    If LenB(m_rowCounter) > 0 Then
        deleteRow = g_actionCounters.Item(m_rowCounter)
    Else
        deleteRow = m_row
    End If

    ' Make sure that the Row to be deleted actually exists
    If deleteRow > 1 And deleteRow <= theTable.Rows.Count Then

        ' Remove any editable bookmarks from the editable bookmarks object before the range is deleted (it is too late after!)
        Set deletionRange = theTable.Rows(deleteRow).Range
        g_editableBookmarks.DeleteInRange deletionRange

        ' Delete the specified row from the Exceptions table
        theTable.Rows(deleteRow).Delete
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

Private Sub IAction_UpdateContentControlXML()
    ' Definition required - But nothing to do
End Sub ' IAction_UpdateContentControlXML

Private Sub IAction_UpdateDateXML()
    ' Definition required - But nothing to do
End Sub ' Sub IAction_UpdateDateXML

Friend Property Get RowToDelete() As Long
    If LenB(m_rowCounter) > 0 Then
        RowToDelete = g_actionCounters.Item(m_rowCounter)
    Else
        RowToDelete = m_row
    End If
End Property ' Get RowToDelete

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeDeleteRow
End Property ' Get IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
