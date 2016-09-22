VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionAddRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionAddRow
' Purpose:      Data class that models the 'addRow' action node-tree of the 'instructions.xml' file.
'
' Note 1:       Adds a row stored as a Builidng Block to a table.
'
' Note 2:       The following attributes are available for use with this Action:
'               buildingBlock [Mandatory]
'                   The name of the Building Block that contains the row to be inserted.
'                   Each cell must contain a bookmark if the cell is to be filled in.
'               insertAfterRow [Mandatory]
'                   The table row after which the new row will be inserted.
'                   This can be either a numeric value or the name of a named counter.
'               tableBookmark   [Mandatory]
'                   The name of a bookmark that wholely or partially contains the table the new row will be inserted in.
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
' History:      19/07/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBuildingBlock            As String = "buildingBlock"             ' The name of a Building Block that contains the table row to be inserted
Private Const mc_ANBreak                    As String = "break"
Private Const mc_ANInsertAfterRow           As String = "insertAfterRow"            ' The number of the row in table the new row will be inserted after
Private Const mc_ANTableBookmark            As String = "tableBookmark"             ' The name of a bookmark that wholely or partially contains the table the new row will be inserted in


Private m_break                 As Boolean              ' Used to cause the code to execute a Stop instruction
Private m_buildingBlock         As String
Private m_insertAfterRow        As Long
Private m_insertAfterRowCounter As String               ' The name of the named counter
Private m_tableBookmark         As String


'===================================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the AddRow action instruction.
' Date:         19/07/16    Created.
'
' On Entry:     actionNode          The xml node containing the addRow instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'===================================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionAddRow.IAction_Parse"

    Dim buildingBlockCount  As Long
    Dim insertAfterRowCount As Long
    Dim tableBookmarkCount  As Long
    Dim theAttribute        As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out the 'addRow' nodes attributes
    For Each theAttribute In actionNode.Attributes
    
        Select Case theAttribute.BaseName
        Case mc_ANBuildingBlock
            buildingBlockCount = buildingBlockCount + 1
            m_buildingBlock = theAttribute.Text

        Case mc_ANTableBookmark
            tableBookmarkCount = tableBookmarkCount + 1
            m_tableBookmark = theAttribute.Text

        Case mc_ANInsertAfterRow
            insertAfterRowCount = insertAfterRowCount + 1

            ' If value is non numeric then it must be a named counter
            If IsNumeric(theAttribute.Text) Then
                m_insertAfterRow = CLng(theAttribute.Text)
            Else
                m_insertAfterRowCounter = Trim$(theAttribute.Text)
                If LenB(m_insertAfterRowCounter) = 0 Then
                    InvalidAttributeValueError actionNode.nodeName, mc_ANInsertAfterRow, m_insertAfterRowCounter, c_proc
                End If
            End If

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure the Mandatory attributes are present
    If buildingBlockCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANBuildingBlock, c_proc
    ElseIf tableBookmarkCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANTableBookmark, c_proc
    ElseIf insertAfterRowCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANInsertAfterRow, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If buildingBlockCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANBuildingBlock, c_proc
    ElseIf tableBookmarkCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANTableBookmark, c_proc
    ElseIf insertAfterRowCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANInsertAfterRow, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'=======================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Inserts a Building Block that contains a single Row somewhere in a table.
' Note 1:       This code is aimed at updating a table in the assessment report in response to a UI action.
'=======================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionAddRow.IAction_BuildAssessmentReport"

    Dim bookmarkRange       As Word.Range
    Dim insertAfterRow      As Long
    Dim insertAfterRange    As Word.Range
    Dim originalSelection   As Word.Range
    Dim tableBookmark       As String
    Dim theTable            As Word.Table
    Dim whereInserted       As Word.Range

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

    ' The insertAfterRow is either a numeric value or a named counter
    If LenB(m_insertAfterRowCounter) > 0 Then
        insertAfterRow = g_actionCounters.Item(m_insertAfterRowCounter)
    Else
        insertAfterRow = m_insertAfterRow
    End If

    ' Make sure the specified insertAfterRow Row exists
    If insertAfterRow <= 0 Or insertAfterRow > theTable.Rows.Count Then
        InvalidInsertAfterRowNumberError insertAfterRow, theTable.Rows.Count, c_proc
    End If

    ' The row we need to insert the new row after
    Set insertAfterRange = theTable.Rows(insertAfterRow).Range
    insertAfterRange.Collapse wdCollapseEnd

    ' If we are inserting after the last row in the table we do not need to split the table
    If insertAfterRow < theTable.Rows.Count Then

        ' We need to split the table so that bookmarks at the begining of the first cell don't get screwed up
        Set originalSelection = Selection.Range.Duplicate
        insertAfterRange.Select
        Selection.SplitTable
    End If

    ' Insert the new row into the table
    Set whereInserted = InsertBuildingBlock2(m_buildingBlock, insertAfterRange)

    If Not originalSelection Is Nothing Then

        ' Now delete the paragraph mark added by SplitTable
        whereInserted.Collapse wdCollapseEnd
        whereInserted.Delete

        ' Restore the original selection object
        originalSelection.Select
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

Friend Property Get InsertAfterLastRow() As Long
    If LenB(m_insertAfterRowCounter) > 0 Then
        InsertAfterLastRow = g_actionCounters.Item(m_insertAfterRowCounter)
    Else
        InsertAfterLastRow = m_insertAfterRow
    End If
End Property ' Get InsertAfterLastRow

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeAddRow
End Property ' Get IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
