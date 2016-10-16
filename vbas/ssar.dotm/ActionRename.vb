VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ActionRename
' Purpose:      Data class that models the 'rename' action node of the 'instructions.xml' file.
'
' Note 1:       The following attributes are available for use with this Action:
'               newName [Mandatory]
'                   The bookmarks new name.
'               oldName [Mandatory]
'                   The bookmarks current name.
'               increment {optional]
'                   A value added to the g_counters counter at the current depth. This attribute is only applied to counter values
'                   used to construct the 'newName' bookmark.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
'
' Note 1:       The purpose of the 'increment' attribute is to be able to automatically rename pattern bookmarks.
'               So a bookmark with a pattern of "KF_s%1_criteria_%2" and a value of "KF_s12_criteria_6" can be automatically renamed
'               to "KF_s12_criteria_7". By using 'increment' with a value of 1. Using 'increment' with a value of -1, would rename
'               "KF_s12_criteria_6" to "KF_s12_criteria_5".
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
' History:      16/11/15    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBreak        As String = "break"
Private Const mc_ANIncrement    As String = "increment"
Private Const mc_ANNewName      As String = "newName"
Private Const mc_ANOldName      As String = "oldName"


Private m_break             As Boolean
Private m_increment         As Long
Private m_newBookmarkName   As String
Private m_oldBookmarkName   As String


'===================================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Rename action instruction.
' Date:         16/11/15    Created.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'===================================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionRename.IAction_Parse"

    On Error GoTo Do_Error

    ' Make sure there are attributes present
    If actionNode.Attributes.Length > 0 Then

        ParseAttributes actionNode, mc_ANNewName, m_newBookmarkName, mc_ANOldName, m_oldBookmarkName, _
                        mc_ANIncrement, m_increment, mc_ANBreak, m_break
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'=======================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Renames a bookmark.
' Notes:        This is generally used to create a unique bookmark name where one does not currently exist or a unique
'               bookmark name is required by a subsequent action block.
'=======================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionRename.IAction_BuildAssessmentReport"

    Dim bookmarkName    As String
    Dim newName         As String
    Dim target          As Word.Range

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    With g_assessmentReport.bookmarks

        ' Update any numeric parameters in the Old Bookmark name with their corresponding values
        bookmarkName = g_counters.UpdatePredicates(m_oldBookmarkName)

        ' Make sure the old bookmark exists before trying to rename it
        If .Exists(bookmarkName) Then

            ' See if the counter for the new bookmark name needs incrementing (for the auto renumbering of bookmark names)
            If m_increment <> 0 Then
                g_counters.Counter = g_counters.Counter + m_increment
            End If

            ' Generate the name for the new bookmark
            newName = g_counters.UpdatePredicates(m_newBookmarkName)

            ' See if the counter should be reset to its original value
            If m_increment <> 0 Then
                g_counters.Counter = g_counters.Counter - m_increment
            End If

            ' Get the old bookmarks Range object so that we can use it to create the new bookmark
            Set target = .Item(bookmarkName).Range

            ' Create a bookmark using the new name
            .Add newName, target

            ' If there is a corresponding editable bookmark rename that as well
            g_editableBookmarks.Rename bookmarkName, newName

            ' Delete the old bookmark name
            .Item(bookmarkName).Delete
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

'=======================================================================================================================
' Procedure:    IAction_ConstructRichText
' Purpose:      Handles the Rename requirements for Rich Text.
'=======================================================================================================================
Private Sub IAction_ConstructRichText()
    ' Definition required - But nothing to do
End Sub ' IAction_ConstructRichText

Private Sub IAction_HTMLForXMLUpdate()
    ' Definition required - But nothing to do
End Sub '  Sub IAction_HTMLForXMLUpdate

Private Sub IAction_UpdateContentControlXML()
    ' Definition required - But nothing to do
End Sub ' IAction_UpdateContentControlXML

'=======================================================================================================================
' Procedure:    IAction_UpdateDateXML
' Purpose:      Handles the Rename requirements for editable area date validation.
'=======================================================================================================================
Private Sub IAction_UpdateDateXML()
    ' Definition required - But nothing to do
End Sub ' Sub IAction_UpdateDateXML

Friend Property Get Increment() As Long
    Increment = m_increment
End Property ' Get Increment

Friend Property Get NewBookmarkName() As String
    NewBookmarkName = m_newBookmarkName
End Property ' NewBookmarkName

Friend Property Get OldBookmarkName() As String
    OldBookmarkName = m_oldBookmarkName
End Property ' OldBookmarkName

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeRename
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
