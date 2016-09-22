VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionLink
' Purpose:      Data class that models the 'link' action node of the 'instructions.xml' file.
'
' Note 1:       Creates a Ref field to reference a piece of bookmarked data.
'
' Note 2:       The following attributes are available for use with this Action:
'               bookmark
'                   .
'               bookmarkPattern
'                   .
'               patternData [Optional]
'                   .
'               source
'                   .
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
' History:      22/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBookmark     As String = "bookmark"          ' Initial bookmark name
Private Const mc_ANBreak        As String = "break"             ' Breakpoint debug aid,
Private Const mc_ANPattern      As String = "pattern"           ' Used to create a unique bookmark name
Private Const mc_ANPatternData  As String = "patternData"       ' The xpath query for the pattern parameter replacement value
Private Const mc_ANSource       As String = "source"            ' The source bookmark name used to update 'pattern' bookmark contents


Private m_bookmark          As String
Private m_bookmarkPattern   As String
Private m_break             As Boolean
Private m_patternData       As String
Private m_source            As String


'===================================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Link action instruction.
' Date:         29/06/16    Created.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'===================================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionLink.IAction_Parse"

    On Error GoTo Do_Error

    ' Make sure there are attributes present
    If actionNode.Attributes.Length > 0 Then

        ParseAttributes actionNode, mc_ANBookmark, m_bookmark, mc_ANPattern, m_bookmarkPattern, mc_ANPatternData, m_patternData, _
                                  mc_ANSource, m_source, mc_ANBreak, m_break
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'=======================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Sets up the Link Action while Building or Refreshing the Assessment Report.
' Notes 1:      The Link Action create a Ref field that references the actual data at a bookmarked location.
' Notes 2:      However, there is a FWB we have to contend with, because the data source (the area referred to by the
'               Bookmark in the Ref Field) has an Editor (it's an editable area of a protected document) for some
'               bizzare reason Word make the Ref Field editable as well!!
' Note 3:       Because we do not want the Ref Fields being editable we have to update them in a very specific manner.
' Note 4:       At the time we insert the Ref the data source Bookmark may not exist, causing the Ref Field to display
'               an error until it is updated when the Assessment Report build process is complete.
'=======================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionLink.IAction_BuildAssessmentReport"

    Dim dataSourceBM    As String
    Dim newBM           As String
    Dim newField        As Word.Field
    Dim rawName         As String
    Dim target          As Word.Range
    Dim targetBM        As String

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    With g_assessmentReport.bookmarks

        ' Make sure the target bookmark (the location where the Ref Field will be inserted) exists
        targetBM = g_counters.UpdatePredicates(m_bookmark)
        If .Exists(targetBM) Then

            ' Create the name for pattern Bookmark (used to give a repeating bookmark a unique name)
            If LenB(m_bookmarkPattern) > 0 Then
                newBM = g_counters.UpdatePredicates(m_bookmarkPattern)
            End If

            ' Create the name of the data source Bookmark
            rawName = ReplacePatternData(m_source, m_patternData)
            dataSourceBM = g_counters.UpdatePredicates(rawName)

            ' Get the target bookmarks Range object so that we can use it to create the new bookmark
            Set target = .Item(targetBM).Range

            ' Insert the Ref Field
            Set newField = g_assessmentReport.Content.Fields.Add(target, wdFieldRef, dataSourceBM, False)

            ' Use the unique new Bookmark name to create a Bookmark for the Ref Field so that we can find it again
            If LenB(newBM) > 0 Then
                .Add newBM, newField.result
            End If

            ' FWB: Delete the unwanted Editor object added by Word to the Field when it was updated!
            If newField.result.Editors.Count > 0 Then
                newField.result.Editors(1).Delete
            End If

            ' CRITICAL: This field MUST be locked or else Word will auto-update it when you print an assessment and that in turn
            ' screws up our Editable Bookmark Object. Because when you rebuild the Editable Bookmarks the document will contain
            ' editable areas that we are not tracking bookmarks for.
            newField.Locked = True

            ' #FWB# Always force the fields ShowCodes to False or when the field is updated if the fields target
            ' range is collapsed (just an insertion point) Word displays the field code instead of the field results!
            newField.ShowCodes = False
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

Friend Property Get Bookmark() As String
    Bookmark = m_bookmark
End Property ' Get Bookmark

Friend Property Get bookmarkPattern() As String
    bookmarkPattern = m_bookmarkPattern
End Property 'Get BookmarkPattern

Friend Property Get PatternData() As String
    PatternData = m_patternData
End Property ' PatternData

Friend Property Get Source() As String
    Source = m_source
End Property ' Get Source

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeLink
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
