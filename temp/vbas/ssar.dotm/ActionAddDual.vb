VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionAddDual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionAddDual
' Purpose:      Data class that models the 'addDual' action node of the 'instructions.xml' file.
'
' Note 1:       The 'addDual' action is basically a loop for inserting Building Blocks into the assessment report.
'
' Note 2:       This class is intended to insert a table, where the table displays two sets of data items per row.
'
' Note 3:       It uses the formula (number of data items returned by the xpath query +1) \ 2 to calculate the number of rows
'               required to hold the number of data items returned by the xpath query.
'
' Note 4:       The following attributes are available for use with this Action:
'               buildingBlock*
'                   The name of the Building Block to insert if the index value is 1.
'               bookmark*
'                   The name of the Bookmark specifying the location to insert 'buildingBlock'.
'               extendBookmark*
'                   True, then 'bookmark' is extended to include the text inserted using 'buildingBlock'.
'                   False, then 'bookmark' stays as it was before the 'add'.
'                   [Default = True]
'               where*
'                   Determines where the building block is inserted relative to 'bookmark'. Values are:
'                   AfterLastParagraph
'                       Specifies create a new paragraph to insert 'buildingBlock' in at the end of the range specified by
'                       'bookmark'.
'                   AtEndOfRange
'                       Specifies insert 'buildingBlock' at the end of the range specified by 'bookmark'.
'               buildingBlockN+
'                   The name of the Building Block to insert if the index value is greater than 1.
'               bookmarkN+
'                   The name of the Bookmark specifying the location to insert 'buildingBlockN'.
'               extendBookmarkN+
'                   True, then 'bookmarkN' is extended to include the text inserted using 'buildingBlockN'.
'                   False, then 'bookmarkN' stays as it was before the 'add'.
'                   [Default = True]
'               whereN+
'                   Determines where the building block is inserted relative to 'bookmarkN'. Values are:
'                   AfterLastParagraph
'                       Specifies create a new paragraph to insert 'buildingBlockN' in at the end of the range specified by
'                       'bookmarkN'.
'                   AtEndOfRange
'                       Specifies insert 'buildingBlockN' at the end of the range specified by 'bookmarkN'.
'               indexAdjustment [Optional]
'                   This is a counter name. The counters value is added to the internal index numbers generated from the number of
'                   nodes returned by the query.
'               deleteIfNull [Optional]
'                   The name of a Bookmark specifying a range to be delete if query 'test' yields no data.
'               pattern [Optional]
'                   The name of a secondary Bookmark created to mark the text inserted using either 'buildingBlock' or
'                   'buildingBlockN'.
'               break [Optional]
'                   True, then code breaks before execution of the 'add' instruction.
'                   [Default = False]
'               test [Mandatory]
'                   An xml xpath query that yields one or more nodes. A Building Block is inserted for each occurrence of a node.
'                   The first Building Block inserted (index = 1) uses 'buildingBlock', the second and subsequent Building Blocks
'                   inserted use 'buildingBlockN'.
'
'               Note *: If 'buildingBlock' is specified, 'bookmark' and 'where' must be specified, 'extendBookmark' is optional and
'                       defaults to True.
'               Note +: If 'buildingBlockN' is specified, 'bookmarkN' and 'whereN' must be specified, 'extendBookmarkN' is optional
'                       and defaults to True.
'
'               Notes 2a:   When the 'buildingBlock' attribute (and its associated attributes) is specified, 'buildingBlockN' (and
'                           its associated attributes) is always specified.
'               Notes 2b:   When only the 'buildingBlockN' attribute (and its associated attributes) is specified, no Add occurs
'                           for the first returned data element.
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
' History:      29/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBookmark                 As String = "bookmark"
Private Const mc_ANBookmarkN                As String = "bookmarkN"
Private Const mc_ANBreak                    As String = "break"
Private Const mc_ANBuildingBlock            As String = "buildingBlock"
Private Const mc_ANBuildingBlockN           As String = "buildingBlockN"
Private Const mc_ANDeleteIfNull             As String = "deleteIfNull"
Private Const mc_ANExtendBookmark           As String = "extendBookmark"
Private Const mc_ANExtendBookmarkN          As String = "extendBookmarkN"
Private Const mc_ANIndexAdjustment          As String = "indexAdjustment"
Private Const mc_ANPattern                  As String = "pattern"
Private Const mc_ANTest                     As String = "test"
Private Const mc_ANWhere                    As String = "where"
Private Const mc_ANWhereN                   As String = "whereN"

Private Enum mc_flags
    mc_flagBuildingBlock = 1
    mc_flagBookmark = 2
    mc_flagWhere = 4
    mc_flagAll = mc_flagBuildingBlock Or mc_flagBookmark Or mc_flagWhere
End Enum


Private m_bookmark                      As String
Private m_bookmarkN                     As String
Private m_break                         As Boolean              ' Used to cause the code to execute a Stop instruction
Private m_buildingBlock                 As String
Private m_buildingBlockN                As String
Private m_deleteIfNullBookmark          As String
Private m_extendBookmark                As Boolean
Private m_extendBookmarkN               As Boolean
Private m_indexAdjustmentCounterName    As String
Private m_pattern                       As String
Private m_subActions                    As Actions              ' Nested Actions in the order they occur.
Private m_test                          As String
Private m_where                         As ssarWhereType
Private m_whereN                        As ssarWhereType


Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionAddDual.IAction_Parse"

    Dim flag            As Long
    Dim flagN           As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Set default values for these attributes
    m_extendBookmark = True
    m_extendBookmarkN = True

    ' Parse out the 'add' nodes attributes
    For Each theAttribute In actionNode.Attributes
    
        Select Case theAttribute.BaseName
        Case mc_ANBookmark
            m_bookmark = theAttribute.Text
            flag = flag Or mc_flagBookmark

        Case mc_ANBookmarkN
            m_bookmarkN = theAttribute.Text
            flagN = flagN Or mc_flagBookmark

        Case mc_ANBuildingBlock
            m_buildingBlock = theAttribute.Text
            flag = flag Or mc_flagBuildingBlock

        Case mc_ANBuildingBlockN
            m_buildingBlockN = theAttribute.Text
            flagN = flagN Or mc_flagBuildingBlock

        Case mc_ANWhere
            m_where = ParseAttributeWhere(actionNode, theAttribute)
            flag = flag Or mc_flagWhere

        Case mc_ANWhereN
            m_whereN = ParseAttributeWhere(actionNode, theAttribute)
            flagN = flagN Or mc_flagWhere

        Case mc_ANTest
            m_test = theAttribute.Text

        Case mc_ANDeleteIfNull
            m_deleteIfNullBookmark = theAttribute.Text

        Case mc_ANPattern
            m_pattern = theAttribute.Text

        Case mc_ANIndexAdjustment
            m_indexAdjustmentCounterName = theAttribute.Text

        Case mc_ANExtendBookmark
            m_extendBookmark = CBool(theAttribute.Text)

        Case mc_ANExtendBookmarkN
            m_extendBookmarkN = CBool(theAttribute.Text)

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Check that all members are present in mandatory attribute groups
    If flag <> 0 And flag <> mc_flagAll Then
        AddNodeMissingAttributeError actionNode.nodeName, actionNode.nodeTypedValue, c_proc
    End If
    If flagN <> 0 And flagN <> mc_flagAll Then
        AddNodeMissingAttributeNError actionNode.nodeName, actionNode.nodeTypedValue, c_proc
    End If

    ' Parse out all of the Add nodes child nodes (sub-actions)
    Set m_subActions = New Actions
    m_subActions.Parse actionNode, addPatternBookmarks

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

Private Function ParseAttributeWhere(ByRef actionNode As MSXML2.IXMLDOMNode, _
                                     ByRef theAttribute As MSXML2.IXMLDOMNode) As ssarWhereType
    Const c_proc As String = "ActionAddDual.ParseAttributeWhere"

    Select Case theAttribute.Text
    Case ssarWhereTextAfterLastParagraph
        ParseAttributeWhere = ssarWhereTypeAfterLastParagraph

    Case ssarWhereTextAtEndOfRange
        ParseAttributeWhere = ssarWhereTypeAtEndOfRange

    Case ssarWhereTextReplaceRange
        ParseAttributeWhere = ssarWhereTypeReplaceRange

    Case Else
        InvalidAttributeWhereValueError actionNode.nodeName, theAttribute.BaseName, theAttribute.Text, c_proc
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ParseAttributeWhere

'=======================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Inserts Building Blocks while Building the Assessment Report (BAR).
' Note 1:       This is designed to insert a table where data is presented in two groups of columns. To this end it
'               maintains a second set of counters (used for predicate replacement) for the right hand side group of
'               columns.
'=======================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionAddDual.IAction_BuildAssessmentReport"

    Dim bbName          As String
    Dim bbWhere         As ssarWhereType
    Dim bookmarkExtend  As Boolean
    Dim bookmarkName    As String
    Dim index           As Long
    Dim indexAdjustment As Long
    Dim rowsRequired    As Long
    Dim theQuery        As String
    Dim subNodes        As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Set up the secondary counter used by the data items inserted into the right hand side of the table.
    ' The normal counters are used for data items inserted into the left hand side of the table.

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_test)

    ' See if index adjustment is required (allows for the nodes retrieved by the query not being the first nodes
    ' in the sequence required). So if there are precceding nodes of the same type they can be skipped over.
    If LenB(m_indexAdjustmentCounterName) > 0 Then

        ' Get the index adjustment value from the named counter
        indexAdjustment = g_actionCounters.Item(m_indexAdjustmentCounterName)
    End If

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Use the ActionAdd objects 'test' string as an xpath query to retrieve all matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Calculate the actual number of rows required using two data item for each row
        rowsRequired = (subNodes.Length + 1) \ 2

        ' Update the counter object so that it generates the correct numbers for the right hand side data set items
        g_counters.Offset = rowsRequired

        ' Iterate each pair of retrieved nodes
        For index = 1 To rowsRequired

            ' Increment the counter for the current level. Always include the index adjustment as it has a zero value if unused.
            g_counters.Counter = index + indexAdjustment

            ' If this Section uses more than one Building Block, choose the Building Block
            ' based on this sections counter. This allows a different Building Block to be
            ' used the first time a Building Block is added to the document.
            If index = 1 Then
                bbName = m_buildingBlock
                bbWhere = m_where
                bookmarkName = m_bookmark
                bookmarkExtend = m_extendBookmark
            Else
                bbName = m_buildingBlockN
                bbWhere = m_whereN
                bookmarkName = m_bookmarkN
                bookmarkExtend = m_extendBookmarkN
            End If

            ' Insert the selected Building Block
            InsertBuildingBlock bbName, bbWhere, bookmarkName, m_pattern, index, bookmarkExtend

            ' Perform any nested actions in the current ActionAddDual object - but not when we delete the block!
            m_subActions.BuildAssessmentReport
        Next
    Else

        ' No nodes were returned by the query, so check to see if there is a bookmark that should be
        ' deleted. The bookmark may in turn contain other bookmarks. This provides a mechanism for
        ' deleting corresponding boiler plate text when there is no data for associated bookmarks.

        ' Delete the bookmarked block of text
        DeleteUnusedTextBlock m_deleteIfNullBookmark
    End If

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.Offset = 0
    g_counters.ResetCounter
    g_counters.DecrementDepth

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

Friend Property Get BookmarkN() As String
    BookmarkN = m_bookmarkN
End Property ' Get BookmarkN

Friend Property Get BuildingBlock() As String
    BuildingBlock = m_buildingBlock
End Property ' Get BuildingBlock

Friend Property Get BuildingBlockN() As String
    BuildingBlockN = m_buildingBlockN
End Property ' Get BuildingBlockN

Friend Property Get DeleteIfNullBookmark() As String
    DeleteIfNullBookmark = m_deleteIfNullBookmark
End Property '  Get DeleteIfNullBookmark

Friend Property Get ExtendBookmark() As Boolean
    ExtendBookmark = m_extendBookmark
End Property ' ExtendBookmark

Friend Property Get ExtendBookmarkN() As Boolean
    ExtendBookmarkN = m_extendBookmarkN
End Property ' ExtendBookmarkN

Friend Property Get IndexAdjustmentCounterName() As String
    IndexAdjustmentCounterName = m_indexAdjustmentCounterName
End Property ' Get IndexAdjustmentCounterName

Friend Property Get Pattern() As String
    Pattern = m_pattern
End Property ' Get Pattern

Friend Property Get SubActions() As Actions
    Set SubActions = m_subActions
End Property ' Get SubActions

Friend Property Get Test() As String
    Test = m_test
End Property ' Get test

Friend Property Get Where() As ssarWhereType
    Where = m_where
End Property ' Get Where

Friend Property Get WhereN() As ssarWhereType
    WhereN = m_whereN
End Property ' Get WhereN

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeCopy
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
