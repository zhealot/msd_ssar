VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ActionAdd
' Purpose:      Data class that models the 'add' action node-tree of the 'instructions.xml' file.
' Note 1:       The 'add' action is basically a loop for inserting Building Blocks into the assessment report.
'
' Note 2:       The <add></add> block contains further actions to insert data into the Building Block being added.
'
' Note 3:       The following attributes are available for use with this Action:
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
'               deleteIfNull
'                   The name of a Bookmark specifying a range to be delete if query 'test' yields no data.
'               pattern
'                   The name of a secondary Bookmark created to mark the text inserted using either 'buildingBlock' or
'                   'buildingBlockN'.
'               refresh
'                   True, then causes the range marked by 'refreshBookmark' to be deleted, then process the Add action block to
'                   rebuild the document content for the deleted block of text.
'               refreshDeleteBM
'                   A bookmark whose range should be deleted prior to doing the fresh.
'               refreshQuery
'                   The xpath query that determines which nodes should have their contents updated to match the corresponding
'                   bookmark specified by ''.
'               refreshTargetBM
'                   The bookmark patern of the bookmarks matching the xpath query. For each bookmark the corresponding xml element
'                   will be updated. This step controls what is actually selected by the 'test' xpath query.
'               refreshBMData
'                   The xpath query that determines which node value from the assessment report xml will be used as part of a
'                   bookmark name to generate a unique bookmark name.
'               break [Optional]
'                   True, then code breaks before execution of the 'add' instruction.
'                   [Default = False]
'               test
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
' History:      03/06/16    1.  Moved to this AddIn as part of the SSAR development.
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
Private Const mc_ANPattern                  As String = "pattern"
Private Const mc_ANRefresh                  As String = "refresh"
Private Const mc_ANRefreshBMData            As String = "refreshBMData"
Private Const mc_ANrefreshDeleteBM          As String = "refreshDeleteBM"
Private Const mc_ANrefreshQuery             As String = "refreshQuery"
Private Const mc_ANrefreshTargetBM          As String = "refreshTargetBM"
Private Const mc_ANTest                     As String = "test"
Private Const mc_ANWhere                    As String = "where"
Private Const mc_ANWhereN                   As String = "whereN"


Private Enum mc_flags
    mc_flagBuildingBlock = 1
    mc_flagBookmark = 2
    mc_flagWhere = 4
    mc_flagAll = mc_flagBuildingBlock Or mc_flagBookmark Or mc_flagWhere
End Enum


Private m_bookmark              As String
Private m_bookmarkN             As String
Private m_break                 As Boolean              ' Used to cause the code to execute a Stop instruction
Private m_buildingBlock         As String
Private m_buildingBlockN        As String
Private m_deleteIfNullBookmark  As String
Private m_extendBookmark        As Boolean
Private m_extendBookmarkN       As Boolean
Private m_pattern               As String
Private m_refresh               As Boolean              ' Refresh an area of the asessment report
Private m_refreshDeleteBM       As String               ' The area of the assessment report to be deleted before the refresh
Private m_refreshQuery          As String               ' The xpath query to determine which xml nodes are to have a dirty update
Private m_refreshTargetBM       As String               ' The bookmark patern of the bookmarks matching the xpath query.
                                                        ' For each bookmark the corresponding xml element will be updated.
Private m_refreshBMData         As String               ' The Xpath query for the 'refreshTargetBM' parameter replacement value.
                                                        ' This is used to generate a unique bookmark name based on the contents of
                                                        ' an assessment report node.
Private m_subActions            As Actions              ' Nested Actions in the order they occur.
Private m_test                  As String
Private m_where                 As ssarWhereType
Private m_whereN                As ssarWhereType


Private Sub Class_Terminate()
    Set m_subActions = Nothing
End Sub ' Class_Terminate

'===================================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Add action instruction.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'===================================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionAdd.IAction_Parse"

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

        Case mc_ANExtendBookmark
            m_extendBookmark = CBool(theAttribute.Text)

        Case mc_ANExtendBookmarkN
            m_extendBookmarkN = CBool(theAttribute.Text)

        Case mc_ANRefresh
            m_refresh = CBool(theAttribute.Text)

        Case mc_ANrefreshDeleteBM
            m_refreshDeleteBM = theAttribute.Text

        Case mc_ANrefreshQuery
            m_refreshQuery = theAttribute.Text

        Case mc_ANrefreshTargetBM
            m_refreshTargetBM = theAttribute.Text

        Case mc_ANRefreshBMData
            m_refreshBMData = theAttribute.Text

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
    Const c_proc As String = "ActionAdd.ParseAttributeWhere"

    On Error GoTo Do_Error

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

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Inserts Building Blocks while Building the Assessment Report (BAR).
' Notes:        This procedure is indirectly recursive.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionAdd.IAction_BuildAssessmentReport"

    Dim bbName          As String
    Dim bbWhere         As ssarWhereType
    Dim bookmarkExtend  As Boolean
    Dim bookmarkName    As String
    Dim index           As Long
    Dim theQuery        As String
    Dim subNode         As MSXML2.IXMLDOMNode
    Dim subNodes        As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_test)

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Use the ActionAdd objects 'test' string as an xpath query to retrieve all matching nodes
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

                ' Insert the specified Building Block
                InsertBuildingBlock bbName, bbWhere, bookmarkName, m_pattern, index, bookmarkExtend

                ' Perform any nested actions in the current ActionAdd object - but not when we delete the block!
                m_subActions.BuildAssessmentReport
            Else

                ' Delete the unused text block if it is present
                DeleteUnusedTextBlock m_deleteIfNullBookmark
            End If
        Next
    Else

        ' No nodes were returned by the query, so check to see if there is a bookmark that should be
        ' deleted. The bookmark may in turn contain other bookmarks. This provides a mechanism for
        ' deleting corresponding boiler plate text when there is no data for associated bookmarks.

        ' Delete the bookmarked block of text
        DeleteUnusedTextBlock m_deleteIfNullBookmark
    End If

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

'===================================================================================================================================
' Procedure:    IAction_ConstructRichText
' Purpose:      Processes Add actions while building the Rich Text data source.
' Note 1:       The only thing really happening here is setting up the counters and carring out all sub-actions.
'===================================================================================================================================
Private Sub IAction_ConstructRichText()
    Const c_proc As String = "ActionAdd.IAction_ConstructRichText"

    Dim index       As Long
    Dim theQuery    As String
    Dim subNodes    As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_test)

    ' Use the ActionAdd objects Test string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            m_subActions.ConstructRichText
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
End Sub ' IAction_ConstructRichText

'===================================================================================================================================
' Procedure:    IAction_HTMLForXMLUpdate
' Purpose:      Processes Add actions while building the HTML to XML data source.
' Note 1:       The only thing really happening here is setting up the counters and carring out all sub-actions.
' Date:         05/08/16    Created.
'===================================================================================================================================
Private Sub IAction_HTMLForXMLUpdate()
    Const c_proc As String = "ActionAdd.IAction_HTMLForXMLUpdate"

    Dim index       As Long
    Dim theQuery    As String
    Dim subNodes    As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_test)

    ' Use the ActionAdd objects Test string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            m_subActions.HTMLForXMLUpdate
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
End Sub '  Sub IAction_HTMLForXMLUpdate

'===================================================================================================================================
' Procedure:    UpdateContentControlXML
' Purpose:      Processes Add actions while updating the assessment report xml using a value stored in a Content Control.
' Note 1:       This is a bit confusing because the Content Controls have their own xml data store. But this is not updating the
'               Content Control data store, this updates the assessment report xml using the Contents Controls value.
' Date:         16/08/16    Created
'===================================================================================================================================
Private Sub IAction_UpdateContentControlXML()
    Const c_proc As String = "ActionAdd.IAction_UpdateContentControlXML"

    Dim index       As Long
    Dim theQuery    As String
    Dim subNodes    As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_test)

    ' Use the ActionAdd objects Test string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            m_subActions.UpdateContentControlXML
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
End Sub ' IAction_UpdateContentControlXML

'===================================================================================================================================
' Procedure:    IAction_UpdateDateXML
' Purpose:       the corresponding xml of any editable area that contains a date with the:
'               ssarDataFormatDateLong, ssarDataFormatDateShort and ssarDataFormatDateShortYear data types.
'
' Note 1:       Before the xml update occurs the dates are validated to ensure that the user has entered a valid date.
' Note 2:       The only thing really happening here is setting up the counters and carring out all sub-actions.
'
' Date:         04/08/16    Created.
'===================================================================================================================================
Private Sub IAction_UpdateDateXML()
    Const c_proc As String = "ActionAdd.IAction_UpdateDateXML"

    Dim index       As Long
    Dim theQuery    As String
    Dim subNodes    As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_test)

    ' Use the ActionAdd objects Test string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            m_subActions.UpdateDateXML
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

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeAdd
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break

Friend Property Get Pattern() As String
    Pattern = m_pattern
End Property ' Get Pattern

Friend Property Get Refresh() As Boolean
    Refresh = m_refresh
End Property ' Get Refresh

Friend Property Get RefreshBMData() As String
    RefreshBMData = m_refreshBMData
End Property ' Get RefreshBMData

Friend Property Get RefreshDeleteBookmark() As String
    RefreshDeleteBookmark = m_refreshDeleteBM
End Property ' Get RefreshDeleteBookmark

Friend Property Get RefreshQuery() As String
    RefreshQuery = m_refreshQuery
End Property ' Get RefreshQuery

Friend Property Get RefreshTargetBM() As String
    RefreshTargetBM = m_refreshTargetBM
End Property ' Get RefreshTargetBM

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
