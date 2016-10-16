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
' Purpose:      Data class that models the 'add' action node-tree of the 'rda instructions.xml' file.
' Note:         The 'add' node can contain nested 'add' nodes.
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
' History:      05/11/15    1.  Created.
'===================================================================================================================================
Option Explicit

' Node Name Consts
Private Const mc_NNAdd                      As String = "add"
Private Const mc_NNInsert                   As String = "insert"
Private Const mc_NNLink                     As String = "link"
Private Const mc_NNRename                   As String = "rename"
Private Const mc_NNSetup                    As String = "setup"

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

' 'where' attribute values
Private Const mc_AVWhere_AfterLastParagraph As String = "AfterLastParagraph"
Private Const mc_AVWhere_AtEndOfRange       As String = "AtEndOfRange"
Private Const mc_AVWhere_ReplaceRange       As String = "ReplaceRange"


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
Private m_subActions            As Collection           ' Nested "adds" and "inserts" in the order they occur.
Private m_test                  As String
Private m_where                 As rdaWhere
Private m_whereN                As rdaWhere


Private Sub Class_Terminate()
    Set m_subActions = Nothing
End Sub ' Class_Terminate

'=======================================================================================================================
' Procedure:    Parse
' Purpose:      .
' Notes:        .
'
' On Entry:     addNode             .
'               addPattenBookmarks  Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Friend Sub Parse(ByRef addNode As MSXML2.IXMLDOMNode, _
                 Optional ByVal addPattenBookmarks As Boolean)
    Const c_proc As String = "ActionAdd.Parse"

    Dim child           As MSXML2.IXMLDOMNode
    Dim errorText       As String
    Dim flag            As Long
    Dim flagN           As Long
    Dim newAdd          As ActionAdd
    Dim newInsert       As ActionInsert
    Dim newLink         As ActionLink
    Dim newRename       As ActionRename
    Dim newSetup        As ActionSetup
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Initialise the SubActions (nested 'add, insert, rename and setup' actions) collection object
    Set m_subActions = New VBA.Collection

    ' Set default values for these attributes
    m_extendBookmark = True
    m_extendBookmarkN = True

    ' Parse out the 'add' nodes attributes
    For Each theAttribute In addNode.Attributes
    
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
            m_where = ParseAttributeWhere(theAttribute)
            flag = flag Or mc_flagWhere

        Case mc_ANWhereN
            m_whereN = ParseAttributeWhere(theAttribute)
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
            errorText = Replace$(mgrErrTextInvalidAttributeName, mgrP1, addNode.nodeName)
            errorText = Replace$(errorText, mgrP2, theAttribute.BaseName)
            Err.Raise mgrErrNoInvalidAttributeName, c_proc, errorText
        End Select
    Next

    ' Check that all members are present in mandatory attribute groups
    If flag <> 0 And flag <> mc_flagAll Then
        errorText = Replace$(mgrErrTextAddNodeMissingAttribute, mgrP1, addNode.nodeTypedValue)
        Err.Raise mgrErrNoAddNodeMissingAttribute, c_proc, errorText
    End If
    If flagN <> 0 And flagN <> mc_flagAll Then
        errorText = Replace$(mgrErrTextAddNodeMissingAttributeN, mgrP1, addNode.nodeTypedValue)
        Err.Raise mgrErrNoAddNodeMissingAttributeN, c_proc, errorText
    End If

    ' If this Add action has a 'refresh' = True attribute add the current
    ' action to the global collection object for Add actions with Refresh
    If m_refresh Then
        If g_addsWithRefresh Is Nothing Then
            Set g_addsWithRefresh = New Collection
        End If

        ' Add this Add action node to the collection object so that we have
        ' an easy way to get to it without walking the entire Actions tree
        g_addsWithRefresh.Add Me
    End If

    ' Parse out root 'add' nodes child nodes
    For Each child In addNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse out the child nodes
            Select Case child.BaseName
            Case mc_NNAdd

                ' Create a new ActionAdd object
                Set newAdd = New ActionAdd

                ' Parse out this 'add' node
                newAdd.Parse child, addPattenBookmarks

                ' Add the new 'add' node to the collection object
                m_subActions.Add newAdd

            Case mc_NNInsert

                ' Create a new ActionInsert object
                Set newInsert = New ActionInsert

                ' Parse out the 'insert' node
                newInsert.Parse child, addPattenBookmarks

                ' Add it to the collection object
                m_subActions.Add newInsert

            Case mc_NNLink

                ' Create a new ActionLink object
                Set newLink = New ActionLink

                ' Parse out the 'link' node
                newLink.Parse child

                ' Add the new 'link' node to the collection object
                m_subActions.Add newLink

            Case mc_NNRename

                ' Create a new ActionRename object
                Set newRename = New ActionRename

                ' Parse out this 'rename' node
                newRename.Parse child

                ' Add the new 'rename' node to the collection object
                m_subActions.Add newRename

            Case mc_NNSetup

                ' Create a new ActionSetup object
                Set newSetup = New ActionSetup

                ' Parse out this 'setup' node
                newSetup.Parse child

                ' Add the new 'rename' node to the collection object
                m_subActions.Add newSetup

            Case Else
                errorText = Replace$(mgrErrTextInvalidNodeName, mgrP1, child.BaseName)
                Err.Raise mgrErrNoInvalidNodeName, c_proc, errorText
            End Select
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

Private Function ParseAttributeWhere(ByRef theAttribute As MSXML2.IXMLDOMNode) As rdaWhere
    Const c_proc As String = "ActionAdd.ParseAttributeWhere"

    Dim errorText As String

    On Error GoTo Do_Error

    Select Case theAttribute.Text
    Case mc_AVWhere_AfterLastParagraph
        ParseAttributeWhere = rdaWhereAfterLastParagraph

    Case mc_AVWhere_AtEndOfRange
        ParseAttributeWhere = rdaWhereAtEndOfRange

    Case mc_AVWhere_ReplaceRange
        ParseAttributeWhere = rdaWhereReplaceRange

    Case Else
        errorText = Replace$(mgrErrTextInvalidAttributeValueExtended, mgrP1, theAttribute.parentNode.BaseName)
        errorText = Replace$(errorText, mgrP2, theAttribute.BaseName)
        errorText = Replace$(errorText, mgrP3, "'AtEndOfRange', 'AfterLastParagraph' or 'ReplaceRange'")
        errorText = Replace$(errorText, mgrP4, theAttribute.Text)
        Err.Raise mgrErrNoInvalidAttributeValue, c_proc, errorText
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ParseAttributeWhere

Friend Property Get Bookmark() As String
    Bookmark = m_bookmark
End Property ' Get Bookmark

Friend Property Get BookmarkN() As String
    BookmarkN = m_bookmarkN
End Property ' Get BookmarkN

Friend Property Get Break() As Boolean
    Break = m_break
End Property ' Break

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

Friend Property Get SubActions() As Collection
    Set SubActions = m_subActions
End Property ' Get SubActions

Friend Property Get Test() As String
    Test = m_test
End Property ' Get test

Friend Property Get Where() As rdaWhere
    Where = m_where
End Property ' Get Where

Friend Property Get WhereN() As rdaWhere
    WhereN = m_whereN
End Property ' Get WhereN
