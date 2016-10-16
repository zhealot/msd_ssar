VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        UserInterface
' Purpose:      Data class that models the 'userInterface' action node-tree of the 'instructions.xml' file.
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
' History:      30/05/16    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit

' NN = Node Name
Private Const mc_NNActionSet                    As String = "actionSet"
Private Const mc_NNBookmarks                    As String = "bookmarks"
Private Const mc_NNNewXML                       As String = "newXML"
Private Const mc_NNQueries                      As String = "queries"

' ANAS = Attribute Names - Action Set
Private Const mc_ANASAddExceptionsRow           As String = "addExceptionsRow"
Private Const mc_ANASAddExceptionsTable         As String = "addExceptionsTable"
Private Const mc_ANASDeleteExceptionsRow        As String = "deleteExceptionsRow"
Private Const mc_ANASDeleteExceptionsTable      As String = "deleteExceptionsTable"

' ANB = Attribute Names - Bookmarks
Private Const mc_ANBExceptionTable              As String = "exceptionTable"

' ANN = Attribute Names - New XML
Private Const mc_ANNNewException                As String = "newException"

' ANQ = Attribute Names - Queries
Private Const mc_ANQAllFindings                 As String = "allFindings"
Private Const mc_ANQBuildDropDown               As String = "buildDropDown"
Private Const mc_ANQDeleteFindingsXML           As String = "deleteFindingsXML"
Private Const mc_ANQNextSiblingDifferentTag     As String = "nextSiblingDifferentTag"
Private Const mc_ANQParentNodeForAddDelete      As String = "parentNodeForAddDelete"
Private Const mc_ANQTickedRecommendationButtons As String = "tickedRecommendationButtons"
Private Const mc_ANQTickedStrengthButtons       As String = "tickedStrengthButtons"
Private Const mc_ANQNumberOfToggleButtons       As String = "numberOfToggleButtons"


' ActionSet variables
Private m_asAddExceptionsRow                As Actions
Private m_asAddExceptionsTable              As Actions
Private m_asDeleteExceptionsRow             As Actions
Private m_asDeleteExceptionsTable           As Actions

Private m_addContent                        As Actions

Private m_bookmarkExceptionsTable           As String

Private m_deleteContentItemBookmark         As String
Private m_deleteContentRenameBookmarks      As String
Private m_deleteContentUnusedBlockBookmark  As String

Private m_newExceptionXML                   As String

' XPath query variables
Private m_queryAllFindings                  As String
Private m_queryBuildDropDown                As String
Private m_queryDeleteFindingsXML            As String
Private m_queryNextSiblingDifferentTag      As String
Private m_queryParentNodeForAddDelete       As String
Private m_queryTickedStrengthButtons        As String
Private m_queryTickedRecommendationButtons  As String
Private m_queryNumberOfToggleButtons        As String


Friend Sub Parse(ByVal userInterfaceNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.Parse"

    Dim child   As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out each child node ('queries', 'addContent' or 'deleteContent')
    For Each child In userInterfaceNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse out each child node
            Select Case child.BaseName
            Case mc_NNQueries
                ParseQueriesNode child

            Case mc_NNBookmarks
                ParseBookmarksNode child

            Case mc_NNNewXML
                ParseNewXMLNode child

            Case mc_NNActionSet
                ParseActionSetNode child

            Case Else
                InvalidNodeNameError child.BaseName, c_proc
            End Select
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

Private Sub ParseActionSetNode(ByVal actionSetNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.ParseActionSetNode"

    Dim addExceptionsRowCount       As Long
    Dim addExceptionsTableCount     As Long
    Dim child                       As MSXML2.IXMLDOMNode
    Dim deleteExceptionsRowCount    As Long
    Dim deleteExceptionsTableCount  As Long

    On Error GoTo Do_Error

    ' Parse the 'actionSet' node
    For Each child In actionSetNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            Select Case child.BaseName
            Case mc_ANASAddExceptionsRow
                addExceptionsRowCount = addExceptionsRowCount + 1
                If addExceptionsRowCount = 1 Then
                    Set m_asAddExceptionsRow = New Actions
                    m_asAddExceptionsRow.Parse child
                Else
                    DuplicateNodeError child.nodeName, actionSetNode.nodeName, c_proc
                End If

            Case mc_ANASAddExceptionsTable
                addExceptionsTableCount = addExceptionsTableCount + 1
                If addExceptionsTableCount = 1 Then
                    Set m_asAddExceptionsTable = New Actions
                    m_asAddExceptionsTable.Parse child
                Else
                    DuplicateNodeError child.nodeName, actionSetNode.nodeName, c_proc
                End If

            Case mc_ANASDeleteExceptionsRow
                deleteExceptionsRowCount = deleteExceptionsRowCount + 1
                If deleteExceptionsRowCount = 1 Then
                    Set m_asDeleteExceptionsRow = New Actions
                    m_asDeleteExceptionsRow.Parse child
                Else
                    DuplicateNodeError child.nodeName, actionSetNode.nodeName, c_proc
                End If

            Case mc_ANASDeleteExceptionsTable
                deleteExceptionsTableCount = deleteExceptionsTableCount + 1
                If deleteExceptionsTableCount = 1 Then
                    Set m_asDeleteExceptionsTable = New Actions
                    m_asDeleteExceptionsTable.Parse child
                Else
                    DuplicateNodeError child.nodeName, actionSetNode.nodeName, c_proc
                End If

            Case Else
                InvalidNodeNameError child.nodeName, c_proc
            End Select
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub 'ParseActionSetNode

Private Sub ParseBookmarksNode(ByVal bookmarksNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.ParseBookmarksNode"

    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse the 'bookmarks' node
    For Each theAttribute In bookmarksNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANBExceptionTable
            m_bookmarkExceptionsTable = theAttribute.Text

        Case Else
            InvalidAttributeNameError bookmarksNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub 'ParseBookmarksNode

Private Sub ParseNewXMLNode(ByVal newXMLNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.ParseNewXMLNode"

    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse the 'newXML' node
    For Each theAttribute In newXMLNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANNNewException
            m_newExceptionXML = theAttribute.Text

        Case Else
            InvalidAttributeNameError newXMLNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub 'ParseNewXMLNode

Private Sub ParseQueriesNode(ByVal queriesNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.ParseQueriesNode"

    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse the 'queries' node
    For Each theAttribute In queriesNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANQAllFindings
            m_queryAllFindings = theAttribute.Text

        Case mc_ANQBuildDropDown
            m_queryBuildDropDown = theAttribute.Text

        Case mc_ANQNumberOfToggleButtons
            m_queryNumberOfToggleButtons = theAttribute.Text

        Case mc_ANQParentNodeForAddDelete
            m_queryParentNodeForAddDelete = theAttribute.Text

        Case mc_ANQTickedRecommendationButtons
            m_queryTickedRecommendationButtons = theAttribute.Text

        Case mc_ANQTickedStrengthButtons
            m_queryTickedStrengthButtons = theAttribute.Text

        Case mc_ANQNextSiblingDifferentTag
            m_queryNextSiblingDifferentTag = theAttribute.Text

        Case mc_ANQDeleteFindingsXML
            m_queryDeleteFindingsXML = theAttribute.Text

        Case Else
            InvalidAttributeNameError queriesNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub 'ParseQueriesNode

Friend Property Get AddContent() As Actions
    Set AddContent = m_addContent
End Property ' AddContent

Friend Property Get ActionSetAddExceptionsRow() As Actions
    Set ActionSetAddExceptionsRow = m_asAddExceptionsRow
End Property ' Get ActionSetAddExceptionsRow

Friend Property Get ActionSetAddExceptionsTable() As Actions
    Set ActionSetAddExceptionsTable = m_asAddExceptionsTable
End Property ' Get ActionSetAddExceptionsTable

Friend Property Get ActionSetDeleteExceptionsRow() As Actions
    Set ActionSetDeleteExceptionsRow = m_asDeleteExceptionsRow
End Property ' Get ActionSetDeleteExceptionsRow

Friend Property Get ActionSetDeleteExceptionsTable() As Actions
    Set ActionSetDeleteExceptionsTable = m_asDeleteExceptionsTable
End Property ' Get DeleteExceptionsTable

Friend Property Get BookmarkExceptionsTable() As String
    BookmarkExceptionsTable = m_bookmarkExceptionsTable
End Property ' BookmarkExceptionsTable

Friend Property Get DeleteContentItemBookmark() As String
    DeleteContentItemBookmark = m_deleteContentItemBookmark
End Property ' DeleteContentItemBookmark

Friend Property Get DeleteContentRenameBookmarks() As String
    DeleteContentRenameBookmarks = m_deleteContentRenameBookmarks
End Property ' DeleteContentRenameBookmarks

Friend Property Get DeleteContentUnusedBlockBookmark() As String
    DeleteContentUnusedBlockBookmark = m_deleteContentUnusedBlockBookmark
End Property ' DeleteContentUnusedBlockBookmark

Friend Property Get NewExceptionXML() As String
    NewExceptionXML = m_newExceptionXML
End Property ' Get NewExceptionXML

Friend Property Get AllFindings() As String
    AllFindings = m_queryAllFindings
End Property ' Get AllFindings

Friend Property Get QueryBuildDropDown() As String
    QueryBuildDropDown = m_queryBuildDropDown
End Property ' Get QueryBuildDropDown

Friend Property Get QueryDeleteFindingsXML() As String
    QueryDeleteFindingsXML = m_queryDeleteFindingsXML
End Property ' QueryDeleteFindingsXML

Friend Property Get QueryNumberOfToggleButtons() As String
    QueryNumberOfToggleButtons = m_queryNumberOfToggleButtons
End Property ' Get QueryNumberOfToggleButtons

Friend Property Get QueryParentNodeForAddDelete() As String
    QueryParentNodeForAddDelete = m_queryParentNodeForAddDelete
End Property ' Get QueryParentNodeForAddDelete

Friend Property Get QueryTickedRecommendationButtons() As String
    QueryTickedRecommendationButtons = m_queryTickedRecommendationButtons
End Property ' Get QueryTickedRecommendationButtons

Friend Property Get QueryTickedStrengthButtons() As String
    QueryTickedStrengthButtons = m_queryTickedStrengthButtons
End Property ' Get QueryTickedStrengthButtons

Friend Property Get QueryNextSiblingDifferentTag() As String
    QueryNextSiblingDifferentTag = m_queryNextSiblingDifferentTag
End Property ' Get QueryNextSiblingDifferentTag
