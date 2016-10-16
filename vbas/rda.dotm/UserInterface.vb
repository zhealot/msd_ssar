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
' Purpose:      Data class that models the 'userInterface' action node-tree of the 'rda instructions.xml' file.
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
' History:      04/11/15    1.  Created.
'===================================================================================================================================
Option Explicit

' NN = Node Name
Private Const mc_NNQueries                  As String = "queries"
Private Const mc_NNAddContent               As String = "addContent"
Private Const mc_NNDeleteContent            As String = "deleteContent"

' ANQ = Attribute Names - Queries
Private Const mc_ANQBuildDropDown           As String = "buildDropDown"
Private Const mc_ANQDeleteXML               As String = "deleteXML"
Private Const mc_ANQNextSiblingDifferentTag As String = "nextSiblingDifferentTag"
Private Const mc_ANQParentNodeForAddDelete  As String = "parentNodeForAddDelete"
Private Const mc_ANQToggleButtonSiblings    As String = "toggleButtonSiblings"
Private Const mc_ANQVisibleToggleButtons    As String = "visibleToggleButtons"

' ANDC = Attribute Names - DeleteContent
Private Const mc_ANDCItemBookmark           As String = "itemBookmark"
Private Const mc_ANDCRenameBookmarks        As String = "renameBookmarks"
Private Const mc_ANDCUnusedBlockBookmark    As String = "unusedBlockBookmark"


Private m_addContent                       As Actions

Private m_deleteContentItemBookmark        As String
Private m_deleteContentRenameBookmarks     As String
Private m_deleteContentUnusedBlockBookmark As String

Private m_queryBuildDropDown               As String
Private m_queryDeleteXML                   As String
Private m_queryNextSiblingDifferentTag     As String
Private m_queryParentNodeForAddDelete      As String
Private m_queryToggleButtonSiblings        As String
Private m_queryVisibleToggleButtons        As String


Friend Sub Initialise(ByVal userInterfaceNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.Initialise"

    Dim child     As MSXML2.IXMLDOMNode
    Dim errorText As String

    On Error GoTo Do_Error

    ' Create new AddContent object that stores the "setup" and "use" steps
    Set m_addContent = New Actions

    ' Parse out each child node ('queries', 'addContent' or 'deleteContent')
    For Each child In userInterfaceNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse out each child node
            Select Case child.BaseName
            Case mc_NNQueries
                ParseQueriesNode child

            Case mc_NNAddContent
                Set m_addContent = New Actions
                m_addContent.Initialise child

            Case mc_NNDeleteContent
                ParseDeleteContentNode child

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
End Sub ' Initialise

Private Sub ParseDeleteContentNode(ByVal deleteContentNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.ParseDeleteContentNode"

    Dim errorText    As String
    Dim theAttribute As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse the 'deleteContent' node
    For Each theAttribute In deleteContentNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANDCItemBookmark
            m_deleteContentItemBookmark = theAttribute.Text

        Case mc_ANDCUnusedBlockBookmark
            m_deleteContentUnusedBlockBookmark = theAttribute.Text

        Case mc_ANDCRenameBookmarks
            m_deleteContentRenameBookmarks = theAttribute.Text

        Case Else
            errorText = Replace$(mgrErrTextInvalidAttributeName, mgrP1, deleteContentNode.BaseName)
            errorText = Replace$(errorText, mgrP2, theAttribute.BaseName)
            Err.Raise mgrErrNoInvalidAttributeName, c_proc, errorText
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub 'ParseDeleteContentNode

Private Sub ParseQueriesNode(ByVal queriesNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "UserInterface.ParseQueriesNode"

    Dim errorText    As String
    Dim theAttribute As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse the 'queries' node
    For Each theAttribute In queriesNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANQBuildDropDown
            m_queryBuildDropDown = theAttribute.Text

        Case mc_ANQVisibleToggleButtons
            m_queryVisibleToggleButtons = theAttribute.Text

        Case mc_ANQParentNodeForAddDelete
            m_queryParentNodeForAddDelete = theAttribute.Text

        Case mc_ANQToggleButtonSiblings
            m_queryToggleButtonSiblings = theAttribute.Text

        Case mc_ANQNextSiblingDifferentTag
            m_queryNextSiblingDifferentTag = theAttribute.Text

        Case mc_ANQDeleteXML
            m_queryDeleteXML = theAttribute.Text

        Case Else
            errorText = Replace$(mgrErrTextInvalidAttributeName, mgrP1, queriesNode.BaseName)
            errorText = Replace$(errorText, mgrP2, theAttribute.BaseName)
            Err.Raise mgrErrNoInvalidAttributeName, c_proc, errorText
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

Friend Property Get DeleteContentItemBookmark() As String
    DeleteContentItemBookmark = m_deleteContentItemBookmark
End Property ' DeleteContentItemBookmark

Friend Property Get DeleteContentRenameBookmarks() As String
    DeleteContentRenameBookmarks = m_deleteContentRenameBookmarks
End Property ' DeleteContentRenameBookmarks

Friend Property Get DeleteContentUnusedBlockBookmark() As String
    DeleteContentUnusedBlockBookmark = m_deleteContentUnusedBlockBookmark
End Property ' DeleteContentUnusedBlockBookmark

Friend Property Get QueryBuildDropDown() As String
    QueryBuildDropDown = m_queryBuildDropDown
End Property ' Get QueryBuildDropDown

Friend Property Get QueryDeleteXML() As String
    QueryDeleteXML = m_queryDeleteXML
End Property ' QueryDeleteXML

Friend Property Get QueryVisibleToggleButtons() As String
    QueryVisibleToggleButtons = m_queryVisibleToggleButtons
End Property ' Get QueryVisibleToggleButtons

Friend Property Get QueryParentNodeForAddDelete() As String
    QueryParentNodeForAddDelete = m_queryParentNodeForAddDelete
End Property ' Get QueryParentNodeForAddDelete

Friend Property Get QueryToggleButtonSiblings() As String
    QueryToggleButtonSiblings = m_queryToggleButtonSiblings
End Property ' Get QueryToggleButtonSiblings

Friend Property Get QueryNextSiblingDifferentTag() As String
    QueryNextSiblingDifferentTag = m_queryNextSiblingDifferentTag
End Property ' Get QueryNextSiblingDifferentTag
