VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Actions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        Actions
' Purpose:      Data class that models the 'actions' and 'addContent' node-tree of the 'rda instructions.xml' file.
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
' History:      05/10/15    1.  Created.
'===================================================================================================================================
Option Explicit

Private Const mc_NNAdd    As String = "add"
Private Const mc_NNInsert As String = "insert"
Private Const mc_NNLink   As String = "link"
Private Const mc_NNRename As String = "rename"
Private Const mc_NNSetup  As String = "setup"


Private m_actions  As VBA.Collection


'=======================================================================================================================
' Procedure:    Initialise
' Purpose:      .
' Notes:        .
'
' On Entry:     actionsNode         .
'               addPattenBookmarks  Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Friend Sub Initialise(ByVal actionsNode As MSXML2.IXMLDOMNode, _
                      Optional ByVal addPattenBookmarks As Boolean)
    Const c_proc As String = "Actions.Initialise"

    Dim newAdd      As ActionAdd
    Dim child       As MSXML2.IXMLDOMNode
    Dim errorText   As String
    Dim newInsert   As ActionInsert
    Dim newLink     As ActionLink
    Dim newRename   As ActionRename
    Dim newSetup    As ActionSetup

    On Error GoTo Do_Error

    ' Parse out the action list

    ' Initialise the 'setup, add, insert and rename' action list collection object
    Set m_actions = New Collection

    ' Parse out each child node ('setup', 'add', 'insert' or 'rename')
    For Each child In actionsNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse the child node ('insert' and 'add' nodes)
            Select Case child.BaseName
            Case mc_NNInsert

                ' Create a new ActionInsert object
                Set newInsert = New ActionInsert

                ' Parse out the 'insert' node
                newInsert.Parse child, addPattenBookmarks

                ' Add it to the collection object so that we know what 'add, insert, link, rename and setup' sub nodes exist and in what order they occur
                m_actions.Add newInsert

            Case mc_NNAdd

                ' Create a new ActionAdd object
                Set newAdd = New ActionAdd

                ' Parse out the 'add' node
                newAdd.Parse child, addPattenBookmarks

                ' Add it to the collection object so that we know what 'add, insert, link, rename and setup' sub nodes exist and in what order they occur
                m_actions.Add newAdd

            Case mc_NNLink

                ' Create a new ActionLink object
                Set newLink = New ActionLink

                ' Parse out the 'link' node
                newLink.Parse child

                ' Add it to the collection object so that we know what 'add, insert, link, rename and setup' sub nodes exist and in what order they occur
                m_actions.Add newLink

            Case mc_NNRename

                ' Create a new ActionRename object
                Set newRename = New ActionRename

                ' Parse out the 'rename' node
                newRename.Parse child

                ' Add it to the collection object so that we know what 'add, insert, link, rename and setup' sub nodes exist and in what order they occur
                m_actions.Add newRename

            Case mc_NNSetup

                ' Create a new ActionSetup object
                Set newSetup = New ActionSetup

                ' Parse out the 'setup' node
                newSetup.Parse child

                ' Add it to the collection object so that we know what 'add, insert, link, rename and setup' sub nodes exist and in what order they occur
                m_actions.Add newSetup

            Case Else
                errorText = Replace$(mgrErrTextInvalidActionVerb, mgrP1, child.BaseName)
                Err.Raise mgrErrNoInvalidActionVerb, c_proc, errorText
            End Select
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

Friend Property Get actionList() As Collection
    Set actionList = m_actions
End Property ' Get ActionList
