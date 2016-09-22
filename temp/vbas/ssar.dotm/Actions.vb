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
' Purpose:      Data class that models the 'actions' and 'addContent' node-tree of the 'instructions.xml' file.
'
' Note 1:       Each action list can use the following actions:
'               add         -   Adds one of two Building Blocks to the Assessment Report at the specified Bookmark location.
'                               A Bookmarked area can be deleted if there is no data for the Building Block being added.
'               addDual     -   Inserts a table, where the table displays two sets of data items per row.
'               addRow      -   Adds a row stored as a Builidng Block to a table.
'               colourMap   -   Applies a contextural <colourMapping> to the specified bookmark using the bookmarks value for the
'                               colour lookups.
'               copy        -   Copies the text (without any formatting) of the from bookmark to the to bookmark.
'               counter     -   Implements a named counter on which very basic (add, subtract, set value and set value from query)
'                               operations can be performed. The counter value can then be used by If and Do actions to control
'                               instruction forks and loops.
'               delete      -   Delete a specified bookmark and the text and tables, or just tables or just the paragraph mark
'                               immediately before the bookmark start position.
'               deleteRow   -   Deletes a specified row from the table in the specified bookmark
'               do          -   Implements a basic loop contruct. Provides both For loop and Do loop functionality.
'               dropDown    -   Implements a DropDown Content Control
'               if          -   Conditional branching true/false instruction path.
'               insert      -   Inserts data into the Assessment Report from either the xml loaded from the '.rep' file or
'                               the HTML document if the dataFormat is RichText or Multiline.
'               link        -   Creates a Ref field to reference a piece of bookmarked data.
'               rename      -   Renames a Bookmark.

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
Private Const mc_NNAdd          As String = "add"
Private Const mc_NNAddDual      As String = "addDual"
Private Const mc_NNAddRow       As String = "addRow"
Private Const mc_NNColourMap    As String = "colourMap"
Private Const mc_NNCopy         As String = "copy"
Private Const mc_NNCounter      As String = "counter"
Private Const mc_NNDelete       As String = "delete"
Private Const mc_NNDeleteRow    As String = "deleteRow"
Private Const mc_NNDo           As String = "do"
Private Const mc_NNDropDown     As String = "dropDown"
Private Const mc_NNIf           As String = "if"
Private Const mc_NNInsert       As String = "insert"
Private Const mc_NNLink         As String = "link"
Private Const mc_NNRename       As String = "rename"
Private Const mc_NNSetup        As String = "setup"


Private m_actions  As VBA.Collection


'=======================================================================================================================
' Procedure:    Parse
' Purpose:      Initialises the classes actions collection by parsing the passed in node.
'
' On Entry:     actionsNode         The node whose child nodes are to be parsed.
'               addPatternBookmarks  Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Friend Sub Parse(ByVal actionsNode As MSXML2.IXMLDOMNode, _
                 Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "Actions.Parse"

    On Error GoTo Do_Error

    ' Parse out the action list
    ActionParser actionsNode, m_actions, addPatternBookmarks

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

Private Sub ActionParser(ByVal startingNode As MSXML2.IXMLDOMNode, _
                         ByRef theActions As VBA.Collection, _
                         Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "Actions.ActionParser"

    Dim child       As MSXML2.IXMLDOMNode
    Dim newAction   As IAction

    On Error GoTo Do_Error

    ' Initialise the action list collection object
    Set theActions = New Collection

    ' Parse out each child node (add, addDual, addRow, colourMap, copy, counter, delete, do, etc.)
    For Each child In startingNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse the child node (add, copy, counter, delete, do, etc.)
            Select Case child.BaseName
            Case mc_NNInsert
                Set newAction = New ActionInsert

            Case mc_NNAdd
                Set newAction = New ActionAdd

            Case mc_NNDo
                Set newAction = New ActionDo

            Case mc_NNDropDown
                Set newAction = New ActionDropDown

            Case mc_NNLink
                Set newAction = New ActionLink

            Case mc_NNIf
                Set newAction = New ActionIf

            Case mc_NNCounter
                Set newAction = New ActionCounter

            Case mc_NNCopy
                Set newAction = New ActionCopy

            Case mc_NNRename
                Set newAction = New ActionRename

            Case mc_NNAddDual
                Set newAction = New ActionAddDual

            Case mc_NNDelete
                Set newAction = New ActionDelete

            Case mc_NNColourMap
                Set newAction = New ActionColourMap

            Case mc_NNAddRow
                Set newAction = New ActionAddRow

            Case mc_NNDeleteRow
                Set newAction = New ActionDeleteRow

            Case Else
                InvalidActionVerbError child.BaseName, c_proc
            End Select

            ' Parse out the node and add it to the collection object (so we end up with a list of Actions in the order they occur)
            newAction.Parse child, addPatternBookmarks
            theActions.Add newAction
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ActionParser

Friend Property Get ActionList() As Collection
    Set ActionList = m_actions
End Property ' Get ActionList

'=======================================================================================================================
' Procedure:    BuildAssessmentReport
' Purpose:      Main processing loop for Building the Assessment Report (BAR).
' Notes:        This procedure is indirectly recursive.
'=======================================================================================================================
Friend Sub BuildAssessmentReport()
    Const c_proc As String = "Actions.BuildAssessmentReport"

    Dim theAction       As IAction

    On Error GoTo Do_Error

    ' Check that the collection exists before trying to use it
    If m_actions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all Action objects in the collection, invoking each Action objects BuildAssessmentReport method
    For Each theAction In m_actions

        ' Assemble the assemment report
        theAction.BuildAssessmentReport
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BuildAssessmentReport

'=======================================================================================================================
' Procedure:    ConstructRichText
' Purpose:      Main processing loop for creating the Rich Text data source used to build the Assessment Report.
' Notes:        This procedure is indirectly recursive.
'=======================================================================================================================
Friend Sub ConstructRichText()
    Const c_proc As String = "Actions.ConstructRichText"

    Dim theAction   As IAction

    On Error GoTo Do_Error

    ' Check that the collection exists before trying to use it
    If m_actions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all Action objects in the collection, invoking each objects ConstructRichText method
    For Each theAction In m_actions

        ' Create the Rich Text data source
        theAction.ConstructRichText
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ConstructRichText

'=======================================================================================================================
' Procedure:    HTMLForXMLUpdate
' Purpose:      Main processing loop for creating the HTML data source which is in turn converted to XHTML and used to
'               update the Assessment Report xml.
' Notes:        This procedure is indirectly recursive.
'=======================================================================================================================
Friend Sub HTMLForXMLUpdate()
    Const c_proc As String = "Actions.HTMLForXMLUpdate"

    Dim theAction   As IAction

    On Error GoTo Do_Error

    ' Check that the collection exists before trying to use it
    If m_actions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all Action objects in the collection, invoking each objects HTMLForXMLUpdate method
    For Each theAction In m_actions

        ' Create the Rich Text data source
        theAction.HTMLForXMLUpdate
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub '  HTMLForXMLUpdate

'=======================================================================================================================
' Procedure:    UpdateContentControlXML
' Purpose:      Main processing loop for updating the assessment report with values from Content Controls.
' Notes:        This procedure is indirectly recursive.
'=======================================================================================================================
Friend Sub UpdateContentControlXML()
    Const c_proc As String = "Actions.UpdateContentControlXML"

    Dim theAction   As IAction

    On Error GoTo Do_Error

    ' Check that the collection exists before trying to use it
    If m_actions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all Action objects in the collection, invoking each objects UpdateContentControlXML method
    For Each theAction In m_actions

        ' Create the Rich Text data source
        theAction.UpdateContentControlXML
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' UpdateContentControlXML

'=======================================================================================================================
' Procedure:    UpdateDateXML
' Purpose:      Main processing loop for validating editable date input areas.
' Notes:        This procedure is indirectly recursive.
'=======================================================================================================================
Friend Sub UpdateDateXML()
    Const c_proc As String = "Actions.UpdateDateXML"

    Dim theAction   As IAction

    On Error GoTo Do_Error

    ' Check that the collection exists before trying to use it
    If m_actions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all Action objects in the collection, invoking each objects UpdateDateXML method
    For Each theAction In m_actions

        ' Create the Rich Text data source
        theAction.UpdateDateXML
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' UpdateDateXML
