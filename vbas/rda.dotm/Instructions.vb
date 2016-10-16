VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        Instructions
' Purpose:      The primary object for mapping the contens of the 'RDA instruction.xml' files used to produce Assessmet Reports.
' Notes:        There is one class for each major node tree in the model:
'
'               Instructions
'                   Actions
'                       ActionAdd
'                       ActionInsert
'                       ActionRename
'                       ActionSetup
'                   UserInterface
'                       Actions
'                           ActionAdd
'                           ActionInsert
'                           ActionRename
'                           ActionSetup
'                       UIDeleteContent
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

Private Const mc_XQActions       As String = "/instructions/actions"
Private Const mc_XQUserInterface As String = "/instructions/userInterface"

Private m_actions As Actions
Private m_ui      As UserInterface


Friend Sub Initialise(ByVal xmlDocument As MSXML2.DOMDocument60)
    Const c_proc As String = "Instructions.Initialise"

    Dim actionsNode       As MSXML2.IXMLDOMNode
    Dim userInterfaceNode As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Get the two child nodes where the real stuff happens
    Set actionsNode = xmlDocument.SelectSingleNode(mc_XQActions)
    Set userInterfaceNode = xmlDocument.SelectSingleNode(mc_XQUserInterface)

    ' Create the global object used to keep track of editable bookmark patterns and bookmarks
    Set g_editableBookmarks = New EditableBookmarks

    ' Initialise the Actions object and parse the 'actions' node
    Set m_actions = New Actions
    m_actions.Initialise actionsNode, True

    ' Initialise the UserInterface object and parse the 'userInterface' node
    Set m_ui = New UserInterface
    m_ui.Initialise userInterfaceNode

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

Friend Property Get Actions() As Actions
    Set Actions = m_actions
End Property ' Get Actions

Friend Property Get UserInterface() As UserInterface
    Set UserInterface = m_ui
End Property '  Get UserInterface
