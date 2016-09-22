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
' Purpose:      The primary object for mapping the contents of the 'instruction.xml' files used to produce Assessment Reports.
' Note 1:       There is one class for each major node tree in the model:
'
'               Instructions
'                   initialise
'                   actions
'                       action...
'                   refresh
'                       action...
'                   reports
'                       fullReport
'                           action...
'                       summaryReport
'                           action...
'                   userInterface
'                       queries
'                       bookmarks
'                       newXML
'                       actionSet
'                           addExceptionsTable
'                               action...
'                           addExceptionsRow
'                               action...
'                           deleteExceptionsTable
'                               action...
'                           deleteExceptionsRow
'                               action...
'
' Note 2:       action... is one or more actions (add, addDual, addRow, colourMap, copy, counter, delete, deleteRow, do, dropDown,
'               if, insert, link or rename.
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
' History:      19/06/16    1.  Created.
'===================================================================================================================================
Option Explicit

Private Const mc_XQActions          As String = "/instructions/actions"
Private Const mc_XQInitialise       As String = "/instructions/initialise"
Private Const mc_XQRefresh          As String = "/instructions/refresh"
Private Const mc_XQFullReport       As String = "/instructions/reports/fullReport"
Private Const mc_XQSummaryReport    As String = "/instructions/reports/summaryReport"
Private Const mc_XQUserInterface    As String = "/instructions/userInterface"

Private m_actions       As Actions
Private m_fullReport    As Actions
Private m_initialise    As InstructionsInitialise
Private m_refresh       As Actions
Private m_summaryReport As Actions
Private m_ui            As UserInterface


Friend Sub Initialise(ByVal xmlDocument As MSXML2.DOMDocument60)
    Const c_proc As String = "Instructions.Initialise"

    Dim actionsNode         As MSXML2.IXMLDOMNode
    Dim fullReportNode      As MSXML2.IXMLDOMNode
    Dim initialiseNode      As MSXML2.IXMLDOMNode
    Dim refreshNode         As MSXML2.IXMLDOMNode
    Dim summaryReportNode   As MSXML2.IXMLDOMNode
    Dim userInterfaceNode   As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Get the three child nodes where the real stuff happens
    Set actionsNode = xmlDocument.SelectSingleNode(mc_XQActions)
    Set initialiseNode = xmlDocument.SelectSingleNode(mc_XQInitialise)
    Set refreshNode = xmlDocument.SelectSingleNode(mc_XQRefresh)
    Set fullReportNode = xmlDocument.SelectSingleNode(mc_XQFullReport)
    Set summaryReportNode = xmlDocument.SelectSingleNode(mc_XQSummaryReport)

    Set userInterfaceNode = xmlDocument.SelectSingleNode(mc_XQUserInterface)

    ' Initialise the Initialise object
    Set m_initialise = New InstructionsInitialise
    m_initialise.Parse initialiseNode

    ' Create the global object used to keep track of editable bookmark patterns and bookmarks
    Set g_editableBookmarks = New EditableBookmarks

    ' Initialise the Actions object and parse the 'actions' node
    Set m_actions = New Actions
    m_actions.Parse actionsNode, True

    ' Initialise the Refresh object and parse the 'refresh' node
    Set m_refresh = New Actions
    m_refresh.Parse refreshNode, True

    ' Initialise the Full Report object and parse the 'fullReport' node
    Set m_fullReport = New Actions
    m_fullReport.Parse fullReportNode, True

    ' Initialise the Summary Report object and parse the 'summaryReport' node
    Set m_summaryReport = New Actions
    m_summaryReport.Parse summaryReportNode, True

    ' Initialise the UserInterface object and parse the 'userInterface' node
    Set m_ui = New UserInterface
    m_ui.Parse userInterfaceNode

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

Friend Property Get Actions() As Actions
    Set Actions = m_actions
End Property ' Get Actions

Friend Property Get FullReport() As Actions
    Set FullReport = m_fullReport
End Property ' Get FullReport

Friend Property Get Refresh() As Actions
    Set Refresh = m_refresh
End Property ' Get Refresh

Friend Property Get SummaryReport() As Actions
    Set SummaryReport = m_summaryReport
End Property ' Get SummaryReport

Friend Property Get UserInterface() As UserInterface
    Set UserInterface = m_ui
End Property '  Get UserInterface
