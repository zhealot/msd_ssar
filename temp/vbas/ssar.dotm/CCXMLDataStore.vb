VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCXMLDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       CCXMLDataStore
' Purpose:      Creates the CustomXMLPart used to store the data used for all xml mapped Content Controls.
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

Private Const mc_rootNodeXPath  As String = "/%1"


Private m_ccXMLData     As Office.CustomXMLPart                 '
Private m_rootNode      As Office.CustomXMLNode                 '
Private m_rootNodeXPath As String                               '


'===================================================================================================================================
' Procedure:    Initialise
' Purpose:      Initialises this class using external data.
' Notes:        This is a single instance class.
' Date:         19/06/16    Created.
'
' On Entry:     rootNodeName        The name of the root node used in the custom xml data store part.
'===================================================================================================================================
Friend Sub Initialise(ByVal rootNodeName As String)
    Const c_proc            As String = "CCXMLDataStore.Initialise"
    Const c_rootNodeXML     As String = "<%1/>"

    Dim rootNodeXML     As String

    On Error GoTo Do_Error

    ' Create the Custom XML Part with a root node that will contain one node for each mapped Content Control
    rootNodeXML = Replace$(c_rootNodeXML, mgrP1, rootNodeName)
    Set m_ccXMLData = g_assessmentReport.CustomXMLParts.Add(rootNodeXML)

    ' Create a reference to the root node since its frequently used
    m_rootNodeXPath = Replace$(mc_rootNodeXPath, mgrP1, rootNodeName)
    Set m_rootNode = m_ccXMLData.SelectSingleNode(m_rootNodeXPath)

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

'===================================================================================================================================
' Procedure:    MapDropDown
' Purpose:      Maps DropDown Content Control to the assessment reports custom xml data store part.
' Notes:        The xml Data Store is updated from the rda xml.
' Date:         18/06/16    Created.
'
' On Entry:     theContentControl   The Content Control to map to the custom xml data store part.
'               rdaXPath            The xpath query used to retrieve the necessary data from the rda xml.
'               ccDataNodeName      The node name in the custom xml data store part the Content Control is mapped to.
'===================================================================================================================================
Friend Sub MapDropDown(ByVal theContentControl As Word.ContentControl, _
                       ByVal rdaXPath As String, _
                       ByVal ccDataNodeName As String)
    Const c_proc As String = "CCXMLDataStore.MapDropDown"

    Dim ccDataNode  As Office.CustomXMLNode
    Dim ccXPath     As String
    Dim rdaDataNode As MSXML2.IXMLDOMNode
    Dim theData     As String

    On Error GoTo Do_Error

    ' Replace any predicate place holders with their real values
    rdaXPath = g_counters.UpdatePredicates(rdaXPath)

    ' Create the real node name to use as it may need a number insert to create a unique node name
    ccDataNodeName = g_counters.UpdatePredicates(ccDataNodeName)

    ' Retrieve the data for the Content Control from rda
    Set rdaDataNode = g_xmlDocument.SelectSingleNode(rdaXPath)
    theData = rdaDataNode.Text

    ' Add the Content Controls data node to the custom xml data store part as its not part of the data store yet
    m_rootNode.AppendChildNode ccDataNodeName, "", msoCustomXMLNodeElement, theData

    ' Create the xpath query to retrieve the node we just added to the custom xml data store part
    ccXPath = m_rootNodeXPath & "/" & ccDataNodeName

    ' Now get a reference to the node we have just added to the custom xml data store part
    Set ccDataNode = m_ccXMLData.SelectSingleNode(ccXPath)

    ' Map the Content COntrol to the appropriate xml node in the custom xml data store part
    theContentControl.LockContentControl = False
    theContentControl.XMLMapping.SetMappingByNode ccDataNode

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' MapDropDown

'===================================================================================================================================
' Procedure:    UpdateRDAXML
' Purpose:      Update the RDA xml with the data from the Content Control mapped to the specified custom xml data store part.
' Notes:        The actual Content Control is not required as all we are doing is copying the xml data from the custom xml data
'               store part to the rda xml node.
' Date:         19/06/16        Created.
'
' On Entry:     rdaXPath            The xpath query used to specify which rda xml node to update.
'               ccDataNodeName      The node name in the custom xml data store part that contains the data to update rda with.
'===================================================================================================================================
Friend Sub UpdateRDAXML(ByVal rdaXPath As String, _
                        ByVal ccDataNodeName As String)
    Const c_proc As String = "CCXMLDataStore.UpdateRDAXML"

    Dim ccDataNode  As Office.CustomXMLNode
    Dim ccXPath     As String
    Dim rdaDataNode As MSXML2.IXMLDOMNode
    Dim theData     As String

    On Error GoTo Do_Error

    ' Create the xpath query to retrieve the node in the custom xml data store part that contains the data for the mapped Content Control
    ccXPath = m_rootNodeXPath & "/" & ccDataNodeName

    ' Create the real node name to use as it may need a number insert to create a unique node name
    ccXPath = g_counters.UpdatePredicates(ccXPath)

    ' Get a reference to the node in the custom xml data store part and retrieve the data from it
    Set ccDataNode = m_ccXMLData.SelectSingleNode(ccXPath)
    theData = ccDataNode.Text

    ' Replace any predicate place holders with their real values
    rdaXPath = g_counters.UpdatePredicates(rdaXPath)

    ' Retrieve the rda xml node whose data needs updating
    Set rdaDataNode = g_xmlDocument.SelectSingleNode(rdaXPath)

    ' Update the rda xml data node with the value from the mapped Content Control data node
    rdaDataNode.Text = theData

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' UpdateRDAXML

'===================================================================================================================================
' Procedure:    Value
' Purpose:      Return the current value of the specified node in the Custom XML Part Content COntrol data store.
' Date:         22/06/16    Created.
'
' On Entry:     ccDataNodeName      The node name in the custom xml data store part that contains the data we want the value of.
' Returns:      The value of the specified node.
'===================================================================================================================================
Friend Property Get Value(ByVal ccDataNodeName As String) As String
    Const c_proc As String = "CCXMLDataStore.Get Value"

    Dim ccXPath     As String

    On Error GoTo Do_Error

    ' Create the xpath query to retrieve the node required node from the custom xml data store part
    ccXPath = m_rootNodeXPath & "/" & ccDataNodeName

    ' Retrieve the node and return its value
    Value = m_ccXMLData.SelectSingleNode(ccXPath).Text

Do_Exit:
    Exit Property

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Property ' Get Value
