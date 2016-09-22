VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionCounter
' Purpose:      Data class that models the 'counter' action node of the 'instructions.xml' file.
'
' Note 1:       Maintains a simple named Long counter on which very basic (add, subtract, set value and set value from query)
'               operations can be performed. The counter value can then be used by ActionIf to control instruction forks.
'
' Note 2:       The following attributes are available for use with this Action:
'               name [Mandatory]
'                   The name of the counter to be used.
'               add*
'                   The value to be added.
'               subtract*
'                   The value to be subtracted.
'               value*
'                   Sets the counter to the specified value.
'               valueFromAddin
'                   Calls the specified function in this Addin, which will return a value that the counter will be set to.
'               valueFromQuery*
'                   Set the counters values to the number of nodes returned by an xpath query.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
'
'               Note *: Only one of these attributes is permitted and likewise one of these attributes is mandatory.
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
' History:      28/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction

' AN = Attribute Name
Private Const mc_ANAdd              As String = "add"               ' Adds the specified value to the named counter
Private Const mc_ANBreak            As String = "break"             ' Break if value is True
Private Const mc_ANName             As String = "name"              ' The name of the counter
Private Const mc_ANSubtract         As String = "subtract"          ' Subtracts the specified value from the named counter
Private Const mc_ANValue            As String = "value"             ' Sets the the named counter to the specified value
Private Const mc_ANValueFromAddin   As String = "valueFromAddin"    ' Sets the counter to the value returned by the specified function
Private Const mc_ANValueFromQuery   As String = "valueFromQuery"    ' Sets the counter to the number of nodes returned by this xpath query


' AVVFA = Attribute Value - Value From Addin
Private Const mc_AVVFADeleteRow             As String = "DeleteRow"
Private Const mc_AVVFAFirstRowToRename      As String = "FirstRowToRename"
Private Const mc_AVVFALastRowToRename       As String = "LastRowToRename"
Private Const mc_AVVFAStandardIndexNumber   As String = "StandardIndexNumber"
Private Const mc_AVVFAInsertRowAfter        As String = "InsertRowAfter"


Private Enum opType
    opTypeAdd = 1
    opTypeSubtract
    opTypeValue
    opTypeValueFromAddin
    opTypeValueFromQuery
End Enum ' opType

' Function Name Type - The names of special functions that provide a value to a counter
Private Enum fnType
    fnTypeStandardIndexNumber
    fnTypeFirstRowToRename
    fnTypeLastRowToRename
    fnTypeDeleteRow
    fnTypeInsertRowAfter
End Enum ' fnType


Private m_break             As Boolean                      ' Used to cause the code to execute a Stop instruction
Private m_counterName       As String                       ' The name of the counter
Private m_functionName      As String                       ' The name of the function to call when the operation is valueFromQuery
Private m_fnType            As fnType                       ' The function name translated to a numeric value
Private m_opType            As opType                       ' The operation to perform on the counter
Private m_query             As String                       ' The xpath query used by 'valueFromQuery'
Private m_value             As Long                         ' The value to set the counter to


'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Counter action instruction.
' Date:         28/06/16    Created.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionCounter.IAction_Parse"

    Dim nameCount       As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode
    Dim theValue        As String

    On Error GoTo Do_Error

    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANName
            nameCount = nameCount + 1
            m_counterName = theAttribute.Text

        Case mc_ANAdd
            If m_opType = 0 Then
                m_opType = opTypeAdd
                theValue = theAttribute.Text
            Else
                OnlyOneAttributeAllowedError actionNode.nodeName, mc_ANAdd, c_proc
            End If

        Case mc_ANSubtract
            If m_opType = 0 Then
                m_opType = opTypeSubtract
                theValue = theAttribute.Text
            Else
                OnlyOneAttributeAllowedError actionNode.nodeName, mc_ANSubtract, c_proc
            End If

        Case mc_ANValue
            If m_opType = 0 Then
                m_opType = opTypeValue
                theValue = theAttribute.Text
            Else
                OnlyOneAttributeAllowedError actionNode.nodeName, mc_ANValue, c_proc
            End If

        Case mc_ANValueFromQuery
            If m_opType = 0 Then
                m_opType = opTypeValueFromQuery
                m_query = theAttribute.Text
            Else
                OnlyOneAttributeAllowedError actionNode.nodeName, mc_ANValueFromQuery, c_proc
            End If

        Case mc_ANValueFromAddin
            If m_opType = 0 Then
                m_opType = opTypeValueFromAddin
                m_functionName = theAttribute.Text

                ' Validate and translate the Function Name in the 'valueFromAddin' attribute if present
                m_fnType = ParseAttributeValueFromAddin(actionNode, theAttribute)
            Else
                OnlyOneAttributeAllowedError actionNode.nodeName, mc_ANValueFromAddin, c_proc
            End If

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure one of the Mandatory attributes was present
    If nameCount < 1 Then
        MissingAttributeError actionNode.nodeName, mc_ANName, c_proc
    ElseIf m_opType = 0 Then
        MissingAttributeError actionNode.nodeName, "add, subtract, value, valueFromAddin or valueFromQuery", c_proc
    End If

    ' Check for duplicate name attributes
    If nameCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANName, c_proc
    End If

    ' Check that a value was supplied for: add, subtract, value, valueFromAddin and valueFromQuery
    If m_opType = opTypeAdd And LenB(theValue) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANAdd, theValue, c_proc
    ElseIf m_opType = opTypeSubtract And LenB(theValue) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANSubtract, theValue, c_proc
    ElseIf m_opType = opTypeValue And LenB(theValue) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANValue, theValue, c_proc
    ElseIf m_opType = opTypeValueFromQuery And LenB(m_query) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANValueFromQuery, m_query, c_proc
    ElseIf m_opType = opTypeValueFromAddin And LenB(m_functionName) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANValueFromAddin, m_query, c_proc
    End If

    ' Store the string value as a numeric value now that we are finished with the string
    If LenB(theValue) > 0 Then
        m_value = CLng(theValue)
    End If

    ' Make sure the counter collection object exists
    If g_actionCounters Is Nothing Then
        Set g_actionCounters = New Scripting.Dictionary
    End If

    ' Add the counter to the counter collection if it is not already present
    If Not g_actionCounters.Exists(m_counterName) Then
        g_actionCounters.Add m_counterName, 0
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'===================================================================================================================================
' Procedure:    ParseAttributeValueFromAddin
' Purpose:      Parses, validates and translates to a numeric value the 'valueFromAddin' attribute.
' Notes:        The advantage of doing the function name translation now is that we can provide better contextural error messages
'               and that we translate the text to a numeric value, so subsequent use is easier.
'
' Date:         12/07/16    Created.
'
' On Entry:     actionNode          The xml node that contains the 'valueFromAddin' attribute.
'               theAttribute        The 'valueFromAddin' attribute node.
' Returns:      The numeric value for the function name in the 'valueFromAddin' attribute.
'===================================================================================================================================
Private Function ParseAttributeValueFromAddin(ByRef actionNode As MSXML2.IXMLDOMNode, _
                                              ByRef theAttribute As MSXML2.IXMLDOMNode) As fnType
    Const c_proc As String = "ActionAdd.ParseAttributeValueFromAddin"

    On Error GoTo Do_Error

    ' Call the function specified by the function name
    Select Case m_functionName
    Case mc_AVVFAStandardIndexNumber
        ParseAttributeValueFromAddin = fnTypeStandardIndexNumber
    Case mc_AVVFAFirstRowToRename
        ParseAttributeValueFromAddin = fnTypeFirstRowToRename
    Case mc_AVVFALastRowToRename
        ParseAttributeValueFromAddin = fnTypeLastRowToRename
    Case mc_AVVFADeleteRow
        ParseAttributeValueFromAddin = fnTypeDeleteRow
    Case mc_AVVFAInsertRowAfter
        ParseAttributeValueFromAddin = fnTypeInsertRowAfter
    Case Else
        UnknownFunctionNameError actionNode.nodeName, theAttribute.BaseName, theAttribute.Text, c_proc
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ParseAttributeValueFromAddin

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Performs an operation on a predefined Action counter while building an assessment report.
' Notes:        The counters are currently used only by ActionIf.
' Date:         28/06/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionCounter.IAction_BuildAssessmentReport"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Perform the required action on the specified named counter
    CounterAction

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

'===================================================================================================================================
' Procedure:    IAction_ConstructRichText
' Purpose:      Performs an operation on a predefined Action counter while building the rich text data source.
' Notes:        The counters are currently used only by ActionIf.
' Date:         28/06/16    Created.
'===================================================================================================================================
Private Sub IAction_ConstructRichText()
    Const c_proc As String = "ActionCounter.IAction_ConstructRichText"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Perform the required action on the specified named counter
    CounterAction

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_ConstructRichText

Private Sub IAction_HTMLForXMLUpdate()
    ' Definition required - But nothing to do
End Sub '  Sub IAction_HTMLForXMLUpdate

Private Sub IAction_UpdateContentControlXML()
    Const c_proc As String = "ActionCounter.IAction_UpdateDateXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Perform the required action on the specified named counter
    CounterAction

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_UpdateContentControlXML

'===================================================================================================================================
' Procedure:    IAction_UpdateDateXML
' Purpose:      Performs an operation on a predefined Action counter while performing editable area date validation.
' Date:         4/08/16     Created.
'===================================================================================================================================
Private Sub IAction_UpdateDateXML()
    Const c_proc As String = "ActionCounter.IAction_UpdateDateXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Perform the required action on the specified named counter
    CounterAction

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Sub IAction_UpdateDateXML

'===================================================================================================================================
' Procedure:    CounterAction
' Purpose:      Perform the actual work when needing to set a named counter value.
' Note 1:       Do not instantiate an error handler, let the caller handle any errors.
' Date:         16/08/16    Created.
'===================================================================================================================================
Private Sub CounterAction()
    Dim counterValue    As Long

    ' Get the counters current value
    counterValue = g_actionCounters.Item(m_counterName)

    ' Perform the required operation on the counter
    Select Case m_opType
    Case opTypeAdd
        counterValue = counterValue + m_value
    Case opTypeSubtract
        counterValue = counterValue - m_value
    Case opTypeValue
        counterValue = m_value
    Case opTypeValueFromQuery
        counterValue = GetValueUsingQuery()
    Case opTypeValueFromAddin
        counterValue = GetValueFromFunction()
    End Select

    ' Update the counter dictionary object
    g_actionCounters.Item(m_counterName) = counterValue
End Sub ' CounterAction

'===================================================================================================================================
' Procedure:    GetValueUsingQuery
' Purpose:      Returns the number of nodes the xpath query (m_query) produced.
'===================================================================================================================================
Private Function GetValueUsingQuery() As Long
    Const c_proc As String = "ActionCounter.GetValueUsingQuery"

    Dim queryResult As MSXML2.IXMLDOMNodeList
    Dim theQuery    As String

    On Error GoTo Do_Error

    ' Replace any predicate placeholder character sequences in the query
    theQuery = g_counters.UpdatePredicates(m_query)

    ' Query the assessment report xml
    Set queryResult = g_xmlDocument.SelectNodes(theQuery)

    ' See if anything was returned by the query
    If Not queryResult Is Nothing Then
        GetValueUsingQuery = queryResult.Length
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' GetValueUsingQuery

'===================================================================================================================================
' Procedure:    GetValueFromFunction
' Purpose:      Returns the value from a called function in the UI module.
' Note 1:       This allows us to pass values determined by the UI through the counter to other action which can use named counters.
'===================================================================================================================================
Private Function GetValueFromFunction() As Long
    Const c_proc As String = "ActionCounter.GetValueFromFunction"

    On Error GoTo Do_Error

    ' Call the function specified by the function name
    Select Case m_fnType
    Case fnTypeStandardIndexNumber
        GetValueFromFunction = CV_StandardIndexNumber()
    Case fnTypeFirstRowToRename
        GetValueFromFunction = CV_FirstRowToRename()
    Case fnTypeLastRowToRename
        GetValueFromFunction = CV_LastRowToRename()
    Case fnTypeDeleteRow
        GetValueFromFunction = CV_DeleteRow()
    Case fnTypeInsertRowAfter
        GetValueFromFunction = CV_InsertRowAfter()
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' GetValueFromFunction

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeCounter
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
