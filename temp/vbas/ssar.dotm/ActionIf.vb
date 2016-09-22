VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionIf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionIf
' Purpose:      Data class that models the 'if' action node-tree of the 'instructions.xml' file.
' Note 1:       The following attributes are available for use with this Action:
'               condition*
'                   An xpath query run against the rda xml data.
'               counter*
'                   The name of a counter whose value is to be compared to an acompanying 'value' attribute.
'               value*
'                   1.  The value the named 'counter' is to be compared to.
'                   2.  The name of another named 'counter', 'counter' is to be compared with.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
'
'               Note *: Only one of these two nodes are permitted 'condition' or 'counter'.
'                       If 'condition' is present, 'value' is optional.
'                       If 'value' is present then the node count from the results of the 'condition' xpath query are used to
'                       evaluate True, using an equality test.
'                       If 'counter' is present 'operator' and 'value' must be present as well.
'
' Note 2:       The 'if' node must have either one or two child nodes. The following combinations are permissible:
'               1. One 'then' node.
'               2. One 'else' node.
'               3. One 'then' node and one 'else' node.
'               All other nodes, such as: add, insert, link, rename etc. must be child nodes of the 'then' and 'else' nodes,
'               not the 'if' node.
'
' Note 3:       Attribute combinations:
'                     condition  counter  operator  value
'               1.        x
'               2.        x                           x
'               3.                  x        x        x
'
'               1.  'condition' is deemed to be True if its xpath query returns one or more node.
'               2.  'condition' is deemed to be True if its xpath query returns the number of nodes specified by 'value'.
'               3.  'condition' is deemed to be True if the named counter and the number specifed by 'value' meet the operator condition.
'
' Note 4:       Valid operator values are:
'               EQ = EQual
'               NE = Not Equal
'               LT = Less Than (the named counter is Less Than the specified value)
'               GT = Greater Than (the named counter is Greater Than the specified value)
'               LE = Less than or Equal to (the named counter is Less than or Equal to the specified value)
'               GE = Greater than or Equal to (the named counter is Greater than or Equal to the specified value)
'
' Note 5:       When the If evaluates to True the 'then' part of the 'if' block will be processed if present.
'               When the If evaluates to False the 'else' part of the 'if' block will be processed if present.
'
' Example 1:    <if condition="/Assessment/report/standard[suggestionsForQualityEnhancement/narrative != '']">
'                   <then>
'                       <insert bookmark   ="disclaimer"
'                               dataFormat ="Text"
'                               data       ="/Assessment/disclaimer"/>
'                   </then>
'                   <else>
'                       <insert bookmark   ="purpose"
'                               dataFormat ="Text"
'                               data       ="/Assessment/purpose"/>
'                   </else>
'               </if>
'
' Example 2:    <if condition="/Assessment/report/standard" value="10">
'                   <then>
'                       <insert bookmark   ="disclaimer"
'                               dataFormat ="Text"
'                               data       ="/Assessment/disclaimer"/>
'                   </then>
'               </if>
'
' Example 2:    <if counter="findings" operator="EQ" value="1">
'                   <else>
'                       <insert bookmark   ="disclaimer"
'                               dataFormat ="Text"
'                               data       ="/Assessment/disclaimer"/>
'                   </else>
'               </if>
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
' History:      10/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBreak        As String = "break"
Private Const mc_ANCondition    As String = "condition"
Private Const mc_ANCounter      As String = "counter"
Private Const mc_ANOperator     As String = "operator"
Private Const mc_ANValue        As String = "value"

' AOV = Attribute 'operator' Value
Private Const mc_AOVEq          As String = "EQ"
Private Const mc_AOVNe          As String = "NE"
Private Const mc_AOVLt          As String = "LT"
Private Const mc_AOVGt          As String = "GT"
Private Const mc_AOVLe          As String = "LE"
Private Const mc_AOVGe          As String = "GE"

' NN = Node Name
Private Const mc_NNThen         As String = "then"
Private Const mc_NNElse         As String = "else"


Private Enum ifType
    ifTypeConditionOnly
    ifTypeConditionValue
    ifTypeCounterOperatorValue
End Enum ' ifType


Private m_break             As Boolean              ' Used to cause the code to execute a Stop instruction
Private m_condition         As String               ' This 'if' blocks xpath query
Private m_counterName       As String               ' The named counter to use for the If test
Private m_ifType            As ifType               ' The type of if statement to evaluate
Private m_operator          As ssarOperatorType     ' The operator as an operator type
Private m_value             As Long                 ' The value the named counter must equal
Private m_valueCounter      As String               ' The named counter to be used for value tests
Private m_elseActions       As Actions              ' List of 'else' (False) actions
Private m_thenActions       As Actions              ' List of 'then' (True) actions


'===================================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the If action instruction.
' Date:         29/06/16    Created.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'===================================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionIf.IAction_Parse"

    Dim child           As MSXML2.IXMLDOMNode
    Dim conditionCount  As Long
    Dim counterCount    As Long
    Dim elseCount       As Long
    Dim errorText       As String
    Dim operatorCount   As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode
    Dim thenCount       As Long
    Dim valueCount      As Long

    On Error GoTo Do_Error

    ' Set a default value for the operator attribute
    m_operator = ssarOperatorTypeNone

    ' Parse out the 'if' nodes attributes
    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANCondition
            conditionCount = conditionCount + 1
            m_condition = theAttribute.Text
            If LenB(Trim$(m_condition)) = 0 Then
                InvalidAttributeValueError actionNode.nodeName, mc_ANCondition, m_condition, c_proc
            End If

        Case mc_ANCounter
            counterCount = counterCount + 1
            m_counterName = Trim$(theAttribute.Text)
            If LenB(m_counterName) = 0 Then
                InvalidAttributeValueError actionNode.nodeName, mc_ANCounter, m_counterName, c_proc
            End If

        Case mc_ANOperator
            operatorCount = operatorCount + 1
            m_operator = ParseAttributeOperator(actionNode, theAttribute)

        Case mc_ANValue
            valueCount = valueCount + 1

            ' If value is non numeric then it must be a named counter
            If IsNumeric(theAttribute.Text) Then
                m_value = CLng(theAttribute.Text)
            Else
                m_valueCounter = Trim$(theAttribute.Text)
                If LenB(m_valueCounter) = 0 Then
                    InvalidAttributeValueError actionNode.nodeName, mc_ANCounter, m_valueCounter, c_proc
                End If
            End If

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Check that there was one 'condition' attribute and it is not a null string
    If conditionCount < 1 And counterCount < 1 Then
        If conditionCount < 1 Then
            MissingAttributeError actionNode.nodeName, mc_ANCondition, c_proc
        Else
            MissingAttributeError actionNode.nodeName, mc_ANCounter, c_proc
        End If
    ElseIf conditionCount > 1 Or counterCount > 1 Or operatorCount > 1 Then
        If conditionCount > 1 Then
            DuplicateAttributeError actionNode.nodeName, mc_ANCondition, c_proc
        ElseIf operatorCount > 1 Then
            DuplicateAttributeError actionNode.nodeName, mc_ANOperator, c_proc
        Else
            DuplicateAttributeError actionNode.nodeName, mc_ANCounter, c_proc
        End If
    End If

    ' If 'counter' is present make sure both 'operator' and 'value' are present as well
    If counterCount = 1 Then
        If operatorCount = 0 Then
            MissingAttributeError actionNode.nodeName, mc_ANOperator, c_proc
        ElseIf valueCount = 0 Then
            MissingAttributeError actionNode.nodeName, mc_ANValue, c_proc
        End If
    End If

    ' Set if type to be evaluated
    If conditionCount = 1 And valueCount = 0 Then
        m_ifType = ifTypeConditionOnly
    ElseIf conditionCount = 1 And valueCount = 1 Then
        m_ifType = ifTypeConditionValue
    Else
        m_ifType = ifTypeCounterOperatorValue
    End If

    ' Parse out the 'if' node to make sure it only contains one 'then' node, or one 'else' node or one 'then' node and one 'else' node
    For Each child In actionNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse the child node ('then' and 'else' nodes)
            Select Case child.BaseName
            Case mc_NNThen
                thenCount = thenCount + 1

                ' Parse out all of the 'then' nodes child nodes
                Set m_thenActions = New Actions
                m_thenActions.Parse child, addPatternBookmarks

            Case mc_NNElse
                elseCount = elseCount + 1

                ' Parse out all of the 'else' nodes child nodes
                Set m_elseActions = New Actions
                m_elseActions.Parse child, addPatternBookmarks

            Case Else
                InvalidActionVerbError child.BaseName, c_proc
            End Select
        End If
    Next

    ' Validate the 'if' nodes child nodes
    If Not (thenCount = 1 And elseCount = 0 Or thenCount = 0 And elseCount = 1 Or thenCount = 1 And elseCount = 1) Then
        errorText = Replace$(mgrErrTextInvalidIfBlock, mgrP1, actionNode.Text)
        Err.Raise mgrErrNoInvalidIfBlock, c_proc, errorText
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

Private Function ParseAttributeOperator(ByRef actionNode As MSXML2.IXMLDOMNode, _
                                        ByRef theAttribute As MSXML2.IXMLDOMNode) As ssarOperatorType
    Const c_proc As String = "ActionIf.ParseAttributeOperator"

    On Error GoTo Do_Error

    Select Case theAttribute.Text
    Case mc_AOVEq
        ParseAttributeOperator = ssarOperatorTypeEQ
    Case mc_AOVNe
        ParseAttributeOperator = ssarOperatorTypeNE
    Case mc_AOVLt
        ParseAttributeOperator = ssarOperatorTypeLT
    Case mc_AOVGt
        ParseAttributeOperator = ssarOperatorTypeGT
    Case mc_AOVLe
        ParseAttributeOperator = ssarOperatorTypeLE
    Case mc_AOVGe
        ParseAttributeOperator = ssarOperatorTypeGE
    Case Else
        InvalidAttributeValueError actionNode.nodeName, theAttribute.BaseName, theAttribute.Text, c_proc
    End Select

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ParseAttributeOperator

'=======================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Allows the instruction file to fork into True and False blocks while Building the Assessment Report.
' Date:         29/06/16    Created.
'=======================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc    As String = "ActionIf.IAction_BuildAssessmentReport"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Do something as a result of the If test
    If EvaluateIf Then

        ' Execute the True Action fork
        If Not m_thenActions Is Nothing Then
            m_thenActions.BuildAssessmentReport
        End If
    Else

        ' Execute the False Action fork
        If Not m_elseActions Is Nothing Then
            m_elseActions.BuildAssessmentReport
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

'===================================================================================================================================
' Procedure:    IAction_ConstructRichText
' Purpose:      Allows the instruction file to fork into True and False blocks while creating the Rich Text data source used to
'               build an ssar assessment report.
' Date:         29/06/16    Created.
'===================================================================================================================================
Private Sub IAction_ConstructRichText()
    Const c_proc    As String = "ActionIf.IAction_ConstructRichText"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Do something as a result of the If test
    If EvaluateIf Then
        m_thenActions.ConstructRichText
    Else
        m_elseActions.ConstructRichText
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_ConstructRichText

'===================================================================================================================================
' Procedure:    IAction_HTMLForXMLUpdate
' Procedure:    IAction_ConstructRichText
' Purpose:      Allows the instruction file to fork into True and False blocks while creating the HTML to XHTML data source used to
'               update the assessment report DOMDocument xml.
' Date:         05/08/16    Created.
'===================================================================================================================================
Private Sub IAction_HTMLForXMLUpdate()
    Const c_proc    As String = "ActionIf.IAction_UpdateDateXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Do something as a result of the If test
    If EvaluateIf Then
        m_thenActions.HTMLForXMLUpdate
    Else
        m_elseActions.HTMLForXMLUpdate
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub '  Sub IAction_HTMLForXMLUpdate

'===================================================================================================================================
' Procedure:    IAction_UpdateContentControlXML
' Purpose:      Allows the instruction file to fork into True and False blocks while updating the assessment report xml using the
'               values of the Content Controls.
' Date:         16/08/16    Created.
'===================================================================================================================================
Private Sub IAction_UpdateContentControlXML()
    Const c_proc    As String = "ActionIf.IAction_UpdateContentControlXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Do something as a result of the If test
    If EvaluateIf Then
        m_thenActions.UpdateContentControlXML
    Else
        m_elseActions.UpdateContentControlXML
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_UpdateContentControlXML

'===================================================================================================================================
' Procedure:    IAction_UpdateDateXML
' Purpose:      Allows the instruction file to fork into True and False blocks while updating the corresponding xml of any editable
'               area that contains a date.
' Note 1:       Before the xml update occurs the dates are validated to ensure that the user has entered a valid date.
' Date:         04/08/16    Created.
'===================================================================================================================================
Private Sub IAction_UpdateDateXML()
    Const c_proc    As String = "ActionIf.IAction_UpdateDateXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Do something as a result of the If test
    If EvaluateIf Then
        m_thenActions.UpdateDateXML
    Else
        m_elseActions.UpdateDateXML
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Sub IAction_UpdateDateXML

'===================================================================================================================================
' Procedure:    ConditionAndValueIf
' Purpose:      Performs a comparison between the results of the xpath query and the specified value.
' Notes:        The value can be a specified numeric value or the name of a named counter. If a named counter is specified then the
'               number of nodes returned by the xpath queery is compared to the counters value. If they are equal then the test is
'               deemed to be True.
'
' Date:         11/07/16    Rewritten.
'===================================================================================================================================
Private Function ConditionAndValueIf() As Boolean
    Const c_proc    As String = "ActionIf.ConditionAndValueIf"

    Dim theResults  As MSXML2.IXMLDOMNodeList
    Dim theQuery    As String
    Dim theValue    As Long

    On Error GoTo Do_Error

    ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct nodes
    theQuery = g_counters.UpdatePredicates(m_condition)

    ' Execute the xpath query
    Set theResults = g_xmlDocument.SelectNodes(theQuery)

    ' If 'value' is a named counter use the value of the named counter
    If LenB(m_valueCounter) > 0 Then
        theValue = g_actionCounters.Item(m_valueCounter)
    Else
        theValue = m_value
    End If

    ' Return the result - The If statement evaluates to True if the xpath query returned the number of nodes specified by 'value'
    ConditionAndValueIf = (theResults.Length = theValue)

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ConditionAndValueIf

'===================================================================================================================================
' Procedure:    ConditionIf
' Purpose:      Executes an xpath query, if the query returns one or more nodes then the test is deemed True.
'
' Date:         10/06/16    Created.
'===================================================================================================================================
Private Function ConditionIf() As Boolean
    Const c_proc    As String = "ActionIf.ConditionIf"

    Dim theResults  As MSXML2.IXMLDOMNodeList
    Dim theQuery    As String

    On Error GoTo Do_Error

    ' Replace the predicate place holder (if present) with the predicate value so that we retrieve the correct node occurrence
    theQuery = g_counters.UpdatePredicates(m_condition)

    ' Execute the xpath query
    Set theResults = g_xmlDocument.SelectNodes(theQuery)
    
    ' Return the result - The If statement evaluates to True if one or more nodes were returned by the xpath query
    ConditionIf = (theResults.Length > 0)

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ConditionIf

'===================================================================================================================================
' Procedure:    CounterOperatorValueIf
' Purpose:      Performs a specifiable comparison (using the operator) between the value of a named counter and either a specified
'               numeric value or the value if another named counter.
'
' Date:         11/07/16    Rewritten.
'===================================================================================================================================
Private Function CounterOperatorValueIf() As Boolean
    Const c_proc    As String = "ActionIf.CounterOperatorValueIf"

    Dim counterValue    As Long
    Dim theValue        As Long

    On Error GoTo Do_Error

    ' Get the counter from the counters dictionary object
    If g_actionCounters.Exists(m_counterName) Then

        counterValue = g_actionCounters.Item(m_counterName)

        ' If 'value' is a named counter use the value of the named counter
        If LenB(m_valueCounter) > 0 Then
            If g_actionCounters.Exists(m_valueCounter) Then
                theValue = g_actionCounters.Item(m_valueCounter)
            Else
                KeyDoesNotExistInDictionaryError m_valueCounter, "g_actionCounters", c_proc
            End If
        Else
            theValue = m_value
        End If

        ' Now return the results of the comparison of the counter value against the value using the specified operator
        Select Case m_operator
        Case ssarOperatorTypeEQ
            CounterOperatorValueIf = (counterValue = theValue)
        Case ssarOperatorTypeNE
            CounterOperatorValueIf = (counterValue <> theValue)
        Case ssarOperatorTypeLT
            CounterOperatorValueIf = (counterValue < theValue)
        Case ssarOperatorTypeGT
            CounterOperatorValueIf = (counterValue > theValue)
        Case ssarOperatorTypeLE
            CounterOperatorValueIf = (counterValue <= theValue)
        Case ssarOperatorTypeGE
            CounterOperatorValueIf = (counterValue >= theValue)
        End Select
    Else
        KeyDoesNotExistInDictionaryError m_counterName, "g_actionCounters", c_proc
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' CounterOperatorValueIf

'===================================================================================================================================
' Procedure:    EvaluateIf
' Purpose:      Evaluates the if type specified for this class instance.
' Date:         07/08/16    Created.
'
' Returns:      True if the If condition test evaluated to True, otherwise False.
'===================================================================================================================================
Private Function EvaluateIf() As Boolean

    ' Evaluate the appropriate If type
    Select Case m_ifType
    Case ifTypeConditionOnly
        EvaluateIf = ConditionIf
    Case ifTypeConditionValue
        EvaluateIf = ConditionAndValueIf
    Case ifTypeCounterOperatorValue
        EvaluateIf = CounterOperatorValueIf
    End Select
End Function ' EvaluateIf

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeIf
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break

Friend Property Get Condition() As String
    Condition = m_condition
End Property ' Condition

Friend Property Get ElseActions() As Actions
    Set ElseActions = m_elseActions
End Property ' ElseActions

Friend Property Get ThenActions() As Actions
    Set ThenActions = m_thenActions
End Property ' ThenActions
