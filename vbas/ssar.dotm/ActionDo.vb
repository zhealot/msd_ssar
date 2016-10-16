VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionDo
' Purpose:      Data class that models the 'do' action node of the 'instructions.xml' file.
'
' Note 1:       The 'do' Actions sets up a basic loop construct. It's been bastardised to perform For loop functionality that
'               really should have formed an ActionFor class. But time is short and needs must, so apologies to whoever is
'               maintaining this!

' Note 2:       The number of loop iterations is the number of nodes returned by the xpath query 'condition' or by the 'start' and
'               'end' attributes.
'
' Note 3:       The following attributes are available for use with this Action:
'               condition*
'                   An xpath query that returns a list of nodes. There is one 'do' iteration for each node in the list.
'               start*
'                   A numeric value or the name of a named counter that specifies the do loop start index.
'               end*
'                   A numeric value or the name of a named counter that specifies the do loop end index.
'               reverse [Optional]
'                   When True, the loop is iterated from highest value to lowest value (eg: 10 To 1, instead of 1 To 10).
'               indexAdjustment [Optional]
'                   This is a counter name. The counters value is added to the internal index numbers generated from the number of
'                   nodes returned by the query.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
'
'               Note *: If 'condition' is specified 'start' and 'end' must be omitted.
'                       If 'start' is specified 'end' must also be specified.
'                       You cannot specify 'indexAdjustment' when using 'start' and 'end'.
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
' History:      27/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBreak                    As String = "break"             ' Break if value is True
Private Const mc_ANCondition                As String = "condition"         ' The xpath query
Private Const mc_ANEnd                      As String = "end"               ' A numeric value or the name of a named counter
Private Const mc_ANIndexAdjustment          As String = "indexAdjustment"   ' A named counter whose value is used to adjust the loop index value
Private Const mc_ANReverse                  As String = "reverse"           ' Reverse the order of the loop iteration
Private Const mc_ANStart                    As String = "start"             ' A numeric value or the name of a named counter


Private Enum doType
    doTypeCondition
    doTypeStartEnd
End Enum ' doType


Private m_break                         As Boolean                          ' Used to cause the code to execute a Stop instruction
Private m_condition                     As String                           ' The xpath query that determins how many loop iterations occur
Private m_doType                        As doType                           ' The type of do loop that will be performed
Private m_end                           As Long                             ' A numeric value that specifies the loop end index
Private m_endCounter                    As String                           ' The named counter that specifies the loop end index
Private m_indexAdjustmentCounterName    As String                           ' The named counter to use for index offset adjustmnent
Private m_reverse                       As Boolean                          ' When True iterate the loop in revrse order (highest to lowest)
Private m_start                         As Long                             ' A numeric value that specifies the loop start index
Private m_startCounter                  As String                           ' The named counter that specifies the loop start index
Private m_subActions                    As Actions                          ' Nested Actions in the order they occur.

'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Do action instruction.
' Date:         27/06/16    Created.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionDo.IAction_Parse"

    Dim conditionCount  As Long
    Dim endCount        As Long
    Dim indexCount      As Long
    Dim reverseCount    As Long
    Dim startCount      As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out the 'do' nodes attributes
    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANCondition
            conditionCount = conditionCount + 1
            m_condition = theAttribute.Text

        Case mc_ANStart
            startCount = startCount + 1

            ' If value is non numeric then it must be a named counter
            If IsNumeric(theAttribute.Text) Then
                m_start = CLng(theAttribute.Text)
            Else
                m_startCounter = Trim$(theAttribute.Text)
                If LenB(m_startCounter) = 0 Then
                    InvalidAttributeValueError actionNode.nodeName, mc_ANStart, m_startCounter, c_proc
                End If
            End If

        Case mc_ANEnd
            endCount = endCount + 1

            ' If value is non numeric then it must be a named counter
            If IsNumeric(theAttribute.Text) Then
                m_end = CLng(theAttribute.Text)
            Else
                m_endCounter = Trim$(theAttribute.Text)
                If LenB(m_endCounter) = 0 Then
                    InvalidAttributeValueError actionNode.nodeName, mc_ANEnd, m_endCounter, c_proc
                End If
            End If

        Case mc_ANIndexAdjustment
            indexCount = indexCount + 1
            m_indexAdjustmentCounterName = theAttribute.Text

        Case mc_ANReverse
            reverseCount = reverseCount + 1
            m_reverse = CBool(theAttribute.Text)

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    'Make sure there are no duplicate attributes
    If conditionCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANCondition, c_proc
    ElseIf startCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANStart, c_proc
    ElseIf endCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANEnd, c_proc
    ElseIf indexCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANIndexAdjustment, c_proc
    ElseIf reverseCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANReverse, c_proc
    End If

    ' Make sure the attributes that are present contain some data
    If conditionCount = 1 And LenB(m_condition) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANCondition, m_condition, c_proc
    ElseIf indexCount = 1 And LenB(m_indexAdjustmentCounterName) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANIndexAdjustment, m_indexAdjustmentCounterName, c_proc
    End If

    ' Make sure mutually exclusive attributes are not present
    If conditionCount > 0 Then
        If startCount > 0 Then
            InvalidAttributeCombinationError actionNode.nodeName, mc_ANStart, c_proc
        ElseIf endCount > 0 Then
            InvalidAttributeCombinationError actionNode.nodeName, mc_ANEnd, c_proc
        End If
    ElseIf startCount = 1 And endCount = 1 And indexCount = 1 Then
        InvalidAttributeCombinationError actionNode.nodeName, mc_ANIndexAdjustment, c_proc
    End If

    ' Make sure the Mandatory attributes are present
    If startCount = 0 And conditionCount = 0 Then
        MissingAttributeError actionNode.nodeName, "condition or start/end", c_proc
    End If

    ' If start was not specified end must not have been specified, if start was specified end must be specified
    If CBool(startCount Xor endCount) Then
        If startCount = 0 Then
            MissingAttributeError actionNode.nodeName, mc_ANStart, c_proc
        Else
            MissingAttributeError actionNode.nodeName, mc_ANEnd, c_proc
        End If
    End If

    ' Set the loop type
    If LenB(m_condition) > 0 Then
        m_doType = doTypeCondition
    Else
        m_doType = doTypeStartEnd
    End If

    ' Parse out all of the Do nodes child nodes (sub-actions)
    Set m_subActions = New Actions
    m_subActions.Parse actionNode, addPatternBookmarks

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Allows the instruction file to perform simple loops while Building the Assessment Report.
' Notes:        The number of loop iterations is governed by the number of nodes returned by the xpath query.
' Date:         27/06/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionDo.IAction_BuildAssessmentReport"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Process the specified do type
    Select Case m_doType
    Case doTypeCondition
        ConditionDo ssarIActionMethodBuildAssessmentReport
    Case doTypeStartEnd
        StartEndDo ssarIActionMethodBuildAssessmentReport
    End Select

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

'===================================================================================================================================
' Procedure:    IAction_ConstructRichText
' Purpose:      Allows the instruction file to perform simple loops while creating the rich text data source.
' Date:         27/06/16    Created.
'===================================================================================================================================
Private Sub IAction_ConstructRichText()
    Const c_proc As String = "ActionDo.IAction_ConstructRichText"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Process the specified do type
    Select Case m_doType
    Case doTypeCondition
        ConditionDo ssarIActionMethodRichText
    Case doTypeStartEnd
        StartEndDo ssarIActionMethodRichText
    End Select

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_ConstructRichText

'===================================================================================================================================
' Procedure:    IAction_HTMLForXMLUpdate
' Purpose:      Allows the instruction file to perform simple loops while creating the HTML to XHTML data source.
' Date:         05/08/16    Created.
'===================================================================================================================================
Private Sub IAction_HTMLForXMLUpdate()
    Const c_proc As String = "ActionDo.IAction_HTMLForXMLUpdate"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Process the specified do type
    Select Case m_doType
    Case doTypeCondition
        ConditionDo ssarIActionMethodHTMLForXMLUpdate
    Case doTypeStartEnd
        StartEndDo ssarIActionMethodHTMLForXMLUpdate
    End Select

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub '  Sub IAction_HTMLForXMLUpdate

'===================================================================================================================================
' Procedure:    IAction_UpdateDateXML
' Purpose:      Allows the instruction file to perform simple loops while updating the corresponding xml of any editable area that
'               contains a date.
' Note 1:       Before the xml update occurs the dates are validated to ensure that the user has entered a valid date.
' Date:         04/08/16    Created.
'===================================================================================================================================
Private Sub IAction_UpdateDateXML()
    Const c_proc As String = "ActionDo.IAction_UpdateDateXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Process the specified do type
    Select Case m_doType
    Case doTypeCondition
        ConditionDo ssarIActionMethodUpdateDateXML
    Case doTypeStartEnd
        StartEndDo ssarIActionMethodUpdateDateXML
    End Select

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Sub IAction_UpdateDateXML

Private Sub IAction_UpdateContentControlXML()
    Const c_proc As String = "ActionDo.IAction_UpdateDateXML"

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Process the specified do type
    Select Case m_doType
    Case doTypeCondition
        ConditionDo ssarIActionMethodUpdateContentControlXML
    Case doTypeStartEnd
        StartEndDo ssarIActionMethodUpdateContentControlXML
    End Select

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_UpdateContentControlXML

'===================================================================================================================================
' Procedure:    ConditionDo
' Purpose:      Performs a do loop based on the number of nodes returned by an xpath query.
' Date:         12/07/16    Created.
'
' On Entry:     iactionMethod       The actual IAction procedure to call.
'===================================================================================================================================
Private Sub ConditionDo(ByVal iactionMethod As ssarIActionMethod)
    Const c_proc As String = "ActionDo.ConditionDo"

    Dim index           As Long
    Dim indexAdjustment As Long
    Dim loopEnd         As Long
    Dim loopStart       As Long
    Dim loopStep        As Long
    Dim theQuery        As String
    Dim subNodes        As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(m_condition)

    ' Use the condition string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' See if index adjustment is required (allows for the nodes retrieved by the query not being the first nodes
        ' in the sequence required). So if there are precceding nodes of the same type they can be skipped over.
        If LenB(m_indexAdjustmentCounterName) > 0 Then

            ' Get the index adjustment value from the named counter
            indexAdjustment = g_actionCounters.Item(m_indexAdjustmentCounterName)
        End If

        ' Set up variables for the loop iteration order. Either lowest to highest or highest to lowest.
        If m_reverse Then
            loopStart = subNodes.Length
            loopEnd = 1
            loopStep = -1
        Else
            loopStart = 1
            loopEnd = subNodes.Length
            loopStep = 1
        End If

        ' Iterate all retrieved nodes
        For index = loopStart To loopEnd Step loopStep

            ' Increment the counter for the current level. Always include the index adjustment as it has a zero value if unused.
            g_counters.Counter = index + indexAdjustment

            ' Perform any nested actions in the current ActionDo object
            Select Case iactionMethod
            Case ssarIActionMethodBuildAssessmentReport
                m_subActions.BuildAssessmentReport
            Case ssarIActionMethodRichText
                m_subActions.ConstructRichText
            Case ssarIActionMethodUpdateContentControlXML
                m_subActions.UpdateContentControlXML
            Case ssarIActionMethodUpdateDateXML
                m_subActions.UpdateDateXML
            Case ssarIActionMethodHTMLForXMLUpdate
                m_subActions.HTMLForXMLUpdate
            End Select
        Next
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ConditionDo

'===================================================================================================================================
' Procedure:    StartEndDo
' Purpose:      Performs a do loop based on the specified start and end values.
' Notes:        The start and end values can either be spoecified numeric values or the name of a named counter.
' Date:         12/07/16    Created.
'
' On Entry:     iactionMethod       The actual IAction procedure to call.
'===================================================================================================================================
Private Sub StartEndDo(ByVal iactionMethod As ssarIActionMethod)
    Const c_proc As String = "ActionDo.StartEndDo"

    Dim endIndex        As Long
    Dim index           As Long
    Dim loopEnd         As Long
    Dim loopStart       As Long
    Dim loopStep        As Long
    Dim startIndex      As Long

    On Error GoTo Do_Error

    ' Get the start index value, use the named counter if one was specified
    If LenB(m_startCounter) > 0 Then
        startIndex = g_actionCounters.Item(m_startCounter)
    Else
        startIndex = m_start
    End If

    ' Get the end index value, use the named counter if one was specified
    If LenB(m_endCounter) > 0 Then
        endIndex = g_actionCounters.Item(m_endCounter)
    Else
        endIndex = m_end
    End If

    ' Set up variables for the loop iteration order. Either lowest to highest or highest to lowest.
    If m_reverse Then
        loopStart = endIndex
        loopEnd = startIndex
        loopStep = -1
    Else
        loopStart = startIndex
        loopEnd = endIndex
        loopStep = 1
    End If

    ' Iterate the specified number of times
    For index = loopStart To loopEnd Step loopStep

        ' Increment the counter for the current level
        g_counters.Counter = index

        ' Perform any nested actions in the current ActionDo object
        Select Case iactionMethod
        Case ssarIActionMethodBuildAssessmentReport
            m_subActions.BuildAssessmentReport
        Case ssarIActionMethodRichText
            m_subActions.ConstructRichText
        Case ssarIActionMethodUpdateContentControlXML
            m_subActions.UpdateContentControlXML
        Case ssarIActionMethodUpdateDateXML
            m_subActions.UpdateDateXML
        Case ssarIActionMethodHTMLForXMLUpdate
            m_subActions.HTMLForXMLUpdate
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' StartEndDo

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeDo
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break

Friend Property Get IndexAdjustmentCounterName() As String
    IndexAdjustmentCounterName = m_indexAdjustmentCounterName
End Property ' Get IndexAdjustmentCounterName

Friend Property Get Condition() As String
    Condition = m_condition
End Property ' Get Condition

Friend Property Get Reverse() As Boolean
    Reverse = m_reverse
End Property ' Get Reverse
