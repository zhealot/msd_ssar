Attribute VB_Name = "modAvailableToAddins"
'===================================================================================================================================
' Module:       modAvailableToAddins
' Purpose:      Contains public procedures that are available to the RDA and SSAR addins.
' Note:         The code housed here would normally fit in one of the other modules, but because that module is using Option Private
'               Module the addin cannot see it. So here we house procedures that do not fit with their parent modules scope rules.
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
' History:      31/05/2016    1.  Created.
'===================================================================================================================================
Option Explicit

' AN = Attribute Name
Private Const mc_ANDataChanged          As String = "dataChanged"

' AVDC = Attribute Value DataChanged
Private Const mc_AVDCNo                 As String = "No"
Private Const mc_AVDCYes                As String = "Yes"


'===================================================================================================================================
' Procedure:    ADDocVarExists
' Purpose:      Determines whether the passed in document variable name exists in the ActiveDocument.
' Date:         19/08/16    Created.
'
' Returns:      True if the passed in document variable exists.
'===================================================================================================================================
Public Function ADDocVarExists(ByVal docVarName As String) As Boolean
    Dim dummy   As String

    On Error GoTo Do_Error
    dummy = ActiveDocument.Variables(docVarName).Value
    ADDocVarExists = True

Do_Exit:
    Exit Function

Do_Error:
    Err.Clear
    Resume Do_Exit
End Function ' ADDocVarExists

'===================================================================================================================================
' Procedure:    ARDocVarExists
' Purpose:      Determines whether the passed in document variable name exists in the assessment report.
' Date:         19/08/16    Created.
'
' Returns:      True if the passed in document variable exists.
'===================================================================================================================================
Public Function ARDocVarExists(ByVal docVarName As String) As Boolean
    Dim dummy   As String

    On Error GoTo Do_Error
    dummy = g_assessmentReport.Variables(docVarName).Value
    ARDocVarExists = True

Do_Exit:
    Exit Function

Do_Error:
    Err.Clear
    Resume Do_Exit
End Function ' ARDocVarExists

'===================================================================================================================================
' Procedure:    ADDocVarBooleanValue
' Purpose:      Returns the boolean value of the passed in document variable in the ActiveDocument.
' Note 1:       Implicit in this is that the document variable contains a value that can be coerced to a boolean value.
' Date:         19/08/16    Created.
'
' Returns:      True if the document variable contained a value that could be coerced into a True boolean value.
'               False, if there was an error or the document variable contained a False coerceable value.
'===================================================================================================================================
Public Function ADDocVarBooleanValue(ByVal docVarName As String) As Boolean
    Dim dummy   As String

    On Error Resume Next
    dummy = ActiveDocument.Variables(docVarName).Value
    ADDocVarBooleanValue = CBool(dummy)
    Err.Clear
End Function ' ADDocVarBooleanValue

'===================================================================================================================================
' Procedure:    ARDocVarBooleanValue
' Purpose:      Returns the boolean value of the passed in document variable in the assessment report.
' Note 1:       Implicit in this is that the document variable contains a value that can be coerced to a boolean value.
' Date:         19/08/16    Created.
'
' Returns:      True if the document variable contained a value that could be coerced into a True boolean value.
'               False, if there was an error or the document variable contained a False coerceable value.
'===================================================================================================================================
Public Function ARDocVarBooleanValue(ByVal docVarName As String) As Boolean
    Dim dummy   As String

    On Error Resume Next
    dummy = g_assessmentReport.Variables(docVarName).Value
    ARDocVarBooleanValue = CBool(dummy)
    Err.Clear
End Function ' ARDocVarBooleanValue

'===================================================================================================================================
' Procedure:    ADDocVarValue
' Purpose:      Returns the value of the passed in document variable in the ActiveDocument.
' Date:         19/08/16    Created.
'
' Returns:      The document variable value or a null string if the document variable did not exist.
'===================================================================================================================================
Public Function ADDocVarValue(ByVal docVarName As String) As String
    On Error Resume Next
    ADDocVarValue = ActiveDocument.Variables(docVarName).Value
    Err.Clear
End Function ' ADDocVarValue

'===================================================================================================================================
' Procedure:    ARDocVarValue
' Purpose:      Returns the value of the passed in document variable in the assessment report.
' Date:         19/08/16    Created.
'
' Returns:      The document variable value or a null string if the document variable did not exist.
'===================================================================================================================================
Public Function ARDocVarValue(ByVal docVarName As String) As String
    On Error Resume Next
    ARDocVarValue = g_assessmentReport.Variables(docVarName).Value
    Err.Clear
End Function ' ARDocVarValue

'===================================================================================================================================
' Procedure:    DoesNodePathHaveData
' Purpose:      Determines whether the node itself or any of its descendant nodes contain any data.
' Note 1:       Attribute data is excluded.
' Note 2:       Would normally be housed in modXML.
'
' On Entry:     subNode             The node to start the determination at.
' Returns:      True, if the node or one of its descendant node contains any data.
'===================================================================================================================================
Public Function DoesNodePathHaveData(ByVal subNode As MSXML2.IXMLDOMNode) As Boolean
    Dim child   As MSXML2.IXMLDOMNode
    Dim hasData As Boolean

    ' Make sure something has been passed in
    If subNode Is Nothing Then
        Exit Function
    End If

    ' Check to see if the passed in subNode contains data
    If subNode.Text <> vbNullString Then
        DoesNodePathHaveData = True

    Else

        ' Now check to see if any of the child nodes contain data
        For Each child In subNode.ChildNodes

            ' Check the current nodes child nodes
            hasData = DoesNodePathHaveData(child)
            If hasData Then
                DoesNodePathHaveData = hasData
                Exit For
            End If
        Next
    End If
End Function ' DoesNodePathHaveData

'===================================================================================================================================
' Procedure:    IsAssessmentReport
' Purpose:      Determines whether the current Word Document is an assessment report.
' Note 2:       Would normally be housed in modUtility.
'===================================================================================================================================
Public Function IsAssessmentReport() As Boolean
    Dim dummy As String

    If Documents.Count > 0 Then

        ' Is the active document an Assessment Report
        On Error Resume Next
        dummy = ActiveDocument.Variables(mgrDVAssessmentReport)
        If Err.Number = 0 Then
            IsAssessmentReport = True
        Else
            Err.Clear
        End If
    End If
End Function ' IsAssessmentReport

'=======================================================================================================================
' Procedure:    LoadAssessmentReport
' Purpose:      Loads the 'rda-asmntrept.rep' file.
' Note 1:       Despite its name 'rda-asmntrept.rep' (generated by Remedy) is actually an html file. The xml we need is
'               embedded within one of the html elements.
' Note 2:       The xml requires correction before we can use it as certain escape sequences have been replaced with
'               their actual character equivalent, whereas we still need the escape sequence to form valid xml.
' Note 3:       Once we have the valid assessment report xml, it is then loaded into an xml DOM Document object.
' Note 4:       We can then interrogate the xml DOMDocument to find out which of the three different assessment report
'               xml versions has been loaded.
' Note 5:       Each different assessment report xml version has its own schema file.
' Note 6:       The schema file appropriate to the assessment report xml version is then loaded and used to validate
'               the assessment report xml against.
' Note 7:       Would normally be housed in modXML.
'
' On Entry:     repFileFullPath     The name and path of the '.rep' Remedy html input file.
' Returns:      True if the .rep input file has been loaded and parsed correctly and the xml successfully extracted.
'=======================================================================================================================
Public Function LoadAssessmentReport(ByVal repFileFullPath As String) As Boolean
    Const c_proc As String = "modXML.LoadAssessmentReport"

    Dim errorText As String

    On Error GoTo Do_Error

    EventLog "Loading: " & repFileFullPath, c_proc

    ' Load the html rep file and extract and clean the xml
    If LoadRemedyHTMLFile(repFileFullPath, g_xmlDocument) Then

        ' Initialise the g_rootData object which contains critical information
        ' specific to the fetchdoc.infopathxml file that has been loaded
        Set g_rootData = New RootData
        g_rootData.Initialise

        '
        If ValidateXMLFile(g_xmlDocument, g_configuration.SchemaFullName) Then
            LoadAssessmentReport = True
        Else
            errorText = Replace$(mgrErrTextXMLSchemaLoad, mgrP1, g_configuration.SchemaFullName)
            XMLErrorReporter g_xmlDocument.parseError, errorText, c_proc
            Exit Function
        End If

    Else
        errorText = Replace$(mgrErrTextRepDocumentLoad, mgrP1, repFileFullPath)
        XMLErrorReporter g_xmlDocument.parseError, errorText, c_proc
        Exit Function
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    LoadAssessmentReport = False
    Resume Do_Exit
End Function ' LoadAssessmentReport

'=======================================================================================================================
' Procedure:    ParseAttributes
' Purpose:      Parses out the passed in nodes attribute values into the supplied ParamArray.
' Note 1:       The ParamArray must have 2 or more elements and must have an even number of elements. The first element
'               of a pair is the attribute name and the second is the attribute value to be returned.
' Note 2:       Would normally be housed in modXML.
'
' On Entry:     theNode             The node containing the attributes to be parsed out.
'               attNameVariable     A ParamaArray of name/value data pairs.
' On Exit:      attNameVariable     The value part of each data pair will have been updated with the attributes value
'                                   if the attribute was present.
'=======================================================================================================================
Public Sub ParseAttributes(ByVal theNode As MSXML2.IXMLDOMNode, _
                           ParamArray attNameVariable() As Variant)
    Const c_proc As String = "modXMLPrivate.ParseAttributes"

    Dim elementCount As Long
    Dim errorText    As String
    Dim found        As Boolean
    Dim index        As Long
    Dim nextElement  As Long
    Dim theAttribute As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Make sure the Attribute Name and Variable array has an even number of elements. The first
    ' for the attribute name and the second for the variable to receive the attributes value.

    ' Calculate how many elements were passed in (there must be an even number, as the arguments must be in pairs)
    elementCount = UBound(attNameVariable) - LBound(attNameVariable) + 1
    If elementCount > 0 And elementCount Mod 2 = 0 Then

        ' Parse out the nodes attributes
        For Each theAttribute In theNode.Attributes

            ' Clear the attribute found flag
            found = False

            ' Because we have a zero based array, attNameVariable(1) is the second element
            For index = LBound(attNameVariable) To UBound(attNameVariable) Step 2

                ' Compare the current attributes name with the current ParamArray item (which should be an attribute name)
                If theAttribute.BaseName = attNameVariable(index) Then

                    ' Update the Variable element of the Attribute Name and Variable pair
                    nextElement = index + 1
                    If VarType(attNameVariable(nextElement)) = vbBoolean Then
                        attNameVariable(nextElement) = CBool(theAttribute.Text)
                    Else
                        attNameVariable(nextElement) = theAttribute.Text
                    End If

                    ' Set the flag to indicate we found the current attribute (which suppresses the error)
                    found = True
                    Exit For
                End If
            Next

            ' Complain if we could not find the specified attribute
            If Not found Then
                errorText = Replace$(mgrErrTextInvalidAttributeName, mgrP1, theAttribute.BaseName)
                Err.Raise mgrErrNoInvalidAttributeName, c_proc, errorText
            End If
        Next
    Else
        Err.Raise mgrErrNoInvalidParamArray, c_proc, mgrErrTextInvalidParamArray
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseAttributes

'=======================================================================================================================
' Procedure:    SetDataChangedAttribute
' Purpose:      Sets the assessment reports dataChanged attribute to indicate whether the data being sent to the
'               webservice has been altered or not.
' Note 1:       The 'dataChanged' attribute is part of the 'Assessment' node.
' Note 2:       Would normally be housed in modXML.
'
' On Entry:     dataState           True if the data has been changed.
'=======================================================================================================================
Public Sub SetDataChangedAttribute(ByVal dataState As Boolean)
    Const c_proc As String = "modXML.SetDataChangedAttribute"

    Dim attributeValue  As String
    Dim child           As MSXML2.IXMLDOMNode
    Dim rootNode        As MSXML2.IXMLDOMNode
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Find the root node ('Asssessment') by traversing all child nodes until we hit the first node of nodeType NODE_ELEMENT.
    ' All other nodes will be of nodeType NODE_PROCESSING_INSTRUCTION.
    For Each child In g_xmlDocument.ChildNodes

        If child.NodeType = NODE_ELEMENT Then
            Set rootNode = child
            Exit For
        End If
    Next

    ' Now parse out the 'Assessment' node attributes until we find the 'dataChanged' attribute
    For Each theAttribute In rootNode.Attributes
        If theAttribute.BaseName = mc_ANDataChanged Then

            ' Set the value for the dataChange attribute to one of the two expected values (Yes/No)
            If dataState Then
                attributeValue = mc_AVDCYes
            Else
                attributeValue = mc_AVDCNo
            End If

            ' Exit the loop now that we have found the attribute we were looking for
            Exit For
        End If
    Next

    ' Make sure we found the 'dataChanged' attribute
    If LenB(attributeValue) Then

        ' Update the attribute nodes value
        theAttribute.Text = attributeValue
    Else
        Err.Raise mgrErrNoUnexpectedCondition, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' SetDataChangedAttribute

