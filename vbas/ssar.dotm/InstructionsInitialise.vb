VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InstructionsInitialise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       InstructionsInitialise
' Purpose:      Contains non Action specific information that needs to be present before processing the Actions.
' Note 1:       Use this to set up anything that requires non instruction specific data.
' Note 2:       The following attributes are available for use with this Instruction:
'               ccRootNode [Mandatory]
'                   Defines the root node name used in the Custom XML Part used for Content Control data mapping.
'               colourMapping
'
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
' History:      18/06/16    1.  Created.
'===================================================================================================================================
Option Explicit

' NN = Node Name
Private Const mc_NNCCCustomXMLPart  As String = "ccCustomXMLPart"
Private Const mc_NNColourMapping    As String = "colourMapping"

' ANCC = Attribute Name Content Control (Custom XML Part)
Private Const mc_ANCCRootNode       As String = "rootNode"

' ANCM = Attribute Names for the Colour Map node
Private Const mc_ANCMBackground     As String = "background"
Private Const mc_ANCMForeground     As String = "foreground"
Private Const mc_ANCMValue          As String = "value"


'===================================================================================================================================
' Procedure:    Parse
' Purpose:      Parses the 'initialise' node instruction.
' Date:         19/06/16    Created.
'
' On Entry:     initialiseNode      The 'initialise' xml node to parse.
'===================================================================================================================================
Friend Sub Parse(ByVal initialiseNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "InstructionsInitialise.Parse"

    Dim child   As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    For Each child In initialiseNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse the child nodes ()
            Select Case child.BaseName
            Case mc_NNCCCustomXMLPart

                ' Parse out the CCCustomXMLPart node
                ParseCCCustomXMLPart child

            Case mc_NNColourMapping

                ' Parse out the colouMap node
                ParseColourMapping child

            Case Else
                InvalidActionVerbError child.BaseName, c_proc
            End Select
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

'===================================================================================================================================
' Procedure:    ParseCCCustomXMLPart
' Purpose:      Parses the 'ccCustomXMLPart' child node.
' Date:         18/06/16    Created.
'
' On Entry:     ccCustomXMLPartNode The 'ccCustomXMLPart' xml node to parse.
'===================================================================================================================================
Private Sub ParseCCCustomXMLPart(ByVal ccCustomXMLPartNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "InstructionsInitialise.ParseCCCustomXMLPart"

    Dim rootNodeCount   As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse the 'ccCustomXMLPart' node
    For Each theAttribute In ccCustomXMLPartNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANCCRootNode

            ' Perform basic validation of the attribute and its text
            rootNodeCount = rootNodeCount + 1
            If rootNodeCount = 1 Then
                If LenB(theAttribute.Text) > 0 Then

                    ' Create and initialise the Content Control Data Store object
                    Set g_ccXMLDataStore = New CCXMLDataStore
                    g_ccXMLDataStore.Initialise theAttribute.Text
                Else
                    InvalidAttributeValueError ccCustomXMLPartNode.nodeName, theAttribute.BaseName, theAttribute.Text, c_proc
                End If
            Else
                DuplicateAttributeError ccCustomXMLPartNode.nodeName, theAttribute.BaseName, c_proc
            End If

        Case Else
            InvalidAttributeNameError ccCustomXMLPartNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    If rootNodeCount = 0 Then
        MissingAttributeError ccCustomXMLPartNode.nodeName, mc_ANCCRootNode, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseCCCustomXMLPart

'===================================================================================================================================
' Procedure:    ParseColourMapping
' Purpose:      Parses the 'colourMapping' child node.
' Date:         18/06/16    Created.
'
' On Entry:     colourMappingNode   The 'colourMapping' xml node to parse.
'===================================================================================================================================
Private Sub ParseColourMapping(ByVal colourMappingNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "InstructionsInitialise.ParseColourMapping"

    Dim backgroundColour        As Long
    Dim backgroundColourName    As String
    Dim backgroundCount         As Long
    Dim foregroundColour        As Long
    Dim foregroundColourName    As String
    Dim foregroundCount         As Long
    Dim nodeName                As String
    Dim theAttribute            As MSXML2.IXMLDOMNode
    Dim theValue                As String
    Dim valueCount              As Long

    On Error GoTo Do_Error

    nodeName = colourMappingNode.nodeName

    ' Parse out the 'colourMapping' nodes attributes
    For Each theAttribute In colourMappingNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANCMValue
            valueCount = valueCount + 1
            theValue = theAttribute.Text

        Case mc_ANCMForeground
            foregroundCount = foregroundCount + 1
            foregroundColourName = theAttribute.Text

        Case mc_ANCMBackground
            backgroundCount = backgroundCount + 1
            backgroundColourName = theAttribute.Text
        Case Else
            InvalidAttributeNameError nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure these attributes are present
    If valueCount = 0 Then
        MissingAttributeError nodeName, mc_ANCMValue, c_proc
    ElseIf foregroundCount = 0 Then
        MissingAttributeError nodeName, mc_ANCMForeground, c_proc
    ElseIf backgroundCount = 0 Then
        MissingAttributeError nodeName, mc_ANCMBackground, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If valueCount > 1 Then
        DuplicateAttributeError nodeName, mc_ANCMValue, c_proc
    ElseIf foregroundCount > 1 Then
        DuplicateAttributeError nodeName, mc_ANCMForeground, c_proc
    ElseIf backgroundCount > 1 Then
        DuplicateAttributeError nodeName, mc_ANCMBackground, c_proc
    End If

    ' Make sure the attributes contain some data
    If LenB(theValue) = 0 Then
        InvalidAttributeValueError nodeName, mc_ANCMValue, theValue, c_proc
    ElseIf LenB(foregroundColourName) = 0 Then
        InvalidAttributeValueError nodeName, mc_ANCMForeground, foregroundColourName, c_proc
    ElseIf LenB(backgroundColourName) = 0 Then
        InvalidAttributeValueError nodeName, mc_ANCMBackground, backgroundColourName, c_proc
    End If

    ' Validate and translate to composite colours the colourMap values
    foregroundColour = ValidateAndTranslateColour(foregroundColourName, nodeName, mc_ANCMForeground, c_proc)
    backgroundColour = ValidateAndTranslateColour(backgroundColourName, nodeName, mc_ANCMBackground, c_proc)

    ' Got through the validation, so see if the dictionary objects need creating
    If g_colourMapForeground Is Nothing Then
        Set g_colourMapForeground = New Scripting.Dictionary
        Set g_colourMapBackground = New Scripting.Dictionary
    End If

    ' Add the composite colour value for the colour maps to the two colour map dictionary objects
    g_colourMapForeground.Add theValue, foregroundColour
    g_colourMapBackground.Add theValue, backgroundColour

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseColourMapping

'===================================================================================================================================
' Procedure:    ValidateAndTranslateColour
' Purpose:      Validates the passed in colour name and returns the numeric colour value for that colour.
' Notes:        It is intentional to have no error handler so that the caller directly handles any errors.
' Date:         18/06/16    Created.
'
' On Entry:     colourName          The name of the colour to validate and translate.
'               nodeName            The xml node name to be used if there is an error.
'               attributeName       The xml attribute name to be used if there is an error.
'               moduleProcedure     The name of the module and procedure to be reported if there is an error.
' Returns:      The composite colour value of the passed in colour name.
'===================================================================================================================================
Private Function ValidateAndTranslateColour(ByVal colourName As String, _
                                            ByVal nodeName As String, _
                                            ByVal attributeName As String, _
                                            ByVal moduleProcedure As String) As Long

    ' Translate the colour name to the composite colour value
    Select Case colourName
    Case ssarColourNameBlack
        ValidateAndTranslateColour = ssarColourBlack
    Case ssarColourNameYellow
        ValidateAndTranslateColour = ssarColourYellow
    Case ssarColourNameGreen
        ValidateAndTranslateColour = ssarColourGreen
    Case ssarColourNameRed
        ValidateAndTranslateColour = ssarColourRed
    Case ssarColourNameWhite
        ValidateAndTranslateColour = ssarColourWhite
    Case ssarColourNameBlue
        ValidateAndTranslateColour = ssarColourBlue
    Case ssarColourNameAqua                             '26/04/2017, tao@allfields.co.nz
        ValidateAndTranslateColour = ssarColourAqua
    Case Else

        ' Failed validation...
        InvalidAttributeValueError nodeName, attributeName, colourName, moduleProcedure
    End Select
End Function ' ValidateAndTranslateColour
