VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ValidationTest
' Purpose:      Stores a Validation test that is run against the loaded 'rda instructions.xml' file.
'               An instance of this class forms part of the Configuration object heirarchy.
'
' Author:       Peter Hewett - Inner Word Limited (innerword@xnet.co.nz)
' Copyright:    Ministry of Social Development (MSD) ©2015 All rights reserved.
' Contact       Inner Word Limited
' details:      134 Kahu Road
'               Paremata
'               Porirua City
'               5024
'               T: +64 4 233 2124
'               M: +64 21 213 5063
'               E: innerword@xnet.co.nz
'
' History:      05/11/15    1.  Created.
'===================================================================================================================================
Option Explicit

' ANV = Attribute Names for the 'validate' node
Private Const mc_ANVManifestVersion   As String = "manifestVersion"
Private Const mc_ANVQuery             As String = "query"
Private Const mc_ANVExpectedResult    As String = "expectedResult"

' Node Values for the 'expectedResult' node
Private Const mc_NVExpectedResultNone As String = "None"
Private Const mc_NVExpectedResultSome As String = "Some"


Private m_xpathQuery      As String
Private m_manifestVersion As Long
Private m_expectedResult  As mgrValidationResult


Friend Sub Initialise(ByRef validateNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "ValidationTest.Initialise"

    Dim errorText As String
    Dim theAttribute As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out the 'validate' nodes attributes
    For Each theAttribute In validateNode.Attributes

        ' Parse each attribute
        Select Case theAttribute.BaseName
        Case mc_ANVManifestVersion
            m_manifestVersion = CLng(theAttribute.Text)

        Case mc_ANVQuery
            m_xpathQuery = theAttribute.Text

        Case mc_ANVExpectedResult
            Select Case theAttribute.Text
            Case mc_NVExpectedResultNone
                m_expectedResult = mgrValidationResultNone

            Case mc_NVExpectedResultSome
                m_expectedResult = mgrValidationResultSome

            Case Else
                errorText = Replace$(mgrErrTextInvalidAttributeValue, mgrP1, theAttribute.BaseName)
                errorText = Replace$(errorText, mgrP2, theAttribute.Text)
                Err.Raise mgrErrNoInvalidAttributeValue, c_proc, errorText

            End Select

        Case Else
            errorText = Replace$(mgrErrTextInvalidAttributeName, mgrP1, theAttribute.BaseName)
            Err.Raise mgrErrNoInvalidAttributeName, c_proc, errorText

        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

Friend Property Get ExpectedResult() As mgrValidationResult
    ExpectedResult = m_expectedResult
End Property ' Get ExpectedResult

Friend Property Get ManifestVersion() As Long
    ManifestVersion = m_manifestVersion
End Property ' Get ManifestVersion

Friend Property Get Query() As String
    Query = m_xpathQuery
End Property ' Get Query
