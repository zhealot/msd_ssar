VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RootData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===================================================================================================================================
' Class:        RootData
' Purpose:      Interface module for the Assessment xml nodes attributes.
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
' History:      12/05/16    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit

' NN = NodeName
Private Const mc_NNAssessmentReportVersion As String = "/Assessment/assessmentReportVersion"

' AN = Attribute Name
Private Const mc_ANDataChanged             As String = "dataChanged"
Private Const mc_ANEnvironment             As String = "environment"
Private Const mc_ANMainSaveURL             As String = "mainSaveURL"
Private Const mc_ANNoNamespace             As String = "noNamespaceSchemaLocation"
Private Const mc_ANOpenView                As String = "openView"
Private Const mc_ANRemedyServerName        As String = "remedyServerName"
Private Const mc_ANRemedyServerPort        As String = "remedyServerPort"
Private Const mc_ANRemedyUserId            As String = "remedyUserId"
Private Const mc_ANTransactionGUID         As String = "transactionGUID"
Private Const mc_ANXhtml                   As String = "xhtml"
Private Const mc_ANXsi                     As String = "xsi"

' OVV = OpenView Values
Private Const mc_OVVNullView              As String = ""                    ' This has been added because of an oddity in the supplied test data
Private Const mc_OVVPrintView             As String = "Print View"
Private Const mc_OVVReadView              As String = "Read Only View"
Private Const mc_OVVWriteView             As String = "Write View"

Private m_initialised      As Boolean
Private m_manifestVersion  As Long
Private m_rootNode         As MSXML2.IXMLDOMNode

' 'Assessment' node attribute values
Private m_dataChanged      As String
Private m_environment      As String
Private m_mainSaveURL      As String
Private m_remedyServerName As String
Private m_remedyServerPort As String
Private m_remedyUserId     As String
Private m_transactionGUID  As String
Private m_viewMode         As mgrViewMode

Private m_reportVersion    As mgrVersionType


Friend Sub Initialise()
    Const c_proc As String = "RootData.Initialise"

    Dim piNode As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Get the processing instruction node that contains the href attribute
    If Not g_xmlDocument Is Nothing Then
        Set piNode = g_xmlDocument.SelectSingleNode(mgrXQFDHrefProcessingInstruction)
    Else
        Err.Raise mgrErrNoInfoPathXMLFileHasNotBeenLoaded, c_proc, mgrErrTextInfoPathXMLFileHasNotBeenLoaded
    End If

    ' Make sure we got the Processing Instruction node
    If piNode Is Nothing Then
        Err.Raise mgrErrNoProcessingInstructionNotFound, c_proc, mgrErrTextProcessingInstructionNotFound
    End If

    ' Parse out the information we need
    GetManifestVersionFromHref piNode

    ' Parse out all of the 'Assessment' nodes attributes
    ParseAssessmentAttributes

    ' Get the Report Version which determines the watermarks displayed
    ParseAssessmentReportVersion

    ' Set initialised flag so that other code does not error
    m_initialised = True

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

Private Sub GetManifestVersionFromHref(ByVal piNode As MSXML2.IXMLDOMNode)
    Const c_proc      As String = "RootData.GetManifestVersionFromHref"
    Const c_hrefStart As String = "href="""
    Const c_hrefEnd   As String = """ name="""

    Dim endPosition   As Long
    Dim errorText     As String
    Dim hrefText      As String
    Dim index         As Long
    Dim parts()       As String
    Dim piText        As String
    Dim startPosition As Long

    On Error GoTo Do_Error

    ' Locate the 'href="' string in the processing instruction
    ' We then need all text between the first double quote character and the last double quote character of the href
    piText = piNode.Text
    startPosition = InStr(piText, c_hrefStart) + Len(c_hrefStart)
    endPosition = InStr(startPosition, piText, c_hrefEnd)
    hrefText = Mid$(piText, startPosition, endPosition - startPosition)

    ' Split the href string into an array of strings
    parts = Split(hrefText, "/")

    ' Iterate the array looking for a purely numeric string (the manifest version number)
    For index = LBound(parts) To UBound(parts)

        ' The version number is the only part that is purely numeric, so we can ignore anything else
        If IsNumeric(parts(index)) Then

            ' Store the manifest version number
            m_manifestVersion = CLng(parts(index))
            Exit Sub
        End If
    Next

    ' No manifest version number found in the href string
    errorText = Replace$(mgrErrTextCouldNotGetManifestVersion, mgrP1, hrefText)
    Err.Raise mgrErrNoCouldNotGetManifestVersion, c_proc, errorText

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' GetManifestVersionFromHref

Private Sub ParseAssessmentAttributes()
    Const c_proc As String = "RootData.ParseAssessmentAttributes"

    Dim child        As MSXML2.IXMLDOMNode
    Dim dummy        As String
    Dim errorText    As String
    Dim openView     As String

    On Error GoTo Do_Error

    ' Find the root node by traversing all child nodes until we hit the first node of nodeType NODE_ELEMENT.
    ' All other nodes will be of nodeType NODE_PROCESSING_INSTRUCTION.
    For Each child In g_xmlDocument.ChildNodes

        If child.NodeType = NODE_ELEMENT Then
            Set m_rootNode = child
            Exit For
        End If
    Next

    ' Parse out the root nodes attributes
    ParseAttributes m_rootNode, mc_ANDataChanged, m_dataChanged, mc_ANEnvironment, m_environment, mc_ANMainSaveURL, m_mainSaveURL, _
                                mc_ANOpenView, openView, mc_ANRemedyServerName, m_remedyServerName, mc_ANRemedyServerPort, m_remedyServerPort, _
                                mc_ANRemedyUserId, m_remedyUserId, mc_ANTransactionGUID, m_transactionGUID, _
                                mc_ANXhtml, dummy, mc_ANXsi, dummy, mc_ANNoNamespace, dummy

    ' Added because of oddities in test data, so it should not occur in the MSD test or production environments
    If openView = mc_OVVNullView And g_configuration.IsDevelopment Then

        ' Default the 'openView' attribute with a null string value
        openView = mc_OVVWriteView
    End If

    ' Parse out the 'openView' attributes value
    Select Case openView
    Case mc_OVVPrintView
        m_viewMode = mgrViewModePrint

    Case mc_OVVReadView
        m_viewMode = mgrViewModeRead

    Case mc_OVVWriteView

        ' Write View is only valid for manifest version 3 or later, if it occurs with an
        ' earlier manifest version then it must be forced to Read Only View
        If m_manifestVersion >= 3 Then
            m_viewMode = mgrViewModeWrite
        Else
            m_viewMode = mgrViewModeRead
        End If

    Case Else
        errorText = Replace$(mgrErrTextInvalidAttributeValue, mgrP1, mc_ANOpenView)
        errorText = Replace$(errorText, mgrP2, openView)
        Err.Raise mgrErrNoInvalidAttributeValue, c_proc, errorText
    End Select

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseAssessmentAttributes

Private Sub ParseAssessmentReportVersion()
    Const c_proc As String = "RootData.ParseAssessmentReportVersion"

    Dim errorText   As String
    Dim versionNode As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    Set versionNode = g_xmlDocument.SelectSingleNode(mc_NNAssessmentReportVersion)
    
    If Not versionNode Is Nothing Then
        Select Case versionNode.Text
        Case mgrReportVersionDraft
            m_reportVersion = mgrVersionTypeDraft

        Case mgrReportVersionManagers
            m_reportVersion = mgrVersionTypeManagers

        Case mgrReportVersionFinal
            m_reportVersion = mgrVersionTypeFinal

        Case Else
            errorText = Replace$(mgrErrTextInvalidNodeValue, mgrP1, mc_NNAssessmentReportVersion)
            errorText = Replace$(errorText, mgrP2, versionNode.Text)
            Err.Raise mgrErrNoInvalidNodeValue, mgrP2, errorText
        
        End Select
   End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseAssessmentReportVersion

Public Property Get ManifestVersion() As Long
    Const c_proc As String = "RootData.ManifestVersion"

    Dim errorText As String

    On Error GoTo Do_Error

    If m_initialised Then
        ManifestVersion = m_manifestVersion
    Else
        errorText = Replace$(mgrErrTextClassMustBeInitialised, mgrP1, "RootData")
        Err.Raise mgrErrNoClassMustBeInitialised, c_proc, errorText
    End If

Do_Exit:
    Exit Property

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Property ' Get ManifestVersion

Public Property Get DataChanged() As String
    DataChanged = m_dataChanged
End Property ' Get DataChanged

Public Property Get Environment() As String
    Environment = m_environment
End Property '  Get Environment

Public Property Get IsWritable() As Boolean
    IsWritable = (m_viewMode = mgrViewModeWrite)
End Property ' IsWritable

Public Property Get MainSaveURL() As String
    MainSaveURL = m_mainSaveURL
End Property ' Get MainSaveURL

Public Property Get RemedyServerName() As String
    RemedyServerName = m_remedyServerName
End Property ' Get RemedyServerName

Public Property Get RemedyServerPort() As String
    RemedyServerPort = m_remedyServerPort
End Property ' Get RemedyServerPort

Public Property Get RemedyUserId() As String
    RemedyUserId = m_remedyUserId
End Property ' Get RemedyUserId

Public Property Get ReportVersion() As mgrVersionType
     ReportVersion = m_reportVersion
End Property ' Get ReportVersion

Public Property Get TransactionGUID() As String
    TransactionGUID = m_transactionGUID
End Property '  Get TransactionGUID

Public Property Get TheRootNode() As MSXML2.IXMLDOMNode
    Set TheRootNode = m_rootNode
End Property ' Get TheRootNode

Public Property Get ViewMode() As mgrViewMode
    ViewMode = m_viewMode
End Property ' Get ViewMode
