VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        Configuration
' Purpose:      Interface module for the configuration file.
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

' NN = Node Name (Second level Node Names)
Private Const mc_NNDevelopment                As String = "development"
Private Const mc_NNDebug                      As String = "debug"
Private Const mc_NNEventLog                   As String = "eventLog"
Private Const mc_NNSchema                     As String = "schema"
Private Const mc_NNTemplates                  As String = "templates"
Private Const mc_NNWordHTMLTextFile           As String = "wordHTMLTextFile"
Private Const mc_NNWordAssessmentReportFile   As String = "wordAssessmentReportFile"
Private Const mc_NNWordFullReportFile         As String = "wordFullReportFile"
Private Const mc_NNWordSummaryReportFile      As String = "wordSummaryReportFile"
Private Const mc_NNWordXHTMLTextFile          As String = "wordXHTMLTextFile"
Private Const mc_NNRDAXMLFile                 As String = "rdaXMLFile"
Private Const mc_NNXMLTextFile                As String = "xmlTextFile"
Private Const mc_NNWordInstructions           As String = "wordInstructions"
Private Const mc_NNDocumentProtection         As String = "documentProtection"
Private Const mc_NNValidation                 As String = "validation"
Private Const mc_NNValidate                   As String = "validate"

' AND = Attribute Names for the 'development' node
Private Const mc_ANDIsDevelopment             As String = "isDevelopment"
Private Const mc_ANDPromptForFile             As String = "promptForFile"

' ANDB = Attribute Names for the 'debug' node
Private Const mc_ANDBLogging                  As String = "logging"
Private Const mc_ANDBCloseWordHTMLDocument    As String = "closeWordHTMLDocument"
Private Const mc_ANDBValidateInstructionsFile As String = "validateInstructionsFile"

' ANEL = Attribute Names for the 'eventLog' node
Private Const mc_ANELEnable                   As String = "enable"
Private Const mc_ANELName                     As String = "name"
Private Const mc_ANELPath                     As String = "path"

' ANS = Attribute Names for the 'schema' node
Private Const mc_ANSManifest1                 As String = "manifest1"
Private Const mc_ANSManifest2                 As String = "manifest2"
Private Const mc_ANSManifest3                 As String = "manifest3"
Private Const mc_ANSSSAR                      As String = "ssar"
Private Const mc_ANSPath                      As String = "path"

' ANT = Attribute Names for the 'templates' node
Private Const mc_ANTPath                      As String = "path"
Private Const mc_ANTManifest1                 As String = "manifest1"
Private Const mc_ANTManifest2                 As String = "manifest2"
Private Const mc_ANTManifest3                 As String = "manifest3"
Private Const mc_ANTSSAR                      As String = "ssar"

' ANWHTF = Attribute Names for the 'wordHTMLTextFile' node
Private Const mc_ANWHTFName                   As String = "name"
Private Const mc_ANWHTFPath                   As String = "path"
Private Const mc_ANWHTFDelete                 As String = "delete"

' ANWARF = Attribute Names for the 'wordAssessmentReportFile' node
Private Const mc_ANWARFName                   As String = "name"
Private Const mc_ANWARFPath                   As String = "path"

' ANWFRF = Attribute Names for the 'wordFullReportFile' node
Private Const mc_ANWFRFName                   As String = "name"
Private Const mc_ANWFRFPath                   As String = "path"

' ANWSRF = Attribute Names for the 'wordSummaryReportFile' node
Private Const mc_ANWSRFName                   As String = "name"
Private Const mc_ANWSRFPath                   As String = "path"


' ANWXTF = Attribute Names for the 'wordXHTMLTextFile' node
Private Const mc_ANWXTFName                   As String = "name"
Private Const mc_ANWXTFPath                   As String = "path"
Private Const mc_ANWXTFDelete                 As String = "delete"

' ANRXF = Attribute Names for the 'rdaXMLFile' node
Private Const mc_ANRXFEnable                  As String = "enable"
Private Const mc_ANRXFName                    As String = "name"
Private Const mc_ANRXFPath                    As String = "path"

' ANXTF = Attribute Names for the 'xmlTextFile' node
Private Const mc_ANXTFEnable                  As String = "enable"
Private Const mc_ANXTFName                    As String = "name"
Private Const mc_ANXTFPath                    As String = "path"

' ANWI = Attribute Names for the 'wordInstructions' node
Private Const mc_ANWIPath                     As String = "path"
Private Const mc_ANWIManifest1                As String = "manifest1"               ' Not used
Private Const mc_ANWIManifest2                As String = "manifest2"               ' The instruction file used to generate a Manifest 2 Assessment Report
Private Const mc_ANWIManifest3                As String = "manifest3"               ' The instruction file used to generate a Manifest 3 Assessment Report
Private Const mc_ANWISSAR                     As String = "ssar"                    ' The instruction file used to generate a SSAR Assessment Report

' ANDP = Attribute Names for the 'protectDocument' node
Private Const mc_ANDPEnable                   As String = "enable"                  ' True, then the document is protected for wdAllowOnlyReading
Private Const mc_ANDPGenerator                As String = "generator"               ' A series of xpath queries whose results yield the password
Private Const mc_ANDPReverse                  As String = "reverse"                 ' Reverse the generated password


Private m_generatedAssessmentReportFileName As String

Private m_passwordGenerated                 As Boolean

Private m_isDevelopment                     As Boolean
Private m_promptForFile                     As Boolean

Private m_debugLogging                      As Boolean
Private m_closeWordHTMLDocument             As Boolean
Private m_validateInstructionsFile          As Boolean                      ' Whether the Actions file should be validated

Private m_eventLogging                      As Boolean
Private m_eventLogName                      As String
Private m_eventLogPath                      As String

Private m_inputXMLFileFullPath              As String

Private m_schemaNameManifest1               As String                       ' The name of the schema file (xsd) that the Manifest 1 rep files xml is validated against
Private m_schemaNameManifest2               As String                       ' The name of the schema file (xsd) that the Manifest 2 rep files xml is validated against
Private m_schemaNameManifest3               As String                       ' The name of the schema file (xsd) that the Manifest 3 rep files xml is validated against
Private m_schemaNameSSAR                    As String                       ' The name of the schema file (xsd) that the SSAR rep files xml is validated against
Private m_schemaPath                        As String                       ' The path to the schema file (xsd) that the rep file xml is validated against

Private m_templatePath                      As String                       ' The path to the folder containing the Word templates used to produce Assessment Reports
Private m_templateManifest1                 As String                       ' The name of the Word template to produce an Assessment Report from a v1 rep file
Private m_templateManifest2                 As String                       ' The name of the Word template to produce an Assessment Report from a v2 rep file
Private m_templateManifest3                 As String                       ' The name of the Word template to produce an Assessment Report from a v3 rep file
Private m_templateSSAR                      As String                       ' The name of the Word template to produce an Assessment Report from a SSAR rep file

Private m_wordRawHTMLTextFileName           As String                       ' The file name with placeholder text
Private m_wordHTMLTextFileName              As String                       ' The file name with placeholders replaced with their appropriate values
Private m_wordHTMLTextFilePath              As String
Private m_wordHTMLTextFileDelete            As Boolean

Private m_wordAssessmentReportFileName      As String
Private m_wordAssessmentReportFilePath      As String

Private m_wordFullReportFileName            As String
Private m_wordFullReportFilePath            As String

Private m_wordSummaryReportFileName         As String
Private m_wordSummaryReportFilePath         As String


Private m_wordRawXHTMLTextFileName          As String                       ' The file name with placeholder text
Private m_wordXHTMLTextFileName             As String                       ' The file name with placeholders replaced with their appropriate values
Private m_wordXHTMLTextFilePath             As String
Private m_wordXHTMLTextFileDelete           As Boolean

Private m_rdaXMLFileEnable                  As Boolean
Private m_rdaXMLFileName                    As String
Private m_rdaXMLFilePath                    As String

Private m_xmlTextFileEnable                 As Boolean
Private m_xmlTextFileName                   As String
Private m_xmlTextFilePath                   As String

Private m_wordInstructionsPath              As String                       ' The path to the folder containing the xml Actions files used to control the production of the various Assessment Reports
Private m_wordInstructionsManifest1         As String                       ' The xml file used to control the production of an Assessment Report from a v1 rep file
Private m_wordInstructionsManifest2         As String                       ' The xml file used to control the production of an Assessment Report from a v2 rep file
Private m_wordInstructionsManifest3         As String                       ' The xml file used to control the production of an Assessment Report from a v3 rep file
Private m_wordInstructionsSSAR              As String                       ' The xml file used to control the production of an Assessment Report from a SSAR rep file

Private m_protectDocument                   As Boolean                      ' Whether the Assessment Report should have read only editing restrictions
Private m_password                          As String                       ' The generated password (the reverse option has been applied to the password)
Private m_passwordGenerator                 As String                       ' The xpath queries to use to generate the password
Private m_reversePassword                   As Boolean                      ' Whether the generated password should be reversed

Private m_validationTests                   As Collection


Public Property Get AssessmentReportFileFullName() As String
    AssessmentReportFileFullName = m_generatedAssessmentReportFileName
End Property ' Get AssessmentReportFileFullName

Public Property Get AssessmentReportPassword() As String

    ' Only do this once
    If Not m_passwordGenerated Then

        ' Generate the Assessment Report password
        GeneratePassword

        ' Set the flag so we only generate the password once
        m_passwordGenerated = True
    End If

    ' Return the previously generated password
    AssessmentReportPassword = m_password
End Property ' Get AssessmentReportPassword


'=======================================================================================================================
' Procedure:    InputXMLFileFullPath
' Purpose:      These let or get the full path of the input xml file (normally fetchdoc.infopathxml).
' Notes:        Set manually withing the code not loaded from the configuration file.
'
' On Entry:     xmlFileFullPath     The full path and name of the xml input file that generates the Assessment Report.
' Returns:      The full path and name of the xml input file that generates the Assessment Report.
'=======================================================================================================================
Public Property Get InputXMLFileFullPath() As String
    InputXMLFileFullPath = m_inputXMLFileFullPath
End Property ' Get InputXMLFileFullPath
Friend Property Let InputXMLFileFullPath(ByVal xmlFileFullPath As String)
    m_inputXMLFileFullPath = xmlFileFullPath
End Property ' Let InputXMLFileFullPath

Public Property Get InstructionsFilePath() As String
    InstructionsFilePath = m_wordInstructionsPath
End Property ' Get InstructionsFilePath

Public Property Get InstructionsFileFullNameManifest1() As String
    InstructionsFileFullNameManifest1 = QualifyPath(m_wordInstructionsPath) & m_wordInstructionsManifest1
End Property ' Get InstructionsFileFullNameManifest1

Public Property Get InstructionsFileFullNameManifest2() As String
    InstructionsFileFullNameManifest2 = QualifyPath(m_wordInstructionsPath) & m_wordInstructionsManifest2
End Property ' Get InstructionsFileFullNameManifest2

Public Property Get InstructionsFileFullNameManifest3() As String
    InstructionsFileFullNameManifest3 = QualifyPath(m_wordInstructionsPath) & m_wordInstructionsManifest3
End Property ' Get InstructionsFileFullNameManifest3

Public Property Get InstructionsFileFullNameSSAR() As String
    InstructionsFileFullNameSSAR = QualifyPath(m_wordInstructionsPath) & m_wordInstructionsSSAR
End Property ' Get InstructionsFileFullNameSSAR

Public Property Get CurrentInstructionsFileFullName() As String
    Const c_proc As String = "Configuration.Get CurrentInstructionsFileFullName"

    Dim errorText As String

    On Error GoTo Do_Error

    Select Case g_rootData.ManifestVersion
    Case 1
        CurrentInstructionsFileFullName = InstructionsFileFullNameManifest1

    Case 2
        CurrentInstructionsFileFullName = InstructionsFileFullNameManifest2

    Case 3
        CurrentInstructionsFileFullName = InstructionsFileFullNameManifest3

    Case 4
        CurrentInstructionsFileFullName = InstructionsFileFullNameSSAR

    Case Else
        errorText = Replace$(mgrErrTextInvalidManifestVersionNumber, mgrP1, CStr(g_rootData.ManifestVersion))
        Err.Raise mgrErrNoInvalidManifestVersionNumber, c_proc, errorText
    End Select

Do_Exit:
    Exit Property

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Property ' CurrentInstructionsFileFullName()

Public Property Get CurrentManifestVersion() As Long
    CurrentManifestVersion = g_rootData.ManifestVersion
End Property ' CurrentManifestVersion

Public Property Get CurrentTemplateFullName() As String
    Const c_proc As String = "Configuration.Get CurrentTemplateFullName"

    Dim errorText As String

    On Error GoTo Do_Error

    Select Case g_rootData.ManifestVersion
    Case 1
        CurrentTemplateFullName = TemplateFullNameManifest1

    Case 2
        CurrentTemplateFullName = TemplateFullNameManifest2

    Case 3
        CurrentTemplateFullName = TemplateFullNameManifest3

    Case 4
        CurrentTemplateFullName = TemplateFullNameSSAR

    Case Else
        errorText = Replace$(mgrErrTextInvalidManifestVersionNumber, mgrP1, CStr(g_rootData.ManifestVersion))
        Err.Raise mgrErrNoInvalidManifestVersionNumber, c_proc, errorText
    End Select

Do_Exit:
    Exit Property

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Property ' Get CurrentTemplateFullName

Public Property Get DebugLogging() As Boolean
    DebugLogging = m_debugLogging
End Property ' Get DebugLogging

Public Property Get CloseWordHTMLDocument() As Boolean
    CloseWordHTMLDocument = m_closeWordHTMLDocument
End Property ' Get CloseWordHTMLDocument

Public Property Get EventLogFile() As String
    EventLogFile = m_eventLogName
End Property ' EventLogFile

Public Property Get EventLogging() As Boolean
    EventLogging = m_eventLogging
End Property ' Get EventLogging

Public Property Get EventLogPath() As String
    EventLogPath = m_eventLogPath
End Property ' EventLogPath

Public Property Get IsDevelopment() As Boolean
    IsDevelopment = m_isDevelopment
End Property ' Get IsDevelopment

Public Property Get PromptForFile() As Boolean
    PromptForFile = m_promptForFile
End Property ' Get PromptForFile

Public Property Get ProtectDocument() As Boolean
    ProtectDocument = m_protectDocument
End Property ' ProtectDocument

Public Property Get RDAXMLFileFullName() As String
     RDAXMLFileFullName = QualifyPath(m_rdaXMLFilePath) & m_rdaXMLFileName
End Property ' Get RDAXMLFileFullName

Public Property Get SaveRDAXMLFile() As Boolean
     SaveRDAXMLFile = m_rdaXMLFileEnable
End Property ' Get SaveRDAXMLFile

Public Property Get SaveXMLTextFile() As Boolean
     SaveXMLTextFile = m_xmlTextFileEnable
End Property ' Get SaveXMLTextFile

Public Property Get ShouldValidateInstructionsFile() As Boolean
    ShouldValidateInstructionsFile = m_validateInstructionsFile
End Property ' Get ShouldValidateInstructionsFile

Public Property Get SchemaFullName() As String
    Const c_proc As String = "Configuration.SchemaFullName"

    Dim errorText As String

    On Error GoTo Do_Error

    ' Return the appropriate schema path and name based on the version of the manifest used by the loaded rep file
    Select Case g_rootData.ManifestVersion
    Case 1
        SchemaFullName = QualifyPath(m_schemaPath) & m_schemaNameManifest1

    Case 2
        SchemaFullName = QualifyPath(m_schemaPath) & m_schemaNameManifest2

    Case 3
        SchemaFullName = QualifyPath(m_schemaPath) & m_schemaNameManifest3
        
    Case 4
        SchemaFullName = QualifyPath(m_schemaPath) & m_schemaNameSSAR

    Case Else
        errorText = Replace$(mgrErrTextInvalidManifestVersionNumber, mgrP1, CStr(g_rootData.ManifestVersion))
        Err.Raise mgrErrNoInvalidManifestVersionNumber, c_proc, errorText
    End Select

Do_Exit:
    Exit Property

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Property ' Get SchemaFullName

Public Property Get SchemaPath() As String
     SchemaPath = QualifyPath(m_schemaPath)
End Property ' Get SchemaPath

Public Property Get TemplatePath() As String
    TemplatePath = m_templatePath
End Property ' Get TemplatePath

Public Property Get TemplateFullNameManifest1() As String
    TemplateFullNameManifest1 = QualifyPath(m_templatePath) & m_templateManifest1
End Property ' Get TemplateFullNameManifest1

Public Property Get TemplateFullNameManifest2() As String
    TemplateFullNameManifest2 = QualifyPath(m_templatePath) & m_templateManifest2
End Property ' Get TemplateFullNameManifest2

Public Property Get TemplateFullNameManifest3() As String
    TemplateFullNameManifest3 = QualifyPath(m_templatePath) & m_templateManifest3
End Property ' Get TemplateFullNameManifest3

Public Property Get TemplateFullNameSSAR() As String
    TemplateFullNameSSAR = QualifyPath(m_templatePath) & m_templateSSAR
End Property ' Get TemplateFullNameSSAR

Public Property Get ValidateInstructionsFile() As Boolean
    ValidateInstructionsFile = m_validateInstructionsFile
End Property ' Get ValidateInstructionsFile

Public Property Get WordHTMLTextFileDelete() As Boolean
     WordHTMLTextFileDelete = m_wordHTMLTextFileDelete
End Property ' Get WordHTMLTextFileDelete

Public Property Get WordHTMLTextFilePath() As String
    WordHTMLTextFilePath = m_wordHTMLTextFilePath
End Property ' Get WordHTMLTextFilePath

Public Property Get WordHTMLTextFileFullName() As String

    ' If the static file name is a null string then we need to generate it
    If LenB(m_wordHTMLTextFileName) = 0 Then

        ' Create the temporary file name by using the supplied text plus a date
        m_wordHTMLTextFileName = MakeUniqueTemporaryFileName(m_wordRawHTMLTextFileName)
    End If

    ' Return the full path of the temporary html file
    WordHTMLTextFileFullName = QualifyPath(m_wordHTMLTextFilePath) & m_wordHTMLTextFileName
End Property ' Get WordHTMLTextFileFullName

Public Property Get FullReportFileFullName() As String
    Dim baseFullName    As String

    baseFullName = QualifyPath(m_wordFullReportFilePath) & m_wordFullReportFileName

    FullReportFileFullName = Replace$(baseFullName, mgrP1, Format$(Now, mgrTemporaryFileDateFormat))
End Property ' Get FullReportFileFullName

Public Property Get SummaryReportFileFullName() As String
    Dim baseFullName    As String

    baseFullName = QualifyPath(m_wordSummaryReportFilePath) & m_wordSummaryReportFileName

    SummaryReportFileFullName = Replace$(baseFullName, mgrP1, Format$(Now, mgrTemporaryFileDateFormat))
End Property ' Get SummaryReportFileFullName

Public Property Get WordXHTMLTextFileDelete() As Boolean
    WordXHTMLTextFileDelete = m_wordXHTMLTextFileDelete
End Property ' Get WordXHTMLTextFileDelete

Public Property Get WordXHTMLTextFileFullName() As String

    ' If the static file name is a null string then we need to generate it
    If LenB(m_wordXHTMLTextFileName) = 0 Then

        ' Create the temporary file name by using the supplied text plus a date
        m_wordXHTMLTextFileName = MakeUniqueTemporaryFileName(m_wordRawXHTMLTextFileName)
    End If

    ' Return the full path of the temporary html file
    WordXHTMLTextFileFullName = QualifyPath(m_wordXHTMLTextFilePath) & m_wordXHTMLTextFileName
End Property ' Get WordXHTMLTextFileFullName

Public Property Get XMLTextFileFullName() As String
     XMLTextFileFullName = QualifyPath(m_xmlTextFilePath) & m_xmlTextFileName
End Property ' Get XMLTextFileFullName

Friend Sub GenerateAssessmentReportFileName()
    Dim baseFullName    As String

    baseFullName = QualifyPath(m_wordAssessmentReportFilePath) & m_wordAssessmentReportFileName

    m_generatedAssessmentReportFileName = Replace$(baseFullName, mgrP1, Format$(Now, mgrTemporaryFileDateFormat))
End Sub ' GenerateAssessmentReportFileName

Friend Sub ResetGeneratedFileNames()
    m_wordHTMLTextFileName = vbNullString
    m_wordXHTMLTextFileName = vbNullString
End Sub ' ResetGeneratedFileNames

'=======================================================================================================================
' Procedure:    GeneratePassword
' Purpose:      Generates a password used to lock the Assessment Report to prevent the user from changing Ranges other
'               than those enabled for editing.
'
' On Entry:     m_passwordGenerator The query strings used to generate the password.
' On Exit:      m_password          Contains the password that is used to lock the Assessment Report.
'=======================================================================================================================
Private Sub GeneratePassword()
    Const c_proc As String = "Configuration.GeneratePassword"

    Dim backwards As String
    Dim index     As Long
    Dim Password  As String
    Dim queries() As String
    Dim theQuery  As Variant

    On Error GoTo Do_Error

    ' Make sure the rep file has been loaded
    If Not g_xmlDocument Is Nothing Then

        ' Only generate a password if the Assessment Report is going to be restricted to Read Only Editing
        If m_protectDocument Then

            ' The generator string may contain multiple comma delimited queries
            queries() = Split(m_passwordGenerator, ",")

            ' Run each query and concatenate the query results to build the password
            For Each theQuery In queries

                Password = Password & g_xmlDocument.SelectSingleNode(theQuery).Text
            Next

            ' See if the password should be reversed
            If m_reversePassword Then

                ' Reverse the password
                For index = Len(Password) To 1 Step -1
                    backwards = backwards & Mid$(Password, index, 1)
                Next

                ' Store the reversed password
                m_password = backwards
            Else

                ' Store the non reversed password
                m_password = Password
            End If
        End If
    Else
        Err.Raise mgrErrNoInfoPathXMLFileHasNotBeenLoaded, c_proc, mgrErrTextInfoPathXMLFileHasNotBeenLoaded
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' GeneratePassword

'=======================================================================================================================
' Procedure:    Initialise
' Purpose:      Procedure to set up the Configuration object model from the 'word rda config.xml' file.
' Note:         Triggers all setup activities.
'=======================================================================================================================
Friend Sub Initialise()
    Const c_proc As String = "Configuration.Initialise"

    Dim configFile         As MSXML2.DOMDocument60
    Dim configFileFullName As String
    Dim errorText          As String
    Dim loadFolder         As String

    On Error GoTo Do_Error

    ' Set up the collection object used to hold ValidationTest objects
    Set m_validationTests = New Collection

    ' Load the configuration xml file, which must be in the same folder as this template (addin)
    loadFolder = QualifyPath(ThisDocument.path)

    ' Create the full name for the configuration file
    configFileFullName = QualifyPath(loadFolder) & mgrFIConfigurationFile

    ' Load the xml configuration file
    If LoadXmlFile(configFileFullName, configFile) Then

        ' Now parse the configuration file
        ParseConfig configFile
    Else
        errorText = Replace$(mgrErrTextFailedToLoadConfigurationFile, mgrP1, configFileFullName)
        errorText = Replace$(errorText, mgrP2, configFile.parseError.reason)
        Err.Raise mgrErrNoFailedToLoadConfigurationFile, c_proc, errorText
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

'=======================================================================================================================
' Procedure:    ParseConfig
' Purpose:      Parses all 'configuration' sub-node data.
' Notes:        The code is structured so that each sub-node has its own parser.
'
' On Entry:     configFile          The xml file containing the word rda configuration data.
'=======================================================================================================================
Private Sub ParseConfig(ByRef configFile As MSXML2.DOMDocument60)
    Const c_proc As String = "Configuration.ParseConfig"

    Dim child     As MSXML2.IXMLDOMNode
    Dim errorText As String
    Dim rootNode  As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Locate the root node
    For Each child In configFile.ChildNodes

        ' The first node of type Element is the root node
        If child.NodeType = NODE_ELEMENT Then
            Set rootNode = child
            Exit For
        End If
    Next

    ' Check that we found the root node
    If rootNode Is Nothing Then
        Err.Raise mgrErrNoUnexpectedCondition, c_proc, mgrErrTextUnexpectedCondition
    End If

    ' Iterate the root nodes children
    For Each child In rootNode.ChildNodes

        ' Filter out any comment nodes
        If child.NodeType <> NODE_COMMENT Then

            ' Parse each child nodes attributes.
            ' The Parse Attribute procedure takes pairs of values - the Attribute Name and Variable, to contain the Attribute value
            Select Case child.BaseName
            Case mc_NNDevelopment
                ParseAttributes child, mc_ANDIsDevelopment, m_isDevelopment, mc_ANDPromptForFile, m_promptForFile

            Case mc_NNDebug
                ParseAttributes child, mc_ANDBLogging, m_debugLogging, mc_ANDBCloseWordHTMLDocument, m_closeWordHTMLDocument, _
                                       mc_ANDBValidateInstructionsFile, m_validateInstructionsFile

            Case mc_NNEventLog
                ParseAttributes child, mc_ANELEnable, m_eventLogging, mc_ANELName, m_eventLogName, mc_ANELPath, m_eventLogPath

            Case mc_NNSchema
                ParseAttributes child, mc_ANSPath, m_schemaPath, mc_ANSManifest1, m_schemaNameManifest1, _
                                       mc_ANSManifest2, m_schemaNameManifest2, mc_ANSManifest3, m_schemaNameManifest3, _
                                       mc_ANSSSAR, m_schemaNameSSAR

            Case mc_NNTemplates
                ParseAttributes child, mc_ANTPath, m_templatePath, mc_ANTManifest1, m_templateManifest1, _
                                       mc_ANTManifest2, m_templateManifest2, mc_ANTManifest3, m_templateManifest3, _
                                       mc_ANTSSAR, m_templateSSAR

            Case mc_NNWordHTMLTextFile
                ParseAttributes child, mc_ANWHTFPath, m_wordHTMLTextFilePath, mc_ANWHTFName, m_wordRawHTMLTextFileName, _
                                       mc_ANWHTFDelete, m_wordHTMLTextFileDelete

            Case mc_NNWordAssessmentReportFile
                ParseAttributes child, mc_ANWARFPath, m_wordAssessmentReportFilePath, mc_ANWARFName, m_wordAssessmentReportFileName

            Case mc_NNWordFullReportFile
                ParseAttributes child, mc_ANWFRFPath, m_wordFullReportFilePath, mc_ANWFRFName, m_wordFullReportFileName

            Case mc_NNWordSummaryReportFile
                ParseAttributes child, mc_ANWSRFPath, m_wordSummaryReportFilePath, mc_ANWSRFName, m_wordSummaryReportFileName

            Case mc_NNWordXHTMLTextFile
                ParseAttributes child, mc_ANWXTFPath, m_wordXHTMLTextFilePath, mc_ANWXTFName, m_wordRawXHTMLTextFileName, _
                                       mc_ANWXTFDelete, m_wordXHTMLTextFileDelete

            Case mc_NNRDAXMLFile
                ParseAttributes child, mc_ANRXFEnable, m_rdaXMLFileEnable, mc_ANRXFPath, m_rdaXMLFilePath, _
                                       mc_ANRXFName, m_rdaXMLFileName

            Case mc_NNXMLTextFile
                ParseAttributes child, mc_ANXTFEnable, m_xmlTextFileEnable, mc_ANXTFPath, m_xmlTextFilePath, _
                                       mc_ANXTFName, m_xmlTextFileName

            Case mc_NNWordInstructions
                ParseAttributes child, mc_ANWIPath, m_wordInstructionsPath, mc_ANWIManifest1, m_wordInstructionsManifest1, _
                                       mc_ANWIManifest2, m_wordInstructionsManifest2, mc_ANWIManifest3, m_wordInstructionsManifest3, _
                                       mc_ANWISSAR, m_wordInstructionsSSAR

            Case mc_NNDocumentProtection
                ParseAttributes child, mc_ANDPGenerator, m_passwordGenerator, mc_ANDPReverse, m_reversePassword, _
                                       mc_ANDPEnable, m_protectDocument

            Case mc_NNValidation
                ParseValidationNode child

            Case Else
                errorText = Replace$(mgrErrTextInvalidNodeName, mgrP1, child.BaseName)
                Err.Raise mgrErrNoInvalidNodeName, c_proc, errorText
            End Select
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseConfig

'=======================================================================================================================
' Procedure:    ParseValidationNode
' Purpose:      Parses the 'validation' nodes attributes.
' Notes:        This node contains multiple 'validate' child nodes, each child node is a complete test.
'
' On Entry:     validationNode      The 'validation' node whose child nodes are to be parsed.
'=======================================================================================================================
Private Sub ParseValidationNode(ByVal validationNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "Configuration.mc_NNValidate"

    Dim child     As MSXML2.IXMLDOMNode
    Dim errorText As String

    On Error GoTo Do_Error

    ' Create the Validation Tests Collection object
    Set m_validationTests = New Collection

    ' Parse out the 'validation' nodes child nodes
    For Each child In validationNode.ChildNodes

        ' Parse each child node
        Select Case child.BaseName
        Case mc_NNValidate
            ParseValidateNode child

        Case Else
            errorText = Replace$(mgrErrTextInvalidNodeName, mgrP1, child.BaseName)
            Err.Raise mgrErrNoInvalidNodeName, c_proc, errorText

        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' mc_NNValidate

'=======================================================================================================================
' Procedure:    ParseValidateNode
' Purpose:      Parses the 'validate' node.
' Notes:        These attributes contain information about a validation test that will be run against the
'               'rda instructions.xml' file. Each 'validate' node is one complete test.
'
' On Entry:     validateNode        The 'validate' node whose attributes are to be parsed.
' On Exit:      m_validationTests   A new ValidationTest object has been added to this collection object.
'=======================================================================================================================
Private Sub ParseValidateNode(ByVal validateNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "Configuration.ParseValidateNode"

    Dim newTest      As ValidationTest

    On Error GoTo Do_Error

    ' Create a new ValidationTest object
    Set newTest = New ValidationTest

    ' Have the ValidationTest object parse the xml
    newTest.Initialise validateNode

    ' Ad the new ValidationTest object to the test objects collection
    m_validationTests.Add newTest

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseValidateNode

'=======================================================================================================================
' Procedure:    PerformActionsFileValidation
' Purpose:      Performs each of the validation tests defined for the current manifest.
'
' On Entry:     m_validationTests   Contains all tests to be performed.
' Returns:      True if there were no tests run or all defined tests returned the expected results.
'=======================================================================================================================
Friend Function PerformActionsFileValidation() As Boolean
    Const c_proc As String = "Configuration.PerformActionsFileValidation"

    Dim allOk    As Boolean
    Dim theNodes As MSXML2.IXMLDOMNodeList
    Dim theTest  As ValidationTest

    On Error GoTo Do_Error

    ' Set initial test result value or the final test results will never be correct
    allOk = True

    ' Iterate all validation test, runing those tests that match the current manifest version
    For Each theTest In m_validationTests

        ' Check that the test is for the current manifest version number
        If theTest.ManifestVersion = g_configuration.CurrentManifestVersion Then

            ' Execute the xpath query against the Actions xml document
            Set theNodes = g_xmlInstructionData.SelectNodes(theTest.Query)

            ' Determine whether the query results are valid
            If theTest.ExpectedResult = mgrValidationResultNone And theNodes.Length = 0 Then
                
                ' Valid result - all ok only remain True if all test results are True
                allOk = allOk And True
            ElseIf theTest.ExpectedResult = mgrValidationResultSome And theNodes.Length > 0 Then

                ' Valid result - all ok only remain True if all test results are True
                allOk = allOk And True
            Else

                ' Failed validation - so report it, log it - DO SOMETHING WITH IT!!
                allOk = False
            End If
        End If
    Next

    ' Return test result
    PerformActionsFileValidation = allOk

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' PerformActionsFileValidation

Friend Sub ResetPassword()
    m_passwordGenerated = False
End Sub ' ResetPassword
