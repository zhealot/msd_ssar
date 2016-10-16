Attribute VB_Name = "modGlobalData"
'===================================================================================================================================
' Module:       modGlobalData
' Purpose:      Contains data used by this addin, and the SSAR and RDA addins.
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
' History:      30/05/16    1.  Created.
'===================================================================================================================================
Option Explicit

Public Const mgrFIRDAXMLDataFileType                        As String = "rep"

Public Const mgrFIConfigurationFile                         As String = "word rda config.xml"

' Used for parameter replacement
Public Const mgrP1                                          As String = "%1"
Public Const mgrP2                                          As String = "%2"
Public Const mgrP3                                          As String = "%3"
Public Const mgrP4                                          As String = "%4"

' Predefined error numbers
Public Const mgrErrNoInvalidProcedureCall                   As Long = 5
Public Const mgrErrNoTypeMismatch                           As Long = 13
Public Const mgrErrNoKeyAlreadyAssociated                   As Long = 457
Public Const mgrErrNoRequestedCollectionMemberDoesNotExist  As Long = 5941
Public Const mgrErrNoUnableToRunTheSpecifiedMacro           As Long = -2147352573

' Custom error numbers (these can be changed with refering to the RDA and SSAR addins)
Public Const mgrErrNoBase                                   As Long = -20000
Public Const mgrErrNoUnexpectedCondition                    As Long = mgrErrNoBase - 10
Public Const mgrErrNoFailedToLoadConfigurationFile          As Long = mgrErrNoBase - 20
Public Const mgrErrNoRepDocumentLoad                        As Long = mgrErrNoBase - 30
Public Const mgrErrNoXMLSchemaLoad                          As Long = mgrErrNoBase - 40
Public Const mgrErrNoXMLSchemaValidate                      As Long = mgrErrNoBase - 50
Public Const mgrErrNoUnexpectedHTMLDocumentStructure        As Long = mgrErrNoBase - 60
Public Const mgrErrNoInvalidIfBlock                         As Long = mgrErrNoBase - 70
Public Const mgrErrNoInvalidNodeName                        As Long = mgrErrNoBase - 80
Public Const mgrErrNoInvalidNodeValue                       As Long = mgrErrNoBase - 90
Public Const mgrErrNoDuplicateNodeName                      As Long = mgrErrNoBase - 95
Public Const mgrErrNoMissingAttribute                       As Long = mgrErrNoBase - 100
Public Const mgrErrNoOnlyOneAttributeAllowed                As Long = mgrErrNoBase - 101
Public Const mgrErrNoDuplicateAttributeName                 As Long = mgrErrNoBase - 110
Public Const mgrErrNoInvalidAttributeCombination            As Long = mgrErrNoBase - 111
Public Const mgrErrNoInvalidAttributeName                   As Long = mgrErrNoBase - 120
Public Const mgrErrNoInvalidAttributeValue                  As Long = mgrErrNoBase - 130
Public Const mgrErrNoUnknownFunctionNameError               As Long = mgrErrNoBase - 135
Public Const mgrErrNoInvalidParamArray                      As Long = mgrErrNoBase - 140
Public Const mgrErrNoKeyDoesNotExistInDictionary            As Long = mgrErrNoBase - 141
Public Const mgrErrNoInsertAfterRowNumber                   As Long = mgrErrNoBase - 142
Public Const mgrErrNoCouldNotGetManifestVersion             As Long = mgrErrNoBase - 150
Public Const mgrErrNoProcessingInstructionNotFound          As Long = mgrErrNoBase - 160
Public Const mgrErrNoClassMustBeInitialised                 As Long = mgrErrNoBase - 170
Public Const mgrErrNoInvalidActionVerb                      As Long = mgrErrNoBase - 180
Public Const mgrErrNoUnknownActionVerbType                  As Long = mgrErrNoBase - 190
Public Const mgrErrNoUnknownAddContentActionVerbType        As Long = mgrErrNoBase - 200
Public Const mgrErrNoCounterDepthExceedsMaximum             As Long = mgrErrNoBase - 210
Public Const mgrErrNoInvalidManifestVersionNumber           As Long = mgrErrNoBase - 220
Public Const mgrErrNoInvalidDataFormatNodeValue             As Long = mgrErrNoBase - 230
Public Const mgrErrNoNoSpecifiedBookmarkExists              As Long = mgrErrNoBase - 240
Public Const mgrErrNoNoSuchBookmark                         As Long = mgrErrNoBase - 250
Public Const mgrErrNoBookmarkIsExpectedToContainATable      As Long = mgrErrNoBase - 251
Public Const mgrErrNoAddNodeMissingAttribute                As Long = mgrErrNoBase - 260
Public Const mgrErrNoAddNodeMissingAttributeN               As Long = mgrErrNoBase - 270
Public Const mgrErrNoInfoPathXMLFileHasNotBeenLoaded        As Long = mgrErrNoBase - 290
Public Const mgrErrNoEditorsDictionaryUndefined             As Long = mgrErrNoBase - 300
Public Const mgrErrNoContentControlMissingFromBookmark      As Long = mgrErrNoBase - 310
Public Const mgrErrNoInvalidDocumentContent                 As Long = mgrErrNoBase - 320
Public Const mgrErrNoFailedToSubmitXMLToWebservice          As Long = mgrErrNoBase - 330

' Overrides for predefined error messages
Public Const mgrErrTextUnableToRunTheSpecifiedMacro         As String = "Unable to run the specified macro: %1"

' Used to override multiple generic error messages
Public Const mgrErrTextBookmarkDoesNotExist                 As String = "The specified bookmark: %1, does not exist"
Public Const mgrErrTextBuildingBlockDoesNotExist            As String = "The specified building block: %1, does not exist"

' Custom error descriptions (these cannot be changed without first refering to the RDA and SSAR addins)
Public Const mgrErrTextUnexpectedCondition                  As String = "Unexpected condition."

Public Const mgrErrTextFailedToLoadConfigurationFile        As String = "Failed to load configuration file: %1" & vbCr & _
                                                                       "Reason: %2"
Public Const mgrErrTextRepDocumentLoad                      As String = "Rep Input file: %1 failed to load."
Public Const mgrErrTextXMLSchemaLoad                        As String = "XML Schema file: %1 failed to load."
Public Const mgrErrTextXMLSchemaValidate                    As String = "XML Input file: %1 failed to validate using schema: %2."

Public Const mgrErrTextUnexpectedHTMLDocumentStructure      As String = "Unexpected HTML document structure encountered in input file: %1"

Public Const mgrErrTextInvalidIfBlock                       As String = "Invalid 'if' node: %1"
Public Const mgrErrTextInvalidNodeName                      As String = "Invalid node name: %1"
Public Const mgrErrTextInvalidNodeValue                     As String = "Invalid node (%1) value: %2"
Public Const mgrErrTextDuplicateNodeName                    As String = "Duplicate node name: %1. Parent node: %2"
Public Const mgrErrTextMissingAttribute                     As String = "Node: %1, missing attribute: (%2)"
Public Const mgrErrTextOnlyOneAttributeAllowed              As String = "Node: %1, only one attribute is allowed: (%2)"
Public Const mgrErrTextDuplicateAttributeName               As String = "Node: %1, duplicate attribute: %2"
Public Const mgrErrTextInvalidAttributeCombination          As String = "Node: %1, invalid attribute combination: %2 value: %3"
Public Const mgrErrTextInvalidAttributeName                 As String = "Node: %1, invalid attribute name: %2"
Public Const mgrErrTextInvalidAttributeValue                As String = "Node: %1, invalid attribute: %2 value: %3"
Public Const mgrErrTextUnknownFunctionNameError             As String = "Node: %1, invalid function name specified in attribute: %2 value: %3"
Public Const mgrErrTextInvalidParamArray                    As String = "ParamArray contains an incorrect number of elements."
Public Const mgrErrTextKeyDoesNotExistInDictionary          As String = "Key: %1, does not exist in dictionary: %2"
Public Const mgrErrTextInsertAfterRowNumber                 As String = "Invalid insertAfterRow row number specified: %1 (maximum: %2)"
Public Const mgrErrTextCouldNotGetManifestVersion           As String = "The Manifest version number could not be obtained from the" & vbCr & _
                                                                        "'mso-infoPathSolution' Processing Instructions 'href' attribute:" & vbCr & "%1"
Public Const mgrErrTextProcessingInstructionNotFound        As String = "Processing Instruction 'mso-infoPathSolution' was not found"
Public Const mgrErrTextClassMustBeInitialised               As String = "The class '%1' instance must be Initialised before accessing its Properties or Methods"

Public Const mgrErrTextInvalidActionVerb                    As String = "Invalid Action verb: %1"
Public Const mgrErrTextUnknownActionVerbType                As String = "Unknown Action verb type: %1"
Public Const mgrErrTextUnknownAddContentActionVerbType      As String = "Unknown AddContent Action verb type: %1"
Public Const mgrErrTextCounterDepthExceedsMaximum           As String = "The maximum counter depth of %1 has been exceeded." & vbCr & _
                                                                        "As a result of nested Add blocks in the Actions file:" & vbCr & _
                                                                        "%2"
Public Const mgrErrTextInvalidManifestVersionNumber         As String = "Invalid manifest version number: %1"
Public Const mgrErrTextInvalidDataFormatNodeValue           As String = "Invalid 'dataFormat' node value: %1"
Public Const mgrErrTextNoSpecifiedBookmarkExists            As String = "None of the specified bookmarks exist: "
Public Const mgrErrTextNoSuchBookmark                       As String = "Bookmark ""%1"" missing from the Assessment Report"
Public Const mgrErrTextBookmarkIsExpectedToContainATable    As String = "Bookmark: ""%1"" is expected to contain a Table"
Public Const mgrErrTextAddNodeMissingAttribute              As String = "Action: %1, missing 'buildingBlock', 'bookmark' and/or 'where' attribute." & vbCr & _
                                                                        "Actual value: %2"
Public Const mgrErrTextAddNodeMissingAttributeN             As String = "Action: %1, missing 'buildingBlockN', 'bookmarkN' and/or 'whereN' attribute." & vbCr & _
                                                                        "Actual value: %2"
Public Const mgrErrTextInvalidAttributeValueExtended        As String = "Action: %1, expected '%2' attribute value of %3." & vbCr & _
                                                                        "Actual value: %4"
Public Const mgrErrTextInfoPathXMLFileHasNotBeenLoaded      As String = "The InfoPath xml file (fetchdoc.infopathxml) has not been loaded"
Public Const mgrErrTextEditorsDictionaryUndefined           As String = "The Editors Dictionary object is undefined"
Public Const mgrErrTextContentControlMissingFromBookmark    As String = "Bookmark: %1, is expected to contain a Content Control"
Public Const mgrErrTextInvalidDocumentContent               As String = "You have entered invalid content (%1) into the Assessment Report." & vbCr & _
                                                                        "This must be removed before you can Submit it."
Public Const mgrErrTextFailedToSubmitXMLToWebservice        As String = "An error has occurred on submitting report data to mgr" & vbCr & _
                                                                        "Webservice: %1" & vbCr & _
                                                                        "WebClient error (%2): %3" & vbCr & _
                                                                        "Please contact IT HELP CYF for assistance"

' These are for use with mgrErrNoKeyAlreadyAssociated
Public Const mgrErrTextDuplicateBookmarkPattern             As String = "Bookmark pattern %1 already in collection"

' Warning messages
Public Const mgrWarnUnexpectedHTMLNodeType                  As String = "Unexpected HTML node type encountered: %1"
Public Const mgrWarnBeforeDenyClose                         As String = "You must Submit or Cancel the Assessment Report from the RDA Tab before trying to close the document."
Public Const mgrWarnOnlyOneActiveAssessmentReport           As String = "You can only have one Active Assessment Report open at any time."

' Xpath queries run against the "rep" file
Public Const mgrXQFDDescendantDivNodes                      As String = "descendant::*[local-name()='div']"
Public Const mgrXQFDDescendantPNodes                        As String = "descendant::*[local-name()='p']"
Public Const mgrXQFDHrefProcessingInstruction               As String = "//processing-instruction('mso-infoPathSolution')"

' Miscellaneous assessment report xml node names
Public Const mgrNNAssessment                                As String = "Assessment"

' HTMLEN = HTML Entity Name
Public Const mgr_HTMLEN_NBSP                                As String = "&nbsp;"                 ' Non breaking space
Public Const mgr_HTMLEN_LT                                  As String = "&lt;"                   ' "<" character
Public Const mgr_HTMLEN_GT                                  As String = "&gt;"                   ' ">" character

' HTMLOT = HTML Original Text
Public Const mgr_HTMLOT_ClassMsoNormal1                     As String = " class=MsoNormal>"      ' This has to be done in two parts     as "MsoNormal" is a prefix to other strings
Public Const mgr_HTMLOT_ClassMsoNormal2                     As String = " class=MsoNormal "

' XMLECV = XML Escaped Character Value
Public Const mgr_XMLECV_NBSP                                As String = "&#160;"                 ' Non breaking space (decimal 160)
Public Const mgr_XMLECV_LT                                  As String = "&#60;"                  ' Left angle bracket "<" (decimal 60)
Public Const mgr_XMLECV_GT                                  As String = "&#62;"                  ' Right angle bracket ">" (decimal 62)

' HTMLRT = HTML Replacement Text
Public Const mgr_HTMLRT_ClassMsoNormal1                     As String = ">"
Public Const mgr_HTMLRT_ClassMsoNormal2                     As String = " "

' Used to locate the start and end of data blocks and bookmark names added to html text so we know which data element the html is for
Public Const mgrHTMLBookmarkedBlockLeadIn                   As String = "¦¦¦@@@"     ' Characters used to indicate the start of a bookmarked block
Public Const mgrHTMLBookmarkNameEnd                         As String = "`~`"        ' Characters used to indicate the end of the bookmark name
Public Const mgrHTMLBookmarkedBlockLeadOut                  As String = "@@@|¦|"     ' Characters used to indicate the end of a bookmarked block

' Hack to allow us to recognise where a bullet should be
Public Const mgrBulletRecognitionSequence                   As String = "§§§"

' Date formats used when displaying dates
Public Const mgrDateFormatLong                              As String = "dd mmmm yyyy"
Public Const mgrDateFormatShort                             As String = "dd/mm/yyyy"
Public Const mgrDateFormatShortYear                         As String = "dd/mm/yy"

' Used to detect the report version within the xml
Public Const mgrReportVersionDraft                          As String = "Draft Report"
Public Const mgrReportVersionManagers                       As String = "Managers Report"
Public Const mgrReportVersionFinal                          As String = "Final Report"

' Building Block names for Watermarks
Public Const mgrBBWatermarkDraft                            As String = "Draft Report"
Public Const mgrBBWatermarkManagers                         As String = "Managers Report"
Public Const mgrBBWatermarkFinal                            As String = "Final Report"

' Document Variable names
Public Const mgrDVAssessmentReport                          As String = "AssessmentReport"          ' Used to determine whether a document is an Assessment Report.
                                                                                                    ' This should have a boolean value.
Public Const mgrDVDataFile                                  As String = "data file"                 ' The data (rep) file used to create the assessment report
Public Const mgrDVInstructionFile                           As String = "instruction file"          ' The instruction file used to create the assessment report
Public Const mgrDVManagerAddinVersion                       As String = "manager addin version"
Public Const mgrDVPreserveUI                                As String = "preserve ui"               ' Used to indicate that the ui should not be reset when opening
                                                                                                    ' a saved assessment report. This should only be used when generating
                                                                                                    ' Full Report or Summary Report variants of an assessment report.
                                                                                                    ' This should have a boolean value.
Public Const mgrDVRDAAddinVersion                           As String = "rda addin version"
Public Const mgrDVSSARAddinVersion                          As String = "ssar addin version"
Public Const mgrDVTemplate                                  As String = "template"                  ' The template used to create the assessment report
Public Const mgrDVVersion                                   As String = "version"                   ' This is defined in this AddIn and is the current version number

Public Const mgrTemporaryFileUniqueId                       As String = "(an=%1 pid=%2) %3"         ' Used to construct helpful and unique file names
Public Const mgrTemporaryFileDateFormat                     As String = "yyyy-mmm-dd hh.mm.ss"      ' Used to insert a date in the temporary html text file name


' Assessment report version types (determines which watermark is used)
Public Enum mgrVersionType
    mgrVersionTypeDraft
    mgrVersionTypeManagers
    mgrVersionTypeFinal
End Enum

' The view mode (determines whether the assessment report is read only)
Public Enum mgrViewMode
    mgrViewModeWrite
    mgrViewModeRead
    mgrViewModePrint
End Enum


Public g_wordEvents             As ApplicationEvents                                            ' Hooks application level events
Public g_eventLog               As LogIt                                                        ' Log file object should be accessed through the EventLog method
Public g_configuration          As Configuration                                                ' Used by all addins for configuration (environment) info
Public g_rootData               As RootData                                                     ' Attribute derived data from the main xml Assessment node attributes

Public g_htmlWordDocument       As Word.Document                                                ' Word document used to store #TODO#
Public g_xmlDocument            As MSXML2.DOMDocument60                                         ' DOM Document for the extracted xml used to create the assessment report
Public g_xmlInstructionData     As MSXML2.DOMDocument60                                         ' DOM Document for the instruction file used to create the assessment report
Public g_assessmentReport       As Word.Document                                                ' Word document used for the actual assessment report
Public g_htmlTextDocument       As HTMLDoc                                                      ' HTMLDoc object used to create the HTML document used as the rich text data source

Public g_richTextData           As Scripting.Dictionary                                         ' Rich text dictionary used as an index to the html document used as the data source


Public g_rdaCallStack           As Collection                                                   ' For parameter passing to the RDA addin
Public g_ssarCallStack          As Collection                                                   ' For parameter passing to the SSAR addin

Public g_hasBeenSubmitted       As Boolean                                                      ' True, when the xml data has been submitted to the web service

Public g_addinVersionManager    As String                                                       ' The version number of the manager addin (ar manager)
Public g_addinVersionRDA        As String                                                       ' The version number of the rda addin
Public g_addinVersionSSAR       As String                                                       ' The version number of the ssar addin
