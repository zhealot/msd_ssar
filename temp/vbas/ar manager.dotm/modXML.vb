Attribute VB_Name = "modXML"
'===================================================================================================================================
' Module:       modXML
' Purpose:      General purpose xml code, private in scope to this addin.
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
Option Private Module


'=======================================================================================================================
' Procedure:    LoadRemedyHTMLFile
' Purpose:      Loads a Remedy '.rep' html file, parses it to locate and extract the embedded xml, corrects the xml and
'               loads it into a DOMDocument object.
'
' On Entry:     theFullPath         The path and name of the Remedy '.rep' html file to be loaded.
'               theXMLFile          Ignored.
' On Exit:      theXMLFile          The newly created assessment report xml DOMDocument object.
' Returns:      True if the '.rep' file was successfully loaded, parsed, the xml extracted, corrected and loaded into a
'               DOMDocument object.
'=======================================================================================================================
Public Function LoadRemedyHTMLFile(ByVal theFullPath As String, _
                                   ByRef theXMLFile As MSXML2.DOMDocument60) As Boolean
    Const c_proc As String = "modXML.LoadRemedyHTMLFile"

    Dim repHTML As String

    On Error GoTo Do_Error

    EventLog c_proc

    ' Load the input '.rep' file, extracting and correcting the xml
    repHTML = ConvertRemedyRepHTMLToXML(theFullPath)

    ' Create a xml document object for the assessment report xml
    Set theXMLFile = New MSXML2.DOMDocument60

    ' Load the Assessment Report xml into the Assessment Report xml DOMDOcument object
    With theXMLFile
        .validateOnParse = False
        .resolveExternals = False
        .preserveWhiteSpace = False
        LoadRemedyHTMLFile = .LoadXML(repHTML)
    End With

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' LoadRemedyHTMLFile

'=======================================================================================================================
' Procedure:    ConvertRemedyRepHTMLToXML
' Purpose:      Loads a Remedy '.rep' html file, parses it to locate, extract and correct the embedded xml.
'
' On Entry:     repFileFullPath     The path and name of the Remedy '.rep' html file to be loaded.
' Returns:      The corrected xml that was extracted from the Remedy html.
'=======================================================================================================================
Public Function ConvertRemedyRepHTMLToXML(ByVal repFileFullPath As String) As String
    Const c_proc As String = "modXML.ConvertRemedyRepHTMLToXML"

    Dim bodyElement             As MSHTML.HTMLBody
    Dim childElement            As MSHTML.IHTMLDOMNode
    Dim indexElement            As MSHTML.IHTMLDOMNode
    Dim loadedHTML              As String
    Dim theDoc                  As MSHTML.HTMLDocument
    Dim xmlFound                As Boolean
    Dim xmlInputFileFullPath    As String
    Dim xmlParseError           As MSXML2.IXMLDOMParseError

    ' Load the html file into a string
    On Error GoTo Do_Error

    EventLog c_proc

    loadedHTML = LoadRepHTMLFile(repFileFullPath)

    Set theDoc = New MSHTML.HTMLDocument
    theDoc.Body.innerHTML = loadedHTML

    ' Free the loaded html string asap as it could be very large
    loadedHTML = vbNullString

    ' Locate the body tag - ignoring everything before it
    Set bodyElement = theDoc.Body

    ' Iterate the 'body' elements child elements until we find the 'TABLE' element
    For Each indexElement In bodyElement.ChildNodes
        If indexElement.nodeName = "TABLE" Then

            ' The 'TABLE' element should have a 'TBODY' child element
            Set childElement = indexElement.FirstChild
            If childElement.nodeName = "TBODY" Then

                ' The 'TBODY' element should have a 'TR' child element
                Set childElement = childElement.FirstChild
                If childElement.nodeName = "TR" Then

                    ' The 'TBODY' element should have two 'TD' child elements, we want the second one
                    Set childElement = childElement.ChildNodes(1)
                    If childElement.nodeName = "TD" Then
                        ConvertRemedyRepHTMLToXML = MakeValidXMLFromHTML(childElement.innerText)
                        xmlFound = True
                        Exit For
                    Else
                        RaiseErrorUnexpectedHTMLDocumentStructure repFileFullPath, c_proc
                    End If
                Else
                    RaiseErrorUnexpectedHTMLDocumentStructure repFileFullPath, c_proc
                End If
            Else
                RaiseErrorUnexpectedHTMLDocumentStructure repFileFullPath, c_proc
            End If
        End If
    Next

    If Not xmlFound Then
        RaiseErrorUnexpectedHTMLDocumentStructure repFileFullPath, c_proc
    End If

    ' Save the extracted and corrected xml if required
    If g_configuration.SaveRDAXMLFile Then
        xmlInputFileFullPath = Replace$(g_configuration.RDAXMLFileFullName, mgrP1, Format$(Now, mgrTemporaryFileDateFormat))
        CreateXMLFile xmlInputFileFullPath, ConvertRemedyRepHTMLToXML
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ConvertRemedyRepHTMLToXML

'=======================================================================================================================
' Procedure:    RaiseErrorUnexpectedHTMLDocumentStructure
' Purpose:      Stub to raise a custom error.
' Notes:        No error handler of our own is instantiated so that the error we raise looks like it comes from our
'               calling procedure.
'
' On Entry:     inputFileFullPath   The path and name to use in the error message.
'               errorProcedure      The procedure name that originally raised the error.
'=======================================================================================================================
Private Sub RaiseErrorUnexpectedHTMLDocumentStructure(ByVal inputFileFullPath As String, _
                                                      ByVal errorProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextUnexpectedHTMLDocumentStructure, mgrP1, inputFileFullPath)
    Err.Raise mgrErrNoUnexpectedHTMLDocumentStructure, errorProcedure, errorText
End Sub ' RaiseErrorUnexpectedHTMLDocumentStructure
                                                                                                
'=======================================================================================================================
' Procedure:    LoadRepHTMLFile
' Purpose:      Loads a ".rep" html file that contains the embedded Assessment Report xml.
'
' On Entry:     repFileFullPath     The path and name of the html .rep file to load.
' Returns:      The text of the loaded .rep file.
'=======================================================================================================================
Private Function LoadRepHTMLFile(ByVal repFileFullPath As String) As String
    Const c_proc As String = "modXML.LoadRepHTMLFile"
    Const c_fsoForReading As Long = 1
    Const c_fsoForWriting As Long = 2

    Dim theFSO  As Scripting.FileSystemObject
    Dim theFile As Scripting.TextStream

    On Error GoTo Do_Error

    EventLog c_proc

    Set theFSO = New Scripting.FileSystemObject

    ' Open the file for reading
    Set theFile = theFSO.OpenTextFile(repFileFullPath, c_fsoForReading, True)

    ' Read the entire file contents
    If theFile.AtEndOfStream Then
        LoadRepHTMLFile = vbNullString
    Else
        LoadRepHTMLFile = CorrectRepHTML(theFile.ReadAll)
    End If

    ' Destroy the FSO objects
    Set theFSO = Nothing
    Set theFile = Nothing

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' LoadRepHTMLFile

'=======================================================================================================================
' Procedure:    CorrectRepHTML
' Purpose:      Corrects the html text loaded from a Remedy '.rep' html file.
' Notes:        Converts html entity names to xml numeric character references.
'
' On Entry:     theHTMLText         The html text to be corrected.
' Returns:      The corrected html text.
'=======================================================================================================================
Private Function CorrectRepHTML(ByVal theHTMLText As String) As String
    Const c_proc As String = "modXML.CorrectRepHTML"

    On Error GoTo Do_Error

    EventLog c_proc

    ' Replace HTML entity names with the appropriate xhtml escaped character values
    theHTMLText = Replace$(theHTMLText, mgr_HTMLEN_NBSP, mgr_XMLECV_NBSP)
    theHTMLText = Replace$(theHTMLText, mgr_HTMLEN_LT, mgr_XMLECV_LT)
    CorrectRepHTML = Replace$(theHTMLText, mgr_HTMLEN_GT, mgr_XMLECV_GT)


Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' CorrectRepHTML

'=======================================================================================================================
' Procedure:    MakeValidXMLFromHTML
' Purpose:      Takes the XML embedded in an HTML wrapper (generated by Remedy) and corrects it so that it is valid XML.
' Notes:        When extracting what should be valid xml from the html wrapper certain characters (such as the ampersand
'               escape sequence '&amp;' are replaced with the actual '&' character which is invalid in the xml) so have
'               to be re-escaped.
'
' On Entry:     xmlInput            The xml extracted from an html wrapper.
' Returns:      A valid xml string.
'=======================================================================================================================
Public Function MakeValidXMLFromHTML(ByRef xmlInput As String) As String
    Const c_proc        As String = "modXML.MakeValidXMLFromHTML"

    ' This matches a naked '&' but not '&#99;', '&#999;' or '&#9999;' numeric entity codes (where 9 represents any number)
    Const c_ampFind     As String = "(?!\&#\d+;)\&"
    Const c_ampReplace  As String = "&amp;"

    Dim location    As Long
    Dim theRegEx    As VBScript_RegExp_55.RegExp

    On Error GoTo Do_Error

    EventLog c_proc

    ' Create and setup the RegEx object we use to clean up the xml
    Set theRegEx = New VBScript_RegExp_55.RegExp
    With theRegEx
        .MultiLine = True
        .Global = True
        .IgnoreCase = False
        .Pattern = c_ampFind
    End With

    ' For some reason the html generated by Remedy contains one or more characters before the very
    ' first '<' character of the xml, so we need to skip over this to get to the very first '<'
    location = InStr(xmlInput, "<")
    If location = 1 Then

        ' Just in case there are no unwanted characters before the first '<'.
        ' Clean up the xml by reinstating the necessary escape codes.
        MakeValidXMLFromHTML = theRegEx.Replace(xmlInput, c_ampReplace)
    Else

        ' When we do have unwanted characters before the first '<'
        ' Clean up the xml by reinstating the necessary escape codes.
        MakeValidXMLFromHTML = theRegEx.Replace(Mid$(xmlInput, location), c_ampReplace)
    End If
    
    ' The sequences '##60;' and '##62;' are used for xml content text to represent the xml numeric entity codes '&#60;' (<)
    ' and '&#62;' (>). This is so that we can tell the difference between the html representation of the xml '<' (&lt;) and
    ' '>' (&gt) which form the actual xml and '<' and '>' characters in the actual xml content.
    MakeValidXMLFromHTML = Replace$(MakeValidXMLFromHTML, "##60;", "&#60;")
    MakeValidXMLFromHTML = Replace$(MakeValidXMLFromHTML, "##62;", "&#62;")

'''EventLog vbCr & c_proc & " After processing:" & vbCr & MakeValidXMLFromHTML

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' MakeValidXMLFromHTML

'=======================================================================================================================
' Procedure:    LoadXmlFile
' Purpose:      Stub procedure to load an xml file.
' Note:         It is not the intent of this procedure to validate the xml file. That is the callers responsibility.
'
' On Entry:     theFullPath         The path and name of the xml file to be loaded.
'               theXMLFile          Ignored.
' On Exit:      theXMLFile          The DOMDocument object for the loaded xml file.
' Returns:      True if the file loaded without any Parse errors.
'=======================================================================================================================
Public Function LoadXmlFile(ByVal theFullPath As String, _
                            ByRef theXMLFile As MSXML2.DOMDocument60) As Boolean
    Const c_proc As String = "modXML.LoadXmlFile"

    On Error GoTo Do_Error

    ' Create a DOM Document object for the file we are going to load
    EventLog "Loading xml file: " & theFullPath, c_proc
    Set theXMLFile = New MSXML2.DOMDocument60

    ' Asynchronously load the specified file
    With theXMLFile
        .Async = False
        .validateOnParse = False
        .resolveExternals = False
        .preserveWhiteSpace = False
        LoadXmlFile = .Load(theFullPath)
    End With

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' LoadXmlFile

'=======================================================================================================================
' Procedure:    ValidateXMLFile
' Purpose:      Validates the passed in xml DOMDOcument against the specifed schema.
' Notes:        The schema is loaded into an xml DOMDocument which is then added to the schema cache used for
'               validation.
'
' On Entry:     xmlDocument         The xml DOMDocument to be validated.
'               xmlSchemaFullPath   The path and name of the schema file.
' On Exit:      xmlDocument         It now has the specified schema attached.
' Returns:      True if the DOMDocument could be validated against the specified schema.
'=======================================================================================================================
Public Function ValidateXMLFile(ByRef xmlDocument As MSXML2.DOMDocument60, _
                                ByVal xmlSchemaFullPath As String) As Boolean
    Const c_proc As String = "modXML.ValidateXMLFile"

    Dim errorText      As String
    Dim xmlParseError  As MSXML2.IXMLDOMParseError
    Dim xmlSchema      As MSXML2.DOMDocument60
    Dim xmlSchemaCache As MSXML2.XMLSchemaCache60

    On Error GoTo Do_Error

    ' Load the schema file
    EventLog "Loading schema: " & xmlSchemaFullPath, c_proc
    If LoadXmlFile(xmlSchemaFullPath, xmlSchema) Then

        ' Create the SchemaCache object
        Set xmlSchemaCache = New XMLSchemaCache60

        ' Add the schema file into the schema cache using the default namespace
        xmlSchemaCache.Add "", xmlSchema
    Else
        errorText = Replace$(mgrErrTextXMLSchemaLoad, mgrP1, xmlSchemaFullPath)
        XMLErrorReporter g_xmlDocument.parseError, errorText, c_proc
        Exit Function
    End If

    ' Bind the schema cache to the xml document object
    Set xmlDocument.schemas = xmlSchemaCache

    ' Validate the xml file
    EventLog "Validitating: " & xmlSchemaFullPath, c_proc
    Set xmlParseError = xmlDocument.Validate

    ' Report any errors
    If xmlParseError.ErrorCode = 0 Then
        ValidateXMLFile = True
    Else

        ' Spit out an error message about whatever caused the validation to fail
        errorText = Replace$(mgrErrTextXMLSchemaValidate, mgrP1, xmlDocument.Url)
        errorText = Replace$(errorText, mgrP2, xmlSchemaFullPath)
        XMLErrorReporter xmlParseError, errorText, c_proc
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ValidateXMLFile

Public Sub CreateXMLFile(ByVal xmlFileFullPath As String, _
                         ByRef xmlData As String)
    Const c_proc As String = "modXML.CreateXMLFile"

    Dim theFSO      As Scripting.FileSystemObject
    Dim theXMLFile  As Scripting.TextStream

    On Error GoTo Do_Error

    ' Create a File System Object to use in turn, to create a TextStream object
    Set theFSO = New Scripting.FileSystemObject

    ' Use the TextStream object to create the HTML Document file.
    ' The file is a ASCII file.
    Set theXMLFile = theFSO.CreateTextFile(xmlFileFullPath, True, False)

    theXMLFile.Write xmlData
    theXMLFile.Close

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CreateXMLFile
