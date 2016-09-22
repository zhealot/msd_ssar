Attribute VB_Name = "modUpdateAssessmentReportXML"
'===================================================================================================================================
' Module:       modUpdateAssessmentReportXML
' Purpose:      Contains the code that updates the Assessment Reports xml using the RichText (editable) areas that have been written
'               to a 'Filtered HTML' file.
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

Private m_newXHTMLNode      As xhtmlNode


'=======================================================================================================================
' Procedure:    ParseHTMLToXML
' Purpose:      Takes a Filtered HTML document and parses it to create xhtml which is then used to update the Assessment
'               Report xml.
'
' On Entry:     htmlFileFullPath    The path and name of the Filtered HTML input file.
' On Exit:      The Assessment Report xml (g_xmlDocument) has been updated.
'=======================================================================================================================
Public Sub ParseHTMLToXML(ByVal htmlFileFullPath As String)
    Const c_proc As String = "modUpdateAssessmentReportXML.ParseHTMLToXML"

    Dim bodyElement    As MSHTML.HTMLBody
    Dim errorText      As String
    Dim indexElement   As MSHTML.IHTMLDOMNode
    Dim loadedHTML     As String
    Dim theChildren    As MSHTML.IHTMLDOMChildrenCollection
    Dim theDoc         As MSHTML.HTMLDocument
    Dim updatedXMLFile As String
    Dim xmlParseError  As MSXML2.IXMLDOMParseError

'    On Error GoTo Do_Error

    EventLog c_proc

    ' Load the html file into a string
    loadedHTML = LoadHTMLFile(htmlFileFullPath)

    Set theDoc = New MSHTML.HTMLDocument
    theDoc.Body.innerHTML = loadedHTML
    
    ' Locate the body tag - ignoring everything before it
    Set bodyElement = theDoc.Body

    ' The 'body' elements first child is always a 'div' element with a class = 'WordSection1'
    ' so we want the 'div' elements child nodes
    Set theChildren = bodyElement.FirstChild.ChildNodes
    Set m_newXHTMLNode = Nothing

    For Each indexElement In theChildren

        ' This statement ensures that all nodes at the same hierarchical level in the input file are added to the correct base node
        If Not m_newXHTMLNode Is Nothing Then
            m_newXHTMLNode.ResetCurrentNode
        End If
        ParseHTMLNode indexElement

'''        Debug.Print m_newXHTMLNode.XML
    Next

    ' Flush the final xml update back to the RDA XML Dom Document.
    ' This needs to be done because normally the xml update is triggered by encountering the next query,
    ' but for the final entry in the file there is no next query so the xml does not get flushed.
    If Not m_newXHTMLNode Is Nothing Then
        m_newXHTMLNode.UpdateRDAXML
    End If

    ' Validate the xml we have just been updating
    Set xmlParseError = g_xmlDocument.Validate

    ' Check to see if the updated xml validated ok
    If xmlParseError.ErrorCode <> 0 Then

        ' Spit out an error message about whatever caused the validation to fail
        errorText = Replace$(mgrErrTextXMLSchemaValidate, mgrP1, g_xmlDocument.Url)
        errorText = Replace$(errorText, mgrP2, g_configuration.SchemaFullName)
        XMLErrorReporter xmlParseError, errorText, c_proc
    End If

    ' Now see if the updated xml should be saved to a file.
    ' The only purpose this serves is to provide a copy of the updated xml in a file for verification purposes.
    If g_configuration.SaveXMLTextFile Then
        updatedXMLFile = g_configuration.XMLTextFileFullName
        updatedXMLFile = MakeUniqueTemporaryFileName(updatedXMLFile)
        g_xmlDocument.Save updatedXMLFile
    End If
    Debug.Print c_proc & " completed."

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ParseHTMLToXML

Private Function ParseHTMLNode(ByVal theHTMLNode As MSHTML.IHTMLDOMNode) As Boolean
    Const c_proc As String = "modUpdateAssessmentReportXML.ParseHTMLNode"

    Dim leadinLength As Long
    Dim indexElement As MSHTML.IHTMLDOMNode
    Dim queryLength  As Long
    Dim theChildren  As MSHTML.IHTMLDOMChildrenCollection
    Dim warnText     As String
    Dim xpathQuery   As String

    On Error GoTo Do_Error

    If Not theHTMLNode Is Nothing Then
    
        ' Assign default return value
        ParseHTMLNode = True

        Select Case TypeName(theHTMLNode)
        Case "DispHTMLDOMTextNode"
            m_newXHTMLNode.AddTextNodeToCurrentXMLNode theHTMLNode
            ParseHTMLNode = False

        Case "HTMLParaElement"
            If theHTMLNode.NodeType = 1 Then
                If Left$(theHTMLNode.innerText, Len(mgrHTMLBookmarkedBlockLeadIn)) = mgrHTMLBookmarkedBlockLeadIn Then
                    leadinLength = Len(mgrHTMLBookmarkedBlockLeadIn)
                    queryLength = Len(theHTMLNode.innerText) - leadinLength - Len(mgrHTMLBookmarkNameEnd)
                    xpathQuery = Mid$(theHTMLNode.innerText, leadinLength + 1, queryLength)

                    ' See if we need to flush the previous XMTML object
                    If Not m_newXHTMLNode Is Nothing Then
                        m_newXHTMLNode.UpdateRDAXML
                    End If
                    
                    ' Create a new XMTML object
                    Set m_newXHTMLNode = New xhtmlNode

                    ' Set the xpath query string for the new XMTML object
                    m_newXHTMLNode.Query = xpathQuery
                    Debug.Print m_newXHTMLNode.Query

                    ParseHTMLNode = False
                    Exit Function
                Else
                    m_newXHTMLNode.AddXMLChildNode theHTMLNode
                End If
            End If

        Case "HTMLSpanElement", "HTMLPhraseElement", "HTMLTable", "HTMLTableSection", "HTMLTableRow", "HTMLTableCol", "HTMLTableCell", _
             "HTMLLIElement", "HTMLUListElement"
            m_newXHTMLNode.AddXMLChildNode theHTMLNode

        Case "HTMLBRElement"
            m_newXHTMLNode.AddXMLChildNode theHTMLNode

        Case "HTMLCommentElement"
            ' It should be safe to ignore Comment Nodes

        Case "HTMLAnchorElement"
            m_newXHTMLNode.AddAnchorNodeToCurrentXMLNode theHTMLNode
            ParseHTMLNode = False

        Case "HTMLUnknownElement"

            ' This is here because we have encountered the situation where Word seems to be able to generate invalid html.
            ' In the specific case we saw the following: <p><span>xxx</span></span></p>
            ' The unmatched </span> tag generates an 'HTMLUnknownElement'. Since it's unmatched there's nothing we can do
            ' with it so we just ignore it and hope for the best.
            ParseHTMLNode = False
            Exit Function
            
        Case Else
            warnText = Replace$(mgrWarnUnexpectedHTMLNodeType, mgrP1, TypeName(theHTMLNode))
            MsgBox warnText, vbExclamation Or vbOKOnly, mgrTitle & " - " & c_proc
        End Select

        If theHTMLNode.HasChildNodes Then
            Set theChildren = theHTMLNode.ChildNodes

            For Each indexElement In theChildren
                If ParseHTMLNode(indexElement) Then

                    ' So that all child nodes of the current HTML node get added to the same xml node we
                    ' rewind the current node back one level to account for the child node added above
                    m_newXHTMLNode.MakeParentTheCurrentNode
                End If
            Next
        Else
        End If
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ParseHTMLNode

Private Function LoadHTMLFile(ByVal theHTMLFileFullPath As String) As String
    Const c_fsoForReading As Long = 1
    Const c_fsoForWriting As Long = 2

    Dim theFSO  As Scripting.FileSystemObject
    Dim theFile As Scripting.TextStream

    Set theFSO = New Scripting.FileSystemObject

    ' Open the file for reading
    Set theFile = theFSO.OpenTextFile(theHTMLFileFullPath, c_fsoForReading, True)

    ' Read the entire file contents
    If theFile.AtEndOfStream Then
        LoadHTMLFile = ""
    Else
        LoadHTMLFile = CorrectHTML(theFile.ReadAll)
    End If
End Function ' LoadHTMLFile

Private Function CorrectHTML(ByVal theHTMLText As String) As String
    ' Replace HTML entity names with the appropriate xhtml escaped character values
    theHTMLText = Replace$(theHTMLText, mgr_HTMLEN_NBSP, mgr_XMLECV_NBSP)
    theHTMLText = Replace$(theHTMLText, mgr_HTMLEN_LT, mgr_XMLECV_LT)
    theHTMLText = Replace$(theHTMLText, mgr_HTMLEN_GT, mgr_XMLECV_GT)

    ' Remove HTML attributes we don't want propagating into the xhtml
    theHTMLText = Replace$(theHTMLText, mgr_HTMLOT_ClassMsoNormal1, mgr_HTMLRT_ClassMsoNormal1)
    CorrectHTML = Replace$(theHTMLText, mgr_HTMLOT_ClassMsoNormal2, mgr_HTMLRT_ClassMsoNormal2)
End Function ' CorrectHTML
