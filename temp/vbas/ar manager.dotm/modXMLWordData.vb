Attribute VB_Name = "modXMLWordData"
'===================================================================================================================================
' Module:       modXMLWordData
' Purpose:      ???.
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


Public Function LoadInstructionData(ByVal theRDAXMLFullPath As String) As Boolean
    Dim errorText     As String
    Dim theParseError As MSXML2.IXMLDOMParseError

    ' Create a new xml document object to hold the Word Action data
    Set g_xmlInstructionData = New MSXML2.DOMDocument60

    With g_xmlInstructionData
        .Async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    ' Load the document xml document
    g_xmlInstructionData.Load theRDAXMLFullPath

    ' Force validation of the xml document
    Set theParseError = g_xmlInstructionData.parseError

    ' Report any error
    If theParseError.ErrorCode = 0 Then

        ' Set return status to success
        LoadInstructionData = True
    Else

        ' Report the error as cogently as possible
        With theParseError
            errorText = "The Word Action Data XML document failed to load due the following error." & vbCrLf & _
            "Error #: " & .ErrorCode & ": " & .reason & _
            "Line #: " & .Line & vbCrLf & _
            "Line Position: " & .linepos & vbCrLf & _
            "Position In File: " & .filepos & vbCrLf & _
            "Source Text: " & .srcText & vbCrLf & _
            "Document URL: " & .Url
        End With

        MsgBox errorText
    End If
End Function ' LoadInstructiondata
