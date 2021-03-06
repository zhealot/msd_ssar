Attribute VB_Name = "modSubmitToRemedyRDA"
'===================================================================================================================================
' Module:       modSubmitToRemedyRDA
' Purpose:      Submits the updated Assessment Report xml to the Remedy/RDA webservice.
'
' Author:       Peter Hewett - Inner Word Limited (innerword@xnet.co.nz)
' Copyright:    Ministry of Social Development (MSD) �2016 All rights reserved.
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

Private Const mc_soapEncoding           As String = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>"
Private Const mc_soapEnvelopeStart      As String = "<SOAP-ENV:Envelope xmlns:SOAPSDK1=""http://www.w3.org/2001/XMLSchema"" " & _
                                                    "xmlns:SOAPSDK2=""http://www.w3.org/2001/XMLSchema-instance"" " & _
                                                    "xmlns:SOAPSDK3=""http://schemas.xmlsoap.org/soap/encoding/"" " & _
                                                    "xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
Private Const mc_soapBodyStart          As String = "<SOAP-ENV:Body>"
Private Const mc_infopathWrapperBegin1  As String = "<tns1:sendInfopathToRemedy xmlns:tns1=""http://infopath.cyf.govt.nz"">"
Private Const mc_infopathWrapperBegin2  As String = "<tns1:infopathDocument>"
Private Const mc_infopathWrapperEnd2    As String = "</tns1:infopathDocument>"
Private Const mc_infopathWrapperEnd1    As String = "</tns1:sendInfopathToRemedy>"
Private Const mc_soapBodyEnd            As String = "</SOAP-ENV:Body>"
Private Const mc_soapEnvelopeEnd        As String = "</SOAP-ENV:Envelope>"

'Public Variables to maintain the session through GET requests
Public mrhSession2 As Variant
Public mrhLast2 As Variant

'=======================================================================================================================
' Procedure:    SubmitXMLToWebService
' Purpose:      Using soap submit the modified Assessment Report xml to the Remedy/RDA webservice.
' Notes 1:      We only close the Assessment Report document if the Submit works.
' Notes 2:      If the Assessment Report document is Closed or Cancelled from the RDA Tab (suppressErrors = True) the
'               Assessment Report document is always closed.
' Notes 3:      Whether it works or not the submit worked the UI is reset as there is nothing more the user can do with
'               the Assessment Report.
'
' On Entry:     suppressErrors      True if webservice (not errors in general) error reporting should be suppressed.
'
' Returns:      True if the submit to webservice worked.
'=======================================================================================================================
Public Function SubmitXMLToWebService(Optional ByVal suppressErrors As Boolean = False) As Boolean
    Const c_proc As String = "modSubmitToRemedyRDA.SubmitXMLToWebService"

    Dim client      As WebClient
    Dim request     As WebRequest
    Dim response    As WebResponse

    On Error GoTo Do_Error

    ' Create WebClient and WebRequest objects
    Set client = New WebClient
    Set request = New WebRequest

    ' Setup the WebClient object
    client.BaseUrl = g_rootData.MainSaveURL
    client.EnableAutoProxy = False
    client.FollowRedirects = True

    ' Setup the WebRequest object
    request.Method = WebMethod.HttpPost
    request.AddHeader "SOAPAction", "InfopathWebServiceService"
    request.Format = WebFormat.XML
    request.Body = BuildSOAPMessage

    'Calls the Function that does the two GET requests and stores
    'the MRH session cookie values to the public variables
    GetMRHSession
    
    'Debug.Print "MRHsession cookie for the post request:" & mrhSession2
    'Debug.Print "LastMRH_Session cookie for the post request:" & mrhLast2
           
    'MRH session cookies added to the POST request
    request.AddCookie "MRHSession", mrhSession2
    request.AddCookie "LastMRH_Session", mrhLast2

    ' Send request to the webservice.
    ' If the response is redirected assume that its the F5 and that we need to establish
    ' a session this should be handled by the changes to the WebClient.Execute method.
    Set response = client.Execute(request)

    ' Check the response to make sure we succeeded in submitting the xml to the webservice
    If response.StatusCode = WebStatusCode.OK Then

        ' Indicate that the webservice submit worked
        SubmitXMLToWebService = True
    Else
        If Not suppressErrors Then

            ' Raise an error to report the WebClient error response
            ReportWebClientError response, g_rootData.MainSaveURL
        End If
    End If

Do_Exit:

    ' Irrespective of whether or not the Submit/Cancel/Close worked, allow the user to close the Assessment Report
    g_hasBeenSubmitted = True

    ' At this point the RDA Tab is no longer usable for this Assessment Report so disable it
    ResetAlmostEverything
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' SubmitXMLToWebService

'=======================================================================================================================
' Procedure:    BuildSOAPMessage
' Purpose:      Wraps the Assessment Report xml being sent to the Remedy/RDA webservice in a SOAP wrapper.
'
' Returns:      The Assessment Reports xml in a SOAP wrapper.
'=======================================================================================================================
Private Function BuildSOAPMessage() As String
    Const c_proc As String = "modSubmitToRemedyRDA.BuildSOAPMessage"

    On Error GoTo Do_Error

    ' Construct the soap and InfoPath wrapper required before the actual xml
    BuildSOAPMessage = mc_soapEncoding & _
                       mc_soapEnvelopeStart & _
                       mc_soapBodyStart & _
                       mc_infopathWrapperBegin1 & _
                       mc_infopathWrapperBegin2
    
    ' Now add the escaped xml
    BuildSOAPMessage = BuildSOAPMessage & EscapeXML(g_xmlDocument.XML)

    ' Now add the InfoPath and soap wrappers required after the actual xml
    BuildSOAPMessage = BuildSOAPMessage & mc_infopathWrapperEnd2 & _
                       mc_infopathWrapperEnd1 & _
                       mc_soapBodyEnd & _
                       mc_soapEnvelopeEnd

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' BuildSOAPMessage

Public Sub CloseAssessmentReportDocument()
    Const c_proc As String = "modSubmitToRemedyRDA.CloseAssessmentReportDocument"

    On Error GoTo Do_Error

    EventLog "Closing assessment report document: " & g_assessmentReport.FullName, c_proc

    ' This global flag must be set or ApplicationEvents.m_wordApp_DocumentBeforeClose
    ' will not allow the Assessment Report document to be closed
    g_hasBeenSubmitted = True

    ' An Assessment Report document is closed this way in an attempt to circumvent a problem
    ' where MSD are being left with an orphan window when the document is closed!
    ' Bizarrely the orphan window is not in Words Windows collection object.
    With g_assessmentReport
        .AttachedTemplate.Saved = True
        .Close wdDoNotSaveChanges
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CloseAssessmentReportDocument

Public Sub SaveAndCloseAssessmentReportDocument()
    Const c_proc As String = "modSubmitToRemedyRDA.SaveAndCloseAssessmentReportDocument"

    Dim arFullName  As String

    On Error GoTo Do_Error

    ' Get the path and name to use for saving the Assessment Report
    arFullName = g_configuration.AssessmentReportFileFullName

    EventLog "Saving assessment report document as: " & arFullName, c_proc

    ' Save the assessment report document, this is only required if the submit to webservice fails. In the event of a failure the
    ' file is reopened so that the user can see that they have not lost any work. If the submit works the document is deleted.
    ' This is all to get around the orphan window problem, where submitting the xml to the webservice and then closing the Assessment
    ' Report document leaves the document window open, but not part of Words Document objects count!
    g_assessmentReport.SaveAs2 arFullName, wdFormatDocumentDefault, True, vbNullString, False

    CloseAssessmentReportDocument

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' SaveAndCloseAssessmentReportDocument

Private Function EscapeXML(ByVal theXML As String) As String
    Dim escapedXML  As String

    ' & is replaced with &amp;
    escapedXML = Replace$(theXML, "&", "&amp;")

    ' < is replaced with &lt;
    escapedXML = Replace$(escapedXML, "<", "&lt;")

    ' > is replaced with &gt;
    EscapeXML = Replace$(escapedXML, ">", "&gt;")
End Function ' EscapeXML

Private Sub ReportWebClientError(ByVal response As WebResponse, _
                                 ByVal webServiceURL As String)
    Dim errorText As String

    errorText = Replace$(mgrErrTextFailedToSubmitXMLToWebservice, mgrP1, webServiceURL)
    errorText = Replace$(errorText, mgrP2, WebClientErrorNumberToText(response.StatusCode))
    errorText = Replace$(errorText, mgrP3, response.StatusDescription)

    ' This is a rare instance where we raise an error without instantiating our own error handler. We
    ' do this because this procedure is just a 'helper' and we want the caller to handle the error.
    Err.Raise mgrErrNoFailedToSubmitXMLToWebservice, , errorText
End Sub ' ReportWebClientError

Private Function WebClientErrorNumberToText(ByVal webClientErrorNumber As WebStatusCode) As String
    Select Case webClientErrorNumber
    Case WebStatusCode.OK
        WebClientErrorNumberToText = "OK"
    Case WebStatusCode.created
        WebClientErrorNumberToText = "Created"
    Case WebStatusCode.BadGateway
        WebClientErrorNumberToText = "Bad Gateway"
    Case WebStatusCode.BadRequest
        WebClientErrorNumberToText = "Bad Request"
    Case WebStatusCode.Forbidden
        WebClientErrorNumberToText = "Forbidden"
    Case WebStatusCode.GatewayTimeout
        WebClientErrorNumberToText = "Gateway Timeout"
    Case WebStatusCode.InternalServerError
        WebClientErrorNumberToText = "Internal Server Error"
    Case WebStatusCode.NoContent
        WebClientErrorNumberToText = "No Content"
    Case WebStatusCode.NotFound
        WebClientErrorNumberToText = "Not Found"
    Case WebStatusCode.NotModified
        WebClientErrorNumberToText = "Not Modified"
    Case WebStatusCode.RequestTimeout
        WebClientErrorNumberToText = "Request Timeout"
    Case WebStatusCode.ServiceUnavailable
        WebClientErrorNumberToText = "Service Unavailable"
    Case WebStatusCode.Unauthorized
        WebClientErrorNumberToText = "Unauthorized"
    Case WebStatusCode.UnsupportedMediaType
        WebClientErrorNumberToText = "Unsupported Media Type"
    End Select
End Function ' WebClientErrorNumberToText

'Function to create session through GET requests that uses MRHSession and LastMRH_Session cookies
'The MRHSession and LastMRH_Session cookie values for the second request are stored in the public variables
Public Function GetMRHSession()
    Const c_proc As String = "modSubmitToRemedyRDA.GetMRHSession"

    Dim client1      As WebClient
    Dim client2      As WebClient
    Dim request1     As WebRequest
    Dim response1    As WebResponse
    Dim request2     As WebRequest
    Dim response2    As WebResponse
    Dim hostUrl As String
    Dim compUrl As String
    Dim urlArraySplit() As String
    Dim mrhSession1 As Variant
    Dim mrhLast1 As Variant

    On Error GoTo Do_Error
   
    compUrl = g_rootData.MainSaveURL
    urlArraySplit = Split(compUrl, "nz/")
    hostUrl = urlArraySplit(0) & "nz"
    ' Create WebClient and WebRequest objects
    Set client1 = New WebClient
    Set client2 = New WebClient
    Set request1 = New WebRequest
    Set request2 = New WebRequest

    ' Setup the WebClient object
    client1.BaseUrl = hostUrl & "/foo/"
    client2.BaseUrl = hostUrl & "/my.policy"
    client1.EnableAutoProxy = False
    client1.FollowRedirects = False
    client2.EnableAutoProxy = False
    client2.FollowRedirects = False

    
    
    
    request1.Method = WebMethod.HttpGet
    request1.Format = WebFormat.plainText
    request1.Body = "GET request to a random url in RDA to retrieve MRH session cookies"
    request2.Method = WebMethod.HttpGet
    request2.Format = WebFormat.plainText
    request2.Body = "GET request to my.policy in RDA to retrieve MRH session cookies. Adding the MRH session cookies from the previous GET request."
    
    ' Send GET request to the webservice.
    Set response1 = client1.Execute(request1)

    'Check the response status code is not a redirect
    'Found ther error handling code in http://stackoverflow.com/questions/1038006/good-patterns-for-vba-error-handling
    If response1.StatusCode <> 302 Then
        Err.Raise vbObjectError + 513, "modSubmitToRemedyRDA.GetMRHSession", "Did not return a redirect response for first Get request"
        'Debug.Print "Get Request 1 status code:" & response1.StatusCode
        'Debug.Print "MRH session cookie in the first request not sent"
    End If
    
    mrhSession1 = WebHelpers.FindInKeyValues(response1.Cookies, "MRHSession")
    mrhLast1 = WebHelpers.FindInKeyValues(response1.Cookies, "LastMRH_Session")
    'Debug.Print "MRH Session:" & mrhSession1
    'Debug.Print "Last MRH:" & mrhLast1
    
    'The Cookies from the response of the first
    'request is added as cookies to the second request
    request2.AddCookie "MRHSession", mrhSession1
    request2.AddCookie "LastMRH_Session", mrhLast1
    
    Set response2 = client2.Execute(request2)

   'Check the response status code is not a redirect
   If response2.StatusCode <> 302 Then
        Err.Raise vbObjectError + 514, "modSubmitToRemedyRDA.GetMRHSession", "Did not return a redirect response for second Get request"
        'Debug.Print "Get Request 2 status code:" & response2.StatusCode
        'Debug.Print "MRH session cookie in the first request not sent"
    End If
    'Check the response to check the redirect
    'The MRH session cookies for the first get request are stored
    
    mrhSession2 = WebHelpers.FindInKeyValues(response2.Cookies, "MRHSession")
    mrhLast2 = WebHelpers.FindInKeyValues(response2.Cookies, "LastMRH_Session")
    'Debug.Print "MRH Session:" & mrhSession2
    'Debug.Print "Last MRH:" & mrhLast2
    


Do_Exit:
    Exit Function


Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
    

End Function ' GetMRHSession

