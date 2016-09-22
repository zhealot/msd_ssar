VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTMLDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===================================================================================================================================
' Class:        HTMLDoc
' Purpose:      Creates an HTML document.
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

Private Const mc_errNoBase                      As Long = -10000
Private Const mc_ErrNoHTMLFileNotInitialised    As Long = mc_errNoBase

Private Const mc_ErrTextHTMLFileNotInitialised  As String = "The html output file has not been initialised!"

Private Const mc_htmlLeadIn                     As String = "<!DOCTYPE html><html><body><span style='font-family:Verdana'>"
Private Const mc_htmlLeadOut                    As String = "</span></body></html>"
Private Const mc_htmlDivTagOpen                 As String = "<div xmlns=""http://www.w3.org/1999/xhtml"">"
Private Const mc_htmlDivTagClose                As String = "</div>"

Private Const mc_searchDivTag                   As String = "<div "

Private Const mc_RegExSearch                    As String = ">\s*<"
Private Const mc_RegExReplace                   As String = "><"


Private m_theRegEx                 As RegExp


Private m_rdaHTMLTextDocumentument As Scripting.TextStream
Private m_closed                   As Boolean
Private m_initialised              As Boolean
Private m_htmlFullName             As String


Private Sub Class_Terminate()
    If Not m_closed Then
        CloseFile
    End If
End Sub ' Class_Terminate

Public Sub AddText(ByVal bookmarkName As String, _
                   ByVal xhtmlText As String)
    Const c_proc As String = "HTMLDoc.AddText"

    Dim htmlText As String

    On Error GoTo Do_Error

    If Not m_initialised Then
        Err.Raise mc_ErrNoHTMLFileNotInitialised, c_proc, mc_ErrTextHTMLFileNotInitialised
    End If

    ' Remove whitespace and newline characters, as they appear in the html file when we open it in Word as spaces and new line characters
    xhtmlText = m_theRegEx.Replace(xhtmlText, mc_RegExReplace)

    ' The first tag should be a 'div' tag, if it's not we need to add one
    If Left$(xhtmlText, Len(mc_searchDivTag)) <> mc_searchDivTag Then
        xhtmlText = mc_htmlDivTagOpen & xhtmlText & mc_htmlDivTagClose
    End If

    ' Add the bookmark name as text so that we can locate the entry we are really interested in
    htmlText = mgrHTMLBookmarkedBlockLeadIn & bookmarkName & mgrHTMLBookmarkNameEnd & xhtmlText & mgrHTMLBookmarkedBlockLeadOut
    WriteHTML htmlText

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' AddText

Friend Sub CloseFile()
    Const c_proc As String = "HTMLDoc.CloseFile"

    On Error GoTo Do_Error

    ' Make sure the file was initialised, if it wasn't there is no valid leadin html so adding lead out html is pointless
    If m_initialised Then

        ' Add the lead out html or else the file will not be valid
        AddHTMLLeadOut
    End If

    ' Close the file so it can be opened with Word or a Web Browser
    m_rdaHTMLTextDocumentument.Close

    ' Destroy the RegEx object
    Set m_theRegEx = Nothing

    ' Set a flag so that we don't end up here again on class termination
    m_closed = True

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CloseFile

Friend Sub Initialise(ByVal htmlFileFullName As String)
    Const c_proc As String = "HTMLDoc.Initialise"

    Dim theFSO As Scripting.FileSystemObject

    On Error GoTo Do_Error

    ' Setup the RegEx object used by the AddText procedure, we do it here since it is single purpose and reused.
    ' The RegEx object is used to remove all newline and whitespace following a close tag of any kind.
    Set m_theRegEx = New RegExp
    With m_theRegEx
        .MultiLine = True
        .Global = True
        .IgnoreCase = False
        .Pattern = mc_RegExSearch
    End With

    ' Store the full file name (path and name)
    m_htmlFullName = htmlFileFullName

    ' Create a File System Object to use in turn, to create a TextStream object
    Set theFSO = New Scripting.FileSystemObject

    ' Use the TextStream object to create the HTML Document file.
    ' The file is a Unicode file as the input file is UTF-8 and thus contains extended character values.
    Set m_rdaHTMLTextDocumentument = theFSO.CreateTextFile(m_htmlFullName, True, True)

    ' Add the leadin html that makes the file a valid html file
    AddHTMLLeadIn

    ' Set the initialised flag so this procedure is not called again
    m_initialised = True

Do_Exit:
    Set theFSO = Nothing
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Initialise

Private Sub AddHTMLLeadIn()

    ' Add the html lead-in wrapper, without it the html is not valid
    WriteHTML mc_htmlLeadIn
End Sub ' AddHTMLLeadIn

Private Sub AddHTMLLeadOut()

    ' Add the html lead-out wrapper, without it the html is not valid
    WriteHTML mc_htmlLeadOut
End Sub ' AddLeadOut

Private Sub WriteHTML(ByVal theHTML As String)

    ' Write the html to the file
    m_rdaHTMLTextDocumentument.Write theHTML
End Sub ' WriteHTML

Friend Property Get FullPath() As String
    FullPath = m_htmlFullName
End Property ' Get FullPath
