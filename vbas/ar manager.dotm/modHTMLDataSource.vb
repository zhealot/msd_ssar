Attribute VB_Name = "modHTMLDataSource"
'===================================================================================================================================
' Module:       modHTMLDataSource
' Purpose:      Contains code used for producing the HTML file which is in turn used as the rich text data source.
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


'=======================================================================================================================
' Procedure:    AssignRangesToHTML
' Purpose:      Updates the g_richTextData dictionary object with Range objects which represent blocks of RichText in
'               the html file.
'               The generated html file is opened as a Word document. The Word document is parsed for block markers
'               which identify individual blocks of RichText. Each RichText block is prefaced with a bookmark name which
'               is used as a index for the dictionary object.
'               When a block is found the bookmark name is extracted and a Range object created for the RichText block.
'               The Range object is then added to the dictionary using the bookmark name as the key.
'
' On Exit:      The g_richTextData dictionary object has had the RichText Range objects added to it.
'=======================================================================================================================
Public Sub AssignRangesToHTML()
    Const c_proc As String = "modHTMLDataSource.AssignRangesToHTML"

    Dim bmEnd           As Long
    Dim bmEndLength     As Long
    Dim bmStart1        As Long
    Dim bmStartN        As Long
    Dim bmStartLength   As Long
    Dim bookmarkName    As String
    Dim eolCharacter    As Boolean
    Dim index           As Long
    Dim isEndOfLine     As Word.Range
    Dim lastSearchArea  As Word.Range
    Dim paragraphText   As String
    Dim penultimate     As Word.Range
    Dim searchArea      As Word.Range

    On Error GoTo Do_Error

    ' The text file containing the RichText (xhtml) has been generated so now open it as a Word Document
    OpenHTMLFileAsWordDocument

    ' Search the Word HTML document for the specified bookmark text
    Set searchArea = g_htmlWordDocument.Content

    ' Used to determine whether the line containing the bookmark name ends immediately
    ' after the bookmark name or the line contains addional data
    Set isEndOfLine = g_htmlWordDocument.Content
    isEndOfLine.Collapse wdCollapseStart

    ' Clear the search object in case it's been used by the user or other code
    ResetFindProperties searchArea

    ' Do this once as it's always the same
    bmStartLength = Len(mgrHTMLBookmarkedBlockLeadIn)
    bmEndLength = Len(mgrHTMLBookmarkNameEnd)
    bmStart1 = bmStartLength + 1
    bmStartN = bmStart1 + Len(mgrHTMLBookmarkedBlockLeadOut)

    ' Search for the bookmark start character sequence.
    ' The odd thing about Words Find is that the result Range is returned in the Range object
    ' we are doing the Find on, so in our case the result is returned in searchArea
    With searchArea.Find
        .Text = mgrHTMLBookmarkedBlockLeadIn
        Do While .Execute
            If .found Then
                index = index + 1

                ' If this is the second or subsequent find then setup the previous finds Range object
                If index > 1 Then

                    ' Set the correct end position for the Range object so that it excludes the final paragraph mark
                    lastSearchArea.End = searchArea.Start - bmStartLength - 1

                    ' Add the Range object to the html dictionary object using the bookmark name as the index
                    g_richTextData.Add bookmarkName, lastSearchArea
                End If

                ' The first paragraph of a successful search always contains the bookmark name
                paragraphText = searchArea.Paragraphs(1).Range.Text
                If index = 1 Then
                    bmEnd = InStr(bmStart1, paragraphText, mgrHTMLBookmarkNameEnd)

                    ' Now extract the bookmark name
                    bookmarkName = Mid$(paragraphText, bmStart1, bmEnd - bmStart1)


                    ' Check to see if the line ends after the bookmark or contains data.
                    ' If there is a vbcr after the bookmark leadout sequence then there is no following data.
                    isEndOfLine.End = searchArea.Start + bmEnd + bmEndLength
                Else
                    bmEnd = InStr(bmStartN, paragraphText, mgrHTMLBookmarkNameEnd)

                    ' Now extract the bookmark name
                    bookmarkName = Mid$(paragraphText, bmStartN, bmEnd - bmStartN)

                    ' Check to see if the line ends after the bookmark or contains data.
                    ' If there is a vbcr after the bookmark leadout sequence then there is no following data.
                    isEndOfLine.End = searchArea.Start + bmStart1 + Len(bookmarkName) + bmEndLength
                End If

                isEndOfLine.Start = isEndOfLine.End - 1
                eolCharacter = (isEndOfLine.Text = vbCr)

                If eolCharacter Then

                    ' Adjust the Search Results Range object so that the Range starts
                    ' after the end of the bookmark name end character sequence
                    searchArea.Start = isEndOfLine.End
                Else
                    searchArea.Start = isEndOfLine.Start
                End If

                ' Create a unique copy of the found search range as we need unique Range objects in the dictionary
                Set lastSearchArea = searchArea.Duplicate
            End If
        Loop
    End With

    ' The final find never finds a match but there is still text from this find that needs to be added to the dictionary
    If index > 1 Then

        ' The document structure results in the last paragraph of the document being a mgrHTMLBookmarkedBlockLeadOut
        ' sequence, so we can always ignore this paragraph. So we choose the end (excluding the paragraph mark) of
        ' the penultimate paragraph of the document.
        With g_htmlWordDocument.Paragraphs
            Set penultimate = .Item(.Count - 1).Range
        End With
        lastSearchArea.End = penultimate.End - 1

        ' Add the Range object to the html dictionary object using the bookmark name as the index
        g_richTextData.Add bookmarkName, lastSearchArea
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' AssignRangesToHTML

Private Sub ResetFindProperties(ByRef findArea As Word.Range)
    With findArea.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
    End With
End Sub ' ResetFindProperties

'=======================================================================================================================
' Procedure:    OpenHTMLFileAsWordDocument
' Purpose:      Opens the html file that contains the RichText as a Word document.
'
' On Exit:      g_HTMLWordDocument is the html file opened as a Word document.
'=======================================================================================================================
Public Sub OpenHTMLFileAsWordDocument()
    Const c_proc As String = "modHTMLDataSource.OpenHTMLFileAsWordDocument"

    On Error GoTo Do_Error

    ' Close the html document
    g_htmlTextDocument.CloseFile

    ' Now open the generated html text file as invisible Word document
    Set g_htmlWordDocument = Documents.Open(g_htmlTextDocument.FullPath, False, False, False, , , False, , , _
                                            WdOpenFormat.wdOpenFormatWebPages, , False, False)

    ' Destroy the HTMLDoc object used to create the xhtml text file
    Set g_htmlTextDocument = Nothing

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' OpenHTMLFileAsWordDocument

'=======================================================================================================================
' Procedure:    SetupHTMLDataSource
' Purpose:      Sets up the HTMLDoc object to create the html text file.
'
' On Exit:      g_HTMLTextDocument setup as the HTMLDoc object.
'=======================================================================================================================
Public Sub SetupHTMLDataSource()
    Const c_proc As String = "modHTMLDataSource.SetupHTMLDataSource"

    ' Create the HTML Object used to create the HTML file we will use as data source for Rich Text
    On Error GoTo Do_Error

    Set g_htmlTextDocument = New HTMLDoc

    ' Initialise the html object using the unique file name
    g_htmlTextDocument.Initialise g_configuration.WordHTMLTextFileFullName

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' SetupHTMLDataSource
