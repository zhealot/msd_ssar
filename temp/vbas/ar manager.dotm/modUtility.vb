Attribute VB_Name = "modUtility"
'===================================================================================================================================
' Module:       modUtility
' Purpose:      Contains general purpose utility code.
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


' XQ = Xpath Queries to obtain data used to construct temporary file names
Private Const mc_XQAssessmentNumber       As String = "/Assessment/assessmentNumber"
Private Const mc_XQProviderId             As String = "/Assessment/provider/id"



'===================================================================================================================================
' Procedure:    PadR
' Purpose:      Right pads the input string with spaces to the required length.
' Notes:        Does not truncate the input string if it is longer than the specified length.
' Date:         31/05/2016
'
' On Entry:     inputText           The text to be right padded.
'               padWidth            The required string length.
' Returns:      The padded input string.
'===================================================================================================================================
Public Function PadR(ByVal inputText As String, _
                     ByVal padWidth As Long) As String

    If Len(inputText) < padWidth Then
        PadR = inputText & String$(padWidth - Len(inputText), " ")
    Else
        PadR = inputText
    End If
End Function ' PadR

Public Function PromptForFile(ByVal dataFileFullPath As String) As String
    Const c_proc As String = "modUtility.PromptForFile"

    Dim theDialog As Office.FileDialog

    On Error GoTo Do_Error

    ' Display the file picker dialog
    Set theDialog = Application.FileDialog(msoFileDialogFilePicker)

    ' Setup the various dialog properties
    With theDialog
        .AllowMultiSelect = False
        .InitialView = msoFileDialogViewDetails

        .title = "Select alternate input file"
        .ButtonName = "Open Rep"

        .Filters.Clear
        .Filters.Add "InfoPath rep file", "*." & mgrFIRDAXMLDataFileType

        .InitialFileName = dataFileFullPath

        If .Show Then
            PromptForFile = .SelectedItems(1)
        Else
            PromptForFile = dataFileFullPath
        End If
    End With

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' PromptForFile

Public Function MakeUniqueTemporaryFileName(ByVal theFileName As String) As String
    Const c_proc As String = "modUtility.MakeUniqueTemporaryFileName"

    Dim assessmentNumber    As String
    Dim dataNode            As MSXML2.IXMLDOMNode
    Dim providerId          As String
    Dim uniqueId            As String

    On Error GoTo Do_Error

    If Not g_xmlDocument Is Nothing Then

        ' Get the Assessment Number
        Set dataNode = g_xmlDocument.SelectSingleNode(mc_XQAssessmentNumber)
        If Not dataNode Is Nothing Then
            assessmentNumber = dataNode.Text
        End If

        ' Get the Provider Id
        Set dataNode = g_xmlDocument.SelectSingleNode(mc_XQProviderId)
        If Not dataNode Is Nothing Then
            providerId = dataNode.Text
        End If

        ' The uniqueId const has three parameters:
        ' %1 = assessment number
        ' %2 = provider id
        ' %3 = date and time
        uniqueId = Replace$(mgrTemporaryFileUniqueId, mgrP1, assessmentNumber)
        uniqueId = Replace$(uniqueId, mgrP2, providerId)
        uniqueId = Replace$(uniqueId, mgrP3, Format$(Now, mgrTemporaryFileDateFormat))

        ' Now use the unique id string we've built to create a unique file name
        MakeUniqueTemporaryFileName = Replace$(theFileName, mgrP1, uniqueId)
    Else
        Err.Raise mgrErrNoUnexpectedCondition, c_proc
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' MakeUniqueTemporaryFileName

'======================================================================================================================
' Procedure:    QualifyPath
' Purpose:      Forces a path name to be terminated with a backslash character. This allows file names to be appended
'               to the path name.
'
' On Entry:     inputPath           Path to append backslash to if necessary
' Returns:      Input path terminated with a backslash
'======================================================================================================================
Public Function QualifyPath(ByVal inputPath As String) As String
    If LenB(inputPath) > 0 Then
        If Right$(inputPath, 1) <> "\" Then
            QualifyPath = inputPath & "\"
        Else
            QualifyPath = inputPath
        End If
    Else
        QualifyPath = vbNullString
    End If
End Function ' QualifyPath

'======================================================================================================================
' Procedure:    UnqualifyPath
' Purpose:      Removes the terminating backslash character in a path name.
'
' On Entry:     inputPath           Path to remove backslash from if necessary
' Returns:      Input path without the terminating backslash
'======================================================================================================================
Public Function UnqualifyPath(ByVal inputPath As String) As String
    If LenB(inputPath) > 0 Then
        If Right$(inputPath, 1) = "\" Then
            UnqualifyPath = Left$(inputPath, Len(inputPath) - 1)
        Else
            UnqualifyPath = inputPath
        End If
    Else
        UnqualifyPath = vbNullString
    End If
End Function ' UnqualifyPath

'======================================================================================================================
' Procedure:    GetBaseName
' Purpose:      Returns the Filename part (no Extension) of the passed in path.
' E.g.:         Passed In                                   Returns
'               C:\my folder\my subfolder\testfile.txt      testfile
'               C:\my folder\my subfolder\                  vbNullString
'               testfile.txt                                testfile
'
' On Entry:     theFullPath         A full path or at least a file name.
' Returns:      The Filename (without the Extension) part of the passed in full path.
'======================================================================================================================
Public Function GetBaseName(ByVal theFullPath As String) As String
    Dim delimiter As Long
    Dim fileName  As String

    fileName = GetFileName(theFullPath)
    delimiter = InStrRev(fileName, ".")
    If delimiter > 0 Then
        GetBaseName = Left$(fileName, delimiter - 1)
    End If
End Function ' GetBaseName

'======================================================================================================================
' Procedure:    GetExtensionName
' Purpose:      Returns the Filename Extension part of the passed in path.
' E.g.:         Passed In                                   Returns
'               C:\my folder\my subfolder\testfile.txt      txt
'               C:\my folder\my subfolder\                  vbNullString
'               testfile.txt                                txt
'
' On Entry:     theFullPath         A full path or at least a file name.
' Returns:      The Extension part of the passed in full path.
'======================================================================================================================
Public Function GetExtensionName(ByVal theFullPath As String) As String
    Dim delimiter As Long
    Dim fileName  As String

    fileName = GetFileName(theFullPath)
    delimiter = InStrRev(fileName, ".")
    If delimiter > 0 Then
        GetExtensionName = Mid$(fileName, delimiter + 1)
    End If
End Function ' GetExtensionName

'======================================================================================================================
' Procedure:    GetFileName
' Purpose:      Returns the Filename and Extension part of the passed in path.
' E.g.:         Passed In                                   Returns
'               C:\my folder\my subfolder\testfile.txt      testfile.txt
'               C:\my folder\my subfolder\                  vbNullString
'               testfile.txt                                testfile.txt
'
' On Entry:     fullPath            A full path or at least a file name.
' Returns:      The filename (and Extension) part of the passed in full path.
'======================================================================================================================
Public Function GetFileName(ByVal FullPath As String) As String
    Dim delimiter As Long

    delimiter = InStrRev(FullPath, "\")
    If delimiter > 0 Then
        GetFileName = Mid$(FullPath, delimiter + 1)
    Else
        GetFileName = FullPath
    End If
End Function ' GetFileName

'======================================================================================================================
' Procedure:    GetPath
' Purpose:      Returns the Path part of the passed in path.
' E.g.:         Passed In                                   Returns
'               C:\my folder\my subfolder\testfile.txt      C:\my folder\my subfolder\
'               C:\my folder\my subfolder\                  C:\my folder\my subfolder\
'               testfile.txt                                vbNullString
'
' On Entry:     theFullPath         A full path or at least a path.
' Returns:      The Path part of the passed in full path.
'======================================================================================================================
Public Function GetPath(ByVal theFullPath As String) As String
    Dim delimiter As Long

    delimiter = InStrRev(theFullPath, "\")
    If delimiter > 0 Then
        GetPath = Left$(theFullPath, delimiter)
    Else
        GetPath = vbNullString
    End If
End Function ' GetPath
