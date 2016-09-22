VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        LogIt
' Purpose:      General purpose file logging class module. Writes output to the specified text file.
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

Private Const mc_defaultDateFormat As String = "dd.mm.yy hh.mm.ss"

Private m_autoPad    As Boolean
Private m_dateFormat As String
Private m_enabled    As Boolean
Private m_file       As String
Private m_path       As String
Private m_timeStamp  As Boolean


'============================================================================================================
' Procedure:    AutoPad
' Purpose:      If the output text contains embedded CR or CRLF sequences any succeeding nonblank lines are
'               padded with spaces. This is so that follow on lines line up neatly in the log file.
'============================================================================================================
Public Property Get AutoPad() As Boolean
    AutoPad = m_autoPad
End Property
Public Property Let AutoPad(ByVal autoPadding As Boolean)
    m_autoPad = autoPadding
End Property

Public Sub BlankLine()
    Const c_proc As String = "LogIt.BlankLine"

    Dim hFile As Long

    On Error GoTo Do_Error

    ' Everything must be right for us to write a blank line to the log file
    If CanOutput Then

        ' Open log file for append
        hFile = FreeFile
        Open m_path & m_file For Append Access Write As hFile

        Print #hFile, vbNullString
        Close hFile
    End If
    
Do_Exit:
    On Error GoTo 0
    Exit Sub

Do_Error:
    enabled = False
    Err.Raise Err.Number, c_proc, Err.Description
End Sub ' BlankLine

'============================================================================================================
' Procedure:    CanOutput
' Purpose:      When True all conditions are correct for the passed in text to be written to the log file.
'============================================================================================================
Public Property Get CanOutput() As Boolean
    CanOutput = LenB(m_path) > 0 And LenB(m_file) > 0 And enabled
End Property

Private Sub Class_Initialize()
    DateFormat = mc_defaultDateFormat
    Me.enabled = True
End Sub

'============================================================================================================
' Procedure:    DateFormat
' Purpose:      Allows the default date/time format to be overridden.
' Note:         Setting the DateFormat does not automatically enable timestamping of log file entries.
'============================================================================================================
Public Property Get DateFormat() As String
    DateFormat = m_dateFormat
End Property
Public Property Let DateFormat(ByVal newDateFormat As String)
    m_dateFormat = newDateFormat
End Property

Public Property Get DateLength() As Long
    DateLength = Len(Format$(Now, DateFormat))
End Property

'============================================================================================================
' Procedure:    Enabled
' Purpose:      Allows logging to be turned on (True) and off (False).
'============================================================================================================
Public Property Get enabled() As Boolean
    enabled = m_enabled
End Property
Public Property Let enabled(ByVal enableLogFile As Boolean)
    m_enabled = enableLogFile
End Property

'============================================================================================================
' Procedure:    File
' Purpose:      Sets the name of the log file.
' Note:         The log file path is set independently (see Path property).
'============================================================================================================
Public Property Get File() As String
    File = m_file
End Property
Public Property Let File(ByVal fileName As String)
    m_file = fileName
End Property

'============================================================================================================
' Procedure:    FormatForOutput
' Purpose:      Performs any special formatting of the text to be written to the log file.
'
' On Entry:     outputText          The text to be output.
' Returns:      The correctly formatted and indented output text.
'============================================================================================================
Private Function FormatForOutput(ByVal outputText As String) As String
    Const c_proc As String = "LogIt.FormatForOutput"

    Dim index       As Long
    Dim lines()     As String
    Dim padSequence As String

    On Error GoTo Do_Error

    ' Prefix the first line of log text with a timestamp
    If TimeStamp Then
        outputText = Format$(Now, DateFormat) & vbTab & outputText
    End If

    ' When writing multiple lines separate each line with CR rather than CRLF
    outputText = Replace$(outputText, vbCrLf, vbCr)

    ' Do AutoPad (padding of follow on lines) if required
    If AutoPad Then

        ' We only need to AutoPad the second and subsequent lines so find out how many lines we have
        lines() = Split(outputText, vbCr)
        If UBound(lines) > 0 Then

            ' Only determine the pad character sequence once since all lines are padded the same
            If TimeStamp Then
                padSequence = Space$(DateLength) & vbTab
            End If

            ' Prefix the second and subsequent lines
            For index = 1 To UBound(lines)
                lines(index) = padSequence & lines(index)
            Next

            ' Now put everything back together as one string to be written to the log file
            outputText = Join$(lines, vbCr)
        End If
    End If

    ' Return the formatted text
    FormatForOutput = outputText

Do_Exit:
    Exit Function

Do_Error:
    enabled = False
    Err.Raise vbObjectError + 8002, c_proc, Err.Description
End Function ' FormatForOutput

'============================================================================================================
' Procedure:    Kill
' Purpose:      Deletes the current log file.
'============================================================================================================
Public Sub Kill()
    Const c_proc As String = "LogIt.Kill"

    On Error GoTo Do_Error

    ' Delete the current log file if it exists
    If LenB(Dir$(m_path & m_file)) > 0 Then
        VBA.Kill m_path & m_file
    End If

Do_Exit:
    On Error GoTo 0
    Exit Sub

Do_Error:
    enabled = False
    Err.Raise vbObjectError + 8001, c_proc, Err.Description
End Sub ' Kill

'============================================================================================================
' Procedure:    Output
' Purpose:      Writes a string of text to the log file.
' Note:         When writting multiple lines use vbCr as the line separator.
'
' On Entry:     outputText          A string of text (which may contain multiple lines) to be written to the
'                                   log file.
'============================================================================================================
Public Sub Output(ByVal outputText As String)
    Const c_proc As String = "LogIt.Output"

    Dim textOut As String

    On Error GoTo Do_Error

    If CanOutput Then

        ' Format the text we need to output
        textOut = FormatForOutput(outputText)

        ' Write the log entry to the log file
        WriteToFile textOut
    End If

Do_Exit:
    On Error GoTo 0
    Exit Sub

Do_Error:
    enabled = False
    Err.Raise Err.Number, c_proc, Err.Description
End Sub ' Output

Private Sub WriteToFile(ByVal outputText As String)
    Const c_proc As String = "LogIt.WriteToFile"

    Dim hFile   As Long

    On Error GoTo Do_Error

    ' Open log file for append
    hFile = FreeFile
    Open m_path & m_file For Append Access Write As hFile

    ' Write the output and tidy up
    Print #hFile, outputText
    Close hFile

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub '

'============================================================================================================
' Procedure:    Path
' Purpose:      Sets the path of the log file.
' Note:         The log file name is set independently (see File property).
'============================================================================================================
Public Property Get path() As String
    path = m_path
End Property
Public Property Let path(ByVal usePath As String)
    usePath = Trim$(usePath)
    If Right$(usePath, 1) <> "\" Then usePath = usePath & "\"
    m_path = Trim$(usePath)
End Property

'============================================================================================================
' Property:     TimeStamp
' Purpose:      When True a timestamp prefixes each line written to the log file.
'============================================================================================================
Public Property Get TimeStamp() As Boolean
    TimeStamp = m_timeStamp
End Property
Public Property Let TimeStamp(ByVal timeStampOutput As Boolean)
    m_timeStamp = timeStampOutput
End Property
