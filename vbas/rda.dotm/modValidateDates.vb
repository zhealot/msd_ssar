Attribute VB_Name = "modValidateDates"
'===================================================================================================================================
' Module:       modValidateDates
' Purpose:      Validates all editable dates and then updates the underlying Assessment Report xml.
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

Private Const mc_errorDateValidation As String = "One or more dates were found to be incorrectly formatted." & vbCr & _
                                                 "To help you identify these dates they have had their background" & vbCr & _
                                                 "set to pinkish colour." & vbCr & vbCr & _
                                                 "Please correct these dates and resubmit the Assessment Report."

Private m_validateErrors As Long


Public Function ValidateAllEditableDates() As Boolean
    Const c_proc As String = "modValidateDates.ValidateAllEditableDates"

    Dim allActions  As VBA.Collection

    On Error GoTo Do_Error

    ' Reset the error counter
    m_validateErrors = 0

    ' Initialise the counters collection
    Set g_counters = New Counters
    
    ' The 'action' list used to build the Assessment Report is now used to validate just the editable date input areas
    Set allActions = g_instructions.Actions.actionList

    ' Parse out each "action" (which results in each of the actions in the list being carried out)
    DateValidateActions allActions

    ' Report any validation errors
    If m_validateErrors = 0 Then
        ValidateAllEditableDates = True
    Else

        MsgBox mc_errorDateValidation, vbExclamation Or vbOKOnly, rdaTitle
        m_validateErrors = 0
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' ValidateAllEditableDates

'=======================================================================================================================
' Procedure:    DateValidateActions
' Purpose:      Main processing loop for validating all editable date input areas.
' Note 1:       This validates 'actionInsert' of type rdaDataFormatDateLong and rdaDataFormatDateShort.
' Note 2:       This procedure is indirectly recursive.
'
' On Entry:     allActions          A Collection object that contains the list of 'actions' to be carried out.
'=======================================================================================================================
Private Sub DateValidateActions(ByVal allActions As VBA.Collection)
    Const c_proc As String = "modValidateDates.DateValidateActions"

    Dim errorText As String
    Dim theAction As Object

    On Error GoTo Do_Error

    ' Check that the collection exists before using it
    If allActions Is Nothing Then
        Exit Sub
    End If

    ' Iterate all ActionSetup, ActionAdd and ActionInsert objects in the collection
    For Each theAction In allActions

        ' There are four types of objects in this collection so choose the method appropriate to the object
        Select Case TypeName(theAction)
        Case rdaTypeActionAdd
            DateValidateAdd theAction

        Case rdaTypeActionInsert
            DateValidateInsert theAction

        Case rdaTypeActionLink, rdaTypeActionRename, rdaTypeActionSetup
            ' We do not need to do anything for these actions

        Case Else
            errorText = Replace$(mgrErrTextUnknownActionVerbType, mgrP1, TypeName(theAction))
            Err.Raise mgrErrNoUnknownActionVerbType, c_proc, errorText
        End Select
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DateValidateActions

'=======================================================================================================================
' Procedure:    DateValidateAdd
' Purpose:      Sub-block processing loop for validating all editable date input areas.
' Notes:        This procedure is indirectly recursive.
'
' On Entry:     info                The ActionAdd object to be actioned.
'=======================================================================================================================
Private Sub DateValidateAdd(ByVal info As ActionAdd)
    Const c_proc As String = "modValidateDates.DateValidateAdd"

    Dim index    As Long
    Dim theQuery As String
    Dim subNodes As MSXML2.IXMLDOMNodeList

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'rda instruction.xml' file
    If info.Break Then Stop

    ' Increment the depth counter so that we know which counter to use
    g_counters.IncrementDepth

    ' Update any predicate placeholders in the xpath query with their appropriate indexed value.
    ' Since we have not started updating the counters for the current nesting depth yet this lags by depth-1
    theQuery = g_counters.UpdatePredicates(info.Test)

    ' Use the ActionAdd objects Test string as an xpath query to retrieve the matching nodes
    Set subNodes = g_xmlDocument.SelectNodes(theQuery)

    ' There is only something to do if the xpath query actually returned one or more nodes
    If subNodes.Length > 0 Then

        ' Iterate all retrieved nodes
        For index = 1 To subNodes.Length

            ' Increment the counter for the current level
            g_counters.Counter = index

            ' Perform any nested actions in the current ActionAdd object
            DateValidateActions info.SubActions
        Next
    End If

    ' Zero the counter for the current level and then decrease the depth (nested) level
    g_counters.ResetCounter
    g_counters.DecrementDepth

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DateValidateAdd

'=======================================================================================================================
' Procedure:    DateValidateInsert
' Purpose:      Processes the Insert action when validating editable date input areas.
' Notes:        Only interested in types rdaDataFormatDateLong and rdaDataFormatDateShort.
'
' On Entry:     info                The ActionInsert object to be actioned.
'=======================================================================================================================
Public Sub DateValidateInsert(ByVal info As ActionInsert)
    Const c_proc As String = "modValidateDates.DateValidateInsert"

    Dim dataBookmarkName As String
    Dim dataNode         As MSXML2.IXMLDOMNode
    Dim dateInputArea    As Word.Range
    Dim dummy            As Date
    Dim thedate          As String
    Dim theQuery         As String

    On Error GoTo Do_Error

    With info

        ' This statement is to aid debugging of the 'rda instruction.xml' file
        If .Break Then Stop

        ' We only need to validate editable input areas
        If .Editable Then

            ' Perform the appropriate update type
            Select Case .DataFormat
            Case rdaDataFormatDateLong, rdaDataFormatDateShort

                ' Replace the predicate place holder (if present) with the predicate
                ' value to build the Bookmark name of an editable date area
                dataBookmarkName = ReplacePatternData(.BookmarkPattern, .PatternData)
                dataBookmarkName = g_counters.UpdatePredicates(dataBookmarkName)

                ' See if the Bookmark actually exists
                If g_assessmentReport.bookmarks.Exists(dataBookmarkName) Then

                    ' Retrieve the bookmark contents
                    Set dateInputArea = g_assessmentReport.bookmarks(dataBookmarkName).Range
                    thedate = dateInputArea.Text

                    ' Null strings and valid dates are legal values
                    If LenB(thedate) > 0 Then

                        ' If the text successfully converts to a date then it's valid
                        On Error Resume Next
                        dummy = CDate(thedate)
                        If Err.Number = mgrErrNoTypeMismatch Then

                            ' Set the font colour of the date input area to make it stand
                            ' out so that the user knows it contains an invalid date
                            dateInputArea.Font.ColorIndex = wdRed
                            Err.Clear
                            m_validateErrors = m_validateErrors + 1
                        ElseIf Err.Number <> 0 Then
                            Err.Raise Err.Number, c_proc
                        Else

                            ' Get the Bookmarks corresponding Assessment Report xml data node so that we can update it
                            theQuery = g_counters.UpdatePredicates(.DataSource)
                            Set dataNode = g_xmlDocument.SelectSingleNode(theQuery)

                            ' Now update the corresponding Assessment Report xml
                            If Not dataNode Is Nothing Then
                                dataNode.Text = thedate
                            End If

                            ' Clear the background colour just in case we previously set it to highlight an error
                            dateInputArea.Font.ColorIndex = wdAuto
                        End If
                    End If
                End If

            Case rdaDataFormatText, rdaDataFormatLong, rdaDataFormatDateLong, rdaDataFormatDateShort, rdaDataFormatTick
                ' There is no need to do anything for these 'actions'

            End Select
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DateValidateInsert
