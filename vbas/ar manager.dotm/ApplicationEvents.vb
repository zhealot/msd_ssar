VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ApplicationEvents
' Purpose:      Sinks Application level events.
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

Private WithEvents m_wordApp As Word.Application
Attribute m_wordApp.VB_VarHelpID = -1

Private Sub Class_Initialize()

    ' Hook application level events
    Set m_wordApp = Application
End Sub ' Class_Initialize

Private Sub m_wordApp_DocumentBeforeClose(ByVal Doc As Document, _
                                          ByRef Cancel As Boolean)
    Const c_proc As String = "ApplicationEvents.m_wordApp_DocumentBeforeClose"

    On Error GoTo Do_Error

    Select Case Doc.Type
    Case wdTypeTemplate
        Exit Sub
    Case wdTypeDocument

        ' If the file is marked as an Assessment Report carry out further checking
        If IsAssessmentReport Then

            ' Carry on checking if we are not closing an html file
            If StrComp(GetExtensionName(Doc), "html", vbTextCompare) <> 0 Then

                ' At this point the file being closed is an Assessment report (and not an html document
                ' created from the Assessment Report template) we need to tell the user they can't close
                ' the document, as the file has not been Submitted or Closed from the RDA Tab
                If g_hasBeenSubmitted Then
                    Set g_assessmentReport = Nothing

                    ' Make sure that the user is not prompted to save changes if they
                    ' are Closing a Read Only View or Print View Assessment Report
                    Doc.Saved = True
                Else
                    MsgBox mgrWarnBeforeDenyClose, vbOKOnly Or vbExclamation, mgrTitle
                    Cancel = True
                End If
            End If
        End If
    End Select

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' m_wordApp_DocumentBeforeClose

Private Sub m_wordApp_DocumentBeforePrint(ByVal Doc As Document, Cancel As Boolean)
    Const c_proc As String = "ApplicationEvents.m_wordApp_DocumentBeforePrint"

    On Error GoTo Do_Error

    Debug.Print "m_wordApp_DocumentBeforePrint"

    ' We need to perform a full refresh (regenerate all refreshable Assessment Report areas) before printing an manifest
    ' 3 or SSAR Assessment Report in case the user has updated editable areas of the document that should be rolled-up into
    ' the refreshable areas of the document. Only manifest 3 and SSAR Assessment Reports need to be updated as these are the only
    ' editable Assessment Reports. Manifest 2 are read only and thus can never be changed from the time they are generated.
    If IsAssessmentReport Then
        EventLog "Printing assessment report: " & Doc.FullName, c_proc
        If g_assessmentReport Is Nothing Then
            EventLog "g_assessmentReport Is Nothing, so cannot proceed further", c_proc
        Else
            If g_rootData Is Nothing Then
                EventLog "g_rootData Is Nothing, so cannot check manifest version number", c_proc
            Else
                If g_rootData.ManifestVersion >= 3 Then

                    ' We only need to do a refresh if there is a chance that the assessment reports Executive
                    ' Summary section could be out of sync with the updatable parts of the document
                    If g_rootData.IsWritable Then
                        Manager_FullRefresh
                    End If
                End If
            End If
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' m_wordApp_DocumentBeforePrint

'===================================================================================================================================
' Procedure:    m_wordApp_DocumentOpen
' Purpose:      This event triggers the creation of an Assessment Report using the 'rep' file being opened.
' Note 1:       The 'rep' file is associated with Word, but is essentially an html file that contains an xml payload. So the 'rep'
'               file is closed as a Word document and loaded by other code into an html parser.
'
' On Entry:     Doc                 The document being opened.
'===================================================================================================================================
Private Sub m_wordApp_DocumentOpen(ByVal Doc As Document)
    Const c_proc As String = "ApplicationEvents.m_wordApp_DocumentOpen"

    Dim fullPathName    As String
    Dim isHTMLFile      As Boolean
    Dim isRepFile       As Boolean

    On Error GoTo Do_Error

    ' See if the file being opened is the one we are interested in
    isRepFile = (StrComp(GetExtensionName(Doc.Name), mgrFIRDAXMLDataFileType, vbTextCompare) = 0)
    If isRepFile Then
        EventLog "File being opened is: " & Doc.FullName, c_proc

        Doc.ActiveWindow.Visible = False
        CommandBars("XML Document").Visible = False

        ' Capture the full path, so that we can reopen the file as an html file
        fullPathName = Doc.FullName

        ' Close the document as we are not interested in it as a Word document
        Doc.Close wdDoNotSaveChanges

        ' We can only support one active Assessment Report so make sure that the user is not trying to create a second
        If g_assessmentReport Is Nothing Then

            ' Kick off the production of the Assessment Report
            GenerateAssessmentReport fullPathName
        Else

            ' Warn the user that they can only have one active Assessment Report
            MsgBox mgrWarnOnlyOneActiveAssessmentReport, vbOKOnly Or vbInformation, mgrTitle
        End If
    ElseIf Doc.Type = wdTypeDocument Then
        If IsAssessmentReport Then

            ' If IsAssessmentReport is true them we are either opening one of our temporary
            ' html files or the user has saved and is reopening an Assessment Report.
            ' Ignore opening an html file
            isHTMLFile = (StrComp(GetExtensionName(Doc), "html", vbTextCompare) = 0)
            If Not isHTMLFile Then

                ' Determine if we are reopening an assessment report that was closed when producing a Full Report or Summary Report.
                ' If we are do not destroy the document variable that we use to indicate that the document is an assessment report
                ' and do not reset the User Interface, so that the user can continue editing or Submit their assessment report.
                If ADDocVarBooleanValue(mgrDVPreserveUI) Then

                    ' This variable should be deleted imediately since it has served its purpose
                    Doc.Variables(mgrDVPreserveUI).Delete
                    Doc.Save
                Else

                    ' This should be a Word document so destroy the Assessment Report Document Variable so that IsAssessmentReport
                    ' will not report the document as an Assessment report and consequently other code will not try to setup the UI
                    Doc.Variables(mgrDVAssessmentReport) = vbNullString
                    Doc.Saved = True

                    ' Make sure all object used to create an Assessment Report are uninitialised as they should not be used
                    ResetAlmostEverything
                End If
            End If
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' m_wordApp_DocumentOpen

Private Sub m_wordApp_WindowActivate(ByVal Doc As Document, _
                                     ByVal Wn As Window)
    ' Invalidate the MSD Tab Group controls so that they are reset appropriate to the document context.
    ' e.g.: Visible for an Assessment Report hidden for all other documents.
    Manager_InvalidateGroupControls
End Sub ' m_wordApp_WindowActivate
