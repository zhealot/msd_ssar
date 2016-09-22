VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentProtection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===================================================================================================================================
' Class:        DocumentProtection
' Purpose:      This class handles the protect and unprotection of the Assessment Report.
'               The Assessment is normally locked so that only the ranges set as editable can be edited. But if the user uses the
'               Fluent UI to make permitted changes, we use this class to unprotect and the reprotect the Assessment Report.
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

Private m_protectOnExit As Boolean


Private Sub Class_Initialize()
    m_protectOnExit = IsProtected
End Sub ' Class_Initialize

Private Sub Class_Terminate()
    If m_protectOnExit Then
        EnableProtection
    End If
End Sub ' Class_Terminate

Friend Property Get ProtectOnExit() As Boolean
    ProtectOnExit = m_protectOnExit
End Property ' Get ProtectOnExit

Public Sub EnableProtection()
    If Not IsProtected Then
        RestrictDocumentEditing
    End If
End Sub ' EnableProtection

Public Sub DisableProtection(Optional ByVal protectWhenExiting As Variant)
    Const c_proc As String = "DocumentProtection.DisableProtection"

    Dim Password As String

    On Error GoTo Do_Error

    ' Save what this class should do when this instance terminates
    If Not IsMissing(protectWhenExiting) Then

        If VarType(protectWhenExiting) = vbBoolean Then
            m_protectOnExit = protectWhenExiting
        Else
            Err.Raise mgrErrNoInvalidProcedureCall, c_proc
        End If
    End If

    ' If the Assessment Report is Protected, then we unprotect it. If it is not protected, then there is nothing to do.
    If IsProtected Then

        ' Unprotect the document if directed to do so by the configuration file
        If g_configuration.ProtectDocument Then

            ' Get the password (if the Assessment Report is password protected)
            Password = g_configuration.AssessmentReportPassword

            ' Unprotect the Assessment Report so that we can make changes to it
            g_assessmentReport.Unprotect Password
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DisableProtection

Private Function IsProtected() As Boolean
    IsProtected = (g_assessmentReport.ProtectionType = wdAllowOnlyReading)
End Function ' IsProtected

Private Sub RestrictDocumentEditing()
    Const c_proc As String = "DocumentProtection.RestrictDocumentEditing"

    Dim Password As String

    On Error GoTo Do_Error

    ' Protect the document if directed to do so by the configuration file
    If g_configuration.ProtectDocument Then

        ' Get the password (if the Assessment Report should be password protected)
        Password = g_configuration.AssessmentReportPassword

        ' Set the Assessment Report document protection to Read Only. This leaves just the Ranges
        ' marked as editable (an Editor has been added to the Range) as being editable by the user.
        g_assessmentReport.Protect wdAllowOnlyReading, False, Password, False, False
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' RestrictDocumentEditing
