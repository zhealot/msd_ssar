Attribute VB_Name = "modCallCodeInAddIns"
'===================================================================================================================================
' Module:       modCallCodeInAddIns
' Purpose:      Calls code in the SSAR and RDA addins.
' Note 1:       This addin cannot directly call code in the RDA and SSAR addins, because Word does not support circular project
'               References so we have to do it using Application.Run.
' Note 2:       Parameter passing does not seem to work as documented, so likewise we have to do that our own way.
' Note 3:       Despite using the Addin name "RDA!" or "SSAR!" prefix, we have seen code in the wrong Addin called! So to avoid this,
'               wholey unambiguous procedure names are used for procedures in an Addin we need to call.
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
' History:      30/05/16    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module

'Public Sub FUI_BShowRDA1(ByVal control As IRibbonControl)
'    CallCodeInSSAR "SSAR_ResetRibbon", "modSSARFluentUI", False
'    CallCodeInRDA23 "RDA23_ResetRibbon", "modRDAFluentUI", True
'End Sub ' FUI_BShowRDA1

'Public Sub FUI_BShowRDA2(ByVal control As IRibbonControl)
'    CallCodeInRDA23 "RDA23_ResetRibbon", "modRDAFluentUI", False
'    CallCodeInSSAR "SSAR_ResetRibbon", "modSSARFluentUI", True
'End Sub ' FUI_BShowRDA2

'===================================================================================================================================
' Procedure:    BuildRDAAssessmentReport
' Purpose:      Creates the RDA (Manifest 2 and 3) assessment report.
' Date:         01/06/2016
'===================================================================================================================================
Public Sub BuildRDAAssessmentReport()
    Const c_proc    As String = "modCallCodeInAddIns.BuildRDAAssessmentReport"

    On Error GoTo Do_Error

    CallCodeInRDA "BuildRDADocument", "modStartHere"

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BuildRDAAssessmentReport

'===================================================================================================================================
' Procedure:    BuildSSARAssessmentReport
' Purpose:      Creates the SSAR assessment report.
' Date:         01/06/2016
'===================================================================================================================================
Public Sub BuildSSARAssessmentReport()
    Const c_proc    As String = "modCallCodeInAddIns.BuildSSARAssessmentReport"

    On Error GoTo Do_Error

    CallCodeInSSAR "BuildSSARDocument", "modStartHere"

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BuildSSARAssessmentReport

'=======================================================================================================================
' Procedure:    CallCodeInRDA
' Purpose:      Calls the specified procedure in the RDA addin.
'
' On Entry:     procedureName       The name of the procedure to call.
'               moduleName          The name of the module that contains procedureName.
'               parameters          A param array containing any parameters to be passed to the called procedure.
'=======================================================================================================================
Public Sub CallCodeInRDA(ByVal procedureName As String, _
                         ByVal moduleName As String, _
                         ParamArray parameters() As Variant)
    Const c_proc        As String = "modCallCodeInAddIns.CallCodeInRDA"
    Const c_rdaRoot     As String = "RDA!"

    Dim procedureToCall As String
    Dim stackItem       As Variant

    On Error GoTo Do_Error

    ' This is the name of the procedure in the RDA AddIn we will call
    procedureToCall = c_rdaRoot & moduleName & "." & procedureName

    ' Create a new stack list
    Set g_rdaCallStack = New Collection

    ' Add each of the passed in parameters to the call stack
    For Each stackItem In parameters
        g_rdaCallStack.Add stackItem
    Next

    Application.Run procedureToCall

Do_Exit:
    Exit Sub

Do_Error:
    If Err.Number = mgrErrNoUnableToRunTheSpecifiedMacro Then
        Err.Description = Replace$(mgrErrTextUnableToRunTheSpecifiedMacro, mgrP1, procedureToCall)
    End If
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CallCodeInRDA
                           
'=======================================================================================================================
' Procedure:    CallCodeInSSAR
' Purpose:      Calls the specified procedure in the SSAR addin.
'
' On Entry:     procedureName       The name of the procedure to call.
'               moduleName          The name of the module that contains procedureName.
'               parameters          A param array containing any parameters to be passed to the called procedure.
'=======================================================================================================================
Public Sub CallCodeInSSAR(ByVal procedureName As String, _
                          ByVal moduleName As String, _
                          ParamArray parameters() As Variant)
    Const c_proc        As String = "modCallCodeInAddIns.CallCodeInSSAR"
    Const c_ssarRoot    As String = "SSAR!"

    Dim procedureToCall As String
    Dim stackItem       As Variant

    On Error GoTo Do_Error

    ' This is the name of the procedure in the RDA23 AddIn we will call
    procedureToCall = c_ssarRoot & moduleName & "." & procedureName

    ' Create a new stack list
    Set g_ssarCallStack = New Collection

    ' Add each of the passed in parameters to the call stack
    For Each stackItem In parameters
        g_ssarCallStack.Add stackItem
    Next

    Application.Run procedureToCall

Do_Exit:
    Exit Sub

Do_Error:
    If Err.Number = mgrErrNoUnableToRunTheSpecifiedMacro Then
        Err.Description = Replace$(mgrErrTextUnableToRunTheSpecifiedMacro, mgrP1, procedureToCall)
    End If
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' CallCodeInSSAR

'===================================================================================================================================
' Procedure:    Manager_FullRefresh
' Purpose:      Causes a full refresh of the rolled-up data in the assessment reports executive summary section.
' Date:         01/06/2016
'===================================================================================================================================
Public Sub Manager_FullRefresh()
    Const c_proc        As String = "modCallCodeInAddIns.Manager_FullRefresh"

    On Error GoTo Do_Error

    ' Different refresh (rebuild of the rolled up executive summary data) requirements for manifest 2/3 and SSAR
    If g_rootData.ManifestVersion <= 3 Then
        EventLog2 "Calling RDAFullRefresh", c_proc
        CallCodeInRDA "RDAFullRefresh", "modRefresh"
    Else
        EventLog2 "Calling SSARFullRefresh", c_proc
        CallCodeInSSAR "SSARFullRefresh", "modRefresh"
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Manager_FullRefresh

'===================================================================================================================================
' Procedure:    Manager_InvalidateGroupControls
' Purpose:      Invalidates the Fluent UI controls in the RDA and SSAR addins.
' Date:         31/05/2016
'===================================================================================================================================
Public Sub Manager_InvalidateGroupControls()
    Const c_proc        As String = "modCallCodeInAddIns.Manager_InvalidateGroupControls"

    On Error GoTo Do_Error

    ' This procedure is always called when the active widow changes, but be aware that the document may not be an assessment report
    If IsAssessmentReport Then

        ' Force both addins to reset their Ribbons as we only want one addin displaying an RDA Tab!
        ' Or neither displaysing an RDA Tab is the ActiveDocument is not an assessment report.
        CallCodeInRDA "RDA_InvalidateGroupControls", "modRDAFluentUI"
        CallCodeInSSAR "SSAR_InvalidateGroupControls", "modSSARFluentUI"
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Manager_InvalidateGroupControls

'===================================================================================================================================
' Procedure:    ResetAlmostEverything
' Purpose:      Resets all necessary variable and causes the UI to reset as well.
' Date:         01/06/2016
'===================================================================================================================================
Public Sub ResetAlmostEverything()
    Const c_proc        As String = "modCallCodeInAddIns.ResetAlmostEverything"

    On Error GoTo Do_Error

    ' g_rootdata will not be Nothing if there is an assessment report
    If g_rootData Is Nothing Then

        ' Reset both UI's as a precautionary measure
        CallCodeInRDA "ResetAlmostEverything", "modRDAFluentUI"
        CallCodeInSSAR "ResetAlmostEverything", "modSSARFluentUI"
    Else

        ' Different resets requirements for manifest 2/3 and SSAR
        If g_rootData.ManifestVersion <= 3 Then
            CallCodeInRDA "ResetAlmostEverything", "modRDAFluentUI"
        Else
            CallCodeInSSAR "ResetAlmostEverything", "modSSARFluentUI"
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' ResetAlmostEverything
