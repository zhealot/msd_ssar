Attribute VB_Name = "AutoExec"
'===================================================================================================================================
' Module:       AutoExec
' Purpose:      The only purpose of this module is to house the AutoExec procedure (aka Main).
' Note 1:       Word seems to have some issues (thanks microsoft) these days in recognising the AutoExec procedure. So here we've
'               abandoned using a procedure called AutoExec in favour of a module named AutoExec and a procedure named Main.
'               Though in practice Word seems equally flakey using either method.
'
' Note 2:       You cannot control the order in which Word automatically loads addins and a lot of the code is dependent upon the
'               ar manager addin being present. Since we cannot guarantee that it is, try to minimise AutoExec activity.
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
' History:      31/05/16    1.  Created.
'===================================================================================================================================
Option Explicit


Public Sub Main()
    Const c_proc As String = "AutoExec.Main"

    On Error GoTo Do_Error

    ' Store the version number of this addin
    g_addinVersionRDA = ThisDocument.Variables(mgrDVVersion).Value

    Debug.Print "RDA." & c_proc & " [version: " & g_addinVersionRDA & "]"

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Main

