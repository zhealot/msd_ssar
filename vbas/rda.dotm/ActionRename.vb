VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ActionRename
' Purpose:      Data class that models the 'rename' action node of the 'rda instructions.xml' file.
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
' History:      16/11/15    1.  Created.
'===================================================================================================================================
Option Explicit


' AN = Attribute Name
Private Const mc_ANBreak   As String = "break"
Private Const mc_ANNewName As String = "newName"
Private Const mc_ANOldName As String = "oldName"


Private m_break           As Boolean
Private m_newBookmarkName As String
Private m_oldBookmarkName As String


Friend Sub Parse(ByRef renameNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "ActionRename.Parse"

    On Error GoTo Do_Error

    ' Make sure there are attributes present
    If renameNode.Attributes.Length > 0 Then

        ParseAttributes renameNode, mc_ANNewName, m_newBookmarkName, mc_ANOldName, m_oldBookmarkName, mc_ANBreak, m_break
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

Friend Property Get Break() As Boolean
    Break = m_break
End Property ' Break

Friend Property Get OldBookmarkName() As String
    OldBookmarkName = m_oldBookmarkName
End Property ' OldBookmarkName

Friend Property Get NewBookmarkName() As String
    NewBookmarkName = m_newBookmarkName
End Property ' NewBookmarkName
