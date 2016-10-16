VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       IAction
' Purpose:      Interface class for all Actions.
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
' History:      28/06/16    1.  Created.
'===================================================================================================================================
Option Explicit


Public Sub Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                 Optional ByVal addPatternBookmarks As Boolean)
    ' No code allowed in interface classes.
End Sub ' Parse

Public Sub BuildAssessmentReport()
    ' No code allowed in interface classes.
End Sub ' BuildAssessmentReport

Public Sub ConstructRichText()
    ' No code allowed in interface classes.
End Sub ' ConstructRichText

Public Sub HTMLForXMLUpdate()
    ' No code allowed in interface classes.
End Sub ' HTMLForXMLUpdate

'===================================================================================================================================
' Procedure:    UpdateContentControlXML
' Purpose:      Update the assessment report xml using a value stored in a Content Control.
' Note 1:       This is a bit confusing because the Content Controls have their own xml data store. But this is not updating the
'               Content Control data store, this updates the assessment report xml using the Contents Controls value.
' Date:         16/08/16    Created
'===================================================================================================================================
Public Sub UpdateContentControlXML()
    ' No code allowed in interface classes.
End Sub ' UpdateContentControlXML

Public Sub UpdateDateXML()
    ' No code allowed in interface classes.
End Sub ' UpdateDateXML

Public Property Get Break() As Boolean
    ' No code allowed in interface classes.
End Property ' Break

Public Property Get ActionType() As ssarActionType
    ' No code allowed in interface classes.
End Property ' Get ActionType
