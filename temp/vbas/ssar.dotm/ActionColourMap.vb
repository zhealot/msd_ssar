VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionColourMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionColourMap
' Purpose:      Data class that models the 'colourMap' action node of the 'instructions.xml' file.
'
' Note 1:       Applies a contextural <colourMapping> to the specified bookmark using the bookmarks value for the colour lookups.
'
' Note 2:       The <initialise/colourMapping> nodes specify the text values and their associated foreground and background colours.
'               It is these colours that are applied to the bookmarked Range by the colorMap action using the value of the
'               bookmarked Range.
'
' Note 3:       The following attributes are available for use with this Action:
'               bookmark [Mandatory]
'                   The name of a Bookmark whose range content is  matched against the Foreground and Background colour tables and
'                   the resulting lookup value used to set the text colour and background colour.
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
' History:      30/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction


' AN = Attribute Name
Private Const mc_ANBookmark     As String = "bookmark"          ' The bookmark whose range should be coloured mapped based on its text content
Private Const mc_ANBreak        As String = "break"             ' Break if value is True


Private m_bookmark          As String                       ' The name of the bookmark whose range is to be colour mapped
Private m_break             As Boolean                      ' Used to cause the code to execute a Stop instruction

'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the ColourMap action instruction.
'
' On Entry:     actionNode          The xml node containing the colourMap instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionColourMap.IAction_Parse"

    Dim bookmarkCount   As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Parse out the 'colourMap' nodes attributes
    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANBookmark
            bookmarkCount = bookmarkCount + 1
            m_bookmark = theAttribute.Text

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure the Mandatory attributes are present
    If bookmarkCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If bookmarkCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANBookmark, c_proc
    End If

    ' Make sure these attributes contain some data
    If LenB(m_bookmark) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANBookmark, m_bookmark, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_Parse

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Copies text from the 'from bookmark' to the 'to bookmark'.
' Notes:        Copies without any formatting.
' Date:         22/06/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionColourMap.IAction_BuildAssessmentReport"

    Dim bookmarkName    As String
    Dim bookmarkRange   As Word.Range

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Make sure the bookmark exist
    With g_assessmentReport.bookmarks

        ' Replace any predicate placeholder character sequences in the bookmark name with their real value
        bookmarkName = g_counters.UpdatePredicates(m_bookmark)
        If .Exists(bookmarkName) Then

            ' Get the Bookmarks range so we can apply the colour map to it
            Set bookmarkRange = .Item(bookmarkName).Range

            ' Apply a font and background colour if requested
            ApplyColourMap bookmarkRange
        End If
    End With

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' IAction_BuildAssessmentReport

Private Sub IAction_ConstructRichText()
    ' Definition required - But nothing to do
End Sub ' IAction_ConstructRichText

Private Sub IAction_HTMLForXMLUpdate()
    ' Definition required - But nothing to do
End Sub '  Sub IAction_HTMLForXMLUpdate

Private Sub IAction_UpdateContentControlXML()
    ' Definition required - But nothing to do
End Sub ' IAction_UpdateContentControlXML

Private Sub IAction_UpdateDateXML()
    ' Definition required - But nothing to do
End Sub ' Sub IAction_UpdateDateXML

Private Property Get IAction_ActionType() As ssarActionType
    IAction_ActionType = ssarActionTypeColourMap
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break
