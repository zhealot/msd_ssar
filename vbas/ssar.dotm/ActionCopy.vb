VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ActionCopy
' Purpose:      Data class that models the 'copy' action node of the 'instructions.xml' file.
'
' Note 1:       Copies the text (without any formatting) of the from bookmark to the to bookmark.
'
' Note 2:       The following attributes are available for use with this Action:
'               fromBookmark [Mandatory]
'                   The name of a Bookmark whose range content will be copied.
'               fromBookmark [Mandatory]
'                   The name of a Bookmark whose range content will be copied.
'               pattern [Optional]
'                   The name of a secondary Bookmark created to mark the text being inserted into the 'toBookmark'.
'               coloured [Optional]
'                   When true applies a text and background colours based on the Foreground and Background colour tables. The actual
'                   colours used are determined by the value of the text in the bookmarked range.
'               break [Optional]
'                   When True causes the code to issue a Stop instruction (used only for debugging).
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
' History:      22/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Implements IAction

' AN = Attribute Name
Private Const mc_ANBreak        As String = "break"             ' Break if value is True
Private Const mc_ANColoured     As String = "coloured"          ' Coloured text and background
Private Const mc_ANFromBookmark As String = "fromBookmark"      ' Initial bookmark name
Private Const mc_ANPattern      As String = "pattern"           ' Secondary boomark name for the 'toBookmark' range
Private Const mc_ANToBookmark   As String = "toBookmark"        ' Initial bookmark name


Private m_break             As Boolean                      ' Used to cause the code to execute a Stop instruction
Private m_coloured          As Boolean                      ' When True, the text and background will be coloured based on the selected text
Private m_fromBookmark      As String                       ' The source of the data being copied
Private m_pattern           As String                       ' The name of (pattern for) a secondary Bookmark allocated to the 'toBookmark' range
Private m_toBookmark        As String                       ' The desination of the data being copied


'=======================================================================================================================
' Procedure:    IAction_Parse
' Purpose:      Parses the Copy action instruction.
'
' On Entry:     actionNode          The xml node containing the add instruction to parse.
'               addPatternBookmarks Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Private Sub IAction_Parse(ByRef actionNode As MSXML2.IXMLDOMNode, _
                          Optional ByVal addPatternBookmarks As Boolean)
    Const c_proc As String = "ActionCopy.Parse"

    Dim colouredCount   As Long
    Dim fromCount       As Long
    Dim patternCount    As Long
    Dim theAttribute    As MSXML2.IXMLDOMNode
    Dim toCount         As Long

    On Error GoTo Do_Error

    ' Parse out the 'copy' nodes attributes
    For Each theAttribute In actionNode.Attributes

        Select Case theAttribute.BaseName
        Case mc_ANFromBookmark
            fromCount = fromCount + 1
            m_fromBookmark = theAttribute.Text

        Case mc_ANToBookmark
            toCount = toCount + 1
            m_toBookmark = theAttribute.Text

        Case mc_ANPattern
            patternCount = patternCount + 1
            m_pattern = theAttribute.Text

        Case mc_ANColoured
            colouredCount = colouredCount + 1
            m_coloured = CBool(theAttribute.Text)

        Case mc_ANBreak
            m_break = CBool(theAttribute.Text)

        Case Else
            InvalidAttributeNameError actionNode.nodeName, theAttribute.BaseName, c_proc
        End Select
    Next

    ' Make sure the Mandatory attributes are present
    If fromCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANFromBookmark, c_proc
    ElseIf toCount = 0 Then
        MissingAttributeError actionNode.nodeName, mc_ANToBookmark, c_proc
    End If

    ' Make sure there are no duplicated attributes
    If fromCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANFromBookmark, c_proc
    ElseIf toCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANToBookmark, c_proc
    ElseIf patternCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANPattern, c_proc
    ElseIf colouredCount > 1 Then
        DuplicateAttributeError actionNode.nodeName, mc_ANColoured, c_proc
    End If

    ' Make sure these attributes contain some data
    If LenB(m_fromBookmark) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANFromBookmark, m_fromBookmark, c_proc
    ElseIf LenB(m_toBookmark) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANToBookmark, m_toBookmark, c_proc
    ElseIf patternCount = 1 And LenB(m_pattern) = 0 Then
        InvalidAttributeValueError actionNode.nodeName, mc_ANPattern, m_pattern, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

'===================================================================================================================================
' Procedure:    IAction_BuildAssessmentReport
' Purpose:      Copies text from the 'from bookmark' to the 'to bookmark'.
' Notes:        Copies without any formatting.
' Date:         22/06/16    Created.
'===================================================================================================================================
Private Sub IAction_BuildAssessmentReport()
    Const c_proc As String = "ActionCopy.IAction_BuildAssessmentReport"

    Dim fromRange           As Word.Range
    Dim realFromBookmark    As String
    Dim realToBookmark      As String
    Dim thePattern          As String
    Dim toRange             As Word.Range

    On Error GoTo Do_Error

    ' This statement is to aid debugging of the 'instruction.xml' file
    If m_break Then Stop

    ' Replace any numeric placeholders with their real value
    realFromBookmark = g_counters.UpdatePredicates(m_fromBookmark)
    realToBookmark = g_counters.UpdatePredicates(m_toBookmark)

    ' Make sure both bookmarks exist
    With g_assessmentReport.bookmarks
        If .Exists(realFromBookmark) Then
            If .Exists(realToBookmark) Then

                ' Get the two bookmarks range objects
                Set fromRange = .Item(realFromBookmark).Range
                Set toRange = .Item(realToBookmark).Range

                ' If the from range contains a Content Control, reset the from range to the Content Controls range, because if the
                ' Content Control is a DropDown Content Control, by convertion there is a space before it and after it (because you
                ' cant set a Bookmark to the Content Controls range). So if we dont adjust the from range in this situation the text
                ' returned will not match the colour map keys if m_coloured = True.
                If fromRange.ContentControls.Count > 0 Then

                    ' We make the relatively safe assumption that the bookmarked area only contains one Content Control
                    Set fromRange = fromRange.ContentControls(1).Range
                End If

                ' Copy the text from the source (from) to the target (to)
                toRange.Text = fromRange.Text

                ' Recreate the (to) bookmark as copying the source text to it will destroy it
                .Add realToBookmark, toRange

                ' See if a secondary (pattern) Bookmark should be created
                If LenB(m_pattern) > 0 Then

                    ' Replace any predicate placeholder characters present in the pattern with their real value
                    thePattern = g_counters.UpdatePredicates(m_pattern)

                    ' Now create the specified bookmark
                    .Add thePattern, toRange
                End If

                ' Apply a font and background colour if requested
                If m_coloured Then
                    ApplyColourMap toRange
                End If
            Else
                BookmarkDoesNotExistError realFromBookmark, c_proc
            End If
        Else
            BookmarkDoesNotExistError realFromBookmark, c_proc
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
    IAction_ActionType = ssarActionTypeCopy
End Property ' IAction_ActionType

Private Property Get IAction_Break() As Boolean
    IAction_Break = m_break
End Property ' Get IAction_Break

Friend Property Get Coloured() As Boolean
    Coloured = m_coloured
End Property ' Coloured

Friend Property Get FromBookmark() As String
    FromBookmark = m_fromBookmark
End Property ' Get FromBookmark

Friend Property Get Pattern() As String
    Pattern = m_pattern
End Property ' Get Pattern

Friend Property Get ToBookmark() As String
    ToBookmark = m_toBookmark
End Property ' Get ToBookmark
