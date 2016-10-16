VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' AN = Attribute Name
Private Const mc_ANBookmark     As String = "bookmark"          ' Initial bookmark name
Private Const mc_ANBreak        As String = "break"             ' Breakpoint debug aid,
Private Const mc_ANPattern      As String = "pattern"           ' Used to create a unique bookmark name
Private Const mc_ANPatternData  As String = "patternData"       ' The xpath query for the pattern parameter replacement value
Private Const mc_ANSource       As String = "source"            ' The source bookmark name used to update 'pattern' bookmark contents


Private m_bookmark          As String
Private m_bookmarkPattern   As String
Private m_break             As Boolean
Private m_patternData       As String
Private m_source            As String


Friend Sub Parse(ByRef linkNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "ActionLink.Parse"

    On Error GoTo Do_Error

    ' Make sure there are attributes present
    If linkNode.Attributes.Length > 0 Then

        ParseAttributes linkNode, mc_ANBookmark, m_bookmark, mc_ANPattern, m_bookmarkPattern, mc_ANPatternData, m_patternData, _
                                  mc_ANSource, m_source, mc_ANBreak, m_break
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

Friend Property Get Bookmark() As String
    Bookmark = m_bookmark
End Property ' Get Bookmark

Friend Property Get BookmarkPattern() As String
    BookmarkPattern = m_bookmarkPattern
End Property 'Get BookmarkPattern

Friend Property Get Break() As Boolean
    Break = m_break
End Property ' Break

Friend Property Get PatternData() As String
    PatternData = m_patternData
End Property ' PatternData

Friend Property Get Source() As String
    Source = m_source
End Property ' Get Source
