VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ActionInsert
' Purpose:      Data class that models the 'insert' action node-tree of the 'rda instructions.xml' file.
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
' History:      05/11/15    1.  Created.
'===================================================================================================================================
Option Explicit

' AN = Attribute Name
Private Const mc_ANBookmark     As String = "bookmark"
Private Const mc_ANBreak        As String = "break"
Private Const mc_ANData         As String = "data"
Private Const mc_ANDataFormat   As String = "dataFormat"
Private Const mc_ANDefaultText  As String = "defaultText"
Private Const mc_ANDeleteIfNull As String = "deleteIfNull"
Private Const mc_ANEditable     As String = "editable"
Private Const mc_ANPattern      As String = "pattern"
Private Const mc_ANPatternData  As String = "patternData"       ' Used for the xpath query


' DFV = DataFormat node Values
Private Const mc_DFVDateLong    As String = "DateLong"
Private Const mc_DFVDateShort   As String = "DateShort"
Private Const mc_DFVLong        As String = "Long"
Private Const mc_DFVMultiline   As String = "Multiline"
Private Const mc_DFVRichText    As String = "RichText"
Private Const mc_DFVText        As String = "Text"
Private Const mc_DFVTick        As String = "Tick"


Private m_bookmark          As String
Private m_bookmarkPattern   As String
Private m_break             As Boolean          ' Used to cause the code to execute a Stop instruction
Private m_defaultText       As String
Private m_deleteIfNull      As String           ' Delete this bookmark if the node/nodeList returned by the xpath query contains no data
Private m_editable          As Boolean
Private m_dataFormat        As rdaDataFormat
Private m_dataSource        As String           ' The xpath query that supplies the data
Private m_patternData       As String           ' The xpath query for the pattern parameter replacement value


'=======================================================================================================================
' Procedure:    Parse
' Purpose:      .
' Notes:        .
'
' On Entry:     insertNode          .
'               addPattenBookmarks  Indicate whether pattern bookmarks should be added to g_editableBookmarks.
'=======================================================================================================================
Friend Sub Parse(ByRef insertNode As MSXML2.IXMLDOMNode, _
                 Optional ByVal addPattenBookmarks As Boolean)
    Const c_proc As String = "ActionInsert.Parse"

    Dim DataFormat   As String
    Dim errorText    As String
    Dim theAttribute As MSXML2.IXMLDOMNode

    On Error GoTo Do_Error

    ' Default values, these values prevail when there is no xml data present for a particular attribute (the xml data is optional).
    ' All strings start out with their inherent value of vbNullString.
    m_editable = False
    m_dataFormat = rdaDataFormatText
    m_break = False

    ' Parse all attributes (each entry is a pair of values - the Attribute Name and Variable, to contain the Attribute value)
    ParseAttributes insertNode, mc_ANBookmark, m_bookmark, mc_ANData, m_dataSource, mc_ANEditable, m_editable, mc_ANDataFormat, DataFormat, _
                                mc_ANDeleteIfNull, m_deleteIfNull, mc_ANPattern, m_bookmarkPattern, mc_ANPatternData, m_patternData, _
                                mc_ANDefaultText, m_defaultText, mc_ANBreak, m_break

    ' Parse out the 'dataSource' nodes value
    Select Case DataFormat
    Case mc_DFVText
        m_dataFormat = rdaDataFormatText

    Case mc_DFVRichText
        m_dataFormat = rdaDataFormatRichText

    Case mc_DFVMultiline
        m_dataFormat = rdaDataFormatMultiline

    Case mc_DFVDateLong
        m_dataFormat = rdaDataFormatDateLong

    Case mc_DFVDateShort
        m_dataFormat = rdaDataFormatDateShort

    Case mc_DFVLong
        m_dataFormat = rdaDataFormatLong

    Case mc_DFVTick
        m_dataFormat = rdaDataFormatTick

    Case Else
        errorText = Replace$(mgrErrTextInvalidDataFormatNodeValue, mgrP1, theAttribute.BaseName)
        Err.Raise mgrErrNoInvalidDataFormatNodeValue, c_proc, errorText

    End Select

    ' If the bookmark is editable add the pattern to the editable patterns list
    If m_editable And addPattenBookmarks Then
        If LenB(m_bookmarkPattern) > 0 Then
            g_editableBookmarks.PatternAdd m_bookmarkPattern
        End If
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

Friend Property Get DefaultText() As String
    DefaultText = m_defaultText
End Property ' Get DefaultText

Friend Property Get DeleteIfNull() As String
    DeleteIfNull = m_deleteIfNull
End Property ' Get DeleteIfNull

Friend Property Get DataFormat() As rdaDataFormat
    DataFormat = m_dataFormat
End Property ' Get DataFormat

Friend Property Get DataSource() As String
    DataSource = m_dataSource
End Property ' Get DataSource

Friend Property Get Editable() As Boolean
    Editable = m_editable
End Property ' Get Editable

Friend Property Get HasDefaultText() As Boolean
    HasDefaultText = (Len(m_defaultText) > 0)
End Property ' HasDefaultText

Friend Property Get PatternData() As String
    PatternData = m_patternData
End Property ' PatternData
