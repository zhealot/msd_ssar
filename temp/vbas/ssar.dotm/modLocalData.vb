Attribute VB_Name = "modLocalData"
'===================================================================================================================================
' Module:       modLocalData
' Purpose:      Contains data used globally by this addin.
' Note:         Option Private Module is used to limit Public variable scope to this addin, so that these declarations are not
'               visible to the other addins.
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
' History:      02/06/16    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module

Public Const ssarAddinName                              As String = "SSAR"
Public Const ssarTitle                                  As String = "Word SSAR AddIn"


' Used for counter 2 (used by AddDual) to insert data into tables where each row displays two sets (not two items) of data
Public Const ssarP1                                     As String = "!1"
Public Const ssarP2                                     As String = "!2"
Public Const ssarP3                                     As String = "!3"
Public Const ssarP4                                     As String = "!4"

' Used for 'patternData' parameter replacement
Public Const ssarPDP1                                   As String = "#1"

' Colour Names used by /instructions/initialise/colourMap
Public Const ssarColourNameBlack                        As String = "Black"
Public Const ssarColourNameBlue                         As String = "Blue"
Public Const ssarColourNameGreen                        As String = "Green"
Public Const ssarColourNameRed                          As String = "Red"
Public Const ssarColourNameWhite                        As String = "White"
Public Const ssarColourNameYellow                       As String = "Yellow"

' Colours used for text and background (composite RGB colours as used by Word)
Public Const ssarColourBlack                            As Long = wdColorBlack
Public Const ssarColourBlue                             As Long = 15773696          ' RGB(  0, 176, 240)
Public Const ssarColourGreen                            As Long = 5287936           ' RGB(  0, 176,  80)
Public Const ssarColourRed                              As Long = 255               ' RGB(255,   0,   0)
Public Const ssarColourWhite                            As Long = wdColorWhite
Public Const ssarColourYellow                           As Long = 65535             ' RGB(255, 255,   0)

' Action 'add' and Action 'addDual' permissible 'where' attribute values
Public Const ssarWhereTextAfterLastParagraph            As String = "AfterLastParagraph"
Public Const ssarWhereTextAtEndOfRange                  As String = "AtEndOfRange"
Public Const ssarWhereTextReplaceRange                  As String = "ReplaceRange"

Public Enum ssarDataFormat
    ssarDataFormatDateLong
    ssarDataFormatDateShort
    ssarDataFormatDateShortYear
    ssarDataFormatLong
    ssarDataFormatMultiline
    ssarDataFormatRichText
    ssarDataFormatText
End Enum ' ssarDataFormat

' ActionDelete' 'what' attribute values
Public Enum ssarWhatType
    ssarWhatTypeAll = 1                                                 ' All is any tables and the bookmarked range
    ssarWhatTypeTable                                                   ' Any tables in the bookmarked range
    ssarWhatTypePMBefore                                                ' The paragraphmark immediately before the start of the bookmarked range
End Enum ' ssarWhatType

' ActionAdd and ActionAddDual 'buildingBlock' and 'buildingBlockN' translated attribute 'where' values
Public Enum ssarWhereType
    ssarWhereTypeNone
    ssarWhereTypeAfterLastParagraph
    ssarWhereTypeAtEndOfRange
    ssarWhereTypeReplaceRange
End Enum ' ssarWhereType

' ActionIf translated 'operator' attribute values
Public Enum ssarOperatorType
    ssarOperatorTypeNone
    ssarOperatorTypeEQ
    ssarOperatorTypeNE
    ssarOperatorTypeLT
    ssarOperatorTypeGT
    ssarOperatorTypeLE
    ssarOperatorTypeGE
End Enum ' ssarOperatorType

' Used to indicate the actual IAction procedure to call
Public Enum ssarIActionMethod
    ssarIActionMethodBuildAssessmentReport = 1
    ssarIActionMethodRichText = 2
    ssarIActionMethodUpdateContentControlXML = 3
    ssarIActionMethodUpdateDateXML = 4
    ssarIActionMethodHTMLForXMLUpdate = 5
End Enum ' ssarIActionMethod

' The Action Type returned by each Action Class
Public Enum ssarActionType
    ssarActionTypeAdd = 1
    ssarActionTypeAddDual
    ssarActionTypeAddRow
    ssarActionTypeColourMap
    ssarActionTypeCopy
    ssarActionTypeCounter
    ssarActionTypeDelete
    ssarActionTypeDeleteRow
    ssarActionTypeDo
    ssarActionTypeDropDown
    ssarActionTypeIf
    ssarActionTypeInsert
    ssarActionTypeLink
    ssarActionTypeRename
End Enum ' ssarActionType


Public g_addsWithRefresh        As VBA.Collection
Public g_actionCounters         As Scripting.Dictionary
Public g_bookmarkCount          As Long
Public g_counters               As Counters
Public g_dateValidationError    As Boolean
Public g_editableBookmarks      As EditableBookmarks
Public g_instructions           As Instructions
Public g_ccXMLDataStore         As CCXMLDataStore
Public g_colourMapBackground    As Scripting.Dictionary         ' Background colour dictionary
Public g_colourMapForeground    As Scripting.Dictionary         ' Foreground colour dictionary
Public g_xhtmlWordDoc           As Word.Document                ' Used to generate html, which in turn is parsed into xhtml
Public g_standardsCount         As Long
