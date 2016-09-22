Attribute VB_Name = "modLocalData"
'===================================================================================================================================
' Module:       modLocalData
' Purpose:      Contains data used globally by this addin.
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

Public Const rdaTitle                                   As String = "Word RDA AddIn"


' Used for 'patternDatea' parameter replacement
Public Const rdaPDP1                                    As String = "#1"

' Action types
Public Const rdaTypeActionAdd                           As String = "ActionAdd"
Public Const rdaTypeActionInsert                        As String = "ActionInsert"
Public Const rdaTypeActionLink                          As String = "ActionLink"
Public Const rdaTypeActionRename                        As String = "ActionRename"
Public Const rdaTypeActionSetup                         As String = "ActionSetup"


Public Enum rdaDataFormat
    rdaDataFormatDateLong
    rdaDataFormatDateShort
    rdaDataFormatLong
    rdaDataFormatMultiline
    rdaDataFormatRichText
    rdaDataFormatText
    rdaDataFormatTick                                   ' Added for Manifest 1 support, but never used
End Enum

' "buildingBlock" and "buildingBlockN" attribute "where" values
Public Enum rdaWhere
    rdaWhereNone
    rdaWhereAfterLastParagraph
    rdaWhereAtEndOfRange
    rdaWhereReplaceRange
End Enum

Public Enum rdaAction
    rdaActionInsert
    rdaActionAdd
    rdaActionSetup
End Enum

Public Enum rdaValidationResult
    rdaValidationResultNone
    rdaValidationResultSome
End Enum

Public Enum rdaViewMode
    rdaViewModeWrite
    rdaViewModeRead
    rdaViewModePrint
End Enum


Public g_addsWithRefresh        As Collection
Public g_bookmarkCount          As Long
Public g_counters               As Counters
Public g_editableBookmarks      As EditableBookmarks
Public g_instructions           As Instructions

