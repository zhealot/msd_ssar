Attribute VB_Name = "modActionsHelper"
'===================================================================================================================================
' Module:       modActionsHelper
' Purpose:      A helper module for the ActionXXX classes.
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
' History:      1/06/16     1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module


Public Sub AddNodeMissingAttributeError(ByVal nodeName As String, _
                                        ByVal nodeValue As String, _
                                        ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextAddNodeMissingAttribute, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, nodeValue)
    Err.Raise mgrErrNoAddNodeMissingAttribute, moduleProcedure, errorText
End Sub ' AddNodeMissingAttributeError

Public Sub AddNodeMissingAttributeNError(ByVal nodeName As String, _
                                         ByVal nodeValue As String, _
                                         ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextAddNodeMissingAttributeN, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, nodeValue)
    Err.Raise mgrErrNoAddNodeMissingAttributeN, moduleProcedure, errorText
End Sub ' AddNodeMissingAttributeNError

Public Sub BookmarkDoesNotExistError(ByVal bookmarkName As String, _
                                     ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextBookmarkDoesNotExist, mgrP1, bookmarkName)
    Err.Raise mgrErrNoRequestedCollectionMemberDoesNotExist, moduleProcedure, errorText
End Sub ' BookmarkDoesNotExistError

Public Sub BookmarkIsExpectedToContainATableError(ByVal bookmarkName As String, _
                                                  ByVal moduleProcedure As String)
    Dim errorText   As String
 
    errorText = Replace$(mgrErrTextBookmarkIsExpectedToContainATable, mgrP1, bookmarkName)
    Err.Raise mgrErrNoBookmarkIsExpectedToContainATable, moduleProcedure, errorText
End Sub ' BookmarkIsExpectedToContainATableError

Public Sub DuplicateAttributeError(ByVal nodeName As String, _
                                   ByVal attributeName As String, _
                                   ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextDuplicateAttributeName, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    Err.Raise mgrErrNoDuplicateAttributeName, moduleProcedure, errorText
End Sub ' DuplicateAttributeError

Public Sub DuplicateNodeError(ByVal nodeName As String, _
                              ByVal parentNodeName As String, _
                              ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextDuplicateNodeName, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, parentNodeName)
    Err.Raise mgrErrNoDuplicateNodeName, moduleProcedure, errorText
End Sub ' DuplicateNodeError

Public Sub InvalidInsertAfterRowNumberError(ByVal InsertAfter As Long, _
                                            ByVal insertAfterLimit As Long, _
                                            ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextInsertAfterRowNumber, mgrP1, CStr(InsertAfter))
    errorText = Replace$(errorText, mgrP2, CStr(insertAfterLimit))
    Err.Raise mgrErrNoInsertAfterRowNumber, moduleProcedure, errorText
End Sub ' InvalidInsertAfterRowNumberError

Public Sub InvalidActionVerbError(ByVal nodeName As String, _
                                  ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidActionVerb, mgrP1, nodeName)
    Err.Raise mgrErrNoInvalidActionVerb, moduleProcedure, errorText
End Sub ' InvalidActionVerbError

Public Sub InvalidAttributeCombinationError(ByVal nodeName As String, _
                                            ByVal attributeName As String, _
                                            ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidAttributeCombination, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    Err.Raise mgrErrNoInvalidAttributeCombination, moduleProcedure, errorText
End Sub ' InvalidAttributeCombinationError

Public Sub InvalidAttributeNameError(ByVal nodeName As String, _
                                     ByVal attributeName As String, _
                                     ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidAttributeName, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    Err.Raise mgrErrNoInvalidAttributeName, moduleProcedure, errorText
End Sub ' InvalidAttributeNameError

Public Sub InvalidAttributeWhatValueError(ByVal nodeName As String, _
                                          ByVal attributeName As String, _
                                          ByVal attributeValue As String, _
                                          ByVal moduleProcedure As String)
    Const c_permissibleAttributes   As String = "'All' or 'Table'."

    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidAttributeValueExtended, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    errorText = Replace$(errorText, mgrP3, c_permissibleAttributes)
    errorText = Replace$(errorText, mgrP4, attributeValue)
    Err.Raise mgrErrNoInvalidAttributeValue, moduleProcedure, errorText
End Sub ' InvalidAttributeWhatValueError

Public Sub InvalidAttributeWhereValueError(ByVal nodeName As String, _
                                           ByVal attributeName As String, _
                                           ByVal attributeValue As String, _
                                           ByVal moduleProcedure As String)
    Const c_permissibleAttributes   As String = "'AfterLastParagraph' or 'AtEndOfRange'."

    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidAttributeValueExtended, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    errorText = Replace$(errorText, mgrP3, c_permissibleAttributes)
    errorText = Replace$(errorText, mgrP4, attributeValue)
    Err.Raise mgrErrNoInvalidAttributeValue, moduleProcedure, errorText
End Sub ' InvalidAttributeWhereValueError

Public Sub InvalidAttributeValueError(ByVal nodeName As String, _
                                      ByVal attributeName As String, _
                                      ByVal attributeValue As String, _
                                      ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidAttributeValue, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    errorText = Replace$(errorText, mgrP3, attributeValue)
    Err.Raise mgrErrNoInvalidAttributeValue, moduleProcedure, errorText
End Sub ' InvalidAttributeValueError

Public Sub InvalidNodeNameError(ByVal nodeName As String, _
                                ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextInvalidNodeName, mgrP1, nodeName)
    Err.Raise mgrErrNoInvalidNodeName, moduleProcedure, errorText
End Sub ' InvalidNodeNameError

Public Sub KeyDoesNotExistInDictionaryError(ByVal keyName As String, _
                                            ByVal dictionaryName As String, _
                                            ByVal moduleProcedure As String)
    Dim errorText   As String
                                            
    errorText = Replace$(mgrErrTextKeyDoesNotExistInDictionary, mgrP1, keyName)
    errorText = Replace$(errorText, mgrP2, dictionaryName)
    Err.Raise mgrErrNoKeyDoesNotExistInDictionary, moduleProcedure, errorText
End Sub ' KeyDoesNotExistInDictionaryError

Public Sub MissingAttributeError(ByVal nodeName As String, _
                                 ByVal attributeName As String, _
                                 ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextMissingAttribute, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    Err.Raise mgrErrNoMissingAttribute, moduleProcedure, errorText
End Sub ' MissingAttributeError

Public Sub OnlyOneAttributeAllowedError(ByVal nodeName As String, _
                                        ByVal attributeName As String, _
                                        ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextOnlyOneAttributeAllowed, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    Err.Raise mgrErrNoOnlyOneAttributeAllowed, moduleProcedure, errorText
End Sub ' OnlyOneAttributeAllowedError

Public Sub UnknownFunctionNameError(ByVal nodeName As String, _
                                    ByVal attributeName As String, _
                                    ByVal attributeValue As String, _
                                    ByVal moduleProcedure As String)
    Dim errorText   As String

    errorText = Replace$(mgrErrTextUnknownFunctionNameError, mgrP1, nodeName)
    errorText = Replace$(errorText, mgrP2, attributeName)
    errorText = Replace$(errorText, mgrP3, attributeValue)
    Err.Raise mgrErrNoUnknownFunctionNameError, moduleProcedure, errorText
End Sub ' UnknownFunctionNameError
