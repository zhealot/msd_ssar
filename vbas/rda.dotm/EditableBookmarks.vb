VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EditableBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        EditableBookmarks
' Purpose:      Tracks Bookmarks with editable areas, so that they can be recreated if necessary.
' Note 1:       These comments apply to Bookmarks with editable areas (i.e.: the Bookmarks Range object has an Editor object
'               associated with it, thus making it editable).
' Note 2:       The reason the Bookmarks need to be recreated is that if the user copies the entire Range of one Bookmark and pastes
'               that into the entire Range of another Bookmark, the Bookmark for the area being pasted into is destroyed.
' Note 3:       The m_actualBookmarks Collection object is maintained in the order the bookmarks occur in the document.
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
' History:      23/12/16    1.  Created.
'===================================================================================================================================
Option Explicit

Private m_actualBookmarks  As VBA.Collection
Private m_patternBookmarks As VBA.Collection


Private Sub Class_Initialize()
    Set m_actualBookmarks = New VBA.Collection
    Set m_patternBookmarks = New VBA.Collection
End Sub

Friend Sub PatternAdd(ByVal newPattern As String)
    Const c_proc As String = "EditableBookmarks.PatternAdd"

    Dim ErrNumber As Long
    Dim errorText As String

    On Error Resume Next
    m_patternBookmarks.Add newPattern, newPattern

    ' Output a custom error message for a duplicate key error
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        On Error GoTo Do_Error

        If ErrNumber = mgrErrNoKeyAlreadyAssociated Then
            errorText = Replace$(mgrErrTextDuplicateBookmarkPattern, mgrP1, newPattern)
            Err.Raise ErrNumber, c_proc, errorText
        Else
            Err.Raise ErrNumber, c_proc
        End If
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' PatternAdd

Friend Function BookmarkExists(ByVal theBookmarkName As String) As Boolean
    Dim dummy As String

    On Error Resume Next
    dummy = m_actualBookmarks(theBookmarkName)

    BookmarkExists = (Err.Number = 0)
    Err.Clear
End Function ' BookmarkExists

Friend Function DumpBookmarks()
    Dim index As Long

    For index = 1 To m_actualBookmarks.Count
        Debug.Print index & "," & m_actualBookmarks(index)
    Next
End Function ' DumpBookmarks

Friend Function DumpPatterns()
    Dim index As Long

    For index = 1 To m_patternBookmarks.Count
        Debug.Print m_patternBookmarks(index)
    Next
End Function ' DumpPatterns

Friend Function PatternExists(ByVal thePattern As String) As Boolean
    Dim dummy As String

    On Error Resume Next
    dummy = m_patternBookmarks(thePattern)

    PatternExists = (Err.Number = 0)
    Err.Clear
End Function ' PatternExists

Friend Sub BookmarkAdd(ByVal theBookmarkName As String, _
                       Optional ByVal afterBookmark As String)
    Const c_proc As String = "EditableBookmarks.BookmarksAdd"

    On Error GoTo Do_Error

    If LenB(theBookmarkName) > 0 Then

        ' Only add the Bookmark if it's not already in the collection.
        ' If it is in the collection assume it is in the correct position.
        If Not BookmarkExists(theBookmarkName) Then
            If LenB(afterBookmark) = 0 Then

                ' See if we can find the name of the Bookmark we should insert after, this keeps the bookmarks in document position order.
                ' If we cant the Bookmark is added to the end, which means the Bookmark is being added to the end of the document.
                afterBookmark = InsertAfter(theBookmarkName)
                If LenB(afterBookmark) = 0 Then
                    m_actualBookmarks.Add theBookmarkName, theBookmarkName
                Else
                    m_actualBookmarks.Add theBookmarkName, theBookmarkName, , afterBookmark
                End If
            Else
                m_actualBookmarks.Add theBookmarkName, theBookmarkName, , afterBookmark
            End If
        End If
    Else
        Err.Raise mgrErrNoInvalidProcedureCall, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BookmarkAdd

Private Function InsertAfter(ByVal theBookmarkName As String) As String
    Const c_proc As String = "EditableBookmarks.InsertAfter"

    Dim index       As Long
    Dim newBMRange  As Word.Range
    Dim thisBMName  As String
    Dim thisBMRange As Word.Range
    Dim newStart    As Long

    On Error GoTo Do_Error

    '
    Set newBMRange = g_assessmentReport.bookmarks(theBookmarkName).Range
    newStart = newBMRange.Start

    With g_assessmentReport.bookmarks
        For index = m_actualBookmarks.Count To 1 Step -1
            thisBMName = m_actualBookmarks(index)
            Set thisBMRange = .Item(thisBMName).Range
 
            If newStart > thisBMRange.End Then
                Exit For
            End If
        Next
    End With

    '
    If LenB(thisBMName) = 0 Then
        InsertAfter = vbNullString
    Else
        InsertAfter = thisBMName
    End If

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' InsertAfter

Friend Sub BookmarkRemove(ByVal theBookmarkName As String)
    Const c_proc As String = "EditableBookmarks.BookmarkRemove"

    On Error GoTo Do_Error

    m_actualBookmarks.Remove theBookmarkName

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BookmarkRemove

Public Property Get Bookmark(ByVal index As Variant) As String
    Bookmark = m_actualBookmarks(index)
End Property ' Get Bookmark

Public Property Get BookmarkCount() As Long
    BookmarkCount = m_actualBookmarks.Count
End Property '  Get BookmarkCount

Public Property Get IsBookmarkEditable(ByVal theBookmarkName As String)
    IsBookmarkEditable = BookmarkExists(theBookmarkName)
End Property ' Get IsBookmarkEditable

Public Property Get IsPatternEditable(ByVal thePattern As String)
    IsPatternEditable = PatternExists(thePattern)
End Property ' Get IsPatternEditable

Public Property Get Pattern(ByVal index As Variant) As String
    Pattern = m_patternBookmarks(index)
End Property ' Get Pattern

Public Property Get PatternCount() As Long
    PatternCount = m_patternBookmarks.Count
End Property '  Get PatternCount
