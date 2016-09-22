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
' History:      30/05/16    1.  Moved to this AddIn as part of the SSAR development.
'===================================================================================================================================
Option Explicit

Private m_actualBookmarks  As VBA.Collection
Private m_patternBookmarks As VBA.Collection

Private Sub Class_Initialize()
    Set m_actualBookmarks = New VBA.Collection
    Set m_patternBookmarks = New VBA.Collection
End Sub

'===================================================================================================================================
' Procedure:    BookmarkAdd
' Purpose:      Adds a bookmark name to the Bookmark collection object.
' Note 1:       The bookmark collection object is kept in document order.
' Note 2:       The bookmark must exist when this code is called or it is not added to the Bookmark collection object.
'
' Date:         30/05/16    Created.
'
' On Entry:     theBookmarkName     The name of the bookmark being added.
'               afterBookmark       The name of a bookmark the passed in bookmark should be added after.
'===================================================================================================================================
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

Friend Function BookmarkExists(ByVal theBookmarkName As String) As Boolean
    Dim dummy As String

    On Error Resume Next
    dummy = m_actualBookmarks(theBookmarkName)

    BookmarkExists = (Err.Number = 0)
    Err.Clear
End Function ' BookmarkExists

'===================================================================================================================================
' Procedure:    BookmarkRemove
' Purpose:      Removes a bookmark from the Bookmark name collection object.
' Date:         30/05/16    Created.
'
' On Entry:     theBookmarkName     The bookmark name to remove from the Bookmark name collection object.
'===================================================================================================================================
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

'===================================================================================================================================
' Procedure:    DeleteInRange
' Purpose:      Deletes all editable bookmarks in the passed in range.
' Note 1:       We don't delete the actualt bookmark, just remove it from the editable bookmarks collection object.
' Note 2:       The actual bookmarks must exist in the document, so call this procedure before deleting the range from the document.
'
' Date:         02/08/16    Created.
'
' On Entry:     targetRange         The range to delete all editable bookmarks from.
'===================================================================================================================================
Friend Sub DeleteInRange(ByVal targetRange As Word.Range)
    Const c_proc As String = "EditableBookmarks.DeleteInRange"

    Dim bookmarkIndex   As Word.Bookmark

    On Error GoTo Do_Error

    ' Loop through all bookmarks in the range and see if there is a corresponding editable bookmark.
    ' If there is an editable bookmark then remove it from the editable bookmarks collection object.
    For Each bookmarkIndex In targetRange.bookmarks
        If BookmarkExists(bookmarkIndex.Name) Then
            BookmarkRemove bookmarkIndex.Name
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' DeleteInRange

Friend Function DumpBookmarks()
    Dim index As Long

    If m_actualBookmarks Is Nothing Then
        Debug.Print "m_actualBookmarks Is Nothing"
    Else
        Debug.Print "Editable bookmark count: " & m_actualBookmarks.Count
        For index = 1 To m_actualBookmarks.Count
            Debug.Print index & "," & m_actualBookmarks(index)
        Next
    End If
End Function ' DumpBookmarks

Friend Function DumpPatterns()
    Dim index As Long

    For index = 1 To m_patternBookmarks.Count
        Debug.Print m_patternBookmarks(index)
    Next
End Function ' DumpPatterns

'===================================================================================================================================
' Procedure:    InsertAfter
' Purpose:      Locates the bookmark immediately before the passed in bookmark name.
' Notes:        This works because:
'               1.  The bookmark collection object is kept in document order.
'               2.  There are no overlapping bookmarks.
'               3.  This is only the subset of editable bookmarks, not all bookmarks.
'
' Date:         30/05/16    Created.
'               02/08/16    Rewrite to optimise search, now only seraches from the end of the specified bookmarks range, rather
'                           than the end of the assessment report.
'
' On Entry:     theBookmarkName     The name of the bookmark, whose predecessor bookmark we require.
' Returns:      The name of the passed in bookmarks immediate predecessor bookmark or a null string if there is no predecessor.
'===================================================================================================================================
'''Private Function InsertAfter(ByVal theBookmarkName As String) As String
'''    Const c_proc As String = "EditableBookmarks.InsertAfter"
'''
'''    Dim index       As Long
'''    Dim newBMRange  As Word.Range
'''    Dim thisBMName  As String
'''    Dim thisBMRange As Word.Range
'''    Dim newStart    As Long
'''
'''    On Error GoTo Do_Error
'''
'''    ' Get the range object of the passed in Bookmark name, then use that to get the bookmarks starting position
'''    Set newBMRange = g_assessmentReport.bookmarks(theBookmarkName).Range
'''    newStart = newBMRange.Start
'''
'''    ' Loop backwards through the assessments reports bookmark collection (as the bookmarks are in document order)
'''    With g_assessmentReport.bookmarks
'''        For index = m_actualBookmarks.Count To 1 Step -1
'''
'''            ' Get the Bookmark name and its Range object
'''            thisBMName = m_actualBookmarks(index)
'''            Set thisBMRange = .Item(thisBMName).Range
'''
'''            ' If the specified bookmark starts before the current bookmarks ends, then we've found its predecessor
'''            If newStart > thisBMRange.End Then
'''                Exit For
'''            End If
'''        Next
'''    End With
'''
'''    ' If no predecessor was found, then we return a null string
'''    InsertAfter = thisBMName
'''
'''Do_Exit:
'''    Exit Function
'''
'''Do_Error:
'''    ErrorReporter c_proc
'''    Resume Do_Exit
'''End Function ' InsertAfter

Private Function InsertAfter(ByVal theBookmarkName As String) As String
    Const c_proc As String = "EditableBookmarks.InsertAfter"

    Dim index       As Long
    Dim newBMRange  As Word.Range
    Dim searchRange As Word.Range
    Dim thisBMName  As String
    Dim thisBMRange As Word.Range
    Dim newStart    As Long

    On Error GoTo Do_Error

    ' Get the range object of the passed in Bookmark name, then use that to get the bookmarks starting position
    Set newBMRange = g_assessmentReport.bookmarks(theBookmarkName).Range
    newStart = newBMRange.Start

    ' The search range is from the start of the assessment report to the end of the specified bookmark range
    Set searchRange = g_assessmentReport.Content
    searchRange.End = newBMRange.End

    ' Loop backwards through the search range
    With searchRange.bookmarks
        For index = .Count To 1 Step -1

            ' Get the Bookmark name
            thisBMName = .Item(index)

            ' See if it is an editable bookmark as they are the only ones we are interested in
            If BookmarkExists(thisBMName) Then
                Set thisBMRange = .Item(index).Range
 
                ' If the specified bookmark starts before the current bookmarks ends, then we've found its predecessor
                If newStart > thisBMRange.End Then
                    Exit For
                End If
            End If
            
            ' Reset the bookmark name so that if there is no match we return a null string
            thisBMName = vbNullString
        Next
    End With

    ' If no predecessor was found, then we return a null string
    InsertAfter = thisBMName

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' InsertAfter

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

Friend Function PatternExists(ByVal thePattern As String) As Boolean
    Dim dummy As String

    On Error Resume Next
    dummy = m_patternBookmarks(thePattern)

    PatternExists = (Err.Number = 0)
    Err.Clear
End Function ' PatternExists

'===================================================================================================================================
' Procedure:    Rename
' Purpose:      Renames a bookmark in the editable bookmarks collection object.
' Date:         02/08/16    Created.
'
' On Entry:     currentEBName       The current name of the editable bookmark.
'               newEBName           The new editable bookmark name.
'===================================================================================================================================
Friend Sub Rename(ByVal currentEBName As String, _
                  ByVal newEBName As String)
    Const c_proc As String = "EditableBookmarks.Rename"

    On Error GoTo Do_Error

    If BookmarkExists(currentEBName) Then

        ' Since we can't actually rename the bookmark we create a new one and then delete the current one
        BookmarkAdd newEBName, currentEBName
        BookmarkRemove currentEBName
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Rename

Public Property Get Bookmark(ByVal index As Variant) As String
    Bookmark = m_actualBookmarks(index)
End Property ' Get Bookmark

Public Property Get bookmarkCount() As Long
    bookmarkCount = m_actualBookmarks.Count
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

Public Property Get patternCount() As Long
    patternCount = m_patternBookmarks.Count
End Property '  Get PatternCount
