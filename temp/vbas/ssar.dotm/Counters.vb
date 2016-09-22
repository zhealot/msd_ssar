VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        Counters
' Purpose:      Maintains a set of counters used for determining the depth (nesting level) of Add, AddDual and Do instructions in
'               the 'instructions.xml' file.
' Note 1:       The counters operate using a stack style mechanism, it is always the counter on the top of the stack that is used
'               by the default properties (but this can be overridden).
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

Private Const mc_size As Long = 10

Private m_counters()    As Long
Private m_depth         As Long
Private m_offset        As Long                     ' Used for Tables that display two data items per row, thus half as many rows required
                                                    ' as there are data items. This provides a virtual secondary set of counters offset from
                                                    ' the main counter by a specified amount (the offset). You apply the offset to all items
                                                    ' in the second data set. The offset is calculated (total data items +1) \ 2.
                                                    ' So in a table displaying 8 data items will display them thus:
                                                    ' ---------------------------------------------
                                                    ' | Item 1.a | Item 1.b | Item 5.a | Item 5.b |
                                                    ' ---------------------------------------------
                                                    ' | Item 2.a | Item 2.b | Item 6.a | Item 6.b |
                                                    ' ---------------------------------------------
                                                    ' | Item 3.a | Item 3.b | Item 7.a | Item 7.b |
                                                    ' ---------------------------------------------
                                                    ' | Item 4.a | Item 4.b | Item 8.a | Item 8.b |
                                                    ' ---------------------------------------------
                                                    ' The offset can be added to the counter to yield the index position of its data counterpart
                                                    ' from the second data set.

Private Sub Class_Initialize()
    ReDim m_counters(1 To mc_size)
End Sub ' Class_Initialize

Friend Sub DecrementDepth()
    m_depth = m_depth - 1
End Sub ' IncrementDepth

Friend Sub DecrementCounter()

    If m_depth <= UBound(m_counters) Then
        m_counters(m_depth) = m_counters(m_depth) - 1
    End If
End Sub ' DecrementCounters

Friend Sub IncrementCounter()

    If m_depth <= UBound(m_counters) Then
        m_counters(m_depth) = m_counters(m_depth) + 1
    End If
End Sub ' IncrementCounters

Friend Sub IncrementDepth()
    m_depth = m_depth + 1

    ' Make space in the counter array if necessary
    If m_depth > UBound(m_counters) Then
        ReDim Preserve m_counters(LBound(m_counters) To UBound(m_counters) + mc_size)
    End If
End Sub ' IncrementDepth

Public Sub SetCounter(ByVal theDepth As Long, _
                      ByVal theValue As Long)
    Const c_proc As String = "Counters.SetCounter"

    On Error GoTo Do_Error

    If theDepth >= LBound(m_counters) And theDepth <= UBound(m_counters) Then
        m_counters(theDepth) = theValue
    Else
        Err.Raise mgrErrNoInvalidProcedureCall, c_proc
    End If

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' SetCounter

Friend Sub ResetCounter(Optional ByVal theDepth As Long = -1)

    If theDepth < 0 Then
        theDepth = m_depth
    End If

    If theDepth <= UBound(m_counters) And theDepth >= 1 Then
        m_counters(m_depth) = 0
    End If
End Sub ' ResetCounter

Friend Property Get Counter() As Long
    Counter = m_counters(m_depth)
End Property
Friend Property Let Counter(ByVal newValue As Long)
    m_counters(m_depth) = newValue
End Property ' Counter

'===================================================================================================================================
' Procedure:    UpdatePredicates
' Purpose:      Replaces predicate placeholder character sequence with a counter value.
'
' Note 1:       Though aimed at xpath predicate expressions this code is also used to do parameter replacement in bookmark names.
'               This allows the bookmark name to reflect the position of the xml item used to fill the bookmark, in the xml items
'               nodeset.
' Note 2:       There are two parameter placeholder character sets. Those that start with a '%' and those that start with a '!'.
'               If the placeholder starts with a '%' then the standard counter values will be used.
'               If the placeholder starts with a '!' then the offset value will be added to the standard counter values before it
'               is used.
'
' Date:         30/06/16    Updated to add functionality for the replacement of offset ('!') placeholder character sequences.
'
' On Entry:     theQuery            An xpath expression or bookmark name (with predicate placeholder cgaracters).
' Returns:      The query string with all placeholder character sequences replaced with the appropriate counter values.
'===================================================================================================================================
Friend Function UpdatePredicates(ByVal theQuery As String) As String
    Const c_proc As String = "Counters.UpdatePredicates"

    Dim errorText   As String
    Dim index       As Long
    Dim parameter1  As String
    Dim parameter2  As String
    Dim theCounter  As Long

    On Error GoTo Do_Error

    For index = 1 To m_depth
        theCounter = m_counters(index)

        Select Case index
        Case 1
            parameter1 = mgrP1
            parameter2 = ssarP1
        Case 2
            parameter1 = mgrP2
            parameter2 = ssarP2
        Case 3
            parameter1 = mgrP3
            parameter2 = ssarP3
        Case 4
            parameter1 = mgrP4
            parameter2 = ssarP4
        Case Else
            errorText = Replace$(mgrErrTextCounterDepthExceedsMaximum, mgrP1, 4)
            errorText = Replace$(errorText, mgrP2, g_configuration.CurrentInstructionsFileFullName)
            Err.Raise mgrErrNoCounterDepthExceedsMaximum, c_proc, errorText
        End Select

        ' Replace both the primary placeholders (those that start with a '%') and the offset placeholders (those that start with a '!')
        theQuery = Replace$(theQuery, parameter1, CStr(theCounter))
        theQuery = Replace$(theQuery, parameter2, CStr(theCounter + m_offset))
    Next

    UpdatePredicates = theQuery

Do_Exit:
    Exit Function

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Function ' UpdatePredicates

Friend Property Get Count(Optional ByVal requiredDepth As Long = -1) As Long

    If requiredDepth < 0 Then

        ' Use the current depth as the index
        Count = m_counters(m_depth)
    Else

        ' Use the supplied depth as the index
        Count = m_counters(requiredDepth)
    End If
End Property ' Get Count

Friend Property Get Depth() As Long
    Depth = m_depth
End Property
Friend Property Let Depth(ByVal newDepth As Long)
    m_depth = newDepth
End Property ' Depth

Friend Property Get Offset() As Long
    Offset = m_offset
End Property
Friend Property Let Offset(ByVal newOffset As Long)
    m_offset = newOffset
End Property
