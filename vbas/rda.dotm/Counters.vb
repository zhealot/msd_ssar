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
' Purpose:      Maintains a set of counters used for determining the depth (nesting level) of Add instructions in the instruction
'               file.
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

Private m_counters() As Long
Private m_depth      As Long

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

Friend Property Let Counter(ByVal newValue As Long)

    m_counters(m_depth) = newValue
End Property ' Let Counter

Friend Function UpdatePredicates(ByVal theQuery As String) As String
    Const c_proc As String = "Counters.UpdatePredicates"

    Dim errorText   As String
    Dim index       As Long
    Dim parameter   As String
    Dim theCounter  As Long

    On Error GoTo Do_Error

    For index = 1 To m_depth
        theCounter = m_counters(index)

        Select Case index
        Case 1
            parameter = mgrP1

        Case 2
            parameter = mgrP2

        Case 3
            parameter = mgrP3

        Case 4
            parameter = mgrP4

        Case Else
            errorText = Replace$(mgrErrTextCounterDepthExceedsMaximum, mgrP1, 4)
            errorText = Replace$(errorText, mgrP2, g_configuration.CurrentInstructionsFileFullName)
            Err.Raise mgrErrNoCounterDepthExceedsMaximum, c_proc, errorText
        End Select

        theQuery = Replace$(theQuery, parameter, CStr(theCounter))
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
