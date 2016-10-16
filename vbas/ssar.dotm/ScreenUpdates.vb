VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScreenUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Module:       ScreenUpdates
' Purpose:      Suppresses screen updates.
' Note 1:       Because this is class based, after instantiating an instance of this class you can rely on VBA to destor the class
'               instance when exiting the procedure that instantiated it. Thus screen updates and a screen refresh is implicit.
' Note 2:       This class is designed to stack, so that if code called futher downstream instantiates another instance of this
'               class when the downstream instance exits, it will not restore screen updates or do a screen refresh unless it is the
'               'outermost' instance.
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
' History:      5/08/16     1.  Created.
'===================================================================================================================================
Option Explicit

Private m_initialState  As Boolean

Private Sub Class_Initialize()

    ' This is key, save the original screen update state
    m_initialState = Application.ScreenUpdating

    ' Only suppress screen updating, if screen updating is enabled
    If m_initialState Then
        Application.ScreenUpdating = False
    End If
End Sub ' Class_Initialize

Private Sub Class_Terminate()
    If m_initialState Then
        Application.ScreenUpdating = True
        Application.ScreenRefresh
    End If
End Sub '  Class_Terminate
