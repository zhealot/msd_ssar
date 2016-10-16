VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================================
' Class:        ActionSetup
' Purpose:      Data class that models the 'setup' action node-tree of the 'rda instructions.xml' file.
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
' History:      07/11/15    1.  Created.
'===================================================================================================================================
Option Explicit


' AN = Attribute Name
Private Const mc_ANPrimaryBookmark    As String = "primaryBookmark"
Private Const mc_ANSecondaryBookmarks As String = "secondaryBookmarks"
Private Const mc_ANSetupBuildingBlock As String = "setupBuildingBlock"
Private Const mc_ANRenameBookmarks    As String = "renameBookmarks"


Private m_primaryBookmark    As String                  ' The bookmark name of the block we need to insert a Building Block into
Private m_secondaryBookmarks As String                  ' If the primary bookmark is absent, work our way through this list to
                                                        ' find the next possible insertion location
Private m_setupBuildingBlock As String                  ' The Building Block to insert when the primarp bookmark is absent
Private m_renameBookmarks    As String                  ' A comma delimited list of bookmark names that need renaming


Friend Sub Parse(ByRef setupNode As MSXML2.IXMLDOMNode)
    Const c_proc As String = "ActionSetup.Parse"

    On Error GoTo Do_Error

    ' Parse the 'setup' nodes attributes
    ParseAttributes setupNode, mc_ANPrimaryBookmark, m_primaryBookmark, mc_ANSecondaryBookmarks, m_secondaryBookmarks, _
                               mc_ANSetupBuildingBlock, m_setupBuildingBlock, mc_ANRenameBookmarks, m_renameBookmarks

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' Parse

Friend Property Get PrimaryBookmark() As String
    PrimaryBookmark = m_primaryBookmark
End Property ' PrimaryBookmark

Friend Property Get SecondaryBookmarks() As String
    SecondaryBookmarks = m_secondaryBookmarks
End Property ' SecondaryBookmarks

Friend Property Get SetupBuildingBlock() As String
    SetupBuildingBlock = m_setupBuildingBlock
End Property ' SetupBuildingBlock

Friend Property Get RenameBookmarks() As String
    RenameBookmarks = m_renameBookmarks
End Property ' Get RenameBookmarks
