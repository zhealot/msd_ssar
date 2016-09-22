Attribute VB_Name = "modStartHere"
'===================================================================================================================================
' Module:       modStartHere
' Purpose:      This is where we start to produce an assessment report.
' Note:         The scope of this module must be public.
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
' History:      1/06/2016       Created.
'===================================================================================================================================
Option Explicit


'=======================================================================================================================
' Procedure:    BuildRDADocument
' Purpose:      Builds the Assessment Report.
' Notes:        The preparatory setup work has already been performed by the time this procedure is called.
'
' On Exit:      g_instructions      Setup.
'               g_counters          Created.
'=======================================================================================================================
Public Sub BuildRDADocument()
    Const c_proc As String = "modStartHere.BuildRDADocument"

    Dim allActions           As VBA.Collection
    Dim doDocumentProtection As DocumentProtection

    On Error GoTo Do_Error

    EventLog "Starting document build: " & Now, c_proc

    ' Initialise the Instructions object
    Set g_instructions = New Instructions
    g_instructions.Initialise g_xmlInstructionData
    Set allActions = g_instructions.Actions.actionList

    ' Dispose of the g_xmlInstructionData DOMDocument object now that we have finished with it
    Set g_xmlInstructionData = Nothing

    ' Initialise the counters collection
    Set g_counters = New Counters

    ' Create the text file we write the xhtml to. Later this file is opened as a Word document and the
    ' data (located by bookmark name) is cut from the html document and pasted into the Word Assessment report.
    SetupHTMLDataSource

    ' Create the dictionary object used for holding the Range objects and Variants (strings) for RichText
    Set g_richTextData = New Scripting.Dictionary

    ' First parse of Actions object to build the html file used as the RichText source
    RichTextActions allActions

    ' Open the xhtml file as a Word document and parse it to set up a Range object for each xhtml entry
    AssignRangesToHTML

    ' Parse out each "action" (which results in each of the actions being carried out)
    BARActions allActions

    ' Update all of the Ref Fields insert by the Link Actions
    UpdateAllRefFields

    ' Add the appropriate Watermark if the manifest version number is three
    If g_rootData.ManifestVersion = 3 Then
        AddWatermark
    End If

    ' Restrict document editing to Read Only so that only those Ranges marked as Editable can be modified by the user
    Set doDocumentProtection = NewDocumentProtection
    doDocumentProtection.EnableProtection

    ' Create the FluentUI Custom Tab
    EventLog "Next step is RDA_RibbonReset", c_proc
    RDA_RibbonReset

    ' Clear the undo buffer so that the user cannot screw up the generated Assessment Report
    g_assessmentReport.UndoClear

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' BuildRDADocument
