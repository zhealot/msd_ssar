Attribute VB_Name = "modDevelopmentAids"
'===================================================================================================================================
' Module:       modDevelopmentAids
' Purpose:      Contains any code and data that does not need to be part of the mainline code and assists in debugging.
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
' History:      1/07/16    1.  Created.
'===================================================================================================================================
Option Explicit
Option Private Module


Private Const mc_ftCriticalAction                       As String = "Critical Action"
Private Const mc_ftRecommendation                       As String = "Recommendation"
Private Const mc_ftRequiredAction                       As String = "Required Action"
Private Const mc_ftStrength                             As String = "Strength"

' Exceptions Table cells
Private Const mc_cellCriteria                           As Long = 1         ' The cell (row relative) that contains the criteria number
Private Const mc_cellFindingsType                       As Long = 3         ' The cell (row relative) that contains the findings type


' FindingsType Type used to determine which FindingType are present dor a specified Criteria number.
' Because these values are AND'ed/OR'ed they must be binary
Private Enum ftType
    ftTypeNone = 0                          ' Only used to reset flag variables
    ftTypeCriticalAction = 1
    ftTypeRequiredAction = 2
    ftTypeStrength = 4
    ftTypeRecommendation = 8
End Enum ' ftType


Private m_selectedDropDownIndex                         As Long


'===================================================================================================================================
' Procedure:    LoadTestInstructionFile
' Purpose:      The whole point of this method is to load a mini instruction file containing just one or two actions you want to
'               execute against an assessment report produced by the mainline code. In this way you can try just one or two Actions
'               to see if you get the desired results.
'
' Note 1:       Calling the Instructions objects Initialise method destroys g_editableBookmarks.
' Note 2:       The g_counters object needs to exist for virtually all Actions to work.
'
' Date:         1/07/16     Created.
'===================================================================================================================================
Private Sub LoadTestInstructionFile()
    Const c_proc As String = "modDevelopmentAids.LoadTestInstructionFile"

    Dim testActions         As Actions
    Dim testInstructions    As Instructions
    Dim undoer              As Word.UndoRecord
    Dim xmlInstructionData  As MSXML2.DOMDocument60

    On Error GoTo Do_Error

    ' Make sure there is an assessment report to execute the Actions against
    If g_assessmentReport Is Nothing Then
        MsgBox "No current assessment report document object"
        Exit Sub
    End If
    
    ' Load the test instruction file
    Set xmlInstructionData = New MSXML2.DOMDocument60
    With xmlInstructionData
        .async = False
        .validateOnParse = False
        .resolveExternals = False
    End With

    ' Load the test instruction xml file
    xmlInstructionData.Load g_configuration.InstructionsFilePath & "test instructions.xml"

    ' Initialise the Instructions object
    Set testInstructions = New Instructions
    testInstructions.Initialise xmlInstructionData
    Set testActions = testInstructions.Actions

    ' Make sure the current undo list is empty, so the only thing to undo after this procedure
    ' terminates is one composite undo of everything that the test instruction file Actions did
    g_assessmentReport.UndoClear

    ' Setup a custom undo record
    Set undoer = Application.UndoRecord
    undoer.StartCustomRecord "Test Instructions Actions"

    ' Execute the main instuction block
    testActions.BuildAssessmentReport

    ' Terminate the custom undo
    undoer.EndCustomRecord

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' LoadTestInstructionFile

Private Sub SimpleRefreshTest()
    Const c_proc As String = "modDevelopmentAids.SimpleRefreshTest"

    Dim doDocumentProtection    As DocumentProtection

    On Error GoTo Do_Error

    ' Disable document protection
    Set doDocumentProtection = NewDocumentProtection
    doDocumentProtection.DisableProtection

    ' Do the refresh
    g_instructions.Refresh.BuildAssessmentReport

    ' Restore document protection
    doDocumentProtection.EnableProtection

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' SimpleRefreshTest

Public Sub ERT()
    Dim insertAfterRow      As Long
    Dim renameFirstRowIndex As Long
    Dim renameLastRowIndex  As Long

    Set g_assessmentReport = ActiveDocument
    m_selectedDropDownIndex = 4                 ' The standard to return information about
    GetExceptionTableInfo 4, mc_ftStrength, insertAfterRow, renameFirstRowIndex, renameLastRowIndex
End Sub '

Private Sub GetExceptionTableInfo(ByVal newCriteria As Long, _
                                  ByVal newFindingType As String, _
                                  ByRef insertAfterRow As Long, _
                                  ByRef renameFirstRowIndex As Long, _
                                  ByRef renameLastRowIndex As Long)
    Const c_proc                As String = "modSSARFluentUI.GetExceptionTableInfo"
    Const c_cellCriteria        As Long = 1         ' The cell (row relative) that contains the criteria number
    Const c_cellFindingsType    As Long = 3         ' The cell (row relative) that contains the findings type

    Dim cellRange               As Word.Range
    Dim currentCriteriaNumber   As Long
    Dim currentFindingsType     As String
    Dim exceptionsRow           As Word.Row
    Dim exceptionsTable         As Word.Table
    Dim rowCount                As Long
    Dim rowIndex                As Long
    Dim theTable                As Word.Table

    On Error GoTo Do_Error

    ' Get a refererence to the Exceptions table that we need to return information about
    Set theTable = GetExceptionsTable(m_selectedDropDownIndex)

    rowCount = theTable.Rows.Count
    For rowIndex = 1 To rowCount
        Set exceptionsRow = theTable.Rows(rowIndex)

        ' Ignore the first row as it is the table Header
        If Not exceptionsRow.IsFirst Then

            ' Criteria number of the current row
            Set cellRange = exceptionsRow.Cells(c_cellCriteria).Range
            cellRange.End = cellRange.End - 1
            currentCriteriaNumber = CLng(cellRange.Text)

            ' If the new Criteria number is less than the current rows criteria number, keep going
            If newCriteria < currentCriteriaNumber Then
                If insertAfterRow = 0 Then
                    insertAfterRow = exceptionsRow.index - 1
                End If
                Exit For
            ElseIf newCriteria = currentCriteriaNumber Then

                ' Extract the current Findings Type
                Set cellRange = exceptionsRow.Cells(c_cellFindingsType).Range
                cellRange.End = cellRange.End - 1
                currentFindingsType = cellRange.Text

                ' The new criteria number and the current rows criteria number match, so now check the finding type
                If newFindingType = mc_ftRecommendation Then

                    ' Recommendations are the lowest sort order item, so a new Recommendation must be after everything else
                    insertAfterRow = exceptionsRow.index
                Else
                    If currentFindingsType = mc_ftCriticalAction Then
                        insertAfterRow = exceptionsRow.index
                    ElseIf currentFindingsType = mc_ftRequiredAction Then
                        insertAfterRow = exceptionsRow.index
                    Else

                        ' We are inserting a Strength and the last item must be a Recommendation, so we insert before the Recommendation row
                        insertAfterRow = exceptionsRow.index - 1
                    End If
                End If
            End If
        End If
    Next

    ' The new Exceptions table row must be inserted after the last row in the table
    If insertAfterRow = 0 Then
        insertAfterRow = rowCount
    End If

    ' Setup the remaining information that we need to return
    If insertAfterRow < rowCount Then
        renameFirstRowIndex = insertAfterRow + 1
        renameLastRowIndex = rowCount
    Else
        renameFirstRowIndex = 0
        renameLastRowIndex = 0
    End If

    MsgBox "Insert after row: " & insertAfterRow & vbCr & _
           "First row index:  " & renameFirstRowIndex & vbCr & _
           "Last row index:   " & renameLastRowIndex

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub '  GetExceptionTableInfo

Public Sub VerifyAllExceptionsTables()
    Const c_proc As String = "modDevelopmentAids.VerifyAllExceptionsTables"

    Dim cellRange                   As Word.Range
    Dim cellText                    As String
    Dim currentCriteriasFindings    As ftType
    Dim currentCriteriaNumber       As Long
    Dim currentFindings             As Long
    Dim currentFindingsType         As String
    Dim errorFlag                   As Boolean
    Dim exceptionsRow               As Word.Row
    Dim previousCriteriaNumber      As Long
    Dim previousFindingsType        As String
    Dim rowCount                    As Long
    Dim rowIndex                    As Long
    Dim tableIndex                  As Long
    Dim theTable                    As Word.Table

    On Error GoTo Do_Error

    EventLog "Verifying all Exceptions Tables", c_proc

    ' Iterate all Exceptions Tables in the assessment report
    For tableIndex = 1 To g_standardsCount

        ' Get a refererence to the Exceptions table that we need to return information about
        Set theTable = GetExceptionsTable(tableIndex)

        ' Make sure there is an Exceptions Table for this standard
        If Not theTable Is Nothing Then

            ' Reset this since we are now using a different table
            currentCriteriasFindings = ftTypeNone
            previousCriteriaNumber = 0

            ' Get the table row count which includes the header row
            rowCount = theTable.Rows.Count
            For rowIndex = 1 To rowCount
                Set exceptionsRow = theTable.Rows(rowIndex)

                ' Always clear the flag as we do not want the previous loop iterations results!
                errorFlag = False

                ' Ignore the first row as it is the table Header
                If Not exceptionsRow.IsFirst Then

                    ' Criteria number of the current row
                    Set cellRange = exceptionsRow.Cells(mc_cellCriteria).Range
                    cellRange.End = cellRange.End - 1

                    ' Make sure the cell contains something
                    cellText = Trim$(cellRange.Text)
                    If LenB(cellText) = 0 Then
                        MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " has an invalid Criteria"
                        errorFlag = True
                    Else
                        If IsNumeric(cellText) Then
                            currentCriteriaNumber = CLng(cellText)
                            If currentCriteriaNumber <= 0 Then
                                MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " Criteria number is invalid"
                                errorFlag = True
                            End If
                        Else
                            MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " Criteria is not numeric"
                            errorFlag = True
                        End If
                    End If

                    ' Findings Type of the current row
                    Set cellRange = exceptionsRow.Cells(mc_cellFindingsType).Range
                    cellRange.End = cellRange.End - 1
                    currentFindingsType = cellRange.Text

                    ' Validate the Findings Type
                    Select Case currentFindingsType
                    Case mc_ftCriticalAction
                        currentFindings = ftTypeCriticalAction
                    Case mc_ftRequiredAction
                        currentFindings = ftTypeRequiredAction
                    Case mc_ftStrength
                        currentFindings = ftTypeStrength
                    Case mc_ftRecommendation
                        currentFindings = ftTypeRecommendation

                    Case Else
                        MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " Finding Type is invalid"
                        errorFlag = True
                    End Select

                    ' Check that the Criteria are in ascending order
                    If currentCriteriaNumber < previousCriteriaNumber Then
                        MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " Criteria is out of order"
                        errorFlag = True
                    ElseIf currentCriteriaNumber = previousCriteriaNumber Then

                        ' The current Criteria number is the same as the previous Criteria number,
                        ' so check that the Finding Type is unique for the current Criteria number
                        If Not errorFlag Then

                            ' Check for a duplicate Findings Type
                            If CBool(currentFindings And currentCriteriasFindings) Then
                                MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " Duplicacte Finding Type"
                                errorFlag = True
                            Else

                                ' Check that the FindingType is in the correct order
                                Select Case currentFindings
                                Case ftTypeCriticalAction
                                    errorFlag = CBool(currentCriteriasFindings And (ftTypeRequiredAction Or ftTypeStrength Or ftTypeRecommendation))
                                Case ftTypeRequiredAction
                                    errorFlag = CBool(currentCriteriasFindings And (ftTypeStrength Or ftTypeRecommendation))
                                Case ftTypeStrength
                                    errorFlag = CBool(currentCriteriasFindings And ftTypeRecommendation)
                                End Select

                                If errorFlag Then
                                     MsgBox "Exceptions Table: " & tableIndex & ", Row: " & rowIndex & " FindingType is out of order"
                                End If

                                ' Update for next loop iteration
                                previousCriteriaNumber = currentCriteriaNumber
                            End If
                        End If
                    Else

                        ' Update for next loop iteration
                        previousCriteriaNumber = currentCriteriaNumber

                        ' Reset this since the current Criteria number is great than the previous Criteria number
                        currentCriteriasFindings = currentFindings
                    End If
                End If
            Next
        End If
    Next

Do_Exit:
    Exit Sub

Do_Error:
    ErrorReporter c_proc
    Resume Do_Exit
End Sub ' VerifyAllExceptionsTables

Private Sub AddXMLNodesFromString()
    Dim docPart         As MSXML2.DOMDocument60
    Dim newNode         As MSXML2.IXMLDOMNode
    Dim nodeToAppendTo  As MSXML2.IXMLDOMNode
    Dim theNodes        As MSXML2.IXMLDOMNodeList
    Dim theXML          As String

    theXML = "<findings><sectionNumber>%1</sectionNumber><narrative>%2</narrative><findingType>%3</findingType></findings>"
    Set docPart = New MSXML2.DOMDocument60
    docPart.LoadXML theXML
    docPart.Validate

    Set nodeToAppendTo = g_xmlDocument.SelectSingleNode("/Assessment/report/standard[1]")
    Set newNode = nodeToAppendTo.appendChild(docPart.FirstChild)

    Set theNodes = g_xmlDocument.SelectNodes("/Assessment/report/standard[1]/findings")
    Debug.Print theNodes.Length
End Sub ' AddXMLNodesFromString

Public Sub DumpFindingsXML(ByVal standardNumber As Long)
    Dim findingsNode    As MSXML2.IXMLDOMNode
    Dim index           As Long
    Dim theFindings     As MSXML2.IXMLDOMNodeList
    Dim theQuery        As String

    theQuery = Replace$("/Assessment/report/standard[%1]/findings", mgrP1, CStr(standardNumber))
    Set theFindings = g_xmlDocument.SelectNodes(theQuery)

    For index = 1 To theFindings.Length
        Set findingsNode = theFindings(index - 1)
        Debug.Print findingsNode.ChildNodes(0).Text, findingsNode.ChildNodes(2).Text
    Next
End Sub ' DumpFindingsXML

Private m_editableBookmarkInfoSets As VBA.Collection

Public Sub BuildEditableBookamrkInfo()
    

End Sub
