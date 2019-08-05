VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultsTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================
'Objective: Display the results for the user in a table.
'Author: Daniel Correa de Castro Freitas
'==================================================================

Private outputTable As Range

Private Sub Class_Initialize()

With Program
    Set outputTable = .Range("outputTable")
End With

End Sub

Private Sub Class_Terminate()

With Program
    Set outputTable = Nothing
End With

End Sub

Sub ReturnResultsTable(ByRef results As Variant)

Dim i As Long
Dim newRowsNumber As Long

newRowsNumber = UBound(results)
prepareTableForNewData newRowsNumber

If results(0, 0) <> "" Then
    fillEachColumn results, newRowsNumber
End If

unlockTable

End Sub

Private Sub prepareTableForNewData(ByVal newRowsNumber As Long)

clearPreviousExecution

resizeTable newRowsNumber

allignCellsContent

End Sub

Private Sub clearPreviousExecution()

With outputTable
    .Clear
    If .Rows.Count > 1 Then
        .Delete
    End If
End With

End Sub

Private Sub resizeTable(ByVal newRowsNumber As Long)

Dim outputTableAsObject As Object
Dim outputHeaders As Range

With Program
    Set outputTableAsObject = .ListObjects("outputTable")
    Set outputHeaders = .Range("outputTable[#Headers]")
End With

outputTableAsObject.Resize Range(outputHeaders, outputHeaders.Offset(newRowsNumber + 1, 0))

Set outputTableAsObject = Nothing
Set outputHeaders = Nothing

End Sub

Private Sub allignCellsContent()

With outputTable
    .VerticalAlignment = xlBottom
    .HorizontalAlignment = xlLeft
End With

End Sub

Private Sub unlockTable()

Program.Range("outputTable").Locked = False

End Sub

Private Sub fillEachColumn(ByRef results As Variant, ByVal newRowsNumber As Integer)

Dim outputIdentifierCells As Range
Dim outputInformationCells As Range
Dim outputIdentifier() As String
Dim outputInformation() As String

With Program.Range("outputTable[#Headers]")
    Set outputIdentifierCells = Program.Range("outputTable[" & .Cells(1, 1) & "]")
    Set outputInformationCells = Program.Range("outputTable[" & .Cells(1, 2) & "]")
End With

ReDim Preserve outputIdentifier(newRowsNumber)
ReDim Preserve outputInformation(newRowsNumber)

splitResultsIntoTwoArrays results, outputIdentifier, outputInformation

outputIdentifierCells = Application.Transpose(outputIdentifier)
outputInformationCells = Application.Transpose(outputInformation)

End Sub

Private Sub splitResultsIntoTwoArrays(ByRef results As Variant, outputIdentifier As Variant, ByRef outputInformation As Variant)

Dim i As Long
For i = 0 To UBound(results)
    outputIdentifier(i) = results(i, 0)
    outputInformation(i) = results(i, 1)
Next

End Sub
