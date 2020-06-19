VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'==================================================================
'Objective: Compare the current input values with the previous one.
'Author: Daniel Correa de Castro Freitas
'==================================================================

Private CurrentInput() As String
Private CellsCurrentInput As Range
Private PreviousExecution() As String
Private CellsPreviousExecution As Range


Private Sub Class_Initialize()
    
    Dim currentHeaderInputTableSecondColumn As String
    Dim currentHeaderSupportTableSecondColumn As String
    
    currentHeaderInputTableSecondColumn = Program.Range("inputTable[#Headers]").Cells(1, 2)
    currentHeaderSupportTableSecondColumn = Support.Range("supportTable[#Headers]").Cells(1, 2)
    
    Set CellsCurrentInput = Program.Range("inputTable[" & currentHeaderInputTableSecondColumn & "]")
    Set CellsPreviousExecution = Support.Range("supportTable[" & currentHeaderSupportTableSecondColumn & "]")
    
    CurrentInput = ReturnValuesFromCells(CellsCurrentInput)
    PreviousExecution = ReturnValuesFromCells(CellsPreviousExecution)

End Sub


Private Sub Class_Terminate()

    Set CellsCurrentInput = Nothing
    Set CellsPreviousExecution = Nothing

End Sub


Private Function ReturnValuesFromCells(ByVal targetRange As Range)

    Dim cellValue As Range
    Dim i As Long
    Dim valuesFromRange() As String
    
    For Each cellValue In targetRange
        ReDim Preserve valuesFromRange(i)
        valuesFromRange(i) = cellValue.Value
        i = i + 1
    Next
    
    ReturnValuesFromCells = valuesFromRange

End Function


Private Sub RegisterCurrentExecutionAsPrevious()
    
    Dim place As Range
    Dim i As Integer
    
    For Each place In CellsPreviousExecution
        place.Value = CurrentInput(i)
        i = i + 1
    Next

End Sub


Function CheckNewInputExistance() As Boolean

    If CheckIfNewInputIsDifferentFromPrevious = True Then
        CheckNewInputExistance = True
        'Keep track of previous execution is a MUST, in this
        'way it'll only execute if a change in cell values
        'is made.
        RegisterCurrentExecutionAsPrevious
    Else
        CheckNewInputExistance = False
    End If

End Function


Private Function CheckIfNewInputIsDifferentFromPrevious() As Boolean

    Dim i As Byte
    
    For i = 0 To UBound(CurrentInput)
        If CurrentInput(i) <> PreviousExecution(i) Then
            CheckIfNewInputIsDifferentFromPrevious = True
            Exit For
        End If
    Next

End Function


Function ReturnFilterWords() As Variant

    Dim i As Integer, j As Integer
    Dim nonBlankValues() As String
    
    For i = 0 To UBound(CurrentInput)
        If CurrentInput(i) <> "" Then
            ReDim Preserve nonBlankValues(j)
            nonBlankValues(j) = CurrentInput(i)
            j = j + 1
        End If
    Next
    
    ReturnFilterWords = nonBlankValues

End Function
