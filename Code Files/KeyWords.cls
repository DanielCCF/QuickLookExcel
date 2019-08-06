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

Private currentInput() As String
Private placeCurrentInput As Range
Private previousExecution() As String
Private placePreviousExecution As Range

Private Sub Class_Initialize()

    Set placeCurrentInput = Program.Range("inputTable[" & Program.Range("inputTable[#Headers]").Cells(1, 2) & "]")
    Set placePreviousExecution = Support.Range("supportTable[" & Support.Range("supportTable[#Headers]").Cells(1, 2) & "]")
    
    currentInput = returnValuesFromPlace(placeCurrentInput)
    previousExecution = returnValuesFromPlace(placePreviousExecution)

End Sub

Private Sub Class_Terminate()

    Set placeCurrentInput = Nothing
    Set placePreviousExecution = Nothing

End Sub

Function checkNewInputExistance() As Boolean

    If checkIfNewInputIsDifferentFromPrevious = True Then
        checkNewInputExistance = True
        'Keep track of previous execution is a MUST
        registerCurrentExecutionAsPrevious
    Else
        checkNewInputExistance = False
    End If

End Function

Private Function checkIfNewInputIsDifferentFromPrevious() As Boolean

    Dim i As Byte
    
    For i = 0 To UBound(currentInput)
        If currentInput(i) <> previousExecution(i) Then
            checkIfNewInputIsDifferentFromPrevious = True
            Exit For
        End If
    Next

End Function

Private Function returnValuesFromPlace(ByVal targetRange As Range)

    Dim cellValue As Range
    Dim i As Long
    Dim valuesFromRange() As String
    
    For Each cellValue In targetRange
        ReDim Preserve valuesFromRange(i)
        valuesFromRange(i) = cellValue.Value
        i = i + 1
    Next
    
    returnValuesFromPlace = valuesFromRange

End Function

Private Sub registerCurrentExecutionAsPrevious()
    
    Dim place As Range
    Dim i As Integer
    
    For Each place In placePreviousExecution
        place.Value = currentInput(i)
        i = i + 1
    Next

End Sub

Function returnFilterWords() As Variant

    Dim i As Integer, j As Integer
    Dim nonBlankValues() As String
    
    For i = 0 To UBound(currentInput)
        If currentInput(i) <> "" Then
            ReDim Preserve nonBlankValues(j)
            nonBlankValues(j) = currentInput(i)
            j = j + 1
        End If
    Next
    
    returnFilterWords = nonBlankValues

End Function
