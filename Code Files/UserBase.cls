VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=================================================================
'Objective: Return the Identifier and Information userBased on a
'           given array. If the word(s) exists in the Identifier,
'           an array is returned.
'Version: v1.0
'Creation Date: 18/04/2019
'Modification Date: 25/04/2019
'Author: Daniel Correa de Castro Freitas
'=================================================================

Private placeIdentifierColumn As Range
Private descriptionColumn() As String
Private informationColumn() As String

Private Sub Class_Initialize()

Set placeIdentifierColumn = ThisWorkbook.Sheets("Base").Range("userBase[Identificador]")
returnSeparateIdentifierAndInformation

End Sub

Private Sub Class_Terminate()
 
Set placeIdentifierColumn = Nothing

End Sub

Private Sub returnSeparateIdentifierAndInformation()

Dim cell As Range
Dim i As Long

For Each cell In placeIdentifierColumn
    ReDim Preserve descriptionColumn(i)
    ReDim Preserve informationColumn(i)
    
   descriptionColumn(i) = cell.Value
   informationColumn(i) = cell.Offset(0, 1).Value
   i = i + 1
Next

End Sub
Function filterIdentifierAndInformation(ByRef targetWords As Variant)

'Objective: Return an filtered array with Identifier and Information
'           using the words given to this method.
'           This function works filtering the Identifier column and
'           using its rows as a referente to bring the Information,
'           joining both into a third array with 2 columns.
'           The Separation was need because of the Filter() use,
'           which only works in 1 dimension arrays.

Dim filteredIdentifier() As String
Dim filteredInformation() As String
Dim filteredIdentifierAndInformation() As String
Dim i As Long

If checkAreNoWords(targetWords) = True Then
    ReDim filteredIdentifierAndInformation(0, 0)
    filterIdentifierAndInformation = filteredIdentifierAndInformation
Else
    filteredIdentifier = returnMatchedIdentifier(targetWords)
    If checkEmptyReturn(filteredIdentifier) = True Then
        filteredInformation = filterInformation(filteredIdentifier)
        
        ReDim Preserve filteredIdentifierAndInformation(UBound(filteredIdentifier), 1)
        For i = 0 To UBound(filteredIdentifier)
            filteredIdentifierAndInformation(i, 0) = filteredIdentifier(i)
            filteredIdentifierAndInformation(i, 1) = filteredInformation(i)
        Next
        
        filterIdentifierAndInformation = filteredIdentifierAndInformation
    Else
        ReDim filteredIdentifierAndInformation(0, 0)
        filterIdentifierAndInformation = filteredIdentifierAndInformation
    End If
End If

End Function

Private Function checkAreNoWords(ByRef targetWords As Variant) As Boolean

If Len(Join(targetWords)) = 0 Then
    checkAreNoWords = True
End If

End Function

Private Function returnMatchedIdentifier(ByRef targetWords As Variant)

Dim descriptionCopy() As String
Dim i As Long

'To avoid the unacessary property manipulation
descriptionCopy = descriptionColumn
For i = 0 To UBound(targetWords)
    descriptionCopy = Filter(descriptionCopy, targetWords(i), True, vbTextCompare)
Next

returnMatchedIdentifier = descriptionCopy

End Function

Private Function checkEmptyReturn(ByRef filteredIdentifier() As String) As Boolean

If UBound(filteredIdentifier) > -1 Then
    checkEmptyReturn = True
End If

End Function

Private Function filterInformation(ByRef filteredIdentifier() As String)

Dim currentIdentifierRow As Long
Dim filteredInformation() As String

ReDim Preserve filteredInformation(UBound(filteredIdentifier))

Dim i As Long
For i = 0 To UBound(filteredIdentifier)
    'This is a common excel formula combination between Index and Match
    currentIdentifierRow = WorksheetFunction.Match(filteredIdentifier(i), descriptionColumn, 0)
    filteredInformation(i) = WorksheetFunction.Index(informationColumn, currentIdentifierRow)
Next

filterInformation = filteredInformation

End Function
