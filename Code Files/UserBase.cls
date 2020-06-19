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
'Author: Daniel Correa de Castro Freitas
'=================================================================

Private PlaceIdentifierColumn As Range
Private DescriptionColumn() As String
Private InformationColumn() As String


Private Sub Class_Initialize()

    With Database
        Set PlaceIdentifierColumn = .Range("userBase[" & .Range("userBase[#Headers]").Cells(1, 1) & "]")
    End With
    
    ReturnSeparateIdentifierAndInformation

End Sub


Private Sub Class_Terminate()
 
    Set PlaceIdentifierColumn = Nothing

End Sub


Private Sub ReturnSeparateIdentifierAndInformation()

    Dim cell As Range
    Dim i As Long
    
    For Each cell In PlaceIdentifierColumn
        ReDim Preserve DescriptionColumn(i)
        ReDim Preserve InformationColumn(i)
        
       DescriptionColumn(i) = cell.Value
       InformationColumn(i) = cell.Offset(0, 1).Value
       i = i + 1
    Next

End Sub


Function FilterIdentifierAndInformation(ByRef targetWords As Variant)

'Objective: Return a filtered array with Identifier and Information
'           using the words given to this method.
'           This function works filtering the Identifier column and
'           using its rows as a reference to bring the Information,
'           joining both into a third array with 2 columns.
'           The Separation was need because of the Filter() use,
'           which only works in 1 dimension arrays.

    Dim filteredIdentifier() As String
    Dim filteredInformation() As String
    Dim filteredIdentifierAndInformation() As String
    Dim i As Long
    
    If CheckAreNoWords(targetWords) = True Then
        ReDim filteredIdentifierAndInformation(0, 0)
        FilterIdentifierAndInformation = filteredIdentifierAndInformation
    Else
        filteredIdentifier = ReturnMatchedIdentifier(targetWords)
        If CheckEmptyReturn(filteredIdentifier) = True Then
            filteredInformation = FilterInformation(filteredIdentifier)
            
            ReDim Preserve filteredIdentifierAndInformation(UBound(filteredIdentifier), 1)
            For i = 0 To UBound(filteredIdentifier)
                filteredIdentifierAndInformation(i, 0) = filteredIdentifier(i)
                filteredIdentifierAndInformation(i, 1) = filteredInformation(i)
            Next
            
            FilterIdentifierAndInformation = filteredIdentifierAndInformation
        Else
            ReDim filteredIdentifierAndInformation(0, 0)
            FilterIdentifierAndInformation = filteredIdentifierAndInformation
        End If
    End If

End Function

Private Function CheckAreNoWords(ByRef targetWords As Variant) As Boolean
    
    If Len(Join(targetWords)) = 0 Then
        CheckAreNoWords = True
    End If

End Function

Private Function ReturnMatchedIdentifier(ByRef targetWords As Variant)

    Dim descriptionCopy() As String
    Dim i As Long
    
    'To avoid the unacessary property manipulation
    descriptionCopy = DescriptionColumn
    For i = 0 To UBound(targetWords)
        descriptionCopy = Filter(descriptionCopy, targetWords(i), True, vbTextCompare)
    Next
    
    ReturnMatchedIdentifier = descriptionCopy

End Function

Private Function CheckEmptyReturn(ByRef filteredIdentifier() As String) As Boolean

    If UBound(filteredIdentifier) > -1 Then
        CheckEmptyReturn = True
    End If

End Function

Private Function FilterInformation(ByRef filteredIdentifier() As String)

    Dim currentIdentifierRow As Long
    Dim filteredInformation() As String
    
    ReDim Preserve filteredInformation(UBound(filteredIdentifier))
    
    Dim i As Long
    For i = 0 To UBound(filteredIdentifier)
        'This is a common excel formula combination between Index and Match
        currentIdentifierRow = WorksheetFunction.Match(filteredIdentifier(i), DescriptionColumn, 0)
        filteredInformation(i) = WorksheetFunction.Index(InformationColumn, currentIdentifierRow)
    Next
    
    FilterInformation = filteredInformation

End Function
