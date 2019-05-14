VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Program"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

DoEvents
Application.EnableEvents = False

Dim KW As KeyWords
Dim UB As UserBase
Dim RT As ResultsTable
Dim filterWords As Variant
Dim filteredIdentifierAndInformation As Variant

Set KW = New KeyWords

If KW.checkNewInputExistance = True Then
    unprotectTable
    showIconAsBusy
    
    Set UB = New UserBase
    Set RT = New ResultsTable
    
    filterWords = KW.returnFilterWords
    filteredIdentifierAndInformation = UB.filterIdentifierAndInformation(filterWords)
    RT.populateTableWithSearchResult (filteredIdentifierAndInformation)
    
    showIconAsAvailable
    protectTable
End If

Application.EnableEvents = True

End Sub

Private Sub showIconAsAvailable()

ThisWorkbook.Sheets("Programa").Shapes.Range(Array("Oval 1")).Fill.ForeColor.RGB = RGB(0, 255, 0)

End Sub

Private Sub unprotectTable()

ThisWorkbook.Sheets("Programa").Unprotect

End Sub

Private Sub showIconAsBusy()

ThisWorkbook.Sheets("Programa").Shapes.Range(Array("Oval 1")).Fill.ForeColor.RGB = RGB(255, 0, 0)

End Sub

Private Sub protectTable()

ThisWorkbook.Sheets("Programa").Protect

End Sub
