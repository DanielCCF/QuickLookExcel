VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub PreviousRow()
 
    ReturnRow -1
 
End Sub


Private Sub NextRow()

    ReturnRow 1

End Sub


Private Sub ReturnRow(ByVal RowChanger As Integer)

    Dim LastRowDatabase As Integer
    Dim CurrentRowReference As Range
    Dim NewRowValue As Integer
    
    LastRowDatabase = Database.Range("B1").CurrentRegion.Rows.Count
    Set CurrentRowReference = Support.Range("CurRows")
    NewRowValue = CurrentRowReference + RowChanger
    
    If NewRowValue < LastRowDatabase - 4 And NewRowValue >= 0 Then _
        CurrentRowReference = NewRowValue
    
End Sub



Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    Dim NewSelectionIsValidForChange As Boolean
    
    NewSelectionIsValidForChange = Intersect(Target, DataAccess.Range("EditDeleteBlock")) Is Nothing
    
    If Not NewSelectionIsValidForChange Then
        DataAccess.Range("NewIdentifier").Select
    End If
End Sub
