VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Program"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'==================================================================
'Objective: Use the events to trigger the methods in the classes,
'           as well changing the progress icon and protecting or
'           unprotecting this sheet.
'Author: Daniel Correa de Castro Freitas
'==================================================================

Private Sub Worksheet_Change(ByVal Target As Range)

    DoEvents
    Application.enableEvents = False
    
    Dim kw As KeyWords
    Dim ub As UserBase
    Dim rt As ResultsTable
    Dim filterWords As Variant
    Dim filteredIdentifierAndInformation As Variant
    
    Set kw = New KeyWords
    
    If kw.checkNewInputExistance = True Then
        unprotectTable
        showIconAsBusy
        
        Set ub = New UserBase
        Set rt = New ResultsTable
        
        filterWords = kw.returnFilterWords
        filteredIdentifierAndInformation = ub.filterIdentifierAndInformation(filterWords)
        rt.ReturnResultsTable (filteredIdentifierAndInformation)
        
        showIconAsAvailable
        protectTable
    End If
    
    Application.enableEvents = True

End Sub

Private Sub showIconAsAvailable()

    Program.Shapes.Range(Array("Oval 1")).Fill.ForeColor.RGB = RGB(0, 255, 0)

End Sub

Private Sub unprotectTable()

    Program.Unprotect

End Sub

Private Sub showIconAsBusy()

    Program.Shapes.Range(Array("Oval 1")).Fill.ForeColor.RGB = RGB(255, 0, 0)

End Sub

Private Sub protectTable()

    Program.Protect

End Sub

