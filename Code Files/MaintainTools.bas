Attribute VB_Name = "MaintainTools"
'==================================================================
'Objective: Module dedicate to hold the maintain tools used during
'           the development
'Author: Daniel Correa de Castro Freitas
'==================================================================

Private Sub enableEvents()

Application.enableEvents = True

End Sub

Private Sub disableEvents()

Application.enableEvents = False

End Sub
Private Sub showRibbonAndFormulaBar()
    
With Application
    .DisplayFormulaBar = True
    .ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",true)"
End With

End Sub

