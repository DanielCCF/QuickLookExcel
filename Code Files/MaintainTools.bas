Attribute VB_Name = "MaintainTools"
'==================================================================
'Objective: Module dedicate to hold the maintain tools used during
'           the development
'Creation Date: 16/04/2019
'Modification Date: 25/04/2019
'Author: Daniel Correa de Castro Freitas
'==================================================================

Private Sub enableEvents()

Application.enableEvents = False

End Sub

Private Sub showRibbonAndFormulaBar()
    
With Application
    .DisplayFormulaBar = True
    .ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",true)"
End With

End Sub

