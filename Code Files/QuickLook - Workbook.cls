VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QuickLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Workbook_Deactivate()

With Application
    .DisplayFormulaBar = True
    .ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",true)"
End With
End Sub

Sub Workbook_Activate()

With Application
    Application.DisplayFormulaBar = False
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"",False)"
End With

End Sub


