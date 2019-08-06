VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LangSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub PTLanguage()

    ChangeLanguage "PT"

End Sub

Sub ENLanguage()

    ChangeLanguage "EN"

End Sub
Private Sub ChangeLanguage(ByVal LanguageIndicator As String)
    
    'Objective: Change the Worbkook language using a support sheet.
    '           This procedure changes the Sheet names and Ranges in
    '           the Tutorial, Program and Database sheets.
    
    Dim i As Long
    Dim RowsLangTable As Long
    i = 0

    If Support.Range("CurLang[Current Language]") <> LanguageIndicator Then
        
        ChangeSheetNames LanguageIndicator
        
        PrepareWorkbook
        
        With LangSupport
            RowsLangTable = Range("A1").CurrentRegion.Rows.Count
            For i = 2 To RowsLangTable
                If .Cells(i, 3) = LanguageIndicator Then
                    If .Cells(i, 4) <> "" Then
                        .Cells(i, 4).MergeArea.Copy Sheets(Cells(i, 1)).Range(Cells(i, 2)).MergeArea
                   End If
                End If
            Next
        End With
        
        Support.Range("CurLang[Current Language]") = LanguageIndicator
        
        ReturnOriginalState
        
    End If
    
End Sub

Private Sub PrepareWorkbook()

    'Objective: Remove the protection and hiding the
    '           execution from the user.
    
    Application.ScreenUpdating = False
    Application.enableEvents = False
    
    Program.Unprotect
    Tutorial.Unprotect
    Database.Unprotect
    
End Sub

Private Sub ChangeSheetNames(ByVal LanguageIndicator As String)

    'Objective: Change the sheet names based on the LanguageIndicator
    
    If LanguageIndicator = "PT" Then
        Program.Name = "Programa"
        Tutorial.Name = "Instruções"
        Database.Name = "Base"
    Else
        Program.Name = "Program"
        Tutorial.Name = "Instructions"
        Database.Name = "Database"
    End If
    
End Sub

Private Sub ReturnOriginalState()
    
    Database.Protect
    Tutorial.Protect
    Program.Protect
    
    Application.ScreenUpdating = True
    Application.enableEvents = True
    
End Sub

