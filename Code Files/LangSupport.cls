VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LangSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub PTLanguage()


ChangeLanguage "PT"

End Sub

Sub ENLanguage()


ChangeLanguage "EN"

End Sub

Private Sub ChangeLanguage(ByVal LanguageIndicator As String)
    
    Dim i As Long
    Dim RowsLangTable As Long
    i = 0

    If Support.Range("CurLang[Current Language]") <> LanguageIndicator Then
        Application.ScreenUpdating = False
        Application.enableEvents = False
        
        Program.Unprotect
        Tutorial.Unprotect
        Database.Unprotect
        
        If LanguageIndicator = "PT" Then
            Program.Name = "Programa"
            Tutorial.Name = "Instruções"
            Database.Name = "Base"
        Else
            Program.Name = "Program"
            Tutorial.Name = "Instructions"
            Database.Name = "Database"
        End If
        
        With LangSupport
            .Visible = xlSheetVisible
            RowsLangTable = Range("A1").CurrentRegion.Rows.Count
            For i = 2 To RowsLangTable
                If .Cells(i, 3) = LanguageIndicator Then
                    If .Cells(i, 4) <> "" Then
                        .Cells(i, 4).MergeArea.Copy Sheets(Cells(i, 1)).Range(Cells(i, 2)).MergeArea
                   End If
                End If
            Next
            .Visible = xlSheetVeryHidden
        End With
        Support.Range("CurLang[Current Language]") = LanguageIndicator
    
        Database.Protect
        Tutorial.Protect
        Program.Protect
        
        Tutorial.Activate
        
        Application.ScreenUpdating = True
        Application.enableEvents = True
    End If
    
End Sub


