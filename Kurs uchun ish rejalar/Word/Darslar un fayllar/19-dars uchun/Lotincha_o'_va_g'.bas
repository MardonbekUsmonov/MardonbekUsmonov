Attribute VB_Name = "Lotincha_o"
Sub Lotincha_o()


' this Macros was made by Sherzod Qodirov www.videokent.uz in 29.03.2009
'

  Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "g'"
        .Replacement.Text = "g‘"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "o'"
        .Replacement.Text = "o‘"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "G'"
        .Replacement.Text = "G‘"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "O'"
        .Replacement.Text = "O‘"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

End Sub
