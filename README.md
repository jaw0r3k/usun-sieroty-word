# usun-sieroty-word
Makro w wordzie usuwające litery (i, a, w , z, o, u) z końca lini

## Kod

```bas 
Sub UsunSieroty()
    Application.ScreenUpdating = False

    Dim doZnalezienia, doZamiany

    doZnalezienia = Array(" a ", " i ", " o ", " u ", " w ", " z ", " A ", " I ", " O ", " U ", " W ", " Z ")
    doZamiany = Array(" a^s", " i^s", " o^s", " u^s", " w^s", " z^s", " A^s", " I^s", " O^s", " U^s", " W^s", " Z^s")

    Selection.HomeKey Unit:=wdStory

    For i = LBound(doZnalezienia) To UBound(doZnalezienia)
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Format = False
            .Forward = True
            .Wrap = wdFindContinue
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchCase = True
            .Text = doZnalezienia(i)
            .Replacement.Text = doZamiany(i)
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
        Selection.HomeKey Unit:=wdStory
    Next i

    Application.ScreenUpdating = True

    MsgBox "Gotowe!", vbInformation
End Sub
``` 

## Jak stworzyć makro

W wordzie w zakładce Widok klikamy Makro > Wyświetl Makra
Następnie tworzymy nowe makro, wklejamy kod i gotowe! Możemy uruchomić je w każdym momencie

## Zasada działania

Kod wyszukuje wolne litery a następnie zamienia spację po nich na twardą spację 
