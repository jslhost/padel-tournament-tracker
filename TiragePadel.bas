Attribute VB_Name = "TiragePadel"
' ================================================================
' TIRAGE AU SORT ANTI-CONFLIT — Tournoi Padel Phase 2
' ================================================================
' INSTALLATION (une seule fois) :
'   1. Ouvrir l Editeur VBA : Alt + F11
'   2. Inserer > Module
'   3. Coller ce code entier dans le module
'   4. Fermer l editeur (Alt+F4)
'   5. Dans l onglet "Classement Phase 1", faire un clic droit
'      sur la cellule V51 (bouton bleu "NOUVEAU TIRAGE")
'      > "Assigner une macro" > choisir "NouveauTirage" > OK
'   6. Sauvegarder le fichier en .xlsm
' ================================================================

Sub NouveauTirage()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Classement Phase 1")
    
    Application.ScreenUpdating = False
    ws.Calculate
    
    ' Lecture 8 equipes TOP (col D = nom, col U = poule origine)
    Dim topSrcRows As Variant
    topSrcRows = Array(54, 55, 56, 57, 58, 60, 61, 62)
    
    Dim topTeams(7) As String, topOrig(7) As String
    Dim i As Integer
    For i = 0 To 7
        topTeams(i) = ws.Cells(topSrcRows(i), 4).Value
        topOrig(i) = ws.Cells(topSrcRows(i), 21).Value
    Next i
    
    ' Lecture 12 equipes CONSOLANTE
    Dim consSrcRows As Variant
    consSrcRows = Array(66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 77, 78)
    
    Dim consTeams(11) As String, consOrig(11) As String
    For i = 0 To 11
        consTeams(i) = ws.Cells(consSrcRows(i), 4).Value
        consOrig(i) = ws.Cells(consSrcRows(i), 21).Value
    Next i
    
    ' Tirage anti-conflit (3000 essais, on garde le meilleur)
    Dim bestTop(7) As String, bestTopO(7) As String
    Dim bestCons(11) As String, bestConsO(11) As String
    Dim bestTopSc As Long: bestTopSc = 999
    Dim bestConsSc As Long: bestConsSc = 999
    Dim tmp(11) As String, tmpO(11) As String
    Dim sc As Long, trial As Long, j As Long, t As String, tempOrigine As String
    
    Randomize
    
    ' -- Tirage TOP (2 poules de 4) --
    For trial = 1 To 3000
        For i = 0 To 7: tmp(i) = topTeams(i): tmpO(i) = topOrig(i): Next i
        For i = 7 To 1 Step -1
            j = Int(Rnd() * (i + 1))
            t = tmp(i): tmp(i) = tmp(j): tmp(j) = t
            tempOrigine = tmpO(i): tmpO(i) = tmpO(j): tmpO(j) = tempOrigine
        Next i
        sc = Conflicts(tmpO, 0, 3) + Conflicts(tmpO, 4, 7)
        If sc < bestTopSc Then
            bestTopSc = sc
            For i = 0 To 7: bestTop(i) = tmp(i): bestTopO(i) = tmpO(i): Next i
            If sc = 0 Then Exit For
        End If
    Next trial
    
    ' -- Tirage CONSOLANTE (3 poules de 4) --
    For trial = 1 To 3000
        For i = 0 To 11: tmp(i) = consTeams(i): tmpO(i) = consOrig(i): Next i
        For i = 11 To 1 Step -1
            j = Int(Rnd() * (i + 1))
            t = tmp(i): tmp(i) = tmp(j): tmp(j) = t
            tempOrigine = tmpO(i): tmpO(i) = tmpO(j): tmpO(j) = tempOrigine
        Next i
        sc = Conflicts(tmpO, 0, 3) + Conflicts(tmpO, 4, 7) + Conflicts(tmpO, 8, 11)
        If sc < bestConsSc Then
            bestConsSc = sc
            For i = 0 To 11: bestCons(i) = tmp(i): bestConsO(i) = tmpO(i): Next i
            If sc = 0 Then Exit For
        End If
    Next trial
    
    ' -- Ecriture des resultats dans le tableau --
    Dim fR As Variant: fR = Array(57, 58, 59, 60)
    Dim gR As Variant: gR = Array(64, 65, 66, 67)
    Dim hR As Variant: hR = Array(73, 74, 75, 76)
    Dim iR As Variant: iR = Array(80, 81, 82, 83)
    Dim jR As Variant: jR = Array(87, 88, 89, 90)
    
    For i = 0 To 3
        ws.Cells(fR(i), 23).Value = bestTop(i)      ' col W = nom equipe
        ws.Cells(fR(i), 24).Value = bestTopO(i)     ' col X = poule Ph.1
        ws.Cells(gR(i), 23).Value = bestTop(i + 4)
        ws.Cells(gR(i), 24).Value = bestTopO(i + 4)
        ws.Cells(hR(i), 23).Value = bestCons(i)
        ws.Cells(hR(i), 24).Value = bestConsO(i)
        ws.Cells(iR(i), 23).Value = bestCons(i + 4)
        ws.Cells(iR(i), 24).Value = bestConsO(i + 4)
        ws.Cells(jR(i), 23).Value = bestCons(i + 8)
        ws.Cells(jR(i), 24).Value = bestConsO(i + 8)
    Next i
    
    ' Mise a jour ligne info conflits
    Dim msgTop As String, msgCons As String
    If bestTopSc = 0 Then msgTop = "Top: OK (0 conflit)" Else msgTop = "Top: " & bestTopSc & " conflit(s) inevitable(s)"
    If bestConsSc = 0 Then msgCons = "Consolante: OK (0 conflit)" Else msgCons = "Consolante: " & bestConsSc & " conflit(s) inevitable(s)"
    ws.Cells(52, 22).Value = msgTop & "   |   " & msgCons
    
    Application.ScreenUpdating = True
    MsgBox "Tirage effectue !" & Chr(10) & Chr(10) & msgTop & Chr(10) & msgCons, vbInformation, "Tirage au sort"
End Sub

Function Conflicts(orig() As String, startI As Integer, endI As Integer) As Long
    Dim sc As Long: sc = 0
    Dim i As Integer, j As Integer
    For i = startI To endI - 1
        For j = i + 1 To endI
            If orig(i) <> "" And orig(j) <> "" And orig(i) = orig(j) Then sc = sc + 1
        Next j
    Next i
    Conflicts = sc
End Function

