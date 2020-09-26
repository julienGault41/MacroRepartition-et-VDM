Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Global nbligne As Integer
Global nbREFClient As Integer
Global nbREFCata As Integer
Global config(10) As String
Global nbCH As Integer
Global nbIM As Integer
Global nbSH As Integer
Global REFClient(1000) As String
Global ACTClient(1000) As Double
Global nbModREF(1000) As Integer

Global InterCH(5, 2) As Double
Global InterIM(5, 2) As Double
Global FamillePT(26) As String
Global TabCree As Boolean
Global NBLPFCH, NBLPFIM As Integer
Global NBJourVDM As Integer
Global ImportSaisie As Boolean
Global NBJourREPET As Integer
Global NBJourCOMP As Integer
Global NBJourAUTRE As Integer
Global DLF As Double 'Derni?re ligne de la feuille fonctionnement'
Global NbFRef As Integer 'Nombre de Ref différentes dans fonctionnement'
Global FeuiF As Boolean 'Flag pour prŽsence feuille fonctionnement'


Function NombreREF(SMN As String) As Integer 'Retourne le nombre de références du tableau fonctionnement'
Dim ct, tot, i As Integer
Dim tp As String
ct = 1
tot = 0
For i = 1 To 1000
tp = Worksheets("Fonctionnement").Cells(i + 3, 12)
If tp = SMN Then
tot = tot + 1
End If
Next i
NombreREF = tot
End Function




Sub contrat()
Dim cxh, cxi, tpx As String
Dim i As Integer

cxh = Worksheets("Repar. CH").Cells(3, 1).Value
cxi = Worksheets("Repar. IM").Cells(3, 1).Value
affMessage "Analyse des données du contrat."
For i = 1 To cxh
    tpx = Worksheets("Repar. CH").Cells(4 + i, 1).Value
    For j = 1 To nbREFClient
        If tpx = REFClient(j) Then
        Worksheets("Repar. CH").Cells(4 + i, 30).Value = nbModREF(j)
        End If
    Next j
Next i
For i = 1 To cxi
    tpx = Worksheets("Repar. IM").Cells(4 + i, 1)
    For j = 1 To nbREFClient
        If tpx = REFClient(j) Then
        Worksheets("Repar. IM").Cells(4 + i, 30).Value = nbModREF(j)
        End If
    Next j
Next i
affMessage ""
End Sub

Sub Init()
Dim tampon, tampon2, car As String
Dim cpt, i, X As Integer

nbligne = 0
cpt = 1
nbCH = 0
nbIM = 0
nbSH = 0
nbREFClient = 0
NBLPFCH = 0
NBLPFIM = 0
TabCree = False

Do
tampon = Worksheets("Saisie").Cells(cpt + 3, 2)
If tampon = "" Then
Exit Do
Else
nbligne = nbligne + 1
config(nbligne) = tampon
End If
cpt = cpt + 1
Loop
For cpt = 1 To nbligne
    tampon = config(cpt)
    For i = 1 To Len(tampon)
        car = VBA.Mid$(tampon, i, 1)
        If car = "S" Then nbSH = nbSH + 1
        If car = "C" Then nbCH = nbCH + 1
        If car = "I" Then nbIM = nbIM + 1
    Next i
Next cpt
cpt = 1
Do
tampon = Worksheets("Saisie").Cells(cpt + 4, 6)
tampon2 = Worksheets("Saisie").Cells(cpt + 4, 7)
If tampon = "" Then
Exit Do
Else
nbREFClient = nbREFClient + 1
REFClient(nbREFClient) = tampon
ACTClient(nbREFClient) = tampon2
cpt = cpt + 1
End If
Loop
cpt = 1
nbREFCata = 0
Do
tampon = Worksheets("Sources").Cells(cpt + 1, 2)
If tampon = "" Then
Exit Do
Else
nbREFCata = nbREFCata + 1
End If
cpt = cpt + 1
Loop
For X = 0 To 26
FamillePT(X) = Worksheets("Sources").Cells(3 + X, 19).Value
Next X
For X = 1 To 5
InterCH(X, 1) = Worksheets("Saisie").Cells(18 + X, 2).Value
InterCH(X, 2) = Worksheets("Saisie").Cells(18 + X, 3).Value
InterIM(X, 1) = Worksheets("Saisie").Cells(27 + X, 2).Value
InterIM(X, 2) = Worksheets("Saisie").Cells(27 + X, 3).Value
Next X
End Sub
Function DejaInclus(SMN As String) As String
Dim i As Integer
For i = 1 To nbREFClient
If SMN = REFClient(i) Then
DejaInclus = "Oui"
End If
Next i
DejaInclus = "Non"
End Function

Sub initdeux()
Dim tp, car As String
Dim cpt, i As Integer
nbSH = 0
nbCH = 0
nbIM = 0
nbligne = 0
cpt = 1
Do
tp = Worksheets("Saisie").Cells(cpt + 3, 2).Value
If tp = "" Then
Exit Do
Else
nbligne = nbligne + 1
config(nbligne) = tp
End If
cpt = cpt + 1
Loop
For cpt = 1 To nbligne
    tp = config(cpt)
    For i = 1 To Len(tp)
        car = VBA.Mid$(tp, i, 1)
        If car = "S" Then nbSH = nbSH + 1
        If car = "C" Then nbCH = nbCH + 1
        If car = "I" Then nbIM = nbIM + 1
    Next i
Next cpt
End Sub

Function FamilleVersID(SMN As String) As String
Dim i As Integer
Dim tampon As String
For i = 1 To nbREFCata
    tampon = Worksheets("Sources").Cells(i + 1, 2)
    If SMN = tampon Then
        FamilleVersID = Worksheets("Sources").Cells(i + 1, 12)
    End If
Next i
End Function
Function FamilleTotActi(fam As Integer) As Double
Dim tot As Double
Dim i As Integer
tot = 0
For i = 1 To nbREFClient
    If Extrait(REFClient(i), 12) = fam Then
    tot = tot + ACTClient(i)
    End If
Next i
FamilleTotActi = tot
End Function
Function Extrait(SMN As String, obj As Integer) As String 'OBJ : 1=Platform, 2=SMN, 3=Nom, 4=Abr_viation, 5=nb Tests, 6=Nb pack, 7=tests/pack, 8=Materiel 9=Opt. Materiel 10=volume 11=Famille 12 ID Famille'
Dim cpt As Integer
Dim tampon As String
For cpt = 1 To nbREFCata
tampon = Worksheets("Sources").Cells(cpt + 1, 2)
If tampon = SMN Then
   Extrait = Worksheets("Sources").Cells(cpt + 1, obj)
   Exit For
End If
Next cpt
End Function
Sub Creer_Tableau()
Dim cnt, i, colch, colim, j, cptCH, cptIM As Integer
Dim TPCH, TPIM, TPLI, car, rangelig As String
TPCH = "CH"
TPIM = "IM"
TPLI = "L"
cptCH = 0
cptIM = 0
colch = 5
colim = 5
MsgBox ("Cliquez sur OK pour lancer la créations des tableaux, puis patientez")
affMessage "Création des tableaux de répartition."
For cnt = 1 To nbREFClient
    If Extrait(REFClient(cnt), 1) = "CH" Then
        NBLPFCH = NBLPFCH + 1
        Worksheets("Repar. CH").Cells(NBLPFCH + 4, 1).Value = Extrait(REFClient(cnt), 2)
        Worksheets("Repar. CH").Cells(NBLPFCH + 4, 2).Value = Extrait(REFClient(cnt), 3)
        Worksheets("Repar. CH").Cells(NBLPFCH + 4, 3).Value = ACTClient(cnt)
    End If
    If Extrait(REFClient(cnt), 1) = "IM" Then
        NBLPFIM = NBLPFIM + 1
        Worksheets("Repar. IM").Cells(NBLPFIM + 4, 1).Value = Extrait(REFClient(cnt), 2)
        Worksheets("Repar. IM").Cells(NBLPFIM + 4, 2).Value = Extrait(REFClient(cnt), 3)
        Worksheets("Repar. IM").Cells(NBLPFIM + 4, 3).Value = ACTClient(cnt)
    End If
Next cnt
Select Case nbCH
    Case 1
        Worksheets("Repar. CH").Columns("F:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("P:X").ColumnWidth = 0
    Case 2
        Worksheets("Repar. CH").Columns("G:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("Q:X").ColumnWidth = 0
    Case 3
        Worksheets("Repar. CH").Columns("H:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("R:X").ColumnWidth = 0
    Case 4
        Worksheets("Repar. CH").Columns("I:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("S:X").ColumnWidth = 0
    Case 5
        Worksheets("Repar. CH").Columns("J:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("T:X").ColumnWidth = 0
    Case 6
        Worksheets("Repar. CH").Columns("K:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("U:X").ColumnWidth = 0
    Case 7
        Worksheets("Repar. CH").Columns("L:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("V:X").ColumnWidth = 0
    Case 8
        Worksheets("Repar. CH").Columns("M:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("W:X").ColumnWidth = 0
    Case 9
        Worksheets("Repar. CH").Columns("N:N").ColumnWidth = 0
        Worksheets("Repar. CH").Columns("X:X").ColumnWidth = 0
End Select
Select Case nbIM
    Case 1
        Worksheets("Repar. IM").Columns("F:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("P:X").ColumnWidth = 0
    Case 2
        Worksheets("Repar. IM").Columns("G:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("Q:X").ColumnWidth = 0
    Case 3
        Worksheets("Repar. IM").Columns("H:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("R:X").ColumnWidth = 0
    Case 4
        Worksheets("Repar. IM").Columns("I:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("S:X").ColumnWidth = 0
    Case 5
        Worksheets("Repar. IM").Columns("J:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("T:X").ColumnWidth = 0
    Case 6
        Worksheets("Repar. IM").Columns("K:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("U:X").ColumnWidth = 0
    Case 7
        Worksheets("Repar. IM").Columns("L:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("V:X").ColumnWidth = 0
    Case 8
        Worksheets("Repar. IM").Columns("M:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("W:X").ColumnWidth = 0
    Case 9
        Worksheets("Repar. IM").Columns("N:N").ColumnWidth = 0
        Worksheets("Repar. IM").Columns("X:X").ColumnWidth = 0
End Select
For i = 1 To nbligne
    For j = 1 To Len(config(i))
        car = VBA.Mid$(config(i), j, 1)
        If car = "C" Then
            cptCH = cptCH + 1
            TPLI = "L" + VBA.Str(i)
            TPCH = "CH" + VBA.Str(cptCH)
            Worksheets("Repar. CH").Cells(4, colch).Value = TPLI + "-" + TPCH
            Worksheets("Repar. CH").Cells(4, colch + 10).Value = "QT" + TPLI + "-" + TPCH
            colch = colch + 1
        End If
        If car = "I" Then
            cptIM = cptIM + 1
            TPLI = "L" + VBA.Str(i)
            TPIM = "IM" + VBA.Str(cptIM)
            Worksheets("Repar. IM").Cells(4, colim).Value = TPLI + "-" + TPIM
            Worksheets("Repar. IM").Cells(4, colim + 10).Value = "QT" + TPLI + "-" + TPIM
            colim = colim + 1
            
        End If
    Next j
Next i
Worksheets("Saisie").Cells(4, 3).Value = "Nombre de Ligne(s)"
Worksheets("Saisie").Cells(5, 3).Value = nbligne
Worksheets("Saisie").Cells(6, 3).Value = "Nombre de CH"
Worksheets("Saisie").Cells(7, 3).Value = nbCH
Worksheets("Saisie").Cells(8, 3).Value = "Nombre de IM"
Worksheets("Saisie").Cells(9, 3).Value = nbIM
rangelig = VBA.Str(NBLPFCH + 5) & ":300"
Worksheets("Repar. CH").Rows(rangelig).RowHeight = 0
rangelig = VBA.Str(NBLPFIM + 5) & ":300"
Worksheets("Repar. IM").Rows(rangelig).RowHeight = 0
affMessage ""
TabCree = True
End Sub
Sub Effacer_Tableau()
Worksheets("Repar. CH").Range("E5:N300").Value = ""
Worksheets("Repar. CH").Range("E4:X4").Value = ""
Worksheets("Repar. CH").Range("A5:C300").Value = ""
Worksheets("Repar. IM").Range("E5:N300").Value = ""
Worksheets("Repar. IM").Range("E4:X4").Value = ""
Worksheets("Repar. IM").Range("A5:C300").Value = ""
Worksheets("Repar. IM").Columns("E:X").ColumnWidth = 11
Worksheets("Repar. CH").Columns("E:X").ColumnWidth = 11
Worksheets("Repar. CH").Rows("5:300").RowHeight = 15
Worksheets("Repar. IM").Rows("5:300").RowHeight = 15
Worksheets("Saisie").Range("C4:C9") = ""
End Sub
Sub FamillecocheX(fam As Integer)
Dim valmin As Double
Dim plateforme, tampon As String
Dim r, i, j, cpt, icol, fa As Integer
cpt = 1
icol = MiniActCH()
    Do
    fa = Extrait(Worksheets("Repar. CH").Cells(cpt + 4, 1), 12)
    If fa = fam Then
        Worksheets("Repar. CH").Cells(cpt + 4, icol).Value = "X"
    End If
    cpt = cpt + 1
    If Worksheets("Repar. CH").Cells(cpt + 4, 1) = "" Then
        Exit Do
    End If
    If cpt > 300 Then
        Exit Do
    End If
    Loop
cpt = 1
icol = MiniActIM()
Do
    fa = Extrait(Worksheets("Repar. IM").Cells(cpt + 4, 1), 12)
    If fa = fam Then
        Worksheets("Repar. IM").Cells(cpt + 4, icol).Value = "X"
    End If
        cpt = cpt + 1
    If Worksheets("Repar. IM").Cells(cpt + 4, 1) = "" Then
        Exit Do
    End If
    If cpt > 300 Then
        Exit Do
    End If
Loop
End Sub
Sub IVF()
Dim cpt, i, j, flag As Integer
Dim modu, cxh, cxi, nbf, flag2 As Integer
Dim tot, dv As Double
Dim tp1, tp2, tp3, tp4, tp5, tpx As String
Dim entete As Variant
DLF = 0
cpt = 1
flag = 0
FeuiF = False
tot = 0
modu = 0
nbf = Sheets.Count
entete = Array("SMN", "Nb de tests nécessaire", "Répartition/Instrument")
affMessage "Importation des données via l'onglet 'Fonctionnement'"
For i = 1 To nbf
If Sheets(i).Name = "Fonctionnement" Then
FeuiF = True
End If
Next i
If FeuiF = False Then
MsgBox (" Attention la feuille d'import intitulée : 'Fonctionnement'.' n'éxiste pas. Veuillez la coller dans ce fichier")
Exit Sub
End If
If entete(0) = Worksheets("Fonctionnement").Cells(3, 12).Value And entete(1) = Worksheets("Fonctionnement").Cells(3, 16).Value And entete(2) = Worksheets("Fonctionnement").Cells(3, 6) Then
FeuiF = True
Else
FeuiF = False
End If
If FeuiF = False Then
MsgBox (" Attention la feuille d'import 'Fonctionnement' n'est pas correctement formatée. Veuillez coller une feuille une compatible dans ce fichier")
Exit Sub
End If
CopieFonctionnement
Do
tp1 = Worksheets("Fonctionnement").Cells(cpt + 3, 21)
tp2 = Worksheets("Fonctionnement").Cells(cpt + 3 + 1, 21)
If tp1 = "" And tp2 = "" Then

Exit Do
End If
cpt = cpt + 1

Loop
DLF = cpt + 2
nbREFClient = 1
REFClient(nbREFClient) = Worksheets("Fonctionnement").Cells(4, 12).Value
For i = 5 To DLF
flag = 0
tp4 = Worksheets("Fonctionnement").Cells(i, 12).Value
If tp4 <> "" Then
    For j = 1 To nbREFClient
        If tp4 = REFClient(j) Then
        flag = 1
        End If
    Next j
    If flag = 0 Then
    nbREFClient = nbREFClient + 1
    REFClient(nbREFClient) = tp4
    End If
Else
End If
Next i
For i = 1 To nbREFClient
tot = 0
modu = 0
    For j = 4 To DLF
        If REFClient(i) = Worksheets("Fonctionnement").Cells(j, 12).Value Then
            tot = tot + Worksheets("Fonctionnement").Cells(j, 16).Value
            
        End If
    Next j
ACTClient(i) = tot
nbModREF(i) = NombreREF(REFClient(i))
tp5 = Extrait2(REFClient(i), 8)
If tp5 = "SU" Then
nbModREF(i) = nbModREF(i) - 1
If nbModREF(i) = 0 Then
nbModREF(i) = 1
End If
End If

Next i
For i = 1 To nbREFClient
Worksheets("Saisie").Cells(i + 4, 6).Value = REFClient(i)
Worksheets("Saisie").Cells(i + 4, 7).Value = ACTClient(i)
tp3 = REFClient(i)
Worksheets("Saisie").Cells(i + 4, 5).Value = Extrait2((tp3), 3)
Next i
modifF
End Sub

Sub SMNcocheX(SMN As String)
Dim valmin As Double
Dim i, indexcol, indexlig As Integer
If Extrait(SMN, 1) = "CH" Then
    For i = 1 To nbREFClient
        If Worksheets("Repar. CH").Cells(4 + i, 1).Value = SMN Then
            indexlig = 4 + i
            i = nbREFClient
        End If
    Next i
    indexcol = MiniActCH()
    Worksheets("Repar. CH").Cells(indexlig, indexcol).Value = "X"
End If
If Extrait(SMN, 1) = "IM" Then
    For i = 1 To nbREFClient
        If Worksheets("Repar. IM").Cells(4 + i, 1).Value = SMN Then
            indexlig = 4 + i
            i = nbREFClient
        End If
    Next i
    indexcol = MiniActIM()
    Worksheets("Repar. IM").Cells(indexlig, indexcol).Value = "X"
    End If
End Sub

Sub AnalyseFamille(fm As Integer)
Dim i, j, limit As Integer
Dim opt, tampon As String
opt = Worksheets("Saisie").Cells(37 + fm, 3).Value
If opt = "0" Then
Exit Sub
Else
If FamillePT(fm) = "CH" Then
    limit = nbCH
End If
If FamillePT(fm) = "IM" Then
    limit = nbIM
End If
End If
If val(opt) > limit Then
tampon = Worksheets("Saisie").Cells(37 + fm, 1).Value
tampon = "Le nombre de modules entré pour la famille " & tampon
tampon = tampon & " est supérieur au nombre de modules disponibles. Merci de modifier la saisie"
Exit Sub
Else
For j = 1 To val(opt)
FamillecocheX (fm)
Next j
End If
End Sub
Function MiniActCH() As Integer
Dim a, colonne As Integer
Dim valmin As Double
valmin = Worksheets("Repar. CH").Cells(3, 15).Value
colonne = 5
If nbCH <= 1 Then
    colonne = 5
Else
    For a = 1 To nbCH
         If Worksheets("Repar. CH").Cells(3, 14 + a).Value < valmin Then
         valmin = Worksheets("Repar. CH").Cells(3, 14 + a).Value
         colonne = 4 + a
         End If
    Next a
End If
MiniActCH = colonne
End Function
Function MiniActIM() As Integer
Dim a, colonne As Integer
Dim valmin As Double
valmin = Worksheets("Repar. IM").Cells(3, 15).Value
colonne = 5
If nbIM <= 1 Then
    colonne = 5
Else
    For a = 1 To nbIM
         If Worksheets("Repar. IM").Cells(3, 14 + a).Value < valmin Then
         valmin = Worksheets("Repar. IM").Cells(3, 14 + a).Value
         colonne = 4 + a
         End If
    Next a
End If
MiniActIM = colonne
End Function
Sub Creation()
Dim m1, m2, m3, m4, m5 As String
UserForm1.Show
If ImportSaisie = False Then
IVF
If FeuiF = False Then
Exit Sub
End If
End If

Init
If nbligne = 0 Then
m1 = "Vous devez renseigner la configuration des lignes Atellica !"
MsgBox m1, vbCritical, "Manque Information"
Exit Sub
End If
If nbREFClient = 0 Then
m1 = "Vous devez renseigner le tableau des données du DXCon !"
MsgBox m1, vbCritical, "Manque information"
Exit Sub
End If
Effacer_Tableau
Creer_Tableau
CreerParam
If ImportSaisie = False Then
contrat
End If
MsgBox ("Création des Tableaux Terminée")
Worksheets("Saisie").Cells(12, 9).Value = ""
End Sub
Sub CreerParam()
Dim cpt, LF As Integer
Dim tampon As String
cpt = 1

Do
affMessage "Création des tableaux de repartition des paramètres pour la Chimie." & Str(cpt)
tampon = Worksheets("Para. CH").Cells(cpt + 3, 2)
'Worksheets("Para. CH").Cells(cpt + 3, 4).Value = Extrait2(tampon, 9)
'Worksheets("Para. CH").Cells(cpt + 3, 6).Value = Extrait2(tampon, 10)
If tampon = 0 Then
LF = cpt
Exit Do
End If
cpt = cpt + 1
Loop


cpt = 1
Do
affMessage "Création des tableaux de repartition des paramètres pour l'Immuno." & Str(cpt)
tampon = Worksheets("Para. IM").Cells(cpt + 3, 2).Value
'Worksheets("Para. IM").Cells(cpt + 3, 4).Value = Extrait2(tampon, 9)
'Worksheets("Para. IM").Cells(cpt + 3, 4).Value = Extrait2(tampon, 10)
If tampon = 0 Then
Exit Do
End If
cpt = cpt + 1
Loop
TrierReparCH
TrierReparIM
affMessage ""
End Sub
Sub Analyse()
Dim i, cpt As Integer
Dim REF As String
If TabCree = False Then
MsgBox ("Vous devez créer le tableau avant de lancer l'analyse")
Exit Sub
End If
MsgBox ("Pour lancer l'analyse cliquez sur OK puis patientez")
Worksheets("Repar. CH").Range("E5:M300").Value = ""
Worksheets("Repar. IM").Range("E5:M300").Value = ""
affMessage "Analyse en Cours"
For i = 0 To 26
AnalyseFamille (i)
Next
cpt = 1
Do
If Worksheets("Repar. CH").Cells(cpt + 4, 1).Value = "" Then
Exit Do
End If
AnalyseRef (Worksheets("Repar. CH").Cells(cpt + 4, 1).Value)
cpt = cpt + 1

If cpt > 300 Then
Exit Do
End If
Loop
cpt = 1
Do
If Worksheets("Repar. IM").Cells(cpt + 4, 1).Value = "" Then
Exit Do
End If
AnalyseRef (Worksheets("Repar. IM").Cells(cpt + 4, 1).Value)
cpt = cpt + 1

If cpt > 300 Then
Exit Do
End If
Loop
affMessage ""
MsgBox ("Analyse Terminée")
End Sub
Function LigneRef(SMN As String) As Integer
Dim i As Integer
Dim PT As String
PT = Extrait(SMN, 1)
If PT = "CH" Then
    For i = 1 To NBLPFCH
    If Worksheets("Repar. CH").Cells(4 + i, 1) = SMN Then
    LigneRef = 4 + i
    End If
    Next i
End If
If PT = "IM" Then
    For i = 1 To NBLPFIM
    If Worksheets("Repar. IM").Cells(4 + i, 1) = SMN Then
    LigneRef = 4 + i
    End If
    Next i
End If
End Function
Function NbModulevsActi(activite As Double, plateforme As String) As Integer
Dim i As Integer
If plateforme = "CH" Then
    For i = 1 To 5
        If activite >= InterCH(i, 1) And activite <= InterCH(i, 2) Then
        NbModulevsActi = i
        End If
    Next i
End If
If plateforme = "IM" Then
    For i = 1 To 5
        If activite >= InterIM(i, 1) And activite <= InterIM(i, 2) Then
        NbModulevsActi = i
        End If
    Next i
End If
End Function
Sub AnalyseRef(SMN As String)
Dim nbmod, i As Integer
Dim acti As Double
Dim PT As String
PT = Extrait(SMN, 1)
If PT = "CH" Then
    nbmod = Worksheets("Repar. CH").Cells(LigneRef(SMN), 4)
    acti = Worksheets("Repar. CH").Cells(LigneRef(SMN), 3)
    If nbmod > 0 Then
    Exit Sub
    Else
    For i = 1 To NbModulevsActi(acti, "CH")
    SMNcocheX (SMN)
    Next i
    End If
End If
If PT = "IM" Then
    nbmod = Worksheets("Repar. IM").Cells(LigneRef(SMN), 4)
    acti = Worksheets("Repar. IM").Cells(LigneRef(SMN), 3)
    If nbmod > 0 Then
    Exit Sub
    Else
    For i = 1 To NbModulevsActi(acti, "IM")
    SMNcocheX (SMN)
    Next i
    End If
End If
End Sub
Sub RAZ()
Dim val As Variant
Dim i As Integer
val = MsgBox("Cliquez sur Oui pour effacer toutes les données saisies ", vbYesNo, "Confirmation de RAZ")

If val = "6" Then
affMessage "Effacement des données en cours..."
Effacer_Tableau
Worksheets("Saisie").Range("B4:B13").Value = ""
Worksheets("Saisie").Cells(5, 2).Value = ""
Worksheets("Saisie").Cells(7, 2).Value = ""
Worksheets("Saisie").Cells(9, 2).Value = ""
Worksheets("Saisie").Range("C37:C63").Value = 0
Worksheets("Saisie").Range("E5:G300").Value = ""
Worksheets("Repar. CH").Range("AD5:AD300").Value = ""
Worksheets("Repar. IM").Range("AD5:AD300").Value = ""
Worksheets("Suivi VDM CH").Range("B4:C300").Value = ""
Worksheets("Suivi VDM CH").Range("H4:H300").Value = ""

Worksheets("Suivi VDM CH").Range("B4:BM300").ColumnWidth = 10
Worksheets("Suivi VDM CH").Range("B4:BM300").RowHeight = 15

Worksheets("Suivi VDM IM").Range("B4:C300").Value = ""

Worksheets("Suivi VDM IM").Range("B4:BM300").ColumnWidth = 10
Worksheets("Suivi VDM IM").Range("B4:BM300").RowHeight = 15

Worksheets("Para. CH").Range("A4:I300").RowHeight = 15
Worksheets("Para. IM").Range("A4:I300").RowHeight = 15
TabCree = False
End If
EffacerFonctionnement
affMessage ""
End Sub
Function Extrait2(SMN As String, obj As Integer) As String  'OBJ=1 Plateforme, OBJ=2 SMN, OBJ= 3 Nom OBJ=4 Abreviation, OBJ=4 Type Echantillon Chaine, OBJ=5 Gamme CU, OBJ=6 Gamme SI, OBJ=7 Type r_duit'
Dim cpt As Integer
Dim tampon As String
cpt = 1
Do
tampon = Worksheets("Sources2").Cells(1 + cpt, 2).Value
If tampon = SMN Then
    Extrait2 = Worksheets("Sources2").Cells(1 + cpt, obj).Value
End If
If tampon = "" Then
Exit Do
End If
cpt = cpt + 1
If cpt > 1000 Then
Exit Do
End If
Loop
End Function
Sub DoSheet()
Dim i, ct, j As Integer
Dim jour, nomf, word(4), plan(100) As String
Dim FID, FIF, RED, REF, COD, COF, AUD, AUF As Integer
Dim nbchVDM, nbimVDM As Integer
Dim nbVDMTOT As Integer

nbVDMTOT = 0
nbchVDM = Worksheets("Suivi VDM CH").Cells(2, 7).Value
nbimVDM = Worksheets("Suivi VDM IM").Cells(2, 7).Value

nbVDMTOT = nbchVDM + nbimVDM
jour = "J"
NBJourVDM = Worksheets("Saisie").Cells(12, 20).Value
NBJourREPET = Worksheets("Saisie").Cells(20, 21).Value
NBJourFI = Worksheets("Saisie").Cells(21, 21).Value
NBJourCOMP = Worksheets("Saisie").Cells(22, 21).Value
NBJourAUTRE = Worksheets("Saisie").Cells(23, 21).Value
Worksheets("Suivi VDM IM").Select
FID = 1
FIF = NBJourFI
RED = 1
REF = NBJourREPET
COD = REF + 1
COF = COD + (NBJourCOMP - 1)
AUD = COF + 1
AUF = AUD + (NBJourAUTRE - 1)

For ct = 1 To NBJourVDM
If ct >= FID And ct <= FIF Then
word(1) = "O"
Else
word(1) = "N"
End If
If ct >= RED And ct <= REF Then
word(2) = "O"
Else
word(2) = "N"
End If
If ct >= COD And ct <= COF Then
word(3) = "O"
Else
word(3) = "N"
End If
If ct >= AUD And ct <= AUF Then
word(4) = "O"
Else
word(4) = "N"
End If
plan(ct) = word(1) & word(2) & word(3) & word(4)
Next ct
For ct = 1 To NBJourVDM
nomf = jour & ct
Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = nomf
Trame (nomf), plan(ct)
Next ct

End Sub


Sub VDM()
Dim cpt, lig As Integer
Dim tampon As String
cpt = 1
lig = 1
If Worksheets("Saisie").Cells(3, 16).Value <> "OK" Then
MsgBox ("Il manque des données de VDM pour lancer la création des tableaux !")
Exit Sub
If TabCree = False Then
MsgBox ("Veuillez créer les tableaux de répartition avant ceux pour la validation des méthodes")
Exit Sub
End If
End If
If Worksheets("REPAR. CH").Cells(3, 1).Value = 0 Or Worksheets("REPAR. IM").Cells(3, 1).Value = 0 Then
MsgBox ("Veuillez créer les tableaux de répartition avant ceux pour la validation des méthodes")
Exit Sub
End If
If Worksheets("Saisie").Cells(12, 20).Value = 0 Or Worksheets("Saisie").Cells(12, 20).Value = "" Then
MsgBox ("Veuillez saisir le nombre de jours de VDM envisagé avant de créer les tableaux de VDM")
Exit Sub
End If
affMessage "Création des tableaux de validation des méthodes (Etape 1)"

Do
tampon = Worksheets("Repar. CH").Cells(lig + 4, 1)
If tampon = "" Then
Exit Do
End If
Select Case Extrait2(tampon, 8)
    Case "S"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Sérum"
    Case "U"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Urine"
    Case "P"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Plasma"
    Case "W"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Sang Total"
    Case "UC"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Urine"
    cpt = cpt + 1
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "LCR"
    Case "SU"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Sérum"
    cpt = cpt + 1
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Urine"
    Case "SUC"
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Sérum"
    cpt = cpt + 1
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "Urine"
    cpt = cpt + 1
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM CH").Cells(cpt + 3, 8).Value = "LCR"
    
End Select
cpt = cpt + 1
lig = lig + 1
Loop
cpt = 1
lig = 1
Do
tampon = Worksheets("Repar. IM").Cells(lig + 4, 1).Value
If tampon = "" Then
Exit Do
End If
Select Case Extrait2(tampon, 8)
    Case "S"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Sérum"
    Case "U"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Urine"
    Case "P"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Plasma"
    Case "W"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Sang Total"
    Case "UC"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Urine"
    cpt = cpt + 1
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "LCR"
    Case "SU"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Sérum"
    cpt = cpt + 1
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Urine"
    Case "SUC"
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Sérum"
    cpt = cpt + 1
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "Urine"
    cpt = cpt + 1
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 2).Value = Extrait2(tampon, 2)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 3).Value = Extrait2(tampon, 3)
    Worksheets("Suivi VDM IM").Cells(cpt + 3, 8).Value = "LCR"
End Select
cpt = cpt + 1
lig = lig + 1
Loop
afficheSuiv

DoSheet
RemplirJ

End Sub

Sub afficheSuiv()
Dim i, j, k, l, colm, coln, lastCH, lastIM, nbl As Integer
Dim nbniveau, nbmodule As Integer
Dim module, nive As String
colm = 9
coln = 9
module = "CH"
niveau = "Niveau"
affMessage "Création des tableaux de validation des méthodes (Etape 2)"
Worksheets("Suivi VDM CH").Range("I2:Z3") = ""
nbniveau = Worksheets("Saisie").Cells(5, 17).Value
nbmodule = Worksheets("Saisie").Cells(7, 3).Value
nbl = Worksheets("Suivi VDM CH").Cells(2, 7).Value
lastCH = nbmodule * nbniveau
For i = 1 To nbmodule
    Worksheets("Suivi VDM CH").Cells(2, colm).Value = module & Str(i)
    For j = 1 To nbniveau
    Worksheets("Suivi VDM CH").Cells(3, coln).Value = niveau & Str(j)
    For k = 1 To nbl
        Worksheets("Suivi VDM CH").Cells(3 + k, coln).Value = 0
    Next k
    coln = coln + 1
    colm = colm + 1
    Next j
Next i
remplir "Suivi VDM CH", 4, 45, 4 + (nbl - 1), 45 + (nbCH - 1), "0"
remplir "Suivi VDM CH", 4, 56, 4 + (nbl - 1), 56 + (nbCH - 1), "0"
For i = lastCH + 1 To 30
Worksheets("Suivi VDM CH").Range(Worksheets("Suivi VDM CH").Cells(4, 8 + i), Worksheets("Suivi VDM CH").Cells(300, 8 + i)).Value = ""
Worksheets("Suivi VDM CH").Columns(8 + i).ColumnWidth = 0
Next i
For l = nbCH + 1 To 10
Worksheets("Suivi VDM CH").Columns(44 + l).ColumnWidth = 0
Worksheets("Suivi VDM CH").Columns(55 + l).ColumnWidth = 0
Worksheets("Suivi VDM CH").Range(Worksheets("Suivi VDM CH").Cells(4, 44 + l), Worksheets("Suivi VDM CH").Cells(300, 45 + l)).Value = ""
Worksheets("Suivi VDM CH").Range(Worksheets("Suivi VDM CH").Cells(4, 55 + l), Worksheets("Suivi VDM CH").Cells(300, 56 + l)).Value = ""
Next l
Worksheets("Suivi VDM CH").Columns(3).ColumnWidth = 55

colm = 9
coln = 9
module = "IM"
niveau = "Niveau"
Worksheets("Suivi VDM IM").Range("I2:Z3") = ""
nbniveau = Worksheets("Saisie").Cells(5, 18).Value
nbmodule = Worksheets("Saisie").Cells(9, 3).Value
nbl = Worksheets("Suivi VDM IM").Cells(2, 7).Value
lastIM = nbmodule * nbniveau
For i = 1 To nbmodule
    Worksheets("Suivi VDM IM").Cells(2, colm).Value = module & Str(i)
    For j = 1 To nbniveau
    Worksheets("Suivi VDM IM").Cells(3, coln).Value = niveau & Str(j)
    For k = 1 To nbl
        Worksheets("Suivi VDM IM").Cells(3 + k, coln).Value = 0
    Next k
    coln = coln + 1
    colm = colm + 1
    Next j
Next i
remplir "Suivi VDM IM", 4, 45, 4 + (nbl - 1), 45 + (nbIM - 1), "0"
remplir "Suivi VDM IM", 4, 56, 4 + (nbl - 1), 56 + (nbIM - 1), "0"
For i = lastIM + 1 To 30
Worksheets("Suivi VDM IM").Range(Worksheets("Suivi VDM IM").Cells(4, 8 + i), Worksheets("Suivi VDM IM").Cells(300, 8 + i)).Value = ""
Worksheets("Suivi VDM IM").Columns(8 + i).ColumnWidth = 0
Next i
For l = nbIM + 1 To 10
Worksheets("Suivi VDM IM").Columns(44 + l).ColumnWidth = 0
Worksheets("Suivi VDM IM").Columns(55 + l).ColumnWidth = 0
Worksheets("Suivi VDM IM").Range(Worksheets("Suivi VDM IM").Cells(4, 44 + l), Worksheets("Suivi VDM IM").Cells(300, 45 + l)).Value = ""
Worksheets("Suivi VDM IM").Range(Worksheets("Suivi VDM IM").Cells(4, 55 + l), Worksheets("Suivi VDM IM").Cells(300, 56 + l)).Value = ""
Next l
Worksheets("Suivi VDM IM").Columns(3).ColumnWidth = 55
affMessage ""
End Sub
Sub modifF() 'Modification de la feuille fonctionnement'
Dim tp, tp2 As String
Dim i As Integer
tp2 = "0"
For i = 1 To 1000
If Worksheets("Fonctionnement").Cells(i + 3, 20) = "" Then
Worksheets("Fonctionnement").Cells(i + 3, 20) = "NA"
Worksheets("Fonctionnement").Cells(i + 3, 21) = "NA"
Worksheets("Fonctionnement").Cells(i + 3, 19) = "NA"
End If
Next i
For i = 1 To 1000
tp = Worksheets("Fonctionnement").Cells(i + 3, 12)
If tp = "" Then
Worksheets("Fonctionnement").Cells(i + 3, 12).Value = tp2
Else
tp2 = tp
Worksheets("Fonctionnement").Cells(i + 3, 12).Value = tp2
End If
Next i
End Sub
Function QCSMN(SMN As String) As String 'Retourne une chaine des QC associés à un SMN'
Dim NomQC(100), SMNQC(100), chaine, tp, tp2 As String
Dim nbQC, cpt, i As Integer
chaine = ""
nbQC = 0
cpt = 1
Do
SMNQC(cpt) = Worksheets("QC").Cells(1 + cpt, 3)
NomQC(cpt) = Worksheets("QC").Cells(1 + cpt, 4)

If SMNQC(cpt) = "" Then
Exit Do
End If
cpt = cpt + 1
Loop
nbQC = cpt
cpt = 1
Do
tp = Worksheets("Fonctionnement").Cells(3 + cpt, 20)
tp2 = Worksheets("Fonctionnement").Cells(3 + cpt, 12)
If tp = "" Then
Exit Do
End If
If tp2 = SMN Then
For i = 1 To nbQC
If tp = SMNQC(i) Then
If InStr(chaine, NomQC(i)) = 0 Then
chaine = chaine & "  " & NomQC(i)
End If
End If
Next i
End If
cpt = cpt + 1
Loop
QCSMN = chaine
End Function
Sub Trame(Feuille As String, typ As String) 'OOOO = FI/REPET/COMP/AUTRES  ONNN=FI/NON/NON/NON
Dim lig As Integer
Dim nbdosCH, nbdosIM As Integer
Dim msg As String
lig = 1
nbdosCH = Worksheets("Repar. CH").Cells(3, 1)
nbdosIM = Worksheets("Repar. IM").Cells(3, 1)
Worksheets(Feuille).Columns(2).ColumnWidth = 14
Worksheets(Feuille).Columns(3).ColumnWidth = 50
Worksheets(Feuille).Columns(4).ColumnWidth = 13
Worksheets(Feuille).Columns(5).ColumnWidth = 55
Worksheets(Feuille).Columns(6).ColumnWidth = 14
Worksheets(Feuille).Columns(7).ColumnWidth = 35
Worksheets(Feuille).Cells(1, 1).Value = Feuille
Worksheets(Feuille).Range("B1:I1").Select
Selection.Merge
With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Bold = True
        
End With
Worksheets(Feuille).Cells(1, 2).Value = "Opérations à Réaliser ce jour"
Worksheets(Feuille).Cells(1, 1).Select
With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Bold = True
        
End With
With Selection.Interior
    .Color = 5296274
End With
msg = "Veuillez lancer les contrôles de qualité pour les " & Str(nbdosCH) & "  dosages de Chimie et les " & Str(nbdosIM) & " dosages d'Immunologie sur chacun des modules"
If Mid(typ, 1, 1) = "O" Then
Worksheets(Feuille).Cells(lig + 1, 2).Value = "Fidélité Intermédiaire"
Worksheets(Feuille).Cells(lig + 1, 2).Select
With Selection.Font
        .Name = "Calibri"
        .Size = 16
    End With
    Range(Worksheets(Feuille).Cells(lig + 1, 2), Worksheets(Feuille).Cells(lig + 1, 10)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
    End With
    Range(Worksheets(Feuille).Cells(lig + 4, 2), Worksheets(Feuille).Cells(lig + 7, 10)).Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
    .Name = "Calibri"
    .Size = 14
    End With
    Worksheets(Feuille).Cells(lig + 4, 2).Value = msg
lig = lig + 10
End If
    
    
If Mid(typ, 2, 1) = "O" Then
Worksheets(Feuille).Cells(lig + 1, 2).Value = "Répétabilité"
Worksheets(Feuille).Cells(lig + 2, 2).Value = "Référence"
Worksheets(Feuille).Cells(lig + 2, 3).Value = "Dosage"
Worksheets(Feuille).Cells(lig + 2, 4).Value = "Echantillon"
Worksheets(Feuille).Cells(lig + 2, 5).Value = "Contrôles Utilisables"
Worksheets(Feuille).Cells(lig + 2, 6).Value = "Nombre de répliquats"
Worksheets(Feuille).Cells(lig + 2, 7).Value = "Volume nécessaire par niveau en µl (Hors VM)"

Worksheets(Feuille).Cells(lig + 1, 2).Select
With Selection.Font
        .Name = "Calibri"
        .Size = 16
    End With
    Range(Worksheets(Feuille).Cells(lig + 1, 2), Worksheets(Feuille).Cells(lig + 1, 10)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
    End With
    lig = lig + 23
End If
If Mid(typ, 3, 1) = "O" Then
Worksheets(Feuille).Cells(lig + 1, 2).Value = "Comparaison"
Worksheets(Feuille).Cells(lig + 2, 2).Value = "Référence"
Worksheets(Feuille).Cells(lig + 2, 3).Value = "Dosage"
Worksheets(Feuille).Cells(lig + 2, 4).Value = "Prélèvement"
Worksheets(Feuille).Cells(lig + 2, 5).Value = "Nombre d'échantillons"
Worksheets(Feuille).Columns(5).ColumnWidth = 26
Worksheets(Feuille).Cells(lig + 2, 6).Value = "Volume nécessaire par échantillon en µl (Hors VM)"
Worksheets(Feuille).Columns(6).ColumnWidth = 41
Worksheets(Feuille).Columns(7).ColumnWidth = 18
Worksheets(Feuille).Cells(lig + 2, 7).Value = "Nombre de modules"
Worksheets(Feuille).Columns(8).ColumnWidth = 41
Worksheets(Feuille).Cells(lig + 2, 8).Value = "Volume Total nécessaire en µl (Hors VM)"
Worksheets(Feuille).Cells(lig + 1, 2).Select
With Selection.Font
        .Name = "Calibri"
        .Size = 16
    End With
    Range(Worksheets(Feuille).Cells(lig + 1, 2), Worksheets(Feuille).Cells(lig + 1, 10)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    lig = lig + 23
End If
If Mid(typ, 4, 1) = "O" Then
Worksheets(Feuille).Cells(lig + 1, 2).Value = "Autre(s) opérations"
Worksheets(Feuille).Columns(2).ColumnWidth = 55
Worksheets("Saisie").Select
Range("P21:P24").Select
Selection.Copy
Worksheets(Feuille).Select
Range(Worksheets(Feuille).Cells(lig + 2, 2), Worksheets(Feuille).Cells(lig + 2, 2)).Select
ActiveSheet.Paste
Worksheets(Feuille).Select
Range(Worksheets(Feuille).Cells(lig + 1, 2), Worksheets(Feuille).Cells(lig + 1, 2)).Select
With Selection.Font
        .Name = "Calibri"
        .Size = 16
    End With
    Range(Worksheets(Feuille).Cells(lig + 1, 2), Worksheets(Feuille).Cells(lig + 1, 15)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    lig = lig + 23
End If
End Sub
Sub remplir(Feuille As String, l1 As Integer, c1 As Integer, l2 As Integer, c2 As Integer, chaine As String)
Worksheets(Feuille).Range(Worksheets(Feuille).Cells(l1, c1), Worksheets(Feuille).Cells(l2, c2)).Value = chaine
End Sub

Function volQCCH(Feuille As String, LD As Integer, LF As Integer, colonne As Integer) As Double
Dim vf, inter, rest As Double
Dim i, j, cpt As Integer
Dim Vref, v2, tampon As String
vr = 0
vf = 0
For i = LD To LF
Vref = Worksheets(Feuille).Cells(i, colonne)
For cpt = 1 To 1000
tampon = Worksheets("Sources").Cells(cpt + 1, 2)
tp2 = Worksheets("Sources").Cells(cpt + 1, 1)
If tampon = Vref And tp2 = "CH" Then
   vr = vr + val(Worksheets("Sources").Cells(cpt + 1, 10).Value)
   Exit For
End If
Next cpt
Next i
If vr > 0 Then
inter = vr / 250
If inter <= 1 Then
vf = 50
Else
If inter <> Int(inter) Then
vf = vf + (Int(inter) * 50) + 50
Else
vf = vf + (Int(inter) * 50)
End If
End If
volQCCH = vf + 150
Else
volQCCH = 0
End If
End Function
Function volQCIM(Feuille As String, LD As Integer, LF As Integer, colonne As Integer) As Integer
Dim vf As Double
Dim i, j, cpt As Integer
Dim Vref, v2, tampon, ActivFeuille, tp2 As String

vf = 0
For i = LD To LF
Vref = Worksheets(Feuille).Cells(i, colonne)
For cpt = 1 To 1000
tampon = Worksheets("Sources").Cells(cpt + 1, 2)
tp2 = Worksheets("Sources").Cells(cpt + 1, 1)
If tampon = Vref And tp2 = "IM" Then
   vf = vf + val(Worksheets("Sources").Cells(cpt + 1, 10).Value) + 15
   Exit For
End If
Next cpt
Next i
If vf > 0 Then
vf = vf + 150
volQCIM = vf
Else
volQCIM = 0
End If
End Function
Sub RemplirJ()
Dim DataREF(1000), DataSMN(1000), DataPrel(1000), ActivFeuille As String
Dim nbDATACH, nbDATAIM, cpt, STEP_REPET, STEP_COMPA, jour, incr, incr2, nbligne, cpt2, vt As Integer
nbDATACH = 0
nbDATEIM = 0
cpt = 1
incr = 1
incr2 = 1
STEP_REPET = Worksheets("Saisie").Cells(20, 22).Value
STEP_COMPA = Worksheets("Saisie").Cells(22, 22).Value
Do
DataREF(cpt) = Worksheets("Suivi VDM CH").Cells(3 + cpt, 2)
DataSMN(cpt) = Worksheets("Suivi VDM CH").Cells(3 + cpt, 3)
DataPrel(cpt) = Worksheets("Suivi VDM CH").Cells(3 + cpt, 8)
If DataSMN(cpt) = "" Then
cpt = cpt - 1
nbDATACH = cpt
nbligne = cpt
Exit Do
End If
cpt = cpt + 1
Loop
cpt2 = 1
Do
DataREF(cpt) = Worksheets("Suivi VDM IM").Cells(3 + cpt2, 2)
DataSMN(cpt) = Worksheets("Suivi VDM IM").Cells(3 + cpt2, 3)
DataPrel(cpt) = Worksheets("Suivi VDM IM").Cells(3 + cpt2, 8)
If DataSMN(cpt) = "" Then
cpt = cpt - 1
nbligne = cpt
Exit Do
End If
cpt = cpt + 1
cpt2 = cpt2 + 1
Loop
jour = 1
For incr = 1 To nbligne 'Repet'

        ActivFeuille = "J" & jour
        Worksheets(ActivFeuille).Cells(13 + incr2, 2).Value = DataREF(incr)
        Worksheets(ActivFeuille).Cells(13 + incr2, 3).Value = DataSMN(incr)
        Worksheets(ActivFeuille).Cells(13 + incr2, 4).Value = DataPrel(incr)
        Worksheets(ActivFeuille).Cells(13 + incr2, 5).Value = QCSMN(Worksheets(ActivFeuille).Cells(13 + incr2, 2).Value)
        Worksheets(ActivFeuille).Cells(13 + incr2, 6).Value = Worksheets("Saisie").Cells(6, 17).Value
        Worksheets(ActivFeuille).Cells(13 + incr2, 7).Value = QTQCREPET((DataREF(incr)), Worksheets("Saisie").Cells(6, 17).Value)
        incr2 = incr2 + 1
    
    If incr2 > STEP_REPET Then
    incr2 = 1
    jour = jour + 1

    End If

Next incr
jour = Worksheets("Saisie").Cells(20, 21).Value + 1
incr2 = 1
For incr = 1 To nbligne 'Comparaison'

        ActivFeuille = "J" & jour
        Worksheets(ActivFeuille).Cells(13 + incr2, 2).Value = DataREF(incr)
        Worksheets(ActivFeuille).Cells(13 + incr2, 3).Value = DataSMN(incr)
        Worksheets(ActivFeuille).Cells(13 + incr2, 4).Value = DataPrel(incr)
        Worksheets(ActivFeuille).Cells(13 + incr2, 5).Value = Worksheets("Saisie").Cells(16, 17).Value
        If Extract((DataREF(incr)), 1) = "Atellica® CH" Then
        Worksheets(ActivFeuille).Cells(13 + incr2, 6).Value = QTQCREPET((DataREF(incr)), 1)
        Worksheets(ActivFeuille).Cells(13 + incr2, 7).Value = Worksheets("Saisie").Cells(17, 17).Value
        vt = Worksheets(ActivFeuille).Cells(13 + incr2, 6).Value * Worksheets(ActivFeuille).Cells(13 + incr2, 7).Value
        Else
        Worksheets(ActivFeuille).Cells(13 + incr2, 6).Value = Extract((DataREF(incr)), 21) + 15
        Worksheets(ActivFeuille).Cells(13 + incr2, 7).Value = Worksheets("Saisie").Cells(17, 18).Value
        vt = Worksheets(ActivFeuille).Cells(13 + incr2, 6).Value * Worksheets(ActivFeuille).Cells(13 + incr2, 7).Value
        End If
        Worksheets(ActivFeuille).Cells(13 + incr2, 8).Value = vt
        incr2 = incr2 + 1
    
    If incr2 > STEP_COMPA Then
    incr2 = 1
    jour = jour + 1

    End If

Next incr

End Sub
Sub affMessage(texte As String)
Worksheets("Saisie").Cells(30, 9).Value = texte
End Sub
Function QTQCREPET(SMN As String, nbrep As Integer) As Double
Dim vol, rep As Double
vol = 0
vol = val(Extract(SMN, 21))
If Extract(SMN, 1) = "Atellica® CH" Then
vol = vol * nbrep
rep = vol / 250
If rep > Int(rep) Then
vol = (Int(rep) + 1) * 50
Else
vol = Int(rep) * 50
End If
Else
vol = (vol + 15) * nbrep
End If
QTQCREPET = vol
End Function
Function Extract(SMN As String, obj As Integer) As Variant
Dim i As Integer
i = 3
Do
If Worksheets("Dosages").Cells(i, 2) = SMN Then
    Extract = Worksheets("Dosages").Cells(i, obj).Value
    Exit Do
End If
i = i + 1
Loop
End Function
Sub Auto_Open()


ActiveWindow.Zoom = 50

End Sub
Sub CopieFonctionnement()
Dim i As Integer
Dim okf As Boolean
okf = False
For i = 1 To Sheets.Count
If Sheets(i).Name = "Fonctionnement" Then
okf = True
End If
Next i
If okf = False Then
Exit Sub
End If
    Sheets("Fonctionnement").Select
    Sheets("Fonctionnement").Copy Before:=Sheets(1)
    Sheets("Fonctionnement (2)").Select
    Sheets("Fonctionnement (2)").Name = "Fonctionnement backup"
    Sheets("Fonctionnement backup").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Saisie").Select
End Sub
Sub EffacerFeuille(nom As String)
    Sheets(nom).Select
    ActiveWindow.SelectedSheets.Delete
End Sub
Sub CacherFeuille(nom As String)
    Sheets(nom).Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
Sub MontrerFeuille(nom As String)
    Sheets(nom).Select
    ActiveWindow.SelectedSheets.Visible = True
End Sub
Sub EffacerFonctionnement()
Dim okf As Boolean
okf = False
For i = 1 To Sheets.Count
If Sheets(i).Name = "Fonctionnement" Then
okf = True
End If
Next i
If okf = False Then
Exit Sub
Else
Sheets("Fonctionnement").Select
ActiveWindow.SelectedSheets.Delete
Sheets("Fonctionnement backup").Visible = True
Sheets("Fonctionnement backup").Select
Sheets("Fonctionnement backup").Name = "Fonctionnement"
Sheets("Fonctionnement").Visible = True
Sheets("Saisie").Select
End If

End Sub
Sub TrierReparCH()

    ActiveWorkbook.Worksheets("Repar. CH").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Repar. CH").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("B4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Repar. CH").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub TrierReparIM()

    ActiveWorkbook.Worksheets("Repar. IM").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Repar. IM").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("B4"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Repar. IM").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


