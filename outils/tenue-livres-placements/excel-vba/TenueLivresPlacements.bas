Attribute VB_Name = "TenueLivresPlacements"
Option Explicit
' ============================================================================
' TenueLivresPlacements : placements américains d'un contribuable canadien
'
' Macro autonome (aucune installation). À importer dans le gabarit de saisie
' (Alt+F11, Fichier, Importer un fichier), puis enregistrer en .xlsm.
'
'   Alt+F8  ->  Calculer            : produit les onglets de sortie
'   Alt+F8  ->  TestsAutomatiques   : vérifie le moteur avec des cas connus
'
' Seule communication externe : les taux de la Banque du Canada (dates seulement).
' Si le réseau est bloqué, coller les taux dans l'onglet Taux (date | FXUSDCAD | valeur).
' ============================================================================

Private Const ENCAISSE As String = "ENCAISSE"
Private Const FENETRE_PA As Long = 30
Private Const SEUIL_RETENUE As Double = 0.155
Private Const ECART_TAUX_MAX As Long = 7
Private Const TOLERANCE As Double = 0.01
Private Const CAPACITE As Long = 20000
Private Const FEUILLE_TAUX As String = "Taux"
Private Const NIV_ERREUR As String = "ERREUR"
Private Const NIV_AVERT As String = "AVERTISSEMENT"
Private Const NIV_INFO As String = "INFO"
Private Const FMT_MONTANT As String = "#,##0.00;[Red]-#,##0.00"
Private Const FMT_QTE As String = "#,##0.####"
Private Const FMT_TAUX As String = "0.0000"
Private Const FMT_DATE As String = "yyyy-mm-dd"

Private Type TTx
    compte As String
    typ As String
    titre As String
    dateTx As Date
    dateRegl As Date
    quantite As Double
    prixUsd As Double
    brutUsd As Double
    brutDonne As Boolean
    commission As Double
    retenue As Double
    contrepartie As String
    pbrCadOuv As Double
    pbrCadDonne As Boolean
    description As String
    note As String
    origine As String
    ordre As Long
    dEff As Date
    valide As Boolean
End Type

Private Type TCompte
    code As String
    courtier As String
    numero As String
    description As String
    encaisse As Double
    encaisseFin As Double
End Type

Private Type TDisp
    d As Date
    compte As String
    titre As String
    quantite As Double
    prix As Double
    commission As Double
    produitUsd As Double
    taux As Double
    produitCad As Double
    pbrUsd As Double
    pbrCad As Double
    gainUsd As Double
    gainCadAvant As Double
    refusee As Double
    gainCad As Double
    note As String
    origine As String
End Type

Private Type TRev
    d As Date
    compte As String
    titre As String
    categorie As String
    brutUsd As Double
    retenueUsd As Double
    taux As Double
    brutCad As Double
    retenueCad As Double
    note As String
    origine As String
End Type

Private Type TJour
    d As Date
    compte As String
    typ As String
    titre As String
    quantite As Double
    prix As Double
    brutUsd As Double
    commission As Double
    retenue As Double
    taux As Double
    aTaux As Boolean
    montantCad As Double
    aCad As Boolean
    poolQte As Double
    poolCad As Double
    poolMoyen As Double
    poolUsd As Double
    encaisse As Double
    note As String
    origine As String
End Type

Private Type TAlerte
    niveau As String
    origine As String
    message As String
End Type

Private Type TReleve
    compte As String
    dFin As Date
    titre As String
    valeur As Double
    origine As String
End Type

Private Type T1099
    compte As String
    annee As Long
    div As Double
    ret As Double
    interets As Double
    prod As Double
    origine As String
End Type

Private Type TRapp
    compte As String
    dFin As Date
    annee As Long
    element As String
    decl As Double
    calcule As Double
    ecart As Double
    statut As String
End Type

Private Type TT1135
    element As String
    description As String
    coutMax As Double
    coutFin As Double
    revenu As Double
    gain As Double
End Type

' ---- état du module ----------------------------------------------------------
Private tx() As TTx, nTx As Long
Private ordreTx() As Long, nOrdre As Long
Private cptes() As TCompte, nCptes As Long
Private idxCompte As Object
Private qteCompte As Object
Private qteFin As Object
Private poolTitre() As String, poolQte() As Double, poolCad() As Double, poolUsd() As Double, nPools As Long
Private poolFinQte() As Double, poolFinCad() As Double, poolFinUsd() As Double
Private idxPool As Object
Private disps() As TDisp, nDisps As Long
Private revs() As TRev, nRevs As Long
Private jour() As TJour, nJour As Long
Private alertes() As TAlerte, nAlertes As Long
Private releves() As TReleve, nReleves As Long
Private l1099() As T1099, n1099 As Long
Private rappRel() As TRapp, nRappRel As Long
Private rapp1099() As TRapp, nRapp1099 As Long
Private t1135() As TT1135, nT1135 As Long
Private maxCout As Object, maxEnc As Object, revTitre As Object, gainTitre As Object, intCompte As Object, descr As Object
Private txDates() As Date, txVals() As Double, nTaux As Long
Private txAnnees() As Long, txAnnVals() As Double, nTauxAnn As Long
Private tauxUtilD() As Date, tauxUtilS() As Date, tauxUtilV() As Double, nTauxUtil As Long
Private tauxUtilVus As Object
Private pClient As String, pAnnee As Long, pConvention As String, pMethode As String
Private tauxFinD As Date, tauxFinV As Double, aTauxFin As Boolean, tauxMoyenUtil As Double, aTauxMoyen As Boolean
Private aCloture As Boolean

' ============================================================================
' Points d'entrée
' ============================================================================

Public Sub Calculer()
    Dim debut As Date, fin As Date, i As Long, nErr As Long, nAv As Long, msg As String
    On Error GoTo erreur
    Application.ScreenUpdating = False
    Reinit
    LireParametres
    LireComptes
    LireTransactions
    LireReleves
    Lire1099
    If nTx = 0 Then
        Application.ScreenUpdating = True
        MsgBox "Aucune opération trouvée dans les onglets de comptes.", vbExclamation
        Exit Sub
    End If
    debut = DateSerial(2100, 1, 1): fin = DateSerial(pAnnee, 12, 31)
    For i = 1 To nTx
        If tx(i).dateTx < debut Then debut = tx(i).dateTx
        If tx(i).dateRegl > 0 And tx(i).dateRegl < debut Then debut = tx(i).dateRegl
        If tx(i).dateTx > fin Then fin = tx(i).dateTx
        If tx(i).dateRegl > fin Then fin = tx(i).dateRegl
    Next
    debut = debut - 45
    If fin > Date Then fin = Date
    AssurerTaux debut, fin
    Moteur
    EcrireSortie
    Application.ScreenUpdating = True
    For i = 1 To nAlertes
        If alertes(i).niveau = NIV_ERREUR Then nErr = nErr + 1
        If alertes(i).niveau = NIV_AVERT Then nAv = nAv + 1
    Next
    msg = "Terminé : " & nJour & " opérations traitées, " & nErr & " erreur(s), " & nAv & " avertissement(s)." & vbCrLf & vbCrLf & _
          "Voir l'onglet Alertes en premier, puis Sommaire."
    ThisWorkbook.Worksheets("Sommaire").Activate
    MsgBox msg, IIf(nErr > 0, vbExclamation, vbInformation), "Tenue de livres de placements"
    Exit Sub
erreur:
    Application.ScreenUpdating = True
    MsgBox "Erreur inattendue : " & Err.Description & vbCrLf & "(" & Err.Source & ")", vbCritical
End Sub

' ============================================================================
' Initialisation
' ============================================================================

Private Sub Reinit()
    ReDim tx(1 To CAPACITE): nTx = 0
    ReDim ordreTx(1 To CAPACITE): nOrdre = 0
    ReDim cptes(1 To 200): nCptes = 0
    ReDim poolTitre(1 To 2000): ReDim poolQte(1 To 2000): ReDim poolCad(1 To 2000): ReDim poolUsd(1 To 2000): nPools = 0
    ReDim poolFinQte(1 To 2000): ReDim poolFinCad(1 To 2000): ReDim poolFinUsd(1 To 2000)
    ReDim disps(1 To CAPACITE): nDisps = 0
    ReDim revs(1 To CAPACITE): nRevs = 0
    ReDim jour(1 To CAPACITE): nJour = 0
    ReDim alertes(1 To CAPACITE): nAlertes = 0
    ReDim releves(1 To CAPACITE): nReleves = 0
    ReDim l1099(1 To 2000): n1099 = 0
    ReDim rappRel(1 To CAPACITE): nRappRel = 0
    ReDim rapp1099(1 To 8000): nRapp1099 = 0
    ReDim t1135(1 To 4000): nT1135 = 0
    ReDim txDates(1 To 1): ReDim txVals(1 To 1): nTaux = 0
    ReDim txAnnees(1 To 200): ReDim txAnnVals(1 To 200): nTauxAnn = 0
    ReDim tauxUtilD(1 To CAPACITE): ReDim tauxUtilS(1 To CAPACITE): ReDim tauxUtilV(1 To CAPACITE): nTauxUtil = 0
    Set idxCompte = CreateObject("Scripting.Dictionary")
    Set qteCompte = CreateObject("Scripting.Dictionary")
    Set qteFin = CreateObject("Scripting.Dictionary")
    Set idxPool = CreateObject("Scripting.Dictionary")
    Set maxCout = CreateObject("Scripting.Dictionary")
    Set maxEnc = CreateObject("Scripting.Dictionary")
    Set revTitre = CreateObject("Scripting.Dictionary")
    Set gainTitre = CreateObject("Scripting.Dictionary")
    Set intCompte = CreateObject("Scripting.Dictionary")
    Set descr = CreateObject("Scripting.Dictionary")
    Set tauxUtilVus = CreateObject("Scripting.Dictionary")
    pClient = "": pAnnee = Year(Date): pConvention = "TRANSACTION": pMethode = "QUOTIDIEN"
    aTauxFin = False: aTauxMoyen = False: aCloture = False
End Sub

Private Sub Alerte(ByVal niveau As String, ByVal origine As String, ByVal message As String)
    nAlertes = nAlertes + 1
    alertes(nAlertes).niveau = niveau
    alertes(nAlertes).origine = origine
    alertes(nAlertes).message = message
End Sub

Private Function DGet(ByVal dict As Object, ByVal cle As String) As Double
    If dict.Exists(cle) Then DGet = CDbl(dict.Item(cle)) Else DGet = 0
End Function

Private Sub DAdd(ByVal dict As Object, ByVal cle As String, ByVal delta As Double)
    dict.Item(cle) = DGet(dict, cle) + delta
End Sub

Private Sub DMax(ByVal dict As Object, ByVal cle As String, ByVal v As Double)
    If Not dict.Exists(cle) Then
        dict.Item(cle) = v
    ElseIf v > CDbl(dict.Item(cle)) Then
        dict.Item(cle) = v
    End If
End Sub

Private Function AjouterCompte(ByVal code As String, ByVal courtier As String, ByVal numero As String, ByVal description As String) As Long
    nCptes = nCptes + 1
    cptes(nCptes).code = code
    cptes(nCptes).courtier = courtier
    cptes(nCptes).numero = numero
    cptes(nCptes).description = description
    idxCompte.Item(code) = nCptes
    AjouterCompte = nCptes
End Function

Private Function IdxPoolOuCreer(ByVal titre As String) As Long
    If idxPool.Exists(titre) Then
        IdxPoolOuCreer = idxPool.Item(titre)
    Else
        nPools = nPools + 1
        poolTitre(nPools) = titre
        poolQte(nPools) = 0: poolCad(nPools) = 0: poolUsd(nPools) = 0
        idxPool.Item(titre) = nPools
        IdxPoolOuCreer = nPools
    End If
End Function

Private Function PoolMoyenCad(ByVal ip As Long) As Double
    If poolQte(ip) <> 0 Then PoolMoyenCad = poolCad(ip) / poolQte(ip) Else PoolMoyenCad = 0
End Function

Private Function PoolMoyenUsd(ByVal ip As Long) As Double
    If poolQte(ip) <> 0 Then PoolMoyenUsd = poolUsd(ip) / poolQte(ip) Else PoolMoyenUsd = 0
End Function

Private Function CleQte(ByVal code As String, ByVal titre As String) As String
    CleQte = code & "|" & titre
End Function

Private Function TypeValide(ByVal typ As String) As Boolean
    Select Case typ
        Case "ACHAT", "VENTE", "DIVIDENDE", "INTERET", "REMB_CAPITAL", "FRAIS", "FRACTIONNEMENT", _
             "TRANSFERT_TITRES", "TRANSFERT_ENCAISSE", "DEPOT", "RETRAIT", "AJUST_PBR", "SOLDE_OUVERTURE"
            TypeValide = True
        Case Else
            TypeValide = False
    End Select
End Function

' ============================================================================
' Taux de change
' ============================================================================

Private Function IdxTauxPrecedent(ByVal d As Date) As Long
    Dim lo As Long, hi As Long, m As Long
    lo = 1: hi = nTaux: IdxTauxPrecedent = 0
    Do While lo <= hi
        m = (lo + hi) \ 2
        If txDates(m) <= d Then
            IdxTauxPrecedent = m: lo = m + 1
        Else
            hi = m - 1
        End If
    Loop
End Function

Private Sub LogTaux(ByVal d As Date, ByVal dSource As Date, ByVal v As Double)
    Dim cle As String
    cle = CStr(CLng(d))
    If tauxUtilVus.Exists(cle) Then Exit Sub
    tauxUtilVus.Item(cle) = 1
    nTauxUtil = nTauxUtil + 1
    tauxUtilD(nTauxUtil) = d: tauxUtilS(nTauxUtil) = dSource: tauxUtilV(nTauxUtil) = v
End Sub

Private Function TauxAu(ByVal d As Date, ByVal origine As String, ByRef ok As Boolean) As Double
    Dim i As Long
    ok = False
    i = IdxTauxPrecedent(d)
    If i = 0 Then
        Alerte NIV_ERREUR, origine, "Aucun taux publié le " & Format(d, FMT_DATE) & " ou avant (onglet Taux)"
        Exit Function
    End If
    If d - txDates(i) > ECART_TAUX_MAX Then
        Alerte NIV_ERREUR, origine, "Aucun taux publié entre le " & Format(txDates(i), FMT_DATE) & " et le " & Format(d, FMT_DATE)
        Exit Function
    End If
    LogTaux d, txDates(i), txVals(i)
    ok = True
    TauxAu = txVals(i)
End Function

Private Function TauxDernier(ByVal d As Date, ByRef dSource As Date, ByRef ok As Boolean) As Double
    Dim i As Long
    ok = False
    i = IdxTauxPrecedent(d)
    If i = 0 Then Exit Function
    dSource = txDates(i)
    LogTaux d, txDates(i), txVals(i)
    ok = True
    TauxDernier = txVals(i)
End Function

Private Function TauxMoyen(ByVal annee As Long, ByRef ok As Boolean) As Double
    Dim i As Long, s As Double, n As Long
    ok = False
    For i = 1 To nTauxAnn
        If txAnnees(i) = annee Then ok = True: TauxMoyen = txAnnVals(i): Exit Function
    Next
    For i = 1 To nTaux
        If Year(txDates(i)) = annee Then s = s + txVals(i): n = n + 1
    Next
    If n > 0 Then ok = True: TauxMoyen = s / n
End Function

Private Sub AjouterTauxQuotidien(ByVal d As Date, ByVal v As Double)
    nTaux = nTaux + 1
    If nTaux > UBound(txDates) Then
        ReDim Preserve txDates(1 To UBound(txDates) * 2 + 16)
        ReDim Preserve txVals(1 To UBound(txVals) * 2 + 16)
    End If
    txDates(nTaux) = d: txVals(nTaux) = v
End Sub

Private Sub AjouterTauxAnnuel(ByVal annee As Long, ByVal v As Double)
    Dim i As Long
    For i = 1 To nTauxAnn
        If txAnnees(i) = annee Then txAnnVals(i) = v: Exit Sub
    Next
    nTauxAnn = nTauxAnn + 1
    txAnnees(nTauxAnn) = annee: txAnnVals(nTauxAnn) = v
End Sub

Private Sub TrierTaux()
    Dim gap As Long, i As Long, j As Long, td As Date, tv As Double
    gap = nTaux \ 2
    Do While gap > 0
        For i = gap + 1 To nTaux
            td = txDates(i): tv = txVals(i): j = i
            Do While j > gap
                If txDates(j - gap) > td Then
                    txDates(j) = txDates(j - gap): txVals(j) = txVals(j - gap): j = j - gap
                Else
                    Exit Do
                End If
            Loop
            txDates(j) = td: txVals(j) = tv
        Next
        gap = gap \ 2
    Loop
End Sub

Private Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not ws Is Nothing
End Function

Private Sub ChargerTauxDepuisFeuille()
    Dim ws As Worksheet, r As Long, derniere As Long, v As Variant, serie As String, d As Date, ok As Boolean
    nTaux = 0: nTauxAnn = 0
    If Not FeuilleExiste(FEUILLE_TAUX) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(FEUILLE_TAUX)
    derniere = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To derniere
        d = VersDate(ws.Cells(r, 1).Value, ok)
        If ok And IsNumeric(ws.Cells(r, 3).Value) Then
            serie = UCase$(Trim$(CStr(ws.Cells(r, 2).Value)))
            If serie = "FXAUSDCAD" Then
                AjouterTauxAnnuel Year(d), CDbl(ws.Cells(r, 3).Value)
            Else
                AjouterTauxQuotidien d, CDbl(ws.Cells(r, 3).Value)
            End If
        End If
    Next
    TrierTaux
End Sub

Private Sub EcrireFeuilleTaux()
    Dim ws As Worksheet, arr() As Variant, i As Long, n As Long
    If Not FeuilleExiste(FEUILLE_TAUX) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = FEUILLE_TAUX
    Else
        Set ws = ThisWorkbook.Worksheets(FEUILLE_TAUX)
        ws.Cells.Clear
    End If
    ws.Cells(1, 1).Value = "date": ws.Cells(1, 2).Value = "serie": ws.Cells(1, 3).Value = "valeur"
    ws.Cells(1, 5).Value = "Cache des taux de la Banque du Canada. Si le réseau est bloqué, coller ici les lignes du CSV Valet (FXUSDCAD)."
    n = nTaux + nTauxAnn
    If n = 0 Then Exit Sub
    ReDim arr(1 To n, 1 To 3)
    For i = 1 To nTaux
        arr(i, 1) = txDates(i): arr(i, 2) = "FXUSDCAD": arr(i, 3) = txVals(i)
    Next
    For i = 1 To nTauxAnn
        arr(nTaux + i, 1) = DateSerial(txAnnees(i), 12, 31): arr(nTaux + i, 2) = "FXAUSDCAD": arr(nTaux + i, 3) = txAnnVals(i)
    Next
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 3)).Value = arr
    ws.Columns(1).NumberFormat = FMT_DATE
    ws.Columns(3).NumberFormat = FMT_TAUX
    ws.Columns(1).ColumnWidth = 12: ws.Columns(2).ColumnWidth = 12: ws.Columns(3).ColumnWidth = 10
End Sub

Private Function TelechargerCsv(ByVal serie As String, ByVal debut As Date, ByVal fin As Date) As String
    Dim http As Object, url As String
    url = "https://www.bankofcanada.ca/valet/observations/" & serie & "/csv?start_date=" & _
          Format(debut, FMT_DATE) & "&end_date=" & Format(fin, FMT_DATE)
    On Error GoTo echec
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.send
    If http.Status = 200 Then TelechargerCsv = http.responseText
    Exit Function
echec:
    TelechargerCsv = ""
End Function

Private Function AjouterTauxDepuisCsv(ByVal csv As String, ByVal annuel As Boolean, ByVal existants As Object) As Long
    Dim lignes() As String, i As Long, champs() As String, s As String, d As Date, ok As Boolean, enTete As Boolean, n As Long
    lignes = Split(Replace(csv, vbCr, ""), vbLf)
    For i = 0 To UBound(lignes)
        s = Replace(lignes(i), """", "")
        If Not enTete Then
            If LCase$(Left$(s, 5)) = "date," Then enTete = True
        Else
            champs = Split(s, ",")
            If UBound(champs) >= 1 Then
                d = VersDate(champs(0), ok)
                If ok And IsNumeric(champs(1)) Then
                    If annuel Then
                        AjouterTauxAnnuel Year(d), Val(champs(1)): n = n + 1
                    ElseIf Not existants.Exists(CStr(CLng(d))) Then
                        AjouterTauxQuotidien d, Val(champs(1)): existants.Item(CStr(CLng(d))) = 1: n = n + 1
                    End If
                End If
            End If
        End If
    Next
    AjouterTauxDepuisCsv = n
End Function

Private Sub AssurerTaux(ByVal debut As Date, ByVal fin As Date)
    Dim couvert As Boolean, csv As String, existants As Object, i As Long, n As Long
    ChargerTauxDepuisFeuille
    If nTaux > 0 Then couvert = (txDates(1) <= debut) And (fin - txDates(nTaux) <= ECART_TAUX_MAX)
    If couvert Then Exit Sub
    Set existants = CreateObject("Scripting.Dictionary")
    For i = 1 To nTaux
        existants.Item(CStr(CLng(txDates(i)))) = 1
    Next
    csv = TelechargerCsv("FXUSDCAD", debut, fin)
    If Len(csv) = 0 Then
        Alerte NIV_AVERT, FEUILLE_TAUX, "Banque du Canada injoignable (réseau bloqué ?). Coller les taux manquants du " & _
               Format(debut, FMT_DATE) & " au " & Format(fin, FMT_DATE) & " dans l'onglet Taux, puis relancer."
        Exit Sub
    End If
    n = AjouterTauxDepuisCsv(csv, False, existants)
    csv = TelechargerCsv("FXAUSDCAD", DateSerial(Year(debut), 1, 1), fin)
    If Len(csv) > 0 Then AjouterTauxDepuisCsv csv, True, existants
    TrierTaux
    EcrireFeuilleTaux
    Alerte NIV_INFO, FEUILLE_TAUX, n & " taux quotidiens téléchargés de la Banque du Canada"
End Sub

' ============================================================================
' Lecture du classeur
' ============================================================================

Private Function VersDate(ByVal v As Variant, ByRef ok As Boolean) As Date
    Dim s As String
    ok = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then ok = True: VersDate = CDate(v): Exit Function
    If IsNumeric(v) And VarType(v) <> vbString Then
        If CDbl(v) > 30000 Then ok = True: VersDate = CDate(CDbl(v))
        Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        On Error Resume Next
        VersDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
        ok = (Err.Number = 0)
        On Error GoTo 0
        Exit Function
    End If
    If IsDate(s) Then ok = True: VersDate = CDate(s)
End Function

Private Function VersNombre(ByVal v As Variant, ByRef donne As Boolean) As Double
    Dim s As String
    donne = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then
        s = Trim$(CStr(v))
        s = Replace(Replace(Replace(s, " ", ""), Chr$(160), ""), "$", "")
        If Len(s) = 0 Then Exit Function
        On Error Resume Next
        VersNombre = CDbl(s)
        donne = (Err.Number = 0)
        On Error GoTo 0
        Exit Function
    End If
    If IsNumeric(v) Then donne = True: VersNombre = CDbl(v)
End Function

Private Function Texte(ByVal v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then Texte = "" Else Texte = Trim$(CStr(v))
End Function

Private Function Entetes(ByVal ws As Worksheet) As Object
    Dim d As Object, c As Long, s As String
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To 60
        s = LCase$(Texte(ws.Cells(1, c).Value))
        If Len(s) > 0 Then If Not d.Exists(s) Then d.Item(s) = c
    Next
    Set Entetes = d
End Function

Private Function Col(ByVal ent As Object, ByVal nom As String, ByVal ws As Worksheet) As Long
    If ent.Exists(LCase$(nom)) Then
        Col = ent.Item(LCase$(nom))
    Else
        Err.Raise vbObjectError + 1, "Lecture", "Onglet « " & ws.Name & "» : colonne « " & nom & "» introuvable en ligne 1"
    End If
End Function

Private Function DerniereLigne(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long, maxR As Long
    maxR = 1
    For c = 1 To 14
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > maxR Then maxR = r
    Next
    DerniereLigne = maxR
End Function

Private Function LigneVide(ByVal ws As Worksheet, ByVal r As Long, ByVal nCols As Long) As Boolean
    Dim c As Long
    For c = 1 To nCols
        If Len(Texte(ws.Cells(r, c).Value)) > 0 Then LigneVide = False: Exit Function
    Next
    LigneVide = True
End Function

Private Sub LireParametres()
    Dim ws As Worksheet, r As Long, cle As String, v As String, donne As Boolean
    If Not FeuilleExiste("Paramètres") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Paramètres")
    For r = 1 To 20
        cle = LCase$(Texte(ws.Cells(r, 1).Value))
        v = Texte(ws.Cells(r, 2).Value)
        If cle = "client" Then
            pClient = v
        ElseIf cle = "année" Or cle = "annee" Then
            If IsNumeric(v) Then pAnnee = CLng(v)
        ElseIf Left$(cle, 10) = "convention" Then
            If UCase$(Left$(v, 1)) = "R" Then pConvention = "REGLEMENT" Else pConvention = "TRANSACTION"
        ElseIf Left$(cle, 21) = "taux pour les revenus" Or Left$(cle, 7) = "méthode" Then
            If UCase$(Left$(v, 3)) = "MOY" Then pMethode = "MOYENNE_ANNUELLE" Else pMethode = "QUOTIDIEN"
        End If
    Next
End Sub

Private Sub LireComptes()
    Dim ws As Worksheet, ent As Object, r As Long, code As String
    Set ws = ThisWorkbook.Worksheets("Comptes")
    Set ent = Entetes(ws)
    For r = 2 To DerniereLigne(ws)
        code = Texte(ws.Cells(r, Col(ent, "Code", ws)).Value)
        If Len(code) > 0 Then
            AjouterCompte code, Texte(ws.Cells(r, Col(ent, "Courtier", ws)).Value), _
                          Texte(ws.Cells(r, Col(ent, "Numéro de compte", ws)).Value), _
                          Texte(ws.Cells(r, Col(ent, "Description", ws)).Value)
        End If
    Next
End Sub

Private Sub LireTransactions()
    Dim ws As Worksheet, ent As Object, r As Long, i As Long, ok As Boolean, donne As Boolean
    Dim cDT As Long, cDR As Long, cTy As Long, cTi As Long, cDe As Long, cQ As Long, cP As Long, cB As Long, cC As Long, cR As Long, cCp As Long, cPb As Long, cN As Long
    For i = 1 To nCptes
        If Not FeuilleExiste(cptes(i).code) Then
            Alerte NIV_ERREUR, "Comptes", "Aucun onglet nommé « " & cptes(i).code & "» pour ce compte"
        Else
            Set ws = ThisWorkbook.Worksheets(cptes(i).code)
            Set ent = Entetes(ws)
            cDT = Col(ent, "Date transaction", ws): cDR = Col(ent, "Date règlement", ws): cTy = Col(ent, "Type", ws)
            cTi = Col(ent, "Titre", ws): cDe = Col(ent, "Description", ws): cQ = Col(ent, "Quantité", ws)
            cP = Col(ent, "Prix USD", ws): cB = Col(ent, "Montant brut USD", ws): cC = Col(ent, "Commission USD", ws)
            cR = Col(ent, "Retenue USD", ws): cCp = Col(ent, "Compte contrepartie", ws)
            cPb = Col(ent, "PBR CAD (ouverture)", ws): cN = Col(ent, "Note", ws)
            For r = 2 To DerniereLigne(ws)
                If Not LigneVide(ws, r, 13) Then
                    nTx = nTx + 1
                    With tx(nTx)
                        .compte = cptes(i).code
                        .origine = cptes(i).code & "!L" & r
                        .dateTx = VersDate(ws.Cells(r, cDT).Value, ok)
                        If Not ok Then
                            Alerte NIV_ERREUR, .origine, "Date transaction manquante ou illisible"
                            nTx = nTx - 1
                        Else
                            .dateRegl = VersDate(ws.Cells(r, cDR).Value, ok)
                            If Not ok Then .dateRegl = 0
                            .typ = UCase$(Texte(ws.Cells(r, cTy).Value))
                            .titre = UCase$(Texte(ws.Cells(r, cTi).Value))
                            .description = Texte(ws.Cells(r, cDe).Value)
                            .quantite = VersNombre(ws.Cells(r, cQ).Value, donne)
                            .prixUsd = VersNombre(ws.Cells(r, cP).Value, donne)
                            .brutUsd = VersNombre(ws.Cells(r, cB).Value, .brutDonne)
                            .commission = VersNombre(ws.Cells(r, cC).Value, donne)
                            .retenue = VersNombre(ws.Cells(r, cR).Value, donne)
                            .contrepartie = Texte(ws.Cells(r, cCp).Value)
                            .pbrCadOuv = VersNombre(ws.Cells(r, cPb).Value, .pbrCadDonne)
                            .note = Texte(ws.Cells(r, cN).Value)
                        End If
                    End With
                End If
            Next
        End If
    Next
End Sub

Private Sub LireReleves()
    Dim ws As Worksheet, ent As Object, r As Long, ok As Boolean, donne As Boolean
    If Not FeuilleExiste("Relevés") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Relevés")
    Set ent = Entetes(ws)
    For r = 2 To DerniereLigne(ws)
        If Not LigneVide(ws, r, 4) Then
            nReleves = nReleves + 1
            With releves(nReleves)
                .compte = Texte(ws.Cells(r, Col(ent, "Compte", ws)).Value)
                .dFin = VersDate(ws.Cells(r, Col(ent, "Date de fin", ws)).Value, ok)
                .titre = UCase$(Texte(ws.Cells(r, Col(ent, "Titre", ws)).Value))
                .valeur = VersNombre(ws.Cells(r, Col(ent, "Quantité ou solde USD", ws)).Value, donne)
                .origine = "Relevés!L" & r
                If Not ok Then Alerte NIV_ERREUR, .origine, "Relevé : date illisible": nReleves = nReleves - 1
            End With
        End If
    Next
End Sub

Private Sub Lire1099()
    Dim ws As Worksheet, ent As Object, r As Long, donne As Boolean
    If Not FeuilleExiste("1099") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("1099")
    Set ent = Entetes(ws)
    For r = 2 To DerniereLigne(ws)
        If Not LigneVide(ws, r, 6) Then
            n1099 = n1099 + 1
            With l1099(n1099)
                .compte = Texte(ws.Cells(r, Col(ent, "Compte", ws)).Value)
                .annee = CLng(VersNombre(ws.Cells(r, Col(ent, "Année", ws)).Value, donne))
                .div = VersNombre(ws.Cells(r, Col(ent, "Dividendes bruts USD", ws)).Value, donne)
                .ret = VersNombre(ws.Cells(r, Col(ent, "Impôt étranger retenu USD", ws)).Value, donne)
                .interets = VersNombre(ws.Cells(r, Col(ent, "Intérêts USD", ws)).Value, donne)
                .prod = VersNombre(ws.Cells(r, Col(ent, "Produits de disposition USD", ws)).Value, donne)
                .origine = "1099!L" & r
            End With
        End If
    Next
End Sub

' ============================================================================
' Moteur
' ============================================================================

Private Function MontantBrut(ByVal i As Long) As Double
    If tx(i).brutDonne Then MontantBrut = tx(i).brutUsd Else MontantBrut = tx(i).quantite * tx(i).prixUsd
End Function

Private Function DetenuFinDeJour(ByVal titre As String, ByVal d As Date) As Double
    Dim k As Long, s As Double
    For k = 1 To nOrdre
        With tx(ordreTx(k))
            If .titre = titre And .dEff <= d Then
                Select Case .typ
                    Case "ACHAT", "SOLDE_OUVERTURE", "FRACTIONNEMENT": s = s + .quantite
                    Case "VENTE": s = s - .quantite
                End Select
            End If
        End With
    Next
    DetenuFinDeJour = s
End Function

Private Function AchatsFenetre(ByVal titre As String, ByVal d As Date) As Double
    Dim k As Long, s As Double
    For k = 1 To nOrdre
        With tx(ordreTx(k))
            If .typ = "ACHAT" And .titre = titre Then
                If Abs(.dEff - d) <= FENETRE_PA Then s = s + .quantite
            End If
        End With
    Next
    AchatsFenetre = s
End Function

Private Sub MajMax(ByVal d As Date, ByVal debutAnnee As Date, ByVal finAnnee As Date)
    Dim i As Long, t As Double, ds As Date, ok As Boolean
    If d < debutAnnee Or d > finAnnee Then Exit Sub
    For i = 1 To nPools
        DMax maxCout, poolTitre(i), poolCad(i)
    Next
    t = TauxDernier(d, ds, ok)
    If Not ok Then Exit Sub
    For i = 1 To nCptes
        DMax maxEnc, cptes(i).code, cptes(i).encaisse * t
    Next
End Sub

Private Sub OuvrirAnnee(ByVal debutAnnee As Date)
    Dim i As Long, t As Double, ds As Date, ok As Boolean
    For i = 1 To nPools
        maxCout.Item(poolTitre(i)) = poolCad(i)
    Next
    t = TauxDernier(debutAnnee - 1, ds, ok)
    If ok Then
        For i = 1 To nCptes
            maxEnc.Item(cptes(i).code) = cptes(i).encaisse * t
        Next
    End If
End Sub

Private Sub CloturerAnnee(ByVal finAnnee As Date)
    Dim i As Long, ok As Boolean, cle As Variant
    For i = 1 To nPools
        poolFinQte(i) = poolQte(i): poolFinCad(i) = poolCad(i): poolFinUsd(i) = poolUsd(i)
    Next
    For i = 1 To nCptes
        cptes(i).encaisseFin = cptes(i).encaisse
    Next
    qteFin.RemoveAll
    For Each cle In qteCompte.Keys
        qteFin.Item(cle) = qteCompte.Item(cle)
    Next
    tauxFinV = TauxDernier(finAnnee, tauxFinD, ok)
    aTauxFin = ok
    If ok Then
        If finAnnee - tauxFinD > ECART_TAUX_MAX Then Alerte NIV_AVERT, "Fin d'année", "Taux de fin d'année : dernier taux disponible le " & Format(tauxFinD, FMT_DATE)
    Else
        Alerte NIV_AVERT, "Fin d'année", "Aucun taux disponible pour évaluer l'encaisse au " & Format(finAnnee, FMT_DATE)
    End If
    aCloture = True
End Sub

Private Sub VerifierReleve(ByVal k As Long)
    Dim calcule As Double, ecart As Double, statut As String, ic As Long
    With releves(k)
        If Not idxCompte.Exists(.compte) Then
            Alerte NIV_ERREUR, .origine, "Relevé : compte inconnu « " & .compte & "»"
            Exit Sub
        End If
        ic = idxCompte.Item(.compte)
        If .titre = ENCAISSE Then calcule = cptes(ic).encaisse Else calcule = DGet(qteCompte, CleQte(.compte, .titre))
        ecart = .valeur - calcule
        If Abs(ecart) <= TOLERANCE Then statut = "OK" Else statut = "ÉCART"
        nRappRel = nRappRel + 1
        rappRel(nRappRel).compte = .compte: rappRel(nRappRel).dFin = .dFin: rappRel(nRappRel).element = .titre
        rappRel(nRappRel).decl = .valeur: rappRel(nRappRel).calcule = calcule: rappRel(nRappRel).ecart = ecart: rappRel(nRappRel).statut = statut
        If statut = "ÉCART" Then
            Alerte NIV_AVERT, .origine, "Relevé " & .compte & " au " & Format(.dFin, FMT_DATE) & " : " & .titre & " selon relevé " & _
                   Format(.valeur, "#,##0.00") & ", calculé " & Format(calcule, "#,##0.00") & ", écart " & Format(ecart, "#,##0.00")
        End If
    End With
End Sub

Private Sub TrierReleves()
    Dim i As Long, j As Long, tmp As TReleve
    For i = 2 To nReleves
        tmp = releves(i): j = i
        Do While j > 1
            If releves(j - 1).dFin > tmp.dFin Then releves(j) = releves(j - 1): j = j - 1 Else Exit Do
        Loop
        releves(j) = tmp
    Next
End Sub

Private Sub Moteur()
    Dim i As Long, k As Long, d As Date, debutAnnee As Date, finAnnee As Date, iRel As Long, ouvert As Boolean, ok As Boolean
    debutAnnee = DateSerial(pAnnee, 1, 1): finAnnee = DateSerial(pAnnee, 12, 31)

    ' --- validation --------------------------------------------------------------
    For i = 1 To nTx
        With tx(i)
            .ordre = i
            .titre = UCase$(Trim$(.titre)): .typ = UCase$(Trim$(.typ))
            .valide = True
            If Not TypeValide(.typ) Then
                Alerte NIV_ERREUR, .origine, "Type d'opération inconnu : « " & .typ & "»": .valide = False
            ElseIf Not idxCompte.Exists(.compte) Then
                Alerte NIV_ERREUR, .origine, "Compte inconnu : « " & .compte & "»": .valide = False
            ElseIf (.typ = "TRANSFERT_TITRES" Or .typ = "TRANSFERT_ENCAISSE") And Not idxCompte.Exists(.contrepartie) Then
                Alerte NIV_ERREUR, .origine, "Compte contrepartie inconnu : « " & .contrepartie & "»": .valide = False
            ElseIf (.typ = "TRANSFERT_TITRES" Or .typ = "TRANSFERT_ENCAISSE") And .contrepartie = .compte Then
                Alerte NIV_ERREUR, .origine, "Transfert vers le même compte": .valide = False
            ElseIf Len(.titre) = 0 Then
                Alerte NIV_ERREUR, .origine, "Titre manquant (utiliser ENCAISSE pour les mouvements de trésorerie)": .valide = False
            ElseIf (.typ = "ACHAT" Or .typ = "VENTE" Or .typ = "FRACTIONNEMENT" Or .typ = "TRANSFERT_TITRES") And .titre = ENCAISSE Then
                Alerte NIV_ERREUR, .origine, .typ & " sur ENCAISSE n'a pas de sens": .valide = False
            ElseIf (.typ = "ACHAT" Or .typ = "VENTE" Or .typ = "TRANSFERT_TITRES") And .quantite <= 0 Then
                Alerte NIV_ERREUR, .origine, .typ & " : la quantité doit être positive": .valide = False
            End If
            If .valide Then
                If pConvention = "REGLEMENT" And .dateRegl > 0 Then .dEff = .dateRegl Else .dEff = .dateTx
                If Len(.description) > 0 And Not descr.Exists(.titre) Then descr.Item(.titre) = .description
            End If
        End With
    Next

    ' --- tri stable par date effective ---------------------------------------------
    nOrdre = 0
    For i = 1 To nTx
        If tx(i).valide Then
            nOrdre = nOrdre + 1: k = nOrdre
            Do While k > 1
                If tx(ordreTx(k - 1)).dEff > tx(i).dEff Then ordreTx(k) = ordreTx(k - 1): k = k - 1 Else Exit Do
            Loop
            ordreTx(k) = i
        End If
    Next
    TrierReleves

    ' --- boucle principale -----------------------------------------------------------
    iRel = 1: ouvert = False: aCloture = False
    For k = 1 To nOrdre
        i = ordreTx(k): d = tx(i).dEff
        If Not ouvert And d >= debutAnnee Then OuvrirAnnee debutAnnee: ouvert = True
        Do While iRel <= nReleves
            If releves(iRel).dFin < d Then VerifierReleve iRel: iRel = iRel + 1 Else Exit Do
        Loop
        If Not aCloture And d > finAnnee Then CloturerAnnee finAnnee
        Traiter i, d, debutAnnee, finAnnee
    Next
    If Not ouvert Then OuvrirAnnee debutAnnee
    Do While iRel <= nReleves
        VerifierReleve iRel: iRel = iRel + 1
    Loop
    If Not aCloture Then CloturerAnnee finAnnee

    Rapprocher1099
    PreparerT1135
    If pMethode = "MOYENNE_ANNUELLE" Then tauxMoyenUtil = TauxMoyen(pAnnee, aTauxMoyen)
End Sub

Private Sub Traiter(ByVal i As Long, ByVal d As Date, ByVal debutAnnee As Date, ByVal finAnnee As Date)
    Dim ic As Long, ip As Long, brut As Double, t As Double, ok As Boolean, montantCad As Double, aCad As Boolean, aTaux As Boolean
    Dim note As String, encAvant As Double, dansAnnee As Boolean
    Dim coutUsd As Double, produitUsd As Double, produitCad As Double, pbrCad As Double, pbrUsd As Double
    Dim gainCadAvant As Double, gainUsd As Double, refusee As Double, achetes As Double, detenu As Double, fraction As Double
    Dim brutCad As Double, retenueCad As Double, ratio As Double, gainRepute As Double, montantPbr As Double, categorie As String

    ic = idxCompte.Item(tx(i).compte)
    ip = 0
    If tx(i).titre <> ENCAISSE Then ip = IdxPoolOuCreer(tx(i).titre)
    brut = MontantBrut(i)
    note = tx(i).note
    encAvant = cptes(ic).encaisse
    dansAnnee = (d >= debutAnnee And d <= finAnnee)
    aCad = False: aTaux = False

    Select Case tx(i).typ
        Case "SOLDE_OUVERTURE"
            If tx(i).titre = ENCAISSE Then
                cptes(ic).encaisse = cptes(ic).encaisse + brut
            Else
                If Not tx(i).pbrCadDonne Then
                    Alerte NIV_ERREUR, tx(i).origine, "Solde d'ouverture : PBR CAD (ouverture) manquant"
                    Exit Sub
                End If
                poolQte(ip) = poolQte(ip) + tx(i).quantite
                poolUsd(ip) = poolUsd(ip) + brut
                poolCad(ip) = poolCad(ip) + tx(i).pbrCadOuv
                DAdd qteCompte, CleQte(tx(i).compte, tx(i).titre), tx(i).quantite
                montantCad = tx(i).pbrCadOuv: aCad = True
            End If

        Case "ACHAT"
            t = TauxAu(d, tx(i).origine, ok): If Not ok Then Exit Sub
            aTaux = True
            coutUsd = brut + tx(i).commission
            montantCad = coutUsd * t: aCad = True
            poolQte(ip) = poolQte(ip) + tx(i).quantite
            poolUsd(ip) = poolUsd(ip) + coutUsd
            poolCad(ip) = poolCad(ip) + montantCad
            DAdd qteCompte, CleQte(tx(i).compte, tx(i).titre), tx(i).quantite
            cptes(ic).encaisse = cptes(ic).encaisse - coutUsd

        Case "VENTE"
            t = TauxAu(d, tx(i).origine, ok): If Not ok Then Exit Sub
            aTaux = True
            If poolQte(ip) + 0.000000001 < tx(i).quantite Then
                Alerte NIV_ERREUR, tx(i).origine, "Vente de " & tx(i).quantite & " " & tx(i).titre & " alors que le client en détient " & poolQte(ip)
                Exit Sub
            End If
            If DGet(qteCompte, CleQte(tx(i).compte, tx(i).titre)) + 0.000000001 < tx(i).quantite Then
                Alerte NIV_AVERT, tx(i).origine, "Vente de " & tx(i).quantite & " " & tx(i).titre & " dans " & tx(i).compte & " qui n'en détient que " & _
                       DGet(qteCompte, CleQte(tx(i).compte, tx(i).titre)) & " (transfert manquant ?)"
            End If
            produitUsd = brut - tx(i).commission
            produitCad = brut * t - tx(i).commission * t
            pbrCad = PoolMoyenCad(ip) * tx(i).quantite
            pbrUsd = PoolMoyenUsd(ip) * tx(i).quantite
            gainCadAvant = produitCad - pbrCad
            gainUsd = produitUsd - pbrUsd
            poolQte(ip) = poolQte(ip) - tx(i).quantite
            poolCad(ip) = poolCad(ip) - pbrCad
            poolUsd(ip) = poolUsd(ip) - pbrUsd
            If Abs(poolQte(ip)) < 0.000000001 Then poolQte(ip) = 0: poolCad(ip) = 0: poolUsd(ip) = 0
            DAdd qteCompte, CleQte(tx(i).compte, tx(i).titre), -tx(i).quantite
            cptes(ic).encaisse = cptes(ic).encaisse + produitUsd
            refusee = 0
            If gainCadAvant < 0 Then
                achetes = AchatsFenetre(tx(i).titre, d)
                detenu = DetenuFinDeJour(tx(i).titre, d + FENETRE_PA)
                If achetes > 0 And detenu > 0 Then
                    fraction = tx(i).quantite
                    If achetes < fraction Then fraction = achetes
                    If detenu < fraction Then fraction = detenu
                    fraction = fraction / tx(i).quantite
                    refusee = -gainCadAvant * fraction
                    poolCad(ip) = poolCad(ip) + refusee
                    note = Concat(note, "Perte apparente : " & Format(refusee, "#,##0.00") & " CAD refusé (" & Format(fraction, "0%") & ") et ajouté au PBR")
                    Alerte NIV_AVERT, tx(i).origine, "Perte apparente sur " & tx(i).titre & " : " & Format(refusee, "#,##0.00") & " CAD refusé, ajouté au PBR"
                End If
            End If
            montantCad = produitCad: aCad = True
            nDisps = nDisps + 1
            With disps(nDisps)
                .d = d: .compte = tx(i).compte: .titre = tx(i).titre: .quantite = tx(i).quantite: .prix = tx(i).prixUsd
                .commission = tx(i).commission: .produitUsd = produitUsd: .taux = t: .produitCad = produitCad
                .pbrUsd = pbrUsd: .pbrCad = pbrCad: .gainUsd = gainUsd: .gainCadAvant = gainCadAvant
                .refusee = refusee: .gainCad = gainCadAvant + refusee: .note = note: .origine = tx(i).origine
            End With
            If dansAnnee Then DAdd gainTitre, tx(i).titre, gainCadAvant + refusee

        Case "DIVIDENDE", "INTERET"
            If pMethode = "MOYENNE_ANNUELLE" Then
                t = TauxMoyen(Year(d), ok)
                If Not ok Then Alerte NIV_ERREUR, tx(i).origine, "Aucun taux pour l'année " & Year(d): Exit Sub
            Else
                t = TauxAu(d, tx(i).origine, ok): If Not ok Then Exit Sub
            End If
            aTaux = True
            brutCad = brut * t: retenueCad = tx(i).retenue * t
            montantCad = brutCad: aCad = True
            cptes(ic).encaisse = cptes(ic).encaisse + brut - tx(i).retenue
            If tx(i).typ = "DIVIDENDE" Then categorie = "Dividende" Else categorie = "Intérêt"
            If tx(i).typ = "DIVIDENDE" And brut > 0 Then
                ratio = tx(i).retenue / brut
                If ratio > SEUIL_RETENUE Then
                    Alerte NIV_AVERT, tx(i).origine, "Retenue de " & Format(ratio, "0.0%") & " sur " & tx(i).titre & _
                           " : excédent au-delà de 15 % non créditable (W-8BEN manquant ? déduction 20(11) à considérer)"
                ElseIf ratio = 0 Then
                    Alerte NIV_INFO, tx(i).origine, "Dividende " & tx(i).titre & " sans retenue étrangère"
                End If
            End If
            AjouterRevenu d, tx(i).compte, tx(i).titre, categorie, brut, tx(i).retenue, t, brutCad, retenueCad, note, tx(i).origine
            If dansAnnee Then
                If tx(i).titre = ENCAISSE Then DAdd intCompte, tx(i).compte, brutCad Else DAdd revTitre, tx(i).titre, brutCad
            End If

        Case "REMB_CAPITAL"
            t = TauxAu(d, tx(i).origine, ok): If Not ok Then Exit Sub
            aTaux = True
            montantCad = brut * t: aCad = True
            cptes(ic).encaisse = cptes(ic).encaisse + brut
            gainRepute = 0
            If poolCad(ip) - montantCad < -0.000000001 Then
                gainRepute = montantCad - poolCad(ip)
                montantPbr = poolCad(ip)
                note = Concat(note, "Remboursement supérieur au PBR : gain en capital réputé de " & Format(gainRepute, "#,##0.00") & " CAD")
                Alerte NIV_AVERT, tx(i).origine, note
            Else
                montantPbr = montantCad
            End If
            poolCad(ip) = poolCad(ip) - montantPbr
            poolUsd(ip) = poolUsd(ip) - brut
            If poolUsd(ip) < 0 Then poolUsd(ip) = 0
            AjouterRevenu d, tx(i).compte, tx(i).titre, "Remboursement de capital", brut, 0, t, montantCad, 0, note, tx(i).origine
            If gainRepute > 0 Then
                nDisps = nDisps + 1
                With disps(nDisps)
                    .d = d: .compte = tx(i).compte: .titre = tx(i).titre: .taux = t
                    .gainCadAvant = gainRepute: .gainCad = gainRepute
                    .note = "Gain réputé : remboursement de capital > PBR": .origine = tx(i).origine
                End With
                If dansAnnee Then DAdd gainTitre, tx(i).titre, gainRepute
            End If

        Case "FRAIS"
            t = TauxAu(d, tx(i).origine, ok): If Not ok Then Exit Sub
            aTaux = True
            montantCad = brut * t: aCad = True
            cptes(ic).encaisse = cptes(ic).encaisse - brut
            AjouterRevenu d, tx(i).compte, tx(i).titre, "Frais", brut, 0, t, montantCad, 0, note, tx(i).origine

        Case "FRACTIONNEMENT"
            If poolQte(ip) + tx(i).quantite < -0.000000001 Then
                Alerte NIV_ERREUR, tx(i).origine, "Fractionnement : quantité résultante négative pour " & tx(i).titre
                Exit Sub
            End If
            poolQte(ip) = poolQte(ip) + tx(i).quantite
            DAdd qteCompte, CleQte(tx(i).compte, tx(i).titre), tx(i).quantite

        Case "TRANSFERT_TITRES"
            If DGet(qteCompte, CleQte(tx(i).compte, tx(i).titre)) + 0.000000001 < tx(i).quantite Then
                Alerte NIV_AVERT, tx(i).origine, "Transfert de " & tx(i).quantite & " " & tx(i).titre & " depuis " & tx(i).compte & _
                       " qui n'en détient que " & DGet(qteCompte, CleQte(tx(i).compte, tx(i).titre))
            End If
            DAdd qteCompte, CleQte(tx(i).compte, tx(i).titre), -tx(i).quantite
            DAdd qteCompte, CleQte(tx(i).contrepartie, tx(i).titre), tx(i).quantite
            note = Concat(note, "Vers " & tx(i).contrepartie & " ; PBR inchangé (même contribuable)")

        Case "TRANSFERT_ENCAISSE"
            cptes(ic).encaisse = cptes(ic).encaisse - brut
            cptes(idxCompte.Item(tx(i).contrepartie)).encaisse = cptes(idxCompte.Item(tx(i).contrepartie)).encaisse + brut
            note = Concat(note, "Vers " & tx(i).contrepartie)

        Case "DEPOT"
            cptes(ic).encaisse = cptes(ic).encaisse + brut

        Case "RETRAIT"
            cptes(ic).encaisse = cptes(ic).encaisse - brut

        Case "AJUST_PBR"
            t = TauxAu(d, tx(i).origine, ok): If Not ok Then Exit Sub
            aTaux = True
            montantCad = brut * t: aCad = True
            poolCad(ip) = poolCad(ip) + montantCad
            poolUsd(ip) = poolUsd(ip) + brut
    End Select

    If ip > 0 Then
        If poolQte(ip) < -0.000000001 Then Alerte NIV_ERREUR, tx(i).origine, "Quantité négative pour " & tx(i).titre & " après cette opération"
    End If
    If cptes(ic).encaisse < -0.005 And encAvant >= -0.005 Then
        Alerte NIV_AVERT, tx(i).origine, "Encaisse de " & tx(i).compte & " négative (" & Format(cptes(ic).encaisse, "#,##0.00") & _
               " USD) après cette opération : dépôt, vente, intérêt ou solde d'ouverture manquant ?"
    End If

    nJour = nJour + 1
    With jour(nJour)
        .d = d: .compte = tx(i).compte: .typ = tx(i).typ: .titre = tx(i).titre: .quantite = tx(i).quantite: .prix = tx(i).prixUsd
        .brutUsd = brut: .commission = tx(i).commission: .retenue = tx(i).retenue
        .taux = t: .aTaux = aTaux: .montantCad = montantCad: .aCad = aCad
        If ip > 0 Then
            .poolQte = poolQte(ip): .poolCad = poolCad(ip): .poolMoyen = PoolMoyenCad(ip): .poolUsd = poolUsd(ip)
        End If
        .encaisse = cptes(ic).encaisse: .note = note: .origine = tx(i).origine
    End With
    MajMax d, debutAnnee, finAnnee
End Sub

Private Function Concat(ByVal a As String, ByVal b As String) As String
    If Len(a) = 0 Then Concat = b Else Concat = a & " " & b
End Function

Private Sub AjouterRevenu(ByVal d As Date, ByVal compte As String, ByVal titre As String, ByVal categorie As String, ByVal brutUsd As Double, ByVal retenueUsd As Double, ByVal _
                          t As Double, ByVal brutCad As Double, ByVal retenueCad As Double, ByVal note As String, ByVal origine As String)
    nRevs = nRevs + 1
    With revs(nRevs)
        .d = d: .compte = compte: .titre = titre: .categorie = categorie: .brutUsd = brutUsd: .retenueUsd = retenueUsd
        .taux = t: .brutCad = brutCad: .retenueCad = retenueCad: .note = note: .origine = origine
    End With
End Sub

Private Sub Rapprocher1099()
    Dim k As Long, j As Long, calc(1 To 4) As Double, decl(1 To 4) As Double, noms(1 To 4) As String, e As Long
    noms(1) = "Dividendes bruts USD": noms(2) = "Impôt étranger retenu USD": noms(3) = "Intérêts USD": noms(4) = "Produits de disposition USD"
    For k = 1 To n1099
        If Not idxCompte.Exists(l1099(k).compte) Then
            Alerte NIV_ERREUR, l1099(k).origine, "1099 : compte inconnu « " & l1099(k).compte & "»"
        Else
            For e = 1 To 4: calc(e) = 0: Next
            For j = 1 To nRevs
                If revs(j).compte = l1099(k).compte And Year(revs(j).d) = l1099(k).annee Then
                    If revs(j).categorie = "Dividende" Then calc(1) = calc(1) + revs(j).brutUsd
                    If revs(j).categorie = "Intérêt" Then calc(3) = calc(3) + revs(j).brutUsd
                    If revs(j).categorie <> "Frais" Then calc(2) = calc(2) + revs(j).retenueUsd
                End If
            Next
            For j = 1 To nDisps
                If disps(j).compte = l1099(k).compte And Year(disps(j).d) = l1099(k).annee Then calc(4) = calc(4) + disps(j).produitUsd
            Next
            decl(1) = l1099(k).div: decl(2) = l1099(k).ret: decl(3) = l1099(k).interets: decl(4) = l1099(k).prod
            For e = 1 To 4
                nRapp1099 = nRapp1099 + 1
                With rapp1099(nRapp1099)
                    .compte = l1099(k).compte: .annee = l1099(k).annee: .element = noms(e)
                    .decl = decl(e): .calcule = calc(e): .ecart = decl(e) - calc(e)
                    If Abs(.ecart) <= TOLERANCE Then .statut = "OK" Else .statut = "ÉCART"
                    If .statut = "ÉCART" Then Alerte NIV_AVERT, l1099(k).origine, "1099 " & .compte & " " & .annee & " : " & .element & " déclaré " & _
                        Format(.decl, "#,##0.00") & ", calculé " & Format(.calcule, "#,##0.00")
                End With
            Next
        End If
    Next
End Sub

Private Sub PreparerT1135()
    Dim i As Long, finCad As Double
    For i = 1 To nPools
        If DGet(maxCout, poolTitre(i)) > 0 Or poolFinCad(i) > 0 Or DGet(revTitre, poolTitre(i)) <> 0 Or DGet(gainTitre, poolTitre(i)) <> 0 Then
            nT1135 = nT1135 + 1
            With t1135(nT1135)
                .element = poolTitre(i)
                If descr.Exists(poolTitre(i)) Then .description = descr.Item(poolTitre(i))
                .coutMax = DGet(maxCout, poolTitre(i)): .coutFin = poolFinCad(i)
                .revenu = DGet(revTitre, poolTitre(i)): .gain = DGet(gainTitre, poolTitre(i))
            End With
        End If
    Next
    For i = 1 To nCptes
        If aTauxFin Then finCad = cptes(i).encaisseFin * tauxFinV Else finCad = 0
        nT1135 = nT1135 + 1
        With t1135(nT1135)
            .element = ENCAISSE & " " & cptes(i).code
            .description = Trim$("Encaisse USD chez " & cptes(i).courtier)
            .coutMax = DGet(maxEnc, cptes(i).code): .coutFin = finCad
            .revenu = DGet(intCompte, cptes(i).code): .gain = 0
        End With
    Next
End Sub

' ============================================================================
' Écriture des onglets de sortie
' ============================================================================

Private Function FeuilleSortie(ByVal nom As String) As Worksheet
    Dim ws As Worksheet
    If FeuilleExiste(nom) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(nom).Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nom
    Set FeuilleSortie = ws
End Function

' formats : une lettre par colonne. D date, M montant, Q quantité, T taux, N entier, S texte
Private Sub EcrireTable(ByVal ws As Worksheet, entetes As Variant, donnees As Variant, ByVal nLignes As Long, ByVal formats As String, ByVal ligneDepart As Long)
    Dim j As Long, nCols As Long, c As Range, f As String
    nCols = UBound(entetes) - LBound(entetes) + 1
    For j = 1 To nCols
        Set c = ws.Cells(ligneDepart, j)
        c.Value = entetes(LBound(entetes) + j - 1)
        c.Font.Bold = True: c.Font.Color = RGB(255, 255, 255): c.Interior.Color = RGB(31, 56, 100)
        c.WrapText = True: c.VerticalAlignment = xlCenter
    Next
    ws.Rows(ligneDepart).RowHeight = 32
    If nLignes > 0 Then ws.Range(ws.Cells(ligneDepart + 1, 1), ws.Cells(ligneDepart + nLignes, nCols)).Value = donnees
    For j = 1 To nCols
        If j <= Len(formats) Then f = Mid$(formats, j, 1) Else f = "S"
        With ws.Range(ws.Cells(ligneDepart + 1, j), ws.Cells(ligneDepart + IIf(nLignes > 0, nLignes, 1), j))
            Select Case f
                Case "D": .NumberFormat = FMT_DATE
                Case "M": .NumberFormat = FMT_MONTANT
                Case "Q": .NumberFormat = FMT_QTE
                Case "T": .NumberFormat = FMT_TAUX
                Case "N": .NumberFormat = "0"
            End Select
        End With
        ws.Columns(j).ColumnWidth = IIf(f = "S", 18, 14)
    Next
End Sub

Private Sub Geler(ByVal ws As Worksheet, ByVal cellule As String)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range(cellule).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub

Private Function R2(ByVal v As Double) As Double
    R2 = Round(v, 2)
End Function

Private Sub EcrireSortie()
    Dim ws As Worksheet, arr() As Variant, i As Long, n As Long, k As Long
    Dim sProdCad As Double, sProdUsd As Double, sPbrCad As Double, sPbrUsd As Double, sGains As Double, sPertes As Double, sRef As Double
    Dim sGainCad As Double, sGainUsd As Double, sDivCad As Double, sDivUsd As Double, sIntCad As Double, sIntUsd As Double
    Dim sRcCad As Double, sRcUsd As Double, sRetCad As Double, sRetUsd As Double, sFrCad As Double, sFrUsd As Double
    Dim nErr As Long, nAv As Long, ligne As Long

    ' ---- Sommaire ----------------------------------------------------------------
    For i = 1 To nDisps
        If Year(disps(i).d) = pAnnee Then
            With disps(i)
                sProdCad = sProdCad + .produitCad: sProdUsd = sProdUsd + .produitUsd
                sPbrCad = sPbrCad + .pbrCad: sPbrUsd = sPbrUsd + .pbrUsd
                If .gainCad > 0 Then sGains = sGains + .gainCad Else sPertes = sPertes + .gainCad
                sRef = sRef + .refusee: sGainCad = sGainCad + .gainCad: sGainUsd = sGainUsd + .gainUsd
            End With
        End If
    Next
    For i = 1 To nRevs
        If Year(revs(i).d) = pAnnee Then
            With revs(i)
                Select Case .categorie
                    Case "Dividende": sDivCad = sDivCad + .brutCad: sDivUsd = sDivUsd + .brutUsd
                    Case "Intérêt": sIntCad = sIntCad + .brutCad: sIntUsd = sIntUsd + .brutUsd
                    Case "Remboursement de capital": sRcCad = sRcCad + .brutCad: sRcUsd = sRcUsd + .brutUsd
                    Case "Frais": sFrCad = sFrCad + .brutCad: sFrUsd = sFrUsd + .brutUsd
                End Select
                If .categorie <> "Frais" Then sRetCad = sRetCad + .retenueCad: sRetUsd = sRetUsd + .retenueUsd
            End With
        End If
    Next
    For i = 1 To nAlertes
        If alertes(i).niveau = NIV_ERREUR Then nErr = nErr + 1
        If alertes(i).niveau = NIV_AVERT Then nAv = nAv + 1
    Next
    Set ws = FeuilleSortie("Sommaire")
    ws.Cells(1, 1).Value = "Placements américains : sommaire fiscal " & pAnnee
    ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 14
    ws.Cells(3, 1).Value = "Client": ws.Cells(3, 2).Value = pClient
    ws.Cells(4, 1).Value = "Année d'imposition": ws.Cells(4, 2).Value = pAnnee
    ws.Cells(5, 1).Value = "Convention de date": ws.Cells(5, 2).Value = IIf(pConvention = "REGLEMENT", "Date de règlement", "Date de transaction")
    ws.Cells(6, 1).Value = "Taux pour les revenus": ws.Cells(6, 2).Value = IIf(pMethode = "MOYENNE_ANNUELLE", "Moyenne annuelle BdC", "Taux quotidien BdC")
    ws.Cells(7, 1).Value = "Taux moyen annuel utilisé": If aTauxMoyen Then ws.Cells(7, 2).Value = tauxMoyenUtil: ws.Cells(7, 2).NumberFormat = FMT_TAUX
    ws.Cells(8, 1).Value = "Taux de fin d'année (date, taux)": If aTauxFin Then ws.Cells(8, 2).Value = Format(tauxFinD, FMT_DATE) & " : " & Format(tauxFinV, FMT_TAUX)
    ws.Cells(9, 1).Value = "Généré le": ws.Cells(9, 2).Value = Format(Now, "yyyy-mm-dd hh:mm")
    ws.Cells(10, 1).Value = "Erreurs": ws.Cells(10, 2).Value = nErr
    ws.Cells(11, 1).Value = "Avertissements": ws.Cells(11, 2).Value = nAv
    For i = 3 To 11: ws.Cells(i, 1).Font.Bold = True: Next
    ReDim arr(1 To 11, 1 To 3)
    arr(1, 1) = "Produits de disposition (nets de commissions)": arr(1, 2) = R2(sProdCad): arr(1, 3) = R2(sProdUsd)
    arr(2, 1) = "PBR des titres disposés": arr(2, 2) = R2(sPbrCad): arr(2, 3) = R2(sPbrUsd)
    arr(3, 1) = "Gains en capital (avant pertes)": arr(3, 2) = R2(sGains)
    arr(4, 1) = "Pertes en capital (après pertes apparentes)": arr(4, 2) = R2(sPertes)
    arr(5, 1) = "Pertes apparentes refusées (ajoutées au PBR)": arr(5, 2) = R2(sRef)
    arr(6, 1) = "Gain (perte) en capital net : annexe 3 / annexe G": arr(6, 2) = R2(sGainCad): arr(6, 3) = R2(sGainUsd)
    arr(7, 1) = "Dividendes bruts (revenu étranger)": arr(7, 2) = R2(sDivCad): arr(7, 3) = R2(sDivUsd)
    arr(8, 1) = "Intérêts bruts (revenu étranger)": arr(8, 2) = R2(sIntCad): arr(8, 3) = R2(sIntUsd)
    arr(9, 1) = "Remboursements de capital (réduisent le PBR)": arr(9, 2) = R2(sRcCad): arr(9, 3) = R2(sRcUsd)
    arr(10, 1) = "Impôt étranger retenu : T2209 / TP-772": arr(10, 2) = R2(sRetCad): arr(10, 3) = R2(sRetUsd)
    arr(11, 1) = "Frais (déductibilité à évaluer, 20(1)(bb))": arr(11, 2) = R2(sFrCad): arr(11, 3) = R2(sFrUsd)
    EcrireTable ws, Array("Élément", "CAD", "USD"), arr, 11, "SMM", 13
    ws.Columns(1).ColumnWidth = 55: ws.Columns(2).ColumnWidth = 26: ws.Columns(3).ColumnWidth = 18

    ' ---- Alertes -------------------------------------------------------------------
    Set ws = FeuilleSortie("Alertes")
    n = nAlertes
    If n > 0 Then
        ReDim arr(1 To n, 1 To 3)
        k = 0
        For i = 1 To nAlertes
            If alertes(i).niveau = NIV_ERREUR Then k = k + 1: arr(k, 1) = alertes(i).niveau: arr(k, 2) = alertes(i).origine: arr(k, 3) = alertes(i).message
        Next
        For i = 1 To nAlertes
            If alertes(i).niveau = NIV_AVERT Then k = k + 1: arr(k, 1) = alertes(i).niveau: arr(k, 2) = alertes(i).origine: arr(k, 3) = alertes(i).message
        Next
        For i = 1 To nAlertes
            If alertes(i).niveau = NIV_INFO Then k = k + 1: arr(k, 1) = alertes(i).niveau: arr(k, 2) = alertes(i).origine: arr(k, 3) = alertes(i).message
        Next
    End If
    EcrireTable ws, Array("Niveau", "Origine (onglet!ligne)", "Message"), arr, n, "SSS", 1
    ws.Columns(1).ColumnWidth = 16: ws.Columns(2).ColumnWidth = 22: ws.Columns(3).ColumnWidth = 110
    For i = 1 To n
        If ws.Cells(i + 1, 1).Value = NIV_ERREUR Then ws.Cells(i + 1, 1).Interior.Color = RGB(248, 203, 173)
    Next
    Geler ws, "A2"

    ' ---- Par titre -------------------------------------------------------------------
    EcrireParTitre

    ' ---- Dispositions ----------------------------------------------------------------
    Set ws = FeuilleSortie("Dispositions")
    n = 0
    For i = 1 To nDisps: If Year(disps(i).d) = pAnnee Then n = n + 1
    Next
    If n > 0 Then
        ReDim arr(1 To n, 1 To 17): k = 0
        For i = 1 To nDisps
            If Year(disps(i).d) = pAnnee Then
                k = k + 1
                With disps(i)
                    arr(k, 1) = .d: arr(k, 2) = .compte: arr(k, 3) = .titre: arr(k, 4) = .quantite: arr(k, 5) = R2(.prix): arr(k, 6) = R2(.commission)
                    arr(k, 7) = R2(.produitUsd): arr(k, 8) = Round(.taux, 4): arr(k, 9) = R2(.produitCad): arr(k, 10) = R2(.pbrUsd): arr(k, 11) = R2(.pbrCad)
                    arr(k, 12) = R2(.gainUsd): arr(k, 13) = R2(.gainCadAvant): arr(k, 14) = R2(.refusee): arr(k, 15) = R2(.gainCad): arr(k, 16) = .note: arr(k, 17) = .origine
                End With
            End If
        Next
    End If
    EcrireTable ws, Array("Date", "Compte", "Titre", "Quantité", "Prix USD", "Commission USD", "Produit net USD", "Taux BdC", "Produit net CAD", _
                          "PBR USD", "PBR CAD", "Gain (perte) USD (indicatif)", "Gain (perte) CAD avant perte apparente", _
                          "Perte apparente refusée CAD", "Gain (perte) CAD imposable", "Note", "Origine"), arr, n, "DSSQMMMTMMMMMMMSS", 1
    ws.Columns(16).ColumnWidth = 60
    Geler ws, "A2"

    ' ---- Revenus ------------------------------------------------------------------------
    Set ws = FeuilleSortie("Revenus")
    n = 0
    For i = 1 To nRevs: If Year(revs(i).d) = pAnnee And revs(i).categorie <> "Frais" Then n = n + 1
    Next
    If n > 0 Then
        ReDim arr(1 To n, 1 To 12): k = 0
        For i = 1 To nRevs
            If Year(revs(i).d) = pAnnee And revs(i).categorie <> "Frais" Then
                k = k + 1
                With revs(i)
                    arr(k, 1) = .d: arr(k, 2) = .compte: arr(k, 3) = .titre: arr(k, 4) = .categorie: arr(k, 5) = R2(.brutUsd): arr(k, 6) = R2(.retenueUsd)
                    arr(k, 7) = R2(.brutUsd - .retenueUsd): arr(k, 8) = Round(.taux, 4): arr(k, 9) = R2(.brutCad): arr(k, 10) = R2(.retenueCad)
                    arr(k, 11) = .note: arr(k, 12) = .origine
                End With
            End If
        Next
    End If
    EcrireTable ws, Array("Date", "Compte", "Titre", "Catégorie", "Brut USD", "Retenue USD", "Net USD", "Taux", "Brut CAD", "Retenue CAD", "Note", "Origine"), _
                arr, n, "DSSSMMMTMMSS", 1
    ws.Columns(11).ColumnWidth = 50
    Geler ws, "A2"

    ' ---- Frais ---------------------------------------------------------------------------
    Set ws = FeuilleSortie("Frais")
    n = 0
    For i = 1 To nRevs: If Year(revs(i).d) = pAnnee And revs(i).categorie = "Frais" Then n = n + 1
    Next
    If n > 0 Then
        ReDim arr(1 To n, 1 To 7): k = 0
        For i = 1 To nRevs
            If Year(revs(i).d) = pAnnee And revs(i).categorie = "Frais" Then
                k = k + 1
                With revs(i)
                    arr(k, 1) = .d: arr(k, 2) = .compte: arr(k, 3) = R2(.brutUsd): arr(k, 4) = Round(.taux, 4): arr(k, 5) = R2(.brutCad): arr(k, 6) = .note: arr(k, 7) = .origine
                End With
            End If
        Next
    End If
    EcrireTable ws, Array("Date", "Compte", "USD", "Taux", "CAD", "Note", "Origine"), arr, n, "DSMTMSS", 1
    ws.Columns(6).ColumnWidth = 50

    ' ---- Positions ------------------------------------------------------------------------
    Set ws = FeuilleSortie("Positions")
    ws.Cells(1, 1).Value = "PBR regroupé par titre au 31 décembre " & pAnnee & " (tous comptes, même contribuable)"
    ws.Cells(1, 1).Font.Bold = True
    n = 0
    For i = 1 To nPools: If Abs(poolFinQte(i)) > 0.000000001 Or Abs(poolFinCad(i)) > 0.005 Then n = n + 1
    Next
    If n > 0 Then
        ReDim arr(1 To n, 1 To 7): k = 0
        For i = 1 To nPools
            If Abs(poolFinQte(i)) > 0.000000001 Or Abs(poolFinCad(i)) > 0.005 Then
                k = k + 1
                arr(k, 1) = poolTitre(i)
                If descr.Exists(poolTitre(i)) Then arr(k, 2) = descr.Item(poolTitre(i))
                arr(k, 3) = poolFinQte(i): arr(k, 4) = R2(poolFinUsd(i)): arr(k, 6) = R2(poolFinCad(i))
                If poolFinQte(i) <> 0 Then arr(k, 5) = R2(poolFinUsd(i) / poolFinQte(i)): arr(k, 7) = R2(poolFinCad(i) / poolFinQte(i))
            End If
        Next
    End If
    EcrireTable ws, Array("Titre", "Description", "Quantité", "PBR total USD", "PBR moyen USD", "PBR total CAD", "PBR moyen CAD"), arr, n, "SSQMMMM", 3
    ws.Columns(2).ColumnWidth = 35
    ligne = n + 6
    ws.Cells(ligne, 1).Value = "Détail par compte au 31 décembre " & pAnnee: ws.Cells(ligne, 1).Font.Bold = True
    n = 0
    Dim cle As Variant, parts() As String, q As Double, ipl As Long
    For i = 1 To nCptes
        For Each cle In qteFin.Keys
            parts = Split(CStr(cle), "|")
            If parts(0) = cptes(i).code Then If Abs(CDbl(qteFin.Item(cle))) > 0.000000001 Then n = n + 1
        Next
        n = n + 1
    Next
    ReDim arr(1 To n, 1 To 5): k = 0
    For i = 1 To nCptes
        For Each cle In qteFin.Keys
            parts = Split(CStr(cle), "|")
            If parts(0) = cptes(i).code Then
                q = CDbl(qteFin.Item(cle))
                If Abs(q) > 0.000000001 Then
                    k = k + 1: ipl = idxPool.Item(parts(1))
                    arr(k, 1) = cptes(i).code: arr(k, 2) = parts(1): arr(k, 3) = q
                    If poolFinQte(ipl) <> 0 Then arr(k, 4) = R2(poolFinCad(ipl) / poolFinQte(ipl) * q): arr(k, 5) = R2(poolFinUsd(ipl) / poolFinQte(ipl) * q)
                End If
            End If
        Next
        k = k + 1
        arr(k, 1) = cptes(i).code: arr(k, 2) = ENCAISSE: arr(k, 3) = R2(cptes(i).encaisseFin): arr(k, 5) = R2(cptes(i).encaisseFin)
        If aTauxFin Then arr(k, 4) = R2(cptes(i).encaisseFin * tauxFinV)
    Next
    EcrireTable ws, Array("Compte", "Titre", "Quantité / solde USD", "PBR ou valeur CAD (part du compte)", "PBR ou solde USD"), arr, n, "SSQMM", ligne + 1
    ws.Columns(2).ColumnWidth = 35: ws.Columns(4).ColumnWidth = 30

    ' ---- Journal -----------------------------------------------------------------------------
    Set ws = FeuilleSortie("Journal")
    n = nJour
    If n > 0 Then
        ReDim arr(1 To n, 1 To 18)
        For i = 1 To nJour
            With jour(i)
                arr(i, 1) = .d: arr(i, 2) = .compte: arr(i, 3) = .typ: arr(i, 4) = .titre
                If .quantite <> 0 Then arr(i, 5) = .quantite
                If .prix <> 0 Then arr(i, 6) = R2(.prix)
                arr(i, 7) = R2(.brutUsd)
                If .commission <> 0 Then arr(i, 8) = R2(.commission)
                If .retenue <> 0 Then arr(i, 9) = R2(.retenue)
                If .aTaux Then arr(i, 10) = Round(.taux, 4)
                If .aCad Then arr(i, 11) = R2(.montantCad)
                arr(i, 12) = .poolQte: arr(i, 13) = R2(.poolCad): arr(i, 14) = R2(.poolMoyen): arr(i, 15) = R2(.poolUsd)
                arr(i, 16) = R2(.encaisse): arr(i, 17) = .note: arr(i, 18) = .origine
            End With
        Next
    End If
    EcrireTable ws, Array("Date effective", "Compte", "Type", "Titre", "Quantité", "Prix USD", "Montant brut USD", "Commission USD", "Retenue USD", _
                          "Taux BdC", "Montant CAD", "Pool : quantité après", "Pool : PBR CAD après", "Pool : PBR moyen CAD", "Pool : PBR USD après", _
                          "Encaisse du compte USD après", "Note", "Origine"), arr, n, "DSSSQMMMMTMQMMMMSS", 1
    ws.Columns(17).ColumnWidth = 60
    Geler ws, "A2"

    ' ---- Rapprochement relevés --------------------------------------------------------------
    Set ws = FeuilleSortie("Rapprochement relevés")
    n = nRappRel
    If n > 0 Then
        ReDim arr(1 To n, 1 To 7)
        For i = 1 To n
            With rappRel(i)
                arr(i, 1) = .compte: arr(i, 2) = .dFin: arr(i, 3) = .element: arr(i, 4) = .decl: arr(i, 5) = .calcule: arr(i, 6) = R2(.ecart): arr(i, 7) = .statut
            End With
        Next
    End If
    EcrireTable ws, Array("Compte", "Date de fin", "Titre", "Selon relevé", "Calculé", "Écart", "Statut"), arr, n, "SDSQQQS", 1
    For i = 1 To n
        If rappRel(i).statut = "OK" Then ws.Cells(i + 1, 7).Interior.Color = RGB(198, 224, 180) Else ws.Cells(i + 1, 7).Interior.Color = RGB(248, 203, 173)
    Next
    Geler ws, "A2"

    ' ---- Rapprochement 1099 --------------------------------------------------------------------
    Set ws = FeuilleSortie("Rapprochement 1099")
    n = nRapp1099
    If n > 0 Then
        ReDim arr(1 To n, 1 To 7)
        For i = 1 To n
            With rapp1099(i)
                arr(i, 1) = .compte: arr(i, 2) = .annee: arr(i, 3) = .element: arr(i, 4) = R2(.decl): arr(i, 5) = R2(.calcule): arr(i, 6) = R2(.ecart): arr(i, 7) = .statut
            End With
        Next
    End If
    EcrireTable ws, Array("Compte", "Année", "Élément", "Selon 1099", "Calculé (USD)", "Écart", "Statut"), arr, n, "SNSMMMS", 1
    ws.Columns(3).ColumnWidth = 32
    For i = 1 To n
        If rapp1099(i).statut = "OK" Then ws.Cells(i + 1, 7).Interior.Color = RGB(198, 224, 180) Else ws.Cells(i + 1, 7).Interior.Color = RGB(248, 203, 173)
    Next

    ' ---- T1135 -------------------------------------------------------------------------------------
    Set ws = FeuilleSortie("T1135")
    ws.Cells(1, 1).Value = "Données pour le T1135 (méthode détaillée). La catégorie (1 : fonds, 2 : actions de sociétés non résidentes, " & _
                           "4 : fiducies non résidentes, etc.) reste à déterminer par titre. Le seuil de 100 000 $ CAD s'évalue sur le coût total."
    n = nT1135
    If n > 0 Then
        ReDim arr(1 To n, 1 To 6)
        For i = 1 To n
            With t1135(i)
                arr(i, 1) = .element: arr(i, 2) = .description: arr(i, 3) = R2(.coutMax): arr(i, 4) = R2(.coutFin): arr(i, 5) = R2(.revenu): arr(i, 6) = R2(.gain)
            End With
        Next
    End If
    EcrireTable ws, Array("Bien", "Description", "Coût maximal dans l'année CAD", "Coût à la fin de l'année CAD", "Revenu brut CAD", "Gain (perte) CAD"), arr, n, "SSMMMM", 3
    ws.Columns(2).ColumnWidth = 35
    ligne = n + 4
    ws.Cells(ligne, 1).Value = "Total": ws.Cells(ligne, 1).Font.Bold = True
    For k = 3 To 6
        If n > 0 Then ws.Cells(ligne, k).Formula = "=SUM(" & ws.Cells(4, k).Address(False, False) & ":" & ws.Cells(3 + n, k).Address(False, False) & ")"
        ws.Cells(ligne, k).NumberFormat = FMT_MONTANT: ws.Cells(ligne, k).Font.Bold = True
    Next

    ' ---- Taux utilisés -----------------------------------------------------------------------------
    Set ws = FeuilleSortie("Taux utilisés")
    n = nTauxUtil
    If n > 0 Then
        ReDim arr(1 To n, 1 To 4)
        For i = 1 To n
            arr(i, 1) = tauxUtilD(i): arr(i, 2) = tauxUtilS(i): arr(i, 3) = Round(tauxUtilV(i), 4)
            If tauxUtilD(i) <> tauxUtilS(i) Then arr(i, 4) = "Jour ouvrable précédent"
        Next
    End If
    EcrireTable ws, Array("Date de l'opération", "Date du taux BdC", "Taux USD/CAD", "Remarque"), arr, n, "DDTS", 1
    ws.Columns(4).ColumnWidth = 28
    If n > 1 Then ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 4)).Sort Key1:=ws.Cells(2, 1), Order1:=xlAscending, Header:=xlNo
End Sub

Private Sub EcrireParTitre()
    Dim ws As Worksheet, agg As Object, titre As Variant, v As Variant, i As Long, k As Long, n As Long, arr() As Variant, j As Long, cles As Variant
    Set agg = CreateObject("Scripting.Dictionary")
    ' v : 1 divCad 2 divUsd 3 intCad 4 intUsd 5 rcCad 6 rcUsd 7 retCad 8 retUsd 9 nbDisp 10 gainCad 11 gainUsd 12 refusee
    For i = 1 To nRevs
        If Year(revs(i).d) = pAnnee And revs(i).categorie <> "Frais" Then
            v = AggLigne(agg, revs(i).titre)
            Select Case revs(i).categorie
                Case "Dividende": v(1) = v(1) + revs(i).brutCad: v(2) = v(2) + revs(i).brutUsd
                Case "Intérêt": v(3) = v(3) + revs(i).brutCad: v(4) = v(4) + revs(i).brutUsd
                Case "Remboursement de capital": v(5) = v(5) + revs(i).brutCad: v(6) = v(6) + revs(i).brutUsd
            End Select
            v(7) = v(7) + revs(i).retenueCad: v(8) = v(8) + revs(i).retenueUsd
            agg.Item(revs(i).titre) = v
        End If
    Next
    For i = 1 To nDisps
        If Year(disps(i).d) = pAnnee Then
            v = AggLigne(agg, disps(i).titre)
            v(9) = v(9) + 1: v(10) = v(10) + disps(i).gainCad: v(11) = v(11) + disps(i).gainUsd: v(12) = v(12) + disps(i).refusee
            agg.Item(disps(i).titre) = v
        End If
    Next
    Set ws = FeuilleSortie("Par titre")
    ws.Cells(1, 1).Value = "Revenus, retenues et gains par placement pour " & pAnnee & " (CAD et USD)"
    ws.Cells(1, 1).Font.Bold = True
    n = agg.Count
    If n > 0 Then
        cles = agg.Keys
        TrierCles cles
        ReDim arr(1 To n + 1, 1 To 14)
        For k = 0 To n - 1
            v = agg.Item(cles(k))
            arr(k + 1, 1) = cles(k)
            If descr.Exists(cles(k)) Then arr(k + 1, 2) = descr.Item(cles(k))
            For j = 1 To 12
                arr(k + 1, j + 2) = IIf(j = 9, v(j), R2(CDbl(v(j))))
                arr(n + 1, j + 2) = CDbl(arr(n + 1, j + 2)) + CDbl(v(j))
            Next
        Next
        arr(n + 1, 1) = "Total"
        For j = 3 To 14: arr(n + 1, j) = IIf(j = 11, arr(n + 1, j), R2(CDbl(arr(n + 1, j)))): Next
        n = n + 1
    Else
        ReDim arr(1 To 1, 1 To 14)
    End If
    EcrireTable ws, Array("Titre", "Description", "Dividendes CAD", "Dividendes USD", "Intérêts CAD", "Intérêts USD", "Remb. de capital CAD", _
                          "Remb. de capital USD", "Impôt étranger retenu CAD", "Impôt étranger retenu USD", "Nb de dispositions", _
                          "Gain (perte) en capital CAD", "Gain (perte) USD (indicatif)", "Perte apparente refusée CAD"), arr, n, "SSMMMMMMMMNMMM", 3
    ws.Columns(2).ColumnWidth = 35
    If n > 0 Then ws.Range(ws.Cells(n + 3, 1), ws.Cells(n + 3, 14)).Font.Bold = True
    Geler ws, "C4"
End Sub

Private Function AggLigne(ByVal agg As Object, ByVal titre As String) As Variant
    Dim v(1 To 12) As Double
    If agg.Exists(titre) Then AggLigne = agg.Item(titre) Else AggLigne = v
End Function

Private Sub TrierCles(ByRef cles As Variant)
    Dim i As Long, j As Long, tmp As Variant
    For i = LBound(cles) + 1 To UBound(cles)
        tmp = cles(i): j = i
        Do While j > LBound(cles)
            If CStr(cles(j - 1)) > CStr(tmp) Then cles(j) = cles(j - 1): j = j - 1 Else Exit Do
        Loop
        cles(j) = tmp
    Next
End Sub

' ============================================================================
' Tests automatiques (taux synthétiques, aucune connexion)
' ============================================================================

Private Sub PrepTest(ByVal annee As Long)
    Reinit
    pClient = "Test": pAnnee = annee: pConvention = "TRANSACTION": pMethode = "QUOTIDIEN"
    AjouterCompte "A", "Courtier X", "", ""
    AjouterCompte "B", "Courtier X", "", ""
    TauxSynthetiques 1.3
End Sub

Private Sub TauxSynthetiques(ByVal v As Double)
    Dim n As Long, d As Date
    nTaux = 0: nTauxAnn = 0
    For n = CLng(DateSerial(2024, 1, 1)) To CLng(DateSerial(2025, 12, 31))
        d = CDate(n)
        If Weekday(d, vbMonday) <= 5 Then AjouterTauxQuotidien d, v
    Next
    TrierTaux
End Sub

Private Sub FixerTaux(ByVal d As Date, ByVal v As Double)
    Dim i As Long
    i = IdxTauxPrecedent(d)
    If i > 0 Then If txDates(i) = d Then txVals(i) = v
End Sub

Private Sub AjTx(ByVal compte As String, ByVal typ As String, ByVal titre As String, ByVal d As Date, Optional quantite As Double = 0, Optional prix As Double = 0, ByVal _
                 Optional brut As Variant, Optional commission As Double = 0, Optional retenue As Double = 0, Optional contrepartie As String = "", ByVal _
                 Optional pbrOuv As Variant)
    nTx = nTx + 1
    With tx(nTx)
        .compte = compte: .typ = typ: .titre = titre: .dateTx = d: .quantite = quantite: .prixUsd = prix
        .commission = commission: .retenue = retenue: .contrepartie = contrepartie: .origine = "test"
        If Not IsMissing(brut) Then .brutUsd = CDbl(brut): .brutDonne = True
        If Not IsMissing(pbrOuv) Then .pbrCadOuv = CDbl(pbrOuv): .pbrCadDonne = True
    End With
End Sub

Private Function Env(ByVal a As Double, ByVal b As Double) As Boolean
    Env = Abs(a - b) < 0.005
End Function

Private Function ContientAlerte(ByVal niveau As String, ByVal fragment As String) As Boolean
    Dim i As Long
    For i = 1 To nAlertes
        If alertes(i).niveau = niveau And InStr(1, alertes(i).message, fragment, vbTextCompare) > 0 Then ContientAlerte = True: Exit Function
    Next
End Function

Private Function NbErreurs() As Long
    Dim i As Long, n As Long
    For i = 1 To nAlertes: If alertes(i).niveau = NIV_ERREUR Then n = n + 1
    Next
    NbErreurs = n
End Function

Public Sub TestsAutomatiques()
    Dim nOk As Long, nEchec As Long, detail As String, i As Long, ip As Long, k As Long

    ' 1. PBR moyen en CAD et effet de change
    PrepTest 2025
    FixerTaux DateSerial(2025, 1, 6), 1.3: FixerTaux DateSerial(2025, 1, 20), 1.4: FixerTaux DateSerial(2025, 3, 3), 1.35
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 20), 100, 12
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 3, 3), 100, 11
    Moteur
    Verif NbErreurs = 0 And nDisps = 1, "1a aucune erreur", nOk, nEchec, detail
    Verif Env(disps(1).pbrCad, 1490), "1b PBR moyen CAD = 1490", nOk, nEchec, detail
    Verif Env(disps(1).produitCad, 1485), "1c produit CAD = 1485", nOk, nEchec, detail
    Verif Env(disps(1).gainCad, -5), "1d gain CAD = -5 (effet de change)", nOk, nEchec, detail
    Verif Env(disps(1).gainUsd, 0), "1e gain USD = 0", nOk, nEchec, detail
    ip = idxPool.Item("XYZ")
    Verif Env(poolFinQte(ip), 100) And Env(poolFinCad(ip), 1490), "1f pool restant 100 titres, 1490 CAD", nOk, nEchec, detail

    ' 2. Commission au coût et déduite du produit
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 10, 100, , 10
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 2, 3), 10, 120, , 10
    Moteur
    Verif Env(disps(1).gainCad, 180 * 1.3), "2a gain CAD avec commissions = 234", nOk, nEchec, detail
    Verif Env(cptes(1).encaisseFin, 180), "2b encaisse = 180 USD", nOk, nEchec, detail

    ' 3. Perte apparente totale
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 3, 3), 100, 8
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 3, 13), 100, 8
    Moteur
    Verif Env(disps(1).refusee, 260) And Env(disps(1).gainCad, 0), "3a perte apparente totale refusée (260)", nOk, nEchec, detail
    ip = idxPool.Item("XYZ")
    Verif Env(poolFinCad(ip), 1040 + 260), "3b perte refusée ajoutée au PBR (1300)", nOk, nEchec, detail

    ' 4. Perte apparente partielle (rachat de 40 sur 100)
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 3, 3), 100, 8
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 3, 20), 40, 8
    Moteur
    Verif Env(disps(1).refusee, 104) And Env(disps(1).gainCad, -156), "4 perte apparente partielle 40 %", nOk, nEchec, detail

    ' 5. Pas de perte apparente si rachat après 30 jours
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 3, 3), 100, 8
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 5, 5), 100, 8
    Moteur
    Verif Env(disps(1).refusee, 0) And Env(disps(1).gainCad, -260), "5 aucune perte apparente après 30 jours", nOk, nEchec, detail

    ' 6. Remboursement de capital
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "REMB_CAPITAL", "XYZ", DateSerial(2025, 6, 2), , , 50
    Moteur
    ip = idxPool.Item("XYZ")
    Verif Env(poolFinCad(ip), 1235) And nDisps = 0, "6a remboursement de capital réduit le PBR (1235)", nOk, nEchec, detail
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 1
    AjTx "A", "REMB_CAPITAL", "XYZ", DateSerial(2025, 6, 2), , , 150
    Moteur
    Verif nDisps = 1 And Env(disps(1).gainCad, 65), "6b remboursement > PBR : gain réputé 65", nOk, nEchec, detail

    ' 7. Fractionnement
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "FRACTIONNEMENT", "XYZ", DateSerial(2025, 6, 2), 100
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 7, 7), 200, 6
    Moteur
    Verif Env(disps(1).pbrCad, 1300) And Env(disps(1).gainCad, 260), "7 fractionnement 2 pour 1", nOk, nEchec, detail

    ' 8. Transfert entre comptes : PBR regroupé
    PrepTest 2025
    FixerTaux DateSerial(2025, 1, 6), 1.3: FixerTaux DateSerial(2025, 1, 20), 1.4
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "B", "ACHAT", "XYZ", DateSerial(2025, 1, 20), 100, 12
    AjTx "A", "TRANSFERT_TITRES", "XYZ", DateSerial(2025, 4, 1), 100, , , , , "B"
    AjTx "B", "VENTE", "XYZ", DateSerial(2025, 5, 5), 50, 11
    Moteur
    Verif NbErreurs = 0, "8a transfert sans erreur", nOk, nEchec, detail
    Verif Env(DGet(qteFin, "A|XYZ"), 0) And Env(DGet(qteFin, "B|XYZ"), 150), "8b quantités par compte après transfert", nOk, nEchec, detail
    Verif Env(disps(1).pbrCad, 745), "8c vente au PBR moyen du client (745)", nOk, nEchec, detail

    ' 9. Dividende avec retenue, et alerte au-delà de 15 %
    PrepTest 2025
    FixerTaux DateSerial(2025, 6, 2), 1.37
    AjTx "A", "DIVIDENDE", "XYZ", DateSerial(2025, 6, 2), , , 100, , 15
    AjTx "A", "DIVIDENDE", "ABC", DateSerial(2025, 6, 2), , , 100, , 30
    Moteur
    Verif Env(revs(1).brutCad, 137) And Env(revs(1).retenueCad, 20.55), "9a dividende converti au taux du jour", nOk, nEchec, detail
    Verif Env(cptes(1).encaisseFin, 85 + 70), "9b encaisse nette des retenues", nOk, nEchec, detail
    Verif ContientAlerte(NIV_AVERT, "15 %"), "9c alerte retenue > 15 %", nOk, nEchec, detail

    ' 10. Solde d'ouverture et historique antérieur
    PrepTest 2025
    FixerTaux DateSerial(2025, 3, 3), 1.4
    AjTx "A", "SOLDE_OUVERTURE", "XYZ", DateSerial(2024, 1, 2), 100, , 1000, , , , 1250
    AjTx "A", "ACHAT", "XYZ", DateSerial(2024, 6, 3), 100, 15
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 3, 3), 100, 20
    Moteur
    Verif Env(disps(1).pbrCad, 1600) And Env(disps(1).produitCad, 2800), "10 solde d'ouverture + achat 2024, vente 2025", nOk, nEchec, detail

    ' 11. Rapprochement des relevés
    PrepTest 2025
    AjTx "A", "SOLDE_OUVERTURE", ENCAISSE, DateSerial(2025, 1, 2), , , 5000
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 10, 100
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 2, 10), 5, 100
    nReleves = 3
    releves(1).compte = "A": releves(1).dFin = DateSerial(2025, 1, 31): releves(1).titre = "XYZ": releves(1).valeur = 10
    releves(2).compte = "A": releves(2).dFin = DateSerial(2025, 1, 31): releves(2).titre = ENCAISSE: releves(2).valeur = 4000
    releves(3).compte = "A": releves(3).dFin = DateSerial(2025, 2, 28): releves(3).titre = "XYZ": releves(3).valeur = 20
    Moteur
    Verif nRappRel = 3 And rappRel(1).statut = "OK" And rappRel(2).statut = "OK", "11a relevés de janvier OK", nOk, nEchec, detail
    Verif rappRel(3).statut = "ÉCART" And Env(rappRel(3).ecart, 5), "11b écart de 5 détecté en février", nOk, nEchec, detail

    ' 12. Rapprochement 1099
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 10, 100
    AjTx "A", "DIVIDENDE", "XYZ", DateSerial(2025, 6, 2), , , 100, , 15
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 9, 8), 10, 110, , 5
    n1099 = 1
    l1099(1).compte = "A": l1099(1).annee = 2025: l1099(1).div = 120: l1099(1).ret = 15: l1099(1).interets = 0: l1099(1).prod = 1095
    Moteur
    k = 0
    For i = 1 To nRapp1099: If rapp1099(i).statut = "ÉCART" Then k = k + 1
    Next
    Verif nRapp1099 = 4 And k = 1 And Env(rapp1099(1).ecart, 20), "12 un seul écart 1099 (dividendes, 20)", nOk, nEchec, detail

    ' 13. T1135 : coût maximal et coût de fin d'année
    PrepTest 2025
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 100, 10
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 3, 3), 100, 10
    AjTx "A", "VENTE", "XYZ", DateSerial(2025, 9, 8), 150, 10
    Moteur
    Verif Env(t1135(1).coutMax, 2600) And Env(t1135(1).coutFin, 650), "13 T1135 coût max 2600, fin 650", nOk, nEchec, detail

    ' 14. Taux : fin de semaine et taux manquant
    PrepTest 2025
    FixerTaux DateSerial(2025, 3, 7), 1.44
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 3, 9), 10, 100
    AjTx "A", "ACHAT", "XYZ", DateSerial(2026, 6, 1), 10, 100
    Moteur
    Verif jour(1).aTaux And Env(jour(1).taux, 1.44), "14a dimanche : taux du vendredi", nOk, nEchec, detail
    Verif ContientAlerte(NIV_ERREUR, "Aucun taux publié"), "14b taux manquant = erreur", nOk, nEchec, detail

    ' 15. Encaisse négative : une seule alerte
    PrepTest 2025
    AjTx "A", "SOLDE_OUVERTURE", ENCAISSE, DateSerial(2025, 1, 2), , , 500
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 6), 10, 100
    AjTx "A", "ACHAT", "XYZ", DateSerial(2025, 1, 7), 1, 100
    Moteur
    k = 0
    For i = 1 To nAlertes: If InStr(alertes(i).message, "Encaisse de A négative") > 0 Then k = k + 1
    Next
    Verif k = 1, "15 encaisse négative signalée une fois", nOk, nEchec, detail

    Reinit
    If nEchec = 0 Then
        MsgBox "Tous les tests réussissent : " & nOk & " vérifications.", vbInformation, "Tests automatiques"
    Else
        MsgBox nEchec & " échec(s) sur " & (nOk + nEchec) & " vérifications :" & vbCrLf & vbCrLf & detail, vbExclamation, "Tests automatiques"
    End If
End Sub

Private Sub Verif(ByVal condition As Boolean, ByVal nom As String, ByRef nOk As Long, ByRef nEchec As Long, ByRef detail As String)
    If condition Then
        nOk = nOk + 1
    Else
        nEchec = nEchec + 1
        detail = detail & "ÉCHEC : " & nom & vbCrLf
    End If
End Sub
