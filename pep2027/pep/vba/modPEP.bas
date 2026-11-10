Option Explicit

' =====================================================================================
' PEP - Synchronisierung Jahreskalender <-> Tabellen
'
' Der Jahreskalender zeigt ausschliesslich Formeln und bleibt immer geschuetzt. Ein Eintrag
' im Kalender (Doppelklick bzw. Rechtsklick) wird hier in die zustaendige Tabelle geschrieben
' (Abwesenheiten, Freistellungen, Vertretungen, Ressourcen); der Kalender zeigt ihn danach
' ueber seine Formeln an. Es gibt damit nur eine Datenquelle und nie abweichende Staende -
' auch ohne Makros, in Excel Online und bei AutoSave.
' =====================================================================================

Private mAktiv As Boolean

Private Const CALC_MANUELL As Long = -4135

' ------------------------------------------------------------------ Status / Raster
' Wird in Formeln verwendet (Startseite): ohne Makros liefert der Aufruf #NAME? -> "Makros nicht aktiv"
Public Function PEPMakrosAktiv() As Boolean
    Application.Volatile
    PEPMakrosAktiv = True
End Function

Public Function KalenderRaster(ByVal ws As Worksheet) As Range
    Set KalenderRaster = ws.Range(ws.Cells(KAL_ZEILE1, KAL_SPALTE1), _
        ws.Cells(KAL_ZEILE1 + KAL_ZEILEN - 1, KAL_SPALTE1 + KAL_TAGE - 1))
End Function

Public Function ImRaster(ByVal Target As Range) As Boolean
    ImRaster = Not Application.Intersect(Target, KalenderRaster(Target.Worksheet)) Is Nothing
End Function

' ------------------------------------------------------------------ Eingabe im Kalender
' Doppelklick (ein Tag) bzw. Rechtsklick (markierte Tage) oeffnen diese Eingabe. Das Raster selbst bleibt
' geschuetzt und enthaelt immer die Formeln; geschrieben wird nur in die Tabellen (eine Datenquelle).
Public Sub EintragDialog(ByVal rng As Range)
    Dim isect As Range, v As Variant, vorschlag As String, hinweis As String
    Set isect = Application.Intersect(rng, KalenderRaster(rng.Worksheet))
    If isect Is Nothing Then Exit Sub
    If isect.Cells.Count > 2000 Then
        MsgBox "Bitte höchstens 2000 Felder auf einmal markieren. Größere Zeiträume im Blatt " & _
            "„Abwesenheiten“ erfassen.", vbExclamation, "PEP"
        Exit Sub
    End If
    If isect.Cells.Count = 1 Then vorschlag = Vorschlag1(isect)
    hinweis = "Eintrag für " & isect.Cells.Count & " Feld(er):" & vbLf & _
        "U Urlaub · UE eTime · K Krank · JAZ · S Unterricht · FDL · Kabel" & vbLf & _
        "FS Freistellung · - Tagdienst · V Vertretung · Z Zus. MA · A AZUBI" & vbLf & _
        "leer = löschen · sonstiger Text = Besonderheit"
    v = Application.InputBox(hinweis, "PEP – Kalendereintrag", vorschlag, Type:=2)
    If VarType(v) = vbBoolean Then Exit Sub
    EintragAnwenden isect, CStr(v)
End Sub

' Vorbelegung des Dialogs mit dem Kuerzel des aktuellen Eintrags
Private Function Vorschlag1(ByVal c As Range) As String
    Dim wsB As Worksheet, code As Long, lo As ListObject, i As Long, k As String
    Set wsB = ThisWorkbook.Worksheets(BLATT_BERECHNUNG)
    code = CLng(Val(CStr(wsB.Cells(c.Row, c.Column).Value)))
    Select Case code
    Case 0: Vorschlag1 = ""
    Case 1, 11: Vorschlag1 = CStr(c.Value)
    Case 9
        If CStr(wsB.Cells(c.Row, 4).Value) <> CStr(wsB.Cells(7, c.Column).Value) Then Vorschlag1 = "V"
    Case Else
        Set lo = Tabelle(TBL_ARTEN)
        For i = 1 To lo.ListRows.Count
            If CLng(lo.ListRows(i).Range.Cells(1, lo.ListColumns("Code").Index).Value) = code Then
                k = CStr(lo.ListRows(i).Range.Cells(1, lo.ListColumns("Eingabekürzel").Index).Value)
                Vorschlag1 = Trim$(Split(k & ",", ",")(0))
            End If
        Next i
    End Select
End Function

Public Sub EintragAnwenden(ByVal isect As Range, ByVal txt As String)
    Dim area As Range, c As Range, calcAlt As Long
    If mAktiv Then Exit Sub
    mAktiv = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    calcAlt = Application.Calculation
    Application.Calculation = CALC_MANUELL
    On Error GoTo Fehler
    For Each area In isect.Areas
        For Each c In area.Cells
            ZelleVerarbeiten c, txt
        Next c
    Next area
Fertig:
    On Error Resume Next
    Application.Calculation = calcAlt
    Application.Calculate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    mAktiv = False
    Exit Sub
Fehler:
    MsgBox "Der Eintrag konnte nicht übernommen werden: " & Err.Description, vbExclamation, "PEP"
    Resume Fertig
End Sub

Private Sub ZelleVerarbeiten(ByVal c As Range, ByVal eingabe As String)
    Dim wsB As Worksheet, r As Long, col As Long, nm As String, art As String
    Dim tag As Date, alt As Long, code As Long, txt As String
    Set wsB = ThisWorkbook.Worksheets(BLATT_BERECHNUNG)
    r = c.Row: col = c.Column
    If CStr(wsB.Cells(r, 1).Value) <> "P" Then Exit Sub
    nm = CStr(wsB.Cells(r, 2).Value)
    If nm = "" Then Exit Sub
    If Not AlsDatum(wsB.Cells(KAL_DATUMZEILE, col).Value, tag) Then Exit Sub
    alt = CLng(Val(CStr(wsB.Cells(r, col).Value)))
    txt = Trim$(eingabe)
    code = EingabeAuswerten(txt, art)

    Select Case code
    Case -1                                   ' leer: sichtbaren Eintrag entfernen
        Select Case alt
        Case 1 To 8, 11: ZeitraumTagEntfernen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag
        Case 10: FreistellungAufheben nm, tag, r, col
        Case 9: ZeitraumTagEntfernen Tabelle(TBL_VER), "Vertreter", nm, tag
        Case 12, 13: ZeitraumTagEntfernen Tabelle(TBL_RES), "Ressource", nm, tag
        End Select
    Case 0                                    ' "-" = regulaerer Tagdienst
        ZeitraumTagEntfernen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag
        ZeitraumTagEntfernen Tabelle(TBL_VER), "Vertreter", nm, tag
        ZeitraumTagEntfernen Tabelle(TBL_RES), "Ressource", nm, tag
        FreistellungAufheben nm, tag, r, col
    Case 1 To 8                               ' Abwesenheit / Besonderheit
        ZeitraumTagEntfernen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag
        If code <= 4 Then FreistellungAufheben nm, tag, r, col
        ' Besonderheiten werden ueber ihren Text unterschieden, Abwesenheiten nur ueber die Art
        ZeitraumHinzufuegen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag, "Art", art, _
            IIf(code = 1, "Bemerkung", ""), IIf(code = 1, txt, "")
    Case 10                                   ' FS = Freistellung
        ZeitraumTagEntfernen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag
        FreistellungSetzen nm, tag, r, col
    Case 9                                    ' V = Vertretung in der Bereitschaft
        ZeitraumTagEntfernen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag
        FreistellungAufheben nm, tag, r, col
        ZeitraumHinzufuegen Tabelle(TBL_VER), "Vertreter", nm, tag, "Funktion", _
            Stammfunktion(CStr(wsB.Cells(r, 3).Value)), "", ""
    Case 12, 13                               ' Ressourcen: Z / A
        ZeitraumTagEntfernen Tabelle(TBL_ABW), "Mitarbeiter", nm, tag
        ZeitraumTagEntfernen Tabelle(TBL_RES), "Ressource", nm, tag
        ZeitraumHinzufuegen Tabelle(TBL_RES), "Ressource", nm, tag, "Art", _
            IIf(code = 12, "Zusätzliche MA", "AZUBIS"), "", ""
    End Select
End Sub

Private Function Stammfunktion(ByVal f As String) As String
    Select Case f
    Case "Meister", "AV", "TVT-Fahrer", "Handwerker": Stammfunktion = f
    Case Else: Stammfunktion = "Handwerker"
    End Select
End Function

' Rueckgabe: Code 1..13, 0 = "-" (Tagdienst), -1 = leer (loeschen). Freitext = 1 (Besonderheit)
Private Function EingabeAuswerten(ByVal txt As String, ByRef art As String) As Long
    Dim t As String, lo As ListObject, i As Long, k As Variant, liste As String
    Dim cCode As Long, cBez As Long, cEin As Long
    t = UCase$(Trim$(txt))
    If t = "" Then EingabeAuswerten = -1: Exit Function
    If t = "-" Then EingabeAuswerten = 0: Exit Function
    Set lo = Tabelle(TBL_ARTEN)
    cCode = lo.ListColumns("Code").Index
    cBez = lo.ListColumns("Bezeichnung").Index
    cEin = lo.ListColumns("Eingabekürzel").Index
    For i = 1 To lo.ListRows.Count
        liste = CStr(lo.ListRows(i).Range.Cells(1, cEin).Value)
        If Len(liste) > 0 Then
            For Each k In Split(liste, ",")
                If UCase$(Trim$(CStr(k))) = t Then
                    EingabeAuswerten = CLng(lo.ListRows(i).Range.Cells(1, cCode).Value)
                    art = CStr(lo.ListRows(i).Range.Cells(1, cBez).Value)
                    Exit Function
                End If
            Next k
        End If
    Next i
    EingabeAuswerten = 1
    art = ART_BESONDERHEIT
End Function

' ------------------------------------------------------------------ Freistellungen
Private Sub FreistellungSetzen(ByVal nm As String, ByVal tag As Date, ByVal r As Long, ByVal col As Long)
    Dim lo As ListObject, neu As Range
    Set lo = Tabelle(TBL_FS)
    FSEintragEntfernen lo, nm, tag, "Freistellung entfällt am"
    If StatusNeu(r, col) <> 10 Then
        Set neu = NeueZeile(lo)
        neu.Cells(1, lo.ListColumns("Mitarbeiter").Index).Value = nm
        neu.Cells(1, lo.ListColumns("Stattdessen frei am").Index).Value = tag
        neu.Cells(1, lo.ListColumns("Bemerkung").Index).Value = "im Kalender gesetzt"
        neu.Cells(1, lo.ListColumns("Quelle").Index).Value = "Kalender"
    End If
End Sub

Private Sub FreistellungAufheben(ByVal nm As String, ByVal tag As Date, ByVal r As Long, ByVal col As Long)
    Dim lo As ListObject, neu As Range
    Set lo = Tabelle(TBL_FS)
    FSEintragEntfernen lo, nm, tag, "Stattdessen frei am"
    If StatusNeu(r, col) = 10 Then
        Set neu = NeueZeile(lo)
        neu.Cells(1, lo.ListColumns("Mitarbeiter").Index).Value = nm
        neu.Cells(1, lo.ListColumns("Freistellung entfällt am").Index).Value = tag
        neu.Cells(1, lo.ListColumns("Bemerkung").Index).Value = "im Kalender aufgehoben"
        neu.Cells(1, lo.ListColumns("Quelle").Index).Value = "Kalender"
    End If
End Sub

' Entfernt die Datumsangabe in Spalte spalte fuer Mitarbeiter nm (ganze Zeile, wenn danach leer)
Private Sub FSEintragEntfernen(ByVal lo As ListObject, ByVal nm As String, ByVal tag As Date, ByVal spalte As String)
    Dim i As Long, rng As Range, d As Date, cMA As Long, cS As Long, cE As Long, cT As Long
    cMA = lo.ListColumns("Mitarbeiter").Index
    cS = lo.ListColumns(spalte).Index
    cE = lo.ListColumns("Freistellung entfällt am").Index
    cT = lo.ListColumns("Stattdessen frei am").Index
    For i = lo.ListRows.Count To 1 Step -1
        Set rng = lo.ListRows(i).Range
        If StrComp(CStr(rng.Cells(1, cMA).Value), nm, vbTextCompare) = 0 Then
            If AlsDatum(rng.Cells(1, cS).Value, d) Then
                If d = tag Then
                    rng.Cells(1, cS).ClearContents
                    If IsEmpty(rng.Cells(1, cE).Value) And IsEmpty(rng.Cells(1, cT).Value) Then
                        If lo.ListRows.Count > 1 Then
                            lo.ListRows(i).Delete
                        Else
                            EingabenLeeren lo, rng
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Status eines Tages nach Aenderung der Tabellen gezielt neu berechnen (Kalender ist im manuellen Modus)
Private Function StatusNeu(ByVal r As Long, ByVal col As Long) As Long
    Dim wsB As Worksheet
    Set wsB = ThisWorkbook.Worksheets(BLATT_BERECHNUNG)
    On Error Resume Next
    Tabelle(TBL_ABW).ListColumns("Code").DataBodyRange.Calculate
    Tabelle(TBL_RES).ListColumns("Code").DataBodyRange.Calculate
    On Error GoTo 0
    wsB.Range(wsB.Cells(9, col), wsB.Cells(11, col)).Calculate
    wsB.Cells(r + BER_ABW_VERSATZ, col).Calculate
    wsB.Cells(r, col).Calculate
    StatusNeu = CLng(Val(CStr(wsB.Cells(r, col).Value)))
End Function

' ------------------------------------------------------------------ Zeitraeume in Tabellen
' Entfernt den Tag aus allen Zeitraeumen von key (teilt Zeitraeume bei Bedarf auf)
Private Sub ZeitraumTagEntfernen(ByVal lo As ListObject, ByVal keyCol As String, ByVal key As String, _
                                 ByVal tag As Date)
    Dim i As Long, rng As Range, neu As Range, v As Date, b As Date
    Dim cK As Long, cV As Long, cB As Long
    cK = lo.ListColumns(keyCol).Index
    cV = lo.ListColumns("Von").Index
    cB = lo.ListColumns("Bis").Index
    For i = lo.ListRows.Count To 1 Step -1
        Set rng = lo.ListRows(i).Range
        If StrComp(CStr(rng.Cells(1, cK).Value), key, vbTextCompare) = 0 Then
            If AlsDatum(rng.Cells(1, cV).Value, v) And AlsDatum(rng.Cells(1, cB).Value, b) Then
                If v <= tag And b >= tag Then
                    If v = tag And b = tag Then
                        If lo.ListRows.Count > 1 Then
                            lo.ListRows(i).Delete
                        Else
                            EingabenLeeren lo, rng
                        End If
                    ElseIf v = tag Then
                        rng.Cells(1, cV).Value = tag + 1
                    ElseIf b = tag Then
                        rng.Cells(1, cB).Value = tag - 1
                    Else
                        Set neu = NeueZeile(lo)
                        EingabenKopieren lo, rng, neu
                        neu.Cells(1, cV).Value = tag + 1
                        neu.Cells(1, cB).Value = b
                        rng.Cells(1, cB).Value = tag - 1
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Fuegt den Tag hinzu: verlaengert einen angrenzenden gleichen Zeitraum oder legt eine Zeile an
Private Sub ZeitraumHinzufuegen(ByVal lo As ListObject, ByVal keyCol As String, ByVal key As String, _
                                ByVal tag As Date, ByVal artCol As String, ByVal art As String, _
                                ByVal bemCol As String, ByVal bem As String)
    Dim i As Long, j As Long, rng As Range, rng2 As Range, neu As Range, v As Date, b As Date, v2 As Date, b2 As Date
    Dim cK As Long, cV As Long, cB As Long, cA As Long, cBe As Long
    cK = lo.ListColumns(keyCol).Index
    cV = lo.ListColumns("Von").Index
    cB = lo.ListColumns("Bis").Index
    cA = lo.ListColumns(artCol).Index
    If bemCol <> "" Then cBe = lo.ListColumns(bemCol).Index
    For i = 1 To lo.ListRows.Count
        Set rng = lo.ListRows(i).Range
        If Passt(rng, cK, key, cA, art, cBe, bem) Then
            If AlsDatum(rng.Cells(1, cV).Value, v) And AlsDatum(rng.Cells(1, cB).Value, b) Then
                If b = tag - 1 Then
                    rng.Cells(1, cB).Value = tag
                    For j = lo.ListRows.Count To 1 Step -1
                        Set rng2 = lo.ListRows(j).Range
                        If j <> i And Passt(rng2, cK, key, cA, art, cBe, bem) Then
                            If AlsDatum(rng2.Cells(1, cV).Value, v2) And AlsDatum(rng2.Cells(1, cB).Value, b2) Then
                                If v2 = tag + 1 Then
                                    rng.Cells(1, cB).Value = b2
                                    lo.ListRows(j).Delete
                                    Exit For
                                End If
                            End If
                        End If
                    Next j
                    Exit Sub
                ElseIf v = tag + 1 Then
                    rng.Cells(1, cV).Value = tag
                    Exit Sub
                End If
            End If
        End If
    Next i
    Set neu = NeueZeile(lo)
    neu.Cells(1, cK).Value = key
    neu.Cells(1, cV).Value = tag
    neu.Cells(1, cB).Value = tag
    neu.Cells(1, cA).Value = art
    If cBe > 0 And bem <> "" Then neu.Cells(1, cBe).Value = bem
    neu.Cells(1, lo.ListColumns("Quelle").Index).Value = "Kalender"
End Sub

Private Function Passt(ByVal rng As Range, ByVal cK As Long, ByVal key As String, ByVal cA As Long, _
                       ByVal art As String, ByVal cBe As Long, ByVal bem As String) As Boolean
    If StrComp(CStr(rng.Cells(1, cK).Value), key, vbTextCompare) <> 0 Then Exit Function
    If StrComp(CStr(rng.Cells(1, cA).Value), art, vbTextCompare) <> 0 Then Exit Function
    If cBe > 0 Then
        If StrComp(CStr(rng.Cells(1, cBe).Value), bem, vbTextCompare) <> 0 Then Exit Function
    End If
    Passt = True
End Function

' ------------------------------------------------------------------ Tabellen-Hilfen
Public Function Tabelle(ByVal tname As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tname, vbTextCompare) = 0 Then
                Set Tabelle = lo
                Exit Function
            End If
        Next lo
    Next ws
    Err.Raise vbObjectError + 513, "PEP", "Tabelle " & tname & " wurde nicht gefunden."
End Function

Private Function NeueZeile(ByVal lo As ListObject) As Range
    Dim lr As ListRow
    If lo.ListRows.Count > 0 Then
        Set lr = lo.ListRows(lo.ListRows.Count)
        If ZeileLeer(lo, lr.Range) Then
            Set NeueZeile = lr.Range
            Exit Function
        End If
    End If
    Set lr = lo.ListRows.Add(AlwaysInsert:=True)
    Set NeueZeile = lr.Range
End Function

Private Function ZeileLeer(ByVal lo As ListObject, ByVal rng As Range) As Boolean
    Dim j As Long
    For j = 1 To lo.ListColumns.Count
        If Not rng.Cells(1, j).HasFormula Then
            If Len(CStr(rng.Cells(1, j).Value)) > 0 Then Exit Function
        End If
    Next j
    ZeileLeer = True
End Function

Private Sub EingabenKopieren(ByVal lo As ListObject, ByVal von As Range, ByVal nach As Range)
    Dim j As Long
    For j = 1 To lo.ListColumns.Count
        If Not von.Cells(1, j).HasFormula Then nach.Cells(1, j).Value = von.Cells(1, j).Value
    Next j
End Sub

Private Sub EingabenLeeren(ByVal lo As ListObject, ByVal rng As Range)
    Dim j As Long
    For j = 1 To lo.ListColumns.Count
        If Not rng.Cells(1, j).HasFormula Then rng.Cells(1, j).ClearContents
    Next j
End Sub

Private Function AlsDatum(ByVal v As Variant, ByRef d As Date) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = DateValue(v): AlsDatum = True
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then d = CDate(Int(CDbl(v))): AlsDatum = True
    ElseIf IsDate(v) Then
        d = DateValue(CDate(v)): AlsDatum = True
    End If
End Function

' ------------------------------------------------------------------ Mitarbeiter -> Funktionsmatrix
Public Sub MitarbeiterGeaendert(ByVal Target As Range)
    Dim loM As ListObject, loF As ListObject, isect As Range, c As Range, neu As Range
    Dim nm As String, fn As String, q As Variant, ziel As Range
    If mAktiv Then Exit Sub
    On Error GoTo Ende
    Set loM = Tabelle(TBL_MA)
    If loM.DataBodyRange Is Nothing Then Exit Sub
    Set isect = Application.Intersect(Target, loM.ListColumns("Name").DataBodyRange)
    If isect Is Nothing Then Set isect = Application.Intersect(Target, loM.ListColumns("Funktion").DataBodyRange)
    If isect Is Nothing Then Exit Sub
    Set loF = Tabelle(TBL_FM)
    mAktiv = True
    Application.EnableEvents = False
    For Each c In isect.Cells
        nm = Trim$(CStr(loM.ListColumns("Name").DataBodyRange.Cells(c.Row - loM.DataBodyRange.Row + 1, 1).Value))
        fn = CStr(loM.ListColumns("Funktion").DataBodyRange.Cells(c.Row - loM.DataBodyRange.Row + 1, 1).Value)
        If nm <> "" Then
            If Application.WorksheetFunction.CountIf(loF.ListColumns("Name").DataBodyRange, nm) = 0 Then
                Set neu = NeueZeile(loF)
                neu.Cells(1, loF.ListColumns("Name").Index).Value = nm
                For Each q In Split(QUALIFIKATIONEN, "|")
                    Set ziel = neu.Cells(1, loF.ListColumns(CStr(q)).Index)
                    If CStr(q) = fn Then ziel.Value = "JA" Else ziel.Value = "NEIN"
                Next q
            End If
        End If
    Next c
Ende:
    Application.EnableEvents = True
    mAktiv = False
End Sub

' ------------------------------------------------------------------ Werkzeuge (Makro-Liste)
Public Sub AllesNeuBerechnen()
    Application.CalculateFull
End Sub

' Reparatur: Kalenderformeln neu schreiben (z. B. nach versehentlichem Aufheben des Blattschutzes)
Public Sub KalenderFormelnWiederherstellen()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(BLATT_KALENDER)
    Application.EnableEvents = False
    On Error Resume Next
    ws.Unprotect PEP_KENNWORT
    On Error GoTo 0
    KalenderRaster(ws).FormulaR1C1 = KAL_FORMEL_R1C1
    KalenderRaster(ws).Locked = True
    ws.Protect Password:=PEP_KENNWORT, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingColumns:=True, AllowFormattingRows:=True
    Application.EnableEvents = True
End Sub
