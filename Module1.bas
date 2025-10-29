Attribute VB_Name = "Module1"
Option Explicit

' =================== KONFIG ===================
Private Const ARK_PLAN As String = "Planlegger"
Private Const ARK_OVERSIKT As String = "AKTIVITETSTYPER - OVERSIKT"
Public Property Get F�RSTE_DATAKOL() As Long
    F�RSTE_DATAKOL = Worksheets(ARK_PLAN).Range("FirstDate").Column
End Property

Public Property Get datoRad() As Long
    datoRad = Worksheets(ARK_PLAN).Range("FirstDate").Row
End Property

Public Property Get F�RSTE_PERSONRAD() As Long
    ' just�r +0/+1 avhengig av oppsettet ditt
    F�RSTE_PERSONRAD = Worksheets(ARK_PLAN).Range("PersonHeader").Row + 1
End Property
' =============================================

' =========================================================
'  MODUL 1 � SAMLET v3
'  Inneholder begge m�ter � legge inn aktivitet:
'    1) LeggInnAktivitet: velg person + datoer (klassisk)
'    2) LeggInnAktivitetP�Markering: marker celler og angi kode
'
'  NYTT I v3 (som du ba om):
'    - N�r markeringen treffer **overlapp** med *annen* aktivitet samme dag
'      (dvs. det ligger aktivitet i spennet og teksten ikke starter med samme kode),
'      legges blokken p� **ny under-rad** (opprettes ved behov) i samme personblokk.
'    - Hvis spennet er tomt eller kun samme aktivitet, bruker vi den **valgte raden**.
'
'  Egenskaper ellers:
'    - Anti-lekk til raden under
'    - Gjenoppretter bakgrunn under/ved nye rader (ingen hvite hull)
'    - ByVal p� verdiparametre s� Const (DATORAD) fungerer
' =========================================================

' ===================== M�TE 1 =====================
Public Sub LeggInnAktivitet()
    Dim wsPlan As Worksheet, wsTyp As Worksheet
    Dim personCell As Range
    Dim personRow As Long
    Dim kode As String, beskrivelse As String, kommentar As String, visTekst As String
    Dim startDato As Date, sluttDato As Date
    Dim startCol As Long, sluttCol As Long, m�lRad As Long
    Dim farge As Long
    Dim farger As Object

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(ARK_PLAN)
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT)
    On Error GoTo 0
    If wsPlan Is Nothing Or wsTyp Is Nothing Then
        MsgBox "Finner ikke arkene '" & ARK_PLAN & "' og/eller '" & ARK_OVERSIKT & "'.", vbCritical
        Exit Sub
    End If

    On Error Resume Next
    Set personCell = Application.InputBox( _
        prompt:="Klikk en celle i kolonne A p� '" & ARK_PLAN & "' (rad " & F�RSTE_PERSONRAD & "+).", _
        Title:="Velg person", Type:=8)
    On Error GoTo 0
    If personCell Is Nothing Then Exit Sub
    If personCell.Column <> 1 Or personCell.Row < F�RSTE_PERSONRAD Then
        MsgBox "Velg i kolonne A fra rad " & F�RSTE_PERSONRAD & " og nedover.", vbExclamation
        Exit Sub
    End If
    personRow = personCell.Row

    kode = UCase$(Trim(InputBox("AktivitetsKODE (f.eks. TL, SIC, SAR):", "Aktivitetskode")))
    If Len(kode) = 0 Then Exit Sub
    If Not Sl�OppAktivitet(wsTyp, kode, beskrivelse, farge) Then
        MsgBox "Fant ikke koden i '" & ARK_OVERSIKT & "'.", vbCritical
        Exit Sub
    End If

    If Not HentDato("Startdato (dd.mm.����):", startDato) Then Exit Sub
    If Not HentDato("Sluttdato (dd.mm.����):", sluttDato) Then Exit Sub
    If sluttDato < startDato Then
        MsgBox "Sluttdato kan ikke v�re f�r startdato.", vbExclamation
        Exit Sub
    End If

    startCol = FinnKolonneForDato_Rad13(wsPlan, startDato, F�RSTE_DATAKOL, datoRad)
    sluttCol = FinnKolonneForDato_Rad13(wsPlan, sluttDato, F�RSTE_DATAKOL, datoRad)
    If startCol = 0 Or sluttCol = 0 Then
        MsgBox "Fant ikke start/sluttdato i rad " & datoRad & ".", vbCritical
        Exit Sub
    End If
    If sluttCol < startCol Then
        Dim t As Long, td As Date
        t = startCol: startCol = sluttCol: sluttCol = t
        td = startDato: startDato = sluttDato: sluttDato = td
    End If

    Set farger = HentAktivitetsFarger(wsTyp)
    m�lRad = FinnEllerOpprettLedigRad_UtenNavn(wsPlan, personRow, startCol, sluttCol, farger)
    If m�lRad = 0 Then
        MsgBox "Fant/skapte ikke ledig rad.", vbCritical
        Exit Sub
    End If

    kommentar = InputBox("Kommentar (valgfritt � vises i blokken):", "Kommentar")
    If Len(Trim$(kommentar)) > 0 Then
        visTekst = kode & " � " & Trim$(kommentar)
    Else
        visTekst = kode & " � " & beskrivelse
    End If

    ApplyBlockFormatting wsPlan, m�lRad, startCol, sluttCol, farge, visTekst, farger
End Sub

' ===================== M�TE 2 =====================
' Legger inn aktivitet i markert omr�de (�n blokk per valgt rad)
' v3: Lager ny under-rad hvis spennet overlapper annen aktivitet (ikke samme kode)
Public Sub LeggInnAktivitetP�Markering()
    Dim wsPlan As Worksheet, wsTyp As Worksheet
    Dim farger As Object
    Dim kode As String, beskrivelse As String, kommentar As String, visTekst As String
    Dim farge As Long
    Dim sel As Range, area As Range
    Dim r As Long, cMin As Long, cMax As Long
    Dim lastDatoCol As Long, m�lRad As Long, hovedRad As Long

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(ARK_PLAN)
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT)
    On Error GoTo 0
    If wsPlan Is Nothing Or wsTyp Is Nothing Then
        MsgBox "Finner ikke arkene '" & ARK_PLAN & "' og/eller '" & ARK_OVERSIKT & "'.", vbCritical
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Marker et omr�de i '" & ARK_PLAN & "' f�rst.", vbExclamation
        Exit Sub
    End If
    Set sel = Intersect(Selection, wsPlan.UsedRange)
    If sel Is Nothing Then
        MsgBox "Markeringen er tom.", vbExclamation
        Exit Sub
    End If

    kode = UCase$(Trim(InputBox("AktivitetsKODE (f.eks. TL, SIC, SAR):", "Aktivitetskode")))
    If Len(kode) = 0 Then Exit Sub
    If Not SlaaOppAktivitet(wsTyp, kode, beskrivelse, farge) Then
        MsgBox "Fant ikke koden i '" & ARK_OVERSIKT & "'.", vbCritical
        Exit Sub
    End If

    kommentar = InputBox("Kommentar (valgfritt � vises i blokken):", "Kommentar")
    If Len(Trim$(kommentar)) > 0 Then
        visTekst = kode & " � " & Trim$(kommentar)
    Else
        visTekst = kode & " � " & beskrivelse
    End If

    Set farger = HentAktivitetsFarger(wsTyp)

    Application.ScreenUpdating = False
    lastDatoCol = SisteDatoKolonne(wsPlan, datoRad)
    If lastDatoCol < F�RSTE_DATAKOL Then lastDatoCol = F�RSTE_DATAKOL

    For Each area In sel.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            If r < F�RSTE_PERSONRAD Then GoTo nesteRad
            cMin = Application.WorksheetFunction.Max(F�RSTE_DATAKOL, area.Column)
            cMax = Application.WorksheetFunction.Min(lastDatoCol, area.Column + area.Columns.Count - 1)
            If cMax < cMin Then GoTo nesteRad

            ' Bestem m�lrad: ny under-rad ved overlapp med annen aktivitet
            hovedRad = FinnHovedRad(wsPlan, r)
            If SpanHarAnnenAktivitet(wsPlan, r, cMin, cMax, farger, kode) Then
                m�lRad = FinnEllerOpprettLedigRad_UtenNavn(wsPlan, hovedRad, cMin, cMax, farger)
                If m�lRad = 0 Then GoTo nesteRad
            Else
                m�lRad = r
            End If

            ApplyBlockFormatting wsPlan, m�lRad, cMin, cMax, farge, visTekst, farger
nesteRad:
        Next r
    Next area

    Application.ScreenUpdating = True
End Sub

' ---------------- HJELPERE (Public der n�dvendig) ----------------

Public Function Sl�OppAktivitet(wsTyp As Worksheet, ByVal kode As String, _
                                ByRef beskrivelse As String, ByRef farge As Long) As Boolean
    Dim r As Long, lastRow As Long
    lastRow = wsTyp.Cells(wsTyp.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If UCase$(Trim$(wsTyp.Cells(r, 1).Value)) = UCase$(Trim$(kode)) Then
            beskrivelse = CStr(wsTyp.Cells(r, 2).Value)
            farge = wsTyp.Cells(r, 1).Interior.Color
            Sl�OppAktivitet = True
            Exit Function
        End If
    Next r
End Function

Public Function SlaaOppAktivitet(wsTyp As Worksheet, ByVal kode As String, _
                                 ByRef beskrivelse As String, ByRef farge As Long) As Boolean
    SlaaOppAktivitet = Sl�OppAktivitet(wsTyp, kode, beskrivelse, farge)
End Function

Private Function FinnKolonneForDato_Rad13(ws As Worksheet, ByVal d As Date, _
                                          ByVal firstDataCol As Long, ByVal headerRow As Long) As Long
    Dim lastCol As Long, c As Long, v
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    For c = firstDataCol To lastCol
        v = ws.Cells(headerRow, c).Value
        If IsDate(v) Then
            If CLng(CDate(v)) = CLng(d) Then
                FinnKolonneForDato_Rad13 = c
                Exit Function
            End If
        End If
    Next c
End Function

' Finn f�rste navnerad (hovedrad) over/lik gitt rad
Private Function FinnHovedRad(ws As Worksheet, ByVal rad As Long) As Long
    Dim r As Long
    For r = rad To F�RSTE_PERSONRAD Step -1
        If Len(Trim$(ws.Cells(r, 1).Value)) > 0 Then FinnHovedRad = r: Exit Function
    Next r
    FinnHovedRad = rad
End Function

' Overlapp med *annen* aktivitet i spennet?
' - Dersom vi finner fet tekst i spennet som **ikke** starter med samme kode � TRUE
' - Dersom vi finner aktivitetsfarge uten tekst � antar annen aktivitet � TRUE
' - Kun samme kode eller tomt � FALSE
Private Function SpanHarAnnenAktivitet(ws As Worksheet, ByVal r As Long, _
                                       ByVal cMin As Long, ByVal cMax As Long, _
                                       ByVal farger As Object, ByVal kode As String) As Boolean
    Dim c As Long, cel As Range, txt As String
    For c = cMin To cMax
        Set cel = ws.Cells(r, c)
        If Len(Trim$(cel.Value)) > 0 And cel.Font.Bold Then
            txt = CStr(cel.Value)
            If StrComp(Left$(Trim$(txt), Len(kode)), kode, vbTextCompare) <> 0 Then
                SpanHarAnnenAktivitet = True: Exit Function
            End If
        ElseIf cel.Interior.ColorIndex <> xlColorIndexNone Then
            If FargeN�rAktivitet(cel.Interior.Color, farger) Then
                SpanHarAnnenAktivitet = True: Exit Function
            End If
        End If
    Next c
End Function

Private Sub FinnPersonBlokk(ws As Worksheet, hovedRad As Long, _
                            ByRef blockStart As Long, ByRef blockEnd As Long)
    Dim r As Long, lastRow As Long, v
    blockStart = hovedRad: blockEnd = hovedRad
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = hovedRad + 1 To lastRow
        v = ws.Cells(r, 1).Value
        If Len(Trim$(v)) = 0 Then blockEnd = r Else Exit For
    Next r
End Sub

Public Function HentAktivitetsFarger(wsTyp As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, col As Long
    lastRow = wsTyp.Cells(wsTyp.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        col = wsTyp.Cells(r, 1).Interior.Color
        If col <> 0 Then
            If Not dict.exists(CStr(col)) Then dict.Add CStr(col), col
        End If
    Next r
    Set HentAktivitetsFarger = dict
End Function

Private Function FargeN�rAktivitet(col As Long, ByVal farger As Object, Optional tol As Long = 18) As Boolean
    Dim k As Variant, refCol As Long
    For Each k In farger.Keys
        refCol = CLng(farger(k))
        If FargeAvstand(col, refCol) <= tol Then
            FargeN�rAktivitet = True
            Exit Function
        End If
    Next k
End Function

Private Function FargeAvstand(c1 As Long, c2 As Long) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = c1 Mod 256: g1 = (c1 \ 256) Mod 256: b1 = (c1 \ 65536) Mod 256
    r2 = c2 Mod 256: g2 = (c2 \ 256) Mod 256: b2 = (c2 \ 65536) Mod 256
    FargeAvstand = Application.WorksheetFunction.Max(Abs(r1 - r2), Abs(g1 - g2), Abs(b1 - b2))
End Function

Private Function SpennErLedig(rng As Range, ByVal farger As Object) As Boolean
    Dim c As Range
    For Each c In rng.Cells
        If Len(Trim$(c.Value)) > 0 Then SpennErLedig = False: Exit Function
        If c.Interior.ColorIndex <> xlColorIndexNone Then
            If FargeN�rAktivitet(c.Interior.Color, farger) Then
                SpennErLedig = False: Exit Function
            End If
        End If
    Next c
    SpennErLedig = True
End Function

Public Function SisteDatoKolonne(ws As Worksheet, ByVal headerRow As Long) As Long
    SisteDatoKolonne = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
End Function

' LIM INN I **Modul 1 � Samlet v3** (eller nyere). Erstatt hele
' `FinnEllerOpprettLedigRad_UtenNavn` + legg til helper `NullstillTilHvitMedGrid`.

Private Function FinnEllerOpprettLedigRad_UtenNavn(ws As Worksheet, personRow As Long, _
                                                   startCol As Long, sluttCol As Long, _
                                                   ByVal farger As Object) As Long
    Dim blockStart As Long, blockEnd As Long, r As Long
    Dim rng As Range
    Dim lastCol As Long, c As Long, cel As Range

    FinnPersonBlokk ws, personRow, blockStart, blockEnd

    ' 1) Finn ledig rad i eksisterende blokk
    For r = blockStart To blockEnd
        Set rng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, sluttCol))
        If SpennErLedig(rng, farger) Then
            FinnEllerOpprettLedigRad_UtenNavn = r
            Exit Function
        End If
    Next r

    ' 2) Opprett ny under-rad under blokken � kopier KUN basisformat (kolbredd/rowheight),
    '    men nullstill ALLE datoceller til HVIT + NORMALT RUTENETT (ikke arv fra hovedrad)
    ws.Rows(blockEnd + 1).Insert Shift:=xlDown
    ' behold h�yde/nummerformater ved � kopiere radh�yde/kolbredder indirekte via formats,
    ' men vi skal uansett blanke ut datofeltene etterp�
    ws.Rows(blockStart).Copy
    ws.Rows(blockEnd + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Cells(blockEnd + 1, 1).ClearContents

    lastCol = SisteDatoKolonne(ws, datoRad)
    For c = F�RSTE_DATAKOL To lastCol
        Set cel = ws.Cells(blockEnd + 1, c)
        ' UANSETT hva som ble kopiert: sett hvit bakgrunn og heltrukne tynne kanter
        NullstillTilHvitMedGrid cel
    Next c

    FinnEllerOpprettLedigRad_UtenNavn = blockEnd + 1
End Function

Private Sub NullstillTilHvitMedGrid(cel As Range)
    ' Ingen diagonaler, hvit bakgrunn, heltrukne tynne ytterkanter
    cel.ClearComments
    cel.ClearContents
    cel.Font.Bold = False
    cel.Font.ColorIndex = xlColorIndexAutomatic
    cel.HorizontalAlignment = xlGeneral
    cel.VerticalAlignment = xlCenter
    cel.WrapText = False

    With cel.Interior
        .Pattern = xlSolid
        .TintAndShade = 0
        .Color = RGB(255, 255, 255)
        .PatternTintAndShade = 0
    End With

    cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
    cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone

    cel.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    cel.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    cel.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

    With cel.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous:  .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeRight)
        .LineStyle = xlContinuous:  .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeTop)
        .LineStyle = xlContinuous:  .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous:  .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
End Sub


Private Sub KopierBakgrunn(ByVal src As Range, ByVal dst As Range)
    With dst.Interior
        .Pattern = src.Interior.Pattern
        .TintAndShade = src.Interior.TintAndShade
        .Color = src.Interior.Color
        .PatternTintAndShade = src.Interior.PatternTintAndShade
    End With
End Sub

Public Sub ApplyBlockFormatting(ws As Worksheet, m�lRad As Long, _
                               startCol As Long, sluttCol As Long, _
                               farge As Long, visTekst As String, _
                               ByVal farger As Object)
    Dim rng As Range, startCell As Range, rngUnder As Range, c As Range, cb As Range
    Application.ScreenUpdating = False

    Set rng = ws.Range(ws.Cells(m�lRad, startCol), ws.Cells(m�lRad, sluttCol))
    Set startCell = ws.Cells(m�lRad, startCol)

    rng.ClearContents
    rng.ClearComments
    rng.Interior.Pattern = xlSolid
    rng.Interior.TintAndShade = 0
    rng.Interior.Color = farge

    rng.Borders.LineStyle = xlLineStyleNone
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
    End With
    rng.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

    If m�lRad < ws.Rows.Count Then
        For Each c In rng.Cells
            Set cb = c.Offset(1, 0)
            If cb.Row <= ws.Rows.Count Then
                If Len(Trim$(cb.Value)) = 0 And cb.Interior.Color = farge Then
                    KopierBakgrunn ws.Cells(m�lRad, cb.Column), cb
                End If
            End If
        Next c

        Set rngUnder = ws.Range(ws.Cells(m�lRad + 1, startCol), ws.Cells(m�lRad + 1, sluttCol))
        With rngUnder.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If

    startCell.Value = visTekst
    rng.HorizontalAlignment = xlCenterAcrossSelection
    rng.VerticalAlignment = xlCenter
    rng.WrapText = True
    rng.Font.Bold = True
    rng.Font.Color = IIf(ErLysFarge(farge), RGB(0, 0, 0), RGB(255, 255, 255))

    Application.ScreenUpdating = True
End Sub

Private Function ErLysFarge(col As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = col Mod 256: g = (col \ 256) Mod 256: b = (col \ 65536) Mod 256
    ErLysFarge = (0.299 * r + 0.587 * g + 0.114 * b) > 160
End Function

Private Function HentDato(prompt As String, ByRef d As Date) As Boolean
    Dim s As String
    s = Trim(InputBox(prompt, "Dato"))
    If Len(s) = 0 Then Exit Function
    On Error GoTo Feil
    d = CDate(s): HentDato = True
    Exit Function
Feil:
    MsgBox "Ugyldig dato: " & s, vbExclamation
End Function




