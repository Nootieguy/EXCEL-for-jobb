Attribute VB_Name = "Module4"

Option Explicit
'
' =================== MODUL 4  UVALGTE (v4.8 - DYNAMISK) ===================
' Ufordelte aktiviteter  liste i UVALGTE. Forhåndsvisning nederst i Planlegger.
' Når Person velges i tabellen, flyttes aktiviteten til riktig person i Planlegger.
'
' ENDRING v4.8: Bruker Named Ranges for dynamiske verdier (som Module1 og FjernMarkert)
' Dette tillater å legge til personer uten å måtte endre konstanter.

' ===== KONFIG =====
Private Const ARK_PLAN As String = "Planlegger"
Private Const ARK_OVERSIKT As String = "AKTIVITETSTYPER - OVERSIKT"
Private Const ARK_UVALGT As String = "UVALGTE"

' Panel
Private Const IP_ROW As Long = 2
Private Const PANEL_H As Single = 24

' Tabell
Private Const TBL_START_ROW As Long = 10
Private Const COL_KODE As Long = 1
Private Const COL_BESKR As Long = 2
Private Const COL_FRA As Long = 3
Private Const COL_TIL As Long = 4
Private Const COL_KOMM As Long = 5
Private Const COL_PERSON As Long = 6

' Farger
Private Const FARGE_HEADER As Long = &HE9D7B9
Private Const FARGE_PANEL As Long = &HEDF3FF
Private Const FARGE_PANEL_TITLE As Long = &HDDE7FF
Private Const FARGE_PREVIEW As Long = &HDCDCDC
Private Const FARGE_BTN As Long = &HE36C2E
Private Const FARGE_BTN_TXT As Long = &HFFFFFF

' Navn
Private Const NM_IP_KODE As String = "IP_KODE"
Private Const NM_IP_FRA As String = "IP_FRA"
Private Const NM_IP_TIL As String = "IP_TIL"
Private Const NM_IP_KOMM As String = "IP_KOMM"
Private Const NM_NM_KODER As String = "AKTIV_KODER"
Private Const NM_NM_PERSON As String = "PERSONLISTE"
Private Const NM_BTN_LEGG As String = "btnUvalgtLegg"
Private Const NM_BTN_PREV As String = "btnUvalgtPreview"
Private Const NM_CB_LEGG As String = "chkUvalgtLeggNå"

' =================== HJELPEFUNKSJONER FOR DYNAMISKE VERDIER ===================
' Disse henter verdier fra Named Ranges i Planlegger-arket

Private Function HentFørsteDatoKol() As Long
    HentFørsteDatoKol = Worksheets(ARK_PLAN).Range("FirstDate").Column
End Function

Private Function HentDatoRad() As Long
    HentDatoRad = Worksheets(ARK_PLAN).Range("FirstDate").Row
End Function

Private Function HentFørstePersonRad() As Long
    HentFørstePersonRad = Worksheets(ARK_PLAN).Range("PersonHeader").Row + 1
End Function

' =================== OPPSETT ===================

Public Sub Uvalgte_Oppsett()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsU As Worksheet, wsP As Worksheet, wsO As Worksheet

    On Error Resume Next
    Set wsU = wb.Worksheets(ARK_UVALGT)
    Set wsP = wb.Worksheets(ARK_PLAN)
    Set wsO = wb.Worksheets(ARK_OVERSIKT)
    On Error GoTo 0
    If wsP Is Nothing Or wsO Is Nothing Then
        MsgBox "Mangler ark: '" & ARK_PLAN & "' og/eller '" & ARK_OVERSIKT & "'.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    If wsU Is Nothing Then
        Set wsU = wb.Worksheets.Add(After:=wsP)
        wsU.Name = ARK_UVALGT
    End If
    wsU.Cells.Clear
    wsU.Columns("A:G").ColumnWidth = 16

    ' Toppstripe
    With wsU.Range("A1:G1")
        .Merge
        .Value = "UVALGTE  arbeidsbenk"
        .Interior.Color = FARGE_PANEL_TITLE
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
        Boxify .Cells
    End With

    ' Panel (uten Person)
    With wsU
        .Range("A" & IP_ROW & ":G" & (IP_ROW + 5)).Interior.Color = FARGE_PANEL
        .Range("A" & IP_ROW).Value = "Legg inn ufordelt aktivitet"
        .Range("A" & IP_ROW).Font.Bold = True
        .Range("A" & IP_ROW).Font.Size = 12

        .Range("A" & (IP_ROW + 1)).Value = "Aktivitetskode:"
        .Range("A" & (IP_ROW + 2)).Value = "Fra (dato):"
        .Range("A" & (IP_ROW + 3)).Value = "Til (dato):"
        .Range("E" & (IP_ROW + 1)).Value = "Kommentar:"

        TryDeleteName NM_IP_KODE: .Range("B" & (IP_ROW + 1)).Name = NM_IP_KODE
        TryDeleteName NM_IP_FRA:  .Range("B" & (IP_ROW + 2)).Name = NM_IP_FRA
        TryDeleteName NM_IP_TIL:  .Range("B" & (IP_ROW + 3)).Name = NM_IP_TIL
        TryDeleteName NM_IP_KOMM: .Range("F" & (IP_ROW + 1)).Name = NM_IP_KOMM

        Boxify .Range("A" & (IP_ROW + 1) & ":B" & (IP_ROW + 3))
        Boxify .Range("E" & (IP_ROW + 1) & ":F" & (IP_ROW + 1))
        .Range(NM_IP_FRA & "," & NM_IP_TIL).NumberFormat = "dd.mm.yyyy"
    End With

    ' Kode-rullegardin
    Dim lastO As Long: lastO = wsO.Cells(wsO.Rows.Count, 1).End(xlUp).Row
    TryDeleteName NM_NM_KODER
    ThisWorkbook.Names.Add Name:=NM_NM_KODER, RefersTo:=wsO.Range("A2:A" & lastO)
    With wsU.Range(NM_IP_KODE).Validation
        .Delete: .Add Type:=xlValidateList, Formula1:="=" & NM_NM_KODER
    End With

    ' Personliste i UVALGTE!K
    DefinerPersonListe wsP
    wsU.Columns("K").Hidden = True

    ' Rydd kontroller
    TryDeleteShape wsU, NM_BTN_LEGG
    TryDeleteShape wsU, NM_BTN_PREV
    TryDeleteShape wsU, NM_CB_LEGG

    ' Anker (ByRef-fix)
    Dim anchAdd As Range, anchPrev As Range
    Set anchAdd = wsU.Cells(IP_ROW + 1, 7)
    Set anchPrev = wsU.Cells(TBL_START_ROW + 1, 7)

    ' Checkbox
    Dim cb As Object
    Set cb = wsU.CheckBoxes.Add(wsU.Cells(IP_ROW + 2, 7).Left, wsU.Cells(IP_ROW + 2, 7).Top, 150, PANEL_H)
    With cb
        .Caption = " Legg til nå"
        .Name = NM_CB_LEGG
        .OnAction = "'" & ThisWorkbook.Name & "'!Uvalgte_PanelCheckToggle"
        .Value = False
    End With

    ' Tabellhoder
    With wsU
        .Range("A" & (TBL_START_ROW - 1) & ":F" & (TBL_START_ROW - 1)).Interior.Color = FARGE_HEADER
        .Range("A" & (TBL_START_ROW - 1)).Resize(1, 6).Value = _
            Array("Kode", "Beskrivelse", "Fra", "Til", "Kommentar", "Person")
        Boxify .Range("A" & (TBL_START_ROW - 1) & ":F" & (TBL_START_ROW - 1))
        .Range(.Cells(TBL_START_ROW, COL_FRA), .Cells(TBL_START_ROW + 800, COL_TIL)).NumberFormat = "dd.mm.yyyy"

        With .Range(.Cells(TBL_START_ROW, COL_PERSON), .Cells(TBL_START_ROW + 800, COL_PERSON))
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, Formula1:="=" & NM_NM_PERSON
        End With
    End With

    ' Knapp: oppfrisk forhåndsvisning
    LagKnapp wsU, NM_BTN_PREV, "Oppfrisk forhåndsvisning", _
             "'" & ThisWorkbook.Name & "'!Uvalgte_RefreshPreview", _
             anchPrev, 190, PANEL_H

    Uvalgte_RefreshPreview
    Application.ScreenUpdating = True
End Sub

' =================== HANDLERE ===================

Public Sub Uvalgte_PanelCheckToggle()
    Dim wsU As Worksheet: Set wsU = ThisWorkbook.Worksheets(ARK_UVALGT)
    Dim cb As Object, v As Variant
    On Error Resume Next
    Set cb = wsU.CheckBoxes(NM_CB_LEGG)
    If cb Is Nothing Then Set cb = wsU.OLEObjects(NM_CB_LEGG).Object
    On Error GoTo 0
    If cb Is Nothing Then Exit Sub

    v = cb.Value
    If VarType(v) = vbBoolean Then
        If v Then Uvalgte_LeggTilFraPanel: cb.Value = False
    Else
        If CLng(v) = 1 Then Uvalgte_LeggTilFraPanel: cb.Value = 0
    End If
End Sub

' Legg én rad i UVALGTE (uten person)
Public Sub Uvalgte_LeggTilFraPanel()
    Dim wsU As Worksheet: Set wsU = ThisWorkbook.Worksheets(ARK_UVALGT)
    Dim wsO As Worksheet: Set wsO = ThisWorkbook.Worksheets(ARK_OVERSIKT)

    Dim kode As String, fraD As Date, tilD As Date, komm As String
    Dim beskrivelse As String, f As Long

    kode = Trim$(wsU.Range(NM_IP_KODE).Value)
    fraD = wsU.Range(NM_IP_FRA).Value
    tilD = wsU.Range(NM_IP_TIL).Value
    komm = Trim$(wsU.Range(NM_IP_KOMM).Value)

    If Len(kode) = 0 Or Not IsDate(fraD) Or Not IsDate(tilD) Then
        MsgBox "Velg kode, fra- og til-dato.", vbExclamation: Exit Sub
    End If
    If tilD < fraD Then MsgBox "Til kan ikke være før Fra.", vbExclamation: Exit Sub

    If Not LookupAktivitet(wsO, UCase$(kode), beskrivelse, f) Then beskrivelse = ""

    Dim r As Long: r = NesteTomRad(wsU, COL_KODE, TBL_START_ROW)
    With wsU
        .Cells(r, COL_KODE).Value = UCase$(kode)
        .Cells(r, COL_BESKR).Value = beskrivelse
        .Cells(r, COL_FRA).Value = fraD
        .Cells(r, COL_TIL).Value = tilD
        .Cells(r, COL_KOMM).Value = komm
        .Cells(r, COL_PERSON).ClearContents
    End With

    wsU.Range(NM_IP_KOMM).ClearContents
    wsU.Range(NM_IP_FRA & "," & NM_IP_TIL).ClearContents

    Uvalgte_RefreshPreview
End Sub

' Flytt EN rad fra UVALGTE til personens kalender MED SMART OVERLAPP
Public Sub Uvalgte_TildelRad(ByVal r As Long)
    ' Hent dynamiske verdier
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    førsteDatoKol = HentFørsteDatoKol()
    datoRad = HentDatoRad()
    førstePersonRad = HentFørstePersonRad()
    
    Dim wsU As Worksheet: Set wsU = ThisWorkbook.Worksheets(ARK_UVALGT)
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    Dim wsO As Worksheet: Set wsO = ThisWorkbook.Worksheets(ARK_OVERSIKT)

    If r < TBL_START_ROW Then Exit Sub
    Dim kode As String, person As String, komm As String
    Dim fraD As Date, tilD As Date
    kode = UCase$(Trim$(wsU.Cells(r, COL_KODE).Value))
    person = Trim$(wsU.Cells(r, COL_PERSON).Value)
    fraD = wsU.Cells(r, COL_FRA).Value
    tilD = wsU.Cells(r, COL_TIL).Value
    komm = Trim$(wsU.Cells(r, COL_KOMM).Value)

    If Len(kode) = 0 Then MsgBox "Mangler kode i raden.", vbExclamation: Exit Sub
    If Len(person) = 0 Then MsgBox "Velg person i rullegardinen.", vbExclamation: Exit Sub
    If Not IsDate(fraD) Or Not IsDate(tilD) Then MsgBox "Ugyldig fra/til-dato.", vbExclamation: Exit Sub
    If CLng(tilD) < CLng(fraD) Then MsgBox "Til-dato kan ikke være før fra-dato.", vbExclamation: Exit Sub

    Dim beskrivelse As String, f As Long
    If Not LookupAktivitet(wsO, kode, beskrivelse, f) Then
        MsgBox "Fant ikke aktivitetskoden '" & kode & "' i oversikten.", vbExclamation
        Exit Sub
    End If

    Dim sCol As Long, eCol As Long, t As Long
    sCol = FinnDatoKolonne(wsP, fraD, datoRad)
    eCol = FinnDatoKolonne(wsP, tilD, datoRad)
    If sCol = 0 Or eCol = 0 Then
        MsgBox "Start-/sluttdato finnes ikke i dato-raden (rad " & datoRad & ").", vbExclamation
        Exit Sub
    End If
    If eCol < sCol Then t = sCol: sCol = eCol: eCol = t

    Dim personRow As Long
    personRow = FinnPersonRad(wsP, person, førstePersonRad)
    If personRow = 0 Then
        MsgBox "Fant ikke personen '" & person & "' i kolonne A i Planlegger.", vbExclamation
        Exit Sub
    End If

    Dim farger As Object
    Set farger = HentAktivitetsFarger(wsO)

    ' SMART: Finn ledig rad (lager ny kun ved overlapp)
    Dim målRad As Long
    målRad = FinnEllerOpprettLedigRad_UtenNavn(wsP, personRow, sCol, eCol, farger, førsteDatoKol, datoRad)
    If målRad = 0 Then
        MsgBox "Fant ikke ledig rad for perioden.", vbExclamation
        Exit Sub
    End If

    Dim visTekst As String
    visTekst = kode & IIf(Len(komm) > 0, "  " & komm, IIf(Len(beskrivelse) > 0, "  " & beskrivelse, ""))

    ApplyBlockFormatting_Safe wsP, målRad, sCol, eCol, f, visTekst

    RyddRadIUvalgte wsU, r
    
    ' Slett tomme under-rader
    Dim lastCol As Long: lastCol = SisteDatoKolonneU4(wsP, datoRad)
    SlettTommeUnderRaderIPerson wsP, personRow, lastCol, førsteDatoKol
    
    ' Refresh preview
    Uvalgte_RefreshPreview
End Sub

' Rydd en rad i UVALGTE-tabellen (gjenopprett rutenett)
Private Sub RyddRadIUvalgte(ws As Worksheet, ByVal rad As Long)
    Dim c As Long, cel As Range
    
    ' Rydd kolonne 1-6 (Kode, Beskrivelse, Fra, Til, Kommentar, Person)
    For c = COL_KODE To COL_PERSON
        Set cel = ws.Cells(rad, c)
        
        ' Slett innhold og kommentarer
        cel.ClearContents
        cel.ClearComments
        
        ' Nullstill formatering
        cel.Font.Bold = False
        cel.Font.ColorIndex = xlColorIndexAutomatic
        cel.HorizontalAlignment = xlGeneral
        cel.VerticalAlignment = xlCenter
        cel.WrapText = False
        cel.Interior.ColorIndex = xlColorIndexNone
        
        ' Gjenopprett validering for Person-kolonnen
        If c = COL_PERSON Then
            On Error Resume Next
            cel.Validation.Delete
            cel.Validation.Add Type:=xlValidateList, Formula1:="=" & NM_NM_PERSON
            On Error GoTo 0
        End If
        
        ' Sett standard rutenett (tynne linjer)
        cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
        cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
        
        With cel.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        End With
        With cel.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        End With
        With cel.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        End With
        With cel.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        End With
    Next c
End Sub

' Forhåndsvis alle uvalgte med SMART overlapp-håndtering (SUPER-OPTIMALISERT)
Public Sub Uvalgte_RefreshPreview()
    ' Hent dynamiske verdier
    Dim førsteDatoKol As Long, datoRad As Long
    førsteDatoKol = HentFørsteDatoKol()
    datoRad = HentDatoRad()
    
    Dim wsU As Worksheet: Set wsU = ThisWorkbook.Worksheets(ARK_UVALGT)
    Dim wsP As Worksheet: Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    Dim wsO As Worksheet: Set wsO = ThisWorkbook.Worksheets(ARK_OVERSIKT)

    Dim prevRow As Long: prevRow = Uvalgte_FinnEllerLagPreviewRad(wsP)
    Dim lastCol As Long: lastCol = SisteDatoKolonneU4(wsP, datoRad)

    ' KRITISK: Deaktiver ALT under oppdatering
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' STEG 1: Finn alle eksisterende preview-rader (RASK VERSJON)
    Dim lastRow As Long
    lastRow = prevRow
    Dim maxSearch As Long: maxSearch = prevRow + 20
    
    Dim r As Long
    For r = prevRow + 1 To maxSearch
        If Len(Trim$(wsP.Cells(r, 1).Value)) > 0 Then Exit For
        lastRow = r
    Next r
    
    ' STEG 2: Rydd BARE eksisterende rader (batch-rensing)
    Dim c As Long
    If lastRow >= prevRow Then
        Dim rngRydd As Range
        Set rngRydd = wsP.Range(wsP.Cells(prevRow, førsteDatoKol), wsP.Cells(lastRow, lastCol))
        
        rngRydd.ClearContents
        rngRydd.Interior.Color = RGB(255, 255, 255)
        rngRydd.Font.Bold = False
        rngRydd.Font.ColorIndex = xlColorIndexAutomatic
        
        ' Sett grid på alle celler samtidig (RASK)
        With rngRydd.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlColorIndexAutomatic
        End With
    End If

    ' STEG 3: Tell aktiviteter
    Dim lr As Long: lr = wsU.Cells(wsU.Rows.Count, COL_KODE).End(xlUp).Row
    If lr < TBL_START_ROW Then
        SlettTommePreviewRaderSmart wsP, prevRow, lastCol, førsteDatoKol
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' STEG 4: Plasser aktiviteter - LAG RADER KUN VED OVERLAPP
    Dim rowIdx As Long, beskrivelse As String, f As Long
    Dim sCol As Long, eCol As Long, t As Long, visTekst As String
    Dim målRad As Long
    
    For rowIdx = TBL_START_ROW To lr
        If Len(Trim$(wsU.Cells(rowIdx, COL_KODE).Value)) > 0 Then
            sCol = FinnDatoKolonne(wsP, wsU.Cells(rowIdx, COL_FRA).Value, datoRad)
            eCol = FinnDatoKolonne(wsP, wsU.Cells(rowIdx, COL_TIL).Value, datoRad)
            
            If sCol > 0 And eCol > 0 Then
                If eCol < sCol Then t = sCol: sCol = eCol: eCol = t
                
                ' Finn ledig rad (lager ny kun ved overlapp)
                målRad = FinnLedigPreviewRadSmart(wsP, prevRow, sCol, eCol, førsteDatoKol)
                
                If målRad > 0 Then
                    If Not LookupAktivitet(wsO, UCase$(wsU.Cells(rowIdx, COL_KODE).Value), beskrivelse, f) Then
                        f = FARGE_PREVIEW
                    End If
                    
                    visTekst = UCase$(wsU.Cells(rowIdx, COL_KODE).Value) & _
                               IIf(Len(Trim$(wsU.Cells(rowIdx, COL_KOMM).Value)) > 0, "  " & Trim$(wsU.Cells(rowIdx, COL_KOMM).Value), _
                                  IIf(Len(beskrivelse) > 0, "  " & beskrivelse, ""))
                    
                    ApplyBlockFormatting_Safe wsP, målRad, sCol, eCol, f, visTekst
                End If
            End If
        End If
    Next rowIdx
    
    ' STEG 5: Slett tomme under-rader
    SlettTommePreviewRaderSmart wsP, prevRow, lastCol, førsteDatoKol
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' =================== HJELPERE ===================

Private Sub Boxify(r As Range)
    With r.Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
End Sub

' Sjekk om spennet har overlapp med ANNEN aktivitet (ikke samme kode)
Private Function SpanHarAnnenAktivitet_U4(ws As Worksheet, ByVal rad As Long, _
                                          ByVal cMin As Long, ByVal cMax As Long, _
                                          ByVal farger As Object, ByVal kode As String) As Boolean
    Dim c As Long, cel As Range, txt As String
    For c = cMin To cMax
        Set cel = ws.Cells(rad, c)
        ' Sjekk om det er fet tekst som ikke starter med samme kode
        If Len(Trim$(cel.Value)) > 0 And cel.Font.Bold Then
            txt = CStr(cel.Value)
            If StrComp(Left$(Trim$(txt), Len(kode)), kode, vbTextCompare) <> 0 Then
                SpanHarAnnenAktivitet_U4 = True
                Exit Function
            End If
        ' Sjekk om det er aktivitetsfarge (ikke hvit)
        ElseIf cel.Interior.ColorIndex <> xlColorIndexNone Then
            If FargeNærAktivitet_U4(cel.Interior.Color, farger) Then
                SpanHarAnnenAktivitet_U4 = True
                Exit Function
            End If
        End If
    Next c
End Function

' Hent alle aktivitetsfarger fra oversiktsarket
Private Function HentAktivitetsFarger(wsTyp As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long, r As Long, col As Long
    lastRow = wsTyp.Cells(wsTyp.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        col = wsTyp.Cells(r, 1).Interior.Color
        If col <> 0 And col <> RGB(255, 255, 255) Then
            If Not dict.exists(CStr(col)) Then dict.Add CStr(col), col
        End If
    Next r
    Set HentAktivitetsFarger = dict
End Function

' Sjekk om en farge er nær en aktivitetsfarge
Private Function FargeNærAktivitet_U4(col As Long, ByVal farger As Object, Optional tol As Long = 18) As Boolean
    If col = RGB(255, 255, 255) Then Exit Function ' Hvit er ikke aktivitet
    Dim k As Variant, refCol As Long
    For Each k In farger.Keys
        refCol = CLng(farger(k))
        If FargeAvstand_U4(col, refCol) <= tol Then
            FargeNærAktivitet_U4 = True
            Exit Function
        End If
    Next k
End Function

' Beregn fargeAvstand mellom to farger
Private Function FargeAvstand_U4(c1 As Long, c2 As Long) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = c1 Mod 256: g1 = (c1 \ 256) Mod 256: b1 = (c1 \ 65536) Mod 256
    r2 = c2 Mod 256: g2 = (c2 \ 256) Mod 256: b2 = (c2 \ 65536) Mod 256
    FargeAvstand_U4 = Application.WorksheetFunction.Max(Abs(r1 - r2), Abs(g1 - g2), Abs(b1 - b2))
End Function

Private Sub LagKnapp(ws As Worksheet, navn As String, txt As String, _
                     makro As String, putAt As Range, _
                     w As Single, h As Single)
    On Error Resume Next: ws.Shapes(navn).Delete: On Error GoTo 0
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, putAt.Left, putAt.Top, w, h)
    With shp
        .Name = navn
        .Fill.ForeColor.RGB = FARGE_BTN
        .Line.ForeColor.RGB = FARGE_BTN
        .TextFrame2.TextRange.Characters.Text = txt
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FARGE_BTN_TXT
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Adjustments.Item(1) = 0.2
        .OnAction = makro
    End With
End Sub

' Sammenhengende personliste i UVALGTE!K (for stabil validering)
Private Sub DefinerPersonListe(wsP As Worksheet)
    ' Hent dynamiske verdier
    Dim førstePersonRad As Long
    førstePersonRad = HentFørstePersonRad()
    
    Dim wsU As Worksheet: Set wsU = ThisWorkbook.Worksheets(ARK_UVALGT)
    Dim lastRow As Long, r As Long, w As Long
    wsU.Columns("K").ClearContents

    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    w = 2
    For r = førstePersonRad To lastRow
        If Len(Trim$(wsP.Cells(r, 1).Value)) > 0 Then
            wsU.Cells(w, "K").Value = Trim$(wsP.Cells(r, 1).Value)
            w = w + 1
        End If
    Next r

    TryDeleteName NM_NM_PERSON
    If w > 2 Then
        ThisWorkbook.Names.Add Name:=NM_NM_PERSON, _
            RefersTo:=wsU.Range(wsU.Cells(2, "K"), wsU.Cells(w - 1, "K"))
    Else
        wsU.Range("K2").Value = ""
        ThisWorkbook.Names.Add Name:=NM_NM_PERSON, RefersTo:=wsU.Range("K2")
    End If
End Sub

Private Function Uvalgte_FinnEllerLagPreviewRad(wsP As Worksheet) As Long
    ' Hent dynamiske verdier
    Dim førstePersonRad As Long
    førstePersonRad = HentFørstePersonRad()
    
    Dim lastRow As Long: lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = lastRow To førstePersonRad Step -1
        If UCase$(Trim$(wsP.Cells(r, 1).Value)) Like "UVALGTE  FORHÅNDSVISNING*" Then
            Uvalgte_FinnEllerLagPreviewRad = r: Exit Function
        End If
    Next r
    wsP.Rows(lastRow + 1).Insert Shift:=xlDown
    wsP.Rows(lastRow + 2).Insert Shift:=xlDown
    wsP.Cells(lastRow + 2, 1).Value = "UVALGTE  forhåndsvisning"
    wsP.Cells(lastRow + 2, 1).Font.Bold = True
    Uvalgte_FinnEllerLagPreviewRad = lastRow + 2
End Function

Private Function NesteTomRad(ws As Worksheet, ByVal col As Long, ByVal startRow As Long) As Long
    Dim lr As Long: lr = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    NesteTomRad = IIf(lr < startRow, startRow, lr + 1)
End Function

Private Function FinnDatoKolonne(ws As Worksheet, d As Date, datoRad As Long) As Long
    ' Hent dynamiske verdier
    Dim førsteDatoKol As Long
    førsteDatoKol = HentFørsteDatoKol()
    
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(datoRad, ws.Columns.Count).End(xlToLeft).Column
    For c = førsteDatoKol To lastCol
        If IsDate(ws.Cells(datoRad, c).Value) Then
            If CLng(CDate(ws.Cells(datoRad, c).Value)) = CLng(d) Then
                FinnDatoKolonne = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function FinnPersonRad(ws As Worksheet, ByVal navn As String, førstePersonRad As Long) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = førstePersonRad To lastRow
        If StrComp(Trim$(ws.Cells(r, 1).Value), Trim$(navn), vbTextCompare) = 0 Then
            FinnPersonRad = r: Exit Function
        End If
    Next r
End Function

Public Function SisteDatoKolonneU4(ws As Worksheet, ByVal headerRow As Long) As Long
    SisteDatoKolonneU4 = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Sub ResetToWhiteGrid(ByVal cel As Range)
    ' Fjern innhold først
    cel.ClearContents
    cel.ClearComments
    
    ' Reset font
    cel.Font.Bold = False
    cel.Font.ColorIndex = xlColorIndexAutomatic
    cel.HorizontalAlignment = xlGeneral
    cel.VerticalAlignment = xlCenter
    cel.WrapText = False
    
    ' Hvit bakgrunn
    With cel.Interior
        .Pattern = xlSolid
        .TintAndShade = 0
        .Color = RGB(255, 255, 255)
        .PatternTintAndShade = 0
    End With
    
    ' Fjern diagonaler
    cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
    cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
    
    ' VIKTIG: Ikke fjern eksisterende borders - bare sett dem til standard
    ' Dette bevarer bunnlinjen til raden over
    With cel.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With cel.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With cel.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    With cel.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
End Sub

Private Sub TryDeleteName(ByVal nm As String)
    On Error Resume Next: ThisWorkbook.Names(nm).Delete: On Error GoTo 0
End Sub

Private Sub TryDeleteShape(ws As Worksheet, ByVal shpName As String)
    On Error Resume Next: ws.Shapes(shpName).Delete: On Error GoTo 0
End Sub

' SlåOppAktivitet fallback (tolerant)
Private Function LookupAktivitet(wsTyp As Worksheet, ByVal kode As String, _
                                 ByRef beskrivelse As String, ByRef farge As Long) As Boolean
    On Error GoTo Lokal
    beskrivelse = vbNullString: farge = 0
    LookupAktivitet = Module1.SlåOppAktivitet(wsTyp, kode, beskrivelse, farge)
    Exit Function
Lokal:
    On Error GoTo 0
    Dim r As Long, lastRow As Long, k As String
    lastRow = wsTyp.Cells(wsTyp.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        k = UCase$(Trim$(wsTyp.Cells(r, 1).Value))
        If k = UCase$(Trim$(kode)) Then
            beskrivelse = CStr(wsTyp.Cells(r, 2).Value)
            farge = wsTyp.Cells(r, 1).Interior.Color
            LookupAktivitet = True
            Exit Function
        End If
    Next r
End Function

' ========= FORMATTERINGS-WRAPPER =========

Private Sub ApplyBlockFormatting_Safe(ws As Worksheet, målRad As Long, _
    startCol As Long, sluttCol As Long, farge As Long, visTekst As String)

    Dim ok As Boolean
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!ApplyBlockFormatting", _
                    ws, målRad, startCol, sluttCol, farge, visTekst
    ok = (Err.Number = 0): Err.Clear
    If Not ok Then
        Application.Run "'" & ThisWorkbook.Name & "'!ApplyBlockFormatting", _
                        ws, målRad, startCol, sluttCol, farge, visTekst, Empty
        ok = (Err.Number = 0): Err.Clear
    End If
    On Error GoTo 0
    If Not ok Then ApplyBlockFormattingU4 ws, målRad, startCol, sluttCol, farge, visTekst
End Sub

Private Sub ApplyBlockFormattingU4(ws As Worksheet, målRad As Long, _
    startCol As Long, sluttCol As Long, farge As Long, visTekst As String)

    Dim rng As Range, startCell As Range, rngUnder As Range
    Application.ScreenUpdating = False

    Set rng = ws.Range(ws.Cells(målRad, startCol), ws.Cells(målRad, sluttCol))
    Set startCell = ws.Cells(målRad, startCol)

    rng.ClearContents
    rng.ClearComments
    rng.Interior.Pattern = xlSolid
    rng.Interior.TintAndShade = 0
    rng.Interior.Color = farge

    ' Fjern alle borders først
    rng.Borders.LineStyle = xlLineStyleNone
    
    ' Sett kraftige ytterkanter på blokken
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
    
    ' Ingen borders inni blokken
    rng.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

    ' KRITISK: Gjenopprett bunnlinje på raden UNDER blokken
    If målRad < ws.Rows.Count Then
        Set rngUnder = ws.Range(ws.Cells(målRad + 1, startCol), ws.Cells(målRad + 1, sluttCol))
        With rngUnder.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)  ' SVART linje (ikke xlColorIndexAutomatic)
        End With
    End If

    startCell.Value = visTekst
    rng.HorizontalAlignment = xlCenterAcrossSelection
    rng.VerticalAlignment = xlCenter
    rng.WrapText = True
    rng.Font.Bold = True
    rng.Font.Color = IIf(U4_ErLysFarge(farge), RGB(0, 0, 0), RGB(255, 255, 255))

    Application.ScreenUpdating = True
End Sub

Private Function U4_ErLysFarge(col As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = col Mod 256: g = (col \ 256) Mod 256: b = (col \ 65536) Mod 256
    U4_ErLysFarge = (0.299 * r + 0.587 * g + 0.114 * b) > 160
End Function

' ========= FALLBACK: FinnEllerOpprettLedigRad_UtenNavn =========
' Signaturen matcher Modul 1. Om originalen finnes brukes den; ellers brukes denne.
Private Function FinnEllerOpprettLedigRad_UtenNavn(ws As Worksheet, _
    personRow As Long, startCol As Long, sluttCol As Long, _
    Optional ByVal farger As Object, _
    Optional førsteDatoKol As Long = 0, _
    Optional datoRad As Long = 0) As Long

    ' Hent dynamiske verdier hvis ikke oppgitt
    If førsteDatoKol = 0 Then førsteDatoKol = HentFørsteDatoKol()
    If datoRad = 0 Then datoRad = HentDatoRad()
    
    ' Forsøk å kalle originalen hvis den finnes
    On Error Resume Next
    FinnEllerOpprettLedigRad_UtenNavn = Application.Run("'" & ThisWorkbook.Name & _
        "'!FinnEllerOpprettLedigRad_UtenNavn", ws, personRow, startCol, sluttCol, farger)
    If Err.Number = 0 And FinnEllerOpprettLedigRad_UtenNavn <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear: On Error GoTo 0

    ' Lokal implementasjon
    Dim blockStart As Long, blockEnd As Long, r As Long
    blockStart = personRow: blockEnd = personRow
    
    ' Finn hele personblokken
    Do While Len(Trim$(ws.Cells(blockEnd + 1, 1).Value)) = 0 And blockEnd < ws.Rows.Count
        blockEnd = blockEnd + 1
    Loop

    ' Finn ledig rad i blokken
    Dim c As Long, fri As Boolean, cel As Range
    For r = blockStart To blockEnd
        fri = True
        For c = startCol To sluttCol
            Set cel = ws.Cells(r, c)
            ' Sjekk om cellen har tekst
            If Len(Trim$(cel.Value)) > 0 Then
                fri = False: Exit For
            End If
            ' Sjekk om cellen har aktivitetsfarge (ikke hvit)
            If cel.Interior.ColorIndex <> xlColorIndexNone Then
                If cel.Interior.Color <> RGB(255, 255, 255) Then
                    ' Hvis vi har farger-objekt, sjekk om det er aktivitetsfarge
                    If Not farger Is Nothing Then
                        If FargeNærAktivitet_U4(cel.Interior.Color, farger) Then
                            fri = False: Exit For
                        End If
                    Else
                        ' Uten farger-objekt, anse all ikke-hvit farge som opptatt
                        fri = False: Exit For
                    End If
                End If
            End If
        Next c
        If fri Then
            FinnEllerOpprettLedigRad_UtenNavn = r
            Exit Function
        End If
    Next r

    ' Opprett ny under-rad etter blokken
    ws.Rows(blockEnd + 1).Insert Shift:=xlDown
    ws.Rows(blockStart).Copy
    ws.Rows(blockEnd + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ws.Cells(blockEnd + 1, 1).ClearContents
    
    ' Nullstill alle datoceller til hvit med grid
    Dim lastCol As Long
    lastCol = SisteDatoKolonneU4(ws, datoRad)
    For c = førsteDatoKol To lastCol
        NullstillCelleTilHvitMedGrid ws.Cells(blockEnd + 1, c)
    Next c
    
    FinnEllerOpprettLedigRad_UtenNavn = blockEnd + 1
End Function

' Nullstill celle til hvit med normalt rutenett
Private Sub NullstillCelleTilHvitMedGrid(ByVal cel As Range)
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

    With cel.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(0, 0, 0)
    End With
End Sub

' =================== SMART OVERLAPP-HÅNDTERING (KUN NÅR NØDVENDIG) ===================

' Finn ledig rad - LAG KUN NY RAD VED OVERLAPP
Private Function FinnLedigPreviewRadSmart(ws As Worksheet, prevRow As Long, _
                                          startCol As Long, sluttCol As Long, _
                                          førsteDatoKol As Long) As Long
    Dim r As Long, c As Long, ledig As Boolean
    Dim lastRow As Long
    
    ' Finn hvor preview-blokken slutter
    lastRow = prevRow
    Do While lastRow < ws.Rows.Count
        If Len(Trim$(ws.Cells(lastRow + 1, 1).Value)) > 0 Then Exit Do
        lastRow = lastRow + 1
    Loop
    
    ' Sjekk rad for rad fra prevRow og nedover
    For r = prevRow To lastRow
        ledig = True
        For c = startCol To sluttCol
            If Len(Trim$(ws.Cells(r, c).Value)) > 0 Then
                ledig = False
                Exit For
            End If
            If ws.Cells(r, c).Interior.ColorIndex <> xlColorIndexNone And _
               ws.Cells(r, c).Interior.Color <> RGB(255, 255, 255) Then
                ledig = False
                Exit For
            End If
        Next c
        
        If ledig Then
            FinnLedigPreviewRadSmart = r
            Exit Function
        End If
    Next r
    
    ' INGEN ledig rad - LAG NY RAD (kun én)
    Dim nyRad As Long, datoRad As Long
    nyRad = lastRow + 1
    datoRad = HentDatoRad()
    
    ws.Rows(nyRad).Insert Shift:=xlDown
    
    ' Kopier format fra prevRow
    ws.Rows(prevRow).Copy
    ws.Rows(nyRad).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    ' Tøm kolonne A
    ws.Cells(nyRad, 1).ClearContents
    
    ' Nullstill alle datoceller
    Dim lastCol As Long
    lastCol = SisteDatoKolonneU4(ws, datoRad)
    For c = førsteDatoKol To lastCol
        ResetToWhiteGrid ws.Cells(nyRad, c)
    Next c
    
    FinnLedigPreviewRadSmart = nyRad
End Function

' Slett BARE tomme under-rader (OPTIMALISERT)
Private Sub SlettTommePreviewRaderSmart(ws As Worksheet, prevRow As Long, lastCol As Long, førsteDatoKol As Long)
    Dim r As Long, tom As Boolean, c As Long
    Dim lastRow As Long
    Dim maxSearch As Long
    
    ' Sett maks søkegrense
    maxSearch = prevRow + 50
    If maxSearch > ws.Rows.Count Then maxSearch = ws.Rows.Count
    
    ' Finn hvor preview-blokken slutter (sikker versjon)
    lastRow = prevRow
    For r = prevRow + 1 To maxSearch
        If Len(Trim$(ws.Cells(r, 1).Value)) > 0 Then Exit For
        lastRow = r
    Next r
    
    ' Slett tomme under-rader (IKKE prevRow selv)
    For r = lastRow To prevRow + 1 Step -1
        tom = True
        For c = førsteDatoKol To lastCol
            If Len(Trim$(ws.Cells(r, c).Value)) > 0 Then
                tom = False
                Exit For
            End If
            If ws.Cells(r, c).Interior.ColorIndex <> xlColorIndexNone And _
               ws.Cells(r, c).Interior.Color <> RGB(255, 255, 255) Then
                tom = False
                Exit For
            End If
        Next c
        
        If tom Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

' Slett tomme under-rader i en personblokk (OPTIMALISERT)
Private Sub SlettTommeUnderRaderIPerson(ws As Worksheet, personRow As Long, lastCol As Long, førsteDatoKol As Long)
    Dim blockEnd As Long
    Dim r As Long, c As Long, tom As Boolean
    
    ' Finn slutten av personblokken
    blockEnd = personRow
    Do While blockEnd < ws.Rows.Count
        If Len(Trim$(ws.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
        blockEnd = blockEnd + 1
    Loop
    
    ' Slett tomme under-rader (ikke hovedraden)
    For r = blockEnd To personRow + 1 Step -1
        tom = True
        For c = førsteDatoKol To lastCol
            If Len(Trim$(ws.Cells(r, c).Value)) > 0 Then
                tom = False
                Exit For
            End If
            If ws.Cells(r, c).Interior.ColorIndex <> xlColorIndexNone And _
               ws.Cells(r, c).Interior.Color <> RGB(255, 255, 255) Then
                tom = False
                Exit For
            End If
        Next c
        
        If tom Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub




