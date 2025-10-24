' ============================================================================
' ARK1: PLANLEGGER - WORKSHEET EVENT HANDLER
' ============================================================================
' FORMÅL:
'   - Markerer dagens dato når arket aktiveres
'   - Beskytter aktivitetsnavn mot utilsiktet sletting
'
' FUNKSJONALITET:
'   Auto-restaurering av aktivitetsnavn:
'   - Hvis bruker sletter FØRSTE celle i aktivitet → auto-restaurer navnet
'   - Hvis bruker sletter MIDT i aktivitet → tillat det (for gaps/splits)
'   - Fungerer med Delete, Cut, og markerte områder
'
' DETEKSJON AV FØRSTE CELLE:
'   En celle er første celle hvis:
'   - Cellen har aktivitetsfarge (ikke hvit/grå)
'   - Cellen til venstre har ANNEN farge (eller det ikke finnes celle til venstre)
' ============================================================================

Option Explicit

' Debug-modus (sett til True for logging til Immediate Window)
Private Const DEBUG_MODE As Boolean = False

' Rekursjons-beskyttelse
Private behandlerEndring As Boolean

' ============================================================================
' EVENT: Når arket aktiveres
' ============================================================================
Private Sub Worksheet_Activate()
    modDagensOutline.MarkerDagensDato_Bla
End Sub

' ============================================================================
' EVENT: Når bruker endrer en celle (Delete, Cut, etc.)
' ============================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Slutt

    ' STEG 1: Sjekk rekursjon (unngå loop)
    If behandlerEndring Then Exit Sub

    ' STEG 2: Valider at vi har Target
    If Target Is Nothing Then Exit Sub

    ' STEG 3: Hent dynamiske verdier fra Named Ranges
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    On Error Resume Next
    førsteDatoKol = Me.Range("FirstDate").Column
    datoRad = Me.Range("FirstDate").Row
    førstePersonRad = Me.Range("PersonHeader").Row + 1
    On Error GoTo Slutt

    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then Exit Sub

    ' STEG 4: Sjekk om endringen er i dato-området
    Dim lastCol As Long
    lastCol = Me.Cells(datoRad, Me.Columns.Count).End(xlToLeft).Column

    Dim datoOmråde As Range
    Set datoOmråde = Me.Range(Me.Cells(førstePersonRad, førsteDatoKol), _
                               Me.Cells(Me.Rows.Count, lastCol))

    Dim relevantEndring As Range
    Set relevantEndring = Intersect(Target, datoOmråde)

    If relevantEndring Is Nothing Then Exit Sub

    ' STEG 5: Sett beskyttelsesflagg og disable events
    behandlerEndring = True
    Application.EnableEvents = False

    ' STEG 6: Behandle hver endret celle
    Dim c As Range
    For Each c In relevantEndring.Cells
        Call BehandleSlettetCelle(c, førsteDatoKol, datoRad)
    Next c

    ' STEG 6B: Håndter splits (når hvite celler lages midt i aktivitet)
    For Each c In relevantEndring.Cells
        Call DetekterOgHåndterSplit(c, førsteDatoKol, datoRad)
    Next c

Slutt:
    ' STEG 7: Re-enable events og fjern beskyttelsesflagg
    Application.EnableEvents = True
    behandlerEndring = False

    ' Debug-logging ved feil
    If Err.Number <> 0 Then
        Debug.Print "Feil i Worksheet_Change (Ark1): " & Err.Description
        Err.Clear
    End If
End Sub

' ============================================================================
' HOVEDLOGIKK: Behandle en slettet celle
' ============================================================================
Private Sub BehandleSlettetCelle(ByVal cel As Range, ByVal førsteDatoKol As Long, _
                                  ByVal datoRad As Long)
    ' KRITERIE 1: Cellen må være tom (etter sletting)
    If Len(Trim$(cel.Value)) > 0 Then Exit Sub

    ' KRITERIE 2: Cellen må ha aktivitetsfarge (ikke hvit/grå)
    If Not HarAktivitetsfarge(cel) Then Exit Sub

    ' KRITERIE 3: Cellen må være FØRSTE celle i blokken
    If Not ErFørsteCelleIBlokk(cel, førsteDatoKol) Then
        If DEBUG_MODE Then Debug.Print "  → Ikke første celle - tillater split/gap"
        Exit Sub
    End If

    If DEBUG_MODE Then
        Debug.Print String(80, "=")
        Debug.Print "AUTO-RESTAURERING TRIGGET [" & Now & "]"
        Debug.Print "Celle: " & cel.Address & " (rad " & cel.Row & ", kol " & cel.Column & ")"
    End If

    ' STEG 1: Slå opp aktivitetskode fra farge
    Dim aktivKode As String, aktivBeskr As String
    aktivKode = SlåOppAktivitetFraFarge(cel.Interior.Color, aktivBeskr)

    If Len(aktivKode) = 0 Then
        If DEBUG_MODE Then Debug.Print "  → Fant ikke aktivitetskode for farge " & cel.Interior.Color
        Exit Sub
    End If

    If DEBUG_MODE Then
        Debug.Print "  → Funnet aktivitet: " & aktivKode & " - " & aktivBeskr
    End If

    ' STEG 2: Restaurer navnet i cellen
    Call RestaurerAktivitetsNavn(cel, aktivKode, aktivBeskr)

    If DEBUG_MODE Then
        Debug.Print "  ✓ Navn restaurert!"
        Debug.Print String(80, "=")
    End If
End Sub

' ============================================================================
' SPLIT-DETEKSJON: Sjekk om slettet celle har skapt en split
' ============================================================================
Private Sub DetekterOgHåndterSplit(ByVal cel As Range, ByVal førsteDatoKol As Long, _
                                    ByVal datoRad As Long)
    ' KRITERIE 1: Cellen må være tom og hvit (etter sletting)
    If Len(Trim$(cel.Value)) > 0 Then Exit Sub
    If HarAktivitetsfarge(cel) Then Exit Sub  ' Har fortsatt aktivitetsfarge = ikke split

    ' KRITERIE 2: Må ha aktivitetsfarge til venstre ELLER høyre (split-indikator)
    Dim harAktivVenstre As Boolean, harAktivHøyre As Boolean
    Dim venstreCelle As Range, høyreCelle As Range

    ' Sjekk venstre
    If cel.Column > førsteDatoKol Then
        Set venstreCelle = cel.Offset(0, -1)
        harAktivVenstre = HarAktivitetsfarge(venstreCelle)
    End If

    ' Sjekk høyre
    If cel.Column < Me.Columns.Count Then
        Set høyreCelle = cel.Offset(0, 1)
        harAktivHøyre = HarAktivitetsfarge(høyreCelle)
    End If

    ' Hvis ingen aktivitetsfarge rundt → ikke en split
    If Not harAktivVenstre And Not harAktivHøyre Then Exit Sub

    ' Dette er en split! Håndter den
    If DEBUG_MODE Then
        Debug.Print String(80, "=")
        Debug.Print "SPLIT DETEKTERT [" & Now & "]"
        Debug.Print "Celle: " & cel.Address & " (rad " & cel.Row & ", kol " & cel.Column & ")"
    End If

    Call HåndterAktivitetSplit(cel, førsteDatoKol, datoRad)

    If DEBUG_MODE Then
        Debug.Print "  ✓ Split håndtert!"
        Debug.Print String(80, "=")
    End If
End Sub

' ============================================================================
' SPLIT-HÅNDTERING: Re-formaterer aktivitet som er splittet
' ============================================================================
Private Sub HåndterAktivitetSplit(ByVal hvitCelle As Range, ByVal førsteDatoKol As Long, _
                                   ByVal datoRad As Long)
    Dim rad As Long
    Dim aktivFarge As Long
    Dim aktivKode As String, aktivBeskr As String

    rad = hvitCelle.Row

    ' STEG 1: Finn aktivitetsfargen (fra venstre eller høyre nabo)
    If hvitCelle.Column > førsteDatoKol And HarAktivitetsfarge(hvitCelle.Offset(0, -1)) Then
        aktivFarge = hvitCelle.Offset(0, -1).Interior.Color
    ElseIf hvitCelle.Column < Me.Columns.Count And HarAktivitetsfarge(hvitCelle.Offset(0, 1)) Then
        aktivFarge = hvitCelle.Offset(0, 1).Interior.Color
    Else
        Exit Sub  ' Ingen aktivitetsfarge funnet
    End If

    ' STEG 2: Slå opp aktivitetskode
    aktivKode = SlåOppAktivitetFraFarge(aktivFarge, aktivBeskr)
    If Len(aktivKode) = 0 Then Exit Sub

    If DEBUG_MODE Then
        Debug.Print "  → Aktivitet: " & aktivKode & " - " & aktivBeskr
        Debug.Print "  → Aktivfarge: " & aktivFarge
    End If

    ' STEG 3: Skann hele raden og identifiser alle segmenter
    Dim lastCol As Long
    lastCol = Me.Cells(datoRad, Me.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    Dim inSegment As Boolean
    Dim segmentStart As Long, segmentEnd As Long
    Dim segmentNr As Integer

    segmentNr = 0
    inSegment = False

    For c = førsteDatoKol To lastCol
        Dim celFarge As Long
        celFarge = Me.Cells(rad, c).Interior.Color

        ' Sjekk om denne cellen er del av aktiviteten
        If FargeAvstand(celFarge, aktivFarge) <= 10 Then
            ' Dette er del av aktiviteten
            If Not inSegment Then
                ' Start nytt segment
                inSegment = True
                segmentStart = c
                segmentNr = segmentNr + 1

                If DEBUG_MODE Then
                    Debug.Print "  → Segment " & segmentNr & " starter i kolonne " & segmentStart
                End If
            End If
            segmentEnd = c
        Else
            ' Dette er IKKE del av aktiviteten (hvit eller annen farge)
            If inSegment Then
                ' Avslutt segment
                Call FormaterSegment(rad, segmentStart, segmentEnd, aktivFarge, aktivKode, _
                                     aktivBeskr, segmentNr)
                inSegment = False
            End If
        End If
    Next c

    ' Håndter siste segment hvis vi fortsatt er i et
    If inSegment Then
        Call FormaterSegment(rad, segmentStart, segmentEnd, aktivFarge, aktivKode, _
                             aktivBeskr, segmentNr)
    End If

    ' STEG 4: Nullstill alle hvite celler (fjern alignment som "blør")
    For c = førsteDatoKol To lastCol
        If Not HarAktivitetsfarge(Me.Cells(rad, c)) Then
            With Me.Cells(rad, c)
                .Value = ""  ' Sikre ingen tekst
                .HorizontalAlignment = xlGeneral  ' Fjern xlCenterAcrossSelection
                .Font.Bold = False
            End With
        End If
    Next c
End Sub

' ============================================================================
' FORMATER SEGMENT: Formaterer et enkelt segment av en splittet aktivitet
' ============================================================================
Private Sub FormaterSegment(ByVal rad As Long, ByVal startCol As Long, ByVal endCol As Long, _
                             ByVal aktivFarge As Long, ByVal aktivKode As String, _
                             ByVal aktivBeskr As String, ByVal segmentNr As Integer)
    Dim rng As Range
    Dim førsteCelle As Range
    Dim visTekst As String

    Set rng = Me.Range(Me.Cells(rad, startCol), Me.Cells(rad, endCol))
    Set førsteCelle = Me.Cells(rad, startCol)

    If DEBUG_MODE Then
        Debug.Print "  → Formaterer segment " & segmentNr & ": " & _
                    Me.Cells(1, startCol).Address & "-" & Me.Cells(1, endCol).Address
    End If

    ' Lag visningstekst
    visTekst = aktivKode
    If Len(aktivBeskr) > 0 Then
        visTekst = visTekst & " – " & aktivBeskr
    End If

    ' VIKTIG: Sjekk om første celle i segment mangler tekst
    If Len(Trim$(førsteCelle.Value)) = 0 Then
        førsteCelle.Value = visTekst
        If DEBUG_MODE Then Debug.Print "    → Tekst restaurert i segment " & segmentNr
    End If

    ' Sett alignment KUN over dette segmentet
    With rng
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Sett font-formatering
    With førsteCelle.Font
        .Bold = True
        If ErLysFarge(aktivFarge) Then
            .Color = RGB(0, 0, 0)  ' Sort på lys bakgrunn
        Else
            .Color = RGB(255, 255, 255)  ' Hvit på mørk bakgrunn
        End If
    End With
End Sub

' ============================================================================
' SJEKK: Har cellen aktivitetsfarge? (ikke hvit/grå)
' ============================================================================
Private Function HarAktivitetsfarge(ByVal cel As Range) As Boolean
    Dim col As Long

    ' Sjekk om cellen har farge
    If cel.Interior.ColorIndex = xlColorIndexNone Then Exit Function

    col = cel.Interior.Color

    ' Sjekk om fargen er hvit eller lys grå (standard bakgrunn)
    If col = RGB(255, 255, 255) Then Exit Function  ' Hvit
    If col = RGB(242, 242, 242) Then Exit Function  ' Lys grå (formelkolonner)
    If col = RGB(250, 250, 250) Then Exit Function  ' Alternerende rad-farge

    ' Hvis vi kom hit, har cellen en aktivitetsfarge
    HarAktivitetsfarge = True
End Function

' ============================================================================
' SJEKK: Er cellen første celle i blokken?
' ============================================================================
Private Function ErFørsteCelleIBlokk(ByVal cel As Range, ByVal førsteDatoKol As Long) As Boolean
    Dim celFarge As Long
    Dim venstreCelle As Range
    Dim venstreFarge As Long

    celFarge = cel.Interior.Color

    ' Hvis vi er i første datokolonne → automatisk første celle
    If cel.Column <= førsteDatoKol Then
        ErFørsteCelleIBlokk = True
        Exit Function
    End If

    ' Sjekk cellen til venstre
    Set venstreCelle = cel.Offset(0, -1)

    ' Sjekk om venstre celle har ANNEN farge
    If venstreCelle.Interior.ColorIndex = xlColorIndexNone Then
        ' Venstre er hvit → vi er første celle
        ErFørsteCelleIBlokk = True
    ElseIf venstreCelle.Interior.Color = RGB(255, 255, 255) Then
        ' Venstre er hvit → vi er første celle
        ErFørsteCelleIBlokk = True
    ElseIf venstreCelle.Interior.Color <> celFarge Then
        ' Venstre har annen farge → vi er første celle
        ErFørsteCelleIBlokk = True
    Else
        ' Venstre har samme farge → vi er IKKE første celle
        ErFørsteCelleIBlokk = False
    End If
End Function

' ============================================================================
' OPPSLAG: Finn aktivitetskode fra farge i AKTIVITETSTYPER - OVERSIKT
' ============================================================================
Private Function SlåOppAktivitetFraFarge(ByVal søkFarge As Long, _
                                          ByRef beskrivelse As String) As String
    Dim wsTyp As Worksheet
    Dim lastRow As Long, r As Long
    Dim cellFarge As Long

    On Error Resume Next
    Set wsTyp = ThisWorkbook.Worksheets("AKTIVITETSTYPER - OVERSIKT")
    On Error GoTo 0

    If wsTyp Is Nothing Then Exit Function

    lastRow = wsTyp.Cells(wsTyp.Rows.Count, 1).End(xlUp).Row

    ' Søk gjennom aktivitetstypene
    For r = 2 To lastRow
        cellFarge = wsTyp.Cells(r, 1).Interior.Color

        ' Match på farge (med toleranse for små fargeforskjeller)
        If FargeAvstand(cellFarge, søkFarge) <= 10 Then
            SlåOppAktivitetFraFarge = Trim$(wsTyp.Cells(r, 1).Value)
            beskrivelse = Trim$(wsTyp.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

' ============================================================================
' HJELPER: Beregn fargeAvstand (RGB-forskjell)
' ============================================================================
Private Function FargeAvstand(c1 As Long, c2 As Long) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long

    ' Ekstraher RGB-komponenter
    r1 = c1 Mod 256
    g1 = (c1 \ 256) Mod 256
    b1 = (c1 \ 65536) Mod 256

    r2 = c2 Mod 256
    g2 = (c2 \ 256) Mod 256
    b2 = (c2 \ 65536) Mod 256

    ' Beregn max-avstand (Chebyshev-distanse)
    FargeAvstand = Application.WorksheetFunction.Max( _
        Abs(r1 - r2), Abs(g1 - g2), Abs(b1 - b2))
End Function

' ============================================================================
' RESTAURERING: Sett aktivitetsnavn tilbake i cellen
' ============================================================================
Private Sub RestaurerAktivitetsNavn(ByVal cel As Range, ByVal kode As String, _
                                     ByVal beskrivelse As String)
    Dim visTekst As String
    Dim farge As Long

    ' Lag visningstekst
    visTekst = kode
    If Len(beskrivelse) > 0 Then
        visTekst = visTekst & " – " & beskrivelse
    End If

    farge = cel.Interior.Color

    ' Sett tekst
    cel.Value = visTekst

    ' Formater tekst
    With cel
        .Font.Bold = True
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .WrapText = True

        ' Velg tekstfarge basert på bakgrunnsfarge (lys/mørk)
        If ErLysFarge(farge) Then
            .Font.Color = RGB(0, 0, 0)  ' Sort tekst på lys bakgrunn
        Else
            .Font.Color = RGB(255, 255, 255)  ' Hvit tekst på mørk bakgrunn
        End If
    End With
End Sub

' ============================================================================
' HJELPER: Sjekk om farge er lys (for å velge tekstfarge)
' ============================================================================
Private Function ErLysFarge(col As Long) As Boolean
    Dim r As Long, g As Long, b As Long

    r = col Mod 256
    g = (col \ 256) Mod 256
    b = (col \ 65536) Mod 256

    ' Luma-beregning (ITU BT.601)
    ErLysFarge = (0.299 * r + 0.587 * g + 0.114 * b) > 160
End Function

' ============================================================================
' UTILITY: Reset event handler (for debugging)
' ============================================================================
Public Sub ResetEventHandler()
    behandlerEndring = False
    Application.EnableEvents = True
    Debug.Print "Event handler reset - Ark1 (Planlegger med aktivitetsbeskyttelse)"
End Sub
