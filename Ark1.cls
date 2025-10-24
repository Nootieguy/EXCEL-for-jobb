' ============================================================================
' ARK1: PLANLEGGER - WORKSHEET EVENT HANDLER
' ============================================================================
' FORMÅL:
'   - Markerer dagens dato når arket aktiveres
'   - Beskytter aktivitetsnavn mot utilsiktet sletting
'   - Gjør Delete-tasten til "Clear All" (fjerner både tekst og farge)
'   - Håndterer splits når du sletter midt i aktivitet
'
' FUNKSJONALITET:
'   Auto-restaurering av aktivitetsnavn:
'   - Hvis bruker sletter FØRSTE celle i aktivitet → auto-restaurer navnet
'   - Hvis bruker sletter MIDT i aktivitet → fjerner farge (lager split)
'   - Fungerer med Delete, "Fjern markert område", Cut, og markerte områder
'
'   Delete-oppførsel (forbedret):
'   - Delete fjerner automatisk BÅDE tekst OG farge
'   - Oppfører seg som "Fjern markert område" (Clear All)
'   - Lager automatisk splits når du sletter midt i aktivitet
'
'   Split-håndtering (forbedret):
'   - Splittede aktiviteter viser aktivitetsnavn i ALLE segmenter
'   - Hvert segment får riktig formatering (alignment, font, farge)
'   - Tekst "blør" ikke over hvite celler
'   - Fungerer med både Delete og "Fjern markert område"
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
    Dim splittRader As Object
    Set splittRader = CreateObject("Scripting.Dictionary")

    For Each c In relevantEndring.Cells
        ' Sjekk om cellen er tom
        If Len(Trim$(c.Value)) = 0 Then
            If HarAktivitetsfarge(c) Then
                ' TILFELLE 1: Tom celle med aktivitetsfarge (Delete ble brukt)
                ' Sjekk om dette er første celle i aktivitetsblokken
                If ErFørsteCelleIBlokk(c, førsteDatoKol) Then
                    ' FØRSTE CELLE: Restaurer navnet (beholder fargen)
                    Call BehandleSlettetCelle(c, førsteDatoKol, datoRad)
                Else
                    ' IKKE første celle: Fjern fargen (Delete = Clear All)
                    If DEBUG_MODE Then
                        Debug.Print "Fjerner farge fra celle " & c.Address & " (Delete = Clear All)"
                    End If

                    c.Interior.ColorIndex = xlColorIndexNone
                    c.Font.Bold = False

                    ' Marker raden for split-håndtering
                    If Not splittRader.Exists(c.Row) Then
                        splittRader.Add c.Row, c.Row
                    End If
                End If
            Else
                ' TILFELLE 2: Tom hvit celle (Fjern markert område ble brukt)
                ' Sjekk om den har aktivitetsfarge til venstre eller høyre (split)
                Dim harAktivVenstre As Boolean, harAktivHøyre As Boolean
                harAktivVenstre = False
                harAktivHøyre = False

                If c.Column > førsteDatoKol Then
                    If HarAktivitetsfarge(c.Offset(0, -1)) Then harAktivVenstre = True
                End If
                If c.Column < Me.Columns.Count Then
                    If HarAktivitetsfarge(c.Offset(0, 1)) Then harAktivHøyre = True
                End If

                ' Hvis det er en split, marker raden for håndtering
                If (harAktivVenstre Or harAktivHøyre) And Not splittRader.Exists(c.Row) Then
                    splittRader.Add c.Row, c.Row
                End If
            End If
        End If
    Next c

    ' STEG 6B: Håndter splits (når hvite celler lages midt i aktivitet)
    ' Behandle hver rad kun ÉN gang
    Dim radNr As Variant
    For Each radNr In splittRader.Keys
        Call HåndterAktivitetSplitForRad(CLng(radNr), førsteDatoKol, datoRad)
    Next radNr

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
' SPLIT-HÅNDTERING: Re-formaterer aktivitet som er splittet (HELE RADEN)
' Kan kalles både internt (fra Worksheet_Change) og eksternt (fra FjernMarkert)
' ============================================================================
Public Sub HåndterAktivitetSplitForRad(ByVal rad As Long, ByVal førsteDatoKol As Long, _
                                         ByVal datoRad As Long)
    Dim aktivFarge As Long
    Dim aktivKode As String, aktivBeskr As String
    Dim c As Long
    Dim lastCol As Long

    ' STEG 1: Finn aktivitetsfargen (skann raden for første aktivitetsfarge)
    lastCol = Me.Cells(datoRad, Me.Columns.Count).End(xlToLeft).Column

    For c = førsteDatoKol To lastCol
        If HarAktivitetsfarge(Me.Cells(rad, c)) Then
            aktivFarge = Me.Cells(rad, c).Interior.Color
            Exit For
        End If
    Next c

    If aktivFarge = 0 Then Exit Sub  ' Ingen aktivitetsfarge funnet i raden

    ' STEG 2: Slå opp aktivitetskode
    aktivKode = SlåOppAktivitetFraFarge(aktivFarge, aktivBeskr)
    If Len(aktivKode) = 0 Then Exit Sub

    If DEBUG_MODE Then
        Debug.Print String(80, "=")
        Debug.Print "SPLIT HÅNDTERING FOR RAD " & rad & " [" & Now & "]"
        Debug.Print "  → Aktivitet: " & aktivKode & " - " & aktivBeskr
        Debug.Print "  → Aktivfarge: " & aktivFarge
    End If

    ' STEG 3A: FØRST - Nullstill alignment på ALLE celler i raden
    ' Dette fjerner den gamle xlCenterAcrossSelection som "blør" over
    If DEBUG_MODE Then Debug.Print "  → Nullstiller alignment på hele raden først..."

    For c = førsteDatoKol To lastCol
        With Me.Cells(rad, c)
            .HorizontalAlignment = xlGeneral  ' Fjern all alignment først
        End With
    Next c

    ' STEG 4: Skann hele raden og identifiser alle segmenter
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
                ' Avslutt segment og formater det
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

    ' STEG 5: Nullstill hvite celler fullstendig (sikre ingen tekst eller formatering)
    For c = førsteDatoKol To lastCol
        If Not HarAktivitetsfarge(Me.Cells(rad, c)) Then
            With Me.Cells(rad, c)
                .Value = ""  ' Sikre ingen tekst
                .Font.Bold = False
            End With
        End If
    Next c

    If DEBUG_MODE Then
        Debug.Print "  ✓ Rad " & rad & " håndtert!"
        Debug.Print String(80, "=")
    End If
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

    ' VIKTIG: Sett tekst i ALLE celler i segmentet først (for å tømme gamle verdier)
    Dim col As Long
    For col = startCol To endCol
        If col = startCol Then
            ' FØRSTE celle: Sett aktivitetsnavn
            Me.Cells(rad, col).Value = visTekst
        Else
            ' ANDRE celler: Tøm (kun første celle skal ha tekst)
            Me.Cells(rad, col).Value = ""
        End If
    Next col

    If DEBUG_MODE Then Debug.Print "    → Aktivitetsnavn satt i segment " & segmentNr

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
' Kan kalles både internt og eksternt (fra FjernMarkert)
' ============================================================================
Public Function HarAktivitetsfarge(ByVal cel As Range) As Boolean
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
