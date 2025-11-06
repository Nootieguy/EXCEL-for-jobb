
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
'   - Hvis bruker sletter FØRSTE celle i aktivitet ? auto-restaurer navnet
'   - Hvis bruker sletter MIDT i aktivitet ? fjerner farge (lager split)
'   - Fungerer med Delete, "Fjern markert område", Cut, og markerte områder
'
'   Delete-oppførsel (forbedret):
'   - Delete fjerner automatisk BÅDE tekst OG farge
'   - Oppfører seg som "Fjern markert område" (Clear All)
'   - Lager automatisk splits når du sletter midt i aktivitet
'
'   Split-håndtering (forbedret):
'   - Splittede aktiviteter viser aktivitetsnavn i ALLE segmenter
'   - BEHOLDER kommentarer: Hvis aktivitet har custom beskrivelse (kommentar),
'     vises denne i alle segmenter i stedet for standard beskrivelse
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

    ' STEG 5: Lagre undo-snapshot før endringer
    On Error Resume Next
    LagUndoSnapshot relevantEndring
    On Error GoTo Slutt

    ' STEG 6: Sett beskyttelsesflagg og disable events
    behandlerEndring = True
    Application.EnableEvents = False

    ' STEG 6: Behandle hver endret celle
    Dim c As Range
    Dim splittRader As Object
    Set splittRader = CreateObject("Scripting.Dictionary")

    ' Dictionary for å lagre celler som skal ha farge fjernet
    ' Key = rad, Value = Collection av kolonnenummer
    Dim cellerÅFjerne As Object
    Set cellerÅFjerne = CreateObject("Scripting.Dictionary")

    ' Dictionary for å tracke celler som allerede er behandlet (paste multi-celle fix)
    Dim behandledeCeller As Object
    Set behandledeCeller = CreateObject("Scripting.Dictionary")

    For Each c In relevantEndring.Cells
        ' Hopp over celler som allerede er behandlet (paste multi-celle)
        If behandledeCeller.Exists(c.Address) Then
            GoTo NesteCelle
        End If

        ' Sjekk om cellen er tom
        If Len(Trim$(c.Value)) = 0 Then
            If HarAktivitetsfarge(c) Then
                ' TILFELLE 1: Tom celle med aktivitetsfarge (Delete ble brukt)
                ' Sjekk om dette er en standalone celle (ingen naboer med samme farge)
                Dim erStandalone As Boolean
                erStandalone = ErStandaloneCelle(c, førsteDatoKol, cellerÅFjerne)

                If erStandalone Then
                    ' STANDALONE CELLE: Fjern helt (ingen aktivitet rundt)
                    If DEBUG_MODE Then
                        Debug.Print "Standalone celle - fjerner: " & c.Address
                    End If

                    ' Lagre denne cellen for senere fjerning
                    Dim radKey As String
                    radKey = CStr(c.Row)
                    If Not cellerÅFjerne.Exists(radKey) Then
                        Dim nyCol As Object
                        Set nyCol = CreateObject("Scripting.Dictionary")
                        cellerÅFjerne.Add radKey, nyCol
                    End If
                    cellerÅFjerne(radKey).Add c.Column, c.Column

                    ' Ikke marker for split - det er ingen aktivitet å splitte
                ElseIf ErFørsteCelleIBlokk(c, førsteDatoKol) Then
                    ' FØRSTE CELLE: Bruk Undo for å sjekke om celle hadde tekst
                    ' Hvis den hadde tekst, flytt til neste celle
                    ' Hvis ikke, slett hele aktiviteten

                    ' Bruk Application.Undo for å hente original tilstand midlertidig
                    Application.Undo
                    Dim hadTekst As Boolean
                    Dim originalTekst As String
                    hadTekst = (Len(Trim$(c.Value)) > 0)
                    originalTekst = c.Value

                    ' Slett på nytt
                    c.ClearContents

                    If hadTekst Then
                        ' Celle 1 hadde tekst - sjekk om andre celler også slettes
                        Dim aktivFarge As Long
                        aktivFarge = c.Interior.Color

                        ' Sjekk om andre celler i aktiviteten også er i relevantEndring
                        Dim andreOgsåSlettet As Boolean
                        andreOgsåSlettet = False
                        Dim scanCol As Long
                        For scanCol = c.Column + 1 To lastCol
                            Dim scanCel As Range
                            Set scanCel = Me.Cells(c.Row, scanCol)

                            ' Stopp hvis vi treffer en hvit celle (gap)
                            If Not HarAktivitetsfarge(scanCel) Then
                                Exit For
                            End If

                            ' Stopp hvis vi treffer en celle med bold tekst (ny aktivitet)
                            If scanCel.Font.Bold And Len(Trim$(scanCel.Value)) > 0 Then
                                Exit For
                            End If

                            If FargeAvstand(scanCel.Interior.Color, aktivFarge) <= 10 Then
                                ' Dette er del av aktiviteten - sjekk om den er i relevantEndring
                                Dim celleISel As Range
                                Set celleISel = Nothing
                                On Error Resume Next
                                Set celleISel = Intersect(scanCel, relevantEndring)
                                On Error GoTo 0
                                If Not celleISel Is Nothing Then
                                    andreOgsåSlettet = True
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Next scanCol

                        If andreOgsåSlettet Then
                            ' Hele aktiviteten slettes - IKKE flytt tekst
                            If DEBUG_MODE Then
                                Debug.Print "Hele aktiviteten slettes (flere celler valgt): " & c.Address
                            End If

                            radKey = CStr(c.Row)
                            If Not cellerÅFjerne.Exists(radKey) Then
                                Dim nyCol3c As Object
                                Set nyCol3c = CreateObject("Scripting.Dictionary")
                                cellerÅFjerne.Add radKey, nyCol3c
                            End If
                            cellerÅFjerne(radKey).Add c.Column, c.Column
                        Else
                            ' Bare første celle slettes - flytt tekst til neste celle
                            If DEBUG_MODE Then
                                Debug.Print "Første celle slettes - flytter tekst til neste: " & c.Address
                            End If

                            ' Finn neste celle i samme aktivitet
                            Dim nesteCelleIAktivitet As Range
                            Set nesteCelleIAktivitet = Nothing

                            ' Skann til høyre for å finne neste celle med samme farge
                            For scanCol = c.Column + 1 To lastCol
                                Dim nextScanCell As Range
                                Set nextScanCell = Me.Cells(c.Row, scanCol)

                                If HarAktivitetsfarge(nextScanCell) Then
                                    ' Stopp hvis vi treffer en celle med bold tekst (ny aktivitet)
                                    If nextScanCell.Font.Bold And Len(Trim$(nextScanCell.Value)) > 0 Then
                                        Exit For
                                    End If

                                    ' Fortsett hvis samme farge
                                    If FargeAvstand(nextScanCell.Interior.Color, aktivFarge) <= 10 Then
                                        Set nesteCelleIAktivitet = nextScanCell
                                        Exit For
                                    End If
                                Else
                                    ' Hvit celle = gap mellom aktiviteter, fortsett å lete
                                    ' (men ikke behandle den hvite cellen som del av aktiviteten)
                                End If
                            Next scanCol

                            ' Hvis det finnes en neste celle, kopier teksten dit
                            If Not nesteCelleIAktivitet Is Nothing Then
                                nesteCelleIAktivitet.Value = originalTekst
                            End If

                            ' Marker denne cellen for fjerning (cleanup vil fjerne fargen osv)
                            radKey = CStr(c.Row)
                            If Not cellerÅFjerne.Exists(radKey) Then
                                Dim nyCol3b As Object
                                Set nyCol3b = CreateObject("Scripting.Dictionary")
                                cellerÅFjerne.Add radKey, nyCol3b
                            End If
                            cellerÅFjerne(radKey).Add c.Column, c.Column
                        End If
                    Else
                        ' Celle hadde IKKE tekst ? hele aktiviteten slettes
                        If DEBUG_MODE Then
                            Debug.Print "Hele aktiviteten slettes (celle 1 var tom): " & c.Address
                        End If

                        radKey = CStr(c.Row)
                        If Not cellerÅFjerne.Exists(radKey) Then
                            Dim nyCol3 As Object
                            Set nyCol3 = CreateObject("Scripting.Dictionary")
                            cellerÅFjerne.Add radKey, nyCol3
                        End If
                        cellerÅFjerne(radKey).Add c.Column, c.Column
                    End If
                Else
                    ' IKKE første celle: Marker for fjerning (Delete = Clear All)
                    If DEBUG_MODE Then
                        Debug.Print "Markerer celle for fjerning: " & c.Address
                    End If

                    ' Lagre denne cellen for senere fjerning
                    radKey = CStr(c.Row)
                    If Not cellerÅFjerne.Exists(radKey) Then
                        Dim nyCol2 As Object
                        Set nyCol2 = CreateObject("Scripting.Dictionary")
                        cellerÅFjerne.Add radKey, nyCol2
                    End If
                    cellerÅFjerne(radKey).Add c.Column, c.Column

                    ' Marker raden for split-håndtering
                    If Not splittRader.Exists(c.Row) Then
                        splittRader.Add c.Row, c.Row
                    End If
                End If
            Else
                ' TILFELLE 2: Tom hvit celle (Fjern markert område / Cut ble brukt)
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

                ' Restaurer rutenett for tom hvit celle (f.eks. etter cut)
                Call RestaurerRutenettForCelle(c)
            End If
        Else
            ' TILFELLE 3: Celle har fått innhold (paste/drag/typing)
            ' Sjekk om dette er en aktivitet som trenger multi-celle formatering
            If HarAktivitetsfarge(c) And c.Font.Bold And Len(Trim$(c.Value)) > 0 Then
                ' Dette kan være en paste/drag operasjon av en aktivitet
                ' Finn ut hvor mange celler til høyre som skal inkluderes
                ' Basert på naboer med samme farge
                Dim sluttCol As Long
                sluttCol = c.Column

                ' Skann til høyre for å finne celler med samme farge
                Dim col As Long
                For col = c.Column + 1 To lastCol
                    Dim scanCell As Range
                    Set scanCell = Me.Cells(c.Row, col)

                    ' Stopp hvis vi treffer en hvit/tom celle (gap mellom aktiviteter)
                    If Not HarAktivitetsfarge(scanCell) Then
                        Exit For
                    End If

                    ' Stopp hvis vi treffer en celle med bold tekst (ny aktivitet starter)
                    If scanCell.Font.Bold And Len(Trim$(scanCell.Value)) > 0 Then
                        Exit For
                    End If

                    ' Fortsett hvis samme farge
                    If FargeAvstand(scanCell.Interior.Color, c.Interior.Color) <= 10 Then
                        sluttCol = col
                    Else
                        Exit For
                    End If
                Next col

                ' OVERLAPP-SJEKK: Sjekk naboer for å detektere overlapp
                Dim harOverlapp As Boolean
                Dim dragFarge As Long
                Dim dragTekst As String
                Dim markerCol As Long
                Dim ryddCel As Range
                Dim høyreNabo As Range
                Dim venstreNabo As Range
                Dim dragRng As Range
                Dim dragCel As Range
                Dim gjenCel As Range
                Dim gjenRng As Range
                harOverlapp = False
                dragFarge = c.Interior.Color
                dragTekst = c.Value

                ' MERK: dragData er ikke lenger brukt, bruker dragTekst og dragFarge istedenfor

                ' Sjekk HØYRE nabo: Er det en annen aktivitet der?
                ' Dette skjer når brukeren drar en aktivitet inn i første celle av en annen aktivitet
                If sluttCol < lastCol Then
                    Set høyreNabo = Me.Cells(c.Row, sluttCol + 1)

                    If HarAktivitetsfarge(høyreNabo) Then
                        ' Høyre nabo har aktivitetsfarge - sjekk om det er en ANNEN aktivitet
                        If FargeAvstand(høyreNabo.Interior.Color, dragFarge) > 10 Then
                            harOverlapp = True

                            If DEBUG_MODE Then
                                Debug.Print "OVERLAPP: Høyre nabo har annen farge " & høyreNabo.Interior.Color & " vs drag-farge " & dragFarge
                            End If
                        End If
                    End If
                End If

                ' Sjekk VENSTRE nabo også (hvis brukeren drar til slutten av en aktivitet)
                If Not harOverlapp And c.Column > førsteDatoKol Then
                    Set venstreNabo = Me.Cells(c.Row, c.Column - 1)

                    If HarAktivitetsfarge(venstreNabo) Then
                        If FargeAvstand(venstreNabo.Interior.Color, dragFarge) > 10 Then
                            harOverlapp = True

                            If DEBUG_MODE Then
                                Debug.Print "OVERLAPP: Venstre nabo har annen farge " & venstreNabo.Interior.Color & " vs drag-farge " & dragFarge
                            End If
                        End If
                    End If
                End If

                ' Hvis det er overlapp, fiks navn på ALLE aktiviteter på raden
                If harOverlapp Then
                    If DEBUG_MODE Then
                        Debug.Print "OVERLAPP DETEKTERT: Fikser navn på alle aktiviteter på raden"
                    End If

                    Dim debugMsg As String
                    debugMsg = "Overlapp detektert! Skanner rad " & c.Row & vbCrLf & vbCrLf

                    ' Skann hele raden fra førsteDatoKol til lastCol
                    ' For hver aktivitet: sjekk om fargen matcher navnet, hvis ikke - fiks det
                    Dim overlapScanCol As Long
                    overlapScanCol = førsteDatoKol

                    Do While overlapScanCol <= lastCol
                        Dim overlapScanCel As Range
                        Set overlapScanCel = Me.Cells(c.Row, overlapScanCol)

                        If HarAktivitetsfarge(overlapScanCel) Then
                            Dim erStartAvBlokk As Boolean
                            erStartAvBlokk = True

                            If overlapScanCol > førsteDatoKol Then
                                Dim celTilVenstre As Range
                                Set celTilVenstre = Me.Cells(c.Row, overlapScanCol - 1)

                                If HarAktivitetsfarge(celTilVenstre) Then
                                    If FargeAvstand(celTilVenstre.Interior.Color, overlapScanCel.Interior.Color) <= 10 Then
                                        erStartAvBlokk = False
                                    End If
                                End If
                            End If

                            If erStartAvBlokk Then
                                Dim aktFarge As Long
                                Dim aktTekst As String
                                Dim aktStartCol As Long
                                Dim aktSluttCol As Long
                                Dim eksisterendeTekst As String
                                Dim harKommentar As Boolean
                                Dim kommentarTekst As String
                                Dim kommentarErThreaded As Boolean

                                aktFarge = overlapScanCel.Interior.Color
                                aktTekst = Trim$(overlapScanCel.Value)
                                aktStartCol = overlapScanCol
                                aktSluttCol = aktStartCol
                                eksisterendeTekst = aktTekst

                                harKommentar = HentKommentarFraCelle(overlapScanCel, kommentarTekst, kommentarErThreaded)

                                debugMsg = debugMsg & "Kol " & overlapScanCol & ": Farge=" & aktFarge & ", Tekst='" & aktTekst & "'" & vbCrLf

                                ' Finn slutten av denne aktiviteten og let etter eksisterende tekst/kommentar
                                For col = aktStartCol + 1 To lastCol
                                    Dim aktScanCell As Range
                                    Set aktScanCell = Me.Cells(c.Row, col)

                                    ' Stopp hvis vi treffer en hvit/tom celle (gap mellom aktiviteter)
                                    If Not HarAktivitetsfarge(aktScanCell) Then
                                        Exit For
                                    End If

                                    ' Stopp hvis vi treffer en celle med bold tekst (ny aktivitet starter)
                                    If aktScanCell.Font.Bold And Len(Trim$(aktScanCell.Value)) > 0 Then
                                        Exit For
                                    End If

                                    ' Fortsett hvis samme farge
                                    If FargeAvstand(aktScanCell.Interior.Color, aktFarge) <= 10 Then
                                        aktSluttCol = col

                                        If Len(eksisterendeTekst) = 0 Then
                                            Dim kandidatTekst As String
                                            kandidatTekst = Trim$(aktScanCell.Value)
                                            If Len(kandidatTekst) > 0 Then
                                                eksisterendeTekst = kandidatTekst
                                            End If
                                        End If

                                        If Not harKommentar Then
                                            harKommentar = HentKommentarFraCelle(aktScanCell, kommentarTekst, kommentarErThreaded)
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next col

                                debugMsg = debugMsg & "  Spenn: " & aktStartCol & "-" & aktSluttCol & vbCrLf

                                Dim riktigBeskrivelse As String, riktigKode As String
                                Dim riktigTekst As String

                                ' Slå opp aktivitetskode fra farge først
                                riktigKode = SlåOppAktivitetFraFarge(aktFarge, riktigBeskrivelse)

                                If Len(eksisterendeTekst) > 0 Then
                                    ' Det finnes eksisterende tekst - sjekk om det er gyldig
                                    If Len(riktigKode) > 0 Then
                                        ' Sjekk om eksisterende tekst starter med riktig kode
                                        If StrComp(Left$(eksisterendeTekst, Len(riktigKode)), riktigKode, vbTextCompare) = 0 Then
                                            ' Teksten starter med riktig kode - behold den (kan være custom beskrivelse)
                                            riktigTekst = eksisterendeTekst
                                        Else
                                            ' Teksten matcher ikke koden - erstatt med standard
                                            If Len(riktigBeskrivelse) > 0 Then
                                                riktigTekst = riktigKode & " - " & riktigBeskrivelse
                                            Else
                                                riktigTekst = riktigKode
                                            End If
                                        End If
                                    Else
                                        ' Fant ikke kode for farge - behold eksisterende tekst
                                        riktigTekst = eksisterendeTekst
                                    End If
                                Else
                                    ' Ingen eksisterende tekst - bruk standard
                                    If Len(riktigKode) > 0 Then
                                        If Len(riktigBeskrivelse) > 0 Then
                                            riktigTekst = riktigKode & " - " & riktigBeskrivelse
                                        Else
                                            riktigTekst = riktigKode
                                        End If
                                    Else
                                        riktigTekst = ""
                                    End If
                                End If

                                If Len(riktigTekst) > 0 Then
                                    debugMsg = debugMsg & "  Riktig navn: '" & riktigTekst & "'" & vbCrLf
                                Else
                                    debugMsg = debugMsg & "  Fant ikke navn for farge " & aktFarge & vbCrLf
                                End If

                                ' ALLTID fiks formateringen for å isolere denne aktiviteten
                                ' Selv om teksten er riktig, må vi sikre at xlCenterAcrossSelection
                                ' ikke spreder seg utenfor denne aktivitetens område

                                ' Sett riktig tekst i første celle
                                overlapScanCel.Value = riktigTekst
                                overlapScanCel.HorizontalAlignment = xlCenterAcrossSelection
                                FjernKommentarer overlapScanCel
                                If harKommentar Then Call LeggTilKommentar(overlapScanCel, kommentarTekst, kommentarErThreaded)

                                ' Tøm teksten i resten av cellene i denne aktiviteten
                                For col = aktStartCol + 1 To aktSluttCol
                                    Set gjenCel = Me.Cells(c.Row, col)
                                    With gjenCel
                                        .Value = ""
                                        .HorizontalAlignment = xlCenterAcrossSelection
                                    End With
                                    FjernKommentarer gjenCel
                                Next col

                                If aktTekst <> riktigTekst Then
                                    debugMsg = debugMsg & "  FIKSET: '" & aktTekst & "' -> '" & riktigTekst & "'" & vbCrLf
                                Else
                                    debugMsg = debugMsg & "  Formatering fikset (tekst var riktig)" & vbCrLf
                                End If

                                debugMsg = debugMsg & vbCrLf

                                ' Hopp til slutten av denne aktiviteten
                                overlapScanCol = aktSluttCol + 1
                            Else
                                overlapScanCol = overlapScanCol + 1
                            End If
                        Else
                            overlapScanCol = overlapScanCol + 1
                        End If
                    Loop

                    If DEBUG_MODE Then MsgBox debugMsg, vbInformation
                End If

                ' Hvis vi fant flere celler med samme farge, fiks formateringen
                If sluttCol > c.Column Then
                    ' Fiks formateringen for hele blokken
                    Call FiksMultiCelleAktivitet(c, sluttCol, c.Interior.Color)

                    ' Marker alle celler i blokken som behandlet
                    For markerCol = c.Column To sluttCol
                        behandledeCeller.Add Me.Cells(c.Row, markerCol).Address, True
                    Next markerCol
                End If
            End If
        End If
NesteCelle:
    Next c

    ' STEG 6B: Håndter splits (når hvite celler lages midt i aktivitet)
    ' Behandle hver rad kun ÉN gang
    Dim radNr As Variant
    For Each radNr In splittRader.Keys
        Call HåndterAlleAktiviteterMedSplitIRad(CLng(radNr), førsteDatoKol, datoRad, cellerÅFjerne)
    Next radNr

    ' STEG 6C: Rydd opp celler som er markert for sletting men ikke ble håndtert av split-logikk
    ' (f.eks. standalone celler som ikke genererer splits)
    Dim radKeyCleanup As Variant
    For Each radKeyCleanup In cellerÅFjerne.Keys
        Dim colDictCleanup As Object
        Set colDictCleanup = cellerÅFjerne(radKeyCleanup)

        Dim colCleanup As Variant
        For Each colCleanup In colDictCleanup.Keys
            Dim celCleanup As Range
            Set celCleanup = Me.Cells(CLng(radKeyCleanup), CLng(colCleanup))

            ' Sjekk om cellen fortsatt har farge (ikke håndtert av split-logikk)
            If celCleanup.Interior.ColorIndex <> xlColorIndexNone Then
                If DEBUG_MODE Then
                    Debug.Print "Rydder opp celle som ikke ble håndtert: " & celCleanup.Address
                End If

                ' Rydd helt som FjernMarkert gjør
                celCleanup.ClearComments
                celCleanup.ClearContents
                celCleanup.Font.Bold = False
                celCleanup.Font.ColorIndex = xlColorIndexAutomatic
                celCleanup.HorizontalAlignment = xlGeneral
                celCleanup.VerticalAlignment = xlCenter
                celCleanup.WrapText = False

                ' Sett bakgrunn til ren hvit
                With celCleanup.Interior
                    .Pattern = xlSolid
                    .TintAndShade = 0
                    .Color = RGB(255, 255, 255)
                    .PatternTintAndShade = 0
                End With

                ' Slå av diagonale kanter
                celCleanup.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
                celCleanup.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone

                ' Nullstill alle kanter først
                celCleanup.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                celCleanup.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                celCleanup.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
                celCleanup.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
                celCleanup.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
                celCleanup.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

                ' Sett standard rutenett (tynne kanter)
                With celCleanup.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With

                With celCleanup.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With

                With celCleanup.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With

                With celCleanup.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
            End If
        Next colCleanup
    Next radKeyCleanup

    ' STEG 6D: Fjern tomme under-rader (som FjernMarkert gjør)
    ' Gå gjennom hver rad som ble endret og slett hvis den er helt tom
    Dim personHeaderRad As Long
    personHeaderRad = Me.Range("PersonHeader").Row

    For Each radKeyCleanup In cellerÅFjerne.Keys
        Dim radNum As Long
        radNum = CLng(radKeyCleanup)

        ' Bare slett under-rader (ikke header-raden)
        If radNum > personHeaderRad Then
            ' Sjekk om raden er helt tom
            If RadErTomIAlleDatoKolonner(radNum, førsteDatoKol, lastCol) Then
                ' Sjekk også at kolonne A er tom
                If Len(Trim$(Me.Cells(radNum, 1).Value)) = 0 Then
                    If DEBUG_MODE Then
                        Debug.Print "Sletter tom under-rad: " & radNum
                    End If
                    Me.Rows(radNum).Delete
                End If
            End If
        End If
    Next radKeyCleanup

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
        If DEBUG_MODE Then Debug.Print "  ? Ikke første celle - tillater split/gap"
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
        If DEBUG_MODE Then Debug.Print "  ? Fant ikke aktivitetskode for farge " & cel.Interior.Color
        Exit Sub
    End If

    If DEBUG_MODE Then
        Debug.Print "  ? Funnet aktivitet: " & aktivKode & " - " & aktivBeskr
    End If

    ' STEG 2: Restaurer navnet i cellen
    Call RestaurerAktivitetsNavn(cel, aktivKode, aktivBeskr)

    If DEBUG_MODE Then
        Debug.Print "  ? Navn restaurert!"
        Debug.Print String(80, "=")
    End If
End Sub

' ============================================================================
' SPLIT-HÅNDTERING: Finn og håndter ALLE aktivitetsblokker med splits i raden
' Kan kalles både internt (fra Worksheet_Change) og eksternt (fra FjernMarkert)
' ============================================================================
Public Sub HåndterAlleAktiviteterMedSplitIRad(ByVal rad As Long, ByVal førsteDatoKol As Long, _
                                                ByVal datoRad As Long, ByVal cellerÅFjerne As Object)
    ' Valider parameter (kan være Nothing når kalt eksternt)
    If cellerÅFjerne Is Nothing Then
        Set cellerÅFjerne = CreateObject("Scripting.Dictionary")
    End If

    Dim lastCol As Long
    Dim c As Long
    Dim aktivitetsBlokker As Object
    Set aktivitetsBlokker = CreateObject("Scripting.Dictionary")

    lastCol = Me.Cells(datoRad, Me.Columns.Count).End(xlToLeft).Column

    ' STEG 1: Identifiser hver unike aktivitetsblokk i raden
    ' En blokk = sammenhengende celler med SAMME farge
    Dim blokNr As Long
    blokNr = 0

    Dim inBlokk As Boolean
    Dim blokkStart As Long, blokkSlutt As Long
    Dim blokkFarge As Long

    inBlokk = False

    For c = førsteDatoKol To lastCol
        If HarAktivitetsfarge(Me.Cells(rad, c)) Then
            Dim celleFarge As Long
            Dim celleBold As Boolean
            Dim celleHarTekst As Boolean
            celleFarge = Me.Cells(rad, c).Interior.Color
            celleBold = Me.Cells(rad, c).Font.Bold
            celleHarTekst = (Len(Trim$(Me.Cells(rad, c).Value)) > 0)

            If Not inBlokk Then
                ' Start ny blokk
                inBlokk = True
                blokkStart = c
                blokkFarge = celleFarge
            ElseIf FargeAvstand(celleFarge, blokkFarge) > 10 Then
                ' Farge endret - avslutt forrige blokk og start ny
                blokNr = blokNr + 1
                aktivitetsBlokker.Add blokNr, Array(blokkStart, c - 1, blokkFarge)

                blokkStart = c
                blokkFarge = celleFarge
            ElseIf celleBold And celleHarTekst Then
                ' Bold tekst funnet - dette er en ny aktivitet med samme farge
                ' Avslutt forrige blokk og start ny
                blokNr = blokNr + 1
                aktivitetsBlokker.Add blokNr, Array(blokkStart, c - 1, blokkFarge)

                blokkStart = c
                blokkFarge = celleFarge
            End If
            blokkSlutt = c
        Else
            ' Hvit celle
            If inBlokk Then
                ' Avslutt blokken
                blokNr = blokNr + 1
                aktivitetsBlokker.Add blokNr, Array(blokkStart, blokkSlutt, blokkFarge)
                inBlokk = False
            End If
        End If
    Next c

    ' Håndter siste blokk hvis vi fortsatt er i en
    If inBlokk Then
        blokNr = blokNr + 1
        aktivitetsBlokker.Add blokNr, Array(blokkStart, blokkSlutt, blokkFarge)
    End If

    ' STEG 2: For hver blokk, sjekk om den har split og håndter den
    Dim k As Variant
    For Each k In aktivitetsBlokker.Keys
        Dim blokkInfo As Variant
        blokkInfo = aktivitetsBlokker(k)

        Dim bStart As Long, bSlutt As Long, bFarge As Long
        bStart = blokkInfo(0)
        bSlutt = blokkInfo(1)
        bFarge = blokkInfo(2)

        ' Sjekk om denne blokken har hvite celler (split) ELLER celler markert for fjerning
        Dim harSplit As Boolean
        harSplit = False

        ' Først sjekk hvite celler
        For c = bStart To bSlutt
            If Not HarAktivitetsfarge(Me.Cells(rad, c)) Then
                harSplit = True
                Exit For
            End If
        Next c

        ' Deretter sjekk om noen celler er markert for fjerning
        If Not harSplit Then
            Dim radKeyCheck As String
            radKeyCheck = CStr(rad)
            If cellerÅFjerne.Exists(radKeyCheck) Then
                Dim colDict As Object
                Set colDict = cellerÅFjerne(radKeyCheck)
                For c = bStart To bSlutt
                    If colDict.Exists(c) Then
                        harSplit = True
                        Exit For
                    End If
                Next c
            End If
        End If

        ' Hvis blokken har split, håndter den
        If harSplit Then
            Call HåndterEnkeltAktivitetSplit(rad, bStart, bSlutt, bFarge, førsteDatoKol, datoRad, cellerÅFjerne)
        End If
    Next k
End Sub

' ============================================================================
' SPLIT-HÅNDTERING: Håndter split for EN spesifikk aktivitetsblokk
' ============================================================================
Private Sub HåndterEnkeltAktivitetSplit(ByVal rad As Long, ByVal blokkStart As Long, _
                                         ByVal blokkSlutt As Long, ByVal aktivFarge As Long, _
                                         ByVal førsteDatoKol As Long, ByVal datoRad As Long, _
                                         ByVal cellerÅFjerne As Object)
    Dim aktivKode As String, aktivBeskr As String
    Dim c As Long
    Dim lastCol As Long

    lastCol = Me.Cells(datoRad, Me.Columns.Count).End(xlToLeft).Column

    ' STEG 0: Fjern farger fra markerte celler FØRST
    Dim radKeyFjern As String
    radKeyFjern = CStr(rad)
    If cellerÅFjerne.Exists(radKeyFjern) Then
        Dim colDictFjern As Object
        Set colDictFjern = cellerÅFjerne(radKeyFjern)

        For c = blokkStart To blokkSlutt
            If colDictFjern.Exists(c) Then
                If DEBUG_MODE Then
                    Debug.Print "  ? Fjerner farge fra celle kolonne " & c
                End If

                Dim celFjern As Range
                Set celFjern = Me.Cells(rad, c)

                ' Rydd celle fullstendig
                celFjern.ClearComments
                celFjern.ClearContents
                celFjern.Interior.ColorIndex = xlColorIndexNone
                celFjern.Font.Bold = False
                celFjern.Font.ColorIndex = xlColorIndexAutomatic
                celFjern.HorizontalAlignment = xlGeneral
                celFjern.VerticalAlignment = xlCenter
                celFjern.WrapText = False

                ' Sett hvit bakgrunn
                With celFjern.Interior
                    .Pattern = xlSolid
                    .TintAndShade = 0
                    .Color = RGB(255, 255, 255)
                    .PatternTintAndShade = 0
                End With

                ' Nullstill alle borders først
                celFjern.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
                celFjern.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
                celFjern.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
                celFjern.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                celFjern.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
                celFjern.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
                celFjern.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
                celFjern.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

                ' Sett standard rutenett
                With celFjern.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
                With celFjern.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
                With celFjern.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
                With celFjern.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .Color = RGB(0, 0, 0)
                End With
            End If
        Next c
    End If

    ' STEG 1: Finn tekst FRA DENNE BLOKKEN
    Dim førsteCelleMedTekst As Range
    Set førsteCelleMedTekst = Nothing

    For c = blokkStart To blokkSlutt
        If HarAktivitetsfarge(Me.Cells(rad, c)) Then
            If Len(Trim$(Me.Cells(rad, c).Value)) > 0 Then
                Set førsteCelleMedTekst = Me.Cells(rad, c)
                Exit For
            End If
        End If
    Next c

    ' STEG 2: Hent aktivitetskode og beskrivelse
    Dim fullTekst As String

    If Not førsteCelleMedTekst Is Nothing Then
        ' Bruk teksten fra cellen (kan være kommentar eller standard beskrivelse)
        fullTekst = Trim$(førsteCelleMedTekst.Value)

        ' Parse ut kode og beskrivelse
        Dim dashPos As Long
        dashPos = InStr(fullTekst, "-")
        If dashPos = 0 Then dashPos = InStr(fullTekst, "-")

        If dashPos > 0 Then
            aktivKode = Trim$(Left$(fullTekst, dashPos - 1))
            aktivBeskr = Trim$(Mid$(fullTekst, dashPos + 1))
        Else
            aktivKode = fullTekst
            aktivBeskr = ""
        End If
    Else
        ' Slå opp standard fra AKTIVITETSTYPER
        aktivKode = SlåOppAktivitetFraFarge(aktivFarge, aktivBeskr)
    End If

    If Len(aktivKode) = 0 Then Exit Sub

    If DEBUG_MODE Then
        Debug.Print String(80, "=")
        Debug.Print "SPLIT FOR BLOKK " & blokkStart & "-" & blokkSlutt & " [" & Now & "]"
        Debug.Print "  ? Aktivitet: " & aktivKode & " - " & aktivBeskr
    End If

    ' STEG 3: Nullstill alignment på DENNE blokken
    For c = blokkStart To blokkSlutt
        Me.Cells(rad, c).HorizontalAlignment = xlGeneral
    Next c

    ' STEG 4: Skann blokken og identifiser segmenter
    Dim inSegment As Boolean
    Dim segmentStart As Long, segmentEnd As Long
    Dim segmentNr As Integer

    segmentNr = 0
    inSegment = False

    For c = blokkStart To blokkSlutt
        If FargeAvstand(Me.Cells(rad, c).Interior.Color, aktivFarge) <= 10 Then
            If Not inSegment Then
                inSegment = True
                segmentStart = c
                segmentNr = segmentNr + 1
            End If
            segmentEnd = c
        Else
            If inSegment Then
                Call FormaterSegment(rad, segmentStart, segmentEnd, aktivFarge, aktivKode, _
                                     aktivBeskr, segmentNr)
                inSegment = False
            End If
        End If
    Next c

    If inSegment Then
        Call FormaterSegment(rad, segmentStart, segmentEnd, aktivFarge, aktivKode, _
                             aktivBeskr, segmentNr)
    End If

    ' STEG 5: Nullstill hvite celler i blokken
    For c = blokkStart To blokkSlutt
        If Not HarAktivitetsfarge(Me.Cells(rad, c)) Then
            With Me.Cells(rad, c)
                .Value = ""
                .Font.Bold = False
            End With
        End If
    Next c

    If DEBUG_MODE Then
        Debug.Print "  ? Blokk håndtert!"
        Debug.Print String(80, "=")
    End If
End Sub

' ============================================================================
' SPLIT-HÅNDTERING: Re-formaterer aktivitet som er splittet (HELE RADEN)
' Kan kalles eksternt (fra FjernMarkert) - bruker gammel logikk
' ============================================================================
Public Sub HåndterAktivitetSplitForRad(ByVal rad As Long, ByVal førsteDatoKol As Long, _
                                         ByVal datoRad As Long)
    Dim aktivFarge As Long
    Dim aktivKode As String, aktivBeskr As String
    Dim c As Long
    Dim lastCol As Long

    ' STEG 1: Finn første celle med aktivitetsfarge OG tekst (for å få kommentar)
    lastCol = Me.Cells(datoRad, Me.Columns.Count).End(xlToLeft).Column

    Dim førsteCelleMedTekst As Range
    Set førsteCelleMedTekst = Nothing

    For c = førsteDatoKol To lastCol
        If HarAktivitetsfarge(Me.Cells(rad, c)) Then
            If aktivFarge = 0 Then aktivFarge = Me.Cells(rad, c).Interior.Color

            ' Hvis denne cellen har tekst, bruk den
            If Len(Trim$(Me.Cells(rad, c).Value)) > 0 Then
                Set førsteCelleMedTekst = Me.Cells(rad, c)
                Exit For
            End If
        End If
    Next c

    If aktivFarge = 0 Then Exit Sub  ' Ingen aktivitetsfarge funnet i raden

    ' STEG 2: Hent aktivitetskode og beskrivelse
    Dim fullTekst As String

    If Not førsteCelleMedTekst Is Nothing Then
        ' Bruk teksten fra cellen (kan være kommentar eller standard beskrivelse)
        fullTekst = Trim$(førsteCelleMedTekst.Value)

        ' Parse ut kode og beskrivelse (format: "KODE - Beskrivelse")
        Dim dashPos As Long
        dashPos = InStr(fullTekst, "-")
        If dashPos = 0 Then dashPos = InStr(fullTekst, "-")  ' Prøv også vanlig bindestrek

        If dashPos > 0 Then
            aktivKode = Trim$(Left$(fullTekst, dashPos - 1))
            aktivBeskr = Trim$(Mid$(fullTekst, dashPos + 1))
        Else
            ' Ingen separator funnet, bruk hele teksten som kode
            aktivKode = fullTekst
            aktivBeskr = ""
        End If
    Else
        ' Ingen celle med tekst funnet, slå opp standard fra AKTIVITETSTYPER
        aktivKode = SlåOppAktivitetFraFarge(aktivFarge, aktivBeskr)
    End If

    If Len(aktivKode) = 0 Then Exit Sub

    If DEBUG_MODE Then
        Debug.Print String(80, "=")
        Debug.Print "SPLIT HÅNDTERING FOR RAD " & rad & " [" & Now & "]"
        Debug.Print "  ? Aktivitet: " & aktivKode & " - " & aktivBeskr
        Debug.Print "  ? Aktivfarge: " & aktivFarge
    End If

    ' STEG 3A: FØRST - Nullstill alignment på ALLE celler i raden
    ' Dette fjerner den gamle xlCenterAcrossSelection som "blør" over
    If DEBUG_MODE Then Debug.Print "  ? Nullstiller alignment på hele raden først..."

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
                    Debug.Print "  ? Segment " & segmentNr & " starter i kolonne " & segmentStart
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
        Debug.Print "  ? Rad " & rad & " håndtert!"
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
        Debug.Print "  ? Formaterer segment " & segmentNr & ": " & _
                    Me.Cells(1, startCol).Address & "-" & Me.Cells(1, endCol).Address
    End If

    ' Lag visningstekst
    visTekst = aktivKode
    If Len(aktivBeskr) > 0 Then
        visTekst = visTekst & " - " & aktivBeskr
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

    If DEBUG_MODE Then Debug.Print "    ? Aktivitetsnavn satt i segment " & segmentNr

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

    ' Hvis vi er i første datokolonne ? automatisk første celle
    If cel.Column <= førsteDatoKol Then
        ErFørsteCelleIBlokk = True
        Exit Function
    End If

    ' Sjekk cellen til venstre
    Set venstreCelle = cel.Offset(0, -1)

    ' Sjekk om venstre celle har ANNEN farge
    If venstreCelle.Interior.ColorIndex = xlColorIndexNone Then
        ' Venstre er hvit ? vi er første celle
        ErFørsteCelleIBlokk = True
    ElseIf venstreCelle.Interior.Color = RGB(255, 255, 255) Then
        ' Venstre er hvit ? vi er første celle
        ErFørsteCelleIBlokk = True
    ElseIf venstreCelle.Interior.Color <> celFarge Then
        ' Venstre har annen farge ? vi er første celle
        ErFørsteCelleIBlokk = True
    Else
        ' Venstre har samme farge ? vi er IKKE første celle
        ErFørsteCelleIBlokk = False
    End If
End Function

' ============================================================================
' SJEKK: Blir hele aktiviteten slettet (alle celler i blokken)?
' ============================================================================
Private Function SjekkOmHeleAktivitetenSlettes(ByVal førsteCelle As Range, _
                                                ByVal førsteDatoKol As Long, _
                                                ByVal cellerÅFjerne As Object) As Boolean
    Dim aktivFarge As Long
    Dim radKey As String
    Dim c As Long

    aktivFarge = førsteCelle.Interior.Color
    radKey = CStr(førsteCelle.Row)
    SjekkOmHeleAktivitetenSlettes = True  ' Anta at hele aktiviteten slettes

    ' Skann fra førsteCelle og utover til høyre
    For c = førsteCelle.Column To Me.Columns.Count
        Dim cel As Range
        Set cel = Me.Cells(førsteCelle.Row, c)

        ' Stopp hvis vi kommer til en hvit celle (gap)
        If Not HarAktivitetsfarge(cel) Then
            Exit For
        End If

        ' Stopp hvis vi treffer en celle med bold tekst (ny aktivitet) - unntatt første celle
        If c > førsteCelle.Column Then
            If cel.Font.Bold And Len(Trim$(cel.Value)) > 0 Then
                Exit For
            End If
        End If

        ' Sjekk om cellen er del av aktiviteten
        If FargeAvstand(cel.Interior.Color, aktivFarge) <= 10 Then
            ' Dette er del av aktiviteten
            ' Sjekk om cellen er tom (brukeren har slettet innholdet)
            Dim erTom As Boolean
            erTom = False

            ' Sjekk om cellen er markert for sletting i dictionary
            If cellerÅFjerne.Exists(radKey) Then
                If cellerÅFjerne(radKey).Exists(c) Then
                    erTom = True
                End If
            End If

            ' Sjekk om cellen er tom (uavhengig av bold-formatering)
            ' Dette fanger opp når brukeren trykker Delete på alle celler
            If Len(Trim$(cel.Value)) = 0 Then
                erTom = True
            End If

            ' Hvis denne cellen IKKE er tom, hele aktiviteten slettes IKKE
            If Not erTom Then
                SjekkOmHeleAktivitetenSlettes = False
                Exit Function
            End If
        ElseIf Not HarAktivitetsfarge(cel) Then
            ' Hvit celle - aktiviteten er slutt
            Exit For
        Else
            ' Annen farge - aktiviteten er slutt
            Exit For
        End If
    Next c
End Function

' ============================================================================
' SJEKK: Er cellen en standalone celle (ingen naboer med samme farge)?
' Tar hensyn til at naboer også kan være markert for sletting
' ============================================================================
Private Function ErStandaloneCelle(ByVal cel As Range, ByVal førsteDatoKol As Long, _
                                    ByVal cellerÅFjerne As Object) As Boolean
    Dim celFarge As Long
    Dim venstreCelle As Range, høyreCelle As Range
    Dim radKey As String

    celFarge = cel.Interior.Color
    radKey = CStr(cel.Row)
    ErStandaloneCelle = False  ' Default: ikke standalone

    ' Sjekk om vi er i første kolonne
    If cel.Column <= førsteDatoKol Then
        ' Sjekk bare høyre
        If cel.Column < Me.Columns.Count Then
            Set høyreCelle = cel.Offset(0, 1)

            ' Sjekk om høyre nabo er markert for sletting
            Dim høyreMarkert As Boolean
            høyreMarkert = False
            If Not cellerÅFjerne Is Nothing Then
                If cellerÅFjerne.Exists(radKey) Then
                    If cellerÅFjerne(radKey).Exists(høyreCelle.Column) Then
                        høyreMarkert = True
                    End If
                End If
            End If

            ' Kun tell som nabo hvis IKKE markert for sletting
            If Not høyreMarkert Then
                If HarAktivitetsfarge(høyreCelle) And FargeAvstand(høyreCelle.Interior.Color, celFarge) <= 10 Then
                    ' Har nabo til høyre med samme farge ? ikke standalone
                    Exit Function
                End If
            End If
        End If
        ' Ingen nabo til høyre ? standalone
        ErStandaloneCelle = True
        Exit Function
    End If

    ' Sjekk venstre nabo
    Set venstreCelle = cel.Offset(0, -1)
    Dim harVenstreNabo As Boolean
    harVenstreNabo = False

    ' Sjekk om venstre nabo er markert for sletting
    Dim venstreMarkert As Boolean
    venstreMarkert = False
    If Not cellerÅFjerne Is Nothing Then
        If cellerÅFjerne.Exists(radKey) Then
            If cellerÅFjerne(radKey).Exists(venstreCelle.Column) Then
                venstreMarkert = True
            End If
        End If
    End If

    ' Kun tell som nabo hvis IKKE markert for sletting
    If Not venstreMarkert Then
        If HarAktivitetsfarge(venstreCelle) Then
            If FargeAvstand(venstreCelle.Interior.Color, celFarge) <= 10 Then
                harVenstreNabo = True
            End If
        End If
    End If

    ' Sjekk høyre nabo
    Dim harHøyreNabo As Boolean
    harHøyreNabo = False

    If cel.Column < Me.Columns.Count Then
        Set høyreCelle = cel.Offset(0, 1)

        ' Sjekk om høyre nabo er markert for sletting
        høyreMarkert = False
        If Not cellerÅFjerne Is Nothing Then
            If cellerÅFjerne.Exists(radKey) Then
                If cellerÅFjerne(radKey).Exists(høyreCelle.Column) Then
                    høyreMarkert = True
                End If
            End If
        End If

        ' Kun tell som nabo hvis IKKE markert for sletting
        If Not høyreMarkert Then
            If HarAktivitetsfarge(høyreCelle) Then
                If FargeAvstand(høyreCelle.Interior.Color, celFarge) <= 10 Then
                    harHøyreNabo = True
                End If
            End If
        End If
    End If

    ' Standalone = ingen naboer med samme farge (som ikke også blir slettet)
    ErStandaloneCelle = (Not harVenstreNabo And Not harHøyreNabo)
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
' KOMMENTAR-HJELPERE: Hent, fjern og legg til (støtter både Note og Threaded)
' ============================================================================
Private Function HentKommentarFraCelle(ByVal cel As Range, ByRef tekst As String, _
                                       ByRef erThreaded As Boolean) As Boolean
    Dim kommentarObj As Object

    tekst = ""
    erThreaded = False
    HentKommentarFraCelle = False

    On Error Resume Next
    Set kommentarObj = cel.Comment
    On Error GoTo 0

    If Not kommentarObj Is Nothing Then
        tekst = kommentarObj.Text
        erThreaded = False
        HentKommentarFraCelle = True
        Exit Function
    End If

    On Error Resume Next
    Set kommentarObj = cel.CommentThreaded
    If Err.Number = 0 Then
        If Not kommentarObj Is Nothing Then
            tekst = kommentarObj.Text
            erThreaded = True
            HentKommentarFraCelle = True
        End If
    Else
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Sub FjernKommentarer(ByVal cel As Range)
    On Error Resume Next
    cel.ClearComments
    cel.DeleteThreadedComment
    On Error GoTo 0
End Sub

Private Sub LeggTilKommentar(ByVal cel As Range, ByVal tekst As String, ByVal erThreaded As Boolean)
    On Error Resume Next

    If erThreaded Then
        cel.AddCommentThreaded tekst
        If Err.Number <> 0 Then
            Err.Clear
            cel.AddComment tekst
        End If
    Else
        cel.AddComment tekst
        If Err.Number <> 0 Then
            Err.Clear
            cel.AddCommentThreaded tekst
        End If
    End If

    On Error GoTo 0
End Sub

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
        visTekst = visTekst & " - " & beskrivelse
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
' FIKS: Gjenopprett multi-celle aktivitet etter paste/drag
' ============================================================================
Private Sub FiksMultiCelleAktivitet(ByVal startCel As Range, ByVal sluttKol As Long, ByVal farge As Long)
    Dim rng As Range
    Dim c As Range

    ' Definer hele blokken
    Set rng = Me.Range(startCel, Me.Cells(startCel.Row, sluttKol))

    ' Sett farge på alle celler (clear content i celle 2-5 for xlCenterAcrossSelection)
    For Each c In rng.Cells
        If c.Address <> startCel.Address Then
            c.ClearContents
        End If
        c.Interior.Color = farge
        c.Interior.Pattern = xlSolid
    Next c

    ' Fjern alle borders først
    rng.Borders.LineStyle = xlLineStyleNone

    ' Sett thick border rundt hele blokken
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 0, 0)
    End With

    ' Fjern inside borders
    rng.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    rng.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

    ' Sett alignment og font
    With rng
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Sett tekstfarge basert på bakgrunn
    With startCel.Font
        .Bold = True
        If ErLysFarge(farge) Then
            .Color = RGB(0, 0, 0)  ' Sort tekst på lys bakgrunn
        Else
            .Color = RGB(255, 255, 255)  ' Hvit tekst på mørk bakgrunn
        End If
    End With

    If DEBUG_MODE Then
        Debug.Print "Fikset multi-celle aktivitet: " & rng.Address
    End If
End Sub

' ============================================================================
' UTILITY: Restaurer rutenett for en enkelt celle
' ============================================================================
Private Sub RestaurerRutenettForCelle(ByVal cel As Range)
    ' Nullstill alle borders først
    cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
    cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone

    ' Sett standard rutenett (tynne kanter)
    With cel.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
    With cel.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With

    If DEBUG_MODE Then
        Debug.Print "Restaurert rutenett for: " & cel.Address
    End If
End Sub

' ============================================================================
' UTILITY: Sjekk om rad er tom i alle dato-kolonner
' ============================================================================
Private Function RadErTomIAlleDatoKolonner(ByVal radNr As Long, ByVal førsteDatoKol As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long
    Dim cel As Range

    ' Gå gjennom alle dato-kolonner
    For c = førsteDatoKol To lastCol
        Set cel = Me.Cells(radNr, c)

        ' Sjekk om celle har innhold
        If Len(Trim$(cel.Value)) > 0 Then
            RadErTomIAlleDatoKolonner = False
            Exit Function
        End If

        ' Sjekk om celle har farge (ikke hvit)
        If cel.Interior.ColorIndex <> xlColorIndexNone Then
            If cel.Interior.Color <> RGB(255, 255, 255) Then
                RadErTomIAlleDatoKolonner = False
                Exit Function
            End If
        End If
    Next c

    ' Hvis vi kom hit, er raden tom
    RadErTomIAlleDatoKolonner = True
End Function

' ============================================================================
' UTILITY: Reset event handler (for debugging)
' ============================================================================
Public Sub ResetEventHandler()
    behandlerEndring = False
    Application.EnableEvents = True
    Debug.Print "Event handler reset - Ark1 (Planlegger med aktivitetsbeskyttelse)"
End Sub







