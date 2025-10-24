VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ark5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ============================================================================
' ARK5: AKTIVITETSOVERSIKT - WORKSHEET EVENT HANDLER (FORBEDRET v4.0)
' ============================================================================
' FORMÅL:
'   Håndterer automatisk synkronisering mellom AKTIVITETSOVERSIKT og Planlegger
'   når brukeren endrer data direkte i oversikten.
'
' FUNKSJONALITET:
'   1. PERSON-SWITCHING: Endre person ? flytt aktivitet mellom personer
'   2. KODE-VALIDERING: Sjekk at aktivitetskoden finnes i AKTIVITETSTYPER
'   3. DATO-VALIDERING: Sjekk at startdato ikke er etter sluttdato
'   4. FORSINKELSE-HÅNDTERING: Respekter forsinkelse ved flytting
'
' REKURSJONSBESKYTTELSE:
'   - behandlerEndring flagg forhindrer evige loops
'   - Application.EnableEvents sjekk forhindrer meldinger under refresh
'
' FORBEDRINGER v4.0:
'   - Mer robust søkelogikk (krever ikke fet tekst)
'   - Detaljert debugging til Immediate Window
'   - Bedre feilmeldinger med diagnoseinformasjon
'   - Ny UltraBred-søkemetode som fallback
'   - Debug-funksjon for å liste alle aktiviteter
' ============================================================================

' ============================================================================
' SEKSJON 1: KONSTANTER OG GLOBALE VARIABLER
' ============================================================================
' Definerer tabellstruktur og beskyttelsesflagg

' Første rad med data i tabellen
Private Const TBL_START_ROW As Long = 4

' Kolonneindekser (hvilken kolonne inneholder hva)
Private Const COL_PERSON As Long = 1       ' A: Personnavn
Private Const COL_KODE As Long = 2         ' B: Aktivitetskode (TL, VA, osv)
Private Const COL_BESKR As Long = 3        ' C: Beskrivelse (auto-fylles)
Private Const COL_OPP_START As Long = 4    ' D: Opprinnelig startdato
Private Const COL_OPP_SLUTT As Long = 5    ' E: Opprinnelig sluttdato
Private Const COL_FORSINKET As Long = 6    ' F: Forsinkelse i dager (redigerbar)
Private Const COL_NY_SLUTT As Long = 7     ' G: Ny sluttdato (beregnet = E + F)
Private Const COL_KOMMENTAR As Long = 10   ' J: Kommentar (valgfri)

' Beskyttelsesflagg mot rekursjon
' Settes til True når vi behandler endring for å unngå nested calls
Private behandlerEndring As Boolean

' Debug-modus (sett til True for detaljert logging)
Private Const DEBUG_MODE As Boolean = True


' ============================================================================
' SEKSJON 2: HOVEDEVENT - WORKSHEET_CHANGE
' ============================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Slutt

    ' STEG 1: Sjekk om vi allerede behandler (unngå rekursjon)
    If behandlerEndring Then Exit Sub

    ' STEG 2: Valider input
    If Target Is Nothing Then Exit Sub
    If Target.Row < TBL_START_ROW Then Exit Sub

    ' STEG 3: Sett beskyttelsesflagg
    behandlerEndring = True
    Application.EnableEvents = False

    ' STEG 4: Behandle hver endret celle
    Dim c As Range
    Dim antall As Long

    For Each c In Target.Cells
        antall = antall + 1

        ' Sikkerhet: maks 50 celler samtidig
        If antall > 50 Then
            If Application.EnableEvents Then
                MsgBox "For mange celler endret samtidig.", vbExclamation
            End If
            Exit For
        End If

        ' Kun behandle rader fra TBL_START_ROW og nedover
        If c.Row >= TBL_START_ROW Then
            ' Rute til riktig handler basert på hvilken kolonne som ble endret
            Select Case c.Column
                Case COL_PERSON
                    Call HåndterPersonEndring(c.Row)

                Case COL_KODE
                    Call HåndterKodeEndring(c.Row)

                Case COL_OPP_START, COL_OPP_SLUTT
                    Call HåndterDatoEndring(c.Row)

                Case COL_FORSINKET
                    Call HåndterForsinkelseEndring(c.Row)
            End Select
        End If
    Next c

Slutt:
    ' STEG 5: Re-enable events og fjern beskyttelsesflagg
    Application.EnableEvents = True
    behandlerEndring = False

    ' Debug-logging ved feil
    If Err.Number <> 0 Then
        Debug.Print "Feil i Worksheet_Change: " & Err.Description
        Err.Clear
    End If
End Sub


' ============================================================================
' SEKSJON 3: PERSON-SWITCHING HANDLER (FORBEDRET)
' ============================================================================
Private Sub HåndterPersonEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler

    ' --- STEG 1: HENT DATA FRA RADEN ---
    Dim nyPerson As String
    Dim kode As String
    Dim startDato As Date
    Dim sluttDato As Date
    Dim kommentar As String
    Dim forsinkelse As Long
    Dim nySluttDato As Date
    Dim gammelPerson As String
    Dim wsP As Worksheet

    ' Hent person og kode
    nyPerson = Trim$(Me.Cells(rad, COL_PERSON).Value)
    kode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))

    ' Valider at vi har nødvendig data
    If Len(nyPerson) = 0 Or Len(kode) = 0 Then Exit Sub

    ' Hent datoer og forsinkelse
    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    forsinkelse = Me.Cells(rad, COL_FORSINKET).Value
    kommentar = Trim$(Me.Cells(rad, COL_KOMMENTAR).Value)
    On Error GoTo ErrorHandler

    ' --- STEG 2: BEREGN NY SLUTTDATO (MED FORSINKELSE) ---
    If forsinkelse > 0 Then
        nySluttDato = sluttDato + forsinkelse
    Else
        nySluttDato = sluttDato
    End If

    ' Valider datoer
    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub

    ' --- STEG 3: FINN PLANLEGGER-ARKET ---
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets("Planlegger")
    On Error GoTo ErrorHandler

    If wsP Is Nothing Then
        If Application.EnableEvents Then
            MsgBox "Finner ikke Planlegger-arket!", vbCritical
        End If
        Exit Sub
    End If

    ' --- DEBUG LOGGING ---
    If DEBUG_MODE Then
        Debug.Print String(80, "=")
        Debug.Print "PERSON-SWITCHING DEBUG [" & Now & "]"
        Debug.Print String(80, "=")
        Debug.Print "Søker etter aktivitet:"
        Debug.Print "  Kode:        " & kode
        Debug.Print "  StartDato:   " & Format(startDato, "dd.mm.yyyy")
        Debug.Print "  SluttDato:   " & Format(sluttDato, "dd.mm.yyyy")
        Debug.Print "  Forsinkelse: " & forsinkelse & " dager"
        Debug.Print "  Ny slutt:    " & Format(nySluttDato, "dd.mm.yyyy")
        Debug.Print "  Ny person:   " & nyPerson
        Debug.Print ""
    End If

    ' --- STEG 4: FINN GAMMEL PERSON (4 SØKEMETODER MED FALLBACK) ---
    ' Metode 1: Eksakt match med ny slutt (hvis forsinket)
    If forsinkelse > 0 Then
        If DEBUG_MODE Then Debug.Print "Metode 1: Søker med ny sluttdato..."
        gammelPerson = FinnPersonForAktivitet(wsP, kode, startDato, nySluttDato)
        If DEBUG_MODE And Len(gammelPerson) > 0 Then Debug.Print "  ? Funnet: " & gammelPerson
    End If

    ' Metode 2: Eksakt match med opprinnelig slutt
    If Len(gammelPerson) = 0 Then
        If DEBUG_MODE Then Debug.Print "Metode 2: Søker med opprinnelig sluttdato..."
        gammelPerson = FinnPersonForAktivitet(wsP, kode, startDato, sluttDato)
        If DEBUG_MODE And Len(gammelPerson) > 0 Then Debug.Print "  ? Funnet: " & gammelPerson
    End If

    ' Metode 3: Bred søk (kun kode og start, krever fet tekst)
    If Len(gammelPerson) = 0 Then
        If DEBUG_MODE Then Debug.Print "Metode 3: Bred søk (kun startdato)..."
        gammelPerson = FinnPersonForAktivitetBred(wsP, kode, startDato)
        If DEBUG_MODE And Len(gammelPerson) > 0 Then Debug.Print "  ? Funnet: " & gammelPerson
    End If

    ' Metode 4: ULTRA BRED - søk uten krav om fet tekst (ny!)
    If Len(gammelPerson) = 0 Then
        If DEBUG_MODE Then Debug.Print "Metode 4: Ultra bred søk (ingen formateringskrav)..."
        gammelPerson = FinnPersonForAktivitetUltraBred(wsP, kode, startDato)
        If DEBUG_MODE And Len(gammelPerson) > 0 Then Debug.Print "  ? Funnet: " & gammelPerson
    End If

    ' Hvis vi fortsatt ikke fant aktiviteten, vis forbedret feilmelding
    If Len(gammelPerson) = 0 Then
        If DEBUG_MODE Then
            Debug.Print "  ? IKKE FUNNET!"
            Debug.Print ""
            Debug.Print "Diagnose: Kjør 'Call Ark5.DebugListAktiviteter' i Immediate Window"
            Debug.Print "for å se alle aktiviteter i Planlegger."
            Debug.Print String(80, "=")
        End If

        If Application.EnableEvents Then
            MsgBox "Fant ikke aktiviteten i Planlegger!" & vbCrLf & vbCrLf & _
                   "Søkte etter:" & vbCrLf & _
                   "  Kode: " & kode & vbCrLf & _
                   "  Start: " & Format(startDato, "dd.mm.yyyy") & vbCrLf & _
                   "  Slutt: " & Format(sluttDato, "dd.mm.yyyy") & vbCrLf & vbCrLf & _
                   "Prøvde 4 søkemetoder uten resultat." & vbCrLf & vbCrLf & _
                   "DEBUGGING:" & vbCrLf & _
                   "Åpne VBA Editor (Alt+F11)" & vbCrLf & _
                   "Trykk Ctrl+G (Immediate Window)" & vbCrLf & _
                   "Skriv: Call Ark5.DebugListAktiviteter" & vbCrLf & _
                   "Dette viser alle aktiviteter i Planlegger.", _
                   vbExclamation, "Kan ikke flytte aktivitet"
        End If
        Exit Sub
    End If

    If DEBUG_MODE Then
        Debug.Print ""
        Debug.Print "Gammel person: " & gammelPerson
        Debug.Print String(80, "=")
    End If

    ' --- STEG 5: SJEKK OM DET ER EN FAKTISK ENDRING ---
    If StrComp(gammelPerson, nyPerson, vbTextCompare) = 0 Then Exit Sub

    ' --- STEG 6: FLYTT AKTIVITET TIL NY PERSON ---
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!FlyttAktivitetTilNyPerson", _
                    gammelPerson, nyPerson, kode, _
                    startDato, nySluttDato, kommentar

    If Err.Number <> 0 Then
        If Application.EnableEvents Then
            MsgBox "Feil ved flytting av aktivitet:" & vbCrLf & vbCrLf & _
                   Err.Description & vbCrLf & vbCrLf & _
                   "Kontroller at Module3.FlyttAktivitetTilNyPerson finnes.", vbCritical
        End If
        Err.Clear
    Else
        If DEBUG_MODE Then Debug.Print "? Aktivitet flyttet fra " & gammelPerson & " til " & nyPerson
    End If
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    If Application.EnableEvents Then
        MsgBox "Feil ved person-endring:" & vbCrLf & Err.Description, vbCritical
    End If
    Debug.Print "FEIL i HåndterPersonEndring: " & Err.Description
End Sub


' ============================================================================
' SEKSJON 4: SØKEFUNKSJONER FOR AKTIVITETER (FORBEDRET)
' ============================================================================

' ----------------------------------------------------------------------------
' FUNKSJON: FinnPersonForAktivitet (EKSAKT MATCH)
' ----------------------------------------------------------------------------
Private Function FinnPersonForAktivitet(wsP As Worksheet, _
                                       kode As String, _
                                       startDato As Date, _
                                       sluttDato As Date) As String
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, r As Long, blockEnd As Long
    Dim startCol As Long, sluttCol As Long
    Dim cel As Range
    Dim personNavn As String

    ' Hent struktur fra Named Ranges
    On Error Resume Next
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    On Error GoTo 0

    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then
        If DEBUG_MODE Then Debug.Print "  [!] Named Ranges mangler (FirstDate/PersonHeader)"
        Exit Function
    End If

    ' Finn kolonne-indekser for start og slutt
    startCol = FinnDatoKol(wsP, startDato, datoRad, førsteDatoKol)
    sluttCol = FinnDatoKol(wsP, sluttDato, datoRad, førsteDatoKol)

    If startCol = 0 Or sluttCol = 0 Then
        If DEBUG_MODE Then
            If startCol = 0 Then Debug.Print "  [!] Fant ikke startdato-kolonne for " & Format(startDato, "dd.mm.yyyy")
            If sluttCol = 0 Then Debug.Print "  [!] Fant ikke sluttdato-kolonne for " & Format(sluttDato, "dd.mm.yyyy")
        End If
        Exit Function
    End If

    If DEBUG_MODE Then Debug.Print "  Kolonner: Start=" & startCol & ", Slutt=" & sluttCol

    ' Søk gjennom alle personer
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row

    For r = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(r, 1).Value)

        If Len(personNavn) > 0 Then
            blockEnd = r
            Do While blockEnd < lastRow
                If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                blockEnd = blockEnd + 1
            Loop

            Dim blockRow As Long
            For blockRow = r To blockEnd
                Set cel = wsP.Cells(blockRow, startCol)

                ' Sjekk om dette er start på en aktivitet
                If cel.Font.Bold And Len(Trim$(cel.Value)) > 0 Then
                    If InStr(1, Trim$(cel.Value), kode, vbTextCompare) = 1 Then
                        If VerifiserAktivitetSpenn(wsP, blockRow, startCol, sluttCol) Then
                            FinnPersonForAktivitet = personNavn
                            Exit Function
                        End If
                    End If
                End If
            Next blockRow
        End If
    Next r
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: FinnPersonForAktivitetBred (BRED SØKNING - kun startdato)
' ----------------------------------------------------------------------------
Private Function FinnPersonForAktivitetBred(wsP As Worksheet, _
                                           kode As String, _
                                           startDato As Date) As String
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, r As Long, blockEnd As Long
    Dim startCol As Long
    Dim cel As Range
    Dim personNavn As String

    On Error Resume Next
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    On Error GoTo 0

    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then Exit Function

    startCol = FinnDatoKol(wsP, startDato, datoRad, førsteDatoKol)
    If startCol = 0 Then Exit Function

    If DEBUG_MODE Then Debug.Print "  Kolonne: Start=" & startCol

    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row

    For r = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(r, 1).Value)
        If Len(personNavn) > 0 Then
            blockEnd = r
            Do While blockEnd < lastRow
                If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                blockEnd = blockEnd + 1
            Loop

            Dim blockRow As Long
            For blockRow = r To blockEnd
                Set cel = wsP.Cells(blockRow, startCol)

                ' Aksepterer både fet og ikke-fet tekst
                If Len(Trim$(cel.Value)) > 0 Then
                    If InStr(1, Trim$(cel.Value), kode, vbTextCompare) = 1 Then
                        FinnPersonForAktivitetBred = personNavn
                        If DEBUG_MODE Then Debug.Print "  Match i rad " & blockRow & ": " & cel.Value
                        Exit Function
                    End If
                End If
            Next blockRow
        End If
    Next r
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: FinnPersonForAktivitetUltraBred (NY! - mest robust)
' ----------------------------------------------------------------------------
' Søker kun basert på kode og startdato
' Krever IKKE fet tekst
' Tillater kode hvor som helst i cellen (ikke bare først)
' ----------------------------------------------------------------------------
Private Function FinnPersonForAktivitetUltraBred(wsP As Worksheet, _
                                                 kode As String, _
                                                 startDato As Date) As String
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, r As Long, blockEnd As Long
    Dim startCol As Long
    Dim cel As Range
    Dim personNavn As String
    Dim cellValue As String

    On Error Resume Next
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    On Error GoTo 0

    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then Exit Function

    startCol = FinnDatoKol(wsP, startDato, datoRad, førsteDatoKol)
    If startCol = 0 Then Exit Function

    If DEBUG_MODE Then Debug.Print "  UltraBred: Søker i kolonne " & startCol

    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row

    For r = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(r, 1).Value)
        If Len(personNavn) > 0 Then
            blockEnd = r
            Do While blockEnd < lastRow
                If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                blockEnd = blockEnd + 1
            Loop

            Dim blockRow As Long
            For blockRow = r To blockEnd
                Set cel = wsP.Cells(blockRow, startCol)
                cellValue = UCase$(Trim$(cel.Value))

                ' Søk etter kode HVOR SOM HELST i cellen (ikke kun start)
                If Len(cellValue) > 0 Then
                    If InStr(1, cellValue, UCase$(kode), vbTextCompare) > 0 Then
                        FinnPersonForAktivitetUltraBred = personNavn
                        If DEBUG_MODE Then Debug.Print "  UltraBred match i rad " & blockRow & ": '" & cel.Value & "'"
                        Exit Function
                    End If
                End If
            Next blockRow
        End If
    Next r

    If DEBUG_MODE Then Debug.Print "  UltraBred: Ingen match funnet"
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: VerifiserAktivitetSpenn
' ----------------------------------------------------------------------------
Private Function VerifiserAktivitetSpenn(wsP As Worksheet, _
                                        rad As Long, _
                                        startCol As Long, _
                                        sluttCol As Long) As Boolean
    Dim c As Long
    Dim cel As Range
    Dim startFarge As Long
    Dim alleFet As Boolean

    startFarge = wsP.Cells(rad, startCol).Interior.Color
    alleFet = True

    For c = startCol To sluttCol
        Set cel = wsP.Cells(rad, c)

        ' Celle må enten ha fet tekst ELLER samme farge som start
        If Not cel.Font.Bold And cel.Interior.Color <> startFarge Then
            alleFet = False
            Exit For
        End If
    Next c

    VerifiserAktivitetSpenn = alleFet
End Function

' ----------------------------------------------------------------------------
' FUNKSJON: FinnDatoKol
' ----------------------------------------------------------------------------
Private Function FinnDatoKol(wsP As Worksheet, _
                            søkDato As Date, _
                            datoRad As Long, _
                            førsteDatoKol As Long) As Long
    Dim lastCol As Long, c As Long
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column

    For c = førsteDatoKol To lastCol
        If IsDate(wsP.Cells(datoRad, c).Value) Then
            ' CLng sammenligner bare datodelen, ignorerer tid
            If CLng(CDate(wsP.Cells(datoRad, c).Value)) = CLng(søkDato) Then
                FinnDatoKol = c
                Exit Function
            End If
        End If
    Next c
End Function


' ============================================================================
' SEKSJON 5: KODE-VALIDERING HANDLER
' ============================================================================
Private Sub HåndterKodeEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler

    Dim nyKode As String
    Dim beskrivelse As String
    Dim farge As Long
    Dim wsTyp As Worksheet
    Dim funnet As Boolean

    nyKode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))
    If Len(nyKode) = 0 Then Exit Sub

    On Error Resume Next
    Set wsTyp = ThisWorkbook.Worksheets("AKTIVITETSTYPER - OVERSIKT")
    On Error GoTo ErrorHandler

    If wsTyp Is Nothing Then Exit Sub

    On Error Resume Next
    funnet = Application.Run("LookupAktivitet", wsTyp, nyKode, beskrivelse, farge)
    On Error GoTo ErrorHandler

    If Not funnet Then
        If Application.EnableEvents Then
            MsgBox "Aktivitetskode '" & nyKode & "' finnes ikke i AKTIVITETSTYPER - OVERSIKT!" & vbCrLf & vbCrLf & _
                   "Legg til koden først, eller velg en eksisterende kode.", vbExclamation, "Ugyldig kode"
            Me.Cells(rad, COL_KODE).Value = ""
        End If
    Else
        Me.Cells(rad, COL_BESKR).Value = beskrivelse
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Feil i HåndterKodeEndring: " & Err.Description
    Application.EnableEvents = True
End Sub


' ============================================================================
' SEKSJON 6A: FORSINKELSE-ENDRING HANDLER
' ============================================================================
Private Sub HåndterForsinkelseEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler

    Dim person As String
    Dim kode As String
    Dim startDato As Date
    Dim sluttDato As Date
    Dim forsinkelse As Long
    Dim nySluttDato As Date
    Dim kommentar As String

    person = Trim$(Me.Cells(rad, COL_PERSON).Value)
    kode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))

    If Len(person) = 0 Or Len(kode) = 0 Then Exit Sub

    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    forsinkelse = Me.Cells(rad, COL_FORSINKET).Value
    kommentar = Trim$(Me.Cells(rad, COL_KOMMENTAR).Value)
    On Error GoTo ErrorHandler

    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub

    Dim wsP As Worksheet
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets("Planlegger")
    On Error GoTo ErrorHandler

    If wsP Is Nothing Then Exit Sub

    ' Verifiser at aktiviteten finnes
    Dim testPerson As String
    testPerson = FinnPersonForAktivitetUltraBred(wsP, kode, startDato)

    If Len(testPerson) = 0 Then
        Debug.Print "HåndterForsinkelseEndring: Aktivitet " & kode & " ikke funnet i Planlegger - hopper over"
        Exit Sub
    End If

    If StrComp(testPerson, person, vbTextCompare) <> 0 Then
        Debug.Print "HåndterForsinkelseEndring: Person mismatch - forventet " & person & ", fant " & testPerson & " - hopper over"
        Exit Sub
    End If

    If forsinkelse > 0 Then
        nySluttDato = sluttDato + forsinkelse
    Else
        nySluttDato = sluttDato
    End If

    Dim opprinneligEventsStatus As Boolean
    opprinneligEventsStatus = Application.EnableEvents

    If opprinneligEventsStatus Then
        Application.EnableEvents = True
    End If

    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!SynkroniserEnkeltAktivitet", _
                    person, kode, startDato, nySluttDato, kommentar

    If Err.Number <> 0 Then
        If opprinneligEventsStatus Then
            MsgBox "Feil ved oppdatering av forsinkelse:" & vbCrLf & vbCrLf & _
                   Err.Description & vbCrLf & vbCrLf & _
                   "Kontroller at Module3.SynkroniserEnkeltAktivitet finnes.", vbCritical
        End If
        Err.Clear
    End If

    Application.EnableEvents = opprinneligEventsStatus
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    If Application.EnableEvents Then
        MsgBox "Feil ved forsinkelse-endring:" & vbCrLf & Err.Description, vbCritical
    End If
    Debug.Print "FEIL i HåndterForsinkelseEndring: " & Err.Description
End Sub


' ============================================================================
' SEKSJON 6B: DATO-VALIDERING HANDLER
' ============================================================================
Private Sub HåndterDatoEndring(ByVal rad As Long)
    On Error GoTo ErrorHandler

    Dim startDato As Date
    Dim sluttDato As Date

    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    On Error GoTo ErrorHandler

    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub

    If startDato > sluttDato Then
        If Application.EnableEvents Then
            MsgBox "Startdato kan ikke være etter sluttdato!", vbExclamation, "Ugyldig dato"
            Me.Cells(rad, COL_OPP_START).Value = sluttDato
        End If
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Feil i HåndterDatoEndring: " & Err.Description
    Application.EnableEvents = True
End Sub


' ============================================================================
' SEKSJON 7: UTILITY-FUNKSJONER OG DEBUGGING
' ============================================================================

' ----------------------------------------------------------------------------
' FUNKSJON: ResetEventHandler
' ----------------------------------------------------------------------------
Public Sub ResetEventHandler()
    behandlerEndring = False
    Application.EnableEvents = True
    Debug.Print "Event handler reset - Ark5 v4.0 (Forbedret søk + debugging)"
End Sub

' ----------------------------------------------------------------------------
' FUNKSJON: DebugListAktiviteter (NY!)
' ----------------------------------------------------------------------------
' FORMÅL: Lister alle aktiviteter i Planlegger for debugging
' BRUK: Kjør i Immediate Window: Call Ark5.DebugListAktiviteter
' ----------------------------------------------------------------------------
Public Sub DebugListAktiviteter()
    Dim wsP As Worksheet
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim personNavn As String
    Dim cel As Range
    Dim aktivitetTeller As Long

    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets("Planlegger")
    On Error GoTo 0

    If wsP Is Nothing Then
        Debug.Print "FEIL: Finner ikke Planlegger-arket!"
        Exit Sub
    End If

    On Error Resume Next
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    On Error GoTo 0

    If førsteDatoKol = 0 Or datoRad = 0 Or førstePersonRad = 0 Then
        Debug.Print "FEIL: Named Ranges mangler (FirstDate/PersonHeader)"
        Exit Sub
    End If

    Debug.Print String(80, "=")
    Debug.Print "AKTIVITETER I PLANLEGGER [" & Now & "]"
    Debug.Print String(80, "=")
    Debug.Print "Struktur:"
    Debug.Print "  Første datokol: " & førsteDatoKol
    Debug.Print "  Datorad:        " & datoRad
    Debug.Print "  Første person:  " & førstePersonRad
    Debug.Print ""

    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column

    For r = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(r, 1).Value)

        If Len(personNavn) > 0 Then
            Debug.Print "PERSON: " & personNavn & " (rad " & r & ")"

            ' Finn slutten på personblokken
            Dim blockEnd As Long
            blockEnd = r
            Do While blockEnd < lastRow
                If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                blockEnd = blockEnd + 1
            Loop

            ' Søk gjennom alle datokolonner
            For c = førsteDatoKol To lastCol
                Dim dato As Variant
                dato = wsP.Cells(datoRad, c).Value

                If IsDate(dato) Then
                    ' Sjekk alle rader i personblokken
                    Dim blockRow As Long
                    For blockRow = r To blockEnd
                        Set cel = wsP.Cells(blockRow, c)

                        If Len(Trim$(cel.Value)) > 0 Then
                            aktivitetTeller = aktivitetTeller + 1
                            Debug.Print "  [" & aktivitetTeller & "] " & _
                                       Format(dato, "dd.mm.yyyy") & " | " & _
                                       "Rad:" & blockRow & " Kol:" & c & " | " & _
                                       IIf(cel.Font.Bold, "FET", "   ") & " | " & _
                                       cel.Value
                        End If
                    Next blockRow
                End If
            Next c

            Debug.Print ""
        End If
    Next r

    Debug.Print String(80, "=")
    Debug.Print "Totalt " & aktivitetTeller & " aktiviteter funnet"
    Debug.Print String(80, "=")
End Sub


