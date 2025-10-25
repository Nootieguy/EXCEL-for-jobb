Option Explicit

' =================== KONFIG ===================
Private Const ARK_PLAN As String = "Planlegger"

' Alle verdier som Property Get for konsistens
Public Property Get FØRSTE_DATAKOL() As Long
    FØRSTE_DATAKOL = Worksheets(ARK_PLAN).Range("FirstDate").Column
End Property

Public Property Get datoRad() As Long
    datoRad = Worksheets(ARK_PLAN).Range("FirstDate").Row
End Property

Public Property Get FØRSTE_PERSONRAD() As Long
    FØRSTE_PERSONRAD = Worksheets(ARK_PLAN).Range("PersonHeader").Row + 1
End Property

Public Property Get FJERN_TOMME_UNDERRADER() As Boolean
    FJERN_TOMME_UNDERRADER = True
End Property
' =============================================
'
'  v3.4  Dynamisk versjon med Named Ranges
'  Bruker PersonHeader og FirstDate
'  Legger alltid tilbake heltrukken toppkant over hele raden
'  Auto-slett tomme rader, auto-flytt opp aktivitet
'
Public Sub FjernAktivitetPåMarkering()
    Dim ws As Worksheet
    Dim sel As Range, area As Range, rng As Range
    Dim lastDatoCol As Long
    Dim r As Long, c As Long
    Dim hovedRad As Long
    Dim berørteHovedrader As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ARK_PLAN)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Finner ikke arket '" & ARK_PLAN & "'.", vbCritical
        Exit Sub
    End If

    If TypeName(Selection) <> "Range" Then
        MsgBox "Marker et område i '" & ARK_PLAN & "' først.", vbExclamation
        Exit Sub
    End If
    Set sel = Intersect(Selection, ws.UsedRange)
    If sel Is Nothing Then
        MsgBox "Markeringen er tom.", vbExclamation
        Exit Sub
    End If

    Set berørteHovedrader = CreateObject("Scripting.Dictionary")
    Dim splittRader As Object
    Set splittRader = CreateObject("Scripting.Dictionary")

    ' Dictionary for å spore hvilke celler som skal fjernes (rad -> kolonner)
    Dim cellerÅFjerne As Object
    Set cellerÅFjerne = CreateObject("Scripting.Dictionary")

    ' Lagre undo-snapshot før endringer
    On Error Resume Next
    LagUndoSnapshot sel
    On Error GoTo 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False  ' Disable events for å unngå rekursjon
    lastDatoCol = SisteDatoKolonne(ws, datoRad)
    If lastDatoCol < FØRSTE_DATAKOL Then lastDatoCol = FØRSTE_DATAKOL

    ' STEG 1: Identifiser hvilke celler som skal fjernes OG hvilke rader som har splits
    For Each area In sel.Areas
        Set rng = Intersect(area, ws.Range(ws.Cells(FØRSTE_PERSONRAD, FØRSTE_DATAKOL), _
                                           ws.Cells(ws.Rows.Count, lastDatoCol)))
        If Not rng Is Nothing Then
            For r = rng.Row To rng.Row + rng.Rows.Count - 1
                hovedRad = FinnHovedRad(ws, r)
                If hovedRad >= FØRSTE_PERSONRAD Then berørteHovedrader(CStr(hovedRad)) = True

                ' Marker raden for split-sjekk (hvis den har aktiviteter)
                If RadHarAktivitet(ws, r) Then
                    If Not splittRader.Exists(r) Then splittRader.Add r, r
                End If

                ' Lagre hvilke celler som skal fjernes
                Dim radKey As String
                radKey = CStr(r)
                For c = rng.Column To rng.Column + rng.Columns.Count - 1
                    If c >= FØRSTE_DATAKOL And c <= lastDatoCol Then
                        If Not cellerÅFjerne.Exists(radKey) Then
                            Dim nyCol As Object
                            Set nyCol = CreateObject("Scripting.Dictionary")
                            cellerÅFjerne.Add radKey, nyCol
                        End If
                        cellerÅFjerne(radKey).Add c, c
                    End If
                Next c
            Next r
        End If
    Next area

    ' STEG 2: Håndter splits FØR vi fjerner cellene (så vi kan se fargene)
    Dim arkPlan As Object
    Set arkPlan = ws  ' ws er allerede Planlegger-arket
    Dim radNr As Variant

    For Each radNr In splittRader.Keys
        ' Sjekk om raden vil ha split etter fjerning
        If SjekkOmRadVilHaSplit(ws, CLng(radNr), FØRSTE_DATAKOL, lastDatoCol, cellerÅFjerne) Then
            ' Kall split-håndteringen MED liste over celler som skal fjernes
            arkPlan.HåndterAlleAktiviteterMedSplitIRad CLng(radNr), FØRSTE_DATAKOL, datoRad, cellerÅFjerne
        End If
    Next radNr

    ' STEG 3: Nå rydd opp cellene som ikke ble håndtert av split-logikken
    ' Bare rydd celler som er i cellerÅFjerne OG fortsatt har farge
    For Each radKey In cellerÅFjerne.Keys
        Dim radNum As Long
        radNum = CLng(radKey)
        hovedRad = FinnHovedRad(ws, radNum)

        Dim colDictFinal As Object
        Set colDictFinal = cellerÅFjerne(radKey)

        Dim colFinal As Variant
        For Each colFinal In colDictFinal.Keys
            Dim cFinal As Long
            cFinal = CLng(colFinal)

            ' Kun rydd hvis cellen fortsatt har farge (split-logikken fjernet ikke den)
            If ws.Cells(radNum, cFinal).Interior.ColorIndex <> xlColorIndexNone Then
                RyddCelleTilHvitMedGrid ws, radNum, cFinal
            End If
        Next colFinal

        ' Trekk toppkant som én sammenhengende linje over hele raden
        TrekkToppkantHeleRaden ws, radNum, FØRSTE_DATAKOL, lastDatoCol

        If FJERN_TOMME_UNDERRADER Then SlettTomUnderRadHvisAktuell ws, radNum, hovedRad
    Next radKey

    ' Etter rydding: komprimer hver berørt personblokk
    Dim k As Variant
    For Each k In berørteHovedrader.Keys
        KomprimerBlokkFlyttOppHvisEnesteUnder ws, CLng(k)
    Next k

    ' Sikre at alle person-skillelinjer er på plass
    GjenopprettPersonSkiller ws

    Application.EnableEvents = True  ' Re-enable events
    Application.ScreenUpdating = True
End Sub

' ----------- KJERNE: Komprimering / flytt opp -----------

Private Sub KomprimerBlokkFlyttOppHvisEnesteUnder(ws As Worksheet, ByVal hovedRad As Long)
    Dim lastRow As Long, startBlokk As Long, endBlokk As Long
    Dim r As Long, underMedAktivitet As Long, antUnderMedAktivitet As Long
    Dim lastCol As Long

    ' Finn blokk for denne hovedraden
    startBlokk = hovedRad
    endBlokk = hovedRad
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = SisteDatoKolonne(ws, datoRad)

    ' Finn enden av blokken (påfølgende rader med tom kol A)
    For r = hovedRad + 1 To lastRow
        If Len(Trim$(ws.Cells(r, 1).Value)) = 0 Then
            endBlokk = r
        Else
            Exit For
        End If
    Next r

    ' Tell under-rader som fortsatt har aktivitet
    antUnderMedAktivitet = 0
    underMedAktivitet = 0
    For r = startBlokk + 1 To endBlokk
        If RadHarAktivitet(ws, r) Then
            antUnderMedAktivitet = antUnderMedAktivitet + 1
            underMedAktivitet = r
        End If
    Next r

    ' Hvis hovedraden er tom og det finnes nøyaktig én under-rad med aktivitet  flytt opp
    If Not RadHarAktivitet(ws, hovedRad) And antUnderMedAktivitet = 1 Then
        FlyttRadInnholdOpp ws, underMedAktivitet, hovedRad
        ws.Rows(underMedAktivitet).Delete
    End If

    ' Etter sletting: fjern eventuelle gjenværende tomme under-rader
    Dim slettet As Boolean: slettet = False
    For r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To hovedRad + 1 Step -1
        If Len(Trim$(ws.Cells(r, 1).Value)) = 0 Then
            If RadErTomIAlleDatoKolonner(ws, r) Then
                ws.Rows(r).Delete
                slettet = True
            End If
        End If
    Next r
    
    ' KRITISK FIX: Hvis vi slettet noen under-rader, gjenopprett toppkant på neste rad
    If slettet Then
        Dim nesteRad As Long
        ' Finn neste rad med navn (neste person)
        For r = hovedRad + 1 To ws.Rows.Count
            If Len(Trim$(ws.Cells(r, 1).Value)) > 0 Then
                nesteRad = r
                Exit For
            End If
        Next r
        
        ' Hvis vi fant en neste person, gjenopprett toppkanten
        If nesteRad > hovedRad Then
            Dim rngTop As Range
            Set rngTop = ws.Range(ws.Cells(nesteRad, FØRSTE_DATAKOL), ws.Cells(nesteRad, lastCol))
            With rngTop.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End If
    End If
End Sub

Private Sub FlyttRadInnholdOpp(ws As Worksheet, ByVal srcRad As Long, ByVal dstRad As Long)
    Dim lastCol As Long
    lastCol = SisteDatoKolonne(ws, datoRad)

    ' Kopier ALT innhold/format fra srcRad (dato-kolonner) til dstRad
    ws.Range(ws.Cells(srcRad, FØRSTE_DATAKOL), ws.Cells(srcRad, lastCol)).Copy
    ws.Cells(dstRad, FØRSTE_DATAKOL).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
End Sub

' ----------- Rydding / strektegning -----------

Private Sub RyddCelleTilHvitMedGrid(ws As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim cel As Range, under As Range
    Set cel = ws.Cells(r, c)

    ' 1) Rens innhold/kommentar/tekstformat
    cel.ClearComments
    cel.ClearContents
    cel.Font.Bold = False
    cel.Font.ColorIndex = xlColorIndexAutomatic
    cel.HorizontalAlignment = xlGeneral
    cel.VerticalAlignment = xlCenter
    cel.WrapText = False

    ' 2) Sett bakgrunn til ren hvit (ingen mønster)
    With cel.Interior
        .Pattern = xlSolid
        .TintAndShade = 0
        .Color = RGB(255, 255, 255)
        .PatternTintAndShade = 0
    End With

    ' 3) Slå av diagonale kanter (for å hindre X-kryss)
    cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
    cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone

    ' 4) Nullstill alle kanter og sett tynne ytterkanter (normalt rutenett)
    cel.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
    cel.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone
    cel.Borders(xlInsideVertical).LineStyle = xlLineStyleNone
    cel.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone

    cel.Borders(xlEdgeLeft).LineStyle = xlContinuous
    cel.Borders(xlEdgeLeft).Weight = xlThin
    cel.Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
    
    cel.Borders(xlEdgeRight).LineStyle = xlContinuous
    cel.Borders(xlEdgeRight).Weight = xlThin
    cel.Borders(xlEdgeRight).Color = RGB(0, 0, 0)
    
    cel.Borders(xlEdgeTop).LineStyle = xlContinuous
    cel.Borders(xlEdgeTop).Weight = xlThin
    cel.Borders(xlEdgeTop).Color = RGB(0, 0, 0)
    
    cel.Borders(xlEdgeBottom).LineStyle = xlContinuous
    cel.Borders(xlEdgeBottom).Weight = xlThin
    cel.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)

    ' 5) Sikre at cellen under ikke lekker farge fra tidligere blokk
    If r < ws.Rows.Count Then
        Set under = ws.Cells(r + 1, c)
        under.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
        under.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
        With under.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If
End Sub

' Trekk en sammenhengende toppkant over hele raden (datokolonner)
Private Sub TrekkToppkantHeleRaden(ws As Worksheet, ByVal rad As Long, ByVal colStart As Long, ByVal colEnd As Long)
    Dim rr As Range
    Set rr = ws.Range(ws.Cells(rad, colStart), ws.Cells(rad, colEnd))
    With rr.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
End Sub

Private Sub SlettTomUnderRadHvisAktuell(ws As Worksheet, ByVal r As Long, ByVal hovedRad As Long)
    If r <= hovedRad Then Exit Sub
    If Len(Trim$(ws.Cells(r, 1).Value)) > 0 Then Exit Sub
    If RadErTomIAlleDatoKolonner(ws, r) Then ws.Rows(r).Delete
End Sub

' ----------- Tilstandssjekker -----------

Private Function RadErTomIAlleDatoKolonner(ws As Worksheet, ByVal r As Long) As Boolean
    Dim lastCol As Long, c As Long, cel As Range
    lastCol = SisteDatoKolonne(ws, datoRad)
    For c = FØRSTE_DATAKOL To lastCol
        Set cel = ws.Cells(r, c)
        If Len(Trim$(cel.Value)) > 0 Then Exit Function
        If cel.Interior.ColorIndex <> xlColorIndexNone Then
            If cel.Interior.Color <> RGB(255, 255, 255) Then Exit Function
        End If
    Next c
    RadErTomIAlleDatoKolonner = True
End Function

Private Function RadHarAktivitet(ws As Worksheet, ByVal r As Long) As Boolean
    Dim lastCol As Long, c As Long, cel As Range
    lastCol = SisteDatoKolonne(ws, datoRad)
    For c = FØRSTE_DATAKOL To lastCol
        Set cel = ws.Cells(r, c)
        If Len(Trim$(cel.Value)) > 0 Then RadHarAktivitet = True: Exit Function
        If cel.Interior.ColorIndex <> xlColorIndexNone Then
            If cel.Interior.Color <> RGB(255, 255, 255) Then RadHarAktivitet = True: Exit Function
        End If
        If cel.Font.Bold And Len(Trim$(cel.Value)) > 0 Then RadHarAktivitet = True: Exit Function
    Next c
End Function

' ----------- Navigasjon / hjelpefunksjoner -----------

Private Function FinnHovedRad(ws As Worksheet, ByVal rad As Long) As Long
    Dim r As Long
    For r = rad To FØRSTE_PERSONRAD Step -1
        If Len(Trim$(ws.Cells(r, 1).Value)) > 0 Then FinnHovedRad = r: Exit Function
    Next r
    FinnHovedRad = rad
End Function

Public Function SisteDatoKolonne(ws As Worksheet, ByVal headerRow As Long) As Long
    SisteDatoKolonne = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
End Function

' Gjenopprett toppkanter mellom alle personblokker
Private Sub GjenopprettPersonSkiller(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = SisteDatoKolonne(ws, datoRad)
    
    ' Gå gjennom alle rader og finn personer (navn i kolonne A)
    For r = FØRSTE_PERSONRAD To lastRow
        If Len(Trim$(ws.Cells(r, 1).Value)) > 0 Then
            ' Dette er en personrad - sett toppkant
            Dim rngTop As Range
            Set rngTop = ws.Range(ws.Cells(r, FØRSTE_DATAKOL), ws.Cells(r, lastCol))
            With rngTop.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(0, 0, 0)
            End With
        End If
    Next r
End Sub

' ----------- Split-deteksjon -----------

' Sjekk om en rad VIL ha split etter at celler fjernes
Private Function SjekkOmRadVilHaSplit(ws As Worksheet, ByVal r As Long, _
                                       ByVal startCol As Long, ByVal endCol As Long, _
                                       ByVal cellerÅFjerne As Object) As Boolean
    Dim c As Long
    Dim harSettFarge As Boolean
    Dim harSettHvit As Boolean
    Dim radKey As String

    harSettFarge = False
    harSettHvit = False
    radKey = CStr(r)

    ' Skann raden fra venstre til høyre
    For c = startCol To endCol
        Dim cel As Range
        Set cel = ws.Cells(r, c)

        ' Sjekk om denne cellen skal fjernes
        Dim skalFjernes As Boolean
        skalFjernes = False
        If cellerÅFjerne.Exists(radKey) Then
            Dim colDict As Object
            Set colDict = cellerÅFjerne(radKey)
            If colDict.Exists(c) Then
                skalFjernes = True
            End If
        End If

        ' Sjekk om cellen har aktivitetsfarge (og ikke skal fjernes)
        Dim harAktivFarge As Boolean
        harAktivFarge = False

        If Not skalFjernes Then
            ' Bruk samme logikk som HarAktivitetsfarge fra Ark1.cls
            If cel.Interior.ColorIndex <> xlColorIndexNone Then
                Dim celCol As Long
                celCol = cel.Interior.Color
                If celCol <> RGB(255, 255, 255) And celCol <> RGB(242, 242, 242) And celCol <> RGB(250, 250, 250) Then
                    harAktivFarge = True
                End If
            End If
        End If

        If harAktivFarge Then
            ' Hvis vi har sett hvit/fjernet før → dette er en split!
            If harSettHvit Then
                SjekkOmRadVilHaSplit = True
                Exit Function
            End If
            harSettFarge = True
        Else
            ' Hvit celle eller skal fjernes
            If harSettFarge Then
                harSettHvit = True
            End If
        End If
    Next c

    SjekkOmRadVilHaSplit = False
End Function

