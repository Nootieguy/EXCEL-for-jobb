

Option Explicit
'
' =================== MODUL 5 - AKTIVITETSOVERSIKT (v1.0) ===================
' Totaloversikt av alle aktiviteter med forsinkelseshåndtering og overlappdeteksjon
' Toveis synkronisering med Planlegger-arket

' ===== KONFIG =====
Private Const ARK_PLAN As String = "Planlegger"
Private Const ARK_OVERSIKT_AKT As String = "AKTIVITETSOVERSIKT"
Private Const ARK_OVERSIKT_TYP As String = "AKTIVITETSTYPER - OVERSIKT"

' Tabell
Private Const TBL_START_ROW As Long = 4
Private Const COL_PERSON As Long = 1        ' A
Private Const COL_KODE As Long = 2          ' B
Private Const COL_BESKR As Long = 3         ' C
Private Const COL_OPP_START As Long = 4     ' D - Opprinnelig startdato
Private Const COL_OPP_SLUTT As Long = 5     ' E - Opprinnelig sluttdato
Private Const COL_FORSINKET As Long = 6     ' F - Forsinkelse i dager (redigerbar)
Private Const COL_NY_SLUTT As Long = 7      ' G - Ny sluttdato (beregnet)
Private Const COL_VARIGHET As Long = 8      ' H - Varighet (dager)
Private Const COL_STATUS As Long = 9        ' I - Status (OK/OVERLAPP)
Private Const COL_KOMMENTAR As Long = 10    ' J - Kommentar

' Farger (samme stil som UVALGTE)
Private Const FARGE_HEADER As Long = &HE9D7B9
Private Const FARGE_PANEL_TITLE As Long = &HDDE7FF
Private Const FARGE_OK As Long = &HC6EFCE         ' Lys grønn
Private Const FARGE_OVERLAPP As Long = &HFFC7CE   ' Lys rød
Private Const FARGE_BTN As Long = &HE36C2E
Private Const FARGE_BTN_TXT As Long = &HFFFFFF

' Panel
Private Const PANEL_ROW As Long = 2
Private Const PANEL_H As Single = 24

' Navn
Private Const NM_BTN_REFRESH As String = "btnAktOversRefresh"
Private Const NM_BTN_APPLY As String = "btnAktOversApply"
Private Const NM_BTN_UNDO As String = "btnAktOversUndo"

' Global variabel for undo
Private SisteForsinkelser As Object

' =================== OPPSETT ===================

Public Sub AktivitetsOversikt_Oppsett()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsAO As Worksheet, wsP As Worksheet
    
    On Error Resume Next
    Set wsAO = wb.Worksheets(ARK_OVERSIKT_AKT)
    Set wsP = wb.Worksheets(ARK_PLAN)
    On Error GoTo 0
    
    If wsP Is Nothing Then
        MsgBox "Mangler ark: '" & ARK_PLAN & "'.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' Opprett ark hvis det ikke finnes
    If wsAO Is Nothing Then
        Set wsAO = wb.Worksheets.Add(After:=wsP)
        wsAO.Name = ARK_OVERSIKT_AKT
    End If
    
    ' VIKTIG: Fjern beskyttelse før vi rydder
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    ' Rydd arket
    wsAO.Cells.Clear
    
    ' Sett kolonnebredder
    wsAO.Columns("A").ColumnWidth = 18  ' Person
    wsAO.Columns("B").ColumnWidth = 10  ' Kode
    wsAO.Columns("C").ColumnWidth = 25  ' Beskrivelse
    wsAO.Columns("D").ColumnWidth = 12  ' Opp. start
    wsAO.Columns("E").ColumnWidth = 12  ' Opp. slutt
    wsAO.Columns("F").ColumnWidth = 12  ' Forsinkelse
    wsAO.Columns("G").ColumnWidth = 12  ' Ny slutt
    wsAO.Columns("H").ColumnWidth = 10  ' Varighet
    wsAO.Columns("I").ColumnWidth = 12  ' Status
    wsAO.Columns("J").ColumnWidth = 30  ' Kommentar
    
    ' Toppstripe
    With wsAO.Range("A1:J1")
        .Merge
        .Value = "AKTIVITETSOVERSIKT - Styringsverktøy"
        .Interior.Color = FARGE_PANEL_TITLE
        .Font.Bold = True
        .Font.Size = 16
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
        Boxify .Cells
    End With
    
    ' Panel med knapper og info
    With wsAO
        .Range("A" & PANEL_ROW & ":J" & PANEL_ROW).Interior.Color = RGB(255, 255, 255)
        .Range("A" & PANEL_ROW).Value = "Rediger 'Forsinkelse (dager)' og trykk 'Oppdater Planlegger' for å anvende endringer."
        .Range("A" & PANEL_ROW).Font.Size = 10
        .Range("A" & PANEL_ROW).Font.Italic = True
        .Rows(PANEL_ROW).RowHeight = 20
    End With
    
    ' Rydd eksisterende kontroller
    TryDeleteShape wsAO, NM_BTN_REFRESH
    TryDeleteShape wsAO, NM_BTN_APPLY
    TryDeleteShape wsAO, NM_BTN_UNDO
    
    ' Anker for knapper
    Dim anchRefresh As Range, anchApply As Range, anchUndo As Range
    Set anchRefresh = wsAO.Cells(PANEL_ROW, 7)
    Set anchApply = wsAO.Cells(PANEL_ROW, 9)
    Set anchUndo = wsAO.Cells(PANEL_ROW, 8)
    
    ' Knapp: Oppdater fra Planlegger
    LagKnapp wsAO, NM_BTN_REFRESH, "Hent fra Planlegger", _
             "'" & ThisWorkbook.Name & "'!AktivitetsOversikt_Refresh", _
             anchRefresh, 120, PANEL_H
    
    ' Knapp: Anvend endringer
    LagKnapp wsAO, NM_BTN_APPLY, "Anvend endringer", _
             "'" & ThisWorkbook.Name & "'!AktivitetsOversikt_Apply", _
             anchApply, 120, PANEL_H
    
    ' Knapp: Angre siste endring
    LagKnapp wsAO, NM_BTN_UNDO, "Angre siste", _
             "'" & ThisWorkbook.Name & "'!AktivitetsOversikt_Undo", _
             anchUndo, 100, PANEL_H
    
    ' Tabellhoder
    With wsAO
        .Range("A" & (TBL_START_ROW - 1) & ":J" & (TBL_START_ROW - 1)).Interior.Color = FARGE_HEADER
        .Range("A" & (TBL_START_ROW - 1)).Resize(1, 10).Value = _
            Array("Person", "Kode", "Beskrivelse", "Opp. Start", "Opp. Slutt", _
                  "Forsinkelse (dager)", "Ny Slutt", "Varighet", "Status", "Kommentar")
        
        With .Range("A" & (TBL_START_ROW - 1) & ":J" & (TBL_START_ROW - 1))
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .RowHeight = 30
            Boxify .Cells
        End With
        
        ' Datoformater
        .Range(.Cells(TBL_START_ROW, COL_OPP_START), .Cells(TBL_START_ROW + 500, COL_OPP_START)).NumberFormat = "dd.mm.yyyy"
        .Range(.Cells(TBL_START_ROW, COL_OPP_SLUTT), .Cells(TBL_START_ROW + 500, COL_OPP_SLUTT)).NumberFormat = "dd.mm.yyyy"
        .Range(.Cells(TBL_START_ROW, COL_NY_SLUTT), .Cells(TBL_START_ROW + 500, COL_NY_SLUTT)).NumberFormat = "dd.mm.yyyy"
        
        ' Tallformat for forsinkelse
        .Range(.Cells(TBL_START_ROW, COL_FORSINKET), .Cells(TBL_START_ROW + 500, COL_FORSINKET)).NumberFormat = "0"
        .Range(.Cells(TBL_START_ROW, COL_VARIGHET), .Cells(TBL_START_ROW + 500, COL_VARIGHET)).NumberFormat = "0"
        
        ' Sentrer kolonner
        .Range(.Cells(TBL_START_ROW, COL_KODE), .Cells(TBL_START_ROW + 500, COL_KODE)).HorizontalAlignment = xlCenter
        .Range(.Cells(TBL_START_ROW, COL_STATUS), .Cells(TBL_START_ROW + 500, COL_STATUS)).HorizontalAlignment = xlCenter
        
        ' Formelkolonner (ikke redigerbare - grå bakgrunn)
        .Range(.Cells(TBL_START_ROW, COL_NY_SLUTT), .Cells(TBL_START_ROW + 500, COL_NY_SLUTT)).Interior.Color = RGB(242, 242, 242)
        .Range(.Cells(TBL_START_ROW, COL_VARIGHET), .Cells(TBL_START_ROW + 500, COL_VARIGHET)).Interior.Color = RGB(242, 242, 242)
        .Range(.Cells(TBL_START_ROW, COL_STATUS), .Cells(TBL_START_ROW + 500, COL_STATUS)).Interior.Color = RGB(242, 242, 242)
        
        ' Lås kolonner (tillat kun redigering av Person, Opp.Start, Opp.Slutt, Forsinkelse, Kommentar)
        ' Lås header-rad
        .Range(.Cells(TBL_START_ROW - 1, 1), .Cells(TBL_START_ROW - 1, 10)).Locked = True
        
        ' Lås Kode og Beskrivelse (auto-generert)
        .Range(.Cells(TBL_START_ROW, COL_KODE), .Cells(TBL_START_ROW + 500, COL_KODE)).Locked = True
        .Range(.Cells(TBL_START_ROW, COL_BESKR), .Cells(TBL_START_ROW + 500, COL_BESKR)).Locked = True
        
        ' Lås formelkolonner
        .Range(.Cells(TBL_START_ROW, COL_NY_SLUTT), .Cells(TBL_START_ROW + 500, COL_NY_SLUTT)).Locked = True
        .Range(.Cells(TBL_START_ROW, COL_VARIGHET), .Cells(TBL_START_ROW + 500, COL_VARIGHET)).Locked = True
        .Range(.Cells(TBL_START_ROW, COL_STATUS), .Cells(TBL_START_ROW + 500, COL_STATUS)).Locked = True
        
        ' Tillat redigering av Person, Opp.Start, Opp.Slutt, Forsinkelse, Kommentar
        .Range(.Cells(TBL_START_ROW, COL_PERSON), .Cells(TBL_START_ROW + 500, COL_PERSON)).Locked = False
        .Range(.Cells(TBL_START_ROW, COL_OPP_START), .Cells(TBL_START_ROW + 500, COL_OPP_START)).Locked = False
        .Range(.Cells(TBL_START_ROW, COL_OPP_SLUTT), .Cells(TBL_START_ROW + 500, COL_OPP_SLUTT)).Locked = False
        .Range(.Cells(TBL_START_ROW, COL_FORSINKET), .Cells(TBL_START_ROW + 500, COL_FORSINKET)).Locked = False
        .Range(.Cells(TBL_START_ROW, COL_KOMMENTAR), .Cells(TBL_START_ROW + 500, COL_KOMMENTAR)).Locked = False
    End With
    
    ' *** Sett begrenset AutoFilter ***
    Call SettBegrensetAutoFilter(wsAO)
    
    ' *** Legg til dropdown på Person-kolonnen ***
    Call SettPersonDropdown(wsAO)
    
    ' *** Aktiver arkbeskyttelse ETTER filter og dropdown er satt ***
    wsAO.Protect Password:="", _
                 DrawingObjects:=False, _
                 Contents:=True, _
                 Scenarios:=False, _
                 AllowFormattingCells:=True, _
                 AllowFormattingColumns:=False, _
                 AllowFormattingRows:=False, _
                 AllowInsertingColumns:=False, _
                 AllowInsertingRows:=False, _
                 AllowDeletingColumns:=False, _
                 AllowDeletingRows:=False, _
                 AllowSorting:=True, _
                 AllowFiltering:=True, _
                 AllowUsingPivotTables:=False
    
    ' Frys rader og kolonner
    wsAO.Activate
    wsAO.Range("A" & TBL_START_ROW).Select
    ActiveWindow.FreezePanes = True
    
    ' Last inn data fra Planlegger
    AktivitetsOversikt_Refresh
    
    Application.ScreenUpdating = True
    
    MsgBox "AKTIVITETSOVERSIKT er klar!" & vbCrLf & vbCrLf & _
           " Rediger 'Forsinkelse (dager)' for å justere aktiviteter" & vbCrLf & _
           " Trykk 'Oppdater Planlegger' for å anvende endringer" & vbCrLf & _
           " Status-kolonnen viser automatisk om det er overlapp", vbInformation
End Sub

' =================== OPPDATERING ===================

Public Sub AktivitetsOversikt_Refresh()
    Dim wsAO As Worksheet, wsP As Worksheet, wsTyp As Worksheet
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastCol As Long, lastRow As Long
    Dim personRad As Long, r As Long, c As Long
    Dim personNavn As String
    Dim aktiviteter As Object  ' Dictionary med alle aktiviteter
    Dim aktKey As String, currentRow As Long
    Dim aktivitet As Object
    
    ' Hent ark
    On Error Resume Next
    Set wsAO = ThisWorkbook.Worksheets(ARK_OVERSIKT_AKT)
    Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT_TYP)
    On Error GoTo 0
    
    If wsAO Is Nothing Or wsP Is Nothing Or wsTyp Is Nothing Then
        MsgBox "Mangler nødvendige ark.", vbCritical
        Exit Sub
    End If
    
    ' VIKTIG: Fjern beskyttelse før refresh
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    ' Hent dynamiske verdier
    førsteDatoKol = HentFørsteDatoKol()
    datoRad = HentDatoRad()
    førstePersonRad = HentFørstePersonRad()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Rydd eksisterende data (behold header)
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    If lastRow >= TBL_START_ROW Then
        wsAO.Range(wsAO.Cells(TBL_START_ROW, 1), wsAO.Cells(lastRow, 10)).ClearContents
    End If
    
    ' Opprett dictionary for aktiviteter
    Set aktiviteter = CreateObject("Scripting.Dictionary")
    
    ' Finn siste kolonne i Planlegger
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    
    ' Gå gjennom alle personer i Planlegger
    For personRad = førstePersonRad To lastRow
        personNavn = Trim$(wsP.Cells(personRad, 1).Value)
        
        ' Sjekk om dette er en personrad (ikke tom og ikke preview)
        If Len(personNavn) > 0 And Not (UCase$(personNavn) Like "UVALGTE*") Then
            ' Skann gjennom alle datocolonner for denne personen (og under-rader)
            Call SkannPersonAktiviteter(wsP, wsTyp, personRad, personNavn, _
                                        førsteDatoKol, lastCol, datoRad, aktiviteter)
        End If
    Next personRad
    
    ' Fyll tabellen fra dictionary (sortert)
    currentRow = TBL_START_ROW
    Call FyllTabellFraAktiviteter(wsAO, aktiviteter, currentRow)
    
    ' Legg til formler for beregnede kolonner
    Call LeggTilFormler(wsAO, TBL_START_ROW, currentRow - 1)
    
    ' Detekter overlapp
    Call DetekterOverlapp(wsAO, TBL_START_ROW, currentRow - 1)
    
    ' Formater tabellen
    Call FormaterTabell(wsAO, TBL_START_ROW, currentRow - 1)
    
    ' *** Gjenopprett AutoFilter etter refresh ***
    If currentRow > TBL_START_ROW Then
        Call SettBegrensetAutoFilter(wsAO)
        Call SettPersonDropdown(wsAO)
    End If
    
    ' VIKTIG: Beskytt arket igjen etter refresh
    On Error Resume Next
    wsAO.Protect Password:="", _
                 DrawingObjects:=False, _
                 Contents:=True, _
                 Scenarios:=False, _
                 AllowFormattingCells:=True, _
                 AllowFormattingColumns:=False, _
                 AllowFormattingRows:=False, _
                 AllowInsertingColumns:=False, _
                 AllowInsertingRows:=False, _
                 AllowDeletingColumns:=False, _
                 AllowDeletingRows:=False, _
                 AllowSorting:=True, _
                 AllowFiltering:=True, _
                 AllowUsingPivotTables:=False
    On Error GoTo 0
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Oversikten er oppdatert med " & (currentRow - TBL_START_ROW) & " aktiviteter.", vbInformation
End Sub

' Skann en personblokk (hovedrad + under-rader) for aktiviteter
Private Sub SkannPersonAktiviteter(wsP As Worksheet, wsTyp As Worksheet, _
                                   personRad As Long, personNavn As String, _
                                   førsteDatoKol As Long, lastCol As Long, datoRad As Long, _
                                   aktiviteter As Object)
    Dim r As Long, c As Long, startCol As Long, endCol As Long
    Dim celVal As String, aktivKode As String, aktivBeskr As String, aktivFarge As Long
    Dim startDato As Date, sluttDato As Date
    Dim kommentar As String
    Dim aktKey As String
    Dim blockEnd As Long
    
    ' Finn slutten av personblokken
    blockEnd = personRad
    Do While blockEnd < wsP.Rows.Count
        If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
        blockEnd = blockEnd + 1
    Loop
    
    ' Skann alle rader i blokken (hovedrad + under-rader)
    For r = personRad To blockEnd
        c = førsteDatoKol
        
        Do While c <= lastCol
            celVal = Trim$(wsP.Cells(r, c).Value)
            
            ' Sjekk om dette er start på en aktivitet (fet tekst)
            If Len(celVal) > 0 And wsP.Cells(r, c).Font.Bold Then
                ' Ekstraher aktivitetskode (første ord før "-")
                aktivKode = ExtractAktivitetsKode(celVal)
                kommentar = ExtractKommentar(celVal)
                
                ' Finn start og slutt av aktivitetsblokken
                startCol = c
                endCol = c
                
                ' Finn slutten av blokken (sammenhengende celler med samme farge)
                Dim blokFarge As Long
                blokFarge = wsP.Cells(r, c).Interior.Color
                
                Do While endCol < lastCol
                    Dim NesteCelle As Range
                    Set NesteCelle = wsP.Cells(r, endCol + 1)
                    
                    ' Stopp ved hvit celle, annen farge, eller ny aktivitet (fet tekst)
                    If NesteCelle.Interior.Color = RGB(255, 255, 255) Or _
                       NesteCelle.Interior.ColorIndex = xlColorIndexNone Then
                        Exit Do
                    ElseIf NesteCelle.Font.Bold And Len(Trim$(NesteCelle.Value)) > 0 Then
                        Exit Do
                    ElseIf NesteCelle.Interior.Color = blokFarge And _
                           Len(Trim$(NesteCelle.Value)) = 0 Then
                        endCol = endCol + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                ' Hent datoer fra kolonnehoder
                Dim datoOK As Boolean
                datoOK = False
                
                On Error Resume Next
                If IsDate(wsP.Cells(datoRad, startCol).Value) Then
                    startDato = wsP.Cells(datoRad, startCol).Value
                    If IsDate(wsP.Cells(datoRad, endCol).Value) Then
                        sluttDato = wsP.Cells(datoRad, endCol).Value
                        datoOK = True
                    End If
                End If
                On Error GoTo 0
                
                ' Hopp over hvis datoer er ugyldige
                If Not datoOK Then
                    c = endCol + 1
                    GoTo NesteCelle
                End If
                
                ' Slå opp aktivitetsbeskrivelse
                If Not LookupAktivitet(wsTyp, aktivKode, aktivBeskr, aktivFarge) Then
                    aktivBeskr = ""
                End If
                
                ' Lag unik nøkkel (person + kode + startdato)
                aktKey = personNavn & "|" & aktivKode & "|" & Format(startDato, "yyyy-mm-dd")
                
                ' Legg til i dictionary hvis ikke allerede finnes
                If Not aktiviteter.exists(aktKey) Then
                    Dim aktInfo As Object
                    Set aktInfo = CreateObject("Scripting.Dictionary")
                    aktInfo("Person") = personNavn
                    aktInfo("Kode") = aktivKode
                    aktInfo("Beskrivelse") = aktivBeskr
                    aktInfo("StartDato") = startDato
                    aktInfo("SluttDato") = sluttDato
                    aktInfo("Kommentar") = kommentar
                    aktInfo("Forsinkelse") = 0  ' Default ingen forsinkelse
                    
                    aktiviteter.Add aktKey, aktInfo
                End If
                
                ' Hopp over resten av denne blokken
                c = endCol + 1
NesteCelle:
            Else
                c = c + 1
            End If
        Loop
    Next r
End Sub

' Ekstraher aktivitetskode fra celle-tekst (før "-")
Private Function ExtractAktivitetsKode(txt As String) As String
    Dim pos As Long
    pos = InStr(txt, "-")
    If pos > 0 Then
        ExtractAktivitetsKode = Trim$(Left$(txt, pos - 1))
    Else
        ' Hvis ingen "-", ta første ord
        pos = InStr(txt, " ")
        If pos > 0 Then
            ExtractAktivitetsKode = Trim$(Left$(txt, pos - 1))
        Else
            ExtractAktivitetsKode = Trim$(txt)
        End If
    End If
End Function

' Ekstraher kommentar fra celle-tekst (etter "-")
Private Function ExtractKommentar(txt As String) As String
    Dim pos As Long
    pos = InStr(txt, "-")
    If pos > 0 Then
        ExtractKommentar = Trim$(Mid$(txt, pos + 1))
    Else
        ExtractKommentar = ""
    End If
End Function

' Fyll tabellen fra aktiviteter-dictionary (sortert)
Private Sub FyllTabellFraAktiviteter(wsAO As Worksheet, aktiviteter As Object, ByRef currentRow As Long)
    Dim sortedKeys() As String
    Dim i As Long, j As Long, k As Variant
    Dim tempKey As String
    Dim aktInfo As Object
    Dim wsTyp As Worksheet
    Dim kodeFarge As Long, beskr As String
    
    ' Hent aktivitetstypearket for farger
    On Error Resume Next
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT_TYP)
    On Error GoTo 0
    
    ' Kopier nøkler til array for sortering
    ReDim sortedKeys(0 To aktiviteter.Count - 1)
    i = 0
    For Each k In aktiviteter.Keys
        sortedKeys(i) = CStr(k)
        i = i + 1
    Next k
    
    ' Sorter etter startdato (enkel bubble sort)
    For i = 0 To UBound(sortedKeys) - 1
        For j = i + 1 To UBound(sortedKeys)
            Set aktInfo = aktiviteter(sortedKeys(i))
            Dim aktInfo2 As Object
            Set aktInfo2 = aktiviteter(sortedKeys(j))
            
            ' Sorter først etter dato, deretter etter kode
            If aktInfo("StartDato") > aktInfo2("StartDato") Or _
               (aktInfo("StartDato") = aktInfo2("StartDato") And _
                aktInfo("Kode") > aktInfo2("Kode")) Then
                tempKey = sortedKeys(i)
                sortedKeys(i) = sortedKeys(j)
                sortedKeys(j) = tempKey
            End If
        Next j
    Next i
    
    ' Fyll tabellen
    For i = 0 To UBound(sortedKeys)
        Set aktInfo = aktiviteter(sortedKeys(i))
        
        With wsAO
            .Cells(currentRow, COL_PERSON).Value = aktInfo("Person")
            .Cells(currentRow, COL_KODE).Value = aktInfo("Kode")
            .Cells(currentRow, COL_BESKR).Value = aktInfo("Beskrivelse")
            .Cells(currentRow, COL_OPP_START).Value = aktInfo("StartDato")
            .Cells(currentRow, COL_OPP_SLUTT).Value = aktInfo("SluttDato")
            .Cells(currentRow, COL_FORSINKET).Value = aktInfo("Forsinkelse")
            .Cells(currentRow, COL_KOMMENTAR).Value = aktInfo("Kommentar")
            
            ' Legg til kodefarge hvis vi har tilgang til oversiktsarket
            If Not wsTyp Is Nothing Then
                If LookupAktivitet(wsTyp, aktInfo("Kode"), beskr, kodeFarge) Then
                    With .Cells(currentRow, COL_KODE)
                        .Interior.Pattern = xlSolid
                        .Interior.Color = kodeFarge
                        .Font.Bold = True
                        .Font.Color = IIf(ErLysFarge(kodeFarge), RGB(0, 0, 0), RGB(255, 255, 255))
                    End With
                End If
            End If
        End With
        
        currentRow = currentRow + 1
    Next i
End Sub

' Sjekk om farge er lys (for å velge tekst-farge)
Private Function ErLysFarge(col As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = col Mod 256: g = (col \ 256) Mod 256: b = (col \ 65536) Mod 256
    ErLysFarge = (0.299 * r + 0.587 * g + 0.114 * b) > 160
End Function

' Legg til formler for beregnede kolonner
Private Sub LeggTilFormler(wsAO As Worksheet, startRow As Long, endRow As Long)
    Dim r As Long
    
    For r = startRow To endRow
        ' Ny slutt = Opp. Slutt + Forsinkelse
        wsAO.Cells(r, COL_NY_SLUTT).Formula = "=E" & r & "+F" & r
        
        ' Varighet = Slutt - Start + 1
        wsAO.Cells(r, COL_VARIGHET).Formula = "=E" & r & "-D" & r & "+1"
    Next r
End Sub

' Detekter overlapp (enkel versjon - kompletteres senere)
Private Sub DetekterOverlapp(wsAO As Worksheet, startRow As Long, endRow As Long)
    Dim r As Long, r2 As Long
    Dim person1 As String, person2 As String
    Dim start1 As Date, slutt1 As Date, nySlut1 As Date
    Dim start2 As Date, slutt2 As Date, nySlut2 As Date
    Dim harOverlapp As Boolean
    
    ' Gå gjennom alle rader
    For r = startRow To endRow
        person1 = wsAO.Cells(r, COL_PERSON).Value
        start1 = wsAO.Cells(r, COL_OPP_START).Value
        slutt1 = wsAO.Cells(r, COL_OPP_SLUTT).Value
        nySlut1 = wsAO.Cells(r, COL_NY_SLUTT).Value
        
        harOverlapp = False
        
        ' Sammenlign med alle andre aktiviteter for samme person
        For r2 = startRow To endRow
            If r2 <> r Then
                person2 = wsAO.Cells(r2, COL_PERSON).Value
                
                ' Bare sjekk samme person
                If person2 = person1 Then
                    start2 = wsAO.Cells(r2, COL_OPP_START).Value
                    slutt2 = wsAO.Cells(r2, COL_OPP_SLUTT).Value
                    nySlut2 = wsAO.Cells(r2, COL_NY_SLUTT).Value
                    
                    ' Sjekk overlapp: aktivitet 1 (med forsinkelse) overlapper aktivitet 2
                    ' Overlapp hvis: start1 <= nySlut2 OG nySlut1 >= start2
                    If start1 <= nySlut2 And nySlut1 >= start2 Then
                        harOverlapp = True
                        Exit For
                    End If
                End If
            End If
        Next r2
        
        ' Sett status MED konsistent fargelegging
        With wsAO.Cells(r, COL_STATUS)
            If harOverlapp Then
                .Value = "OVERLAPP"
                .Interior.Pattern = xlSolid
                .Interior.Color = FARGE_OVERLAPP
                .Font.Bold = True
                .Font.Color = RGB(0, 0, 0)
            Else
                .Value = "OK"
                .Interior.Pattern = xlSolid
                .Interior.Color = FARGE_OK
                .Font.Bold = False
                .Font.Color = RGB(0, 0, 0)
            End If
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next r
End Sub

' Sett dropdown-validering på Person-kolonnen
Private Sub SettPersonDropdown(wsAO As Worksheet)
    Dim wsP As Worksheet
    Dim lastRowP As Long, lastRowAO As Long
    Dim r As Long, w As Long
    Dim personListe As String
    Dim førstePersonRad As Long
    
    ' VIKTIG: Fjern beskyttelse først
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    On Error GoTo 0
    If wsP Is Nothing Then Exit Sub
    
    ' Hent førstePersonRad fra Named Range
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    
    ' Bygg person-liste fra Planlegger
    lastRowP = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    personListe = ""
    
    For r = førstePersonRad To lastRowP
        If Len(Trim$(wsP.Cells(r, 1).Value)) > 0 Then
            If Len(personListe) > 0 Then
                personListe = personListe & "," & Trim$(wsP.Cells(r, 1).Value)
            Else
                personListe = Trim$(wsP.Cells(r, 1).Value)
            End If
        End If
    Next r
    
    If Len(personListe) = 0 Then Exit Sub
    
    ' Finn siste rad i AKTIVITETSOVERSIKT
    lastRowAO = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    If lastRowAO < TBL_START_ROW Then lastRowAO = TBL_START_ROW + 100
    
    ' Legg til dropdown-validering på Person-kolonnen (kun data-rader, ikke header)
    With wsAO.Range(wsAO.Cells(TBL_START_ROW, COL_PERSON), wsAO.Cells(lastRowAO, COL_PERSON)).Validation
        .Delete  ' Fjern eksisterende validering
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Operator:=xlBetween, _
             Formula1:=personListe
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
        .ErrorTitle = "Ugyldig person"
        .ErrorMessage = "Velg en person fra listen."
    End With
End Sub

' Sett AutoFilter med begrensede dropdown-piler
Private Sub SettBegrensetAutoFilter(wsAO As Worksheet)
    ' VIKTIG: Fjern beskyttelse først
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    ' Fjern eksisterende filter
    If wsAO.AutoFilterMode Then wsAO.AutoFilterMode = False
    
    ' Sett AutoFilter på header-raden
    wsAO.Range(wsAO.Cells(TBL_START_ROW - 1, 1), wsAO.Cells(TBL_START_ROW - 1, 10)).AutoFilter
    
    ' Fjern dropdown-piler fra kolonner vi ikke vil filtrere
    ' Behold kun: Person (1), Kode (2), Opp.Start (4), Opp.Slutt (5), Status (9)
    With wsAO
        .AutoFilter.Range.AutoFilter Field:=3, VisibleDropDown:=False  ' Beskrivelse
        .AutoFilter.Range.AutoFilter Field:=6, VisibleDropDown:=False  ' Forsinkelse
        .AutoFilter.Range.AutoFilter Field:=7, VisibleDropDown:=False  ' Ny Slutt
        .AutoFilter.Range.AutoFilter Field:=8, VisibleDropDown:=False  ' Varighet
        .AutoFilter.Range.AutoFilter Field:=10, VisibleDropDown:=False ' Kommentar
    End With
End Sub

' Formater tabellen
Private Sub FormaterTabell(wsAO As Worksheet, startRow As Long, endRow As Long)
    Dim r As Long, c As Long
    
    If endRow < startRow Then Exit Sub
    
    ' Legg til grid på hele tabellen
    With wsAO.Range(wsAO.Cells(startRow, 1), wsAO.Cells(endRow, 10)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlColorIndexAutomatic
    End With
    
    ' Altererende radfarger for lesbarhet (UNNTATT Status og Kode kolonner)
    For r = startRow To endRow
        If r Mod 2 = 0 Then
            ' Altererende rad - lys grå bakgrunn
            For c = 1 To 10
                ' HOPP OVER Status-kolonnen (9) og Kode-kolonnen (2) - de har egne farger
                If c <> COL_STATUS And c <> COL_KODE Then
                    wsAO.Cells(r, c).Interior.Color = RGB(250, 250, 250)
                End If
            Next c
        Else
            ' Hvit bakgrunn på oddetallsrader (unntatt Status og Kode)
            For c = 1 To 10
                If c <> COL_STATUS And c <> COL_KODE Then
                    ' Sjekk om dette er formelkolonne (grå)
                    If c = COL_NY_SLUTT Or c = COL_VARIGHET Then
                        wsAO.Cells(r, c).Interior.Color = RGB(242, 242, 242)
                    Else
                        wsAO.Cells(r, c).Interior.Color = RGB(255, 255, 255)
                    End If
                End If
            Next c
        End If
    Next r
    
    ' Sikre at formelkolonner alltid har grå bakgrunn (på de som ikke er Status)
    wsAO.Range(wsAO.Cells(startRow, COL_NY_SLUTT), wsAO.Cells(endRow, COL_NY_SLUTT)).Interior.Color = RGB(242, 242, 242)
    wsAO.Range(wsAO.Cells(startRow, COL_VARIGHET), wsAO.Cells(endRow, COL_VARIGHET)).Interior.Color = RGB(242, 242, 242)
    ' Status-kolonnen får sin egen farge fra DetekterOverlapp - ikke rør den her
    ' Kode-kolonnen får sin egen farge fra FyllTabellFraAktiviteter - ikke rør den her
End Sub

' Lookup aktivitet (kopierer fra andre moduler)
Private Function LookupAktivitet(wsTyp As Worksheet, ByVal kode As String, _
                                 ByRef beskrivelse As String, ByRef farge As Long) As Boolean
    Dim r As Long, lastRow As Long
    lastRow = wsTyp.Cells(wsTyp.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If UCase$(Trim$(wsTyp.Cells(r, 1).Value)) = UCase$(Trim$(kode)) Then
            beskrivelse = CStr(wsTyp.Cells(r, 2).Value)
            farge = wsTyp.Cells(r, 1).Interior.Color
            LookupAktivitet = True
            Exit Function
        End If
    Next r
End Function

Public Sub AktivitetsOversikt_Apply()
    MsgBox "Apply-funksjonen starter nå...", vbInformation, "DEBUG"
    
    On Error GoTo ErrorHandler
    
    Dim wsAO As Worksheet, wsP As Worksheet, wsTyp As Worksheet
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim lastRow As Long, r As Long
    Dim person As String, kode As String, beskrivelse As String, farge As Long
    Dim oppStart As Date, oppSlutt As Date, forsinkelse As Long, nySlutDato As Date
    Dim kommentar As String, visTekst As String
    Dim personRow As Long, startCol As Long, sluttCol As Long, nySluttCol As Long
    Dim farger As Object
    Dim antallOppdatert As Long, antallOverlapp As Long
    Dim overlappListe As Object
    
    ' Hent ark
    On Error Resume Next
    Set wsAO = ThisWorkbook.Worksheets(ARK_OVERSIKT_AKT)
    Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT_TYP)
    On Error GoTo 0
    
    If wsAO Is Nothing Then
        MsgBox "Finner ikke arket '" & ARK_OVERSIKT_AKT & "'.", vbCritical
        Exit Sub
    End If
    If wsP Is Nothing Then
        MsgBox "Finner ikke arket '" & ARK_PLAN & "'.", vbCritical
        Exit Sub
    End If
    If wsTyp Is Nothing Then
        MsgBox "Finner ikke arket '" & ARK_OVERSIKT_TYP & "'.", vbCritical
        Exit Sub
    End If
    
    ' Hent dynamiske verdier
    førsteDatoKol = HentFørsteDatoKol()
    datoRad = HentDatoRad()
    førstePersonRad = HentFørstePersonRad()
    
    ' Bekreft med bruker
    If MsgBox("Dette vil oppdatere Planlegger med forsinkelser fra oversikten." & vbCrLf & vbCrLf & _
              "Vil du fortsette?", vbYesNo + vbQuestion, "Bekreft oppdatering") <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' KRITISK: Lagre tilstand for UNDO (hele Planlegger-tilstanden)
    Set SisteForsinkelser = LagrePlanleggerTilstand(wsP, førsteDatoKol, datoRad, førstePersonRad)
    
    ' Lagre også forsinkelser for å gjenopprette etter refresh
    Dim forsinkelser As Object
    Set forsinkelser = LagreForsinkelser(wsAO)
    
    ' Hent aktivitetsfarger
    Set farger = HentAktivitetsFarger(wsTyp)
    
    ' Opprett dictionary for overlapp-tracking
    Set overlappListe = CreateObject("Scripting.Dictionary")
    
    ' Finn siste rad i oversiktstabellen
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    If lastRow < TBL_START_ROW Then
        MsgBox "Ingen aktiviteter å oppdatere.", vbInformation
        GoTo Cleanup
    End If
    
    antallOppdatert = 0
    
    ' Gå gjennom alle aktiviteter i oversikten
    For r = TBL_START_ROW To lastRow
        person = Trim$(wsAO.Cells(r, COL_PERSON).Value)
        kode = Trim$(wsAO.Cells(r, COL_KODE).Value)
        oppStart = wsAO.Cells(r, COL_OPP_START).Value
        oppSlutt = wsAO.Cells(r, COL_OPP_SLUTT).Value
        forsinkelse = wsAO.Cells(r, COL_FORSINKET).Value
        nySlutDato = wsAO.Cells(r, COL_NY_SLUTT).Value
        kommentar = Trim$(wsAO.Cells(r, COL_KOMMENTAR).Value)
        
        ' Bare oppdater hvis det er forsinkelse
        If forsinkelse > 0 Then
            ' Finn person i Planlegger
            personRow = FinnPersonRadIPlanlegger(wsP, person, førstePersonRad)
            If personRow > 0 Then
                ' Finn aktiviteten i Planlegger
                startCol = FinnDatoKolonneIPlanlegger(wsP, oppStart, datoRad, førsteDatoKol)
                sluttCol = FinnDatoKolonneIPlanlegger(wsP, oppSlutt, datoRad, førsteDatoKol)
                nySluttCol = FinnDatoKolonneIPlanlegger(wsP, nySlutDato, datoRad, førsteDatoKol)
                
                If startCol > 0 And sluttCol > 0 And nySluttCol > 0 Then
                    ' Oppdater aktiviteten i Planlegger
                    If OppdaterAktivitetIPlanlegger(wsP, wsTyp, personRow, kode, _
                                                    startCol, sluttCol, nySluttCol, _
                                                    kommentar, farger, førsteDatoKol, datoRad) Then
                        antallOppdatert = antallOppdatert + 1
                    End If
                End If
            End If
        End If
        
        ' Samle overlapp-info
        If wsAO.Cells(r, COL_STATUS).Value = "OVERLAPP" Then
            Dim overlappKey As String
            overlappKey = person & "|" & kode & "|" & Format(oppStart, "yyyy-mm-dd")
            If Not overlappListe.exists(overlappKey) Then
                overlappListe.Add overlappKey, Array(person, oppStart, nySlutDato)
                antallOverlapp = antallOverlapp + 1
            End If
        End If
    Next r
    
    ' Tegn skravering for overlapp
    If antallOverlapp > 0 Then
        Call TegnOverlappSkravering(wsP, overlappListe, førsteDatoKol, datoRad, førstePersonRad)
    End If
    
    ' VIKTIG: Lagre ALLE redigerbare kolonner før refresh (Person, Forsinkelse, Kommentar)
    Dim alleEndringer As Object
    Set alleEndringer = LagreAlleRedigerbareFelt(wsAO)
    
    ' Oppdater oversikten for å reflektere endringer fra Planlegger
    Call AktivitetsOversikt_Refresh
    
    ' KRITISK: Gjenopprett ALLE endringer etter refresh
    Call GjenopprettAlleRedigerbareFelt(wsAO, alleEndringer)
    
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If antallOppdatert > 0 Then
        MsgBox "Oppdatert " & antallOppdatert & " aktivitet(er) i Planlegger." & vbCrLf & _
               IIf(antallOverlapp > 0, "Fant " & antallOverlapp & " overlapp - markert med skravering.", ""), _
               vbInformation
    Else
        MsgBox "Ingen aktiviteter med forsinkelse funnet.", vbInformation
    End If
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Feil i AktivitetsOversikt_Apply:" & vbCrLf & vbCrLf & _
           "Feilnummer: " & Err.Number & vbCrLf & _
           "Beskrivelse: " & Err.Description & vbCrLf & vbCrLf & _
           "Rad i kode: " & Erl, vbCritical
End Sub

' Oppdater én aktivitet i Planlegger (utvid blokken)
Private Function OppdaterAktivitetIPlanlegger(wsP As Worksheet, wsTyp As Worksheet, _
                                              personRow As Long, kode As String, _
                                              startCol As Long, gammeltSluttCol As Long, nyttSluttCol As Long, _
                                              kommentar As String, farger As Object, _
                                              førsteDatoKol As Long, datoRad As Long) As Boolean
    Dim r As Long, blockEnd As Long, målRad As Long
    Dim c As Long, cel As Range
    Dim funnet As Boolean
    Dim beskrivelse As String, farge As Long, visTekst As String
    Dim overlappStartCol As Long
    Dim overlappAktivitetRad As Long, overlappAktivitetStartCol As Long, overlappAktivitetSluttCol As Long
    Dim overlappFarge As Long
    
    ' Finn personblokken
    blockEnd = personRow
    Do While blockEnd < wsP.Rows.Count
        If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
        blockEnd = blockEnd + 1
    Loop
    
    ' Finn raden med denne aktiviteten (den som skal forsinkes)
    funnet = False
    For r = personRow To blockEnd
        Set cel = wsP.Cells(r, startCol)
        If Len(Trim$(cel.Value)) > 0 And cel.Font.Bold Then
            If InStr(1, cel.Value, kode, vbTextCompare) > 0 Then
                målRad = r
                funnet = True
                Exit For
            End If
        End If
    Next r
    
    If Not funnet Then Exit Function
    
    ' Slå opp aktivitetsinfo
    If Not LookupAktivitet(wsTyp, kode, beskrivelse, farge) Then
        Exit Function
    End If
    
    ' Utvid blokken til ny sluttdato
    If nyttSluttCol > gammeltSluttCol Then
        ' Finn overlappende aktivitet i utvidelsesområdet
        overlappAktivitetRad = 0
        overlappStartCol = 0
        
        For c = gammeltSluttCol + 1 To nyttSluttCol
            Set cel = wsP.Cells(målRad, c)
            If Len(Trim$(cel.Value)) > 0 And cel.Font.Bold Then
                ' Funnet overlappende aktivitet!
                overlappAktivitetRad = målRad
                overlappStartCol = c
                overlappFarge = cel.Interior.Color
                
                ' Finn hvor denne aktiviteten starter
                Dim tempC As Long
                For tempC = c To førsteDatoKol Step -1
                    If wsP.Cells(målRad, tempC).Interior.Color <> overlappFarge Or _
                       (Len(Trim$(wsP.Cells(målRad, tempC).Value)) > 0 And tempC < c) Then
                        overlappAktivitetStartCol = tempC + 1
                        Exit For
                    End If
                    If tempC = førsteDatoKol Then overlappAktivitetStartCol = førsteDatoKol
                Next tempC
                
                ' Finn hvor denne aktiviteten slutter
                Dim lastCol As Long
                lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
                For tempC = c To lastCol
                    If wsP.Cells(målRad, tempC).Interior.Color <> overlappFarge Then
                        overlappAktivitetSluttCol = tempC - 1
                        Exit For
                    End If
                    If tempC = lastCol Then overlappAktivitetSluttCol = lastCol
                Next tempC
                
                Exit For
            End If
        Next c
        
        If overlappAktivitetRad > 0 Then
            ' OVERLAPP FUNNET - flytt den ANDRE aktiviteten (den som kommer senere) til ny rad
            
            ' 1. Finn ledig rad for den overlappende aktiviteten
            Dim nyRad As Long
            nyRad = FinnEllerOpprettLedigRadU5(wsP, personRow, overlappAktivitetStartCol, overlappAktivitetSluttCol, farger, førsteDatoKol, datoRad)
            
            If nyRad > 0 Then
                ' 2. Flytt den overlappende aktiviteten til ny rad
                Call FlyttHeleAktivitetTilNyRad(wsP, målRad, nyRad, overlappAktivitetStartCol, overlappAktivitetSluttCol)
                
                ' 3. Utvid den forsinkede aktiviteten MED SMART TEKST-SENTRERING
                Dim overlappSluttCol As Long
                overlappSluttCol = Application.WorksheetFunction.Min(nyttSluttCol, overlappAktivitetSluttCol)
                
                visTekst = kode & IIf(Len(kommentar) > 0, " - " & kommentar, IIf(Len(beskrivelse) > 0, " - " & beskrivelse, ""))
                
                ' Utvid blokken, men teksten sentreres kun til overlapp-start
                Call ApplyBlockFormattingMedOverlapp(wsP, målRad, startCol, gammeltSluttCol, overlappStartCol, nyttSluttCol, farge, visTekst)
                
                ' 4. Legg RØD skravering i overlapp-området
                Call LeggSkraveringIOverlapp(wsP, målRad, overlappStartCol, overlappSluttCol)
                
                OppdaterAktivitetIPlanlegger = True
            End If
        Else
            ' Ingen overlapp - bare utvid på samme rad
            visTekst = kode & IIf(Len(kommentar) > 0, " - " & kommentar, IIf(Len(beskrivelse) > 0, " - " & beskrivelse, ""))
            Call ApplyBlockFormattingExtend(wsP, målRad, startCol, nyttSluttCol, farge, visTekst)
            OppdaterAktivitetIPlanlegger = True
        End If
    End If
End Function

' Utvid blokk med smart tekst-sentrering (unngår skravert område)
Private Sub ApplyBlockFormattingMedOverlapp(wsP As Worksheet, målRad As Long, _
                                            startCol As Long, gammeltSluttCol As Long, _
                                            overlappStartCol As Long, sluttCol As Long, _
                                            farge As Long, visTekst As String)
    Dim c As Long, cel As Range
    Dim heleBlokken As Range, tekstDel As Range
    
    ' STEG 1: Rydd alt først (blank slate)
    Set heleBlokken = wsP.Range(wsP.Cells(målRad, startCol), wsP.Cells(målRad, sluttCol))
    heleBlokken.ClearFormats
    heleBlokken.ClearContents
    
    ' STEG 2: Sett bakgrunnsfarge på ALLE celler individuelt
    For c = startCol To sluttCol
        With wsP.Cells(målRad, c).Interior
            .Pattern = xlSolid
            .Color = farge
        End With
    Next c
    
    ' STEG 3: Sett kraftige ytterkanter
    With heleBlokken
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = RGB(0, 0, 0)
        End With
    End With
    
    ' STEG 4: Legg grid-linjer på HVER celle
    For c = startCol To sluttCol - 1
        With wsP.Cells(målRad, c).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(180, 180, 180)
        End With
    Next c
    
    ' STEG 5: Sett font-formatering på alle celler
    With heleBlokken.Font
        .Bold = True
        .Color = IIf(ErLysFarge(farge), RGB(0, 0, 0), RGB(255, 255, 255))
    End With
    heleBlokken.VerticalAlignment = xlCenter
    
    ' STEG 6: Sentrer tekst-delen (ikke-skravert område) - UTEN Å MERGE
    Set tekstDel = wsP.Range(wsP.Cells(målRad, startCol), wsP.Cells(målRad, overlappStartCol - 1))
    
    ' Bruk CenterAcrossSelection i stedet for Merge
    tekstDel.HorizontalAlignment = xlCenterAcrossSelection
    
    ' Sett teksten
    wsP.Cells(målRad, startCol).Value = visTekst
    wsP.Cells(målRad, startCol).WrapText = True
    
    ' STEG 7: Gjenopprett bunnlinje
    If målRad < wsP.Rows.Count Then
        With wsP.Range(wsP.Cells(målRad + 1, startCol), wsP.Cells(målRad + 1, sluttCol)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If
End Sub

' Legg rød skravering i overlapp-området (ETTER at alt annet er satt opp)
Private Sub LeggSkraveringIOverlapp(wsP As Worksheet, rad As Long, startCol As Long, sluttCol As Long)
    Dim c As Long
    
    For c = startCol To sluttCol
        With wsP.Cells(rad, c).Interior
            .Pattern = xlPatternLightDown
            .PatternColor = RGB(255, 0, 0)
            ' Color forblir som den var
        End With
        
        ' Gjenopprett grid (må gjøres etter pattern)
        If c < sluttCol Then
            With wsP.Cells(rad, c).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(180, 180, 180)
            End With
        End If
    Next c
    
    ' Sikre at kraftige kanter forblir
    Dim rng As Range
    Set rng = wsP.Range(wsP.Cells(rad, startCol), wsP.Cells(rad, sluttCol))
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
End Sub

' Flytt en HEL aktivitet fra en rad til en annen
Private Sub FlyttHeleAktivitetTilNyRad(wsP As Worksheet, gammelRad As Long, nyRad As Long, _
                                       startCol As Long, sluttCol As Long)
    Dim c As Long
    
    ' Kopier hele aktiviteten til ny rad
    wsP.Range(wsP.Cells(gammelRad, startCol), wsP.Cells(gammelRad, sluttCol)).Copy
    wsP.Cells(nyRad, startCol).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    
    ' Rydd gammel rad
    For c = startCol To sluttCol
        Dim cel As Range
        Set cel = wsP.Cells(gammelRad, c)
        cel.ClearContents
        cel.Interior.Color = RGB(255, 255, 255)
        cel.Font.Bold = False
        cel.Interior.Pattern = xlSolid
        
        ' Gjenopprett grid
        With cel.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
        End With
        With cel.Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
        End With
        With cel.Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
        End With
        With cel.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
        End With
    Next c
End Sub

Private Function FinnEllerOpprettLedigRadU5(wsP As Worksheet, personRow As Long, _
                                            startCol As Long, sluttCol As Long, _
                                            farger As Object, førsteDatoKol As Long, datoRad As Long) As Long
    Dim blockStart As Long, blockEnd As Long, r As Long
    Dim c As Long, fri As Boolean, cel As Range
    Dim lastCol As Long
    
    blockStart = personRow
    blockEnd = personRow
    
    ' Finn hele personblokken
    Do While Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) = 0 And blockEnd < wsP.Rows.Count
        blockEnd = blockEnd + 1
    Loop
    
    ' Finn ledig rad i blokken
    For r = blockStart To blockEnd
        fri = True
        For c = startCol To sluttCol
            Set cel = wsP.Cells(r, c)
            ' Sjekk om cellen har tekst
            If Len(Trim$(cel.Value)) > 0 Then
                fri = False: Exit For
            End If
            ' Sjekk om cellen har aktivitetsfarge (ikke hvit eller skravering)
            If cel.Interior.ColorIndex <> xlColorIndexNone Then
                If cel.Interior.Color <> RGB(255, 255, 255) And _
                   cel.Interior.Color <> RGB(255, 220, 220) Then ' Ignorer skravering
                    If FargeNærAktivitetU5(cel.Interior.Color, farger) Then
                        fri = False: Exit For
                    End If
                End If
            End If
        Next c
        If fri Then
            FinnEllerOpprettLedigRadU5 = r
            Exit Function
        End If
    Next r
    
    ' Opprett ny under-rad etter blokken
    wsP.Rows(blockEnd + 1).Insert Shift:=xlDown
    wsP.Rows(blockStart).Copy
    wsP.Rows(blockEnd + 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    wsP.Cells(blockEnd + 1, 1).ClearContents
    
    ' Nullstill alle datoceller til hvit med grid
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
    For c = førsteDatoKol To lastCol
        NullstillCelleTilHvitMedGridU5 wsP.Cells(blockEnd + 1, c)
    Next c
    
    FinnEllerOpprettLedigRadU5 = blockEnd + 1
End Function

' Nullstill celle til hvit med normalt rutenett
Private Sub NullstillCelleTilHvitMedGridU5(ByVal cel As Range)
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

' Sjekk om farge er nær aktivitetsfarge
Private Function FargeNærAktivitetU5(col As Long, ByVal farger As Object, Optional tol As Long = 18) As Boolean
    If col = RGB(255, 255, 255) Then Exit Function
    If col = RGB(255, 220, 220) Then Exit Function ' Ignorer skravering
    
    Dim k As Variant, refCol As Long
    For Each k In farger.Keys
        refCol = CLng(farger(k))
        If FargeAvstandU5(col, refCol) <= tol Then
            FargeNærAktivitetU5 = True
            Exit Function
        End If
    Next k
End Function

' Beregn fargeAvstand
Private Function FargeAvstandU5(c1 As Long, c2 As Long) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = c1 Mod 256: g1 = (c1 \ 256) Mod 256: b1 = (c1 \ 65536) Mod 256
    r2 = c2 Mod 256: g2 = (c2 \ 256) Mod 256: b2 = (c2 \ 65536) Mod 256
    FargeAvstandU5 = Application.WorksheetFunction.Max(Abs(r1 - r2), Abs(g1 - g2), Abs(b1 - b2))
End Function

' Utvid en eksisterende blokk (ikke lag ny)
Private Sub ApplyBlockFormattingExtend(wsP As Worksheet, målRad As Long, _
                                       startCol As Long, sluttCol As Long, _
                                       farge As Long, visTekst As String)
    Dim rng As Range, startCell As Range, rngUnder As Range
    Dim c As Long
    
    Set rng = wsP.Range(wsP.Cells(målRad, startCol), wsP.Cells(målRad, sluttCol))
    Set startCell = wsP.Cells(målRad, startCol)
    
    ' Fyll fargen over hele spennet
    For c = startCol To sluttCol
        wsP.Cells(målRad, c).Interior.Pattern = xlSolid
        wsP.Cells(målRad, c).Interior.TintAndShade = 0
        wsP.Cells(målRad, c).Interior.Color = farge
    Next c
    
    ' Sett kraftige ytterkanter på hele blokken
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
    
    ' Gjenopprett tynne grid-linjer INNI blokken
    For c = startCol To sluttCol
        With wsP.Cells(målRad, c).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(200, 200, 200)  ' Lys grå for subtile grid-linjer
        End With
    Next c
    
    ' Ingen borders på innsiden horisontalt
    rng.Borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    
    ' Sett tekst (kun i første celle)
    startCell.Value = visTekst
    startCell.Font.Bold = True
    startCell.Font.Color = IIf(ErLysFarge(farge), RGB(0, 0, 0), RGB(255, 255, 255))
    
    ' Sentrer tekst over HELE spennet (inkl. skravering)
    rng.HorizontalAlignment = xlCenterAcrossSelection
    rng.VerticalAlignment = xlCenter
    rng.WrapText = True
    
    ' Gjenopprett bunnlinje på raden under
    If målRad < wsP.Rows.Count Then
        Set rngUnder = wsP.Range(wsP.Cells(målRad + 1, startCol), wsP.Cells(målRad + 1, sluttCol))
        With rngUnder.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(0, 0, 0)
        End With
    End If
End Sub

' Tegn diagonal skravering for overlappende aktiviteter
Private Sub TegnOverlappSkravering(wsP As Worksheet, overlappListe As Object, _
                                   førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long)
    Dim k As Variant
    Dim info As Variant
    Dim person As String, personRow As Long
    Dim startDato As Date, sluttDato As Date
    Dim startCol As Long, sluttCol As Long
    Dim r As Long, blockEnd As Long
    Dim rng As Range
    
    ' Gå gjennom alle overlapp
    For Each k In overlappListe.Keys
        info = overlappListe(k)
        person = info(0)
        startDato = info(1)
        sluttDato = info(2)
        
        ' Finn person
        personRow = FinnPersonRadIPlanlegger(wsP, person, førstePersonRad)
        If personRow > 0 Then
            ' Finn datokolonner
            startCol = FinnDatoKolonneIPlanlegger(wsP, startDato, datoRad, førsteDatoKol)
            sluttCol = FinnDatoKolonneIPlanlegger(wsP, sluttDato, datoRad, førsteDatoKol)
            
            If startCol > 0 And sluttCol > 0 Then
                ' Finn personblokken
                blockEnd = personRow
                Do While blockEnd < wsP.Rows.Count
                    If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
                    blockEnd = blockEnd + 1
                Loop
                
                ' Tegn skravering på alle rader i personblokken for dette tidsrommet
                For r = personRow To blockEnd
                    Set rng = wsP.Range(wsP.Cells(r, startCol), wsP.Cells(r, sluttCol))
                    
                    ' Bare legg skravering UNDER eksisterende aktiviteter (ikke overskrive)
                    ' Sjekk hver celle
                    Dim c As Long, cel As Range
                    For c = startCol To sluttCol
                        Set cel = wsP.Cells(r, c)
                        
                        ' Hvis cellen ikke har aktivitet (ikke fet tekst og ikke aktivitetsfarge)
                        If Not cel.Font.Bold And _
                           (cel.Interior.ColorIndex = xlColorIndexNone Or cel.Interior.Color = RGB(255, 255, 255)) Then
                            ' Legg diagonal skravering i RØDT
                            With cel.Interior
                                .Pattern = xlPatternLightDown  ' Diagonal skravering
                                .PatternColor = RGB(255, 0, 0)  ' RØD skravering (ikke grå)
                                .Color = RGB(255, 220, 220)  ' Lys rød bakgrunn
                            End With
                        End If
                    Next c
                Next r
            End If
        End If
    Next k
End Sub

' Hjelpefunksjoner for å finne person og dato i Planlegger
Private Function FinnPersonRadIPlanlegger(ws As Worksheet, ByVal navn As String, førstePersonRad As Long) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = førstePersonRad To lastRow
        If StrComp(Trim$(ws.Cells(r, 1).Value), Trim$(navn), vbTextCompare) = 0 Then
            FinnPersonRadIPlanlegger = r: Exit Function
        End If
    Next r
End Function

Private Function FinnDatoKolonneIPlanlegger(ws As Worksheet, d As Date, datoRad As Long, førsteDatoKol As Long) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(datoRad, ws.Columns.Count).End(xlToLeft).Column
    For c = førsteDatoKol To lastCol
        If IsDate(ws.Cells(datoRad, c).Value) Then
            If CLng(CDate(ws.Cells(datoRad, c).Value)) = CLng(d) Then
                FinnDatoKolonneIPlanlegger = c: Exit Function
            End If
        End If
    Next c
End Function

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

' =================== HJELPERE ===================

Private Sub Boxify(r As Range)
    With r.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
End Sub

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
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Adjustments.Item(1) = 0.2
        .OnAction = makro
    End With
End Sub

Private Sub TryDeleteShape(ws As Worksheet, ByVal shpName As String)
    On Error Resume Next: ws.Shapes(shpName).Delete: On Error GoTo 0
End Sub

' =================== DYNAMISKE VERDIER ===================
' Henter fra Named Ranges som de andre modulene

Private Function HentFørsteDatoKol() As Long
    HentFørsteDatoKol = Worksheets(ARK_PLAN).Range("FirstDate").Column
End Function

Private Function HentDatoRad() As Long
    HentDatoRad = Worksheets(ARK_PLAN).Range("FirstDate").Row
End Function

Private Function HentFørstePersonRad() As Long
    HentFørstePersonRad = Worksheets(ARK_PLAN).Range("PersonHeader").Row + 1
End Function

' =================== FORSINKELSESHÅNDTERING ===================
' Lagre og gjenopprett forsinkelser ved refresh

' Lagre alle forsinkelser før refresh
Private Function LagreForsinkelser(wsAO As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, r As Long
    Dim person As String, kode As String, oppStart As Date
    Dim forsinkelse As Long
    Dim key As String
    
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    
    For r = TBL_START_ROW To lastRow
        person = Trim$(wsAO.Cells(r, COL_PERSON).Value)
        kode = Trim$(wsAO.Cells(r, COL_KODE).Value)
        
        If Len(person) > 0 And Len(kode) > 0 Then
            If IsDate(wsAO.Cells(r, COL_OPP_START).Value) Then
                oppStart = wsAO.Cells(r, COL_OPP_START).Value
                forsinkelse = 0
                
                On Error Resume Next
                forsinkelse = CLng(wsAO.Cells(r, COL_FORSINKET).Value)
                On Error GoTo 0
                
                ' Lag unik nøkkel
                key = person & "|" & kode & "|" & Format(oppStart, "yyyy-mm-dd")
                
                If Not dict.exists(key) Then
                    dict.Add key, forsinkelse
                End If
            End If
        End If
    Next r
    
    Set LagreForsinkelser = dict
End Function

' Gjenopprett forsinkelser etter refresh
Private Sub GjenopprettForsinkelser(wsAO As Worksheet, forsinkelser As Object)
    Dim lastRow As Long, r As Long
    Dim person As String, kode As String, oppStart As Date
    Dim key As String
    
    If forsinkelser Is Nothing Then Exit Sub
    If forsinkelser.Count = 0 Then Exit Sub
    
    ' VIKTIG: Fjern beskyttelse først
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    
    For r = TBL_START_ROW To lastRow
        person = Trim$(wsAO.Cells(r, COL_PERSON).Value)
        kode = Trim$(wsAO.Cells(r, COL_KODE).Value)
        
        If Len(person) > 0 And Len(kode) > 0 Then
            If IsDate(wsAO.Cells(r, COL_OPP_START).Value) Then
                oppStart = wsAO.Cells(r, COL_OPP_START).Value
                
                ' Lag samme nøkkel
                key = person & "|" & kode & "|" & Format(oppStart, "yyyy-mm-dd")
                
                ' Gjenopprett forsinkelse hvis den finnes
                If forsinkelser.exists(key) Then
                    wsAO.Cells(r, COL_FORSINKET).Value = forsinkelser(key)
                End If
            End If
        End If
    Next r
    
    ' Trigger ny beregning av formler og overlappdeteksjon
    Application.Calculate
    
    ' Kjør overlappdeteksjon på nytt
    Call DetekterOverlapp(wsAO, TBL_START_ROW, lastRow)
End Sub

' Lagre alle redigerbare felt (Person, Forsinkelse, Kommentar)
Private Function LagreAlleRedigerbareFelt(wsAO As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, r As Long
    Dim person As String, kode As String, oppStart As Date
    Dim key As String
    Dim forsinkelse As Long, kommentar As String
    
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    
    For r = TBL_START_ROW To lastRow
        person = Trim$(wsAO.Cells(r, COL_PERSON).Value)
        kode = Trim$(wsAO.Cells(r, COL_KODE).Value)
        
        If Len(person) > 0 And Len(kode) > 0 Then
            If IsDate(wsAO.Cells(r, COL_OPP_START).Value) Then
                oppStart = wsAO.Cells(r, COL_OPP_START).Value
                forsinkelse = wsAO.Cells(r, COL_FORSINKET).Value
                kommentar = Trim$(wsAO.Cells(r, COL_KOMMENTAR).Value)
                
                ' Lag nøkkel basert på person + kode + startdato
                key = person & "|" & kode & "|" & Format(oppStart, "yyyy-mm-dd")
                
                ' Lagre alle felt
                Dim feltDict As Object
                Set feltDict = CreateObject("Scripting.Dictionary")
                feltDict("Person") = person
                feltDict("Forsinkelse") = forsinkelse
                feltDict("Kommentar") = kommentar
                
                dict.Add key, feltDict
            End If
        End If
    Next r
    
    Set LagreAlleRedigerbareFelt = dict
End Function

' Gjenopprett alle redigerbare felt etter refresh
Private Sub GjenopprettAlleRedigerbareFelt(wsAO As Worksheet, alleEndringer As Object)
    Dim lastRow As Long, r As Long
    Dim person As String, kode As String, oppStart As Date
    Dim key As String
    Dim feltDict As Object
    
    If alleEndringer Is Nothing Then Exit Sub
    If alleEndringer.Count = 0 Then Exit Sub
    
    ' VIKTIG: Fjern beskyttelse først
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    
    ' Først: prøv å matche på person+kode+dato (eksakt match)
    For r = TBL_START_ROW To lastRow
        person = Trim$(wsAO.Cells(r, COL_PERSON).Value)
        kode = Trim$(wsAO.Cells(r, COL_KODE).Value)
        
        If Len(person) > 0 And Len(kode) > 0 Then
            If IsDate(wsAO.Cells(r, COL_OPP_START).Value) Then
                oppStart = wsAO.Cells(r, COL_OPP_START).Value
                
                ' Prøv nøkkel med NÅVÆRENDE person (etter refresh)
                key = person & "|" & kode & "|" & Format(oppStart, "yyyy-mm-dd")
                
                If alleEndringer.exists(key) Then
                    Set feltDict = alleEndringer(key)
                    
                    ' Gjenopprett felt
                    wsAO.Cells(r, COL_FORSINKET).Value = feltDict("Forsinkelse")
                    wsAO.Cells(r, COL_KOMMENTAR).Value = feltDict("Kommentar")
                    
                    ' Merk som behandlet
                    alleEndringer.Remove key
                End If
            End If
        End If
    Next r
    
    ' Andre pass: Hvis person ble endret, må vi finne aktiviteten basert på kode+dato hos NY person
    If alleEndringer.Count > 0 Then
        ' Gå gjennom de som ikke ble matchet (person ble endret)
        Dim k As Variant
        For Each k In alleEndringer.Keys
            Set feltDict = alleEndringer(k)
            Dim nyPerson As String
            nyPerson = feltDict("Person")
            
            ' Finn aktiviteten hos den nye personen
            Dim parts() As String
            parts = Split(CStr(k), "|")
            Dim søkKode As String, søkDato As String
            søkKode = parts(1)
            søkDato = parts(2)
            
            For r = TBL_START_ROW To lastRow
                person = Trim$(wsAO.Cells(r, COL_PERSON).Value)
                kode = Trim$(wsAO.Cells(r, COL_KODE).Value)
                
                ' Match på person (NY), kode og dato
                If person = nyPerson And kode = søkKode Then
                    If IsDate(wsAO.Cells(r, COL_OPP_START).Value) Then
                        oppStart = wsAO.Cells(r, COL_OPP_START).Value
                        If Format(oppStart, "yyyy-mm-dd") = søkDato Then
                            ' TREFF! Dette er aktiviteten som ble flyttet
                            wsAO.Cells(r, COL_FORSINKET).Value = feltDict("Forsinkelse")
                            wsAO.Cells(r, COL_KOMMENTAR).Value = feltDict("Kommentar")
                            Exit For
                        End If
                    End If
                End If
            Next r
        Next k
    End If
    
    ' Trigger ny beregning
    Application.Calculate
End Sub

' =================== UNDO-FUNKSJONALITET ===================

' Angre siste endring - gjenoppretter Planlegger til forrige tilstand
Public Sub AktivitetsOversikt_Undo()
    Dim wsAO As Worksheet, wsP As Worksheet
    
    On Error Resume Next
    Set wsAO = ThisWorkbook.Worksheets(ARK_OVERSIKT_AKT)
    Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    On Error GoTo 0
    
    If wsAO Is Nothing Or wsP Is Nothing Then
        MsgBox "Mangler nødvendige ark.", vbCritical
        Exit Sub
    End If
    
    If SisteForsinkelser Is Nothing Then
        MsgBox "Ingen endringer å angre.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Dette vil angre siste endring og gjenopprette Planlegger." & vbCrLf & vbCrLf & _
              "Vil du fortsette?", vbYesNo + vbQuestion, "Angre endring") <> vbYes Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Gjenopprett Planlegger-tilstand
    Call GjenopprettPlanleggerTilstand(wsP, SisteForsinkelser)
    
    ' Nullstill forsinkelser i oversikten
    Call NullstillAlleForsinkelser(wsAO)
    
    ' Refresh oversikten
    Call AktivitetsOversikt_Refresh
    
    ' Tøm undo-lageret
    Set SisteForsinkelser = Nothing
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Endringen er angret.", vbInformation
End Sub

' Lagre hele Planlegger-tilstanden for Undo
Private Function LagrePlanleggerTilstand(wsP As Worksheet, førsteDatoKol As Long, _
                                         datoRad As Long, førstePersonRad As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim key As String
    Dim cellInfo As Object
    
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
    
    ' Lagre alle celler i dato-området
    For r = førstePersonRad To lastRow
        For c = førsteDatoKol To lastCol
            key = r & "|" & c
            
            Set cellInfo = CreateObject("Scripting.Dictionary")
            cellInfo("Value") = wsP.Cells(r, c).Value
            cellInfo("Color") = wsP.Cells(r, c).Interior.Color
            cellInfo("Pattern") = wsP.Cells(r, c).Interior.Pattern
            cellInfo("PatternColor") = wsP.Cells(r, c).Interior.PatternColor
            cellInfo("Bold") = wsP.Cells(r, c).Font.Bold
            cellInfo("FontColor") = wsP.Cells(r, c).Font.Color
            
            dict.Add key, cellInfo
        Next c
    Next r
    
    Set LagrePlanleggerTilstand = dict
End Function

' Gjenopprett Planlegger-tilstand fra Undo-lager
Private Sub GjenopprettPlanleggerTilstand(wsP As Worksheet, tilstand As Object)
    If tilstand Is Nothing Then Exit Sub
    If tilstand.Count = 0 Then Exit Sub
    
    Dim k As Variant
    Dim parts() As String
    Dim r As Long, c As Long
    Dim cellInfo As Object
    
    For Each k In tilstand.Keys
        parts = Split(CStr(k), "|")
        r = CLng(parts(0))
        c = CLng(parts(1))
        
        Set cellInfo = tilstand(k)
        
        With wsP.Cells(r, c)
            .Value = cellInfo("Value")
            .Interior.Color = cellInfo("Color")
            .Interior.Pattern = cellInfo("Pattern")
            
            On Error Resume Next
            .Interior.PatternColor = cellInfo("PatternColor")
            On Error GoTo 0
            
            .Font.Bold = cellInfo("Bold")
            .Font.Color = cellInfo("FontColor")
        End With
    Next k
End Sub

' Nullstill alle forsinkelser i oversikten
Private Sub NullstillAlleForsinkelser(wsAO As Worksheet)
    Dim lastRow As Long, r As Long
    
    ' VIKTIG: Fjern beskyttelse først
    On Error Resume Next
    wsAO.Unprotect Password:=""
    On Error GoTo 0
    
    lastRow = wsAO.Cells(wsAO.Rows.Count, COL_PERSON).End(xlUp).Row
    
    For r = TBL_START_ROW To lastRow
        wsAO.Cells(r, COL_FORSINKET).Value = 0
    Next r
End Sub

' =================== DIREKTE SYNKRONISERING FRA OVERSIKT TIL PLANLEGGER ===================

' Oppdater aktivitet i Planlegger basert på endringer i Oversikt
Public Sub SynkroniserEnkeltAktivitet(person As String, kode As String, _
                                      nyStartDato As Date, nySluttDato As Date, _
                                      Optional kommentar As String = "")
    Dim wsP As Worksheet, wsTyp As Worksheet
    Dim personRow As Long, målRad As Long
    Dim gammelStartCol As Long, gammelSluttCol As Long
    Dim nyStartCol As Long, nySluttCol As Long
    Dim beskrivelse As String, farge As Long
    Dim r As Long, c As Long, blockEnd As Long
    Dim funnet As Boolean
    Dim førsteDatoKol As Long, datoRad As Long
    
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT_TYP)
    On Error GoTo 0
    
    If wsP Is Nothing Or wsTyp Is Nothing Then Exit Sub
    
    ' Få konstanter
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    Dim førstePersonRad As Long
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    
    ' Finn personen
    personRow = FinnPersonRad(wsP, person, førstePersonRad)
    If personRow = 0 Then
        If Application.EnableEvents Then
            MsgBox "Finner ikke person '" & person & "' i Planlegger.", vbExclamation
        End If
        Exit Sub
    End If
    
    ' Finn aktiviteten i Planlegger
    blockEnd = personRow
    Do While blockEnd < wsP.Rows.Count
        If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
        blockEnd = blockEnd + 1
    Loop
    
    funnet = False
    For r = personRow To blockEnd
        If Len(Trim$(wsP.Cells(r, førsteDatoKol).Value)) > 0 And wsP.Cells(r, førsteDatoKol).Font.Bold Then
            If InStr(1, wsP.Cells(r, førsteDatoKol).Value, kode, vbTextCompare) > 0 Then
                målRad = r
                funnet = True
                Exit For
            End If
        End If
    Next r
    
    If Not funnet Then
        If Application.EnableEvents Then
            MsgBox "Finner ikke aktivitet '" & kode & "' for '" & person & "' i Planlegger.", vbExclamation
        End If
        Exit Sub
    End If
    
    ' Hent aktivitetsinfo
    If Not LookupAktivitet(wsTyp, kode, beskrivelse, farge) Then
        If Application.EnableEvents Then
            MsgBox "Finner ikke aktivitetskode '" & kode & "' i oversikten.", vbExclamation
        End If
        Exit Sub
    End If
    
    ' Finn kolonner for datoer
    nyStartCol = FinnDatoKolonne(wsP, nyStartDato, førsteDatoKol, datoRad)
    nySluttCol = FinnDatoKolonne(wsP, nySluttDato, førsteDatoKol, datoRad)
    
    If nyStartCol = 0 Or nySluttCol = 0 Then
        If Application.EnableEvents Then
            MsgBox "Datoene finnes ikke i Planlegger. Utvid datoene først.", vbExclamation
        End If
        Exit Sub
    End If
    
    ' Finn gamle kolonner (sammenhengende blokk kun)
    gammelStartCol = 0
    gammelSluttCol = 0
    
    ' Finn startkolonnen
    For c = førsteDatoKol To wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
        If wsP.Cells(målRad, c).Interior.Color = farge And _
           wsP.Cells(målRad, c).Interior.Pattern <> xlPatternLightDown Then
            gammelStartCol = c
            Exit For
        End If
    Next c
    
    ' Finn slutt (sammenhengende blokk)
    If gammelStartCol > 0 Then
        gammelSluttCol = gammelStartCol
        For c = gammelStartCol + 1 To wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
            If wsP.Cells(målRad, c).Interior.Color = farge Then
                gammelSluttCol = c
            ElseIf wsP.Cells(målRad, c).Interior.Color = RGB(255, 255, 255) Or _
                   wsP.Cells(målRad, c).Interior.ColorIndex = xlColorIndexNone Then
                Exit For
            Else
                Exit For
            End If
        Next c
    End If
    
    ' Rydd gammel aktivitet
    If gammelStartCol > 0 Then
        For c = gammelStartCol To gammelSluttCol
            With wsP.Cells(målRad, c)
                .ClearContents
                .Interior.Color = RGB(255, 255, 255)
                .Interior.Pattern = xlSolid
                .Font.Bold = False
                
                ' Grid
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
                End With
            End With
        Next c
    End If
    
    ' Tegn ny aktivitet
    Dim visTekst As String
    visTekst = kode & IIf(Len(kommentar) > 0, " - " & kommentar, IIf(Len(beskrivelse) > 0, " - " & beskrivelse, ""))
    
    Call ApplyBlockFormattingExtend(wsP, målRad, nyStartCol, nySluttCol, farge, visTekst)
    
    ' MsgBox "Aktivitet oppdatert i Planlegger!", vbInformation
End Sub

' Finn dato-kolonne i Planlegger
Private Function FinnDatoKolonne(wsP As Worksheet, d As Date, førsteDatoKol As Long, datoRad As Long) As Long
    Dim lastCol As Long, c As Long
    lastCol = wsP.Cells(datoRad, wsP.Columns.Count).End(xlToLeft).Column
    
    For c = førsteDatoKol To lastCol
        If IsDate(wsP.Cells(datoRad, c).Value) Then
            If CLng(CDate(wsP.Cells(datoRad, c).Value)) = CLng(d) Then
                FinnDatoKolonne = c
                Exit Function
            End If
        End If
    Next c
End Function

' Finn person-rad i Planlegger
Private Function FinnPersonRad(wsP As Worksheet, person As String, førstePersonRad As Long) As Long
    Dim lastRow As Long, r As Long
    lastRow = wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row
    
    For r = førstePersonRad To lastRow
        If StrComp(Trim$(wsP.Cells(r, 1).Value), Trim$(person), vbTextCompare) = 0 Then
            FinnPersonRad = r
            Exit Function
        End If
    Next r
End Function

' Flytt aktivitet fra en person til en annen
Public Sub FlyttAktivitetTilNyPerson(gammelPerson As String, nyPerson As String, _
                                     kode As String, startDato As Date, sluttDato As Date, _
                                     Optional kommentar As String = "")
    Dim wsP As Worksheet, wsTyp As Worksheet
    Dim gammelPersonRow As Long, nyPersonRow As Long
    Dim gammelRad As Long, nyRad As Long
    Dim startCol As Long, sluttCol As Long
    Dim r As Long, c As Long, blockEnd As Long
    Dim beskrivelse As String, farge As Long
    Dim funnet As Boolean
    Dim førsteDatoKol As Long, datoRad As Long, førstePersonRad As Long
    Dim farger As Object
    
    On Error Resume Next
    Set wsP = ThisWorkbook.Worksheets(ARK_PLAN)
    Set wsTyp = ThisWorkbook.Worksheets(ARK_OVERSIKT_TYP)
    On Error GoTo 0
    
    If wsP Is Nothing Or wsTyp Is Nothing Then
        MsgBox "Finner ikke nødvendige ark.", vbCritical
        Exit Sub
    End If
    
    ' Hent konstanter
    førsteDatoKol = wsP.Range("FirstDate").Column
    datoRad = wsP.Range("FirstDate").Row
    førstePersonRad = wsP.Range("PersonHeader").Row + 1
    
    ' Finn gammel person
    gammelPersonRow = FinnPersonRad(wsP, gammelPerson, førstePersonRad)
    If gammelPersonRow = 0 Then
        ' Kanskje aktiviteten er allerede hos ny person? Prøv å synkroniser i stedet
        Call SynkroniserEnkeltAktivitet(nyPerson, kode, startDato, sluttDato, kommentar)
        Exit Sub
    End If
    
    ' Finn ny person
    nyPersonRow = FinnPersonRad(wsP, nyPerson, førstePersonRad)
    If nyPersonRow = 0 Then
        MsgBox "Finner ikke ny person '" & nyPerson & "' i Planlegger.", vbExclamation
        Exit Sub
    End If
    
    ' Finn aktiviteten hos gammel person
    blockEnd = gammelPersonRow
    Do While blockEnd < wsP.Rows.Count
        If Len(Trim$(wsP.Cells(blockEnd + 1, 1).Value)) > 0 Then Exit Do
        blockEnd = blockEnd + 1
    Loop
    
    funnet = False
    For r = gammelPersonRow To blockEnd
        If Len(Trim$(wsP.Cells(r, førsteDatoKol).Value)) > 0 And wsP.Cells(r, førsteDatoKol).Font.Bold Then
            If InStr(1, wsP.Cells(r, førsteDatoKol).Value, kode, vbTextCompare) > 0 Then
                gammelRad = r
                funnet = True
                Exit For
            End If
        End If
    Next r
    
    If Not funnet Then
        MsgBox "Finner ikke aktivitet '" & kode & "' hos '" & gammelPerson & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Hent aktivitetsinfo
    If Not LookupAktivitet(wsTyp, kode, beskrivelse, farge) Then
        MsgBox "Finner ikke aktivitetskode '" & kode & "'.", vbExclamation
        Exit Sub
    End If
    
    ' Finn kolonner for datoer
    startCol = FinnDatoKolonne(wsP, startDato, førsteDatoKol, datoRad)
    sluttCol = FinnDatoKolonne(wsP, sluttDato, førsteDatoKol, datoRad)
    
    If startCol = 0 Or sluttCol = 0 Then
        MsgBox "Datoene finnes ikke i Planlegger.", vbExclamation
        Exit Sub
    End If
    
    ' STEG 1: Rydd aktiviteten fra gammel person
    For c = startCol To sluttCol
        With wsP.Cells(gammelRad, c)
            .ClearContents
            .Interior.Color = RGB(255, 255, 255)
            .Interior.Pattern = xlSolid
            .Font.Bold = False
            .Font.ColorIndex = xlColorIndexAutomatic
            
            ' Grid
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlThin: .ColorIndex = xlColorIndexAutomatic
            End With
        End With
    Next c
    
    ' STEG 2: Finn ledig rad hos ny person
    Set farger = HentAktivitetsFarger(wsTyp)
    nyRad = FinnEllerOpprettLedigRadU5(wsP, nyPersonRow, startCol, sluttCol, farger, førsteDatoKol, datoRad)
    
    If nyRad = 0 Then
        MsgBox "Fant ikke ledig rad hos '" & nyPerson & "'.", vbExclamation
        Exit Sub
    End If
    
    ' STEG 3: Tegn aktivitet hos ny person
    Dim visTekst As String
    visTekst = kode & IIf(Len(kommentar) > 0, " - " & kommentar, IIf(Len(beskrivelse) > 0, " - " & beskrivelse, ""))
    
    Call ApplyBlockFormattingExtend(wsP, nyRad, startCol, sluttCol, farge, visTekst)
    
    MsgBox "Aktivitet flyttet fra '" & gammelPerson & "' til '" & nyPerson & "'!", vbInformation
End Sub








