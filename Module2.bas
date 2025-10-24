Option Explicit

' =================== KONFIG ===================
Private Const ARK_PLAN As String = "Planlegger"
Private Const FØRSTE_DATAKOL As Long = 2    ' B = 2
Private Const datoRad As Long = 15          ' rad med datoer (ekte datoer)
Private Const FØRSTE_PERSONRAD As Long = 16
' Standard grid: tynn, automatisk farge
Private Const GRID_WEIGHT As Long = xlHairline   ' bruk xlThin for sterkere ruter
' =============================================

' RYDD: fjerner alt i valgt spenn, tegner rutenett på nytt, sletter tomme under-rader
Public Sub RyddBlokkForPerson()
    Dim ws As Worksheet
    Dim personCell As Range
    Dim startDato As Date, sluttDato As Date
    Dim startCol As Long, sluttCol As Long
    Dim blockStart As Long, blockEnd As Long
    Dim lastCol As Long, r As Long
    Dim rng As Range
    
    Set ws = ThisWorkbook.Worksheets(ARK_PLAN)
    
    ' Velg person (hovedrad)
    On Error Resume Next
    Set personCell = Application.InputBox( _
        prompt:="Klikk personens HOVEDRAD i kolonne A (rad " & FØRSTE_PERSONRAD & "+).", _
        Title:="Velg person", Type:=8)
    On Error GoTo 0
    If personCell Is Nothing Then Exit Sub
    If personCell.Column <> 1 Or personCell.Row < FØRSTE_PERSONRAD Then
        MsgBox "Velg i kol A fra rad " & FØRSTE_PERSONRAD & ".", vbExclamation: Exit Sub
    End If
    
    ' Datoer
    If Not HentDato("Startdato (dd.mm.åååå) som skal ryddes:", startDato) Then Exit Sub
    If Not HentDato("Sluttdato (dd.mm.åååå):", sluttDato) Then Exit Sub
    If sluttDato < startDato Then
        MsgBox "Sluttdato < Startdato.", vbExclamation
        Exit Sub
    End If
    
    ' Kolonner for dato-spenn
    startCol = FinnKolonneForDato_Rad13(ws, startDato, FØRSTE_DATAKOL, datoRad)
    sluttCol = FinnKolonneForDato_Rad13(ws, sluttDato, FØRSTE_DATAKOL, datoRad)
    If startCol = 0 Or sluttCol = 0 Then
        MsgBox "Fant ikke datoene i rad " & datoRad & ".", vbCritical: Exit Sub
    End If
    If sluttCol < startCol Then
        Dim t As Long
        t = startCol: startCol = sluttCol: sluttCol = t
    End If
    
    ' Siste datokolonne i visningen
    lastCol = ws.Cells(datoRad, ws.Columns.Count).End(xlToLeft).Column
    
    ' Finn personblokken
    FinnPersonBlokk ws, personCell.Row, blockStart, blockEnd
    
    Application.ScreenUpdating = False
    
    ' 1) Rydd valgt spenn på hele blokken
    For r = blockStart To blockEnd
        Set rng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, sluttCol))
        
        ' Tøm
        rng.ClearContents
        rng.Interior.ColorIndex = xlColorIndexNone
        rng.Borders.LineStyle = xlLineStyleNone
        rng.HorizontalAlignment = xlGeneral
        rng.VerticalAlignment = xlCenter
        rng.Font.Bold = False
        rng.Font.ColorIndex = xlColorIndexAutomatic
        rng.WrapText = False
        On Error Resume Next
        ws.Cells(r, startCol).ClearComments
        On Error GoTo 0
        
        ' Tegn rutenett på nytt (tynne inndelingslinjer)
        With rng.Borders
            .LineStyle = xlContinuous
            .ColorIndex = xlColorIndexAutomatic
            .Weight = GRID_WEIGHT
        End With
    Next r
    
    ' 2) Gjenopprett overordnet formatering på under-rader fra hovedraden
    If blockEnd > blockStart Then
        ws.Rows(blockStart).Copy
        ws.Range(ws.Rows(blockStart + 1), ws.Rows(blockEnd)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
    
    ' 3) Slett tomme under-rader (aldri hovedraden)
    If blockEnd > blockStart Then
        SlettTommeUnderRader ws, blockStart, blockEnd, FØRSTE_DATAKOL, lastCol
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "Ryddet " & Format(startDato, "dd.mm.yyyy") & "" & _
           Format(sluttDato, "dd.mm.yyyy") & " for personblokken, og rutenett er gjenopprettet.", vbInformation
End Sub

' ---------- HJELPERE ----------

Private Sub FinnPersonBlokk(ws As Worksheet, hovedRad As Long, _
                            ByRef blockStart As Long, ByRef blockEnd As Long)
    Dim r As Long, lastRow As Long, v
    blockStart = hovedRad
    blockEnd = hovedRad
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = hovedRad + 1 To lastRow
        v = ws.Cells(r, 1).Value
        If Len(Trim$(v)) = 0 Then
            blockEnd = r
        Else
            Exit For
        End If
    Next r
End Sub

Private Sub SlettTommeUnderRader(ws As Worksheet, blockStart As Long, blockEnd As Long, _
                                 firstDataCol As Long, lastDataCol As Long)
    Dim r As Long
    For r = blockEnd To blockStart + 1 Step -1
        If ErRadTom(ws, r, firstDataCol, lastDataCol) Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

Private Function ErRadTom(ws As Worksheet, rowNum As Long, firstDataCol As Long, lastDataCol As Long) As Boolean
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, firstDataCol), ws.Cells(rowNum, lastDataCol))
    ErRadTom = Application.WorksheetFunction.CountA(rng) = 0
End Function

Private Function FinnKolonneForDato_Rad13(ws As Worksheet, d As Date, _
                                          firstDataCol As Long, headerRow As Long) As Long
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

Private Function HentDato(prompt As String, ByRef d As Date) As Boolean
    Dim s As String
    s = Trim(InputBox(prompt, "Dato"))
    If Len(s) = 0 Then Exit Function
    On Error GoTo Feil
    d = CDate(s)
    HentDato = True
    Exit Function
Feil:
    MsgBox "Ugyldig dato: " & s, vbExclamation
End Function



