Attribute VB_Name = "modDagensOutline"
' ===== modul: modDagensDatoBlå =====
Option Explicit

' Oppsett – juster ved behov
Private Const ARK As String = "Planlegger"
Private Const datoRad As Long = 15          ' datolinjen (ikke 14)
Private Const FORSTE_DATAKOL As Long = 2     ' B = 2

' Marker dagens dato i rad 15 med blått fyll (kun den cellen)
Public Sub MarkerDagensDato_Bla( _
    Optional ByVal ArkNavn As String = ARK, _
    Optional ByVal datoRad As Long = datoRad, _
    Optional ByVal kolStart As Long = FORSTE_DATAKOL)

    Dim ws As Worksheet
    Dim lastCol As Long, c As Long, v As Variant
    Dim colIdag As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ArkNavn)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    ' Finn siste kolonne i datoraden
    lastCol = ws.Cells(datoRad, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < kolStart Then Exit Sub

    ' 1) Nullstill all fyllfarge i rad 15
    With ws.Range(ws.Cells(datoRad, kolStart), ws.Cells(datoRad, lastCol)).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlColorIndexAutomatic
        .ColorIndex = xlColorIndexNone  ' ingen fyll
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    ' 2) Finn kolonne for dagens dato
    colIdag = 0
    For c = kolStart To lastCol
        v = ws.Cells(datoRad, c).Value
        If IsDate(v) Then
            If CLng(CDate(v)) = CLng(Date) Then
                colIdag = c
                Exit For
            End If
        End If
    Next c

    If colIdag = 0 Then Exit Sub  ' ingen treff i dag – ferdig

    ' 3) Sett blått fyll kun i datocellen på rad 15
    With ws.Cells(datoRad, colIdag).Interior
        .Pattern = xlSolid
        .Color = RGB(170, 200, 255)  ' myk blå – juster etter smak
        .TintAndShade = 0
    End With
End Sub


