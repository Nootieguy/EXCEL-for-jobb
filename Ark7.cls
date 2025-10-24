Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Slutt
    If Target Is Nothing Then Exit Sub
    
    ' Disable events for � unng� loop
    Application.EnableEvents = False
    
    ' Personkolonne = F (6), fra rad 10 og ned
    Dim rng As Range, c As Range
    Set rng = Application.Intersect(Target, Me.Range(Me.Cells(10, 6), Me.Cells(Me.Rows.Count, 6)))
    
    If Not rng Is Nothing Then
        ' Person er valgt - tildel rad
        For Each c In rng.Cells
            If Len(Trim$(c.Value)) > 0 Then
                c.Value = Trim$(Replace(CStr(c.Value), ChrW(160), " "))
                Call Uvalgte_TildelRad(c.Row)
            End If
        Next c
    Else
        ' Sjekk om det er endringer i kolonne A-E (kode, dato osv.)
        Set rng = Application.Intersect(Target, Me.Range(Me.Cells(10, 1), Me.Cells(Me.Rows.Count, 5)))
        If Not rng Is Nothing Then
            ' Refresh preview automatisk (med forsinkelse for � unng� spam)
            Application.OnTime Now + TimeValue("00:00:01"), "Uvalgte_RefreshPreview"
        End If
    End If
    
Slutt:
    Application.EnableEvents = True
End Sub
