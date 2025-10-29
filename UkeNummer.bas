Option Explicit

' Fyller rad 14 med ukenumre sentrert over ukens dager, med linjer på sidene
Public Sub FyllInnUkenumreMedMerge()
    Dim ws As Worksheet
    Dim c As Long, lastCol As Long
    Dim dato As Date, ukeNr As Long, forrigeUke As Long
    Dim ukeStartCol As Long, ukeSluttCol As Long
    Dim rng As Range
    
    Set ws = ThisWorkbook.Worksheets("Planlegger")
    
    ' Finn siste kolonne med dato i rad 15
    lastCol = ws.Cells(15, ws.Columns.Count).End(xlToLeft).Column
    
    Application.ScreenUpdating = False
    
    ' Først, un-merge alle celler i rad 14 og rydd
    ws.Rows(14).UnMerge
    ws.Range(ws.Cells(14, 2), ws.Cells(14, lastCol)).ClearContents
    ws.Range(ws.Cells(14, 2), ws.Cells(14, lastCol)).Borders.LineStyle = xlLineStyleNone
    
    forrigeUke = -1
    ukeStartCol = 2 ' Start fra kolonne B
    
    ' Gå gjennom alle kolonner fra B og utover
    For c = 2 To lastCol + 1 ' +1 for å fange siste uke
        
        If c <= lastCol And IsDate(ws.Cells(15, c).Value) Then
            dato = ws.Cells(15, c).Value
            ukeNr = Application.WorksheetFunction.WeekNum(dato, 2) ' 2 = uken starter på mandag
            
            ' Når vi treffer en ny uke (eller første gang)
            If ukeNr <> forrigeUke And forrigeUke <> -1 Then
                ' Merge forrige uke
                ukeSluttCol = c - 1
                Set rng = ws.Range(ws.Cells(14, ukeStartCol), ws.Cells(14, ukeSluttCol))
                
                If ukeSluttCol >= ukeStartCol Then
                    If ukeSluttCol > ukeStartCol Then rng.Merge
                    rng.Value = "Uke " & forrigeUke
                    rng.Font.Bold = True
                    rng.Font.Size = 10
                    rng.HorizontalAlignment = xlCenter
                    rng.VerticalAlignment = xlCenter
                    
                    ' Sett linjer på sidene
                    With rng.Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .Color = RGB(0, 0, 0)
                    End With
                    With rng.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlMedium
                        .Color = RGB(0, 0, 0)
                    End With
                End If
                
                ' Start ny uke
                ukeStartCol = c
            End If
            
            forrigeUke = ukeNr
            
        ElseIf forrigeUke <> -1 Then
            ' Siste uke når vi når slutten
            ukeSluttCol = c - 1
            Set rng = ws.Range(ws.Cells(14, ukeStartCol), ws.Cells(14, ukeSluttCol))
            
            If ukeSluttCol >= ukeStartCol Then
                If ukeSluttCol > ukeStartCol Then rng.Merge
                rng.Value = "Uke " & forrigeUke
                rng.Font.Bold = True
                rng.Font.Size = 10
                rng.HorizontalAlignment = xlCenter
                rng.VerticalAlignment = xlCenter
                
                ' Sett linjer på sidene
                With rng.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
                With rng.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .Color = RGB(0, 0, 0)
                End With
            End If
            
            Exit For
        End If
    Next c
    
    Application.ScreenUpdating = True
    
    MsgBox "Ukenumre lagt til sentrert med linjer!", vbInformation
End Sub

