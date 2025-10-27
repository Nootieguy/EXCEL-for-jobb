Option Explicit

' =================== UNDO MANAGER ===================
' Håndterer undo-funksjonalitet for makroer
' Siden Excel's Ctrl+Z ikke fungerer med makroer
' ====================================================

' Global variabel for å lagre undo-tilstand
Private Type CellState
    Address As String
    Value As Variant
    Formula As String
    InteriorColor As Long
    InteriorColorIndex As Long
    InteriorPattern As Long
    FontBold As Boolean
    FontColor As Long
    FontColorIndex As Long
    HorizontalAlignment As Long
    VerticalAlignment As Long
    WrapText As Boolean
    HasComment As Boolean
    CommentText As String
    ' Borders - alle kanter
    BorderLeftStyle As Long
    BorderLeftWeight As Long
    BorderLeftColor As Long
    BorderLeftColorIndex As Long
    BorderRightStyle As Long
    BorderRightWeight As Long
    BorderRightColor As Long
    BorderRightColorIndex As Long
    BorderTopStyle As Long
    BorderTopWeight As Long
    BorderTopColor As Long
    BorderTopColorIndex As Long
    BorderBottomStyle As Long
    BorderBottomWeight As Long
    BorderBottomColor As Long
    BorderBottomColorIndex As Long
    ' Diagonal borders (viktig for komplette borders)
    BorderDiagDownStyle As Long
    BorderDiagUpStyle As Long
End Type

Private undoStack() As CellState
Private undoStackSize As Long
Private Const MAX_UNDO_LEVELS As Long = 1  ' Bare én undo-nivå (siste operasjon)

' =====================================================
' PUBLIC API
' =====================================================

' Lagre tilstanden til et område før endring
Public Sub LagUndoSnapshot(ByVal rng As Range)
    On Error GoTo ErrorHandler

    ' Validering
    If rng Is Nothing Then Exit Sub
    If rng.Cells.Count = 0 Then Exit Sub
    If rng.Cells.Count > 10000 Then
        ' For store områder, vis advarsel
        Debug.Print "UNDO WARNING: Snapshot av " & rng.Cells.Count & " celler kan ta tid"
    End If

    ' Reset stack
    undoStackSize = 0
    ReDim undoStack(1 To rng.Cells.Count)

    Dim cel As Range
    Dim i As Long
    i = 1

    For Each cel In rng.Cells
        ' Adresse
        undoStack(i).Address = cel.Address

        ' Verdi og formel
        undoStack(i).Value = cel.Value
        undoStack(i).Formula = cel.Formula

        ' Interior (bakgrunn)
        undoStack(i).InteriorColor = cel.Interior.Color
        undoStack(i).InteriorColorIndex = cel.Interior.ColorIndex
        undoStack(i).InteriorPattern = cel.Interior.Pattern

        ' Font
        undoStack(i).FontBold = cel.Font.Bold
        undoStack(i).FontColor = cel.Font.Color
        undoStack(i).FontColorIndex = cel.Font.ColorIndex

        ' Alignment
        undoStack(i).HorizontalAlignment = cel.HorizontalAlignment
        undoStack(i).VerticalAlignment = cel.VerticalAlignment
        undoStack(i).WrapText = cel.WrapText

        ' Comment
        undoStack(i).HasComment = Not cel.Comment Is Nothing
        If undoStack(i).HasComment Then
            undoStack(i).CommentText = cel.Comment.Text
        End If

        ' Borders - Edge borders
        undoStack(i).BorderLeftStyle = cel.Borders(xlEdgeLeft).LineStyle
        undoStack(i).BorderLeftWeight = cel.Borders(xlEdgeLeft).Weight
        undoStack(i).BorderLeftColor = cel.Borders(xlEdgeLeft).Color
        undoStack(i).BorderLeftColorIndex = cel.Borders(xlEdgeLeft).ColorIndex

        undoStack(i).BorderRightStyle = cel.Borders(xlEdgeRight).LineStyle
        undoStack(i).BorderRightWeight = cel.Borders(xlEdgeRight).Weight
        undoStack(i).BorderRightColor = cel.Borders(xlEdgeRight).Color
        undoStack(i).BorderRightColorIndex = cel.Borders(xlEdgeRight).ColorIndex

        undoStack(i).BorderTopStyle = cel.Borders(xlEdgeTop).LineStyle
        undoStack(i).BorderTopWeight = cel.Borders(xlEdgeTop).Weight
        undoStack(i).BorderTopColor = cel.Borders(xlEdgeTop).Color
        undoStack(i).BorderTopColorIndex = cel.Borders(xlEdgeTop).ColorIndex

        undoStack(i).BorderBottomStyle = cel.Borders(xlEdgeBottom).LineStyle
        undoStack(i).BorderBottomWeight = cel.Borders(xlEdgeBottom).Weight
        undoStack(i).BorderBottomColor = cel.Borders(xlEdgeBottom).Color
        undoStack(i).BorderBottomColorIndex = cel.Borders(xlEdgeBottom).ColorIndex

        ' Diagonal borders
        undoStack(i).BorderDiagDownStyle = cel.Borders(xlDiagonalDown).LineStyle
        undoStack(i).BorderDiagUpStyle = cel.Borders(xlDiagonalUp).LineStyle

        i = i + 1
    Next cel

    undoStackSize = rng.Cells.Count
    Debug.Print "UNDO: Snapshot lagret - " & undoStackSize & " celler i " & rng.Address
    Exit Sub

ErrorHandler:
    Debug.Print "UNDO ERROR i LagUndoSnapshot: " & Err.Description
    undoStackSize = 0
End Sub

' Gjenopprett siste lagrede tilstand (UNDO)
Public Sub Undo(ByVal ws As Worksheet)
    If undoStackSize = 0 Then
        MsgBox "Ingen undo-tilstand tilgjengelig.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim i As Long
    Dim cel As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Debug.Print "UNDO: Starter restore av " & undoStackSize & " celler..."

    For i = 1 To undoStackSize
        Set cel = ws.Range(undoStack(i).Address)

        ' Restore value/formula
        If Len(undoStack(i).Formula) > 0 And undoStack(i).Formula <> "=" Then
            cel.Formula = undoStack(i).Formula
        Else
            cel.Value = undoStack(i).Value
        End If

        ' Restore interior
        If undoStack(i).InteriorColorIndex = xlColorIndexNone Then
            cel.Interior.ColorIndex = xlColorIndexNone
        Else
            cel.Interior.Color = undoStack(i).InteriorColor
            cel.Interior.Pattern = undoStack(i).InteriorPattern
        End If

        ' Restore font
        cel.Font.Bold = undoStack(i).FontBold
        If undoStack(i).FontColorIndex = xlColorIndexAutomatic Then
            cel.Font.ColorIndex = xlColorIndexAutomatic
        Else
            cel.Font.Color = undoStack(i).FontColor
        End If

        ' Restore alignment
        cel.HorizontalAlignment = undoStack(i).HorizontalAlignment
        cel.VerticalAlignment = undoStack(i).VerticalAlignment
        cel.WrapText = undoStack(i).WrapText

        ' Restore comment
        cel.ClearComments
        If undoStack(i).HasComment Then
            cel.AddComment undoStack(i).CommentText
        End If

        ' KRITISK: Nullstill ALLE borders først
        cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
        cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone

        ' Restore diagonal borders først
        cel.Borders(xlDiagonalDown).LineStyle = undoStack(i).BorderDiagDownStyle
        cel.Borders(xlDiagonalUp).LineStyle = undoStack(i).BorderDiagUpStyle

        ' Restore edge borders med FULL kontroll
        If undoStack(i).BorderLeftStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeLeft)
                .LineStyle = undoStack(i).BorderLeftStyle
                .Weight = undoStack(i).BorderLeftWeight
                If undoStack(i).BorderLeftColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderLeftColor
                End If
            End With
        End If

        If undoStack(i).BorderRightStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeRight)
                .LineStyle = undoStack(i).BorderRightStyle
                .Weight = undoStack(i).BorderRightWeight
                If undoStack(i).BorderRightColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderRightColor
                End If
            End With
        End If

        If undoStack(i).BorderTopStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeTop)
                .LineStyle = undoStack(i).BorderTopStyle
                .Weight = undoStack(i).BorderTopWeight
                If undoStack(i).BorderTopColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderTopColor
                End If
            End With
        End If

        If undoStack(i).BorderBottomStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeBottom)
                .LineStyle = undoStack(i).BorderBottomStyle
                .Weight = undoStack(i).BorderBottomWeight
                If undoStack(i).BorderBottomColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderBottomColor
                End If
            End With
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Clear stack after undo
    undoStackSize = 0
    Debug.Print "UNDO: Restore fullført!"

    Exit Sub

ErrorHandler:
    Debug.Print "UNDO ERROR i Undo: " & Err.Description & " (celle " & i & ")"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    undoStackSize = 0
End Sub

' Gjenopprett siste lagrede tilstand (STILLE - ingen melding)
Private Sub UndoSilent(ByVal ws As Worksheet)
    If undoStackSize = 0 Then Exit Sub

    On Error GoTo ErrorHandler

    Dim i As Long
    Dim cel As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Debug.Print "UNDO SILENT: Starter restore av " & undoStackSize & " celler..."

    For i = 1 To undoStackSize
        Set cel = ws.Range(undoStack(i).Address)

        ' Restore value/formula
        If Len(undoStack(i).Formula) > 0 And undoStack(i).Formula <> "=" Then
            cel.Formula = undoStack(i).Formula
        Else
            cel.Value = undoStack(i).Value
        End If

        ' Restore interior
        If undoStack(i).InteriorColorIndex = xlColorIndexNone Then
            cel.Interior.ColorIndex = xlColorIndexNone
        Else
            cel.Interior.Color = undoStack(i).InteriorColor
            cel.Interior.Pattern = undoStack(i).InteriorPattern
        End If

        ' Restore font
        cel.Font.Bold = undoStack(i).FontBold
        If undoStack(i).FontColorIndex = xlColorIndexAutomatic Then
            cel.Font.ColorIndex = xlColorIndexAutomatic
        Else
            cel.Font.Color = undoStack(i).FontColor
        End If

        ' Restore alignment
        cel.HorizontalAlignment = undoStack(i).HorizontalAlignment
        cel.VerticalAlignment = undoStack(i).VerticalAlignment
        cel.WrapText = undoStack(i).WrapText

        ' Restore comment
        cel.ClearComments
        If undoStack(i).HasComment Then
            cel.AddComment undoStack(i).CommentText
        End If

        ' KRITISK: Nullstill ALLE borders først
        cel.Borders(xlDiagonalDown).LineStyle = xlLineStyleNone
        cel.Borders(xlDiagonalUp).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeRight).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeTop).LineStyle = xlLineStyleNone
        cel.Borders(xlEdgeBottom).LineStyle = xlLineStyleNone

        ' Restore diagonal borders først
        cel.Borders(xlDiagonalDown).LineStyle = undoStack(i).BorderDiagDownStyle
        cel.Borders(xlDiagonalUp).LineStyle = undoStack(i).BorderDiagUpStyle

        ' Restore edge borders med FULL kontroll
        If undoStack(i).BorderLeftStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeLeft)
                .LineStyle = undoStack(i).BorderLeftStyle
                .Weight = undoStack(i).BorderLeftWeight
                If undoStack(i).BorderLeftColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderLeftColor
                End If
            End With
        End If

        If undoStack(i).BorderRightStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeRight)
                .LineStyle = undoStack(i).BorderRightStyle
                .Weight = undoStack(i).BorderRightWeight
                If undoStack(i).BorderRightColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderRightColor
                End If
            End With
        End If

        If undoStack(i).BorderTopStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeTop)
                .LineStyle = undoStack(i).BorderTopStyle
                .Weight = undoStack(i).BorderTopWeight
                If undoStack(i).BorderTopColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderTopColor
                End If
            End With
        End If

        If undoStack(i).BorderBottomStyle <> xlLineStyleNone Then
            With cel.Borders(xlEdgeBottom)
                .LineStyle = undoStack(i).BorderBottomStyle
                .Weight = undoStack(i).BorderBottomWeight
                If undoStack(i).BorderBottomColorIndex = xlColorIndexAutomatic Then
                    .ColorIndex = xlColorIndexAutomatic
                Else
                    .Color = undoStack(i).BorderBottomColor
                End If
            End With
        End If
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Clear stack after undo
    undoStackSize = 0
    Debug.Print "UNDO SILENT: Restore fullført!"

    Exit Sub

ErrorHandler:
    Debug.Print "UNDO SILENT ERROR: " & Err.Description & " (celle " & i & ")"
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    undoStackSize = 0
End Sub

' Offentlig Undo for Planlegger-arket (MED melding)
Public Sub UndoPlanlegger()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Planlegger")

    If undoStackSize = 0 Then
        MsgBox "Ingen undo-tilstand tilgjengelig." & vbCrLf & vbCrLf & _
               "Undo fungerer bare etter makro-operasjoner som:" & vbCrLf & _
               "• Delete på aktivitetsceller" & vbCrLf & _
               "• Fjern markert område" & vbCrLf & _
               "• Legg inn aktivitet" & vbCrLf & _
               "• Rydd blokk for person", _
               vbExclamation, "Undo"
        Exit Sub
    End If

    Debug.Print "UNDO: Starter fra UndoPlanlegger..."
    Call Undo(ws)
    MsgBox "Undo utført! " & vbCrLf & vbCrLf & _
           "Siste makro-operasjon er angret.", _
           vbInformation, "Undo"
    Exit Sub

ErrorHandler:
    MsgBox "Feil under undo: " & Err.Description, vbCritical, "Undo feilet"
    Debug.Print "ERROR i UndoPlanlegger: " & Err.Description
End Sub

' Sjekk om undo er tilgjengelig
Public Function UndoTilgjengelig() As Boolean
    UndoTilgjengelig = (undoStackSize > 0)
End Function

' =====================================================
' CTRL+Z INTEGRATION
' =====================================================

' Smart Ctrl+Z handler som fungerer med både makro-operasjoner og manuelle endringer
Public Sub CtrlZ_Handler()
    On Error Resume Next

    ' SMART LOGIKK:
    ' - Hvis vi har custom undo tilgjengelig = makro kjørte nettopp → bruk custom undo
    ' - Hvis ikke = brukeren gjorde manuelle endringer → bruk Excel's undo

    If UndoTilgjengelig() Then
        ' Vi har custom undo snapshot (makro kjørte) → bruk den (STILLE - ingen melding)
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("Planlegger")
        Call UndoSilent(ws)
    Else
        ' Ingen custom undo → prøv Excel's innebygde undo for manuelle endringer
        On Error Resume Next
        Application.Undo
        If Err.Number <> 0 Then
            ' Excel's undo feilet også → ingen undo tilgjengelig
            ' (Ingen melding - Ctrl+Z skal være stille som i Excel)
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

' Initialiser Ctrl+Z til å bruke smart undo-handler
' Dette gjør at Ctrl+Z fungerer BÅDE for manuelle endringer OG makro-operasjoner
Public Sub InitializeCtrlZ()
    Application.OnKey "^z", "CtrlZ_Handler"
    MsgBox "Ctrl+Z er nå konfigurert!" & vbCrLf & vbCrLf & _
           "Smart undo-håndtering aktivert:" & vbCrLf & _
           "• Angrer makro-operasjoner (Delete, LeggInn, Rydd, etc.)" & vbCrLf & _
           "• Angrer også manuelle endringer (Excel standard)" & vbCrLf & vbCrLf & _
           "Bruk ResetCtrlZ() for å deaktivere.", _
           vbInformation, "Undo konfigurert"
End Sub

' Fjern Ctrl+Z override (tilbakestill til Excel's standard)
Public Sub ResetCtrlZ()
    Application.OnKey "^z"
    MsgBox "Ctrl+Z er tilbakestilt til Excel's standard oppførsel.", vbInformation
End Sub

' Auto-initialiser når Excel åpnes (VALGFRI - kan kalles manuelt)
Public Sub Auto_Open()
    ' Uncomment neste linje for å aktivere Ctrl+Z automatisk ved oppstart:
    ' InitializeCtrlZ
End Sub
