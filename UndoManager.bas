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
    FontBold As Boolean
    FontColor As Long
    HorizontalAlignment As Long
    VerticalAlignment As Long
    WrapText As Boolean
    HasComment As Boolean
    CommentText As String
    ' Borders (lagres separat)
    BorderLeftStyle As Long
    BorderLeftWeight As Long
    BorderLeftColor As Long
    BorderRightStyle As Long
    BorderRightWeight As Long
    BorderRightColor As Long
    BorderTopStyle As Long
    BorderTopWeight As Long
    BorderTopColor As Long
    BorderBottomStyle As Long
    BorderBottomWeight As Long
    BorderBottomColor As Long
End Type

Private undoStack() As CellState
Private undoStackSize As Long
Private Const MAX_UNDO_LEVELS As Long = 1  ' Bare én undo-nivå (siste operasjon)

' =====================================================
' PUBLIC API
' =====================================================

' Lagre tilstanden til et område før endring
Public Sub LagUndoSnapshot(ByVal rng As Range)
    On Error Resume Next

    ' Reset stack
    undoStackSize = 0
    ReDim undoStack(1 To rng.Cells.Count)

    Dim cel As Range
    Dim i As Long
    i = 1

    For Each cel In rng.Cells
        undoStack(i).Address = cel.Address
        undoStack(i).Value = cel.Value
        undoStack(i).Formula = cel.Formula

        ' Interior
        undoStack(i).InteriorColor = cel.Interior.Color
        undoStack(i).InteriorColorIndex = cel.Interior.ColorIndex

        ' Font
        undoStack(i).FontBold = cel.Font.Bold
        undoStack(i).FontColor = cel.Font.Color

        ' Alignment
        undoStack(i).HorizontalAlignment = cel.HorizontalAlignment
        undoStack(i).VerticalAlignment = cel.VerticalAlignment
        undoStack(i).WrapText = cel.WrapText

        ' Comment
        undoStack(i).HasComment = Not cel.Comment Is Nothing
        If undoStack(i).HasComment Then
            undoStack(i).CommentText = cel.Comment.Text
        End If

        ' Borders
        undoStack(i).BorderLeftStyle = cel.Borders(xlEdgeLeft).LineStyle
        undoStack(i).BorderLeftWeight = cel.Borders(xlEdgeLeft).Weight
        undoStack(i).BorderLeftColor = cel.Borders(xlEdgeLeft).Color

        undoStack(i).BorderRightStyle = cel.Borders(xlEdgeRight).LineStyle
        undoStack(i).BorderRightWeight = cel.Borders(xlEdgeRight).Weight
        undoStack(i).BorderRightColor = cel.Borders(xlEdgeRight).Color

        undoStack(i).BorderTopStyle = cel.Borders(xlEdgeTop).LineStyle
        undoStack(i).BorderTopWeight = cel.Borders(xlEdgeTop).Weight
        undoStack(i).BorderTopColor = cel.Borders(xlEdgeTop).Color

        undoStack(i).BorderBottomStyle = cel.Borders(xlEdgeBottom).LineStyle
        undoStack(i).BorderBottomWeight = cel.Borders(xlEdgeBottom).Weight
        undoStack(i).BorderBottomColor = cel.Borders(xlEdgeBottom).Color

        i = i + 1
    Next cel

    undoStackSize = rng.Cells.Count

    On Error GoTo 0
End Sub

' Gjenopprett siste lagrede tilstand (UNDO)
Public Sub Undo(ByVal ws As Worksheet)
    If undoStackSize = 0 Then
        MsgBox "Ingen undo-tilstand tilgjengelig.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next

    Dim i As Long
    Dim cel As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 1 To undoStackSize
        Set cel = ws.Range(undoStack(i).Address)

        ' Restore value/formula
        If Len(undoStack(i).Formula) > 0 Then
            cel.Formula = undoStack(i).Formula
        Else
            cel.Value = undoStack(i).Value
        End If

        ' Restore interior
        If undoStack(i).InteriorColorIndex = xlColorIndexNone Then
            cel.Interior.ColorIndex = xlColorIndexNone
        Else
            cel.Interior.Color = undoStack(i).InteriorColor
        End If

        ' Restore font
        cel.Font.Bold = undoStack(i).FontBold
        cel.Font.Color = undoStack(i).FontColor

        ' Restore alignment
        cel.HorizontalAlignment = undoStack(i).HorizontalAlignment
        cel.VerticalAlignment = undoStack(i).VerticalAlignment
        cel.WrapText = undoStack(i).WrapText

        ' Restore comment
        cel.ClearComments
        If undoStack(i).HasComment Then
            cel.AddComment undoStack(i).CommentText
        End If

        ' Restore borders
        With cel.Borders(xlEdgeLeft)
            .LineStyle = undoStack(i).BorderLeftStyle
            .Weight = undoStack(i).BorderLeftWeight
            .Color = undoStack(i).BorderLeftColor
        End With

        With cel.Borders(xlEdgeRight)
            .LineStyle = undoStack(i).BorderRightStyle
            .Weight = undoStack(i).BorderRightWeight
            .Color = undoStack(i).BorderRightColor
        End With

        With cel.Borders(xlEdgeTop)
            .LineStyle = undoStack(i).BorderTopStyle
            .Weight = undoStack(i).BorderTopWeight
            .Color = undoStack(i).BorderTopColor
        End With

        With cel.Borders(xlEdgeBottom)
            .LineStyle = undoStack(i).BorderBottomStyle
            .Weight = undoStack(i).BorderBottomWeight
            .Color = undoStack(i).BorderBottomColor
        End With
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Clear stack after undo
    undoStackSize = 0

    On Error GoTo 0
End Sub

' Gjenopprett siste lagrede tilstand (STILLE - ingen melding)
Private Sub UndoSilent(ByVal ws As Worksheet)
    If undoStackSize = 0 Then Exit Sub

    On Error Resume Next

    Dim i As Long
    Dim cel As Range

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 1 To undoStackSize
        Set cel = ws.Range(undoStack(i).Address)

        ' Restore value/formula
        If Len(undoStack(i).Formula) > 0 Then
            cel.Formula = undoStack(i).Formula
        Else
            cel.Value = undoStack(i).Value
        End If

        ' Restore interior
        If undoStack(i).InteriorColorIndex = xlColorIndexNone Then
            cel.Interior.ColorIndex = xlColorIndexNone
        Else
            cel.Interior.Color = undoStack(i).InteriorColor
        End If

        ' Restore font
        cel.Font.Bold = undoStack(i).FontBold
        cel.Font.Color = undoStack(i).FontColor

        ' Restore alignment
        cel.HorizontalAlignment = undoStack(i).HorizontalAlignment
        cel.VerticalAlignment = undoStack(i).VerticalAlignment
        cel.WrapText = undoStack(i).WrapText

        ' Restore comment
        cel.ClearComments
        If undoStack(i).HasComment Then
            cel.AddComment undoStack(i).CommentText
        End If

        ' Restore borders
        With cel.Borders(xlEdgeLeft)
            .LineStyle = undoStack(i).BorderLeftStyle
            .Weight = undoStack(i).BorderLeftWeight
            .Color = undoStack(i).BorderLeftColor
        End With

        With cel.Borders(xlEdgeRight)
            .LineStyle = undoStack(i).BorderRightStyle
            .Weight = undoStack(i).BorderRightWeight
            .Color = undoStack(i).BorderRightColor
        End With

        With cel.Borders(xlEdgeTop)
            .LineStyle = undoStack(i).BorderTopStyle
            .Weight = undoStack(i).BorderTopWeight
            .Color = undoStack(i).BorderTopColor
        End With

        With cel.Borders(xlEdgeBottom)
            .LineStyle = undoStack(i).BorderBottomStyle
            .Weight = undoStack(i).BorderBottomWeight
            .Color = undoStack(i).BorderBottomColor
        End With
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' Clear stack after undo
    undoStackSize = 0

    On Error GoTo 0
End Sub

' Offentlig Undo for Planlegger-arket (MED melding)
Public Sub UndoPlanlegger()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Planlegger")

    If undoStackSize = 0 Then
        MsgBox "Ingen undo-tilstand tilgjengelig.", vbExclamation
        Exit Sub
    End If

    Call Undo(ws)
    MsgBox "Undo utført!", vbInformation
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
