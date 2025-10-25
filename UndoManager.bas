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

    MsgBox "Undo utført!", vbInformation

    On Error GoTo 0
End Sub

' Offentlig Undo for Planlegger-arket
Public Sub UndoPlanlegger()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Planlegger")
    Call Undo(ws)
End Sub

' Sjekk om undo er tilgjengelig
Public Function UndoTilgjengelig() As Boolean
    UndoTilgjengelig = (undoStackSize > 0)
End Function
