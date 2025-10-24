Attribute VB_Name = "ArkAktivitetsOversikt"

' ========================================================================
' WORKSHEET EVENT: AKTIVITETSOVERSIKT
' Håndterer automatisk synkronisering til Planlegger når bruker endrer data
' ========================================================================

Private Const TBL_START_ROW As Long = 10
Private Const COL_PERSON As Long = 1
Private Const COL_KODE As Long = 2
Private Const COL_OPP_START As Long = 4
Private Const COL_OPP_SLUTT As Long = 5
Private Const COL_FORSINKET As Long = 6
Private Const COL_KOMMENTAR As Long = 10

' Lagre gammel person-verdi når bruker klikker i Person-kolonnen
Private gammelPersonVerdi As String
Private gammelPersonRad As Long

' =====================================================================
' EVENT: Når bruker klikker i en celle (før endring)
' =====================================================================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error Resume Next
    
    ' Hvis bruker klikker i Person-kolonnen, lagre gammel verdi
    If Not Target Is Nothing Then
        If Target.Column = COL_PERSON And Target.Row >= TBL_START_ROW Then
            If Target.Rows.Count = 1 And Target.Columns.Count = 1 Then
                gammelPersonVerdi = Trim$(CStr(Target.Value))
                gammelPersonRad = Target.Row
            End If
        End If
    End If
    
    On Error GoTo 0
End Sub

' =====================================================================
' EVENT: Når bruker endrer en celle
' =====================================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Slutt
    
    If Target Is Nothing Then Exit Sub
    If Target.Row < TBL_START_ROW Then Exit Sub
    
    ' Disable events for å unngå loop
    Application.EnableEvents = False
    
    Dim c As Range
    
    ' Gå gjennom hver endret celle
    For Each c In Target.Cells
        If c.Row >= TBL_START_ROW Then
            Select Case c.Column
                Case COL_PERSON
                    ' Person endret - flytt aktivitet til ny person
                    Call HåndterPersonEndring(c.Row)
                    
                Case COL_KODE
                    ' Kode endret - valider og oppdater beskrivelse
                    Call HåndterKodeEndring(c.Row)
                    
                Case COL_OPP_START, COL_OPP_SLUTT
                    ' Dato endret - valider
                    Call HåndterDatoEndring(c.Row)
            End Select
        End If
    Next c
    
Slutt:
    Application.EnableEvents = True
End Sub

' =====================================================================
' Håndter person-endring
' =====================================================================
Private Sub HåndterPersonEndring(ByVal rad As Long)
    Dim nyPerson As String, kode As String
    Dim startDato As Date, sluttDato As Date, kommentar As String
    
    ' Hent verdier fra raden
    nyPerson = Trim$(Me.Cells(rad, COL_PERSON).Value)
    kode = Trim$(Me.Cells(rad, COL_KODE).Value)
    
    ' Valider at vi har nødvendig data
    If Len(nyPerson) = 0 Or Len(kode) = 0 Then Exit Sub
    If Len(gammelPersonVerdi) = 0 Then Exit Sub
    If gammelPersonRad <> rad Then Exit Sub
    If gammelPersonVerdi = nyPerson Then Exit Sub ' Ingen endring
    
    ' Hent datoer
    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    kommentar = Trim$(Me.Cells(rad, COL_KOMMENTAR).Value)
    On Error GoTo 0
    
    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub
    
    ' Kall funksjonen som flytter aktivitet i Planlegger
    On Error Resume Next
    Application.Run "FlyttAktivitetTilNyPerson", _
                    gammelPersonVerdi, nyPerson, kode, _
                    startDato, sluttDato, kommentar
    On Error GoTo 0
    
    ' Nullstill gammel verdi
    gammelPersonVerdi = ""
    gammelPersonRad = 0
End Sub

' =====================================================================
' Håndter kode-endring
' =====================================================================
Private Sub HåndterKodeEndring(ByVal rad As Long)
    Dim nyKode As String, beskrivelse As String, farge As Long
    Dim wsTyp As Worksheet
    Dim funnet As Boolean
    
    nyKode = UCase$(Trim$(Me.Cells(rad, COL_KODE).Value))
    If Len(nyKode) = 0 Then Exit Sub
    
    ' Finn aktivitetstyper-arket
    On Error Resume Next
    Set wsTyp = ThisWorkbook.Worksheets("AKTIVITETSTYPER - OVERSIKT")
    On Error GoTo 0
    
    If wsTyp Is Nothing Then Exit Sub
    
    ' Valider at koden finnes
    On Error Resume Next
    funnet = Application.Run("LookupAktivitet", wsTyp, nyKode, beskrivelse, farge)
    On Error GoTo 0
    
    If Not funnet Then
        MsgBox "Aktivitetskode '" & nyKode & "' finnes ikke i AKTIVITETSTYPER - OVERSIKT!" & vbCrLf & vbCrLf & _
               "Legg til koden først, eller velg en eksisterende kode.", vbExclamation, "Ugyldig kode"
        Me.Cells(rad, COL_KODE).Value = ""
    Else
        ' Oppdater beskrivelsen automatisk
        Me.Cells(rad, COL_KODE + 1).Value = beskrivelse
    End If
End Sub

' =====================================================================
' Håndter dato-endring
' =====================================================================
Private Sub HåndterDatoEndring(ByVal rad As Long)
    Dim startDato As Date, sluttDato As Date
    
    On Error Resume Next
    startDato = Me.Cells(rad, COL_OPP_START).Value
    sluttDato = Me.Cells(rad, COL_OPP_SLUTT).Value
    On Error GoTo 0
    
    If Not IsDate(startDato) Or Not IsDate(sluttDato) Then Exit Sub
    
    ' Valider at startdato ikke er etter sluttdato
    If startDato > sluttDato Then
        MsgBox "Startdato kan ikke være etter sluttdato!", vbExclamation, "Ugyldig dato"
        Me.Cells(rad, COL_OPP_START).Value = sluttDato
    End If
End Sub
