Attribute VB_Name = "Sheet_AutoDetect"
Option Explicit

'==========================================================
' AUTO CLEAR ROW RESULTS WHEN USER EDITS COLUMNS A TO D
' Paste this into: Sheet Module (Right-click sheet tab > View Code)
'==========================================================
Private Sub Worksheet_Change(ByVal Target As Range)

    Dim ws As Worksheet
    Dim changedRow As Long

    Set ws = Me
    changedRow = Target.Row

    '=== Only watch input columns A to D ===
    If Target.Column > 4 Then Exit Sub

    '=== Only watch data rows — not the header ===
    If changedRow < 2 Then Exit Sub

    '=== Only clear if results already exist ===
    If Trim(ws.Cells(changedRow, 5).Value) <> "" Then

        Application.EnableEvents = False
        ws.Range(ws.Cells(changedRow, 5), ws.Cells(changedRow, 11)).ClearContents
        Application.EnableEvents = True

        MsgBox "Row " & changedRow & " has been edited!" & vbCrLf & vbCrLf & _
               "Previous results cleared." & vbCrLf & _
               "Run the macro again to reanalyse this row.", _
               vbInformation, "Row Marked for Reprocessing"

    End If

End Sub