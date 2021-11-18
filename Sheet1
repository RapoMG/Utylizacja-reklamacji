Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next 'bez tego wysypuje się przy usówaniu wiersza
    If Target.Column = 7 And Target.Value > 0 Then
    Target.Cells(1, 2) = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm")
    ElseIf Target.Column = 7 And Target.Value = "" Then
    Target.Cells(1, 2) = Null
    End If
End Sub
