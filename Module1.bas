Dim Adres As String
Sub OddajPierscien()
Dim WierMaksTab As Integer, WierZap As Integer, Rekl As Integer
Dim Data_zwrotuT As String
Dim Data_pobrania As Date, Data_zwrotu As Date
Dim KopRek As Integer
Dim KopKod As String, KopNaz As String, KopPrzyj As String, KopDec As String, KopZw As String
Dim Plik As String, wynik As Integer
Dim Nowe As String

Adres = Range("C2").Value
Data_pobrania = Range("H2").Value

' Brak migotania - skoków między kartami
Application.ScreenUpdating = False

'Test adresu
If Adres = "" Or FileExists(Adres) = False Then
    Adres = ""
    Call Plik_scieazka
End If

If Adres = "" Then
    MsgBox ("Nie wybrano pliku reklamacji.")
    Exit Sub
End If
Workbooks.Open Filename:=Adres

wynik = InStrRev(Adres, "\")
Plik = Right(Adres, Len(Adres) - wynik)
'Ostatnia Reklamacje
'Workbooks(Plik).Sheets("TABELA").Activate
Workbooks(Plik).Sheets("TABELA").Select
Range("b5").Select

Do Until IsEmpty(ActiveCell)
    ActiveCell.Offset(2, 0).Select
Loop
WierMaksTab = ActiveCell.Offset(-3, 0).Row

'Nowy wiersz
ThisWorkbook.Activate
Sheets(1).Range("a3").Select
Do Until IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
Loop
WierZap = ActiveCell.Row

'Wypłaty
Workbooks(Plik).Activate
Sheets("TABELA").Range("M4").Select
Nowe = False

Dim JuzUtyl As Boolean

Do
    Data_zwrotuT = Right(ActiveCell.Value, 17)
   
    If Data_zwrotuT = "" Then Data_zwrotuT = 1
    Data_zwrotu = CDate(Data_zwrotuT)
    If Left(ActiveCell.Value, 5) = "ZWROT" And Data_zwrotu > Data_pobrania Then
        Rekl = ActiveCell.Row
        KopRek = Cells(Rekl, 1).Value
        KopKod = Cells(Rekl, 8).Value
        KopNaz = Cells(Rekl, 3).Value
        KopPrzyj = Cells(Rekl + 1, 2).Value
        KopDec = Right(Cells(Rekl, 12).Value, 17)
        KopZw = Right(Cells(Rekl, 13).Value, 17)
        '------------------------------------------
        If Cells(Rekl, 14).Value = "utylizacja" Or Cells(Rekl, 14).Value = "utylizacja" Then JuzUtyl = True
        '-------------------------------------------
        
        'Kopiowanie i wklejanie
        ThisWorkbook.Activate
        Sheets(1).Select
        Cells(WierZap, 1).FormulaR1C1 = KopRek
        Cells(WierZap, 2).FormulaR1C1 = KopNaz
        Cells(WierZap, 3).FormulaR1C1 = KopKod
        Cells(WierZap, 4).FormulaR1C1 = KopPrzyj
        Cells(WierZap, 5).FormulaR1C1 = KopDec
        Cells(WierZap, 6).FormulaR1C1 = KopZw
        
        If JuzUtyl = True Then Cells(WierZap, 7).FormulaR1C1 = "Tak"
        JuzUtyl = False
        If WierZap Mod 2 = 0 Then
        With Range(Cells(WierZap, 1), Cells(WierZap, 9)).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 16772085
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        End If
        
        WierZap = WierZap + 1
        
        'Powrót
        Workbooks(Plik).Activate
        Sheets(2).Cells(Rekl, 13).Select
        Nowe = True
   End If
ActiveCell.Offset(2, 0).Activate
Loop Until ActiveCell.Row = WierMaksTab



'Dodać znacznik pobierania na wypadek braku nowych reklamacji

' Data pobrania
ThisWorkbook.Activate
Range("H2").FormulaR1C1 = Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm")

If Nowe = False Then MsgBox ("Nie znaleziono nowych reklamacji.")

Windows(Plik).Close SaveChanges:=False
End Sub
Sub Plik_scieazka()
Dim plik_Rekl As Variant
plik_Rekl = Application.GetOpenFilename("Pliki Microsoft Excel,*.xlsm")
If plik_Rekl = False Then
    Exit Sub
Else
    Range("C2").FormulaR1C1 = plik_Rekl
End If
End Sub
Private Function FileExists(fname) As Boolean
FileExists = (Dir(fname) <> "")
End Function
