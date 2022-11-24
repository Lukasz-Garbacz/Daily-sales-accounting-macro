Attribute VB_Name = "Module3"
Sub Naglowek()

Print #1, "INFO{"
Print #1, vbTab; "Nazwa programu ='Sage Symfonia 2.0 Handel 2019.c' Symfonia 2.0 Handel 2019.c"
Print #1, vbTab; "Wersja_programu ="
Print #1, vbTab; "Wersja szablonu ="
Print #1, vbTab; "dane_z_oddzialu ="
Print #1, vbTab; "Kontrahent{"
Print #1, vbTab; vbTab; "id ="
Print #1, vbTab; vbTab; "kod ="
Print #1, vbTab; vbTab; "nazwa ="
Print #1, vbTab; vbTab; "nip ="
Print #1, vbTab; "}"
Print #1, "}"

End Sub


'find a row containing given string assuming its in the first column of rng, return as int
Function find_row(str As String) As Integer
Dim iter1 As Integer

For iter1 = 1 To rng.Rows.Count
    If InStr(1, rng.Cells(iter1, 1), str, vbTextCompare) > 0 Then
        find_row = iter1
        Exit Function
    End If
Next iter1
find_row = -1
End Function


'remove all clutter from value cell from given row
Function fix_str(temp_str As String) As String
temp_str = Replace(temp_str, " ", "")
temp_str = Replace(temp_str, "PLN", "")
temp_str = Replace(temp_str, "(", "")
temp_str = Replace(temp_str, ")", "")
temp_str = Replace(temp_str, "-", "")
temp_str = Replace(temp_str, "+", "")
fix_str = temp_str
End Function


'True if negative, false if positive or 0
Function check_neg(row_num As Integer) As Boolean
If InStr(rng.Cells(row_num, 2).Value, "(") <> 0 Then
    check_neg = True
Else
    check_neg = False
    End If

End Function


'if sign_overr = 0 then do nothing; if 1 then force -; if 2 then force +
Sub Dok_pien(nazwa As String, schemat As String, sign_overr As Integer)
        
        Dim row_num As Integer
        Dim wyplata As Boolean
        
        row_num = find_row(nazwa)
        If row_num = -1 Then
            Exit Sub
        End If
        
        wyplata = check_neg(row_num)
        
        If sign_overr = 1 Then
            wyplata = True
        ElseIf sign_overr = 2 Then
            wyplata = False
        End If
        
        Print #1, "Z oddzia³u. Dok. pieniê¿ny{"
        Print #1, vbTab; "Notatka_Dl{"
        Print #1, vbTab; vbTab; "opis ="
        Print #1, vbTab; "}"
        Print #1, vbTab; "rodzaj_dok =pieniê¿ny"
        Print #1, vbTab; "id =" & + iter_wersji
        Print #1, vbTab; "flag =0"
        Print #1, vbTab; "typ =2"
        Print #1, vbTab; "pusty =0"
        Print #1, vbTab; "rejestr ="
        Print #1, vbTab; "znaczniki =0"
        Print #1, vbTab; "osoba =Admin"
        Print #1, vbTab; "plattypi =0"
        
        If wyplata Then
            Print #1, vbTab; "typdk ="
            Print #1, vbTab; "seria ="
            Print #1, vbTab; "serianr =" & iter_wyplat
        Else
            Print #1, vbTab; "typdk ="
            Print #1, vbTab; "seria ="
            Print #1, vbTab; "serianr =" & iter_wplat
        End If
        
        Print #1, vbTab; "okres =30286"
        Print #1, vbTab; "data =" & data
        Print #1, vbTab; "datarozl ="
        Print #1, vbTab; "termin =" & data
        Print #1, vbTab; "dkid =0"
        Print #1, vbTab; "opis =" & rng.Cells(row_num, 1).Value
        Print #1, vbTab; "khid =0"
        Print #1, vbTab; "khkod ="
        
        If wyplata Then
            Print #1, vbTab; "kwota =-" & fix_str(rng.Cells(row_num, 2).Value)
            Print #1, vbTab; "wyplatai =1"
        Else
            Print #1, vbTab; "kwota =" & fix_str(rng.Cells(row_num, 2).Value)
            Print #1, vbTab; "wyplatai =0"
        End If
        
        Print #1, vbTab; "kwotarozl =0"
        Print #1, vbTab; "stan =0"
        Print #1, vbTab; "typkhi =0"
        Print #1, vbTab; "exp_fki =0"
        Print #1, vbTab; "dzial =0"
        
        If wyplata Then
            Print #1, vbTab; "subtypi ="
        Else
            Print #1, vbTab; "subtypi ="
        End If
        
        Print #1, vbTab; "schemat =" & schemat
        Print #1, vbTab; "waluta ="
        Print #1, vbTab; "kurs =1"
        
        If wyplata Then
            Print #1, vbTab; "kwotawal=-" & fix_str(rng.Cells(row_num, 2).Value)
        Else
            Print #1, vbTab; "kwotawal=" & fix_str(rng.Cells(row_num, 2).Value)
        End If
        
        Print #1, vbTab; "kwotarozlwal =0"
        Print #1, vbTab; "e_status =0"
        Print #1, vbTab; "guid ="
        Print #1, vbTab; "rodzajpn =0"
        Print #1, vbTab; "zapas ="
        Print #1, vbTab; "typi =2"
        Print #1, vbTab; "rejestr_platnosci =KASA"
        Print #1, "}"
        
    If wyplata Then
        iter_wyplat = iter_wyplat + 1
    Else
        iter_wplat = iter_wplat + 1
    End If
    
    iter_wersji = iter_wersji + 1

End Sub


Sub Naglowek2()

    Dim row_num As Integer
    row_num = find_row("Dochód (+)") - 2

    Print #1, "Dokument{"
    Print #1, vbTab; "symbol FK ="
    Print #1, vbTab; "kod =" & rap_nr
    Print #1, vbTab; "opis =rejestr KASA za dzieñ " & data
    Print #1, vbTab; "data =" & data
    Print #1, vbTab; "datasp =" & data
    Print #1, vbTab; "kwota =" & fix_str(rng.Cells(row_num, 2).Value)
    Print #1, vbTab; "SaldoPRK =0.00"
    Print #1, vbTab; "SaldoZRK =0.00"
    Print #1, vbTab; "Sygnatura =Admin"
    Print #1, vbTab; "KontoKasy =100"
    Print #1, vbTab; "obsluguj jak ="
    Print #1, vbTab; "FK nazwa =" & rap_nr
    Print #1, vbTab; "opis FK =rejestr KASA za dzieñ " & data

End Sub


Sub Zapis_part(strona As String, konto As String, row_num As Integer, wyplata As Boolean)

        Print #1, vbTab; "Zapis{"
        Print #1, vbTab; vbTab; "strona =" & strona
        Print #1, vbTab; vbTab; "kwota =" & fix_str(rng.Cells(row_num, 2).Value)
        Print #1, vbTab; vbTab; "konto =" & konto
        Print #1, vbTab; vbTab; "IdDlaRozliczen =" & iter_wersji
        Print #1, vbTab; vbTab; "opis =" & rng.Cells(row_num, 1).Value
    
        If wyplata Then
            If iter_wyplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Right(data, 2) & "-" & Mid(data, 4, 2) & "/000" & iter_wyplat; "/KW"
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Right(data, 2) & "-" & Mid(data, 4, 2) & "/00" & iter_wyplat; "/KW"
            End If
        Else
            If iter_wplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Right(data, 2) & "-" & Mid(data, 4, 2) & "/000" & iter_wplat; "/KP"
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Right(data, 2) & "-" & Mid(data, 4, 2) & "/00" & iter_wplat; "/KP"
            End If
        End If
        

        Print #1, vbTab; vbTab; "Pozycja =" & iter_petli
        Print #1, vbTab; vbTab; "ZapisRownolegly =0"
        Print #1, vbTab; vbTab; "dataKPKW =" & data
        Print #1, vbTab; "}"

End Sub


Sub Zapis(nazwa As String, konto As String, sign_overr As Integer)

    Dim row_num As Integer
    Dim wyplata As Boolean
        
    row_num = find_row(nazwa)
    If row_num = -1 Then
        Exit Sub
    End If
        
    wyplata = check_neg(row_num)
        
    If sign_overr = 1 Then
        wyplata = True
    ElseIf sign_overr = 2 Then
        wyplata = False
    End If

If wyplata Then
    Call Zapis_part("WN", konto, row_num, wyplata)
    iter_wersji = iter_wersji + 1
    Call Zapis_part("MA", 100, row_num, wyplata)
    iter_wersji = iter_wersji + 1
    iter_wplat = iter_wyplat + 1
    iter_petli = iter_petli + 1
Else
    Call Zapis_part("WN", "100", row_num, wyplata)
    iter_wersji = iter_wersji + 1
    Call Zapis_part("MA", konto, row_num, wyplata)
    iter_wersji = iter_wersji + 1
    iter_wplat = iter_wplat + 1
    iter_petli = iter_petli + 1
End If

End Sub



Sub Dok_pien_wyplaty(kwota As String, opis As String, schemat As String)
        
        Dim row_num As Integer
        Dim wyplata As Boolean
        wyplata = True
        
        Print #1, "Z oddzia³u. Dok. pieniê¿ny{"
        Print #1, vbTab; "Notatka_Dl{"
        Print #1, vbTab; vbTab; "opis ="
        Print #1, vbTab; "}"
        Print #1, vbTab; "rodzaj_dok =pieniê¿ny"
        Print #1, vbTab; "id =" & 219 + iter_wersji
        Print #1, vbTab; "flag =0"
        Print #1, vbTab; "typ =2"
        Print #1, vbTab; "pusty =0"
        Print #1, vbTab; "rejestr =130"
        Print #1, vbTab; "znaczniki =0"
        Print #1, vbTab; "osoba =Admin"
        Print #1, vbTab; "plattypi =0"
        
        Print #1, vbTab; "typdk =KW"
        Print #1, vbTab; "seria =sKW"
        Print #1, vbTab; "serianr =" & iter_wyplat
        
        Print #1, vbTab; "okres =30286"
        Print #1, vbTab; "data =" & data
        Print #1, vbTab; "datarozl ="
        Print #1, vbTab; "termin =" & data
        Print #1, vbTab; "dkid =0"
        Print #1, vbTab; "opis =" & opis
        Print #1, vbTab; "khid =0"
        Print #1, vbTab; "khkod ="
        
        Print #1, vbTab; "kwota =-" & fix_str(kwota)
        Print #1, vbTab; "wyplatai =1"
        
        Print #1, vbTab; "kwotarozl =0"
        Print #1, vbTab; "stan =0"
        Print #1, vbTab; "typkhi =0"
        Print #1, vbTab; "exp_fki =0"
        Print #1, vbTab; "dzial =0"
        
        Print #1, vbTab; "subtypi =61"
        
        Print #1, vbTab; "schemat =" & schemat
        Print #1, vbTab; "waluta ="
        Print #1, vbTab; "kurs =1"
        
        Print #1, vbTab; "kwotawal=-" & fix_str(kwota)
        
        Print #1, vbTab; "kwotarozlwal =0"
        Print #1, vbTab; "e_status =0"
        Print #1, vbTab; "guid ="
        Print #1, vbTab; "rodzajpn =0"
        Print #1, vbTab; "zapas ="
        Print #1, vbTab; "typi =2"
        Print #1, vbTab; "rejestr_platnosci =KASA"
        Print #1, "}"
        
        iter_wyplat = iter_wyplat + 1
        iter_wersji = iter_wersji + 1

End Sub


Sub Zapis_part_wyplaty(strona As String, konto As String, opis As String, kwota As String, wyplata As Boolean)

        Print #1, vbTab; "Zapis{"
        Print #1, vbTab; vbTab; "strona =" & strona
        Print #1, vbTab; vbTab; "kwota =" & fix_str(kwota)
        Print #1, vbTab; vbTab; "konto =" & konto
        Print #1, vbTab; vbTab; "IdDlaRozliczen =" & iter_wersji
        Print #1, vbTab; vbTab; "opis =" & opis
    
            If iter_wyplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Right(data, 2) & "-" & Mid(data, 4, 2) & "/000" & iter_wyplat; "/KW"
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Right(data, 2) & "-" & Mid(data, 4, 2) & "/00" & iter_wyplat; "/KW"
            End If
            
        Print #1, vbTab; vbTab; "Pozycja =" & iter_petli
        Print #1, vbTab; vbTab; "ZapisRownolegly =0"
        Print #1, vbTab; vbTab; "dataKPKW =" & data
        Print #1, vbTab; "}"

End Sub


Sub Zapis_wyplaty(opis As String, kwota As String)

    Dim wyplata As Boolean
    wyplata = True

    Call Zapis_part_wyplaty("WN", "202-2-1-", opis, kwota, wyplata)
    iter_wersji = iter_wersji + 1
    Call Zapis_part_wyplaty("MA", 100, opis, kwota, wyplata)
    iter_wersji = iter_wersji + 1
    iter_wyplat = iter_wyplat + 1
    iter_petli = iter_petli + 1


End Sub

Sub OpenFile1()
Dim fNameAndPath    As Variant
Dim iter1           As Integer
iter1 = 7

fNameAndPath = Application.GetOpenFilename(FileFilter:="Excel Files (*.XLS*), *.XLS*", Title:="Wybierz plik wyplaty")
If fNameAndPath = False Then End
Set wb2 = Workbooks.Open(fNameAndPath)

If Right(wb2.Worksheets(1).Cells(3, 1).Value, 10) = data Then
    While IsEmpty(wb2.Worksheets(1).Cells(iter1, 11).Value) = False
        Call Dok_pien_wyplaty(wb2.Worksheets(1).Cells(iter1, 6).Value, wb2.Worksheets(1).Cells(iter1, 11).Value, "BP")
        iter1 = iter1 + 1
    Wend
Else
    MsgBox ("daty sie nie pokrywaja")
End If

End Sub

Sub OpenFile2()
Dim fNameAndPath    As Variant
Dim iter1           As Integer
iter1 = 7

If Right(wb2.Worksheets(1).Cells(3, 1).Value, 10) = data Then
    While IsEmpty(wb2.Worksheets(1).Cells(iter1, 11).Value) = False
        Call Zapis_wyplaty(wb2.Worksheets(1).Cells(iter1, 11).Value, wb2.Worksheets(1).Cells(iter1, 6).Value)
        iter1 = iter1 + 1
    Wend

End If
wb2.Close savechanges:=False
End Sub



