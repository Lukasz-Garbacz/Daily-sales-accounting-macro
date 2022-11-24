Attribute VB_Name = "Module1"
Sub RK_FKJ_Makro()
Attribute RK_FKJ_Makro.VB_Description = "dane do RK"
Attribute RK_FKJ_Makro.VB_ProcData.VB_Invoke_Func = "r\n14"
Dim myFile              As String
Dim rng                 As Range
Dim cellValue           As Variant
Dim i                   As Integer
Dim j                   As Integer
Dim data                As String
Dim rap_nr              As Integer
Dim iter_wplat          As Integer
Dim iter_wyplat         As Integer
Dim WersjaProgramu      As Integer
Dim Okres               As Integer
Dim iter_wersji         As Integer
Dim KontoKasy           As Integer
Dim iter_petli          As Integer
Dim numbers             As Range

myFile = Application.ActiveWorkbook.FullName & ".txt"
Set rng = Application.Range(Cell1:="B6", Cell2:="E55")
Set numbers = Application.Range(Cell1:="A6", Cell2:="A55")
data = Application.Cells(1, 3)
rap_nr = Application.Cells(2, 3)
Open myFile For Output As #1




KontoKasy = 100
Okres = 30286
WersjaProgramu = 219
Print #1, "INFO{"
Print #1, vbTab; "Nazwa programu ='Sage Symfonia 2.0 Handel 2019.c' Symfonia 2.0 Handel 2019.c"
Print #1, vbTab; "Wersja_programu =" & WersjaProgramu
Print #1, vbTab; "Wersja szablonu ="
Print #1, vbTab; "dane_z_oddzialu ="
Print #1, vbTab; "Kontrahent{"
Print #1, vbTab; vbTab; "id ="
Print #1, vbTab; vbTab; "kod ="
Print #1, vbTab; vbTab; "nazwa ="
Print #1, vbTab; vbTab; "nip ="
Print #1, vbTab; "}"
Print #1, "}"






iter_wplat = 1
iter_wyplat = 1
For i = 1 To rng.Rows.Count

    If rng.Cells(i, 2).Value <> 0 And rng.Cells(i, 3) = 0 Then
        Print #1, "Z oddzia³u. Dok. pieniê¿ny{"
        Print #1, vbTab; "Notatka_Dl{"
        Print #1, vbTab; vbTab; "opis ="
        Print #1, vbTab; "}"
        Print #1, vbTab; "rodzaj_dok =pieniê¿ny"
        Print #1, vbTab; "id =" & WersjaProgramu + iter_wersji
        Print #1, vbTab; "flag =0"
        Print #1, vbTab; "typ =2"
        Print #1, vbTab; "pusty =0"
        Print #1, vbTab; "rejestr =130"
        Print #1, vbTab; "znaczniki =0"
        Print #1, vbTab; "osoba =Admin"
        Print #1, vbTab; "plattypi =0"
        Print #1, vbTab; "typdk =KP"
        Print #1, vbTab; "seria =sKP"
        Print #1, vbTab; "serianr =" & iter_wplat
        Print #1, vbTab; "okres =30286"
        Print #1, vbTab; "data =" & data
        Print #1, vbTab; "datarozl ="
        Print #1, vbTab; "termin =" & data
        Print #1, vbTab; "dkid =0"
        Print #1, vbTab; "opis =" & rng.Cells(i, 1).Value
        Print #1, vbTab; "khid =0"
        Print #1, vbTab; "khkod ="
        Print #1, vbTab; "kwota =" & rng.Cells(i, 2).Value
        Print #1, vbTab; "wyplatai =0"
        Print #1, vbTab; "kwotarozl =0"
        Print #1, vbTab; "stan =0"
        Print #1, vbTab; "typkhi =0"
        Print #1, vbTab; "exp_fki =0"
        Print #1, vbTab; "dzial =0"
        Print #1, vbTab; "subtypi =60"
        
        If numbers.Cells(i, 1) = 1 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) < 5 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) < 9 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 9 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) < 12 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) < 40 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 40 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 41 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 42 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 43 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 44 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 45 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 46 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 47 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) < 50 Then
        Print #1, vbTab; "schemat ="
        ElseIf numbers.Cells(i, 1) = 50 Then
        Print #1, vbTab; "schemat ="
        End If
        
        Print #1, vbTab; "waluta ="
        Print #1, vbTab; "kurs =1"
        Print #1, vbTab; "kwotawal=" & rng.Cells(i, 2).Value
        Print #1, vbTab; "kwotarozlwal =0"
        Print #1, vbTab; "e_status =0"
        Print #1, vbTab; "guid ="
        Print #1, vbTab; "rodzajpn =0"
        Print #1, vbTab; "zapas ="
        Print #1, vbTab; "typi =2"
        Print #1, vbTab; "rejestr_platnosci =KASA"
        Print #1, "}"
        
        
    iter_wplat = iter_wplat + 1
    iter_wersji = iter_wersji + 1
    End If
    
    
    If rng.Cells(i, 3).Value <> 0 And rng.Cells(i, 2) = 0 Then
        Print #1, "Z oddzia³u. Dok. pieniê¿ny{"
        Print #1, vbTab; "Notatka_Dl{"
        Print #1, vbTab; vbTab; "opis ="
        Print #1, vbTab; "}"
        Print #1, vbTab; "rodzaj_dok =pieniê¿ny"
        Print #1, vbTab; "id =" & WersjaProgramu + iter_wersji
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
        Print #1, vbTab; "opis =" & rng.Cells(i, 1).Value
        Print #1, vbTab; "khid =0"
        Print #1, vbTab; "khkod ="
        Print #1, vbTab; "kwota =" & -rng.Cells(i, 3).Value
        Print #1, vbTab; "wyplatai =1"
        Print #1, vbTab; "kwotarozl =0"
        Print #1, vbTab; "stan =0"
        Print #1, vbTab; "typkhi =0"
        Print #1, vbTab; "exp_fki =0"
        Print #1, vbTab; "dzial =0"
        Print #1, vbTab; "subtypi =61"
        
        If iter_wersji = 0 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji < 3 And iter_wersji <> 0 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji < 5 And iter_wersji > 2 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji = 5 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji = 6 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji < 35 And iter_wersji > 6 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji = 35 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji = 36 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji = 37 Then
        Print #1, vbTab; "schemat ="
        ElseIf iter_wersji = 38 Then
        Print #1, vbTab; "schemat ="
        End If
        
        Print #1, vbTab; "waluta ="
        Print #1, vbTab; "kurs =1"
        Print #1, vbTab; "kwotawal=" & -rng.Cells(i, 3).Value
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
    End If
    
Next i
    
    Print #1, "Dokument{"
    Print #1, vbTab; "symbol FK =RK"
    Print #1, vbTab; "kod =" & rap_nr
    Print #1, vbTab; "opis =rejestr KASA za dzieñ " & data
    Print #1, vbTab; "data =" & data
    Print #1, vbTab; "datasp =" & data
    Print #1, vbTab; "kwota =" & WorksheetFunction.Sum(rng.Range(Cells(1, 2), Cells(rng.Rows.Count, 3)))
    Print #1, vbTab; "SaldoPRK =0.00"
    Print #1, vbTab; "SaldoZRK =0.00"
    Print #1, vbTab; "Sygnatura =Admin"
    Print #1, vbTab; "KontoKasy =100"
    Print #1, vbTab; "obsluguj jak =RK"
    Print #1, vbTab; "FK nazwa =" & rap_nr
    Print #1, vbTab; "opis FK =rejestr KASA za dzieñ " & data
    


iter_wersji = 1
iter_wplat = 1
iter_wyplat = 1
For i = 1 To rng.Rows.Count
    
    If rng.Cells(i, 2).Value <> 0 And rng.Cells(i, 3) = 0 Then
    
        Print #1, vbTab; "Zapis{"
        Print #1, vbTab; vbTab; "strona =WN"
        Print #1, vbTab; vbTab; "kwota =" & rng.Cells(i, 2).Value + rng.Cells(i, 3).Value
        Print #1, vbTab; vbTab; "konto =" & KontoKasy
        Print #1, vbTab; vbTab; "IdDlaRozliczen =" & iter_wersji
        Print #1, vbTab; vbTab; "opis =" & rng.Cells(i, 1).Value
    
            If iter_wplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/000" & iter_wplat; "/KP"
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/00" & iter_wplat; "/KP"
            End If

        Print #1, vbTab; vbTab; "Pozycja =" & iter_petli
        Print #1, vbTab; vbTab; "ZapisRownolegly =0"
        Print #1, vbTab; vbTab; "dataKPKW =" & data
        Print #1, vbTab; "}"

        Print #1, vbTab; "Zapis{"
        Print #1, vbTab; vbTab; "strona =MA"
        Print #1, vbTab; vbTab; "kwota =" & rng.Cells(i, 2).Value + rng.Cells(i, 3).Value
        Print #1, vbTab; vbTab; "konto =" & rng.Cells(i, 4).Value
        Print #1, vbTab; vbTab; "IdDlaRozliczen =" & iter_wersji + 1
        Print #1, vbTab; vbTab; "opis =" & rng.Cells(i, 1).Value
    
            If iter_wplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/000" & iter_wplat; "/KP"
                iter_wplat = iter_wplat + 1
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/00" & iter_wplat; "/KP"
                iter_wplat = iter_wplat + 1
            End If
    
        Print #1, vbTab; vbTab; "Pozycja =" & iter_petli
        Print #1, vbTab; vbTab; "ZapisRownolegly =0"
        Print #1, vbTab; vbTab; "dataKPKW =" & data
        Print #1, vbTab; "}"
        iter_wersji = iter_wersji + 2
 
     iter_petli = iter_petli + 1
    End If
    
    
     If rng.Cells(i, 3).Value <> 0 And rng.Cells(i, 2) = 0 Then
    
        Print #1, vbTab; "Zapis{"
        Print #1, vbTab; vbTab; "strona =WN"
        Print #1, vbTab; vbTab; "kwota =" & rng.Cells(i, 2).Value + rng.Cells(i, 3).Value
        Print #1, vbTab; vbTab; "konto =" & rng.Cells(i, 4).Value
        Print #1, vbTab; vbTab; "IdDlaRozliczen =" & iter_wersji
        Print #1, vbTab; vbTab; "opis =" & rng.Cells(i, 1).Value
    
            If iter_wyplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/000" & iter_wyplat; "/KW"
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/00" & iter_wyplat; "/KW"
            End If

        Print #1, vbTab; vbTab; "Pozycja =" & iter_petli
        Print #1, vbTab; vbTab; "ZapisRownolegly =0"
        Print #1, vbTab; vbTab; "dataKPKW =" & data
        Print #1, vbTab; "}"

        Print #1, vbTab; "Zapis{"
        Print #1, vbTab; vbTab; "strona =MA"
        Print #1, vbTab; vbTab; "kwota =" & rng.Cells(i, 2).Value + rng.Cells(i, 3).Value
        Print #1, vbTab; vbTab; "konto =" & KontoKasy
        Print #1, vbTab; vbTab; "IdDlaRozliczen =" & iter_wersji + 1
        Print #1, vbTab; vbTab; "opis =" & rng.Cells(i, 1).Value
    
            If iter_wyplat < 10 Then
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/000" & iter_wyplat; "/KW"
                iter_wyplat = iter_wyplat + 1
            Else
                Print #1, vbTab; vbTab; "NumerDok =" & Mid(data, 3, 5) & "/00" & iter_wyplat; "/KW"
                iter_wyplat = iter_wyplat + 1
            End If
    
        Print #1, vbTab; vbTab; "Pozycja =" & iter_petli
        Print #1, vbTab; vbTab; "ZapisRownolegly =0"
        Print #1, vbTab; vbTab; "dataKPKW =" & data
        Print #1, vbTab; "}"
        iter_wersji = iter_wersji + 2
 
     iter_petli = iter_petli + 1
    End If


Next i
Print #1, "}"
Close #1
End Sub
