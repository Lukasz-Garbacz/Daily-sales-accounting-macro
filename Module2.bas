Attribute VB_Name = "Module2"
Public myFile               As String
Public rng                  As Range
Public iter_wplat           As Integer
Public iter_wyplat          As Integer
Public iter_wersji          As Integer
Public iter_petli           As Integer
Public data                 As String
Public rap_nr               As Integer
Public fs                   As Object
Public wb1                  As Workbook
Public wb2                  As Workbook
Public open_file            As Boolean
Public f


Sub RK_macro_2021()

Dim fNameAndPath    As Variant
fNameAndPath = Application.GetOpenFilename(FileFilter:="Excel Files (*.XLS*), *.XLS*", Title:="Wybierz plik raport dzienny")
If fNameAndPath = False Then End
Set wb1 = Workbooks.Open(fNameAndPath)

myFile = Application.ActiveWorkbook.FullName & ".txt"
Set rng = wb1.Worksheets(2).Range(Cell1:="B7", Cell2:="C70")
Open myFile For Output As #1

iter_wplat = 1
iter_wyplat = 1
data = Right(wb1.Worksheets(2).Cells(3, 2), 10)
rap_nr = 0


Call Naglowek

'0 nie robi nic, 1 wymusza zapisanie jako wplate, 2 wymusza zapisanie jako wyplate
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'|                                                                                 |
'|                          IMPORTOWANE POZYCJE CZ. 1                              |
'|                                                                                 |
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call Dok_pien("Sprzeda¿ (brutto) przed rabatami i zwrotami", "WEW", 0)              'Sprzeda¿ (brutto) przed rabatami i zwrotami
Call Dok_pien("Zwroty (-)", "WWY", 0)                                               'Zwroty (-)
Call Dok_pien("Suma wp³at (+)", "TR-", 0)                                           'Suma wp³at (+)

'jesli suma wyplat jest rozna od 0 to pobierz dane z drugiego pliku
If StrComp(rng.Cells(find_row("Suma wyp³at (-)"), 2).Value, "PLN 0,00") Then
    Call OpenFile1                                                                  'Suma wyp³at (-)
End If

Call Dok_pien("Routex International", "", 0)                                      'Routex International
Call Dok_pien("UTA", "", 0)                                                       'UTA
Call Dok_pien("DKV", "", 0)                                                       'DKV
Call Dok_pien("Platnosc Punktami Payback", "", 0)                                 'Platnosc Punktami Payback
Call Dok_pien("Drive Off", "", 0)                                                'Drive Off
Call Dok_pien("BP Gift Card", "", 0)                                              'BP Gift Card
Call Dok_pien("Local Account", "", 2)                                            'Local Account
Call Dok_pien("Elavon", "", 0)                                                   'Elavon
Call Dok_pien("Dummy Tender", "", 2)                                              'Dummy Tender

'jesli korekty s¹ ujemne to zapisz z schematem 246W, jesli dodatnie to 246P
If check_neg(find_row("Korekty dostêpnych funduszy (-)")) Then                      'Korekty dostêpnych funduszy (-)
    Call Dok_pien("Korekty dostêpnych funduszy (-)", "", 1)
Else
    Call Dok_pien("Korekty dostêpnych funduszy (-)", "", 1)
End If

Call Dok_pien("Depozyty (-)", "", 0)                                             'Depozyty (-)
Call Dok_pien("Suma Superat/(Mank) dla zmian", "", 0)                           'Suma Superat/(Mank) dla zmian
Call Dok_pien("Suma Superat/(Mank) dla sejfu", "", 0)                           'Suma Superat/(Mank) dla sejfu
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'Koniec cz. 1

Call Naglowek2
iter_wersji = 1
iter_wplat = 1
iter_wyplat = 1


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'|                                                                                 |
'|                          IMPORTOWANE POZYCJE CZ. 2                              |
'|                                                                                 |
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Call Zapis("Sprzeda¿ (brutto) przed rabatami i zwrotami", "", 0)               'Sprzeda¿ (brutto) przed rabatami i zwrotami
Call Zapis("Zwroty (-)", "", 0)                                                'Zwroty (-)
Call Zapis("Suma wp³at (+)", "", 0)                                              'Suma wp³at (+)

'jesli suma wyplat jest rozna od 0 to pobierz dane z drugiego pliku
If StrComp(rng.Cells(find_row("Suma wyp³at (-)"), 2).Value, "PLN 0,00") Then
    Call OpenFile2                                                                  'Suma wyp³at (-)
End If

Call Zapis("Routex International", "", 0)                                'Routex International
Call Zapis("UTA", "", 0)                                                 'UTA
Call Zapis("DKV", "", 0)                                                 'DKV
Call Zapis("Platnosc Punktami Payback", "", 0)                           'Platnosc Punktami Payback
Call Zapis("Drive Off", "", 0)                                                   'Drive Off
Call Zapis("BP Gift Card", "", 0)                                        'BP Gift Card
Call Zapis("Local Account", "", 2)                                          'Local Account
Call Zapis("Elavon", "", 0)                                                 'Elavon
Call Zapis("Dummy Tender", "", 2)                                           'Dummy Tender
Call Zapis("Korekty dostêpnych funduszy (-)", "", 1)                             'Korekty dostêpnych funduszy (-)
Call Zapis("Depozyty (-)", "", 0)                                                'Depozyty (-)
Call Zapis("Suma Superat/(Mank) dla zmian", "", 0)                               'Suma Superat/(Mank) dla zmian
Call Zapis("Suma Superat/(Mank) dla sejfu", "", 0)                               'Suma Superat/(Mank) dla sejfu
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'Koniec cz. 2

Print #1, "}"
Close #1
wb1.Close savechanges:=False
End Sub

