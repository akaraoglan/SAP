# INPUTBOX

deger = InputBox("Tutar") 'Klasik Inputbox 

deger = Trim(InputBox("Tutar")) 'Baştaki ve Sondaki Boşlukları Temizler

deger = Val(Trim(InputBox("Tutar"))) 'Stringi sayıya çevirir. 

# DÖNGÜ İÇİ INPUTBOX

    If i = 2 Then
        deger2 = Trim(InputBox("deger2 için bir değer girin:", "deger2"))
        If deger2 = "" Then
            If MsgBox("deger2 boş bırakıldı. İşleme devam edilsin mi?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        End If
    End If

' Döngü içinde If i = 2 Then kullanıldığında, InputBox yalnızca ilk döngüde çalışır, deger2 bir kez alınır ve döngü sonuna kadar aynı değerle işlemler devam eder.


