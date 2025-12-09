# INPUTBOX

deger = InputBox("Tutar") 'Klasik Inputbox 

deger = Trim(InputBox("Tutar")) 'Baştaki ve Sondaki Boşlukları Temizler

deger = Val(Trim(InputBox("Tutar"))) 'Stringi sayıya çevirir. 

# DÖNGÜ İÇİ INPUTBOX
For i = 2 To Lastrow
    deger1 = Trim(ActiveWorkbook.Sheets("Sayfa1").Cells(i, 1).Value)
    
    If i = 2 Then
        deger2 = Trim(InputBox("deger2 için bir değer girin:", "deger2"))
        If deger2 = "" Then
            If MsgBox("deger2 boş bırakıldı. İşleme devam edilsin mi?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        End If
    End If

    deger3 = ActiveWorkbook.Sheets("Sayfa1").Cells(i, 3).Value

    ' SAP işlemleri...
Next i

