Attribute VB_Name = "mdlTarjousUtils"
Public Sub LuoTarjousLomakkeelta(frm As Object)
    Dim wsTarjous As Worksheet
    Const TARJOUS_SHEET_NAME As String = "Tarjous"

    Dim parser As clsTextParser
    Set parser = New clsTextParser

    On Error Resume Next
    Set wsTarjous = ThisWorkbook.Worksheets(TARJOUS_SHEET_NAME)
    On Error GoTo 0 ' Palauta normaali virheenk�sittely heti tarkistuksen j�lkeen

    If wsTarjous Is Nothing Then
        MsgBox "V�lilehte� '" & TARJOUS_SHEET_NAME & "' ei l�ytynyt. Tarjouksen luonti peruttu.", vbCritical, "Virhe"
        Set parser = Nothing
        Exit Sub
    End If

    ' --- Muuttujien esittely (pysyy samana) ---
    Dim asiakas As String, lastausOsoiteTaysi As String, purkuOsoiteTaysi As String, m3mInput As String
    Dim lastausOsoiteLahi As String, lastausOsoiteKaupunki As String
    Dim purkuOsoiteLahi As String, purkuOsoiteKaupunki As String
    Dim m3tarjottu As String, m3varattu As String
    Dim lastausMaa As String, purkuMaa As String, puhelin As String, sahkoposti As String
    Dim lastausPaiva As String, purkuPaiva As String
    Dim tarjousTehty As String
    Dim valimatka As String

    ' --- Lue tiedot V�LITETYST� lomakkeesta (frm) ---
    ' On Error Resume Next ' POISTETTU TOISTAISEKSI virheiden paremmaksi paikantamiseksi
                        ' Jos t�m� aiheuttaa ongelmia, voidaan palauttaa tai lis�t� tarkempi k�sittely.
    On Error GoTo FormReadError ' Lis�t��n virheenk�sittelij� lomakkeen luvulle

    asiakas = frm.txtAsiakas.value
    lastausOsoiteTaysi = frm.txtLastausosoite.value
    purkuOsoiteTaysi = frm.txtPurkuosoite.value
    m3mInput = frm.txtM3m.value
    
    lastausMaa = UCase(frm.txtLastausmaa.value)
    If lastausMaa = "" Then lastausMaa = "L�ht�maa avoinna"
    
    purkuMaa = UCase(frm.txtPurkumaa.value)
    If purkuMaa = "" Then purkuMaa = "Kohdemaa avoinna"
    
    lastausPaiva = frm.txtLastauspaiva.value
    If lastausPaiva = "" Then lastausPaiva = "Lastausp�iv� avoinna"
    
    purkuPaiva = frm.txtPurkupaiva.value
    If purkuPaiva = "" Then purkuPaiva = "Purkup�iv� avoinna"
    
    puhelin = frm.txtPuhelin.value
    If puhelin = "" Then puhelin = "Puhelinnumero ei tiedossa"
    
    sahkoposti = frm.txtSahkoposti.value
    If sahkoposti = "" Then sahkoposti = "S�hk�postiosoite ei tiedossa"
    
    valimatka = frm.txtValimatka.value
    If valimatka = "" Then valimatka = "V�limatka avoinna" Else valimatka = valimatka & " km"
    
    tarjousTehty = frm.txtTarjousTehty.value ' Oletetaan .value

    On Error GoTo 0 ' Palauta normaali virheenk�sittely, jos kaikki meni hyvin

    ' --- K�yt� clsTextParseria tietojen j�sent�miseen ---

    ' 1. J�senn� lastausosoite (erotin: pilkku)
    If lastausOsoiteTaysi = "" Then
        lastausOsoiteLahi = "Lastausosoite avoinnna, "
        lastausOsoiteKaupunki = ""
    Else
        parser.InputText = lastausOsoiteTaysi
        parser.Delimiter = ","
        parser.Parse ' Suorita j�sent�minen
        lastausOsoiteLahi = parser.Part1 & ", "       ' Osa ennen pilkkua (tai koko, jos ei pilkkua)
        If parser.Part2 = "" Then
            lastausOsoiteKaupunki = ""              ' Jos tyhj�, niin ei pilkkua
        Else
            lastausOsoiteKaupunki = parser.Part2 & ", "   ' Jos on tieto, niin laitetaan pilkku
        End If
    End If
    
    ' 2. J�senn� purkuosoite (erotin: pilkku)
    If purkuOsoiteTaysi = "" Then
        purkuOsoiteLahi = "Purkuosoite avoinna, "
        purkuOsoiteKaupunki = ""
    Else
        parser.InputText = purkuOsoiteTaysi
        parser.Delimiter = "," ' Varmuuden vuoksi asetetaan uudelleen, vaikka se on sama
        parser.Parse
        purkuOsoiteLahi = parser.Part1 & ", "
        If parser.Part2 = "" Then
            purkuOsoiteKaupunki = ""
        Else
        purkuOsoiteKaupunki = parser.Part2 & ", "
        End If
    End If

    ' 3. J�senn� M3 (erotin: viiva "-"), k�yt� erityiss��nt��
    parser.InputText = m3mInput
    parser.Delimiter = "-"
    parser.Parse

    If parser.Part1 = "" Then
        m3tarjottu = "Kuutiot avoinna"
        m3varattu = "Kuutiot avoinna"
    Else
        If parser.IsSplit Then ' Tarkista, l�ytyik� erotin ("-")
            ' Viiva l�ytyi, ota osat normaalisti
            m3tarjottu = parser.Part1
            m3varattu = parser.Part2
        Else
            ' Viivaa EI l�ytynyt, laita sama arvo molempiin
            m3tarjottu = parser.Part1 ' Part1 sis�lt�� koko sy�tteen, jos ei jaettu
            m3varattu = parser.Part1 ' Sama arvo my�s t�h�n
        End If
    End If

    ' --- Kirjoita tiedot Tarjous-v�lilehdelle ---
    With wsTarjous
        .Range("D5").value = IsotAlkukirjaimet(asiakas)
        .Range("D6").value = puhelin
        .Range("D7").value = sahkoposti
        
        .Range("D13").value = lastausPaiva
        .Range("D14").value = purkuPaiva
        
        .Range("D16").value = IsotAlkukirjaimet(lastausOsoiteLahi & lastausOsoiteKaupunki & lastausMaa)
        .Range("D17").value = IsotAlkukirjaimet(purkuOsoiteLahi & purkuOsoiteKaupunki & purkuMaa)

        .Range("D18").value = m3tarjottu
        .Range("D19").value = m3varattu
        .Range("D20").value = valimatka
        
        .Range("G2").value = tarjousTehty

    End With

CleanExit_LuoTarjous:
    Set wsTarjous = Nothing
    Set parser = Nothing
    Exit Sub

FormReadError:
    MsgBox "Virhe luettaessa tietoja tarjouslomakkeelta." & vbCrLf & _
           "Tarkista, ett� kaikki tarvittavat kent�t ovat olemassa ja oikein nimetty." & vbCrLf & _
           "Virhe: " & Err.Description, vbCritical, "Lomakkeen Lukuvirhe"
    Resume CleanExit_LuoTarjous ' Poistu siististi virheen j�lkeen

End Sub
