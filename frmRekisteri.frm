VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRekisteri 
   Caption         =   "Rekisteri"
   ClientHeight    =   8844.001
   ClientLeft      =   288
   ClientTop       =   1092
   ClientWidth     =   14652
   OleObjectBlob   =   "frmRekisteri.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRekisteri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' --- frmRekisterit-lomakkeen Initialize-rutiini ---
Private Sub UserForm_Initialize()
    On Error GoTo Initialize_Error

    ' --- 1. Tyhjennä lomakkeen kontrollit oletusarvoihin ---
    ' Varmista, että mdlClearForm-moduuli on projektissasi.
    mdlClearForm.ClearForm Me

    ' --- 2. Lataa olemassa olevat rekisteritiedot ListBoxeihin ---
    ' Olettaa, että mdlRegisterUtils.LoadRegisterDataToListBox on luotu
    ' ja välilehdet on nimetty oikein.
    LoadRegisterDataToListBox Me.lstPalvelut, "Palvelut"
    LoadRegisterDataToListBox Me.lstAutot, "Autot"
    LoadRegisterDataToListBox Me.lstKontit, "Kontit"
    LoadRegisterDataToListBox Me.lstKuljettajat, "Kuljettajat"
    LoadRegisterDataToListBox Me.lstApulaiset, "Apulaiset"

    ' --- 3. (Valinnainen) Muut alustukset ---
    ' Esimerkiksi aseta kohdistus ensimmäiseen syötekenttään
    ' Me.txtPalvelu.SetFocus

    ' Tai poista Poista/Muokkaa-painikkeet käytöstä aluksi
    ' (Tämä riippuu tarkasta käyttöliittymälogiikasta, usein _Click-event hoitaa tämän)


CleanExit_Initialize:
    Exit Sub

Initialize_Error:
    MsgBox "Virhe alustettaessa rekisteritietojen hallintalomaketta:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Lomakkeen Alustusvirhe"
    ' Voit harkita lomakkeen sulkemista tässä, jos alustus epäonnistuu kriittisesti
    ' On Error Resume Next
    ' Unload Me
    Resume CleanExit_Initialize
End Sub

' --- PALVELUT-OSION PAINIKKEET ---

Private Sub cmdTallennaPalvelu_Click()
    Dim serviceName As String
    Dim existingRow As Long
    Dim newID As Long
    Dim data(1 To 2) As Variant ' Taulukko datalle (1=ID, 2=Nimi)
    Const SHEET_NAME As String = "Palvelut" ' Välilehden nimi

    On Error GoTo SaveServiceError
'MsgBox "HÄ?"
    ' 1. Lue ja validoi syöte
    serviceName = Trim$(Me.txtPalvelu.Text)
    If serviceName = "" Then
        MsgBox "Palvelun nimi ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtPalvelu.SetFocus
        Exit Sub
    End If

    ' 2. Tarkista duplikaatit (Case-insensitive)
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, serviceName, NAME_COL)
    If existingRow <> 0 Then
        MsgBox "Palvelu nimellä '" & serviceName & "' on jo olemassa rivillä " & existingRow & ".", vbExclamation, "Duplikaatti"
        Me.txtPalvelu.SetFocus
        Exit Sub
    End If

    ' 3. Hae seuraava ID
    newID = mdlRegisterUtils.GetNextRegisterID(SHEET_NAME)
    If newID = 0 Then ' GetNextRegisterID palauttaa 0 virhetilanteessa
        MsgBox "Uutta ID:tä ei voitu hakea. Lisäys epäonnistui.", vbCritical, "ID Virhe"
        Exit Sub
    End If

    ' 4. Kokoa data lisäystä varten
    data(1) = newID
    data(2) = serviceName

    ' 5. Lisää tieto välilehdelle mdlRegisterUtils-funktion avulla
    If mdlRegisterUtils.AddRegisterItem(SHEET_NAME, data) Then
        ' 6. Päivitä UI onnistuneen lisäyksen jälkeen
        LoadRegisterDataToListBox Me.lstPalvelut, SHEET_NAME ' Päivitä listbox
        Me.txtPalvelu.Text = "" ' Tyhjennä syötekenttä
        Me.lstPalvelut.listIndex = -1 ' Poista valinta listalta
        ' Me.cmdPoistaPalvelu.Enabled = False ' Poista Poista-nappi käytöstä (koska valinta poistui)
        MsgBox "Uusi palvelu '" & serviceName & "' (ID: " & newID & ") lisätty onnistuneesti.", vbInformation, "Lisäys Onnistui"
    Else
        ' Virheviesti tuli jo AddRegisterItem-funktiosta
    End If

CleanExit_SaveService:
    Exit Sub

SaveServiceError:
    MsgBox "Odottamaton virhe tallennettaessa palvelua:" & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    Resume CleanExit_SaveService
End Sub


Private Sub cmdPoistaPalvelu_Click()
    Dim itemID As Long
    Dim serviceName As String
    Dim listIndex As Long
    Dim response As VbMsgBoxResult
    Const SHEET_NAME As String = "Palvelut"

    On Error GoTo DeleteServiceError

    ' 1. Varmista, että jokin on valittuna listassa
    listIndex = Me.lstPalvelut.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse poistettava palvelu listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    ' 2. Hae valitun kohteen ID ja nimi (ID piilotetusta sarakkeesta 0)
    On Error Resume Next ' Virheenkäsittely, jos Column(0) ei ole numero
    itemID = CLng(Me.lstPalvelut.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Poisto epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo DeleteServiceError ' Palauta normaali käsittely
        Exit Sub
    End If
    On Error GoTo DeleteServiceError ' Palauta normaali käsittely
    serviceName = Me.lstPalvelut.Column(1, listIndex) ' Nimi sarakkeesta 1

    ' 3. Varmista poisto käyttäjältä
    response = MsgBox("Haluatko varmasti poistaa palvelun:" & vbCrLf & vbCrLf & _
                      "ID: " & itemID & vbCrLf & _
                      "Nimi: " & serviceName & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    If response = vbNo Then Exit Sub

    ' 4. Poista tieto välilehdeltä mdlRegisterUtils-funktion avulla
    If mdlRegisterUtils.DeleteRegisterItem(SHEET_NAME, itemID) Then
        ' 5. Päivitä UI onnistuneen poiston jälkeen
        LoadRegisterDataToListBox Me.lstPalvelut, SHEET_NAME ' Päivitä listbox
        Me.txtPalvelu.Text = "" ' Tyhjennä syötekenttä
        Me.lstPalvelut.listIndex = -1 ' Poista valinta
        Me.cmdPoistaPalvelu.Enabled = False ' Poista Poista-nappi käytöstä
        MsgBox "Palvelu '" & serviceName & "' (ID: " & itemID & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"
    Else
        ' Virheviesti tuli jo DeleteRegisterItem-funktiosta
        ' Varmistetaan, ettei poistonappi jää turhaan päälle, jos poisto epäonnistui
         Me.cmdPoistaPalvelu.Enabled = (Me.lstPalvelut.listIndex > -1)
    End If


CleanExit_DeleteService:
    Exit Sub

DeleteServiceError:
     MsgBox "Odottamaton virhe poistettaessa palvelua:" & vbCrLf & Err.Description, vbCritical, "Poistovirhe"
     Resume CleanExit_DeleteService
End Sub

Private Sub lstPalvelut_Click()
    Dim listIndex As Long
    listIndex = Me.lstPalvelut.listIndex

    If listIndex > -1 Then
        ' Näytä valitun palvelun nimi tekstikentässä
        Me.txtPalvelu.Text = Me.lstPalvelut.Column(1, listIndex)
        ' Aktivoi Poista-painike
        Me.cmdPoistaPalvelu.Enabled = True
    Else
        ' Jos valinta poistuu, tyhjennä kenttä ja deaktivoi nappi
        Me.txtPalvelu.Text = ""
        Me.cmdPoistaPalvelu.Enabled = False
    End If
End Sub

Private Sub cmdTallennaAuto_Click()
    Dim regNum As String
    Dim existingRow As Long
    Dim newID As Long
    Dim data(1 To 2) As Variant ' Taulukko datalle (1=ID, 2=Nimi)
    Const SHEET_NAME As String = "Autot" ' Välilehden nimi

    On Error GoTo SaveServiceError

    ' 1. Lue ja validoi syöte
    regNum = Trim$(Me.txtAuto.Text)
    If regNum = "" Then
        MsgBox "Auton rekisterinumero ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtAuto.SetFocus
        Exit Sub
    End If

    ' 2. Tarkista duplikaatit (Case-insensitive)
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, regNum, NAME_COL)
    If existingRow <> 0 Then
        MsgBox "Auto rekisterinumerolla '" & regNum & "' on jo olemassa rivillä " & existingRow & ".", vbExclamation, "Duplikaatti"
        Me.txtPalvelu.SetFocus
        Exit Sub
    End If

    ' 3. Hae seuraava ID
    newID = mdlRegisterUtils.GetNextRegisterID(SHEET_NAME)
    If newID = 0 Then ' GetNextRegisterID palauttaa 0 virhetilanteessa
        MsgBox "Uutta ID:tä ei voitu hakea. Lisäys epäonnistui.", vbCritical, "ID Virhe"
        Exit Sub
    End If

    ' 4. Kokoa data lisäystä varten
    data(1) = newID
    data(2) = regNum

    ' 5. Lisää tieto välilehdelle mdlRegisterUtils-funktion avulla
    If mdlRegisterUtils.AddRegisterItem(SHEET_NAME, data) Then
        ' 6. Päivitä UI onnistuneen lisäyksen jälkeen
        LoadRegisterDataToListBox Me.lstAutot, SHEET_NAME ' Päivitä listbox
        Me.txtPalvelu.Text = "" ' Tyhjennä syötekenttä
        Me.lstPalvelut.listIndex = -1 ' Poista valinta listalta
        ' Me.cmdPoistaPalvelu.Enabled = False ' Poista Poista-nappi käytöstä (koska valinta poistui)
        MsgBox "Uusi auto rekisterinumerolla '" & regNum & "' (ID: " & newID & ") lisätty onnistuneesti.", vbInformation, "Lisäys Onnistui"
    Else
        ' Virheviesti tuli jo AddRegisterItem-funktiosta
    End If

CleanExit_SaveService:
    Exit Sub

SaveServiceError:
    MsgBox "Odottamaton virhe tallennettaessa palvelua:" & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    Resume CleanExit_SaveService
End Sub


Private Sub cmdPoistaAuto_Click()
    Dim itemID As Long
    Dim regNum As String
    Dim listIndex As Long
    Dim response As VbMsgBoxResult
    Const SHEET_NAME As String = "Autot"

    On Error GoTo DeleteServiceError

    ' 1. Varmista, että jokin on valittuna listassa
    listIndex = Me.lstAutot.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse poistettava auto listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    ' 2. Hae valitun kohteen ID ja nimi (ID piilotetusta sarakkeesta 0)
    On Error Resume Next ' Virheenkäsittely, jos Column(0) ei ole numero
    itemID = CLng(Me.lstAutot.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Poisto epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo DeleteServiceError ' Palauta normaali käsittely
        Exit Sub
    End If
    On Error GoTo DeleteServiceError ' Palauta normaali käsittely
    regNum = Me.lstAutot.Column(1, listIndex) ' Nimi sarakkeesta 1

    ' 3. Varmista poisto käyttäjältä
    response = MsgBox("Haluatko varmasti poistaa auton:" & vbCrLf & vbCrLf & _
                      "ID: " & itemID & vbCrLf & _
                      "Rekisterinumero: " & regNum & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    If response = vbNo Then Exit Sub

    ' 4. Poista tieto välilehdeltä mdlRegisterUtils-funktion avulla
    If mdlRegisterUtils.DeleteRegisterItem(SHEET_NAME, itemID) Then
        ' 5. Päivitä UI onnistuneen poiston jälkeen
        LoadRegisterDataToListBox Me.lstAutot, SHEET_NAME ' Päivitä listbox
        Me.txtAuto.Text = "" ' Tyhjennä syötekenttä
        Me.lstAutot.listIndex = -1 ' Poista valinta
        Me.cmdPoistaAuto.Enabled = False ' Poista Poista-nappi käytöstä
        MsgBox "Auto rekisterinumerolla '" & regNum & "' (ID: " & itemID & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"
    Else
        ' Virheviesti tuli jo DeleteRegisterItem-funktiosta
        ' Varmistetaan, ettei poistonappi jää turhaan päälle, jos poisto epäonnistui
         Me.cmdPoistaAuto.Enabled = (Me.lstAutot.listIndex > -1)
    End If


CleanExit_DeleteService:
    Exit Sub

DeleteServiceError:
     MsgBox "Odottamaton virhe poistettaessa autoa:" & vbCrLf & Err.Description, vbCritical, "Poistovirhe"
     Resume CleanExit_DeleteService
End Sub

Private Sub lstAutot_Click()
    Dim listIndex As Long
    listIndex = Me.lstAutot.listIndex

    If listIndex > -1 Then
        ' Näytä valitun palvelun nimi tekstikentässä
        Me.txtAuto.Text = Me.lstAutot.Column(1, listIndex)
        ' Aktivoi Poista-painike
        Me.cmdPoistaAuto.Enabled = True
    Else
        ' Jos valinta poistuu, tyhjennä kenttä ja deaktivoi nappi
        Me.txtAuto.Text = ""
        Me.cmdPoistaAuto.Enabled = False
    End If
End Sub

Private Sub cmdTallennaKontti_Click()
    Dim regNum As String
    Dim existingRow As Long
    Dim newID As Long
    Dim data(1 To 2) As Variant ' Taulukko datalle (1=ID, 2=Nimi)
    Const SHEET_NAME As String = "Kontit" ' Välilehden nimi

    On Error GoTo SaveServiceError

    ' 1. Lue ja validoi syöte
    regNum = Trim$(Me.txtKontti.Text)
    If regNum = "" Then
        MsgBox "Kontin rekisterinumero ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtKontti.SetFocus
        Exit Sub
    End If

    ' 2. Tarkista duplikaatit (Case-insensitive)
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, regNum, NAME_COL)
    If existingRow <> 0 Then
        MsgBox "Kontti rekisterinumerolla '" & regNum & "' on jo olemassa rivillä " & existingRow & ".", vbExclamation, "Duplikaatti"
        Me.txtKontti.SetFocus
        Exit Sub
    End If

    ' 3. Hae seuraava ID
    newID = mdlRegisterUtils.GetNextRegisterID(SHEET_NAME)
    If newID = 0 Then ' GetNextRegisterID palauttaa 0 virhetilanteessa
        MsgBox "Uutta ID:tä ei voitu hakea. Lisäys epäonnistui.", vbCritical, "ID Virhe"
        Exit Sub
    End If

    ' 4. Kokoa data lisäystä varten
    data(1) = newID
    data(2) = regNum

    ' 5. Lisää tieto välilehdelle mdlRegisterUtils-funktion avulla
    If mdlRegisterUtils.AddRegisterItem(SHEET_NAME, data) Then
        ' 6. Päivitä UI onnistuneen lisäyksen jälkeen
        LoadRegisterDataToListBox Me.lstKontit, SHEET_NAME ' Päivitä listbox
        Me.txtKontti.Text = "" ' Tyhjennä syötekenttä
        Me.lstKontit.listIndex = -1 ' Poista valinta listalta
        ' Me.cmdPoistaPalvelu.Enabled = False ' Poista Poista-nappi käytöstä (koska valinta poistui)
        MsgBox "Uusi kontti rekisterinumerolla '" & regNum & "' (ID: " & newID & ") lisätty onnistuneesti.", vbInformation, "Lisäys Onnistui"
    Else
        ' Virheviesti tuli jo AddRegisterItem-funktiosta
    End If

CleanExit_SaveService:
    Exit Sub

SaveServiceError:
    MsgBox "Odottamaton virhe tallennettaessa konttia:" & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    Resume CleanExit_SaveService
End Sub


Private Sub cmdPoistaKontti_Click()
    Dim itemID As Long
    Dim regNum As String
    Dim listIndex As Long
    Dim response As VbMsgBoxResult
    Const SHEET_NAME As String = "Kontit"

    On Error GoTo DeleteServiceError

    ' 1. Varmista, että jokin on valittuna listassa
    listIndex = Me.lstKontit.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse poistettava kontti listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    ' 2. Hae valitun kohteen ID ja nimi (ID piilotetusta sarakkeesta 0)
    On Error Resume Next ' Virheenkäsittely, jos Column(0) ei ole numero
    itemID = CLng(Me.lstKontit.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Poisto epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo DeleteServiceError ' Palauta normaali käsittely
        Exit Sub
    End If
    On Error GoTo DeleteServiceError ' Palauta normaali käsittely
    regNum = Me.lstKontit.Column(1, listIndex) ' Nimi sarakkeesta 1

    ' 3. Varmista poisto käyttäjältä
    response = MsgBox("Haluatko varmasti poistaa kontin:" & vbCrLf & vbCrLf & _
                      "ID: " & itemID & vbCrLf & _
                      "Rekisterinumero: " & regNum & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    If response = vbNo Then Exit Sub

    ' 4. Poista tieto välilehdeltä mdlRegisterUtils-funktion avulla
    If mdlRegisterUtils.DeleteRegisterItem(SHEET_NAME, itemID) Then
        ' 5. Päivitä UI onnistuneen poiston jälkeen
        LoadRegisterDataToListBox Me.lstKontit, SHEET_NAME ' Päivitä listbox
        Me.txtKontti.Text = "" ' Tyhjennä syötekenttä
        Me.lstKontit.listIndex = -1 ' Poista valinta
        Me.cmdPoistaKontti.Enabled = False ' Poista Poista-nappi käytöstä
        MsgBox "Kontti rekisterinumerolla '" & regNum & "' (ID: " & itemID & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"
    Else
        ' Virheviesti tuli jo DeleteRegisterItem-funktiosta
        ' Varmistetaan, ettei poistonappi jää turhaan päälle, jos poisto epäonnistui
         Me.cmdPoistaKontti.Enabled = (Me.lstKontit.listIndex > -1)
    End If


CleanExit_DeleteService:
    Exit Sub

DeleteServiceError:
     MsgBox "Odottamaton virhe poistettaessa konttia:" & vbCrLf & Err.Description, vbCritical, "Poistovirhe"
     Resume CleanExit_DeleteService
End Sub

Private Sub lstKontit_Click()
    Dim listIndex As Long
    listIndex = Me.lstKontit.listIndex

    If listIndex > -1 Then
        ' Näytä valitun palvelun nimi tekstikentässä
        Me.txtKontti.Text = Me.lstKontit.Column(1, listIndex)
        ' Aktivoi Poista-painike
        Me.cmdPoistaKontti.Enabled = True
    Else
        ' Jos valinta poistuu, tyhjennä kenttä ja deaktivoi nappi
        Me.txtKontti.Text = ""
        Me.cmdPoistaKontti.Enabled = False
    End If
End Sub

' --- KULJETTAJAT-OSION KÄSITTELIJÄT ---

Private Sub lstKuljettajat_Click()
    Dim listIndex As Long
    listIndex = Me.lstKuljettajat.listIndex

    On Error GoTo ListClickError ' Virheenkäsittely

    If listIndex > -1 Then
        ' --- Täytä kentät valitun rivin tiedoilla ---
        ' Sarakkeet: 0=ID, 1=Nimi, 2=Puhelin, 3=Sposti, 4=Osoite
        Me.txtKuljettaja.Text = Me.lstKuljettajat.Column(1, listIndex)
        Me.txtKuljettajaPuhelin.Text = Me.lstKuljettajat.Column(2, listIndex)
        Me.txtKuljettajaSposti.Text = Me.lstKuljettajat.Column(3, listIndex)
        Me.txtKuljettajaOsoite.Text = Me.lstKuljettajat.Column(4, listIndex)

        ' --- Aktivoi Muokkaa ja Poista -painikkeet ---
        Me.cmdMuokkaaKuljettaja.Enabled = True
        Me.cmdPoistaKuljettaja.Enabled = True
        ' Tallenna (Lisää) -painike voidaan poistaa käytöstä, kun muokataan
        ' Me.cmdTallennaKuljettaja.Enabled = False
    Else
        ' --- Tyhjennä kentät ja deaktivoi napit, jos valinta poistuu ---
        Me.txtKuljettaja.Text = ""
        Me.txtKuljettajaPuhelin.Text = ""
        Me.txtKuljettajaSposti.Text = ""
        Me.txtKuljettajaOsoite.Text = ""
        Me.cmdMuokkaaKuljettaja.Enabled = False
        Me.cmdPoistaKuljettaja.Enabled = False
       '  Aktivoi Tallenna (Lisää) -painike
        ' Me.cmdTallennaKuljettaja.Enabled = True
    End If

CleanExit_ListClick:
    Exit Sub

ListClickError:
     MsgBox "Virhe käsiteltäessä Kuljettaja-listan valintaa:" & vbCrLf & Err.Description, vbCritical, "Lista Virhe"
     ' Yritä nollata tila
     Me.txtKuljettaja.Text = ""
     Me.txtKuljettajaPuhelin.Text = ""
     Me.txtKuljettajaSposti.Text = ""
     Me.txtKuljettajaOsoite.Text = ""
     Me.cmdMuokkaaKuljettaja.Enabled = False
     Me.cmdPoistaKuljettaja.Enabled = False
     Resume CleanExit_ListClick
End Sub


Private Sub cmdTallennaKuljettaja_Click() ' LISÄÄ UUSI KULJETTAJA
    Dim driverName As String, phone As String, email As String, address As String
    Dim existingRow As Long
    Dim newID As Long
    Dim data(1 To 5) As Variant ' ID, Nimi, Puh, Sposti, Osoite
    Const SHEET_NAME As String = "Kuljettajat"

    On Error GoTo SaveDriverError

    ' 1. Lue ja validoi syötteet
    driverName = Trim$(Me.txtKuljettaja.Text)
    phone = Trim$(Me.txtKuljettajaPuhelin.Text)
    email = Trim$(Me.txtKuljettajaSposti.Text)
    address = Trim$(Me.txtKuljettajaOsoite.Text)

    If driverName = "" Then
        MsgBox "Kuljettajan nimi ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtKuljettaja.SetFocus
        Exit Sub
    End If

    ' 2. Tarkista duplikaattinimi
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, driverName, NAME_COL)
    If existingRow <> 0 Then
        MsgBox "Kuljettaja nimellä '" & driverName & "' on jo olemassa rivillä " & existingRow & ".", vbExclamation, "Duplikaatti"
        Me.txtKuljettaja.SetFocus
        Exit Sub
    End If

    ' 3. Hae seuraava ID
    newID = mdlRegisterUtils.GetNextRegisterID(SHEET_NAME)
    If newID = 0 Then
        MsgBox "Uutta ID:tä ei voitu hakea. Lisäys epäonnistui.", vbCritical, "ID Virhe"
        Exit Sub
    End If

    ' 4. Kokoa data lisäystä varten
    data(1) = newID
    data(2) = driverName
    data(3) = phone
    data(4) = email
    data(5) = address

    ' 5. Lisää tieto välilehdelle
    If mdlRegisterUtils.AddRegisterItem(SHEET_NAME, data) Then
        ' 6. Päivitä UI
        LoadRegisterDataToListBox Me.lstKuljettajat, SHEET_NAME ' Päivitä lista
        ' Tyhjennä kentät
        Me.txtKuljettaja.Text = ""
        Me.txtKuljettajaPuhelin.Text = ""
        Me.txtKuljettajaSposti.Text = ""
        Me.txtKuljettajaOsoite.Text = ""
        Me.lstKuljettajat.listIndex = -1 ' Poista valinta
        Me.cmdMuokkaaKuljettaja.Enabled = False ' Varmista napin tila
        Me.cmdPoistaKuljettaja.Enabled = False
        MsgBox "Uusi kuljettaja '" & driverName & "' (ID: " & newID & ") lisätty onnistuneesti.", vbInformation, "Lisäys Onnistui"
    End If

CleanExit_SaveDriver:
    Exit Sub

SaveDriverError:
    MsgBox "Odottamaton virhe tallennettaessa kuljettajaa:" & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    Resume CleanExit_SaveDriver
End Sub


Private Sub cmdMuokkaaKuljettaja_Click() ' TALLENNA MUUTOKSET OLEMASSA OLEVAAN
    Dim driverName As String, phone As String, email As String, address As String
    Dim itemID As Long
    Dim listIndex As Long
    Dim existingRow As Long
    Dim existingID As Long
    Dim data(1 To 5) As Variant ' ID, Nimi, Puh, Sposti, Osoite
    Const SHEET_NAME As String = "Kuljettajat"
    Dim ws As Worksheet ' Tarvitaan duplikaatin ID:n tarkistukseen

    On Error GoTo EditDriverError

    ' 1. Varmista, että jokin on valittuna ja hae ID
    listIndex = Me.lstKuljettajat.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse muokattava kuljettaja listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    On Error Resume Next ' Virheenkäsittely ID:n luvussa
    itemID = CLng(Me.lstKuljettajat.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Muokkaus epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo EditDriverError
        Exit Sub
    End If
    On Error GoTo EditDriverError ' Palauta normaali

    ' 2. Lue ja validoi MUOKATUT tiedot kentistä
    driverName = Trim$(Me.txtKuljettaja.Text)
    phone = Trim$(Me.txtKuljettajaPuhelin.Text)
    email = Trim$(Me.txtKuljettajaSposti.Text)
    address = Trim$(Me.txtKuljettajaOsoite.Text)

    If driverName = "" Then
        MsgBox "Kuljettajan nimi ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtKuljettaja.SetFocus
        Exit Sub
    End If

    ' 3. Tarkista duplikaattinimi (mutta salli sama nimi ITSELLÄÄN)
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, driverName, NAME_COL)
    If existingRow <> 0 Then
        ' Nimi löytyi, tarkista onko se eri ID kuin muokattavalla
        Set ws = ThisWorkbook.Worksheets(SHEET_NAME) ' Hae worksheet-objekti
        If Not ws Is Nothing Then
            On Error Resume Next ' ID:n luku voi epäonnistua
            existingID = CLng(ws.Cells(existingRow, ID_COL).value)
            On Error GoTo EditDriverError ' Palauta normaali
            If existingID <> itemID Then
                 MsgBox "Kuljettaja nimellä '" & driverName & "' on jo olemassa (ID: " & existingID & "). Anna eri nimi.", vbExclamation, "Duplikaatti"
                 Me.txtKuljettaja.SetFocus
                 Set ws = Nothing
                 Exit Sub
            End If
        End If
        Set ws = Nothing
    End If

    ' 4. Kokoa data päivitystä varten (käytä alkuperäistä itemID:tä)
    data(1) = itemID ' Tärkeää: ID säilyy samana
    data(2) = driverName
    data(3) = phone
    data(4) = email
    data(5) = address

    ' 5. Päivitä tieto välilehdelle
    If mdlRegisterUtils.UpdateRegisterItem(SHEET_NAME, itemID, data) Then
        ' 6. Päivitä UI
        LoadRegisterDataToListBox Me.lstKuljettajat, SHEET_NAME ' Päivitä lista
        ' Tyhjennä kentät ja deaktivoi napit
        Me.txtKuljettaja.Text = ""
        Me.txtKuljettajaPuhelin.Text = ""
        Me.txtKuljettajaSposti.Text = ""
        Me.txtKuljettajaOsoite.Text = ""
        Me.lstKuljettajat.listIndex = -1
        Me.cmdMuokkaaKuljettaja.Enabled = False
        Me.cmdPoistaKuljettaja.Enabled = False
        MsgBox "Kuljettajan '" & driverName & "' (ID: " & itemID & ") tiedot päivitetty onnistuneesti.", vbInformation, "Päivitys Onnistui"
    End If

CleanExit_EditDriver:
    Set ws = Nothing
    Exit Sub

EditDriverError:
     MsgBox "Odottamaton virhe muokattaessa kuljettajaa:" & vbCrLf & Err.Description, vbCritical, "Muokkausvirhe"
     Resume CleanExit_EditDriver
End Sub


Private Sub cmdPoistaKuljettaja_Click()
    Dim itemID As Long
    Dim driverName As String
    Dim listIndex As Long
    Dim response As VbMsgBoxResult
    Const SHEET_NAME As String = "Kuljettajat"

    On Error GoTo DeleteDriverError

    ' 1. Varmista valinta ja hae ID + Nimi
    listIndex = Me.lstKuljettajat.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse poistettava kuljettaja listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    On Error Resume Next ' ID:n luku
    itemID = CLng(Me.lstKuljettajat.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Poisto epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo DeleteDriverError
        Exit Sub
    End If
    On Error GoTo DeleteDriverError ' Palauta normaali
    driverName = Me.lstKuljettajat.Column(1, listIndex)

    ' 2. Varmista poisto
    response = MsgBox("Haluatko varmasti poistaa kuljettajan:" & vbCrLf & vbCrLf & _
                      "ID: " & itemID & vbCrLf & _
                      "Nimi: " & driverName & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    If response = vbNo Then Exit Sub

    ' 3. Poista tieto välilehdeltä
    If mdlRegisterUtils.DeleteRegisterItem(SHEET_NAME, itemID) Then
        ' 4. Päivitä UI
        LoadRegisterDataToListBox Me.lstKuljettajat, SHEET_NAME ' Päivitä lista
        ' Tyhjennä kentät ja deaktivoi napit
        Me.txtKuljettaja.Text = ""
        Me.txtKuljettajaPuhelin.Text = ""
        Me.txtKuljettajaSposti.Text = ""
        Me.txtKuljettajaOsoite.Text = ""
        Me.lstKuljettajat.listIndex = -1
        Me.cmdMuokkaaKuljettaja.Enabled = False
        Me.cmdPoistaKuljettaja.Enabled = False
        MsgBox "Kuljettaja '" & driverName & "' (ID: " & itemID & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"
    End If

CleanExit_DeleteDriver:
    Exit Sub

DeleteDriverError:
     MsgBox "Odottamaton virhe poistettaessa kuljettajaa:" & vbCrLf & Err.Description, vbCritical, "Poistovirhe"
     Resume CleanExit_DeleteDriver
End Sub

' --- APULAISET-OSION KÄSITTELIJÄT ---

Private Sub lstApulaiset_Click()
    Dim listIndex As Long
    listIndex = Me.lstApulaiset.listIndex

    On Error GoTo ListClickError ' Virheenkäsittely

    If listIndex > -1 Then
        ' --- Täytä kentät valitun rivin tiedoilla ---
        ' Sarakkeet: 0=ID, 1=Nimi, 2=Puhelin, 3=Sposti, 4=Osoite
        Me.txtApulainen.Text = Me.lstApulaiset.Column(1, listIndex)
        Me.txtApulainenPuhelin.Text = Me.lstApulaiset.Column(2, listIndex)
        Me.txtApulainenSposti.Text = Me.lstApulaiset.Column(3, listIndex)
        Me.txtApulainenOsoite.Text = Me.lstApulaiset.Column(4, listIndex)

        ' --- Aktivoi Muokkaa ja Poista -painikkeet ---
        Me.cmdMuokkaaApulainen.Enabled = True
        Me.cmdPoistaApulainen.Enabled = True
        ' Tallenna (Lisää) -painike voidaan poistaa käytöstä, kun muokataan
        ' Me.cmdTallennaKuljettaja.Enabled = False
    Else
        ' --- Tyhjennä kentät ja deaktivoi napit, jos valinta poistuu ---
        Me.txtApulainen.Text = ""
        Me.txtApulainenPuhelin.Text = ""
        Me.txtApulainenSposti.Text = ""
        Me.txtApulainenOsoite.Text = ""
        Me.cmdMuokkaaApulainen.Enabled = False
        Me.cmdPoistaApulainen.Enabled = False
        ' Aktivoi Tallenna (Lisää) -painike
        ' Me.cmdTallennaKuljettaja.Enabled = True
    End If

CleanExit_ListClick:
    Exit Sub

ListClickError:
     MsgBox "Virhe käsiteltäessä Apulainen-listan valintaa:" & vbCrLf & Err.Description, vbCritical, "Lista Virhe"
     ' Yritä nollata tila
     Me.txtApulainen.Text = ""
     Me.txtApulainenPuhelin.Text = ""
     Me.txtApulainenSposti.Text = ""
     Me.txtApulainenOsoite.Text = ""
     Me.cmdMuokkaaApulainen.Enabled = False
     Me.cmdPoistaApulainen.Enabled = False
     Resume CleanExit_ListClick
End Sub


Private Sub cmdTallennaApulainen_Click() ' LISÄÄ UUSI APULAINEN
    Dim driverName As String, phone As String, email As String, address As String
    Dim existingRow As Long
    Dim newID As Long
    Dim data(1 To 5) As Variant ' ID, Nimi, Puh, Sposti, Osoite
    Const SHEET_NAME As String = "Apulaiset"

    On Error GoTo SaveDriverError

    ' 1. Lue ja validoi syötteet
    driverName = Trim$(Me.txtApulainen.Text)
    phone = Trim$(Me.txtApulainenPuhelin.Text)
    email = Trim$(Me.txtApulainenSposti.Text)
    address = Trim$(Me.txtApulainenOsoite.Text)

    If driverName = "" Then
        MsgBox "Apulaisen nimi ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtApulainen.SetFocus
        Exit Sub
    End If

    ' 2. Tarkista duplikaattinimi
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, driverName, NAME_COL)
    If existingRow <> 0 Then
        MsgBox "Apulainen nimellä '" & driverName & "' on jo olemassa rivillä " & existingRow & ".", vbExclamation, "Duplikaatti"
        Me.txtApulainen.SetFocus
        Exit Sub
    End If

    ' 3. Hae seuraava ID
    newID = mdlRegisterUtils.GetNextRegisterID(SHEET_NAME)
    If newID = 0 Then
        MsgBox "Uutta ID:tä ei voitu hakea. Lisäys epäonnistui.", vbCritical, "ID Virhe"
        Exit Sub
    End If

    ' 4. Kokoa data lisäystä varten
    data(1) = newID
    data(2) = driverName
    data(3) = phone
    data(4) = email
    data(5) = address

    ' 5. Lisää tieto välilehdelle
    If mdlRegisterUtils.AddRegisterItem(SHEET_NAME, data) Then
        ' 6. Päivitä UI
        LoadRegisterDataToListBox Me.lstApulaiset, SHEET_NAME ' Päivitä lista
        ' Tyhjennä kentät
        Me.txtApulainen.Text = ""
        Me.txtApulainenPuhelin.Text = ""
        Me.txtApulainenSposti.Text = ""
        Me.txtApulainenOsoite.Text = ""
        Me.lstApulaiset.listIndex = -1 ' Poista valinta
        Me.cmdMuokkaaApulainen.Enabled = False ' Varmista napin tila
        Me.cmdPoistaApulainen.Enabled = False
        MsgBox "Uusi apulainen '" & driverName & "' (ID: " & newID & ") lisätty onnistuneesti.", vbInformation, "Lisäys Onnistui"
    End If

CleanExit_SaveDriver:
    Exit Sub

SaveDriverError:
    MsgBox "Odottamaton virhe tallennettaessa apulaista:" & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    Resume CleanExit_SaveDriver
End Sub


Private Sub cmdMuokkaaApulainen_Click() ' TALLENNA MUUTOKSET OLEMASSA OLEVAAN
    Dim driverName As String, phone As String, email As String, address As String
    Dim itemID As Long
    Dim listIndex As Long
    Dim existingRow As Long
    Dim existingID As Long
    Dim data(1 To 5) As Variant ' ID, Nimi, Puh, Sposti, Osoite
    Const SHEET_NAME As String = "Apulaiset"
    Dim ws As Worksheet ' Tarvitaan duplikaatin ID:n tarkistukseen

    On Error GoTo EditDriverError

    ' 1. Varmista, että jokin on valittuna ja hae ID
    listIndex = Me.lstApulaiset.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse muokattava apulainen listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    On Error Resume Next ' Virheenkäsittely ID:n luvussa
    itemID = CLng(Me.lstApulaiset.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Muokkaus epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo EditDriverError
        Exit Sub
    End If
    On Error GoTo EditDriverError ' Palauta normaali

    ' 2. Lue ja validoi MUOKATUT tiedot kentistä
    driverName = Trim$(Me.txtApulainen.Text)
    phone = Trim$(Me.txtApulainenPuhelin.Text)
    email = Trim$(Me.txtApulainenSposti.Text)
    address = Trim$(Me.txtApulainenOsoite.Text)

    If driverName = "" Then
        MsgBox "Apulaisen nimi ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtApulainen.SetFocus
        Exit Sub
    End If

    ' 3. Tarkista duplikaattinimi (mutta salli sama nimi ITSELLÄÄN)
    existingRow = mdlRegisterUtils.FindRowByValue(SHEET_NAME, driverName, NAME_COL)
    If existingRow <> 0 Then
        ' Nimi löytyi, tarkista onko se eri ID kuin muokattavalla
        Set ws = ThisWorkbook.Worksheets(SHEET_NAME) ' Hae worksheet-objekti
        If Not ws Is Nothing Then
            On Error Resume Next ' ID:n luku voi epäonnistua
            existingID = CLng(ws.Cells(existingRow, ID_COL).value)
            On Error GoTo EditDriverError ' Palauta normaali
            If existingID <> itemID Then
                 MsgBox "Apulainen nimellä '" & driverName & "' on jo olemassa (ID: " & existingID & "). Anna eri nimi.", vbExclamation, "Duplikaatti"
                 Me.txtApulainen.SetFocus
                 Set ws = Nothing
                 Exit Sub
            End If
        End If
        Set ws = Nothing
    End If

    ' 4. Kokoa data päivitystä varten (käytä alkuperäistä itemID:tä)
    data(1) = itemID ' Tärkeää: ID säilyy samana
    data(2) = driverName
    data(3) = phone
    data(4) = email
    data(5) = address

    ' 5. Päivitä tieto välilehdelle
    If mdlRegisterUtils.UpdateRegisterItem(SHEET_NAME, itemID, data) Then
        ' 6. Päivitä UI
        LoadRegisterDataToListBox Me.lstApulaiset, SHEET_NAME ' Päivitä lista
        ' Tyhjennä kentät ja deaktivoi napit
        Me.txtApulainen.Text = ""
        Me.txtApulainenPuhelin.Text = ""
        Me.txtApulainenSposti.Text = ""
        Me.txtApulainenOsoite.Text = ""
        Me.lstApulaiset.listIndex = -1
        Me.cmdMuokkaaApulainen.Enabled = False
        Me.cmdPoistaApulainen.Enabled = False
        MsgBox "Apulaisen '" & driverName & "' (ID: " & itemID & ") tiedot päivitetty onnistuneesti.", vbInformation, "Päivitys Onnistui"
    End If

CleanExit_EditDriver:
    Set ws = Nothing
    Exit Sub

EditDriverError:
     MsgBox "Odottamaton virhe muokattaessa apulaista:" & vbCrLf & Err.Description, vbCritical, "Muokkausvirhe"
     Resume CleanExit_EditDriver
End Sub


Private Sub cmdPoistaApulainen_Click()
    Dim itemID As Long
    Dim driverName As String
    Dim listIndex As Long
    Dim response As VbMsgBoxResult
    Const SHEET_NAME As String = "Apulaiset"

    On Error GoTo DeleteDriverError

    ' 1. Varmista valinta ja hae ID + Nimi
    listIndex = Me.lstApulaiset.listIndex
    If listIndex = -1 Then
        MsgBox "Valitse poistettava apulainen listasta.", vbExclamation, "Ei Valintaa"
        Exit Sub
    End If

    On Error Resume Next ' ID:n luku
    itemID = CLng(Me.lstApulaiset.Column(0, listIndex))
    If Err.Number <> 0 Then
        MsgBox "Valitun rivin ID:tä ei voitu lukea. Poisto epäonnistui.", vbCritical, "ID Virhe"
        On Error GoTo DeleteDriverError
        Exit Sub
    End If
    On Error GoTo DeleteDriverError ' Palauta normaali
    driverName = Me.lstApulaiset.Column(1, listIndex)

    ' 2. Varmista poisto
    response = MsgBox("Haluatko varmasti poistaa apulaisen:" & vbCrLf & vbCrLf & _
                      "ID: " & itemID & vbCrLf & _
                      "Nimi: " & driverName & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    If response = vbNo Then Exit Sub

    ' 3. Poista tieto välilehdeltä
    If mdlRegisterUtils.DeleteRegisterItem(SHEET_NAME, itemID) Then
        ' 4. Päivitä UI
        LoadRegisterDataToListBox Me.lstApulaiset, SHEET_NAME ' Päivitä lista
        ' Tyhjennä kentät ja deaktivoi napit
        Me.txtApulainen.Text = ""
        Me.txtApulainenPuhelin.Text = ""
        Me.txtApulainenSposti.Text = ""
        Me.txtApulainenOsoite.Text = ""
        Me.lstApulaiset.listIndex = -1
        Me.cmdMuokkaaApulainen.Enabled = False
        Me.cmdPoistaApulainen.Enabled = False
        MsgBox "Apulainen '" & driverName & "' (ID: " & itemID & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"
    End If

CleanExit_DeleteDriver:
    Exit Sub

DeleteDriverError:
     MsgBox "Odottamaton virhe poistettaessa apulaista:" & vbCrLf & Err.Description, vbCritical, "Poistovirhe"
     Resume CleanExit_DeleteDriver
End Sub

' Sulje -nappi
Private Sub cmdCancel_Click()
    Unload Me
End Sub

