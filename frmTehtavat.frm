VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTehtavat 
   Caption         =   "Tehtävät"
   ClientHeight    =   11448
   ClientLeft      =   1092
   ClientTop       =   4332
   ClientWidth     =   15264
   OleObjectBlob   =   "frmTehtavat.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmTehtavat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Julkinen muuttuja ID:lle (pysyy samana)
Public TaskIDToEdit As Long

' --- Suoritetaan, kun lomake ladataan muistiin (ennen näyttämistä) ---
Private Sub UserForm_Initialize()
    On Error GoTo Initialize_Error ' Lisätty rakenteellinen virheenkäsittely

    'Debug.Print "frmTehtavat Initialize - TaskIDToEdit: " & TaskIDToEdit ' Voit kommentoida myöhemmin

    ' 1. Tyhjennetään kontrollit
    mdlClearForm.ClearForm Me
    ' Poistettu 'On Error GoTo 0'

    ' 2. Asetetaan oletusvalinnat OptionButtoneille
    Me.optTilaTarjous.value = True
    Me.optApulaisetTilattuEiTarvita.value = True
    Me.optHissiEiTarvita.value = True
    Me.optPysakointilupaEiTarvita.value = True
    Me.optLaivalippuEiTarvita.value = True

    ' --- 3. TÄYTETÄÄN LISTAT JA VALIKOT APUVÄLILEHDILTÄ ---
    ' Oletetaan, että Populate-rutiinit käsittelevät itse virheet (esim. puuttuva välilehti)
    PopulateListBox Me.lstPalvelut, "Palvelut", "B"
    PopulateListBox Me.lstKuljettajat, "Kuljettajat", "B"
    PopulateListBox Me.lstAutot, "Autot", "B"
    PopulateListBox Me.lstKontit, "Kontit", "B"
    PopulateListBox Me.lstApulaiset, "Apulaiset", "B"

    ' --- 4. Asetetaan lomakkeen otsikko ---
    ' Activate-metodi asettaa lopullisen otsikon, mutta jokin oletus voi olla tässä
    If Me.TaskIDToEdit > 0 Then
         Me.Caption = "Muokataan tehtävää..." ' Esim. väliaikainen otsikko
    Else
         Me.Caption = "Lisätään uutta tehtävää..."
    End If

    'Debug.Print "UserForm_Initialize finished. Lists populated." ' Voit kommentoida myöhemmin

CleanExit_Initialize: ' Etiketti siistille poistumiselle
    On Error GoTo 0 ' Nollaa virheenkäsittely ennen poistumista
    Exit Sub

Initialize_Error: ' Virheenkäsittelijä tälle rutiinille
    MsgBox "Virhe alustettaessa tehtävälomaketta:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Lomakkeen Alustusvirhe"
    Resume CleanExit_Initialize ' Yritä poistua siististi virheen jälkeen
End Sub


' Tämä ajetaan, kun lomake aktivoituu (ennen näyttämistä)
Private Sub UserForm_Activate()
    'Debug.Print Now & " frmTehtavat Activate - TaskIDToEdit: " & TaskIDToEdit

    If Me.TaskIDToEdit > 0 Then
        ' --- MUOKKAUSTILA ---
        'Debug.Print Now & " Activate: Muokkaustila, kutsutaan LoadTaskDataIntoForm ID:llä: " & Me.TaskIDToEdit

        If LoadTaskDataIntoForm(Me.TaskIDToEdit) Then ' Tarkista palautusarvo
            ' Lataus onnistui
            Me.Caption = "Muokkaa Tehtävää (ID: " & Me.TaskIDToEdit & ")" ' Aseta otsikko VAIN jos lataus onnistui
            Me.cmdEdit.Enabled = True
            Me.cmdEdit.Visible = True
            Me.cmdSave.Enabled = False
            Me.cmdSave.Visible = False
            Me.txtAsiakas.SetFocus
        Else
            ' Lataus epäonnistui (MsgBox näytetty jo LoadTaskDataIntoForm:ssa)
            Unload Me ' Tuhoa lomake TÄÄLLÄ
            Exit Sub  ' Älä jatka Activate-metodin suoritusta
        End If

    Else
        ' --- LISÄYSTILA ---
        Me.Caption = "Lisää Uusi Tehtävä"
        ' Initialize on jo tyhjentänyt kentät oletusarvoihin
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.Visible = False
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Visible = False
        Me.cmdSave.Enabled = True
        Me.cmdSave.Visible = True
        Me.txtAsiakas.SetFocus
        Me.txtTarjousTehty.value = Format(Date, "dd.mm.yyyy")
    End If
End Sub


Private Sub cmdSave_Click() ' UUDEN LISÄYS
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim taskData As clsTaskItem ' Luodaan uusi
    Const STORAGE_SHEET As String = mdlMain.STORAGE_SHEET_NAME
    Const DISPLAY_SHEET As String = mdlMain.DISPLAY_SHEET_NAME
    Dim tempDate As Variant
    Dim lastausPaivaDate As Variant ' Muuttuja lastauspäivälle
    Dim lastausPaivaText As String ' Lisätty

    On Error GoTo SaveErrorHandler

    ' --- VALIDointi: Asiakas (pakollinen) ---
    If Trim$(Me.txtAsiakas.Text) = "" Then
        MsgBox "Asiakkaan nimi on pakollinen tieto.", vbExclamation, "Pakollinen tieto"
        Me.txtAsiakas.SetFocus
        Exit Sub ' Poistutaan, käyttäjän on syötettävä nimi ja painettava Tallenna uudelleen
    End If

    ' --- VALIDointi: Lastauspäivä (jos annettu, oltava validi) ---
    lastausPaivaText = Trim$(Me.txtLastauspaiva.Text)
    If lastausPaivaText <> "" Then ' Tarkistetaan vain, jos kentässä on jotain
        lastausPaivaDate = mdlDateUtils.ConvertToDate(lastausPaivaText)
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Antamasi lastauspäivä '" & lastausPaivaText & "' ei ole kelvollinen." & vbCrLf & _
                   "Syötä päivämäärä muodossa pp.kk.vvvv tai jätä kenttä tyhjäksi, jos kyseessä on kontaktimerkintä.", vbExclamation, "Virheellinen Lastauspäivä"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistutaan, käyttäjän on korjattava tai tyhjennettävä ja painettava Tallenna uudelleen
        End If
    End If

    ' Hae Manager-oliot
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa.", vbCritical, "Virhe"
        Exit Sub
    End If

    ' Luo TaskItem-olio UUDELLE tiedolle
    Set taskData = New clsTaskItem

    ' --- 3. Lue KAIKKI tiedot lomakkeen kontrolleista taskData-olioon ---
    taskData.kontaktiPaiva = Date

    ' --- Lastauspäivän käsittely ja RecordType/Tila määritys ---
    lastausPaivaText = Trim$(Me.txtLastauspaiva.Text)
    lastausPaivaDate = mdlDateUtils.ConvertToDate(lastausPaivaText) ' Muuttuja lastausPaivaDate on jo esitelty
    
    If lastausPaivaText = "" Then
        ' Lastauspäivä on tyhjä -> Kontakti
        taskData.RecordType = "Kontakti"
        taskData.Tila = "KONTAKTI"
        taskData.lastausPaiva = Null ' Varmistetaan, että clsTaskItem.lastausPaiva on Null
    Else
        ' Lastauspäivä-kentässä on syötettä, validoi se
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Syötä lastauspäivä oikeassa muodossa (pp.kk.vvvv) tai jätä kenttä tyhjäksi kontaktimerkintää varten.", vbExclamation, "Virheellinen Lastauspäivä"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistu, kunnes lastauspäivä on validi tai tyhjä
        Else
            ' Lastauspäivä on validi -> Task
            taskData.RecordType = "Task"
            taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppinä
    
            ' Aseta Tila Task-tyypille (tämä lohko siirretään tähän)
            If Me.fraTila.optTilaHyvaksytty.value = True Then
                taskData.Tila = "HYVÄKSYTTY"
            ElseIf Me.fraTila.optTilaTarjous.value = True Then
                taskData.Tila = "TARJOUS"
            Else
                taskData.Tila = "TARJOUS" ' Oletus Task-tyypille
            End If
        End If
    End If

    taskData.asiakas = Me.txtAsiakas.Text
    taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppinä

    ' Muut päivämäärät (Tallenna Date-tyyppinä tai Null/Empty)
    tempDate = mdlDateUtils.ConvertToDate(Me.txtPurkupaiva.Text)
    If IsDate(tempDate) Then taskData.purkuPaiva = tempDate Else taskData.purkuPaiva = Null ' Tai Empty

    tempDate = mdlDateUtils.ConvertToDate(Me.txtLastausLoppuu.Text)
    If IsDate(tempDate) Then taskData.LastausLoppuu = tempDate Else taskData.LastausLoppuu = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtPurkuLoppuu.Text)
    If IsDate(tempDate) Then taskData.PurkuLoppuu = tempDate Else taskData.PurkuLoppuu = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtTarjousTehty.Text)
    If IsDate(tempDate) Then taskData.tarjousTehty = tempDate Else taskData.tarjousTehty = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtTarjousHyvaksytty.Text)
    If IsDate(tempDate) Then taskData.TarjousHyvaksytty = tempDate Else taskData.TarjousHyvaksytty = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtTarjousHylatty.Text)
    If IsDate(tempDate) Then taskData.TarjousHylatty = tempDate Else taskData.TarjousHylatty = Null

    ' Muut tekstikentät
    taskData.sahkoposti = LCase(Me.txtSahkoposti.Text)
    taskData.lastausMaa = UCase(Me.txtLastausmaa.Text)
    taskData.purkuMaa = UCase(Me.txtPurkumaa.Text)
    taskData.M3m = Me.txtM3m.Text
    taskData.Huomioitavaa = Me.txtHuomioitavaa.Text
    taskData.puhelin = Me.txtPuhelin.Text
    taskData.lastausOsoite = Me.txtLastausosoite.Text
    taskData.purkuOsoite = Me.txtPurkuosoite.Text
    taskData.Vakuutus = Me.txtVakuutus.Text
    taskData.Arvo = Me.txtArvo.Text ' Käsittele Varianttina
    taskData.hinta = Me.txtHinta.Text ' Käsittele Varianttina
    taskData.M3t = Me.txtM3t.Text
    taskData.valimatka = Me.txtValimatka.Text

    ' OptionButtonit
    If Me.fraTila.optTilaHyvaksytty.value = True Then
        taskData.Tila = "HYVÄKSYTTY"
    ElseIf Me.fraTila.optTilaTarjous.value = True Then
        taskData.Tila = "TARJOUS"
    Else
        taskData.Tila = "TARJOUS" ' Oletus
    End If
    ' ... muut optionbutton ryhmät ...
     If Me.optApulaisetTilattuOk.value Then
        taskData.ApulaisetTilattu = "OK"
    ElseIf Me.optApulaisetTilattuTarvitaan.value Then
        taskData.ApulaisetTilattu = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.ApulaisetTilattu = "EI TARVITA"
    End If

    If Me.optHissiOk.value Then
        taskData.hissi = "OK"
    ElseIf Me.optHissiTarvitaan.value Then
        taskData.hissi = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.hissi = "EI TARVITA"
    End If

    If Me.optPysakointilupaOk.value Then
        taskData.Pysakointilupa = "OK"
    ElseIf Me.optPysakointilupaTarvitaan.value Then
        taskData.Pysakointilupa = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.Pysakointilupa = "EI TARVITA"
    End If

    If Me.optLaivalippuOk.value Then
        taskData.Laivalippu = "OK"
    ElseIf Me.optLaivalippuTarvitaan.value Then
        taskData.Laivalippu = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.Laivalippu = "EI TARVITA"
    End If

    ' CheckBoxit
    taskData.LastauspaivaVarmistunut = Me.chkLastauspaivaVarmistunut.value
    taskData.PurkupaivaVarmistunut = Me.chkPurkupaivaVarmistunut.value
    taskData.Rahtikirja = Me.chkRahtikirja.value
    taskData.Laskutus = Me.chkLaskutus.value
    taskData.Muuttomaailma = Me.chkMuuttomaailma.value

    ' ListBoxit
    taskData.palvelu = mdlStringUtils.GetListBoxMultiSelection(Me.lstPalvelut, ";")
    taskData.Kuljettajat = mdlStringUtils.GetListBoxMultiSelection(Me.lstKuljettajat, ";")
    taskData.Autot = mdlStringUtils.GetListBoxMultiSelection(Me.lstAutot, ";")
    taskData.Kontit = mdlStringUtils.GetListBoxMultiSelection(Me.lstKontit, ";")
    taskData.Apulaiset = mdlStringUtils.GetListBoxMultiSelection(Me.lstApulaiset, ";")

    taskData.AttentionSortDate = Null ' Varmista tyhjäksi

    ' --- 4. Lisää UUSI tehtävä TaskManagerin kokoelmaan ---
    tm.AddTask taskData ' Antaa ID:n

    ' --- 5. Tallenna KOKO kokoelma takaisin välilehdelle ---
    Application.StatusBar = "Tallennetaan uutta tehtävää..."
    tm.SaveToSheet STORAGE_SHEET ' Olettaa, että SaveToSheet osaa kirjoittaa Date/Null

    ' --- 6. Päivitä näyttö ---
    Application.StatusBar = "Päivitetään näyttöä..."
    dm.UpdateDisplay tm.tasks, DISPLAY_SHEET

    ' --- 7. Tarjous ---
    Call mdlTarjousUtils.LuoTarjousLomakkeelta(Me)

    Application.StatusBar = False
    'MsgBox "Uusi tehtävä (ID: " & taskData.ID & ") tallennettu onnistuneesti!", vbInformation

    ' --- 8. Sulje lomake ---
    Unload Me
    Exit Sub

SaveErrorHandler:
    Application.StatusBar = False
    Set tm = Nothing ' Vapauta viittaukset
    Set dm = Nothing
    Set taskData = Nothing
    MsgBox "Virhe tallennettaessa uutta tehtävää:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
End Sub


' --- Korvaa tämä kokonaan frmTehtavat.txt -tiedostoon ---
Private Sub cmdEdit_Click() ' MUOKKAUKSEN TALLENNUS
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim taskData As clsTaskItem ' Haetaan olemassa oleva päivitystä varten
    Const STORAGE_SHEET As String = mdlMain.STORAGE_SHEET_NAME
    Const DISPLAY_SHEET As String = mdlMain.DISPLAY_SHEET_NAME
    Dim tempDate As Variant
    Dim lastausPaivaDate As Variant ' Muuttuja lastauspäivälle
    Dim lastausPaivaText As String ' Lisätty

    On Error GoTo EditErrorHandler

    ' --- VALIDointi: Asiakas (pakollinen) ---
    If Trim$(Me.txtAsiakas.Text) = "" Then
        MsgBox "Asiakkaan nimi on pakollinen tieto.", vbExclamation, "Pakollinen tieto"
        Me.txtAsiakas.SetFocus
        Exit Sub ' Poistutaan, käyttäjän on syötettävä nimi ja painettava Tallenna uudelleen
    End If

    ' --- VALIDointi: Lastauspäivä (jos annettu, oltava validi) ---
    lastausPaivaText = Trim$(Me.txtLastauspaiva.Text)
    If lastausPaivaText <> "" Then ' Tarkistetaan vain, jos kentässä on jotain
        lastausPaivaDate = mdlDateUtils.ConvertToDate(lastausPaivaText)
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Antamasi lastauspäivä '" & lastausPaivaText & "' ei ole kelvollinen." & vbCrLf & _
                   "Syötä päivämäärä muodossa pp.kk.vvvv tai jätä kenttä tyhjäksi, jos kyseessä on kontaktimerkintä.", vbExclamation, "Virheellinen Lastauspäivä"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistutaan, käyttäjän on korjattava tai tyhjennettävä ja painettava Tallenna uudelleen
        End If
    End If


    ' Hae Manager-oliot
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa.", vbCritical, "Virhe"
        Exit Sub
    End If

    ' --- 3. Hae olemassa oleva TaskItem ID:n perusteella ---
    If Me.TaskIDToEdit <= 0 Then
        MsgBox "Muokattavan tehtävän ID:tä ei löytynyt!", vbCritical
        Exit Sub
    End If
    Set taskData = tm.GetTaskByID(Me.TaskIDToEdit) ' Hae muokattava olio
    If taskData Is Nothing Then
        MsgBox "Muokattavaa tehtävää ID:llä " & Me.TaskIDToEdit & " ei löytynyt muistista!", vbCritical
        Exit Sub
    End If

     ' --- 4. Päivitä tiedot lomakkeen kontrolleista taskData-olioon ---

    ' Pakolliset kentät (jo tarkistettu)
    taskData.asiakas = Me.txtAsiakas.Text
    taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppinä
    
    ' Lastauspäivän ja RecordType/Tila -logiikka
   'Huom: taskData.kontaktiPaiva EI PÄIVITETÄ TÄSSÄ
    If Trim$(Me.txtLastauspaiva.Text) = "" Then
        ' Lastauspäivä on tyhjä -> Kontakti
        taskData.RecordType = "Kontakti"
        taskData.Tila = "KONTAKTI"
        taskData.lastausPaiva = Null
    Else
        ' Lastauspäivä-kentässä on syötettä, validoi se (lastausPaivaDate on jo alustettu)
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Syötä lastauspäivä oikeassa muodossa (pp.kk.vvvv) tai jätä kenttä tyhjäksi kontaktimerkintää varten.", vbExclamation, "Virheellinen Lastauspäivä"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistu, kunnes lastauspäivä on validi tai tyhjä
        Else
            ' Lastauspäivä on validi -> Task
            taskData.RecordType = "Task"
            taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppinä
            ' Aseta Tila Task-tyypille
            If Me.fraTila.optTilaHyvaksytty.value = True Then
                taskData.Tila = "HYVÄKSYTTY"
            ElseIf Me.fraTila.optTilaTarjous.value = True Then
                taskData.Tila = "TARJOUS"
            Else
                taskData.Tila = "TARJOUS" ' Oletus Task-tyypille
            End If
        End If
    End If

    ' Muut päivämäärät (Tallenna Date-tyyppinä tai Null/Empty)
    tempDate = mdlDateUtils.ConvertToDate(Me.txtPurkupaiva.Text)
    If IsDate(tempDate) Then taskData.purkuPaiva = tempDate Else taskData.purkuPaiva = Null ' Tai Empty

    tempDate = mdlDateUtils.ConvertToDate(Me.txtLastausLoppuu.Text)
    If IsDate(tempDate) Then taskData.LastausLoppuu = tempDate Else taskData.LastausLoppuu = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtPurkuLoppuu.Text)
    If IsDate(tempDate) Then taskData.PurkuLoppuu = tempDate Else taskData.PurkuLoppuu = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtTarjousTehty.Text)
    If IsDate(tempDate) Then taskData.tarjousTehty = tempDate Else taskData.tarjousTehty = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtTarjousHyvaksytty.Text)
    If IsDate(tempDate) Then taskData.TarjousHyvaksytty = tempDate Else taskData.TarjousHyvaksytty = Null

    tempDate = mdlDateUtils.ConvertToDate(Me.txtTarjousHylatty.Text)
    If IsDate(tempDate) Then taskData.TarjousHylatty = tempDate Else taskData.TarjousHylatty = Null

    ' Muut tekstikentät
    taskData.sahkoposti = Me.txtSahkoposti.Text
    taskData.lastausMaa = UCase(Me.txtLastausmaa.Text)
    taskData.purkuMaa = UCase(Me.txtPurkumaa.Text)
    taskData.M3m = Me.txtM3m.Text
    taskData.Huomioitavaa = Me.txtHuomioitavaa.Text
    taskData.puhelin = Me.txtPuhelin.Text
    taskData.lastausOsoite = Me.txtLastausosoite.Text
    taskData.purkuOsoite = Me.txtPurkuosoite.Text
    taskData.Vakuutus = Me.txtVakuutus.Text
    taskData.Arvo = Me.txtArvo.Text ' Käsittele Varianttina
    taskData.hinta = Me.txtHinta.Text ' Käsittele Varianttina
    taskData.M3t = Me.txtM3t.Text
    taskData.valimatka = Me.txtValimatka.Text

    ' OptionButtonit
    If Me.fraTila.optTilaHyvaksytty.value = True Then
        taskData.Tila = "HYVÄKSYTTY"
    ElseIf Me.fraTila.optTilaTarjous.value = True Then
        taskData.Tila = "TARJOUS"
    Else
        taskData.Tila = "TARJOUS" ' Oletus
    End If
    ' ... muut optionbutton ryhmät ...
     If Me.optApulaisetTilattuOk.value Then
        taskData.ApulaisetTilattu = "OK"
    ElseIf Me.optApulaisetTilattuTarvitaan.value Then
        taskData.ApulaisetTilattu = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.ApulaisetTilattu = "EI TARVITA"
    End If

    If Me.optHissiOk.value Then
        taskData.hissi = "OK"
    ElseIf Me.optHissiTarvitaan.value Then
        taskData.hissi = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.hissi = "EI TARVITA"
    End If

    If Me.optPysakointilupaOk.value Then
        taskData.Pysakointilupa = "OK"
    ElseIf Me.optPysakointilupaTarvitaan.value Then
        taskData.Pysakointilupa = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.Pysakointilupa = "EI TARVITA"
    End If

    If Me.optLaivalippuOk.value Then
        taskData.Laivalippu = "OK"
    ElseIf Me.optLaivalippuTarvitaan.value Then
        taskData.Laivalippu = "TARVITAAN"
    Else ' Oletus tai "Ei tarvita"
        taskData.Laivalippu = "EI TARVITA"
    End If

    ' CheckBoxit
    taskData.LastauspaivaVarmistunut = Me.chkLastauspaivaVarmistunut.value
    taskData.PurkupaivaVarmistunut = Me.chkPurkupaivaVarmistunut.value
    taskData.Rahtikirja = Me.chkRahtikirja.value
    taskData.Laskutus = Me.chkLaskutus.value
    taskData.Muuttomaailma = Me.chkMuuttomaailma.value

    ' ListBoxit
    taskData.palvelu = mdlStringUtils.GetListBoxMultiSelection(Me.lstPalvelut, ";")
    taskData.Kuljettajat = mdlStringUtils.GetListBoxMultiSelection(Me.lstKuljettajat, ";")
    taskData.Autot = mdlStringUtils.GetListBoxMultiSelection(Me.lstAutot, ";")
    taskData.Kontit = mdlStringUtils.GetListBoxMultiSelection(Me.lstKontit, ";")
    taskData.Apulaiset = mdlStringUtils.GetListBoxMultiSelection(Me.lstApulaiset, ";")

    ' RecordType (pysyy samana)
    taskData.AttentionSortDate = Null ' Varmista tyhjäksi

    ' --- 5. Päivitä tehtävä TaskManagerin kokoelmaan ---
    tm.UpdateTask taskData ' Kutsu PÄIVITYSmetodia

    ' --- 6. Tallenna KOKO kokoelma takaisin välilehdelle ---
    Application.StatusBar = "Tallennetaan muutoksia..."
    tm.SaveToSheet STORAGE_SHEET ' Olettaa, että SaveToSheet osaa kirjoittaa Date/Null

    ' --- 7. Päivitä näyttö ---
    Application.StatusBar = "Päivitetään näyttöä..."
    dm.UpdateDisplay tm.tasks, DISPLAY_SHEET

    ' --- 8. Tarjous ---
    Call mdlTarjousUtils.LuoTarjousLomakkeelta(Me)
    
    Application.StatusBar = False
    'MsgBox "Muutokset tallennettu onnistuneesti ID:lle " & taskData.ID, vbInformation

    ' --- 9. Sulje lomake ---
    Unload Me
    Exit Sub

EditErrorHandler:
    Application.StatusBar = False
    Set tm = Nothing ' Vapauta viittaukset
    Set dm = Nothing
    Set taskData = Nothing
    MsgBox "Virhe tallennettaessa muutoksia:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Muokkausvirhe"
End Sub

Private Sub cmdDelete_Click()
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim taskIDToDelete As Long
    Dim response As VbMsgBoxResult
    Const STORAGE_SHEET As String = mdlMain.STORAGE_SHEET_NAME
    Const DISPLAY_SHEET As String = mdlMain.DISPLAY_SHEET_NAME

    taskIDToDelete = Me.TaskIDToEdit

    If taskIDToDelete <= 0 Then
        MsgBox "Poistettavan tehtävän ID:tä ei voitu määrittää.", vbExclamation
        Exit Sub
    End If

    response = MsgBox("Haluatko varmasti poistaa tämän tehtävän (ID: " & taskIDToDelete & ")?" & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista poisto")

    If response = vbNo Then Exit Sub

    On Error GoTo DeleteErrorHandler

    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Poisto epäonnistui.", vbCritical, "Virhe"
        GoTo CleanExit_Delete
    End If

    tm.DeleteTask taskIDToDelete

    Application.StatusBar = "Poistetaan ja tallennetaan..."
    tm.SaveToSheet STORAGE_SHEET

    Application.StatusBar = "Päivitetään näyttöä..."
    dm.UpdateDisplay tm.tasks, DISPLAY_SHEET
    
    Application.StatusBar = False
    MsgBox "Tehtävä (ID: " & taskIDToDelete & ") poistettu onnistuneesti.", vbInformation

    Unload Me
    
CleanExit_Delete:
    Application.StatusBar = False
    Exit Sub

DeleteErrorHandler:
    MsgBox "Virhe poistettaessa tehtävää (ID: " & taskIDToDelete & "):" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Poistovirhe"
    Resume CleanExit_Delete
End Sub
    
' --- Apurutiini ComboBoxin täyttämiseen ---
Private Sub PopulateComboBox(cmb As MSForms.ComboBox, sheetName As String, colLetter As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    'On Error Resume Next ' Käsittele virheet (esim. välilehteä ei löydy)
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "Varoitus: Välilehteä '" & sheetName & "' ei löytynyt ComboBoxin täyttöä varten."
        Exit Sub
    End If
    ' Hae data-alue määritellystä sarakkeesta (alkaen riviltä 1)
    Set rng = ws.Range(colLetter & "1", ws.Cells(ws.rows.Count, colLetter).End(xlUp))
    If rng Is Nothing Or rng.rows.Count = 0 Then Exit Sub

    cmb.Clear ' Tyhjennä vanhat valinnat
    For Each cell In rng.Cells ' Käy läpi vain solut alueelta
        If Trim$(cell.value) <> "" Then ' Lisää vain ei-tyhjät arvot
            cmb.AddItem cell.value
        End If
    Next cell
    On Error GoTo 0 ' Palauta normaali virheiden käsittely
End Sub

' --- Apurutiini ListBoxin täyttämiseen ---
Private Sub PopulateListBox(lst As MSForms.ListBox, sheetName As String, colLetter As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    On Error GoTo PopulateError ' Lisätään virheenkäsittelijä

    Set ws = Nothing
    ' Yritä hakea välilehti. Jos ei löydy, PopulateError hoitaa.
    Set ws = ThisWorkbook.Worksheets(sheetName)

    ' Etsi viimeinen rivi
    lastRow = ws.Cells(ws.rows.Count, colLetter).End(xlUp).row
    
    ' Jos dataa ei ole (vain otsikko tai tyhjä), tai jos lastRow on virheellisesti 0 tai 1
    If lastRow < 2 Then ' Oletetaan, että data alkaa aina vähintään riviltä 2
        ' Debug.Print "PopulateListBox: Ei ladattavaa dataa välilehdeltä '" & sheetName & "', sarakkeesta " & colLetter
        GoTo CleanExit ' Poistu siististi, jos ei dataa
    End If

    Set rng = ws.Range(colLetter & "2:" & colLetter & lastRow)
    ' If rng Is Nothing Then GoTo CleanExit ' Tämä tarkistus on epätodennäköinen, jos lastRow >=2

    lst.Clear
    For Each cell In rng.Cells
         ' Käytä CStr(cell.value) varmistaaksesi, että arvo käsitellään merkkijonona
         ' ja vältä mahdollinen virhe, jos cell.value on esim. virhearvo (#N/A)
         Dim cellValueStr As String
         On Error Resume Next ' Ohita virhe, jos solun arvoa ei voi muuntaa stringiksi
         cellValueStr = CStr(cell.value)
         If Err.Number <> 0 Then
             cellValueStr = "" ' Aseta tyhjäksi, jos muunnos epäonnistui
             Err.Clear
         End If
         On Error GoTo PopulateError ' Palauta päävirheenkäsittelijä

         If Trim$(cellValueStr) <> "" Then
            lst.AddItem cellValueStr
         End If
    Next cell

CleanExit:
    On Error GoTo 0 ' Palauta normaali virheenkäsittely
    Set ws = Nothing
    Set rng = Nothing
    Set cell = Nothing
    Exit Sub

PopulateError:
    Debug.Print "Virhe PopulateListBox: Välilehti='" & sheetName & "', Sarake='" & colLetter & "', ListBox='" & lst.Name & "'. Virhe: " & Err.Description
    ' Tässä voisi olla myös MsgBox käyttäjälle, jos listan täyttö on kriittistä lomakkeen toiminnalle
    ' Esim. MsgBox "Listan '" & lst.Name & "' täyttö välilehdeltä '" & sheetName & "' epäonnistui." & vbCrLf & "Virhe: " & Err.Description, vbExclamation
    Resume CleanExit ' Yritä siivota ja poistua
End Sub

' --- Apurutiini olemassa olevan tehtävän tietojen lataamiseksi lomakkeelle ---
Private Function LoadTaskDataIntoForm(taskID As Long) As Boolean
    Dim tm As clsTaskManager
    Dim taskToEdit As clsTaskItem
    LoadTaskDataIntoForm = False
    On Error GoTo LoadDataError

    Set tm = mdlMain.GetTaskManagerInstance()
    'Debug.Print Now & " LoadTaskDataIntoForm: TaskManagerin kokoelman koko HETI haun jälkeen: " & IIf(tm Is Nothing, "tm on Nothing", tm.tasks.Count)

    If tm Is Nothing Then
        'Debug.Print Now & " KRIITTINEN VIRHE: TaskManager-instanssi on Nothing LoadTaskDataIntoFormissa!"
        MsgBox "Kriittinen virhe: Tehtävänhallintaa ei voitu alustaa!", vbCritical
        Exit Function
    End If
    'Debug.Print Now & " TaskManagerin kokoelman koko ennen hakua: " & tm.tasks.Count

    Set taskToEdit = tm.GetTaskByID(taskID)

    If taskToEdit Is Nothing Then
        'Debug.Print Now & " LoadTaskDataIntoForm: GetTaskByID palautti Nothing ID:lle " & taskID
        MsgBox "Muokattavaa tehtävää ID:llä " & taskID & " ei löytynyt muistista!", vbExclamation, "Virhe"
        Exit Function
    'Else
        'Debug.Print Now & " LoadTaskDataIntoForm: Tehtävä löytyi ID:llä " & taskID
    End If

    ' --- Täytä lomakkeen kentät haetun olion tiedoilla ---
    Me.txtAsiakas.Text = mdlStringUtils.DefaultIfNull(taskToEdit.asiakas, "")
    Me.txtSahkoposti.Text = mdlStringUtils.DefaultIfNull(taskToEdit.sahkoposti, "")
    Me.txtTarjousTehty.Text = mdlDateUtils.FormatDateToString(taskToEdit.tarjousTehty, "")
    Me.txtLastauspaiva.Text = mdlDateUtils.FormatDateToString(taskToEdit.lastausPaiva, "")
    Me.txtLastausmaa.Text = mdlStringUtils.DefaultIfNull(taskToEdit.lastausMaa, "")
    Me.txtPurkumaa.Text = mdlStringUtils.DefaultIfNull(taskToEdit.purkuMaa, "")
    Me.txtPurkupaiva.Text = mdlDateUtils.FormatDateToString(taskToEdit.purkuPaiva, "")
    Me.txtM3m.Text = mdlStringUtils.DefaultIfNull(taskToEdit.M3m, "")
    Me.txtHuomioitavaa.Text = mdlStringUtils.DefaultIfNull(taskToEdit.Huomioitavaa, "")
    Me.txtPuhelin.Text = mdlStringUtils.DefaultIfNull(taskToEdit.puhelin, "")
    Me.txtLastausosoite.Text = mdlStringUtils.DefaultIfNull(taskToEdit.lastausOsoite, "")
    Me.txtPurkuosoite.Text = mdlStringUtils.DefaultIfNull(taskToEdit.purkuOsoite, "")
    Me.txtVakuutus.Text = mdlStringUtils.DefaultIfNull(taskToEdit.Vakuutus, "")
    Me.txtArvo.Text = mdlStringUtils.DefaultIfNull(taskToEdit.Arvo, "") ' Variant, DefaultIfNull käsittelee
    Me.txtHinta.Text = mdlStringUtils.DefaultIfNull(taskToEdit.hinta, "") ' Variant, DefaultIfNull käsittelee
    Me.txtTarjousHyvaksytty.Text = mdlDateUtils.FormatDateToString(taskToEdit.TarjousHyvaksytty, "")
    Me.txtTarjousHylatty.Text = mdlDateUtils.FormatDateToString(taskToEdit.TarjousHylatty, "")
    Me.txtM3t.Text = mdlStringUtils.DefaultIfNull(taskToEdit.M3t, "") ' M3t on String clsTaskItem:ssa
    Me.txtValimatka.Text = mdlStringUtils.DefaultIfNull(taskToEdit.valimatka, "")
    Me.txtLastausLoppuu.Text = mdlDateUtils.FormatDateToString(taskToEdit.LastausLoppuu, "")
    Me.txtPurkuLoppuu.Text = mdlDateUtils.FormatDateToString(taskToEdit.PurkuLoppuu, "")
    
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstPalvelut, taskToEdit.palvelu, ";")
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstKuljettajat, taskToEdit.Kuljettajat, ";")
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstAutot, taskToEdit.Autot, ";")
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstKontit, taskToEdit.Kontit, ";")
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstApulaiset, taskToEdit.Apulaiset, ";")

    ' CheckBoxit (oletetaan, että clsTaskItem:ssa nämä ovat Boolean)
    Me.chkLastauspaivaVarmistunut.value = taskToEdit.LastauspaivaVarmistunut
    Me.chkPurkupaivaVarmistunut.value = taskToEdit.PurkupaivaVarmistunut
    ' Seuraavat olivat clsTaskItem:ssa String, mutta frmTehtavat.Save/Edit asettaa ne Boolean-arvoina CheckBoxeista.
    ' Oletetaan, että clsTaskItem.Rahtikirja, .Laskutus, .Muuttomaailma ovat String, jotka voivat olla "KYLLÄ" tai "EI"
    ' tai True/False Boolean-arvoina. Jos ne ovat String, tarvitaan muunnos.
    ' Koska Save/Edit asettaa ne CheckBox.Value (Boolean), oletetaan että ne ovatkin Boolean clsTaskItem:ssa.
    ' Jos ne ovat String, tarvitaan:
    ' Me.chkRahtikirja.value = (UCase(taskToEdit.Rahtikirja) = "KYLLÄ" Or UCase(taskToEdit.Rahtikirja) = "TRUE" Or taskToEdit.Rahtikirja = "-1")
    ' Mutta oletetaan nyt että ne ovat Boolean:
    Me.chkRahtikirja.value = CBool(taskToEdit.Rahtikirja) ' CBool muuntaa String "True"/"False" tai numeerisen booleaniksi
    Me.chkLaskutus.value = CBool(taskToEdit.Laskutus)
    Me.chkMuuttomaailma.value = CBool(taskToEdit.Muuttomaailma)


    ' OptionButtonien asetus (ApulaisetTilattu, Hissi, Pysakointilupa, Laivalippu, Tila)
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(taskToEdit.ApulaisetTilattu, "EI TARVITA")))
        Case "EI TARVITA": Me.optApulaisetTilattuEiTarvita.value = True
        Case "TARVITAAN": Me.optApulaisetTilattuTarvitaan.value = True
        Case "OK": Me.optApulaisetTilattuOk.value = True
        Case Else: Me.optApulaisetTilattuEiTarvita.value = True ' Oletus
    End Select
    
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(taskToEdit.hissi, "EI TARVITA")))
        Case "EI TARVITA": Me.optHissiEiTarvita.value = True
        Case "TARVITAAN": Me.optHissiTarvitaan.value = True
        Case "OK": Me.optHissiOk.value = True
        Case Else: Me.optHissiEiTarvita.value = True
    End Select

    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(taskToEdit.Pysakointilupa, "EI TARVITA")))
        Case "EI TARVITA": Me.optPysakointilupaEiTarvita.value = True
        Case "TARVITAAN": Me.optPysakointilupaTarvitaan.value = True
        Case "OK": Me.optPysakointilupaOk.value = True
        Case Else: Me.optPysakointilupaEiTarvita.value = True
    End Select

    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(taskToEdit.Laivalippu, "EI TARVITA")))
        Case "EI TARVITA": Me.optLaivalippuEiTarvita.value = True
        Case "TARVITAAN": Me.optLaivalippuTarvitaan.value = True
        Case "OK": Me.optLaivalippuOk.value = True
        Case Else: Me.optLaivalippuEiTarvita.value = True
    End Select
    
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(taskToEdit.Tila, "TARJOUS")))
        Case "TARJOUS": Me.optTilaTarjous.value = True
        Case "HYVÄKSYTTY": Me.optTilaHyvaksytty.value = True
        Case Else: Me.optTilaTarjous.value = True ' Oletus
    End Select

    LoadTaskDataIntoForm = True
    Exit Function

LoadDataError:
    MsgBox "Odottamaton virhe ladattaessa tehtävän tietoja (ID: " & taskID & "):" & vbCrLf & Err.Description, vbCritical, "Latausvirhe"
End Function

' Sulje -nappi
Private Sub cmdCancel_Click()
    Unload Me
End Sub

