VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTehtavat 
   Caption         =   "Teht�v�t"
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

' --- Suoritetaan, kun lomake ladataan muistiin (ennen n�ytt�mist�) ---
Private Sub UserForm_Initialize()
    On Error GoTo Initialize_Error ' Lis�tty rakenteellinen virheenk�sittely

    'Debug.Print "frmTehtavat Initialize - TaskIDToEdit: " & TaskIDToEdit ' Voit kommentoida my�hemmin

    ' 1. Tyhjennet��n kontrollit
    mdlClearForm.ClearForm Me
    ' Poistettu 'On Error GoTo 0'

    ' 2. Asetetaan oletusvalinnat OptionButtoneille
    Me.optTilaTarjous.value = True
    Me.optApulaisetTilattuEiTarvita.value = True
    Me.optHissiEiTarvita.value = True
    Me.optPysakointilupaEiTarvita.value = True
    Me.optLaivalippuEiTarvita.value = True

    ' --- 3. T�YTET��N LISTAT JA VALIKOT APUV�LILEHDILT� ---
    ' Oletetaan, ett� Populate-rutiinit k�sittelev�t itse virheet (esim. puuttuva v�lilehti)
    PopulateListBox Me.lstPalvelut, "Palvelut", "B"
    PopulateListBox Me.lstKuljettajat, "Kuljettajat", "B"
    PopulateListBox Me.lstAutot, "Autot", "B"
    PopulateListBox Me.lstKontit, "Kontit", "B"
    PopulateListBox Me.lstApulaiset, "Apulaiset", "B"

    ' --- 4. Asetetaan lomakkeen otsikko ---
    ' Activate-metodi asettaa lopullisen otsikon, mutta jokin oletus voi olla t�ss�
    If Me.TaskIDToEdit > 0 Then
         Me.Caption = "Muokataan teht�v��..." ' Esim. v�liaikainen otsikko
    Else
         Me.Caption = "Lis�t��n uutta teht�v��..."
    End If

    'Debug.Print "UserForm_Initialize finished. Lists populated." ' Voit kommentoida my�hemmin

CleanExit_Initialize: ' Etiketti siistille poistumiselle
    On Error GoTo 0 ' Nollaa virheenk�sittely ennen poistumista
    Exit Sub

Initialize_Error: ' Virheenk�sittelij� t�lle rutiinille
    MsgBox "Virhe alustettaessa teht�v�lomaketta:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Lomakkeen Alustusvirhe"
    Resume CleanExit_Initialize ' Yrit� poistua siististi virheen j�lkeen
End Sub


' T�m� ajetaan, kun lomake aktivoituu (ennen n�ytt�mist�)
Private Sub UserForm_Activate()
    'Debug.Print Now & " frmTehtavat Activate - TaskIDToEdit: " & TaskIDToEdit

    If Me.TaskIDToEdit > 0 Then
        ' --- MUOKKAUSTILA ---
        'Debug.Print Now & " Activate: Muokkaustila, kutsutaan LoadTaskDataIntoForm ID:ll�: " & Me.TaskIDToEdit

        If LoadTaskDataIntoForm(Me.TaskIDToEdit) Then ' Tarkista palautusarvo
            ' Lataus onnistui
            Me.Caption = "Muokkaa Teht�v�� (ID: " & Me.TaskIDToEdit & ")" ' Aseta otsikko VAIN jos lataus onnistui
            Me.cmdEdit.Enabled = True
            Me.cmdEdit.Visible = True
            Me.cmdSave.Enabled = False
            Me.cmdSave.Visible = False
            Me.txtAsiakas.SetFocus
        Else
            ' Lataus ep�onnistui (MsgBox n�ytetty jo LoadTaskDataIntoForm:ssa)
            Unload Me ' Tuhoa lomake T��LL�
            Exit Sub  ' �l� jatka Activate-metodin suoritusta
        End If

    Else
        ' --- LIS�YSTILA ---
        Me.Caption = "Lis�� Uusi Teht�v�"
        ' Initialize on jo tyhjent�nyt kent�t oletusarvoihin
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


Private Sub cmdSave_Click() ' UUDEN LIS�YS
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim taskData As clsTaskItem ' Luodaan uusi
    Const STORAGE_SHEET As String = mdlMain.STORAGE_SHEET_NAME
    Const DISPLAY_SHEET As String = mdlMain.DISPLAY_SHEET_NAME
    Dim tempDate As Variant
    Dim lastausPaivaDate As Variant ' Muuttuja lastausp�iv�lle
    Dim lastausPaivaText As String ' Lis�tty

    On Error GoTo SaveErrorHandler

    ' --- VALIDointi: Asiakas (pakollinen) ---
    If Trim$(Me.txtAsiakas.Text) = "" Then
        MsgBox "Asiakkaan nimi on pakollinen tieto.", vbExclamation, "Pakollinen tieto"
        Me.txtAsiakas.SetFocus
        Exit Sub ' Poistutaan, k�ytt�j�n on sy�tett�v� nimi ja painettava Tallenna uudelleen
    End If

    ' --- VALIDointi: Lastausp�iv� (jos annettu, oltava validi) ---
    lastausPaivaText = Trim$(Me.txtLastauspaiva.Text)
    If lastausPaivaText <> "" Then ' Tarkistetaan vain, jos kent�ss� on jotain
        lastausPaivaDate = mdlDateUtils.ConvertToDate(lastausPaivaText)
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Antamasi lastausp�iv� '" & lastausPaivaText & "' ei ole kelvollinen." & vbCrLf & _
                   "Sy�t� p�iv�m��r� muodossa pp.kk.vvvv tai j�t� kentt� tyhj�ksi, jos kyseess� on kontaktimerkint�.", vbExclamation, "Virheellinen Lastausp�iv�"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistutaan, k�ytt�j�n on korjattava tai tyhjennett�v� ja painettava Tallenna uudelleen
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

    ' --- Lastausp�iv�n k�sittely ja RecordType/Tila m��ritys ---
    lastausPaivaText = Trim$(Me.txtLastauspaiva.Text)
    lastausPaivaDate = mdlDateUtils.ConvertToDate(lastausPaivaText) ' Muuttuja lastausPaivaDate on jo esitelty
    
    If lastausPaivaText = "" Then
        ' Lastausp�iv� on tyhj� -> Kontakti
        taskData.RecordType = "Kontakti"
        taskData.Tila = "KONTAKTI"
        taskData.lastausPaiva = Null ' Varmistetaan, ett� clsTaskItem.lastausPaiva on Null
    Else
        ' Lastausp�iv�-kent�ss� on sy�tett�, validoi se
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Sy�t� lastausp�iv� oikeassa muodossa (pp.kk.vvvv) tai j�t� kentt� tyhj�ksi kontaktimerkint�� varten.", vbExclamation, "Virheellinen Lastausp�iv�"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistu, kunnes lastausp�iv� on validi tai tyhj�
        Else
            ' Lastausp�iv� on validi -> Task
            taskData.RecordType = "Task"
            taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppin�
    
            ' Aseta Tila Task-tyypille (t�m� lohko siirret��n t�h�n)
            If Me.fraTila.optTilaHyvaksytty.value = True Then
                taskData.Tila = "HYV�KSYTTY"
            ElseIf Me.fraTila.optTilaTarjous.value = True Then
                taskData.Tila = "TARJOUS"
            Else
                taskData.Tila = "TARJOUS" ' Oletus Task-tyypille
            End If
        End If
    End If

    taskData.asiakas = Me.txtAsiakas.Text
    taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppin�

    ' Muut p�iv�m��r�t (Tallenna Date-tyyppin� tai Null/Empty)
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

    ' Muut tekstikent�t
    taskData.sahkoposti = LCase(Me.txtSahkoposti.Text)
    taskData.lastausMaa = UCase(Me.txtLastausmaa.Text)
    taskData.purkuMaa = UCase(Me.txtPurkumaa.Text)
    taskData.M3m = Me.txtM3m.Text
    taskData.Huomioitavaa = Me.txtHuomioitavaa.Text
    taskData.puhelin = Me.txtPuhelin.Text
    taskData.lastausOsoite = Me.txtLastausosoite.Text
    taskData.purkuOsoite = Me.txtPurkuosoite.Text
    taskData.Vakuutus = Me.txtVakuutus.Text
    taskData.Arvo = Me.txtArvo.Text ' K�sittele Varianttina
    taskData.hinta = Me.txtHinta.Text ' K�sittele Varianttina
    taskData.M3t = Me.txtM3t.Text
    taskData.valimatka = Me.txtValimatka.Text

    ' OptionButtonit
    If Me.fraTila.optTilaHyvaksytty.value = True Then
        taskData.Tila = "HYV�KSYTTY"
    ElseIf Me.fraTila.optTilaTarjous.value = True Then
        taskData.Tila = "TARJOUS"
    Else
        taskData.Tila = "TARJOUS" ' Oletus
    End If
    ' ... muut optionbutton ryhm�t ...
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

    taskData.AttentionSortDate = Null ' Varmista tyhj�ksi

    ' --- 4. Lis�� UUSI teht�v� TaskManagerin kokoelmaan ---
    tm.AddTask taskData ' Antaa ID:n

    ' --- 5. Tallenna KOKO kokoelma takaisin v�lilehdelle ---
    Application.StatusBar = "Tallennetaan uutta teht�v��..."
    tm.SaveToSheet STORAGE_SHEET ' Olettaa, ett� SaveToSheet osaa kirjoittaa Date/Null

    ' --- 6. P�ivit� n�ytt� ---
    Application.StatusBar = "P�ivitet��n n�ytt��..."
    dm.UpdateDisplay tm.tasks, DISPLAY_SHEET

    ' --- 7. Tarjous ---
    Call mdlTarjousUtils.LuoTarjousLomakkeelta(Me)

    Application.StatusBar = False
    'MsgBox "Uusi teht�v� (ID: " & taskData.ID & ") tallennettu onnistuneesti!", vbInformation

    ' --- 8. Sulje lomake ---
    Unload Me
    Exit Sub

SaveErrorHandler:
    Application.StatusBar = False
    Set tm = Nothing ' Vapauta viittaukset
    Set dm = Nothing
    Set taskData = Nothing
    MsgBox "Virhe tallennettaessa uutta teht�v��:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
End Sub


' --- Korvaa t�m� kokonaan frmTehtavat.txt -tiedostoon ---
Private Sub cmdEdit_Click() ' MUOKKAUKSEN TALLENNUS
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim taskData As clsTaskItem ' Haetaan olemassa oleva p�ivityst� varten
    Const STORAGE_SHEET As String = mdlMain.STORAGE_SHEET_NAME
    Const DISPLAY_SHEET As String = mdlMain.DISPLAY_SHEET_NAME
    Dim tempDate As Variant
    Dim lastausPaivaDate As Variant ' Muuttuja lastausp�iv�lle
    Dim lastausPaivaText As String ' Lis�tty

    On Error GoTo EditErrorHandler

    ' --- VALIDointi: Asiakas (pakollinen) ---
    If Trim$(Me.txtAsiakas.Text) = "" Then
        MsgBox "Asiakkaan nimi on pakollinen tieto.", vbExclamation, "Pakollinen tieto"
        Me.txtAsiakas.SetFocus
        Exit Sub ' Poistutaan, k�ytt�j�n on sy�tett�v� nimi ja painettava Tallenna uudelleen
    End If

    ' --- VALIDointi: Lastausp�iv� (jos annettu, oltava validi) ---
    lastausPaivaText = Trim$(Me.txtLastauspaiva.Text)
    If lastausPaivaText <> "" Then ' Tarkistetaan vain, jos kent�ss� on jotain
        lastausPaivaDate = mdlDateUtils.ConvertToDate(lastausPaivaText)
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Antamasi lastausp�iv� '" & lastausPaivaText & "' ei ole kelvollinen." & vbCrLf & _
                   "Sy�t� p�iv�m��r� muodossa pp.kk.vvvv tai j�t� kentt� tyhj�ksi, jos kyseess� on kontaktimerkint�.", vbExclamation, "Virheellinen Lastausp�iv�"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistutaan, k�ytt�j�n on korjattava tai tyhjennett�v� ja painettava Tallenna uudelleen
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
        MsgBox "Muokattavan teht�v�n ID:t� ei l�ytynyt!", vbCritical
        Exit Sub
    End If
    Set taskData = tm.GetTaskByID(Me.TaskIDToEdit) ' Hae muokattava olio
    If taskData Is Nothing Then
        MsgBox "Muokattavaa teht�v�� ID:ll� " & Me.TaskIDToEdit & " ei l�ytynyt muistista!", vbCritical
        Exit Sub
    End If

     ' --- 4. P�ivit� tiedot lomakkeen kontrolleista taskData-olioon ---

    ' Pakolliset kent�t (jo tarkistettu)
    taskData.asiakas = Me.txtAsiakas.Text
    taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppin�
    
    ' Lastausp�iv�n ja RecordType/Tila -logiikka
   'Huom: taskData.kontaktiPaiva EI P�IVITET� T�SS�
    If Trim$(Me.txtLastauspaiva.Text) = "" Then
        ' Lastausp�iv� on tyhj� -> Kontakti
        taskData.RecordType = "Kontakti"
        taskData.Tila = "KONTAKTI"
        taskData.lastausPaiva = Null
    Else
        ' Lastausp�iv�-kent�ss� on sy�tett�, validoi se (lastausPaivaDate on jo alustettu)
        If IsNull(lastausPaivaDate) Or Not IsDate(lastausPaivaDate) Then
            MsgBox "Sy�t� lastausp�iv� oikeassa muodossa (pp.kk.vvvv) tai j�t� kentt� tyhj�ksi kontaktimerkint�� varten.", vbExclamation, "Virheellinen Lastausp�iv�"
            Me.txtLastauspaiva.SetFocus
            Exit Sub ' Poistu, kunnes lastausp�iv� on validi tai tyhj�
        Else
            ' Lastausp�iv� on validi -> Task
            taskData.RecordType = "Task"
            taskData.lastausPaiva = lastausPaivaDate ' Tallenna Date-tyyppin�
            ' Aseta Tila Task-tyypille
            If Me.fraTila.optTilaHyvaksytty.value = True Then
                taskData.Tila = "HYV�KSYTTY"
            ElseIf Me.fraTila.optTilaTarjous.value = True Then
                taskData.Tila = "TARJOUS"
            Else
                taskData.Tila = "TARJOUS" ' Oletus Task-tyypille
            End If
        End If
    End If

    ' Muut p�iv�m��r�t (Tallenna Date-tyyppin� tai Null/Empty)
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

    ' Muut tekstikent�t
    taskData.sahkoposti = Me.txtSahkoposti.Text
    taskData.lastausMaa = UCase(Me.txtLastausmaa.Text)
    taskData.purkuMaa = UCase(Me.txtPurkumaa.Text)
    taskData.M3m = Me.txtM3m.Text
    taskData.Huomioitavaa = Me.txtHuomioitavaa.Text
    taskData.puhelin = Me.txtPuhelin.Text
    taskData.lastausOsoite = Me.txtLastausosoite.Text
    taskData.purkuOsoite = Me.txtPurkuosoite.Text
    taskData.Vakuutus = Me.txtVakuutus.Text
    taskData.Arvo = Me.txtArvo.Text ' K�sittele Varianttina
    taskData.hinta = Me.txtHinta.Text ' K�sittele Varianttina
    taskData.M3t = Me.txtM3t.Text
    taskData.valimatka = Me.txtValimatka.Text

    ' OptionButtonit
    If Me.fraTila.optTilaHyvaksytty.value = True Then
        taskData.Tila = "HYV�KSYTTY"
    ElseIf Me.fraTila.optTilaTarjous.value = True Then
        taskData.Tila = "TARJOUS"
    Else
        taskData.Tila = "TARJOUS" ' Oletus
    End If
    ' ... muut optionbutton ryhm�t ...
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
    taskData.AttentionSortDate = Null ' Varmista tyhj�ksi

    ' --- 5. P�ivit� teht�v� TaskManagerin kokoelmaan ---
    tm.UpdateTask taskData ' Kutsu P�IVITYSmetodia

    ' --- 6. Tallenna KOKO kokoelma takaisin v�lilehdelle ---
    Application.StatusBar = "Tallennetaan muutoksia..."
    tm.SaveToSheet STORAGE_SHEET ' Olettaa, ett� SaveToSheet osaa kirjoittaa Date/Null

    ' --- 7. P�ivit� n�ytt� ---
    Application.StatusBar = "P�ivitet��n n�ytt��..."
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
        MsgBox "Poistettavan teht�v�n ID:t� ei voitu m��ritt��.", vbExclamation
        Exit Sub
    End If

    response = MsgBox("Haluatko varmasti poistaa t�m�n teht�v�n (ID: " & taskIDToDelete & ")?" & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista poisto")

    If response = vbNo Then Exit Sub

    On Error GoTo DeleteErrorHandler

    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Poisto ep�onnistui.", vbCritical, "Virhe"
        GoTo CleanExit_Delete
    End If

    tm.DeleteTask taskIDToDelete

    Application.StatusBar = "Poistetaan ja tallennetaan..."
    tm.SaveToSheet STORAGE_SHEET

    Application.StatusBar = "P�ivitet��n n�ytt��..."
    dm.UpdateDisplay tm.tasks, DISPLAY_SHEET
    
    Application.StatusBar = False
    MsgBox "Teht�v� (ID: " & taskIDToDelete & ") poistettu onnistuneesti.", vbInformation

    Unload Me
    
CleanExit_Delete:
    Application.StatusBar = False
    Exit Sub

DeleteErrorHandler:
    MsgBox "Virhe poistettaessa teht�v�� (ID: " & taskIDToDelete & "):" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Poistovirhe"
    Resume CleanExit_Delete
End Sub
    
' --- Apurutiini ComboBoxin t�ytt�miseen ---
Private Sub PopulateComboBox(cmb As MSForms.ComboBox, sheetName As String, colLetter As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    'On Error Resume Next ' K�sittele virheet (esim. v�lilehte� ei l�ydy)
    Set ws = Nothing
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then
        Debug.Print "Varoitus: V�lilehte� '" & sheetName & "' ei l�ytynyt ComboBoxin t�ytt�� varten."
        Exit Sub
    End If
    ' Hae data-alue m��ritellyst� sarakkeesta (alkaen rivilt� 1)
    Set rng = ws.Range(colLetter & "1", ws.Cells(ws.rows.Count, colLetter).End(xlUp))
    If rng Is Nothing Or rng.rows.Count = 0 Then Exit Sub

    cmb.Clear ' Tyhjenn� vanhat valinnat
    For Each cell In rng.Cells ' K�y l�pi vain solut alueelta
        If Trim$(cell.value) <> "" Then ' Lis�� vain ei-tyhj�t arvot
            cmb.AddItem cell.value
        End If
    Next cell
    On Error GoTo 0 ' Palauta normaali virheiden k�sittely
End Sub

' --- Apurutiini ListBoxin t�ytt�miseen ---
Private Sub PopulateListBox(lst As MSForms.ListBox, sheetName As String, colLetter As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    On Error GoTo PopulateError ' Lis�t��n virheenk�sittelij�

    Set ws = Nothing
    ' Yrit� hakea v�lilehti. Jos ei l�ydy, PopulateError hoitaa.
    Set ws = ThisWorkbook.Worksheets(sheetName)

    ' Etsi viimeinen rivi
    lastRow = ws.Cells(ws.rows.Count, colLetter).End(xlUp).row
    
    ' Jos dataa ei ole (vain otsikko tai tyhj�), tai jos lastRow on virheellisesti 0 tai 1
    If lastRow < 2 Then ' Oletetaan, ett� data alkaa aina v�hint��n rivilt� 2
        ' Debug.Print "PopulateListBox: Ei ladattavaa dataa v�lilehdelt� '" & sheetName & "', sarakkeesta " & colLetter
        GoTo CleanExit ' Poistu siististi, jos ei dataa
    End If

    Set rng = ws.Range(colLetter & "2:" & colLetter & lastRow)
    ' If rng Is Nothing Then GoTo CleanExit ' T�m� tarkistus on ep�todenn�k�inen, jos lastRow >=2

    lst.Clear
    For Each cell In rng.Cells
         ' K�yt� CStr(cell.value) varmistaaksesi, ett� arvo k�sitell��n merkkijonona
         ' ja v�lt� mahdollinen virhe, jos cell.value on esim. virhearvo (#N/A)
         Dim cellValueStr As String
         On Error Resume Next ' Ohita virhe, jos solun arvoa ei voi muuntaa stringiksi
         cellValueStr = CStr(cell.value)
         If Err.Number <> 0 Then
             cellValueStr = "" ' Aseta tyhj�ksi, jos muunnos ep�onnistui
             Err.Clear
         End If
         On Error GoTo PopulateError ' Palauta p��virheenk�sittelij�

         If Trim$(cellValueStr) <> "" Then
            lst.AddItem cellValueStr
         End If
    Next cell

CleanExit:
    On Error GoTo 0 ' Palauta normaali virheenk�sittely
    Set ws = Nothing
    Set rng = Nothing
    Set cell = Nothing
    Exit Sub

PopulateError:
    Debug.Print "Virhe PopulateListBox: V�lilehti='" & sheetName & "', Sarake='" & colLetter & "', ListBox='" & lst.Name & "'. Virhe: " & Err.Description
    ' T�ss� voisi olla my�s MsgBox k�ytt�j�lle, jos listan t�ytt� on kriittist� lomakkeen toiminnalle
    ' Esim. MsgBox "Listan '" & lst.Name & "' t�ytt� v�lilehdelt� '" & sheetName & "' ep�onnistui." & vbCrLf & "Virhe: " & Err.Description, vbExclamation
    Resume CleanExit ' Yrit� siivota ja poistua
End Sub

' --- Apurutiini olemassa olevan teht�v�n tietojen lataamiseksi lomakkeelle ---
Private Function LoadTaskDataIntoForm(taskID As Long) As Boolean
    Dim tm As clsTaskManager
    Dim taskToEdit As clsTaskItem
    LoadTaskDataIntoForm = False
    On Error GoTo LoadDataError

    Set tm = mdlMain.GetTaskManagerInstance()
    'Debug.Print Now & " LoadTaskDataIntoForm: TaskManagerin kokoelman koko HETI haun j�lkeen: " & IIf(tm Is Nothing, "tm on Nothing", tm.tasks.Count)

    If tm Is Nothing Then
        'Debug.Print Now & " KRIITTINEN VIRHE: TaskManager-instanssi on Nothing LoadTaskDataIntoFormissa!"
        MsgBox "Kriittinen virhe: Teht�v�nhallintaa ei voitu alustaa!", vbCritical
        Exit Function
    End If
    'Debug.Print Now & " TaskManagerin kokoelman koko ennen hakua: " & tm.tasks.Count

    Set taskToEdit = tm.GetTaskByID(taskID)

    If taskToEdit Is Nothing Then
        'Debug.Print Now & " LoadTaskDataIntoForm: GetTaskByID palautti Nothing ID:lle " & taskID
        MsgBox "Muokattavaa teht�v�� ID:ll� " & taskID & " ei l�ytynyt muistista!", vbExclamation, "Virhe"
        Exit Function
    'Else
        'Debug.Print Now & " LoadTaskDataIntoForm: Teht�v� l�ytyi ID:ll� " & taskID
    End If

    ' --- T�yt� lomakkeen kent�t haetun olion tiedoilla ---
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
    Me.txtArvo.Text = mdlStringUtils.DefaultIfNull(taskToEdit.Arvo, "") ' Variant, DefaultIfNull k�sittelee
    Me.txtHinta.Text = mdlStringUtils.DefaultIfNull(taskToEdit.hinta, "") ' Variant, DefaultIfNull k�sittelee
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

    ' CheckBoxit (oletetaan, ett� clsTaskItem:ssa n�m� ovat Boolean)
    Me.chkLastauspaivaVarmistunut.value = taskToEdit.LastauspaivaVarmistunut
    Me.chkPurkupaivaVarmistunut.value = taskToEdit.PurkupaivaVarmistunut
    ' Seuraavat olivat clsTaskItem:ssa String, mutta frmTehtavat.Save/Edit asettaa ne Boolean-arvoina CheckBoxeista.
    ' Oletetaan, ett� clsTaskItem.Rahtikirja, .Laskutus, .Muuttomaailma ovat String, jotka voivat olla "KYLL�" tai "EI"
    ' tai True/False Boolean-arvoina. Jos ne ovat String, tarvitaan muunnos.
    ' Koska Save/Edit asettaa ne CheckBox.Value (Boolean), oletetaan ett� ne ovatkin Boolean clsTaskItem:ssa.
    ' Jos ne ovat String, tarvitaan:
    ' Me.chkRahtikirja.value = (UCase(taskToEdit.Rahtikirja) = "KYLL�" Or UCase(taskToEdit.Rahtikirja) = "TRUE" Or taskToEdit.Rahtikirja = "-1")
    ' Mutta oletetaan nyt ett� ne ovat Boolean:
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
        Case "HYV�KSYTTY": Me.optTilaHyvaksytty.value = True
        Case Else: Me.optTilaTarjous.value = True ' Oletus
    End Select

    LoadTaskDataIntoForm = True
    Exit Function

LoadDataError:
    MsgBox "Odottamaton virhe ladattaessa teht�v�n tietoja (ID: " & taskID & "):" & vbCrLf & Err.Description, vbCritical, "Latausvirhe"
End Function

' Sulje -nappi
Private Sub cmdCancel_Click()
    Unload Me
End Sub

