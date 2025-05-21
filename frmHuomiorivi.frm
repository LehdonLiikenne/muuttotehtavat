VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHuomiorivi 
   Caption         =   "frmHuomiorivi"
   ClientHeight    =   7464
   ClientLeft      =   300
   ClientTop       =   1188
   ClientWidth     =   5376
   OleObjectBlob   =   "frmHuomiorivi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHuomiorivi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Julkinen muuttuja ID:lle, jota käytetään määrittämään,
' ollaanko lisäämässä uutta (ID=0) vai muokkaamassa olemassa olevaa.
Public TaskIDToEdit As Long


' --- Suoritetaan, kun lomake ladataan muistiin (ennen näyttämistä) ---
Private Sub UserForm_Initialize()
    On Error GoTo Initialize_Error

    ' 1. Tyhjennetään kontrollit käyttäen mdlClearForm-moduulia
    '    Olettaa, että mdlClearForm.ClearForm toimii oikein tälle lomakkeelle.
    mdlClearForm.ClearForm Me

    ' 2. TÄYTETÄÄN LISTBOXIT APUVÄLILEHDILTÄ
    '    Oletetaan välilehtien nimet: "Kuljettajat", "Autot", "Kontit"
    '    ja datan olevan sarakkeessa A.
    PopulateListBox Me.lstHuomioriviKuljettajat, "Kuljettajat", "B"
    PopulateListBox Me.lstHuomioriviAutot, "Autot", "B"
    PopulateListBox Me.lstHuomioriviKontit, "Kontit", "B"

CleanExit_Initialize:
    On Error GoTo 0 ' Nollaa virheenkäsittely ennen poistumista
    Exit Sub

Initialize_Error:
    MsgBox "Virhe alustettaessa huomiorivilomaketta:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Lomakkeen Alustusvirhe"
    Resume CleanExit_Initialize
End Sub

' --- Apurutiini ListBoxin täyttämiseen ---
Private Sub PopulateListBox(lst As MSForms.ListBox, sheetName As String, colLetter As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    On Error GoTo PopulateError ' Kattava virheenkäsittelijä

    Set ws = ThisWorkbook.Worksheets(sheetName) ' Anna tämän aiheuttaa virhe, jos sheetName on väärin

    lastRow = ws.Cells(ws.rows.Count, colLetter).End(xlUp).row
    If lastRow < 2 Then GoTo CleanExit_Populate ' Oletetaan data alkavan riviltä 2

    Set rng = ws.Range(colLetter & "2:" & colLetter & lastRow)

    lst.Clear
    Dim cellValueStr As String
    For Each cell In rng.Cells
        On Error Resume Next ' Suojaa CStr(cell.value)
        cellValueStr = CStr(cell.value)
        If Err.Number <> 0 Then
            cellValueStr = ""
            Err.Clear
        End If
        On Error GoTo PopulateError ' Palauta pääkäsittelijä

        If Trim$(cellValueStr) <> "" Then
           lst.AddItem cellValueStr
        End If
    Next cell

CleanExit_Populate:
    On Error GoTo 0
    Set ws = Nothing
    Set rng = Nothing
    Set cell = Nothing
    Exit Sub

PopulateError:
    If ws Is Nothing Then ' Tarkempi virheilmoitus, jos välilehteä ei löytynyt
        Debug.Print "PopulateListBox: Välilehteä '" & sheetName & "' ei löytynyt ListBoxin '" & lst.Name & "' täyttöä varten."
        MsgBox "Virhe: Tarvittavaa välilehteä '" & sheetName & "' ei löytynyt listan '" & lst.Name & "' täyttämiseksi.", vbExclamation, "Listan Täyttövirhe"
    Else ' Muu virhe
        MsgBox "Virhe täytettäessä listaa '" & lst.Name & "' välilehdeltä '" & sheetName & "':" & vbCrLf & Err.Description, vbExclamation, "Listan Täyttövirhe"
    End If
    Resume CleanExit_Populate
End Sub

' --- Ajetaan, kun lomake aktivoituu (juuri ennen näyttämistä) ---
Private Sub UserForm_Activate()
    Dim loadSuccess As Boolean

    On Error GoTo Activate_Error

    If Me.TaskIDToEdit > 0 Then
        ' --- MUOKKAUSTILA ---
        Me.Caption = "Muokkaa Huomioriviä (Ladataan...)"
        loadSuccess = LoadAttentionDataIntoForm(Me.TaskIDToEdit)

        If loadSuccess Then
            ' Lataus onnistui
            Me.Caption = "Muokkaa Huomioriviä (ID: " & Me.TaskIDToEdit & ")"
            ' Aseta painikkeet muokkaustilaan
            Me.cmdSave.Visible = False ' Piilota Lisää-painike
            Me.cmdSave.Enabled = False
            Me.cmdEdit.Visible = True  ' Näytä Tallenna Muutokset -painike
            Me.cmdEdit.Enabled = True
            Me.cmdDelete.Visible = True ' Näytä Poista-painike
            Me.cmdDelete.Enabled = True
            Me.cmdEdit.SetFocus ' Kohdistus Tallenna-painikkeeseen (tai txtHuomioriviHuomio)
        Else
            ' Lataus epäonnistui
            Me.Hide
            Unload Me
            Exit Sub
        End If

    Else
        ' --- LISÄYSTILA ---
        Me.Caption = "Lisää Uusi Huomiorivi"
        ' Initialize on jo tyhjentänyt kentät
        ' Aseta painikkeet lisäystilaan
        Me.cmdSave.Visible = True   ' Näytä Lisää-painike
        Me.cmdSave.Enabled = True
        Me.cmdEdit.Visible = False ' Piilota Tallenna Muutokset -painike
        Me.cmdEdit.Enabled = False
        Me.cmdDelete.Visible = False ' Piilota Poista-painike
        Me.cmdDelete.Enabled = False
        Me.txtHuomioriviHuomio.SetFocus ' Kohdistus ensimmäiseen kenttään
    End If

CleanExit_Activate:
    Exit Sub

Activate_Error:
     MsgBox "Virhe aktivoitaessa huomiorivilomaketta:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Lomakkeen Aktivointivirhe"
     On Error Resume Next
     Me.Hide
     Unload Me
     On Error GoTo 0
     Resume CleanExit_Activate
End Sub

Private Function LoadAttentionDataIntoForm(taskID As Long) As Boolean
    Dim tm As clsTaskManager
    Dim taskToEdit As clsTaskItem
    LoadAttentionDataIntoForm = False ' Oletus: epäonnistui

    On Error GoTo LoadDataError

    ' Hae TaskManager-instanssi
    Set tm = mdlMain.GetTaskManagerInstance()
    If tm Is Nothing Then
        MsgBox "Kriittinen virhe: Tehtävänhallintaa ei voitu alustaa!", vbCritical, "Virhe"
        Exit Function ' Palauta False
    End If

    ' Hae TaskItem ID:n perusteella
    Set taskToEdit = tm.GetTaskByID(taskID)

    ' Tarkista, löytyikö olio
    If taskToEdit Is Nothing Then
        MsgBox "Muokattavaa tietuetta ID:llä " & taskID & " ei löytynyt muistista!", vbExclamation, "Virhe"
        Exit Function ' Palauta False
    End If

    ' TÄRKEÄ TARKISTUS: Varmista, että kyseessä on Huomiorivi
    If taskToEdit.RecordType <> "Attention" Then
         MsgBox "Tietue ID:llä " & taskID & " ei ole Huomiorivi." & vbCrLf & _
                "Avaa oikea muokkauslomake.", vbExclamation, "Väärä Tyyppi"
         Exit Function ' Palauta False
    End If

    ' --- Täytä lomakkeen kentät haetun olion tiedoilla ---
    Me.txtHuomioriviHuomio.Text = mdlStringUtils.DefaultIfNull(taskToEdit.Huomioitavaa, "")
    Me.txtHuomioriviPaiva.Text = mdlDateUtils.FormatDateToString(taskToEdit.AttentionSortDate, "")

    ' Aseta ListBoxien valinnat (käytä funktiota mdlStringUtils-moduulista)
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstHuomioriviKuljettajat, taskToEdit.Kuljettajat, ";")
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstHuomioriviAutot, taskToEdit.Autot, ";")
    Call mdlStringUtils.SetListBoxMultiSelection(Me.lstHuomioriviKontit, taskToEdit.Kontit, ";")

    ' Jos kaikki meni hyvin
    LoadAttentionDataIntoForm = True
    Exit Function ' Poistu onnistuneesti

LoadDataError:
    MsgBox "Odottamaton virhe ladattaessa huomiorivin tietoja (ID: " & taskID & "):" & vbCrLf & Err.Description, vbCritical, "Latausvirhe"
    ' Palauttaa edelleen False (oletusarvo)
End Function

' --- Tallenna-painikkeen toiminto (Lisää uusi tai Tallenna muutokset) ---
Private Sub cmdSave_Click()
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim attnData As clsTaskItem
    Dim isNew As Boolean
    Dim tempDate As Variant

    ' Hae Manager-oliot (käytä mdlMain:n funktioita)
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    ' Tarkista, että managerit saatiin alustettua
    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Tallennus epäonnistui.", vbCritical, "Virhe"
        Exit Sub
    End If

    On Error GoTo SaveErrorHandler

    isNew = (Me.TaskIDToEdit <= 0) ' Määritä, onko kyseessä lisäys vai muokkaus

    ' --- Validoi syötteet ennen jatkamista ---
    If Trim$(Me.txtHuomioriviHuomio.Text) = "" Then
        MsgBox "Huomio-teksti ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtHuomioriviHuomio.SetFocus
        GoTo CleanExit_Save ' Poistu siististi ilman tallennusta
    End If

    tempDate = mdlDateUtils.ConvertToDate(Me.txtHuomioriviPaiva.Text)
    If IsNull(tempDate) Or Not IsDate(tempDate) Then
        MsgBox "Antamasi päivämäärä '" & Me.txtHuomioriviPaiva.Text & "' ei ole kelvollinen.", vbExclamation, "Virheellinen Päivämäärä"
        Me.txtHuomioriviPaiva.SetFocus
        GoTo CleanExit_Save ' Poistu siististi ilman tallennusta
    End If
    ' --- Validointi OK ---


    If isNew Then
        ' --- Lisää uusi huomiorivi ---
        Set attnData = New clsTaskItem
        ' attnData.InitDefaults ' Voi kutsua, jos InitDefaults tekee jotain hyödyllistä huomioriveille
        attnData.RecordType = "Attention" ' Aseta tyyppi
        ' ID annetaan AddTask-metodissa
    Else
        ' --- Muokkaa olemassa olevaa huomioriviä ---
        Set attnData = tm.GetTaskByID(Me.TaskIDToEdit)
        If attnData Is Nothing Then
            MsgBox "Muokattavaa tietuetta (ID: " & Me.TaskIDToEdit & ") ei löytynyt. Tallennus peruttu.", vbCritical, "Virhe"
            GoTo CleanExit_Save
        End If
        ' Varmistus tyypille (vaikka Activate teki sen jo)
        If attnData.RecordType <> "Attention" Then
             MsgBox "Tietue (ID: " & Me.TaskIDToEdit & ") ei ole Huomiorivi. Tallennus peruttu.", vbCritical, "Väärä Tyyppi"
             GoTo CleanExit_Save
        End If
        ' ID säilyy samana
    End If

    ' --- Siirrä tiedot lomakkeelta attnData-olioon ---
    attnData.Huomioitavaa = Me.txtHuomioriviHuomio.Text
    attnData.AttentionSortDate = tempDate ' Käytä validoitua päivämäärää
    attnData.Kuljettajat = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKuljettajat, ";")
    attnData.Autot = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviAutot, ";")
    attnData.Kontit = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKontit, ";")
    attnData.RecordType = "Attention" ' Varmistetaan vielä

    ' --- Nollaa/Tyhjennä Task-tyyppiin liittyvät kentät ---
    ' Tämä varmistaa, ettei turhaa dataa tallennu Attention-riveille
    attnData.asiakas = vbNullString
    attnData.tarjousTehty = Null ' Tai Empty
    attnData.lastausPaiva = Null
    attnData.lastausMaa = vbNullString
    attnData.purkuMaa = vbNullString
    attnData.purkuPaiva = Null
    attnData.M3m = vbNullString
    attnData.palvelu = vbNullString
    attnData.puhelin = vbNullString
    attnData.lastausOsoite = vbNullString
    attnData.purkuOsoite = vbNullString
    attnData.Apulaiset = vbNullString ' Huom: Eri kuin ApulaisetTilattu
    attnData.Rahtikirja = False ' Tai vbNullString riippuen TaskItemista
    attnData.ApulaisetTilattu = vbNullString
    attnData.Pysakointilupa = vbNullString
    attnData.hissi = vbNullString
    attnData.Laivalippu = vbNullString
    attnData.Laskutus = False ' Tai vbNullString
    attnData.Vakuutus = vbNullString
    attnData.Arvo = Null
    attnData.hinta = Null
    attnData.Muuttomaailma = False ' Tai vbNullString
    attnData.M3t = vbNullString
    attnData.LastauspaivaVarmistunut = False
    attnData.PurkupaivaVarmistunut = False
    attnData.TarjousHyvaksytty = Null
    attnData.TarjousHylatty = Null
    attnData.Tila = "HUOMIO" ' Voidaan asettaa oletusarvo, jos halutaan
    attnData.LastausLoppuu = Null
    attnData.PurkuLoppuu = Null
    ' --- Kenttien nollaus valmis ---


    ' --- Suorita tallennustoiminnot ---
    Application.StatusBar = "Tallennetaan huomioriviä..."

    If isNew Then
        tm.AddTask attnData ' Lisää uusi (antaa ID:n)
    Else
        tm.UpdateTask attnData ' Päivitä olemassa oleva
    End If

    ' Tallenna KOKO kokoelma (sisältäen muutoksen/lisäyksen) takaisin välilehdelle
    tm.SaveToSheet mdlMain.STORAGE_SHEET_NAME

    ' Päivitä näyttö
    dm.UpdateDisplay tm.tasks, mdlMain.DISPLAY_SHEET_NAME

    Application.StatusBar = False
    If isNew Then
        MsgBox "Uusi huomiorivi (ID: " & attnData.ID & ") tallennettu onnistuneesti!", vbInformation, "Lisäys Onnistui"
    Else
        MsgBox "Muutokset huomioriviin (ID: " & attnData.ID & ") tallennettu onnistuneesti!", vbInformation, "Muokkaus Onnistui"
    End If

    ' Sulje lomake onnistuneen tallennuksen jälkeen
    Unload Me
    GoTo CleanExit_Save ' Hyppää siivoukseen

SaveErrorHandler:
    Application.StatusBar = False
    MsgBox "Virhe tallennettaessa huomioriviä:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    ' ÄLÄ sulje lomaketta virhetilanteessa, jotta käyttäjä voi korjata

CleanExit_Save:
    ' Vapauta oliomuuttujat
    Set attnData = Nothing
    Set tm = Nothing
    Set dm = Nothing
End Sub

' --- Tallenna Muutokset -painikkeen toiminto ---
Private Sub cmdEdit_Click()
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim attnData As clsTaskItem
    Dim tempDate As Variant

    ' Varmista, että ollaan muokkaustilassa
    If Me.TaskIDToEdit <= 0 Then
        MsgBox "Virhe: Muokkaustoimintoa kutsuttiin ilman validia ID:tä.", vbCritical
        Exit Sub
    End If

    ' Hae Manager-oliot
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Tallennus epäonnistui.", vbCritical, "Virhe"
        Exit Sub
    End If

    On Error GoTo EditErrorHandler

    ' --- Validoi syötteet ---
    If Trim$(Me.txtHuomioriviHuomio.Text) = "" Then
        MsgBox "Huomio-teksti ei voi olla tyhjä.", vbExclamation, "Puuttuva Tieto"
        Me.txtHuomioriviHuomio.SetFocus
        GoTo CleanExit_Edit
    End If
    tempDate = mdlDateUtils.ConvertToDate(Me.txtHuomioriviPaiva.Text)
    If IsNull(tempDate) Or Not IsDate(tempDate) Then
        MsgBox "Antamasi päivämäärä '" & Me.txtHuomioriviPaiva.Text & "' ei ole kelvollinen.", vbExclamation, "Virheellinen Päivämäärä"
        Me.txtHuomioriviPaiva.SetFocus
        GoTo CleanExit_Edit
    End If
    ' --- Validointi OK ---

    ' Hae muokattava olio TaskManagerista
    Set attnData = tm.GetTaskByID(Me.TaskIDToEdit)

    If attnData Is Nothing Then
        MsgBox "Muokattavaa tietuetta (ID: " & Me.TaskIDToEdit & ") ei löytynyt. Tallennus peruttu.", vbCritical, "Virhe"
        GoTo CleanExit_Edit
    End If
    If attnData.RecordType <> "Attention" Then
         MsgBox "Tietue (ID: " & Me.TaskIDToEdit & ") ei ole Huomiorivi. Tallennus peruttu.", vbCritical, "Väärä Tyyppi"
         GoTo CleanExit_Edit
    End If

    ' --- Päivitä tiedot lomakkeelta attnData-olioon ---
    attnData.Huomioitavaa = Me.txtHuomioriviHuomio.Text
    attnData.AttentionSortDate = tempDate
    attnData.Kuljettajat = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKuljettajat, ";")
    attnData.Autot = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviAutot, ";")
    attnData.Kontit = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKontit, ";")
    ' RecordType ja ID säilyvät ennallaan
    ' Varmistetaan myös tässä, että Task-kentät ovat tyhjiä (jos joku aiempi vaihe epäonnistui)
    ' (Voit kopioida nollauskoodin cmdAdd_Click:stä tai luoda erillisen ResetTaskFields-apurutiinin)
    attnData.Tila = "HUOMIO" ' Varmistetaan Tila

    Application.StatusBar = "Tallennetaan muutoksia huomioriviin..."

    ' Päivitä olio TaskManagerissa
    tm.UpdateTask attnData

    ' Tallenna kokoelma levylle
    tm.SaveToSheet mdlMain.STORAGE_SHEET_NAME

    ' Päivitä näyttö
    dm.UpdateDisplay tm.tasks, mdlMain.DISPLAY_SHEET_NAME

    Application.StatusBar = False
    MsgBox "Muutokset huomioriviin (ID: " & attnData.ID & ") tallennettu onnistuneesti!", vbInformation, "Muokkaus Onnistui"

    Unload Me ' Sulje lomake
    GoTo CleanExit_Edit

EditErrorHandler:
    Application.StatusBar = False
    MsgBox "Virhe tallennettaessa muutoksia huomioriviin:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Muokkausvirhe"
    ' Älä sulje lomaketta

CleanExit_Edit:
    Set attnData = Nothing
    Set tm = Nothing
    Set dm = Nothing
End Sub

' --- Poista-painikkeen toiminto ---
Private Sub cmdDelete_Click()
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim taskIDToDelete As Long
    Dim response As VbMsgBoxResult
    Dim itemToDelete As clsTaskItem ' Lisätty tarkistusta varten

    taskIDToDelete = Me.TaskIDToEdit ' Haetaan poistettava ID lomakkeelta

    ' Varmista, että ollaan muokkaustilassa ja ID on validi
    If taskIDToDelete <= 0 Then
        MsgBox "Poistettavan huomiorivin ID:tä ei voitu määrittää. Toiminto peruttu.", vbExclamation, "Virhe"
        Exit Sub
    End If

    ' Kysy varmistus käyttäjältä
    response = MsgBox("Haluatko varmasti poistaa tämän huomiorivin (ID: " & taskIDToDelete & ")?" & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    ' Jos käyttäjä ei halua poistaa, lopeta
    If response = vbNo Then Exit Sub

    ' Jos käyttäjä vahvisti poiston (vbYes)

    ' Hae Manager-oliot
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Poisto epäonnistui.", vbCritical, "Virhe"
        Exit Sub
    End If

    On Error GoTo DeleteErrorHandler

    ' --- Lisäturvatarkistus: Varmista, että ID viittaa Huomioriviin ---
    Set itemToDelete = tm.GetTaskByID(taskIDToDelete)
    If itemToDelete Is Nothing Then
         MsgBox "Poistettavaa tietuetta (ID: " & taskIDToDelete & ") ei löytynyt. Poisto peruttu.", vbExclamation, "Virhe"
         GoTo CleanExit_Delete
    ElseIf itemToDelete.RecordType <> "Attention" Then
         MsgBox "Tietue (ID: " & taskIDToDelete & ") ei ole Huomiorivi. Poisto peruttu.", vbExclamation, "Väärä Tyyppi"
         GoTo CleanExit_Delete
    End If
    ' --- Tarkistus OK ---

    Application.StatusBar = "Poistetaan huomioriviä..."

    ' 1. Poista tehtävä muistista TaskManagerin avulla
    tm.DeleteTask taskIDToDelete

    ' 2. Tallenna muuttunut kokoelma välilehdelle
    tm.SaveToSheet mdlMain.STORAGE_SHEET_NAME

    ' 3. Päivitä näyttö
    dm.UpdateDisplay tm.tasks, mdlMain.DISPLAY_SHEET_NAME

    Application.StatusBar = False
    MsgBox "Huomiorivi (ID: " & taskIDToDelete & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"

    ' 4. Sulje lomake onnistuneen poiston jälkeen
    Unload Me
    GoTo CleanExit_Delete ' Hyppää siivoukseen

DeleteErrorHandler:
    Application.StatusBar = False
    MsgBox "Virhe poistettaessa huomioriviä (ID: " & taskIDToDelete & "):" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Poistovirhe"
    ' ÄLÄ sulje lomaketta virhetilanteessa

CleanExit_Delete:
    Application.StatusBar = False ' Varmistaa, että status bar tyhjenee aina
    Set itemToDelete = Nothing
    Set tm = Nothing
    Set dm = Nothing
End Sub

' --- Peruuta-painikkeen toiminto ---
Private Sub cmdCancel_Click()
    Unload Me
End Sub

