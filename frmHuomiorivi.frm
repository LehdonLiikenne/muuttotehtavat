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

' Julkinen muuttuja ID:lle, jota k�ytet��n m��ritt�m��n,
' ollaanko lis��m�ss� uutta (ID=0) vai muokkaamassa olemassa olevaa.
Public TaskIDToEdit As Long


' --- Suoritetaan, kun lomake ladataan muistiin (ennen n�ytt�mist�) ---
Private Sub UserForm_Initialize()
    On Error GoTo Initialize_Error

    ' 1. Tyhjennet��n kontrollit k�ytt�en mdlClearForm-moduulia
    '    Olettaa, ett� mdlClearForm.ClearForm toimii oikein t�lle lomakkeelle.
    mdlClearForm.ClearForm Me

    ' 2. T�YTET��N LISTBOXIT APUV�LILEHDILT�
    '    Oletetaan v�lilehtien nimet: "Kuljettajat", "Autot", "Kontit"
    '    ja datan olevan sarakkeessa A.
    PopulateListBox Me.lstHuomioriviKuljettajat, "Kuljettajat", "B"
    PopulateListBox Me.lstHuomioriviAutot, "Autot", "B"
    PopulateListBox Me.lstHuomioriviKontit, "Kontit", "B"

CleanExit_Initialize:
    On Error GoTo 0 ' Nollaa virheenk�sittely ennen poistumista
    Exit Sub

Initialize_Error:
    MsgBox "Virhe alustettaessa huomiorivilomaketta:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Lomakkeen Alustusvirhe"
    Resume CleanExit_Initialize
End Sub

' --- Apurutiini ListBoxin t�ytt�miseen ---
Private Sub PopulateListBox(lst As MSForms.ListBox, sheetName As String, colLetter As String)
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long

    On Error GoTo PopulateError ' Kattava virheenk�sittelij�

    Set ws = ThisWorkbook.Worksheets(sheetName) ' Anna t�m�n aiheuttaa virhe, jos sheetName on v��rin

    lastRow = ws.Cells(ws.rows.Count, colLetter).End(xlUp).row
    If lastRow < 2 Then GoTo CleanExit_Populate ' Oletetaan data alkavan rivilt� 2

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
        On Error GoTo PopulateError ' Palauta p��k�sittelij�

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
    If ws Is Nothing Then ' Tarkempi virheilmoitus, jos v�lilehte� ei l�ytynyt
        Debug.Print "PopulateListBox: V�lilehte� '" & sheetName & "' ei l�ytynyt ListBoxin '" & lst.Name & "' t�ytt�� varten."
        MsgBox "Virhe: Tarvittavaa v�lilehte� '" & sheetName & "' ei l�ytynyt listan '" & lst.Name & "' t�ytt�miseksi.", vbExclamation, "Listan T�ytt�virhe"
    Else ' Muu virhe
        MsgBox "Virhe t�ytett�ess� listaa '" & lst.Name & "' v�lilehdelt� '" & sheetName & "':" & vbCrLf & Err.Description, vbExclamation, "Listan T�ytt�virhe"
    End If
    Resume CleanExit_Populate
End Sub

' --- Ajetaan, kun lomake aktivoituu (juuri ennen n�ytt�mist�) ---
Private Sub UserForm_Activate()
    Dim loadSuccess As Boolean

    On Error GoTo Activate_Error

    If Me.TaskIDToEdit > 0 Then
        ' --- MUOKKAUSTILA ---
        Me.Caption = "Muokkaa Huomiorivi� (Ladataan...)"
        loadSuccess = LoadAttentionDataIntoForm(Me.TaskIDToEdit)

        If loadSuccess Then
            ' Lataus onnistui
            Me.Caption = "Muokkaa Huomiorivi� (ID: " & Me.TaskIDToEdit & ")"
            ' Aseta painikkeet muokkaustilaan
            Me.cmdSave.Visible = False ' Piilota Lis��-painike
            Me.cmdSave.Enabled = False
            Me.cmdEdit.Visible = True  ' N�yt� Tallenna Muutokset -painike
            Me.cmdEdit.Enabled = True
            Me.cmdDelete.Visible = True ' N�yt� Poista-painike
            Me.cmdDelete.Enabled = True
            Me.cmdEdit.SetFocus ' Kohdistus Tallenna-painikkeeseen (tai txtHuomioriviHuomio)
        Else
            ' Lataus ep�onnistui
            Me.Hide
            Unload Me
            Exit Sub
        End If

    Else
        ' --- LIS�YSTILA ---
        Me.Caption = "Lis�� Uusi Huomiorivi"
        ' Initialize on jo tyhjent�nyt kent�t
        ' Aseta painikkeet lis�ystilaan
        Me.cmdSave.Visible = True   ' N�yt� Lis��-painike
        Me.cmdSave.Enabled = True
        Me.cmdEdit.Visible = False ' Piilota Tallenna Muutokset -painike
        Me.cmdEdit.Enabled = False
        Me.cmdDelete.Visible = False ' Piilota Poista-painike
        Me.cmdDelete.Enabled = False
        Me.txtHuomioriviHuomio.SetFocus ' Kohdistus ensimm�iseen kentt��n
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
    LoadAttentionDataIntoForm = False ' Oletus: ep�onnistui

    On Error GoTo LoadDataError

    ' Hae TaskManager-instanssi
    Set tm = mdlMain.GetTaskManagerInstance()
    If tm Is Nothing Then
        MsgBox "Kriittinen virhe: Teht�v�nhallintaa ei voitu alustaa!", vbCritical, "Virhe"
        Exit Function ' Palauta False
    End If

    ' Hae TaskItem ID:n perusteella
    Set taskToEdit = tm.GetTaskByID(taskID)

    ' Tarkista, l�ytyik� olio
    If taskToEdit Is Nothing Then
        MsgBox "Muokattavaa tietuetta ID:ll� " & taskID & " ei l�ytynyt muistista!", vbExclamation, "Virhe"
        Exit Function ' Palauta False
    End If

    ' T�RKE� TARKISTUS: Varmista, ett� kyseess� on Huomiorivi
    If taskToEdit.RecordType <> "Attention" Then
         MsgBox "Tietue ID:ll� " & taskID & " ei ole Huomiorivi." & vbCrLf & _
                "Avaa oikea muokkauslomake.", vbExclamation, "V��r� Tyyppi"
         Exit Function ' Palauta False
    End If

    ' --- T�yt� lomakkeen kent�t haetun olion tiedoilla ---
    Me.txtHuomioriviHuomio.Text = mdlStringUtils.DefaultIfNull(taskToEdit.Huomioitavaa, "")
    Me.txtHuomioriviPaiva.Text = mdlDateUtils.FormatDateToString(taskToEdit.AttentionSortDate, "")

    ' Aseta ListBoxien valinnat (k�yt� funktiota mdlStringUtils-moduulista)
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

' --- Tallenna-painikkeen toiminto (Lis�� uusi tai Tallenna muutokset) ---
Private Sub cmdSave_Click()
    Dim tm As clsTaskManager
    Dim dm As clsDisplayManager
    Dim attnData As clsTaskItem
    Dim isNew As Boolean
    Dim tempDate As Variant

    ' Hae Manager-oliot (k�yt� mdlMain:n funktioita)
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    ' Tarkista, ett� managerit saatiin alustettua
    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Tallennus ep�onnistui.", vbCritical, "Virhe"
        Exit Sub
    End If

    On Error GoTo SaveErrorHandler

    isNew = (Me.TaskIDToEdit <= 0) ' M��rit�, onko kyseess� lis�ys vai muokkaus

    ' --- Validoi sy�tteet ennen jatkamista ---
    If Trim$(Me.txtHuomioriviHuomio.Text) = "" Then
        MsgBox "Huomio-teksti ei voi olla tyhj�.", vbExclamation, "Puuttuva Tieto"
        Me.txtHuomioriviHuomio.SetFocus
        GoTo CleanExit_Save ' Poistu siististi ilman tallennusta
    End If

    tempDate = mdlDateUtils.ConvertToDate(Me.txtHuomioriviPaiva.Text)
    If IsNull(tempDate) Or Not IsDate(tempDate) Then
        MsgBox "Antamasi p�iv�m��r� '" & Me.txtHuomioriviPaiva.Text & "' ei ole kelvollinen.", vbExclamation, "Virheellinen P�iv�m��r�"
        Me.txtHuomioriviPaiva.SetFocus
        GoTo CleanExit_Save ' Poistu siististi ilman tallennusta
    End If
    ' --- Validointi OK ---


    If isNew Then
        ' --- Lis�� uusi huomiorivi ---
        Set attnData = New clsTaskItem
        ' attnData.InitDefaults ' Voi kutsua, jos InitDefaults tekee jotain hy�dyllist� huomioriveille
        attnData.RecordType = "Attention" ' Aseta tyyppi
        ' ID annetaan AddTask-metodissa
    Else
        ' --- Muokkaa olemassa olevaa huomiorivi� ---
        Set attnData = tm.GetTaskByID(Me.TaskIDToEdit)
        If attnData Is Nothing Then
            MsgBox "Muokattavaa tietuetta (ID: " & Me.TaskIDToEdit & ") ei l�ytynyt. Tallennus peruttu.", vbCritical, "Virhe"
            GoTo CleanExit_Save
        End If
        ' Varmistus tyypille (vaikka Activate teki sen jo)
        If attnData.RecordType <> "Attention" Then
             MsgBox "Tietue (ID: " & Me.TaskIDToEdit & ") ei ole Huomiorivi. Tallennus peruttu.", vbCritical, "V��r� Tyyppi"
             GoTo CleanExit_Save
        End If
        ' ID s�ilyy samana
    End If

    ' --- Siirr� tiedot lomakkeelta attnData-olioon ---
    attnData.Huomioitavaa = Me.txtHuomioriviHuomio.Text
    attnData.AttentionSortDate = tempDate ' K�yt� validoitua p�iv�m��r��
    attnData.Kuljettajat = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKuljettajat, ";")
    attnData.Autot = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviAutot, ";")
    attnData.Kontit = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKontit, ";")
    attnData.RecordType = "Attention" ' Varmistetaan viel�

    ' --- Nollaa/Tyhjenn� Task-tyyppiin liittyv�t kent�t ---
    ' T�m� varmistaa, ettei turhaa dataa tallennu Attention-riveille
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
    Application.StatusBar = "Tallennetaan huomiorivi�..."

    If isNew Then
        tm.AddTask attnData ' Lis�� uusi (antaa ID:n)
    Else
        tm.UpdateTask attnData ' P�ivit� olemassa oleva
    End If

    ' Tallenna KOKO kokoelma (sis�lt�en muutoksen/lis�yksen) takaisin v�lilehdelle
    tm.SaveToSheet mdlMain.STORAGE_SHEET_NAME

    ' P�ivit� n�ytt�
    dm.UpdateDisplay tm.tasks, mdlMain.DISPLAY_SHEET_NAME

    Application.StatusBar = False
    If isNew Then
        MsgBox "Uusi huomiorivi (ID: " & attnData.ID & ") tallennettu onnistuneesti!", vbInformation, "Lis�ys Onnistui"
    Else
        MsgBox "Muutokset huomioriviin (ID: " & attnData.ID & ") tallennettu onnistuneesti!", vbInformation, "Muokkaus Onnistui"
    End If

    ' Sulje lomake onnistuneen tallennuksen j�lkeen
    Unload Me
    GoTo CleanExit_Save ' Hypp�� siivoukseen

SaveErrorHandler:
    Application.StatusBar = False
    MsgBox "Virhe tallennettaessa huomiorivi�:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Tallennusvirhe"
    ' �L� sulje lomaketta virhetilanteessa, jotta k�ytt�j� voi korjata

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

    ' Varmista, ett� ollaan muokkaustilassa
    If Me.TaskIDToEdit <= 0 Then
        MsgBox "Virhe: Muokkaustoimintoa kutsuttiin ilman validia ID:t�.", vbCritical
        Exit Sub
    End If

    ' Hae Manager-oliot
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Tallennus ep�onnistui.", vbCritical, "Virhe"
        Exit Sub
    End If

    On Error GoTo EditErrorHandler

    ' --- Validoi sy�tteet ---
    If Trim$(Me.txtHuomioriviHuomio.Text) = "" Then
        MsgBox "Huomio-teksti ei voi olla tyhj�.", vbExclamation, "Puuttuva Tieto"
        Me.txtHuomioriviHuomio.SetFocus
        GoTo CleanExit_Edit
    End If
    tempDate = mdlDateUtils.ConvertToDate(Me.txtHuomioriviPaiva.Text)
    If IsNull(tempDate) Or Not IsDate(tempDate) Then
        MsgBox "Antamasi p�iv�m��r� '" & Me.txtHuomioriviPaiva.Text & "' ei ole kelvollinen.", vbExclamation, "Virheellinen P�iv�m��r�"
        Me.txtHuomioriviPaiva.SetFocus
        GoTo CleanExit_Edit
    End If
    ' --- Validointi OK ---

    ' Hae muokattava olio TaskManagerista
    Set attnData = tm.GetTaskByID(Me.TaskIDToEdit)

    If attnData Is Nothing Then
        MsgBox "Muokattavaa tietuetta (ID: " & Me.TaskIDToEdit & ") ei l�ytynyt. Tallennus peruttu.", vbCritical, "Virhe"
        GoTo CleanExit_Edit
    End If
    If attnData.RecordType <> "Attention" Then
         MsgBox "Tietue (ID: " & Me.TaskIDToEdit & ") ei ole Huomiorivi. Tallennus peruttu.", vbCritical, "V��r� Tyyppi"
         GoTo CleanExit_Edit
    End If

    ' --- P�ivit� tiedot lomakkeelta attnData-olioon ---
    attnData.Huomioitavaa = Me.txtHuomioriviHuomio.Text
    attnData.AttentionSortDate = tempDate
    attnData.Kuljettajat = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKuljettajat, ";")
    attnData.Autot = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviAutot, ";")
    attnData.Kontit = mdlStringUtils.GetListBoxMultiSelection(Me.lstHuomioriviKontit, ";")
    ' RecordType ja ID s�ilyv�t ennallaan
    ' Varmistetaan my�s t�ss�, ett� Task-kent�t ovat tyhji� (jos joku aiempi vaihe ep�onnistui)
    ' (Voit kopioida nollauskoodin cmdAdd_Click:st� tai luoda erillisen ResetTaskFields-apurutiinin)
    attnData.Tila = "HUOMIO" ' Varmistetaan Tila

    Application.StatusBar = "Tallennetaan muutoksia huomioriviin..."

    ' P�ivit� olio TaskManagerissa
    tm.UpdateTask attnData

    ' Tallenna kokoelma levylle
    tm.SaveToSheet mdlMain.STORAGE_SHEET_NAME

    ' P�ivit� n�ytt�
    dm.UpdateDisplay tm.tasks, mdlMain.DISPLAY_SHEET_NAME

    Application.StatusBar = False
    MsgBox "Muutokset huomioriviin (ID: " & attnData.ID & ") tallennettu onnistuneesti!", vbInformation, "Muokkaus Onnistui"

    Unload Me ' Sulje lomake
    GoTo CleanExit_Edit

EditErrorHandler:
    Application.StatusBar = False
    MsgBox "Virhe tallennettaessa muutoksia huomioriviin:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Muokkausvirhe"
    ' �l� sulje lomaketta

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
    Dim itemToDelete As clsTaskItem ' Lis�tty tarkistusta varten

    taskIDToDelete = Me.TaskIDToEdit ' Haetaan poistettava ID lomakkeelta

    ' Varmista, ett� ollaan muokkaustilassa ja ID on validi
    If taskIDToDelete <= 0 Then
        MsgBox "Poistettavan huomiorivin ID:t� ei voitu m��ritt��. Toiminto peruttu.", vbExclamation, "Virhe"
        Exit Sub
    End If

    ' Kysy varmistus k�ytt�j�lt�
    response = MsgBox("Haluatko varmasti poistaa t�m�n huomiorivin (ID: " & taskIDToDelete & ")?" & vbCrLf & vbCrLf & _
                      "Toimintoa ei voi peruuttaa.", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "Vahvista Poisto")

    ' Jos k�ytt�j� ei halua poistaa, lopeta
    If response = vbNo Then Exit Sub

    ' Jos k�ytt�j� vahvisti poiston (vbYes)

    ' Hae Manager-oliot
    Set tm = mdlMain.GetTaskManagerInstance()
    Set dm = mdlMain.GetDisplayManagerInstance()

    If tm Is Nothing Or dm Is Nothing Then
        MsgBox "Kriittinen virhe: Sovelluksen komponentteja ei voitu alustaa. Poisto ep�onnistui.", vbCritical, "Virhe"
        Exit Sub
    End If

    On Error GoTo DeleteErrorHandler

    ' --- Lis�turvatarkistus: Varmista, ett� ID viittaa Huomioriviin ---
    Set itemToDelete = tm.GetTaskByID(taskIDToDelete)
    If itemToDelete Is Nothing Then
         MsgBox "Poistettavaa tietuetta (ID: " & taskIDToDelete & ") ei l�ytynyt. Poisto peruttu.", vbExclamation, "Virhe"
         GoTo CleanExit_Delete
    ElseIf itemToDelete.RecordType <> "Attention" Then
         MsgBox "Tietue (ID: " & taskIDToDelete & ") ei ole Huomiorivi. Poisto peruttu.", vbExclamation, "V��r� Tyyppi"
         GoTo CleanExit_Delete
    End If
    ' --- Tarkistus OK ---

    Application.StatusBar = "Poistetaan huomiorivi�..."

    ' 1. Poista teht�v� muistista TaskManagerin avulla
    tm.DeleteTask taskIDToDelete

    ' 2. Tallenna muuttunut kokoelma v�lilehdelle
    tm.SaveToSheet mdlMain.STORAGE_SHEET_NAME

    ' 3. P�ivit� n�ytt�
    dm.UpdateDisplay tm.tasks, mdlMain.DISPLAY_SHEET_NAME

    Application.StatusBar = False
    MsgBox "Huomiorivi (ID: " & taskIDToDelete & ") poistettu onnistuneesti.", vbInformation, "Poisto Onnistui"

    ' 4. Sulje lomake onnistuneen poiston j�lkeen
    Unload Me
    GoTo CleanExit_Delete ' Hypp�� siivoukseen

DeleteErrorHandler:
    Application.StatusBar = False
    MsgBox "Virhe poistettaessa huomiorivi� (ID: " & taskIDToDelete & "):" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Poistovirhe"
    ' �L� sulje lomaketta virhetilanteessa

CleanExit_Delete:
    Application.StatusBar = False ' Varmistaa, ett� status bar tyhjenee aina
    Set itemToDelete = Nothing
    Set tm = Nothing
    Set dm = Nothing
End Sub

' --- Peruuta-painikkeen toiminto ---
Private Sub cmdCancel_Click()
    Unload Me
End Sub

