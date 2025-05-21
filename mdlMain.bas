Attribute VB_Name = "mdlMain"
' --- Standard Module: mdlMain ---
Option Explicit

' --- Globaalit (moduulitasoiset) muuttujat Managereille ---
' Käytetään Private-määritystä, jotta ne ovat vain tämän moduulin käytössä,
' ellei niitä tarvita suoraan muualla (esim. lomakkeilta).
' Tarjotaan funktiot GetTaskManager/GetDisplayManager niiden hakemiseksi.
Private gTaskManager As clsTaskManager
Private gDisplayManager As clsDisplayManager

' --- Vakiot tiedostonimille ---
Public Const STORAGE_SHEET_NAME As String = "Tietovarasto"
Public Const DISPLAY_SHEET_NAME As String = "Tehtävät"
Public Const CONFIG_SHEET_NAME As String = "Config"

Public Const TASK_DATA_END_COLUMN As Long = 102
Public Const TASK_META_DATA_START_COLUMN As Long = 100
Public Const TASK_ID_COLUMN As Long = TASK_META_DATA_START_COLUMN ' 100
Public Const TASK_RECORD_TYPE_COLUMN As Long = TASK_META_DATA_START_COLUMN + 1 ' 101
Public Const TASK_ATTENTION_DATE_COLUMN As Long = TASK_META_DATA_START_COLUMN + 2 ' 102

' --- Alustaa Manager-oliot, jos niitä ei ole vielä luotu ---
' Tätä kutsutaan muiden metodien alussa varmistamaan, että oliot ovat olemassa.
Public Sub InitializeAppObjects()
    On Error GoTo Init_Error

    If gTaskManager Is Nothing Then
        ' Luo TaskManager-olio VAIN jos sitä ei ole
        Set gTaskManager = New clsTaskManager
        'Debug.Print Now & " InitializeAppObjects: Uusi TaskManager luotu. Ladataan data..."

        ' --- Lataa data heti olion luonnin jälkeen ---
        On Error Resume Next ' Käytä varovasti, jos lataus voi epäonnistua
        gTaskManager.LoadFromSheet STORAGE_SHEET_NAME ' Kutsu latausta TÄSSÄ
        If Err.Number <> 0 Then
            MsgBox "Virhe ladattaessa alkutietoja TaskManageriin:" & vbCrLf & Err.Description, vbCritical, "Alustusvirhe"
            ' Harkitse, mitä tässä virhetilanteessa tehdään. Nyt jatketaan tyhjällä kokoelmalla.
            Err.Clear
        End If
        On Error GoTo Init_Error ' Palauta päävirheenkäsittelijä
        'Debug.Print Now & " InitializeAppObjects: Datan lataus valmis. TaskManager Count: " & gTaskManager.tasks.Count ' Tulosta koko heti latauksen jälkeen
        ' ------------------------------------------------------
    Else
        ' Jos TaskManager on jo olemassa, älä tee mitään (tai tulosta vain tieto)
        'Debug.Print Now & " InitializeAppObjects: TaskManager oli jo olemassa. Count: " & gTaskManager.tasks.Count
    End If

    ' Käsittele DisplayManager samoin (luo vain jos puuttuu)
    If gDisplayManager Is Nothing Then
        Set gDisplayManager = New clsDisplayManager
        'Debug.Print Now & " InitializeAppObjects: Uusi DisplayManager luotu."
    Else
         'Debug.Print Now & " InitializeAppObjects: DisplayManager oli jo olemassa."
    End If

CleanExit_Init:
    Exit Sub

Init_Error:
    MsgBox "Kriittinen virhe alustettaessa Manager-olioita (InitializeAppObjects):" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Alustusvirhe"
    ' Nollaa oliot virhetilanteessa?
    Set gTaskManager = Nothing
    Set gDisplayManager = Nothing
    Resume CleanExit_Init
End Sub

' --- Pääaliohjelma, joka lataa datan ja päivittää näytön ---
Public Sub UpdateMainView()
    Dim tasksToDisplay As Collection ' Määritellään funktion alussa

    On Error GoTo ErrorHandler ' Aseta virheenkäsittelijä
    'Debug.Print Now & " !!! UpdateMainView ALKAA !!!"

    ' 1. Varmista, että Manager-oliot ovat alustettu ja olemassa
    InitializeAppObjects
    If gTaskManager Is Nothing Or gDisplayManager Is Nothing Then
        MsgBox "Sovelluksen pääkomponentteja (TaskManager/DisplayManager) ei voitu alustaa!", vbCritical
        GoTo CleanUp ' Poistu siististi, jos alustus epäonnistui
    End If

    ' 2. Hae suodatettu lista tehtävistä (EI ladata uudelleen levyltä!)
    '    Käyttää muistissa olevaa gTaskManager.tasks -kokoelmaa
    Set tasksToDisplay = GetFilteredTasks()

    ' 3. Päivitä näyttö VAIN KERRAN käyttäen suodatettua listaa
    Application.StatusBar = "Päivitetään näyttöä..."
    gDisplayManager.UpdateDisplay tasksToDisplay, DISPLAY_SHEET_NAME ' Anna suodatettu kokoelma

    ' 4. Näytä valmistumisviesti
    Application.StatusBar = False ' Tyhjennä statuspalkki
    'MsgBox "Näkymä '" & DISPLAY_SHEET_NAME & "' päivitetty!", vbInformation, "Valmis"

    ' 5. (Valinnainen) Tulosta lopullinen muistissa olevan kokoelman koko
    'If Not gTaskManager Is Nothing Then
        'Debug.Print Now & " UpdateMainView LOPPU: gTaskManager.tasks.Count (muistissa) = " & gTaskManager.tasks.Count
    'Else
        'Debug.Print Now & " UpdateMainView LOPPU: gTaskManager on Nothing!"
    'End If

CleanUp: ' Yhteinen poistumispiste onnistumiselle ja virheille (ErrorHandlerista tullaan tänne)
    On Error Resume Next ' Ohita virheet siivouksessa
    Set tasksToDisplay = Nothing ' Vapauta kokoelmamuuttuja
    Application.StatusBar = False ' Varmista, että status bar on tyhjä
    Exit Sub ' Normaali poistuminen

ErrorHandler:
    ' Virheenkäsittelijä: Näytä virheilmoitus
    MsgBox "Pääohjelmassa (UpdateMainView) tapahtui odottamaton virhe:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Ohjelmavirhe"
    ' Siirry siivoukseen virheen jälkeen
    Resume CleanUp
End Sub


' --- Funktiot, joilla lomakkeet tai muut moduulit voivat pyytää Manager-olioita ---
' Tämä on siistimpi tapa kuin Public-globaalit muuttujat.
Public Function GetTaskManagerInstance() As clsTaskManager
    InitializeAppObjects
    Set GetTaskManagerInstance = gTaskManager
End Function

Public Function GetDisplayManagerInstance() As clsDisplayManager
    InitializeAppObjects ' Varmista alustus
    Set GetDisplayManagerInstance = gDisplayManager
End Function

' --- ALIOHJELMAT LOMAKKEEN AVAAMISEKSI ---

' Avaa frmTehtavat-lomakkeen UUDEN tehtävän lisäämistä varten
Public Sub ShowTaskForm_AddNew()
    Dim taskForm As frmTehtavat ' Käytä oikeaa lomakkeen nimeä

    On Error GoTo ErrorHandler
    InitializeAppObjects ' Varmista, että Managerit ovat alustettu

    Set taskForm = New frmTehtavat ' Luo uusi lomake-instanssi
    taskForm.TaskIDToEdit = 0       ' Aseta moodi: Lisää uusi (ID=0)
    taskForm.Show vbModal           ' Näytä lomake modaalisesti (odottaa sulkemista)

CleanUp:
    On Error Resume Next ' Ohita virhe, jos lomake on jo purettu
    Unload taskForm      ' Pura lomake muistista
    Set taskForm = Nothing ' Vapauta muuttuja
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa lisäyslomaketta: " & Err.Description, vbCritical
    Resume CleanUp ' Yritä siivota virhetilanteessa
End Sub


' Avaa frmTehtavat-lomakkeen olemassa olevan tehtävän MUOKKAAMISTA varten
Public Sub ShowTaskForm_Edit(ByVal taskID As Long)
    Dim taskForm As frmTehtavat

    On Error GoTo ErrorHandler
    
    'Debug.Print Now & " Avaa ShowTaskForm_Edit ID:llä: " & taskID ' Kommentoitu pois
    
    InitializeAppObjects

    ' Tarkista, että annettu ID on validi
    If taskID <= 0 Then
        MsgBox "Muokattavan tehtävän ID (" & taskID & ") on virheellinen.", vbExclamation
        GoTo CleanUp
    End If

    Set taskForm = New frmTehtavat
    taskForm.TaskIDToEdit = taskID
    taskForm.Show vbModal

CleanUp:
    On Error Resume Next
    Unload taskForm
    Set taskForm = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa muokkauslomaketta ID:llä " & taskID & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' --- ALIOHJELMAT HUOMIORIVI-LOMAKKEEN AVAAMISEKSI ---

' Avaa frmHuomiorivi-lomakkeen UUDEN huomiorivin lisäämistä varten
Public Sub ShowAttentionForm_AddNew()
    Dim attnForm As frmHuomiorivi ' Käytä oikeaa lomakkeen nimeä

    On Error GoTo ErrorHandler
    InitializeAppObjects ' Varmista, että Managerit ovat alustettu

    Set attnForm = New frmHuomiorivi ' Luo uusi lomake-instanssi
    attnForm.TaskIDToEdit = 0       ' Aseta moodi: Lisää uusi (ID=0)
    attnForm.Show vbModal           ' Näytä lomake modaalisesti (odottaa sulkemista)

CleanUp:
    On Error Resume Next ' Ohita virhe, jos lomake on jo purettu
    Unload attnForm      ' Pura lomake muistista
    Set attnForm = Nothing ' Vapauta muuttuja
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa uuden huomiorivin lisäyslomaketta: " & vbCrLf & Err.Description, vbCritical, "Lomakkeen Avausvirhe"
    Resume CleanUp ' Yritä siivota virhetilanteessa
End Sub


' Avaa frmHuomiorivi-lomakkeen olemassa olevan huomiorivin MUOKKAAMISTA varten
Public Sub ShowAttentionForm_Edit(ByVal taskID As Long)
    Dim attnForm As frmHuomiorivi
    Dim tm As clsTaskManager ' Tarvitaan tyypin tarkistukseen
    Dim itemToCheck As clsTaskItem

    On Error GoTo ErrorHandler

    InitializeAppObjects ' Varmista managerit

    ' --- Tarkistukset ennen lomakkeen avaamista ---
    ' 1. Onko ID kelvollinen?
    If taskID <= 0 Then
        MsgBox "Muokattavan huomiorivin ID (" & taskID & ") on virheellinen.", vbExclamation, "Virheellinen ID"
        GoTo CleanUp
    End If

    ' 2. Varmista, että ID viittaa Huomioriviin (ei Task-riviin)
    Set tm = GetTaskManagerInstance() ' Hae TaskManager
    If Not tm Is Nothing Then
        Set itemToCheck = tm.GetTaskByID(taskID)
        If itemToCheck Is Nothing Then
            MsgBox "Tietuetta ID:llä " & taskID & " ei löytynyt.", vbExclamation, "Ei Löytynyt"
            GoTo CleanUp
        ElseIf itemToCheck.RecordType <> "Attention" Then
             MsgBox "Tietue ID:llä " & taskID & " on normaali tehtävä, ei huomiorivi." & vbCrLf & _
                    "Avaa tehtävien muokkauslomake.", vbExclamation, "Väärä Tyyppi"
             GoTo CleanUp ' Poistu, koska tämä on väärä lomake tälle tyypille
        End If
        ' Jos tyyppi oli oikea ("Attention"), vapautetaan itemToCheck
        Set itemToCheck = Nothing
    Else
        MsgBox "TaskManageria ei voitu alustaa. Muokkausta ei voi jatkaa.", vbCritical, "Virhe"
        GoTo CleanUp
    End If
    ' --- Tarkistukset OK ---


    ' Luo ja näytä lomake muokkaustilassa
    Set attnForm = New frmHuomiorivi
    attnForm.TaskIDToEdit = taskID ' Aseta ID muokkausta varten
    attnForm.Show vbModal

CleanUp:
    On Error Resume Next
    Unload attnForm
    Set attnForm = Nothing
    Set itemToCheck = Nothing ' Varmuuden vuoksi
    Set tm = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa huomiorivin muokkauslomaketta ID:llä " & taskID & ": " & vbCrLf & Err.Description, vbCritical, "Lomakkeen Avausvirhe"
    Resume CleanUp
End Sub

' --- ALIOHJELMA REKISTERITIETOJEN HALLINTALOMAKKEEN AVAAMISEKSI ---

Public Sub ShowRegisterForm()
    Dim regForm As frmRekisteri

    On Error GoTo ErrorHandler

    ' Luo uusi instanssi lomakkeesta
    Set regForm = New frmRekisteri

    ' Näytä lomake modaalisesti (koodin suoritus pysähtyy tähän, kunnes lomake suljetaan)
    regForm.Show vbModal

CleanUp:
    ' Siivoa lomake muistista, kun se suljetaan
    On Error Resume Next ' Ohita virhe, jos lomake on jo purettu
    Unload regForm
    Set regForm = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa rekisteritietojen hallintalomaketta:" & vbCrLf & Err.Description, vbCritical, "Lomakkeen Avausvirhe"
    Resume CleanUp ' Yritä siivota virhetilanteessa
End Sub

Public Function GetFilteredTasks() As Collection
    ' --- Esittele KAIKKI muuttujat ensin ---
    Dim allTasks As Collection
    Dim filteredTasks As Collection
    Dim taskItem As clsTaskItem
    Dim displaySheet As Object ' Sheet1 sijaan käytetään Objectia tai Worksheet-tyyppiä
    Dim filterMode As String
    Dim showLastausOK As Boolean
    Dim showPurkuOK As Boolean
    Dim showLaskuttamatta As Boolean
    Dim showKontaktiRowsPlaceholder As Boolean ' Muutos 2: Placeholder
    Dim anyCheckboxTicked As Boolean
    Dim filterSettings As Variant

    Set filteredTasks = New Collection

    On Error GoTo FilterError

    If gTaskManager Is Nothing Then InitializeAppObjects
    If gTaskManager Is Nothing Then GoTo FilterErrorHandler_Specific

    Set allTasks = gTaskManager.tasks
    If allTasks Is Nothing Then GoTo CleanExit_Filter

    ' Hae välilehti (varmista, että Tehtävät-välilehden koodinimi on oikein tai käytä nimeä)
    On Error Resume Next ' Jos välilehteä ei löydy
    Set displaySheet = ThisWorkbook.Worksheets(DISPLAY_SHEET_NAME)
    If displaySheet Is Nothing Then
        Err.Raise vbObjectError + 1000, "GetFilteredTasks", "Välilehteä '" & DISPLAY_SHEET_NAME & "' ei löytynyt."
    End If
    On Error GoTo FilterError ' Palauta normaali virheenkäsittely

    filterSettings = displaySheet.GetFilterSettings()

    If Not IsArray(filterSettings) Then Err.Raise vbObjectError + 1001, "GetFilteredTasks", "GetFilterSettings ei palauttanut taulukkoa."
    If UBound(filterSettings) <> 4 Then Err.Raise vbObjectError + 1002, "GetFilteredTasks", "GetFilterSettings palautti väärän kokoisen taulukon."
    
    filterMode = filterSettings(1)
    showLastausOK = filterSettings(2)
    showPurkuOK = filterSettings(3)
    showLaskuttamatta = filterSettings(4)
    
    showKontaktiRowsPlaceholder = True ' Oletus: näytä kontaktit
    anyCheckboxTicked = (showLastausOK Or showPurkuOK)

    If allTasks.Count > 0 Then
        For Each taskItem In allTasks
            Select Case taskItem.RecordType
                Case "Attention"
                    filteredTasks.Add taskItem

                Case "Kontakti"
                    If showKontaktiRowsPlaceholder Then
                        filteredTasks.Add taskItem
                    End If

                Case "Task"
                    If showLaskuttamatta Then
                        ' --- A. Laskuttamatta-suodatin ---
                        Dim isUninvoiced As Boolean: isUninvoiced = False
                        Dim laskutusValue As String
                        laskutusValue = UCase(Trim(mdlStringUtils.DefaultIfNull(taskItem.Laskutus, "EI")))
                        Select Case laskutusValue
                            Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK"
                                ' isUninvoiced remains False
                            Case Else
                                isUninvoiced = True
                        End Select
                        
                        Dim taskTilaLask As String ' Käytä eri nimeä kuin filterMode
                        taskTilaLask = UCase(Trim(taskItem.Tila))
                        If taskTilaLask = "HYVÄKSYTTY" And isUninvoiced Then
                            filteredTasks.Add taskItem
                        End If
                    Else
                        ' --- B. Normaalit suodattimet for Tasks ---
                        Dim passesPrimary As Boolean: passesPrimary = False
                        Select Case filterMode
                            Case "Kaikki": passesPrimary = True
                            Case "Tarjoukset": If UCase(taskItem.Tila) = "TARJOUS" Then passesPrimary = True
                            Case "Varmistuneet": If UCase(taskItem.Tila) = "HYVÄKSYTTY" Then passesPrimary = True
                        End Select

                        If passesPrimary Then
                            Dim passesSecondary As Boolean: passesSecondary = False
                            If Not anyCheckboxTicked Then
                                passesSecondary = True
                            Else
                                Dim lpv As Boolean, ppv As Boolean
                                lpv = taskItem.LastauspaivaVarmistunut
                                ppv = taskItem.PurkupaivaVarmistunut
                                Select Case True
                                    Case showLastausOK And Not showPurkuOK: If lpv And Not ppv Then passesSecondary = True
                                    Case Not showLastausOK And showPurkuOK: If Not lpv And ppv Then passesSecondary = True
                                    Case showLastausOK And showPurkuOK: If lpv And ppv Then passesSecondary = True
                                End Select
                            End If

                            If passesSecondary Then
                                filteredTasks.Add taskItem
                            End If ' Sulkee If passesSecondary
                        End If ' Sulkee If passesPrimary
                    End If ' Sulkee If showLaskuttamatta
            End Select ' Sulkee Select Case taskItem.RecordType
        Next taskItem
    End If

    Set GetFilteredTasks = filteredTasks
CleanExit_Filter:
    Set allTasks = Nothing
    Set taskItem = Nothing
    Set displaySheet = Nothing
    Exit Function
FilterError:
    Dim errorDesc As String
    errorDesc = "Yleinen virhe suodatettaessa tehtäviä (GetFilteredTasks):" & vbCrLf & _
                "Virhekoodi: " & Err.Number & vbCrLf & _
                "Kuvaus: " & Err.Description
    'Debug.Print "*** " & errorDesc & " ***"
    MsgBox errorDesc, vbCritical, "Suodatusvirhe"
    Set GetFilteredTasks = filteredTasks ' Palauta tyhjä tai osittainen kokoelma
    Resume CleanExit_Filter
FilterErrorHandler_Specific:
     MsgBox "Virhe suodatuksen alustuksessa:" & vbCrLf & "TaskManageria tai Tehtävät-välilehteä ei voitu alustaa/löytää. Tarkista asetukset.", vbCritical, "Suodatusvirhe"
     Set GetFilteredTasks = filteredTasks
     Resume CleanExit_Filter
End Function

