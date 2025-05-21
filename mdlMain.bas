Attribute VB_Name = "mdlMain"
' --- Standard Module: mdlMain ---
Option Explicit

' --- Globaalit (moduulitasoiset) muuttujat Managereille ---
' K�ytet��n Private-m��rityst�, jotta ne ovat vain t�m�n moduulin k�yt�ss�,
' ellei niit� tarvita suoraan muualla (esim. lomakkeilta).
' Tarjotaan funktiot GetTaskManager/GetDisplayManager niiden hakemiseksi.
Private gTaskManager As clsTaskManager
Private gDisplayManager As clsDisplayManager

' --- Vakiot tiedostonimille ---
Public Const STORAGE_SHEET_NAME As String = "Tietovarasto"
Public Const DISPLAY_SHEET_NAME As String = "Teht�v�t"
Public Const CONFIG_SHEET_NAME As String = "Config"

Public Const TASK_DATA_END_COLUMN As Long = 102
Public Const TASK_META_DATA_START_COLUMN As Long = 100
Public Const TASK_ID_COLUMN As Long = TASK_META_DATA_START_COLUMN ' 100
Public Const TASK_RECORD_TYPE_COLUMN As Long = TASK_META_DATA_START_COLUMN + 1 ' 101
Public Const TASK_ATTENTION_DATE_COLUMN As Long = TASK_META_DATA_START_COLUMN + 2 ' 102

' --- Alustaa Manager-oliot, jos niit� ei ole viel� luotu ---
' T�t� kutsutaan muiden metodien alussa varmistamaan, ett� oliot ovat olemassa.
Public Sub InitializeAppObjects()
    On Error GoTo Init_Error

    If gTaskManager Is Nothing Then
        ' Luo TaskManager-olio VAIN jos sit� ei ole
        Set gTaskManager = New clsTaskManager
        'Debug.Print Now & " InitializeAppObjects: Uusi TaskManager luotu. Ladataan data..."

        ' --- Lataa data heti olion luonnin j�lkeen ---
        On Error Resume Next ' K�yt� varovasti, jos lataus voi ep�onnistua
        gTaskManager.LoadFromSheet STORAGE_SHEET_NAME ' Kutsu latausta T�SS�
        If Err.Number <> 0 Then
            MsgBox "Virhe ladattaessa alkutietoja TaskManageriin:" & vbCrLf & Err.Description, vbCritical, "Alustusvirhe"
            ' Harkitse, mit� t�ss� virhetilanteessa tehd��n. Nyt jatketaan tyhj�ll� kokoelmalla.
            Err.Clear
        End If
        On Error GoTo Init_Error ' Palauta p��virheenk�sittelij�
        'Debug.Print Now & " InitializeAppObjects: Datan lataus valmis. TaskManager Count: " & gTaskManager.tasks.Count ' Tulosta koko heti latauksen j�lkeen
        ' ------------------------------------------------------
    Else
        ' Jos TaskManager on jo olemassa, �l� tee mit��n (tai tulosta vain tieto)
        'Debug.Print Now & " InitializeAppObjects: TaskManager oli jo olemassa. Count: " & gTaskManager.tasks.Count
    End If

    ' K�sittele DisplayManager samoin (luo vain jos puuttuu)
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

' --- P��aliohjelma, joka lataa datan ja p�ivitt�� n�yt�n ---
Public Sub UpdateMainView()
    Dim tasksToDisplay As Collection ' M��ritell��n funktion alussa

    On Error GoTo ErrorHandler ' Aseta virheenk�sittelij�
    'Debug.Print Now & " !!! UpdateMainView ALKAA !!!"

    ' 1. Varmista, ett� Manager-oliot ovat alustettu ja olemassa
    InitializeAppObjects
    If gTaskManager Is Nothing Or gDisplayManager Is Nothing Then
        MsgBox "Sovelluksen p��komponentteja (TaskManager/DisplayManager) ei voitu alustaa!", vbCritical
        GoTo CleanUp ' Poistu siististi, jos alustus ep�onnistui
    End If

    ' 2. Hae suodatettu lista teht�vist� (EI ladata uudelleen levylt�!)
    '    K�ytt�� muistissa olevaa gTaskManager.tasks -kokoelmaa
    Set tasksToDisplay = GetFilteredTasks()

    ' 3. P�ivit� n�ytt� VAIN KERRAN k�ytt�en suodatettua listaa
    Application.StatusBar = "P�ivitet��n n�ytt��..."
    gDisplayManager.UpdateDisplay tasksToDisplay, DISPLAY_SHEET_NAME ' Anna suodatettu kokoelma

    ' 4. N�yt� valmistumisviesti
    Application.StatusBar = False ' Tyhjenn� statuspalkki
    'MsgBox "N�kym� '" & DISPLAY_SHEET_NAME & "' p�ivitetty!", vbInformation, "Valmis"

    ' 5. (Valinnainen) Tulosta lopullinen muistissa olevan kokoelman koko
    'If Not gTaskManager Is Nothing Then
        'Debug.Print Now & " UpdateMainView LOPPU: gTaskManager.tasks.Count (muistissa) = " & gTaskManager.tasks.Count
    'Else
        'Debug.Print Now & " UpdateMainView LOPPU: gTaskManager on Nothing!"
    'End If

CleanUp: ' Yhteinen poistumispiste onnistumiselle ja virheille (ErrorHandlerista tullaan t�nne)
    On Error Resume Next ' Ohita virheet siivouksessa
    Set tasksToDisplay = Nothing ' Vapauta kokoelmamuuttuja
    Application.StatusBar = False ' Varmista, ett� status bar on tyhj�
    Exit Sub ' Normaali poistuminen

ErrorHandler:
    ' Virheenk�sittelij�: N�yt� virheilmoitus
    MsgBox "P��ohjelmassa (UpdateMainView) tapahtui odottamaton virhe:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Ohjelmavirhe"
    ' Siirry siivoukseen virheen j�lkeen
    Resume CleanUp
End Sub


' --- Funktiot, joilla lomakkeet tai muut moduulit voivat pyyt�� Manager-olioita ---
' T�m� on siistimpi tapa kuin Public-globaalit muuttujat.
Public Function GetTaskManagerInstance() As clsTaskManager
    InitializeAppObjects
    Set GetTaskManagerInstance = gTaskManager
End Function

Public Function GetDisplayManagerInstance() As clsDisplayManager
    InitializeAppObjects ' Varmista alustus
    Set GetDisplayManagerInstance = gDisplayManager
End Function

' --- ALIOHJELMAT LOMAKKEEN AVAAMISEKSI ---

' Avaa frmTehtavat-lomakkeen UUDEN teht�v�n lis��mist� varten
Public Sub ShowTaskForm_AddNew()
    Dim taskForm As frmTehtavat ' K�yt� oikeaa lomakkeen nime�

    On Error GoTo ErrorHandler
    InitializeAppObjects ' Varmista, ett� Managerit ovat alustettu

    Set taskForm = New frmTehtavat ' Luo uusi lomake-instanssi
    taskForm.TaskIDToEdit = 0       ' Aseta moodi: Lis�� uusi (ID=0)
    taskForm.Show vbModal           ' N�yt� lomake modaalisesti (odottaa sulkemista)

CleanUp:
    On Error Resume Next ' Ohita virhe, jos lomake on jo purettu
    Unload taskForm      ' Pura lomake muistista
    Set taskForm = Nothing ' Vapauta muuttuja
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa lis�yslomaketta: " & Err.Description, vbCritical
    Resume CleanUp ' Yrit� siivota virhetilanteessa
End Sub


' Avaa frmTehtavat-lomakkeen olemassa olevan teht�v�n MUOKKAAMISTA varten
Public Sub ShowTaskForm_Edit(ByVal taskID As Long)
    Dim taskForm As frmTehtavat

    On Error GoTo ErrorHandler
    
    'Debug.Print Now & " Avaa ShowTaskForm_Edit ID:ll�: " & taskID ' Kommentoitu pois
    
    InitializeAppObjects

    ' Tarkista, ett� annettu ID on validi
    If taskID <= 0 Then
        MsgBox "Muokattavan teht�v�n ID (" & taskID & ") on virheellinen.", vbExclamation
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
    MsgBox "Virhe avattaessa muokkauslomaketta ID:ll� " & taskID & ": " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' --- ALIOHJELMAT HUOMIORIVI-LOMAKKEEN AVAAMISEKSI ---

' Avaa frmHuomiorivi-lomakkeen UUDEN huomiorivin lis��mist� varten
Public Sub ShowAttentionForm_AddNew()
    Dim attnForm As frmHuomiorivi ' K�yt� oikeaa lomakkeen nime�

    On Error GoTo ErrorHandler
    InitializeAppObjects ' Varmista, ett� Managerit ovat alustettu

    Set attnForm = New frmHuomiorivi ' Luo uusi lomake-instanssi
    attnForm.TaskIDToEdit = 0       ' Aseta moodi: Lis�� uusi (ID=0)
    attnForm.Show vbModal           ' N�yt� lomake modaalisesti (odottaa sulkemista)

CleanUp:
    On Error Resume Next ' Ohita virhe, jos lomake on jo purettu
    Unload attnForm      ' Pura lomake muistista
    Set attnForm = Nothing ' Vapauta muuttuja
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa uuden huomiorivin lis�yslomaketta: " & vbCrLf & Err.Description, vbCritical, "Lomakkeen Avausvirhe"
    Resume CleanUp ' Yrit� siivota virhetilanteessa
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

    ' 2. Varmista, ett� ID viittaa Huomioriviin (ei Task-riviin)
    Set tm = GetTaskManagerInstance() ' Hae TaskManager
    If Not tm Is Nothing Then
        Set itemToCheck = tm.GetTaskByID(taskID)
        If itemToCheck Is Nothing Then
            MsgBox "Tietuetta ID:ll� " & taskID & " ei l�ytynyt.", vbExclamation, "Ei L�ytynyt"
            GoTo CleanUp
        ElseIf itemToCheck.RecordType <> "Attention" Then
             MsgBox "Tietue ID:ll� " & taskID & " on normaali teht�v�, ei huomiorivi." & vbCrLf & _
                    "Avaa teht�vien muokkauslomake.", vbExclamation, "V��r� Tyyppi"
             GoTo CleanUp ' Poistu, koska t�m� on v��r� lomake t�lle tyypille
        End If
        ' Jos tyyppi oli oikea ("Attention"), vapautetaan itemToCheck
        Set itemToCheck = Nothing
    Else
        MsgBox "TaskManageria ei voitu alustaa. Muokkausta ei voi jatkaa.", vbCritical, "Virhe"
        GoTo CleanUp
    End If
    ' --- Tarkistukset OK ---


    ' Luo ja n�yt� lomake muokkaustilassa
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
    MsgBox "Virhe avattaessa huomiorivin muokkauslomaketta ID:ll� " & taskID & ": " & vbCrLf & Err.Description, vbCritical, "Lomakkeen Avausvirhe"
    Resume CleanUp
End Sub

' --- ALIOHJELMA REKISTERITIETOJEN HALLINTALOMAKKEEN AVAAMISEKSI ---

Public Sub ShowRegisterForm()
    Dim regForm As frmRekisteri

    On Error GoTo ErrorHandler

    ' Luo uusi instanssi lomakkeesta
    Set regForm = New frmRekisteri

    ' N�yt� lomake modaalisesti (koodin suoritus pys�htyy t�h�n, kunnes lomake suljetaan)
    regForm.Show vbModal

CleanUp:
    ' Siivoa lomake muistista, kun se suljetaan
    On Error Resume Next ' Ohita virhe, jos lomake on jo purettu
    Unload regForm
    Set regForm = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Virhe avattaessa rekisteritietojen hallintalomaketta:" & vbCrLf & Err.Description, vbCritical, "Lomakkeen Avausvirhe"
    Resume CleanUp ' Yrit� siivota virhetilanteessa
End Sub

Public Function GetFilteredTasks() As Collection
    ' --- Esittele KAIKKI muuttujat ensin ---
    Dim allTasks As Collection
    Dim filteredTasks As Collection
    Dim taskItem As clsTaskItem
    Dim displaySheet As Object ' Sheet1 sijaan k�ytet��n Objectia tai Worksheet-tyyppi�
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

    ' Hae v�lilehti (varmista, ett� Teht�v�t-v�lilehden koodinimi on oikein tai k�yt� nime�)
    On Error Resume Next ' Jos v�lilehte� ei l�ydy
    Set displaySheet = ThisWorkbook.Worksheets(DISPLAY_SHEET_NAME)
    If displaySheet Is Nothing Then
        Err.Raise vbObjectError + 1000, "GetFilteredTasks", "V�lilehte� '" & DISPLAY_SHEET_NAME & "' ei l�ytynyt."
    End If
    On Error GoTo FilterError ' Palauta normaali virheenk�sittely

    filterSettings = displaySheet.GetFilterSettings()

    If Not IsArray(filterSettings) Then Err.Raise vbObjectError + 1001, "GetFilteredTasks", "GetFilterSettings ei palauttanut taulukkoa."
    If UBound(filterSettings) <> 4 Then Err.Raise vbObjectError + 1002, "GetFilteredTasks", "GetFilterSettings palautti v��r�n kokoisen taulukon."
    
    filterMode = filterSettings(1)
    showLastausOK = filterSettings(2)
    showPurkuOK = filterSettings(3)
    showLaskuttamatta = filterSettings(4)
    
    showKontaktiRowsPlaceholder = True ' Oletus: n�yt� kontaktit
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
                            Case "KYLL�", "K", "TRUE", "YES", "1", "-1", "OK"
                                ' isUninvoiced remains False
                            Case Else
                                isUninvoiced = True
                        End Select
                        
                        Dim taskTilaLask As String ' K�yt� eri nime� kuin filterMode
                        taskTilaLask = UCase(Trim(taskItem.Tila))
                        If taskTilaLask = "HYV�KSYTTY" And isUninvoiced Then
                            filteredTasks.Add taskItem
                        End If
                    Else
                        ' --- B. Normaalit suodattimet for Tasks ---
                        Dim passesPrimary As Boolean: passesPrimary = False
                        Select Case filterMode
                            Case "Kaikki": passesPrimary = True
                            Case "Tarjoukset": If UCase(taskItem.Tila) = "TARJOUS" Then passesPrimary = True
                            Case "Varmistuneet": If UCase(taskItem.Tila) = "HYV�KSYTTY" Then passesPrimary = True
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
    errorDesc = "Yleinen virhe suodatettaessa teht�vi� (GetFilteredTasks):" & vbCrLf & _
                "Virhekoodi: " & Err.Number & vbCrLf & _
                "Kuvaus: " & Err.Description
    'Debug.Print "*** " & errorDesc & " ***"
    MsgBox errorDesc, vbCritical, "Suodatusvirhe"
    Set GetFilteredTasks = filteredTasks ' Palauta tyhj� tai osittainen kokoelma
    Resume CleanExit_Filter
FilterErrorHandler_Specific:
     MsgBox "Virhe suodatuksen alustuksessa:" & vbCrLf & "TaskManageria tai Teht�v�t-v�lilehte� ei voitu alustaa/l�yt��. Tarkista asetukset.", vbCritical, "Suodatusvirhe"
     Set GetFilteredTasks = filteredTasks
     Resume CleanExit_Filter
End Function

