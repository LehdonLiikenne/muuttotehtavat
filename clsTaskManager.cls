VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Class Module: clsTaskManager ---
Option Explicit

' Yksityinen kokoelma TaskItem-olioiden s‰ilytt‰miseen
Private pTasks As Collection
Private Const TASK_DATA_START_ROW = 2
Private Const TASK_DATA_END_COLUMN As Long = 102 ' Viimeinen sarake, jota luetaan dataArray:hyn
' Private Const TASK_META_DATA_START_COLUMN As Long = 100 ' Ei suoraan k‰ytˆss‰ LoadFromSheetin yleislogiikassa
Private Const TASK_ID_COLUMN As Long = 100 ' ID:n sarake Tietovarastossa (CV)
Private Const TASK_RECORD_TYPE_COLUMN As Long = 101 ' RecordTypen sarake (CW)
Private Const TASK_ATTENTION_DATE_COLUMN As Long = 102 ' AttentionSortDaten sarake (CX)
Private Const TASK_LAST_DATA_COLUMN As Long = 102 ' K‰ytet‰‰n dataRange:n m‰‰rittelyyn

' Alustaa kokoelman automaattisesti, kun olio luodaan
Private Sub Class_Initialize()
    Init
End Sub

' Julkinen metodi kokoelman alustamiseen tai tyhjent‰miseen
Public Sub Init()
    Set pTasks = New Collection
End Sub

' Sallii kokoelman lukemisen ulkopuolelta (vain luku)
Public Property Get tasks() As Collection
    Set tasks = pTasks
End Property

' Lataa TaskItem-oliot m‰‰ritellylt‰ v‰lilehdelt‰
' Tiedostossa: clsTaskManager.txt
' Korvaa koko LoadFromSheet-metodi t‰ll‰:

Public Sub LoadFromSheet(ByVal storageSheetName As String)
    ' Debug.Print Now & " $$$ LoadFromSheet ALKAA $$$ Sheet: " & storageSheetName ' T‰m‰ voi olla aktiivinen, jos haluat

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim r As Long
    Dim c As Long
    Dim currentTask As clsTaskItem
    Dim cellValue As Variant

    On Error GoTo ErrorHandler_LoadFromSheet ' YLEINEN virheenk‰sittelij‰ koko metodille

    Set ws = ThisWorkbook.Worksheets(storageSheetName)
    Me.Init ' Varmistaa, ett‰ pTasks-kokoelma on tyhj‰ ennen latausta
    
    lastRow = mdlSheetUtils.GetLastRow(ws, TASK_ID_COLUMN) ' Olettaen, ett‰ ID-sarake on luotettavin viimeisen rivin m‰‰ritt‰miseen

    If lastRow < TASK_DATA_START_ROW Then
        ' Debug.Print "LoadFromSheet: Ei ladattavaa dataa." ' T‰m‰ voi olla aktiivinen
        GoTo CleanUp_LoadFromSheet
    End If

    Set dataRange = ws.Range(ws.Cells(TASK_DATA_START_ROW, 1), ws.Cells(lastRow, TASK_LAST_DATA_COLUMN))
    
    If dataRange.rows.Count = 1 Then
        ReDim dataArray(1 To 1, 1 To TASK_LAST_DATA_COLUMN)
        For c = 1 To TASK_LAST_DATA_COLUMN
            dataArray(1, c) = dataRange.Cells(1, c).value
        Next c
    Else
        dataArray = dataRange.value
    End If

    For r = LBound(dataArray, 1) To UBound(dataArray, 1)
        Set currentTask = New clsTaskItem
        
        On Error GoTo NextRowOnError_LoadFromSheet ' RIVIKOHTAINEN virheenk‰sittelij‰ PƒƒLLE
        ' Debug.Print "LoadFromSheet: K‰sitell‰‰n dataArray-rivi " & r & " (Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArray, 1)) & ")"

        ' Kent‰t 1-36
        currentTask.asiakas = SafeReadString(dataArray, r, 1)
        currentTask.tarjousTehty = SafeReadDate(dataArray, r, 2)
        currentTask.lastausPaiva = SafeReadDate(dataArray, r, 3)
        currentTask.lastausMaa = SafeReadString(dataArray, r, 4)
        currentTask.purkuMaa = SafeReadString(dataArray, r, 5)
        currentTask.purkuPaiva = SafeReadDate(dataArray, r, 6)
        currentTask.M3m = SafeReadString(dataArray, r, 7)
        currentTask.palvelu = SafeReadString(dataArray, r, 8)
        currentTask.Huomioitavaa = SafeReadString(dataArray, r, 9)
        currentTask.puhelin = SafeReadString(dataArray, r, 10)
        currentTask.lastausOsoite = SafeReadString(dataArray, r, 11)
        currentTask.purkuOsoite = SafeReadString(dataArray, r, 12)
        currentTask.Kuljettajat = SafeReadString(dataArray, r, 13)
        currentTask.Autot = SafeReadString(dataArray, r, 14)
        currentTask.Kontit = SafeReadString(dataArray, r, 15)
        currentTask.Apulaiset = SafeReadString(dataArray, r, 16)
        currentTask.Rahtikirja = SafeReadString(dataArray, r, 17)
        currentTask.ApulaisetTilattu = SafeReadString(dataArray, r, 18)
        currentTask.Pysakointilupa = SafeReadString(dataArray, r, 19)
        currentTask.hissi = SafeReadString(dataArray, r, 20)
        currentTask.Laivalippu = SafeReadString(dataArray, r, 21)
        currentTask.Laskutus = SafeReadString(dataArray, r, 22)
        currentTask.Vakuutus = SafeReadString(dataArray, r, 23)
        currentTask.Arvo = SafeReadVariant(dataArray, r, 24)
        currentTask.hinta = SafeReadVariant(dataArray, r, 25)
        currentTask.Muuttomaailma = SafeReadString(dataArray, r, 26)
        currentTask.M3t = SafeReadString(dataArray, r, 27)
        currentTask.LastauspaivaVarmistunut = SafeReadBoolean(dataArray, r, 28)
        currentTask.PurkupaivaVarmistunut = SafeReadBoolean(dataArray, r, 29)
        currentTask.TarjousHyvaksytty = SafeReadDate(dataArray, r, 30)
        currentTask.TarjousHylatty = SafeReadDate(dataArray, r, 31)
        currentTask.sahkoposti = SafeReadString(dataArray, r, 32)
        currentTask.Tila = SafeReadString(dataArray, r, 33)
        currentTask.LastausLoppuu = SafeReadDate(dataArray, r, 34)
        currentTask.PurkuLoppuu = SafeReadDate(dataArray, r, 35)
        currentTask.kontaktiPaiva = SafeReadDate(dataArray, r, 36)
        currentTask.valimatka = SafeReadString(dataArray, r, 37)

        ' Metadata (sarakkeet 100, 101, 102)
        
        ' --- ID:n LUKU JA KƒSITTELY ALKAA TƒSTƒ (SISƒLTƒƒ UUDET DEBUG.PRINTIT) ---
        ' Debug.Print "  LoadFromSheet: Luetaan ID sarakkeesta " & TASK_ID_COLUMN & " rivilt‰ r=" & r ' Pidet‰‰n t‰m‰ kommentoituna
        cellValue = SafeReadVariant(dataArray, r, TASK_ID_COLUMN)
        Dim idProcessingStr As String
        
        ' ID:n muunnos stringiksi ja sen tarkistus
        If IsError(cellValue) Then
            idProcessingStr = "0"
            'Debug.Print "    ID_Debug: Raw cellValue (col " & TASK_ID_COLUMN & ", Excel Row ~" & (TASK_DATA_START_ROW + r - 1) & ") was Error. idProcessingStr set to '0'."
        ElseIf IsNull(cellValue) Then
            idProcessingStr = "0"
            'Debug.Print "    ID_Debug: Raw cellValue (col " & TASK_ID_COLUMN & ", Excel Row ~" & (TASK_DATA_START_ROW + r - 1) & ") was Null. idProcessingStr set to '0'."
        ElseIf IsEmpty(cellValue) Then
            idProcessingStr = "0"
            'Debug.Print "    ID_Debug: Raw cellValue (col " & TASK_ID_COLUMN & ", Excel Row ~" & (TASK_DATA_START_ROW + r - 1) & ") was Empty. idProcessingStr set to '0'."
        Else
            Dim tempStrForId As String
            On Error Resume Next ' Suojaa CStr-kutsua
            tempStrForId = CStr(cellValue)
            If Err.Number <> 0 Then
                'Debug.Print "    ID_Debug: VIRHE CStr(cellValue) ID:n luvussa. Excel Row ~" & (TASK_DATA_START_ROW + r - 1) & ", cellValue tyyppi: " & TypeName(cellValue) & ". Asetetaan idProcessingStr arvoksi '0'."
                idProcessingStr = "0"
                Err.Clear
            Else
                idProcessingStr = tempStrForId
                'Debug.Print "    ID_Debug: Raw cellValue (col " & TASK_ID_COLUMN & ", Excel Row ~" & (TASK_DATA_START_ROW + r - 1) & ", Type: " & TypeName(cellValue) & "): '" & cellValue & "'. Successfully converted to tempStrForId: '" & tempStrForId & "'"
            End If
            On Error GoTo NextRowOnError_LoadFromSheet
        End If
        
        ' Tarkemmat Debug-tulostukset ID:n k‰sittelylle
        'Debug.Print "    ID_Debug: Final idProcessingStr BEFORE IsNumeric check (Excel Row ~" & (TASK_DATA_START_ROW + r - 1) & "): '" & idProcessingStr & "'"
        Dim isNum As Boolean
        isNum = IsNumeric(idProcessingStr)
        'Debug.Print "    ID_Debug: IsNumeric(idProcessingStr) result: " & isNum
        
        Dim valResult As Double
        If isNum Then
            valResult = val(idProcessingStr)
            'Debug.Print "    ID_Debug: Val(idProcessingStr) result: " & valResult
            'Debug.Print "    ID_Debug: Condition (Val(idProcessingStr) > 0) result: " & (valResult > 0)
        Else
            valResult = 0 ' Asetetaan oletusarvo, jos ei numeerinen, jotta myˆhempi viittaus ei kaadu
            'Debug.Print "    ID_Debug: Val() not called because IsNumeric was False."
        End If
        
        If isNum And valResult > 0 Then
            currentTask.ID = CLng(valResult)
            ' Debug.Print "    ID_Debug: currentTask.ID set to: " & currentTask.ID & " (from CLng(valResult))" ' T‰m‰ voi pysy‰ kommentoituna
        Else
            currentTask.ID = 0
            ' Alkuper‰inen varoitusviesti, joka n‰kyi lokissa:
            Debug.Print "  Varoitus LoadFromSheet: Virheellinen/puuttuva ID Excel-rivill‰ n. " & (TASK_DATA_START_ROW + r - LBound(dataArray, 1)) & ". Alkuper‰inen arvo (idProcessingStr) sarakk. " & TASK_ID_COLUMN & ": '" & idProcessingStr & "'"
        End If
        ' Debug.Print "    currentTask.ID asetettu: " & currentTask.ID ' Kommentoitu, koska yll‰ oleva logiikka kattaa t‰m‰n
        ' --- ID:N LUKU JA KƒSITTELY LOPPUU ---

        ' --- RecordTypen luku ---
        ' Debug.Print "  LoadFromSheet: Luetaan RecordType sarakkeesta " & TASK_RECORD_TYPE_COLUMN & " rivilt‰ r=" & r ' Kommentoitu
        currentTask.RecordType = Trim$(SafeReadString(dataArray, r, TASK_RECORD_TYPE_COLUMN, "Task"))
        ' Debug.Print "    currentTask.RecordType asetettu: '" & currentTask.RecordType & "'" ' Kommentoitu
        If currentTask.RecordType <> "Task" And currentTask.RecordType <> "Attention" And currentTask.RecordType <> "Kontakti" Then
             ' Debug.Print "  Varoitus LoadFromSheet: Virheellinen RecordType '" & currentTask.RecordType & "' ... Oletetaan 'Task'." ' Kommentoitu
             currentTask.RecordType = "Task"
        End If

        ' --- AttentionSortDaten luku ---
        ' Debug.Print "  LoadFromSheet: Luetaan AttentionSortDate sarakkeesta " & TASK_ATTENTION_DATE_COLUMN & " rivilt‰ r=" & r ' Kommentoitu
        currentTask.AttentionSortDate = SafeReadDate(dataArray, r, TASK_ATTENTION_DATE_COLUMN)
        
        ' ---- Rivin k‰sittely valmis, yritet‰‰n lis‰t‰ kokoelmaan ----
        On Error GoTo ErrorHandler_LoadFromSheet ' YLEINEN virheenk‰sittelij‰ PƒƒLLE ennen kokoelmaoperaatioita

        If currentTask.ID > 0 Then
            Dim existingItem As clsTaskItem
            Set existingItem = Nothing
            On Error Resume Next
            Set existingItem = pTasks(CStr(currentTask.ID))
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo ErrorHandler_LoadFromSheet

            If existingItem Is Nothing Then
                pTasks.Add currentTask, CStr(currentTask.ID)
            Else
                ' Debug.Print "  Varoitus LoadFromSheet: Duplikaatti ID " & currentTask.ID & " ohitetaan." ' Kommentoitu
                Set existingItem = Nothing
            End If
        Else
            ' T‰m‰ Debug.Print on se, joka tulostaa "Varoitus LoadFromSheet: Ohitetaan rivin lis‰ys, ID <= 0..."
            ' J‰tet‰‰n se aktiiviseksi, koska se kertoo ongelmasta.
            Debug.Print "  Varoitus LoadFromSheet: Ohitetaan rivin lis‰ys, ID <= 0 (Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArray, 1)) & ")."
        End If
        
        GoTo ContinueLoop_LoadFromSheet

NextRowOnError_LoadFromSheet:
        Debug.Print "KRIITTINEN VIRHE LoadFromSheet: K‰sitelt‰ess‰ dataArray-rivi‰ " & r & " (Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArray, 1)) & "), Virhe " & Err.Number & ": " & Err.Description
        Err.Clear

ContinueLoop_LoadFromSheet:
        Set currentTask = Nothing
    Next r

    On Error GoTo ErrorHandler_LoadFromSheet
    Debug.Print "Ladattu " & pTasks.Count & " tietuetta v‰lilehdelt‰ '" & storageSheetName & "'."

CleanUp_LoadFromSheet:
    Set ws = Nothing
    Set dataRange = Nothing
    Set currentTask = Nothing
    If IsArray(dataArray) Then Erase dataArray
    On Error GoTo 0
    Exit Sub

ErrorHandler_LoadFromSheet:
    Debug.Print "--- Virhe ErrorHandler_LoadFromSheet ---"
    Debug.Print "Virhe " & Err.Number & ": " & Err.Description
    ' MsgBox "Yleinen virhe ladattaessa tietoja: " & Err.Description, vbCritical, "Latausvirhe"
    GoTo CleanUp_LoadFromSheet
End Sub

' --- Apufunktiot turvalliseen lukemiseen dataArray:sta ---
Private Function SafeReadString(dataArr As Variant, r As Long, c As Long, Optional defaultValue As String = "") As String
    Dim val As Variant
    'Debug.Print "    SafeReadString KUTSUTTU: r=" & r & ", c=" & c
    
    On Error Resume Next
    val = dataArr(r, c)
    If Err.Number <> 0 Then
        'Debug.Print "      SafeReadString: VIRHE dataArr(r, c) -kutsussa. c=" & c & ", Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArr, 1)) & ", Virhe: " & Err.Description
        SafeReadString = defaultValue
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If IsNull(val) Then
        'Debug.Print "      SafeReadString: Luettu val (r=" & r & ",c=" & c & "): NULL (tyyppi: " & TypeName(val) & ")"
        SafeReadString = defaultValue
    ElseIf IsEmpty(val) Then
        'Debug.Print "      SafeReadString: Luettu val (r=" & r & ",c=" & c & "): EMPTY (tyyppi: " & TypeName(val) & ")"
        SafeReadString = defaultValue
    ElseIf IsError(val) Then
        'Debug.Print "      SafeReadString: Luettu val (r=" & r & ",c=" & c & "): ERROR " & CStr(CVErr(val)) & " (tyyppi: " & TypeName(val) & ")"
        SafeReadString = defaultValue
    Else
        'Debug.Print "      SafeReadString: Luettu val (r=" & r & ",c=" & c & "): '" & val & "' (tyyppi: " & TypeName(val) & ")"
        'Debug.Print "      SafeReadString: Yritet‰‰n CStr(val)."
        On Error Resume Next
        SafeReadString = CStr(val)
        If Err.Number <> 0 Then
            'Debug.Print "        SafeReadString: VIRHE CStr(val) -kutsussa. val oli tyyppi‰ '" & TypeName(val) & "'. Palautetaan defaultValue. Virhe: " & Err.Description
            SafeReadString = defaultValue
            Err.Clear
        Else
            'Debug.Print "        SafeReadString: CStr(val) onnistui: '" & SafeReadString & "'"
        End If
        On Error GoTo 0
    End If
End Function

Private Function SafeReadDate(dataArr As Variant, r As Long, c As Long) As Variant
    Dim val As Variant
    'Debug.Print "    SafeReadDate KUTSUTTU: r=" & r & ", c=" & c
    
    On Error Resume Next
    val = dataArr(r, c)
    If Err.Number <> 0 Then
        'Debug.Print "      SafeReadDate: VIRHE dataArr(r, c) -kutsussa. c=" & c & ", Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArr, 1)) & ", Virhe: " & Err.Description
        SafeReadDate = Null
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    If IsNull(val) Then
        'Debug.Print "      SafeReadDate: Luettu val (r=" & r & ",c=" & c & "): NULL (tyyppi: " & TypeName(val) & ")"
    ElseIf IsEmpty(val) Then
        'Debug.Print "      SafeReadDate: Luettu val (r=" & r & ",c=" & c & "): EMPTY (tyyppi: " & TypeName(val) & ")"
    ElseIf IsError(val) Then
        'Debug.Print "      SafeReadDate: Luettu val (r=" & r & ",c=" & c & "): ERROR " & CStr(CVErr(val)) & " (tyyppi: " & TypeName(val) & ")"
    Else
        'Debug.Print "      SafeReadDate: Luettu val (r=" & r & ",c=" & c & "): '" & val & "' (tyyppi: " & TypeName(val) & ")"
    End If

    If IsError(val) Then
        'Debug.Print "      SafeReadDate: val ON VIRHE. Palautetaan Null."
        SafeReadDate = Null
    Else
        'Debug.Print "      SafeReadDate: Yritet‰‰n mdlDateUtils.ConvertToDate(val)."
        On Error Resume Next
        SafeReadDate = mdlDateUtils.ConvertToDate(val)
        If Err.Number <> 0 Then
            'Debug.Print "        SafeReadDate: VIRHE mdlDateUtils.ConvertToDate -kutsussa. Palautetaan Null. Virhe: " & Err.Description
            SafeReadDate = Null
            Err.Clear
        Else
            If IsNull(SafeReadDate) Then
                 'Debug.Print "        SafeReadDate: mdlDateUtils.ConvertToDate palautti Null."
            Else
                 'Debug.Print "        SafeReadDate: mdlDateUtils.ConvertToDate onnistui: '" & CDate(SafeReadDate) & "'"
            End If
        End If
        On Error GoTo 0
    End If
End Function

Private Function SafeReadVariant(dataArr As Variant, r As Long, c As Long) As Variant
    Dim val As Variant
    'Debug.Print "    SafeReadVariant KUTSUTTU: r=" & r & ", c=" & c
    
    On Error Resume Next
    val = dataArr(r, c)
    If Err.Number <> 0 Then
        'Debug.Print "      SafeReadVariant: VIRHE dataArr(r, c) -kutsussa. c=" & c & ", Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArr, 1)) & ", Virhe: " & Err.Description
        SafeReadVariant = Null
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If IsNull(val) Then
        'Debug.Print "      SafeReadVariant: Luettu val (r=" & r & ",c=" & c & "): NULL (tyyppi: " & TypeName(val) & ")"
    ElseIf IsEmpty(val) Then
        'Debug.Print "      SafeReadVariant: Luettu val (r=" & r & ",c=" & c & "): EMPTY (tyyppi: " & TypeName(val) & ")"
    ElseIf IsError(val) Then
        'Debug.Print "      SafeReadVariant: Luettu val (r=" & r & ",c=" & c & "): ERROR " & CStr(CVErr(val)) & " (tyyppi: " & TypeName(val) & ")"
    Else
        'Debug.Print "      SafeReadVariant: Luettu val (r=" & r & ",c=" & c & "): '" & val & "' (tyyppi: " & TypeName(val) & ")"
    End If
    
    If IsError(val) Then
        'Debug.Print "      SafeReadVariant: val ON VIRHE. Palautetaan Null."
        SafeReadVariant = Null
    Else
        SafeReadVariant = val
        'Debug.Print "      SafeReadVariant: Palautetaan val."
    End If
End Function

Private Function SafeReadBoolean(dataArr As Variant, r As Long, c As Long) As Boolean
    Dim val As Variant
    'Debug.Print "    SafeReadBoolean KUTSUTTU: r=" & r & ", c=" & c

    On Error Resume Next
    val = dataArr(r, c)
    If Err.Number <> 0 Then
        'Debug.Print "      SafeReadBoolean: VIRHE dataArr(r, c) -kutsussa. c=" & c & ", Excel-rivi n. " & (TASK_DATA_START_ROW + r - LBound(dataArr, 1)) & ", Virhe: " & Err.Description
        SafeReadBoolean = False
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If IsNull(val) Then
        'Debug.Print "      SafeReadBoolean: Luettu val (r=" & r & ",c=" & c & "): NULL (tyyppi: " & TypeName(val) & ")"
    ElseIf IsEmpty(val) Then
        'Debug.Print "      SafeReadBoolean: Luettu val (r=" & r & ",c=" & c & "): EMPTY (tyyppi: " & TypeName(val) & ")"
    ElseIf IsError(val) Then
        'Debug.Print "      SafeReadBoolean: Luettu val (r=" & r & ",c=" & c & "): ERROR " & CStr(CVErr(val)) & " (tyyppi: " & TypeName(val) & ")"
    Else
        'Debug.Print "      SafeReadBoolean: Luettu val (r=" & r & ",c=" & c & "): '" & CStr(val) & "' (tyyppi: " & TypeName(val) & ")"
    End If

    If IsError(val) Then
        'Debug.Print "      SafeReadBoolean: val ON VIRHE. Palautetaan False."
        SafeReadBoolean = False
    ElseIf IsNull(val) Or IsEmpty(val) Then
        'Debug.Print "      SafeReadBoolean: val ON NULL tai EMPTY. Palautetaan False."
        SafeReadBoolean = False
    Else
        'Debug.Print "      SafeReadBoolean: Yritet‰‰n muuntaa Booleaniksi."
        Dim strVal As String
        On Error Resume Next
        strVal = UCase(CStr(val))
        If Err.Number <> 0 Then
            'Debug.Print "        SafeReadBoolean: VIRHE CStr(val) boolean muunnoksessa. val oli tyyppi‰ '" & TypeName(val) & "'. Palautetaan False. Virhe: " & Err.Description
            SafeReadBoolean = False
            Err.Clear
        Else
            Select Case strVal
                Case "TRUE", "-1", "1", "YES", "KYLLƒ", "K"
                    SafeReadBoolean = True
                    'Debug.Print "        SafeReadBoolean: Muunnettu True."
                Case Else
                    SafeReadBoolean = False
                    'Debug.Print "        SafeReadBoolean: Muunnettu False."
            End Select
        End If
        On Error GoTo 0
    End If
End Function


' Tallentaa KOKO pTasks-kokoelman m‰‰ritellylle v‰lilehdelle, KORVATEN vanhan datan.
Public Sub SaveToSheet(ByVal storageSheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputArray() As Variant ' Taulukko, johon data kootaan kirjoitusta varten
    Dim taskItem As clsTaskItem
    Dim r As Long ' Rivi-indeksi outputArray:lle

    On Error GoTo ErrorHandler
    Application.StatusBar = "Tallennetaan tietoja..."

    ' --- 1. Hae Kohdev‰lilehti ---
    Set ws = Nothing
    On Error Resume Next ' Vaimennetaan hetkeksi
    Set ws = ThisWorkbook.Worksheets(storageSheetName)
    On Error GoTo ErrorHandler ' Palautetaan normaali
    If ws Is Nothing Then
        Err.Raise vbObjectError + 520, "TaskManager.SaveToSheet", "V‰lilehte‰ '" & storageSheetName & "' ei lˆytynyt."
        GoTo CleanUp
    End If

    ' --- 2. Poista V‰lilehden Suojaus (jos k‰ytˆss‰) ---
    ' (Lis‰‰ koodi tarvittaessa)

    ' --- 3. Tyhjenn‰ Olemassaoleva Data ---
    ' MUOKATTU: K‰yt‰ ID-saraketta (36) viimeisen rivin m‰‰ritt‰miseen
    lastRow = mdlSheetUtils.GetLastRow(ws, TASK_ID_COLUMN)
    'Debug.Print "SaveToSheet: Viimeinen rivi ID-sarakkeen (" & TASK_ID_COLUMN & ") perusteella: " & lastRow ' Debug-tuloste

    If lastRow >= TASK_DATA_START_ROW Then
        Dim clearRange As Range
        ' M‰‰rit‰ tyhjennett‰v‰ alue (A2:AL<viimeinen>)
        Set clearRange = ws.Range(ws.Cells(TASK_DATA_START_ROW, 1), ws.Cells(lastRow, TASK_LAST_DATA_COLUMN))
        Debug.Print "SaveToSheet: Tyhjennet‰‰n alue: " & clearRange.address ' Debug-tuloste
        mdlSheetUtils.ClearRangeContents clearRange
        ' Harkitse, haluatko poistaa myˆs muotoilut:
        ' mdlSheetUtils.ClearRangeFormats clearRange
        Set clearRange = Nothing
    Else
         'Debug.Print "SaveToSheet: Ei tyhjennett‰v‰‰ dataa (lastRow=" & lastRow & ", TASK_DATA_START_ROW= " & TASK_DATA_START_ROW & ")" ' Debug-tuloste
    End If

    ' --- 4. Valmistele Kirjoitettava Data Taulukkoon ---
    If pTasks Is Nothing Or pTasks.Count = 0 Then
        'Debug.Print "SaveToSheet: Ei tallennettavia teht‰vi‰ muistissa."
        GoTo ProtectAndCleanUp ' Ei dataa, siirry suojaamaan v‰lilehti
    End If

    ' M‰‰rit‰ taulukon koko kokoelman perusteella
    ReDim outputArray(1 To pTasks.Count, 1 To TASK_LAST_DATA_COLUMN)
    ' --- DEBUG version of the loop inside clsTaskManager.SaveToSheet ---
    r = 0
'Debug.Print "SaveToSheet: Aloitetaan silmukka. Tallennettavia: " & pTasks.Count

' K‰ytet‰‰n metodin p‰‰virheenk‰sittelij‰‰ (On Error GoTo ErrorHandler)
On Error GoTo ErrorHandler

For Each taskItem In pTasks
    r = r + 1 ' Kasvata taulukon rivi-indeksi‰

    ' --- Kopioi clsTaskItem-olion ominaisuudet taulukon sarakkeisiin (1-38) ---
    ' K‰ytet‰‰n suoria sijoituksia (Variant-taulukko hyv‰ksyy eri tyyppej‰)

    outputArray(r, 1) = taskItem.asiakas                                 ' String
    outputArray(r, 2) = taskItem.tarjousTehty                             ' Variant Date/Null/Empty
    outputArray(r, 3) = taskItem.lastausPaiva                            ' Variant Date/Null/Empty
    outputArray(r, 4) = taskItem.lastausMaa                               ' String
    outputArray(r, 5) = taskItem.purkuMaa                                 ' String
    outputArray(r, 6) = taskItem.purkuPaiva                              ' Variant Date/Null/Empty
    outputArray(r, 7) = taskItem.M3m                                      ' String
    outputArray(r, 8) = taskItem.palvelu                                  ' String
    outputArray(r, 9) = taskItem.Huomioitavaa                           ' String
    outputArray(r, 10) = taskItem.puhelin                                ' String
    outputArray(r, 11) = taskItem.lastausOsoite                          ' String
    outputArray(r, 12) = taskItem.purkuOsoite                            ' String
    outputArray(r, 13) = taskItem.Kuljettajat                            ' String
    outputArray(r, 14) = taskItem.Autot                                  ' String
    outputArray(r, 15) = taskItem.Kontit                                 ' String
    outputArray(r, 16) = taskItem.Apulaiset                              ' String
    outputArray(r, 17) = taskItem.Rahtikirja                             ' String
    outputArray(r, 18) = taskItem.ApulaisetTilattu                       ' String
    outputArray(r, 19) = taskItem.Pysakointilupa                         ' String
    outputArray(r, 20) = taskItem.hissi                                  ' String
    outputArray(r, 21) = taskItem.Laivalippu                             ' String
    outputArray(r, 22) = taskItem.Laskutus                               ' String
    outputArray(r, 23) = taskItem.Vakuutus                               ' String
    outputArray(r, 24) = taskItem.Arvo                                   ' Variant
    outputArray(r, 25) = taskItem.hinta                                  ' Variant
    outputArray(r, 26) = taskItem.Muuttomaailma                          ' String
    outputArray(r, 27) = taskItem.M3t                                    ' String
    outputArray(r, 28) = taskItem.LastauspaivaVarmistunut               ' Boolean
    outputArray(r, 29) = taskItem.PurkupaivaVarmistunut                ' Boolean
    outputArray(r, 30) = taskItem.TarjousHyvaksytty                      ' Variant Date/Null/Empty
    outputArray(r, 31) = taskItem.TarjousHylatty                         ' Variant Date/Null/Empty
    outputArray(r, 32) = taskItem.sahkoposti                            ' String sahkosposti
    outputArray(r, 33) = taskItem.Tila                                   ' AG = String (Tila)
    outputArray(r, 34) = taskItem.LastausLoppuu                          ' AH = Variant Date/Null/Empty
    outputArray(r, 35) = taskItem.PurkuLoppuu                            ' AI = Variant Date/Null/Empty
    outputArray(r, 36) = taskItem.kontaktiPaiva
    outputArray(r, 37) = taskItem.valimatka
    
    outputArray(r, 100) = taskItem.ID                                     ' AJ = Long
    outputArray(r, 101) = taskItem.RecordType                             ' AK = String
    outputArray(r, 102) = taskItem.AttentionSortDate                      ' AL = Variant Date/Null/Empty

Next taskItem
'Debug.Print "SaveToSheet: Silmukka suoritettu."
    ' --- The rest of SaveToSheet continues... ---

    ' --- 5. Kirjoita Koko Taulukko Kerralla V‰lilehdelle ---
    ws.Cells(TASK_DATA_START_ROW, 1).Resize(pTasks.Count, TASK_LAST_DATA_COLUMN).value = outputArray

    'Debug.Print "Tallennettu " & pTasks.Count & " tietuetta v‰lilehdelle '" & storageSheetName & "'."

ProtectAndCleanUp:
    ' --- 6. Palauta V‰lilehden Suojaus (k‰yt‰ myˆhemmin mdlSheetUtils?) ---
    ' On Error Resume Next
    ' ws.Protect Password:="YourPassword"
    ' On Error GoTo 0

CleanUp:
    Application.StatusBar = False
    Set ws = Nothing
    Set taskItem = Nothing
    If IsArray(outputArray) Then Erase outputArray ' Vapauta taulukon muisti
    Exit Sub

ErrorHandler:
    MsgBox "Virhe TaskManager.SaveToSheet: " & Err.Description & vbCrLf & "(Virhe " & Err.Number & ")", vbCritical, "Tallennusvirhe"
    ' Yrit‰ silti suojata v‰lilehti
    GoTo ProtectAndCleanUp
End Sub


' Apufunktio, joka palauttaa arvon tai Null, sopii kirjoitukseen
' Kirjoittaa Date-tyypit Date-arvoina, muuten Null jos tyhj‰/Null.
Private Function GetValueOrNullForSaving(value As Variant) As Variant
    If IsDate(value) Then
        GetValueOrNullForSaving = CDate(value) ' Kirjoita Date-tyyppin‰
    ElseIf IsNull(value) Or IsEmpty(value) Or CStr(value) = "" Then
        GetValueOrNullForSaving = Null ' Kirjoita Null (-> tyhj‰ solu)
    Else
        GetValueOrNullForSaving = value ' Kirjoita alkuper‰inen arvo (esim. numero/teksti)
    End If
End Function


' --- Metodit kokoelman k‰sittelyyn MUISTISSA ---

' Lis‰‰ uuden TaskItem-olion pTasks-kokoelmaan.
' Antaa ID:n tarvittaessa.
Public Sub AddTask(ByVal newTask As clsTaskItem)
    On Error GoTo AddTaskError

    If pTasks Is Nothing Then Me.Init ' Varmista, ett‰ kokoelma on alustettu

    ' Anna uusi ID, jos ID puuttuu tai on 0
    If newTask.ID <= 0 Then
        newTask.ID = mdlConfigUtils.GetNextID()
        If newTask.ID <= 0 Then ' Tarkista, ep‰onnistuiko ID:n haku
             Err.Raise vbObjectError + 515, "TaskManager.AddTask", "Ei saatu haettua validia uutta ID:t‰."
             Exit Sub
        End If
        'Debug.Print "Annettiin uusi ID: " & newTask.ID
    End If

    ' Tarkista duplikaatti-ID ennen lis‰yst‰
    Dim key As String
    key = CStr(newTask.ID)
    On Error Resume Next ' Tarkista, onko avain jo k‰ytˆss‰
    Dim existing As clsTaskItem
    Set existing = pTasks(key)
    On Error GoTo AddTaskError ' Palauta virheenk‰sittely

    If Not existing Is Nothing Then
        Err.Raise vbObjectError + 516, "TaskManager.AddTask", "TaskItem ID:ll‰ " & newTask.ID & " on jo kokoelmassa."
        Exit Sub
    End If

    ' Lis‰‰ uusi olio kokoelmaan, k‰yt‰ ID:t‰ avaimena
    pTasks.Add newTask, key
    'Debug.Print "Lis‰tty TaskItem ID:ll‰: " & newTask.ID

    Exit Sub
AddTaskError:
     MsgBox "Virhe TaskManager.AddTask: " & Err.Description, vbCritical, "Lis‰ysvirhe"
     'Stop
End Sub


' P‰ivitt‰‰ olemassaolevan TaskItem-olion tiedot pTasks-kokoelmassa.
Public Sub UpdateTask(ByVal updatedTask As clsTaskItem)
    Dim key As String
    On Error GoTo UpdateTaskError

    If pTasks Is Nothing Then Err.Raise vbObjectError + 517, "TaskManager.UpdateTask", "Teht‰v‰kokoelmaa ei ole alustettu."

    key = CStr(updatedTask.ID)

    ' Varmista, ett‰ p‰ivitett‰v‰ olio on olemassa
    Dim existing As clsTaskItem
    'On Error Resume Next
    Set existing = pTasks(key)
    On Error GoTo UpdateTaskError

    If existing Is Nothing Then
         Err.Raise vbObjectError + 518, "TaskManager.UpdateTask", "TaskItemia ID:ll‰ " & updatedTask.ID & " ei lˆytynyt p‰ivityst‰ varten."
         Exit Sub
    End If

    ' Korvaa vanha olio uudella (poista ensin, lis‰‰ sitten p‰ivitetty)
    pTasks.Remove key
    pTasks.Add updatedTask, key
    'Debug.Print "P‰ivitetty TaskItem ID:ll‰: " & updatedTask.ID

    Exit Sub
UpdateTaskError:
     MsgBox "Virhe TaskManager.UpdateTask: " & Err.Description, vbCritical, "P‰ivitysvirhe"
End Sub


' Poistaa TaskItem-olion pTasks-kokoelmasta ID:n perusteella.
Public Sub DeleteTask(ByVal taskID As Long)
    Dim key As String
    On Error GoTo DeleteTaskError

    If pTasks Is Nothing Then Exit Sub ' Ei mit‰‰n poistettavaa

    key = CStr(taskID)

    ' Varmista, ett‰ poistettava olio on olemassa
    Dim existing As clsTaskItem
    'On Error Resume Next
    Set existing = pTasks(key)
    On Error GoTo DeleteTaskError

    If existing Is Nothing Then
         'Debug.Print "DeleteTask: TaskItemia ID:ll‰ " & taskID & " ei lˆytynyt."
         ' Ei v‰ltt‰m‰tt‰ virhe, voi poistua hiljaa
         Exit Sub
    End If

    ' Poista olio kokoelmasta
    pTasks.Remove key
    'Debug.Print "Poistettu TaskItem ID:ll‰: " & taskID

    Exit Sub
DeleteTaskError:
    MsgBox "Virhe TaskManager.DeleteTask: " & Err.Description, vbCritical, "Poistovirhe"
End Sub


' Hakee tietyn TaskItem-olion pTasks-kokoelmasta ID:n perusteella.
' Palauttaa Nothing, jos oliota ei lˆydy.
Public Function GetTaskByID(ByVal taskID As Long) As clsTaskItem
    Dim key As String
    On Error Resume Next ' T‰rke‰, jos avainta ei lˆydy

    Set GetTaskByID = Nothing ' Oletuspalautusarvo

    If pTasks Is Nothing Then Exit Function ' Kokoelmaa ei alustettu

    key = CStr(taskID)
    Set GetTaskByID = pTasks(key) ' Yrit‰ hakea olio

    On Error GoTo 0 ' Palauta normaali virheenk‰sittely
End Function

' Luokan sis‰inen apufunktio Boolean-arvojen turvalliseen lukemiseen taulukosta
Private Function GetBooleanValue(value As Variant) As Boolean
    'On Error Resume Next ' Ohita mahdolliset virheet muunnoksessa
    Dim result As Boolean
    result = False ' Oletusarvo

    If Not IsEmpty(value) And Not IsNull(value) Then
        ' Kokeile yleisimpi‰ tapoja ilmaista True
        Select Case UCase(CStr(value))
            Case "TRUE", "-1", "1", "YES", "KYLLƒ", "K"
                result = True
            ' Muut tapaukset (esim. "FALSE", "0", "", teksti) tulkitaan Falseksi
        End Select
    End If

    GetBooleanValue = result
    'On Error GoTo 0 ' Palauta normaali virheenk‰sittely
End Function

