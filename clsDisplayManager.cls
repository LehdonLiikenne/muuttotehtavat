VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDisplayManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Class Module: DisplayManager ---
Option Explicit

' --- Vakiot ---
Private Const DISPLAY_COLUMN_COUNT As Long = 102 ' Näytettävät sarakkeet (A-AL)
Private Const HEADER_TEXT_COL As Long = 1       ' Sarake, johon "VIIKKO X" -teksti tulee (A)
Private Const STATUS_COLUMN_INDEX As Long = 33  ' Sarake, josta luetaan Tila ehdollista muotoilua varten (AG)
Private Const DATA_START_ROW As Long = 2        ' Rivin numero, josta data alkaa
Private Const DISPLAY_ID_COL As Long = 100 ' ID
Private Const DISPLAY_RECORD_TYPE_COL As Long = 101 ' Record type
Private Const DISPLAY_ATTENTION_DATE_COL As Long = 102 ' Attention sort date

' --- Päämetodi näytön päivittämiseen ---
Public Sub UpdateDisplay(ByVal tasks As Collection, ByVal targetSheetName As String)
    ' --- MUUTTUJIEN MÄÄRITTELYT ALUSSA ---
    Dim targetSheet As Worksheet
    Dim displayRowsColl As Collection
    Dim displayRowsArray() As clsDisplayRow
    Dim outputArray() As Variant
    Dim lastDataRowOnSheet As Long
    Dim outputRowIndex As Long
    Dim i As Long
    Dim debugMsg As String
    Dim taskItem As clsTaskItem
    Dim generatedRows As Collection
    Dim objToSort As clsDisplayRow
    Dim dispRow As clsDisplayRow
    Dim clearRange As Range
    Dim k As Long
    Dim j As Long
    Dim weekText As String
    Dim columnData As Variant
    Dim valueToWrite As Variant
    Dim previousWeekNum As Integer
    Dim previousYear As Integer
    Dim currentWeekNum As Integer
    Dim currentYear As Integer

    ' --- Kaikki Käytetyt Sarakkeet (määritä oikeat sarakkeet Excelissä) ---
    Const DISPLAY_ASIAKAS_COL As Long = 2 ' A
    Const DISPLAY_LASTAUSPAIVA_COL As Long = 3 ' C
    Const DISPLAY_LASTAUSMAA_COL As Long = 4 ' D
    Const DISPLAY_PURKUMAA_COL As Long = 5 ' E
    Const DISPLAY_PURKUPAIVA_COL As Long = 6 ' F
    Const DISPLAY_M3M_COL As Long = 7 ' G
    Const DISPLAY_PALVELU_COL As Long = 8 ' H
    Const DISPLAY_HUOMIOITAVAA_COL As Long = 9 ' I
    Const DISPLAY_PUHELIN_COL As Long = 10 ' J
    Const DISPLAY_LASTAUSOSOITE_COL As Long = 11 ' K
    Const DISPLAY_PURKUOSOITE_COL As Long = 12 ' L
    Const DISPLAY_KULJETTAJAT_COL As Long = 13 ' M
    Const DISPLAY_AUTOT_COL As Long = 14 ' N
    Const DISPLAY_KONTIT_COL As Long = 15 ' O
    Const DISPLAY_APULAISET_COL As Long = 16 ' P
    Const DISPLAY_RAHTIKIRJA_COL As Long = 17 ' Q
    Const DISPLAY_APULAISET_TILATTU_COL As Long = 18 ' R
    Const DISPLAY_PYSAKOINTILUPA_COL As Long = 19 ' S
    Const DISPLAY_HISSIA_COL As Long = 20 ' T
    Const DISPLAY_LAIVALIPPU_COL As Long = 21 ' U
    Const DISPLAY_LASKUTUS_COL As Long = 22 ' V
    Const DISPLAY_VAKUUTUS_COL As Long = 23 ' W
    Const DISPLAY_ARVO_COL As Long = 24 ' X
    Const DISPLAY_HINTA_COL As Long = 25 ' Y
    Const DISPLAY_MUUTTOMAAILMA_COL As Long = 26 ' Z
    Const DISPLAY_M3T_COL As Long = 27 ' AA
    Const DISPLAY_LASTAUSPAIVAVARMISTUNUT_COL As Long = 28 ' AB
    Const DISPLAY_PURKUPAIVAVARMISTUNUT_COL As Long = 29 ' AC
    Const DISPLAY_TARJOUSHYVAKSYTTY_COL As Long = 31 ' AE
    Const DISPLAY_TARJOUSHYLATTY_COL As Long = 32 ' AF
    Const DISPLAY_TILA_COL As Long = 33 ' AG
    Const DISPLAY_LASTAUSLOPPUU_COL As Long = 34 ' AH
    Const DISPLAY_PURKULOPPUU_COL As Long = 35 ' AI
    Const DISPLAY_VALIMATKA_COL As Long = 36
    Const DISPLAY_ID_COL As Long = 100
    Const DISPLAY_RECORD_TYPE_COL As Long = 101
    Const DISPLAY_ATTENTION_DATE_COL As Long = 102
    
    'Sähköpostin teksti pienellä
    Const SARAKE_PIENILLA_KIRJAIMILLA_INDEKSI As Long = 20


    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "Päivitetään näyttöä..."

    ' --- Vaihe 0: Alustukset ja Kohdevälilehden Haku ---
    Set displayRowsColl = New Collection

    Set targetSheet = Nothing
    On Error GoTo 0
    Set targetSheet = ThisWorkbook.Worksheets(targetSheetName)
    On Error GoTo ErrorHandler
    If targetSheet Is Nothing Then
        Err.Raise vbObjectError + 514, "DisplayManager.UpdateDisplay", "Kohdevälilehteä '" & targetSheetName & "' ei löytynyt."
        GoTo CleanUp
    End If

    ' Tyhjennä vanha sisältö ja muotoilut
    lastDataRowOnSheet = mdlSheetUtils.GetLastRow(targetSheet, 1)
    If lastDataRowOnSheet >= DATA_START_ROW Then
        Set clearRange = targetSheet.Range(targetSheet.Cells(DATA_START_ROW, 1), targetSheet.Cells(lastDataRowOnSheet, DISPLAY_COLUMN_COUNT))
        mdlSheetUtils.ClearRangeContents clearRange
        mdlSheetUtils.ClearRangeFormats clearRange
        On Error Resume Next
        clearRange.EntireRow.RowHeight = targetSheet.StandardHeight
        On Error GoTo ErrorHandler
        Set clearRange = Nothing
    End If

    ' --- Vaihe 1: Generoi DisplayRow-oliot kaikista TaskItemeista ---
    Application.StatusBar = "Generoidaan näyttörivit..."
    If tasks Is Nothing Or tasks.Count = 0 Then
        targetSheet.Cells(DATA_START_ROW, 1).value = "Ei näytettäviä tietoja."
        GoTo CleanUp
    End If

    For Each taskItem In tasks
        Set generatedRows = GenerateDisplayRowsForItem(taskItem)
        If Not generatedRows Is Nothing Then
            Dim rowToAdd As clsDisplayRow
            For Each rowToAdd In generatedRows
                If Not rowToAdd Is Nothing Then
                     On Error Resume Next
                     displayRowsColl.Add rowToAdd
                     On Error GoTo ErrorHandler
                Else
                End If
            Next rowToAdd
        End If
        Set generatedRows = Nothing
    Next taskItem

    ' --- Vaihe 2: Käsittele tyhjä tulos ---
    If displayRowsColl.Count = 0 Then
        targetSheet.Cells(DATA_START_ROW, 1).value = "Ei generoitunut näyttörivejä."
        GoTo CleanUp
    End If

    ' --- Vaihe 3: Lajittele DisplayRow-oliot ---
    Application.StatusBar = "Lajitellaan rivejä..."
    ReDim displayRowsArray(1 To displayRowsColl.Count)
    i = 1
    For Each objToSort In displayRowsColl
        Set displayRowsArray(i) = objToSort
        i = i + 1
    Next objToSort
    Set displayRowsColl = Nothing

    ' Tarkista SortDate ENNEN lajittelua
    For i = LBound(displayRowsArray) To UBound(displayRowsArray)
        Set dispRow = displayRowsArray(i)
        If dispRow Is Nothing Then
            Debug.Print "!!! UpdateDisplay VAROITUS: displayRowsArray(" & i & ") on Nothing. Ohitetaan rivi."
        Else
            If IsNull(dispRow.SortDate) Then
                'Debug.Print " !!! PRE-SORT WARNING: displayRowsArray(" & i & ").SortDate is Null! Task ID: " & dispRow.SourceRecordID
            ElseIf Not IsDate(dispRow.SortDate) Then
                 'Debug.Print " !!! PRE-SORT WARNING: displayRowsArray(" & i & ").SortDate is not a valid Date! Task ID: " & dispRow.SourceRecordID & ", Value: " & CStr(dispRow.SortDate)
            End If
        End If
    Next i
    Set dispRow = Nothing

    ' Lajittele taulukko
    On Error Resume Next
    If UBound(displayRowsArray) >= LBound(displayRowsArray) Then
         mdlSortUtils.QuickSortDisplayRowsBySortDate displayRowsArray, LBound(displayRowsArray), UBound(displayRowsArray)
    End If
    On Error GoTo ErrorHandler

    ' --- Vaihe 4: Rakenna Lopullinen Tulostustaulukko (outputArray) & Lisää Viikko-otsikot ---
    Application.StatusBar = "Rakennetaan tulostetta..."
    ReDim outputArray(1 To UBound(displayRowsArray) * 2, 1 To DISPLAY_COLUMN_COUNT)
    outputRowIndex = 0

    previousWeekNum = -1
    previousYear = -1

    For i = LBound(displayRowsArray) To UBound(displayRowsArray)
        Set dispRow = displayRowsArray(i)

        If dispRow Is Nothing Then
            GoTo NextIterationInLoop
        End If
        If IsNull(dispRow.SortDate) Or Not IsDate(dispRow.SortDate) Then
             GoTo NextIterationInLoop
        End If

        ' Hae viikko ja vuosi
        currentWeekNum = mdlDateUtils.GetWeekNumberISO8601(dispRow.SortDate)
        currentYear = Year(dispRow.SortDate)

        ' Lisää Viikko-otsikko
        If (currentWeekNum <> previousWeekNum Or currentYear <> previousYear) Then
            outputRowIndex = outputRowIndex + 1
            Dim currentDate As Date
            Dim monthName As String
            Dim weekStartDate As Date
            Dim weekEndDate As Date
            Dim startDateStr As String
            Dim endDateStr As String

            ' Hae tarvittavat tiedot käyttäen apufunktioita
            currentDate = dispRow.SortDate
            monthName = mdlDateUtils.GetFinnishMonthName(currentDate)
            weekStartDate = mdlDateUtils.GetFirstDayOfWeekISO(currentYear, currentWeekNum)
            weekEndDate = mdlDateUtils.GetLastDayOfWeekISO(currentYear, currentWeekNum)

            startDateStr = Format$(weekStartDate, "dd.mm.yyyy")
            endDateStr = Format$(weekEndDate, "dd.mm.yyyy")

            weekText = "VIIKKO " & currentWeekNum & ", " & monthName & " (" & startDateStr & " - " & endDateStr & ")"

            outputArray(outputRowIndex, 1) = weekText
            outputArray(outputRowIndex, 20) = weekText

            For k = 21 To DISPLAY_COLUMN_COUNT
                 outputArray(outputRowIndex, k) = vbNullString
            Next k
        End If
        previousWeekNum = currentWeekNum
        previousYear = currentYear

        ' Lisää DataRivi
        outputRowIndex = outputRowIndex + 1
        columnData = dispRow.DisplayColumns

        If IsArray(columnData) Then
            For j = 1 To DISPLAY_COLUMN_COUNT
                valueToWrite = columnData(j)
                If IsNull(valueToWrite) Then valueToWrite = ""

                If VarType(valueToWrite) = vbString Then
                    If j = SARAKE_PIENILLA_KIRJAIMILLA_INDEKSI Then
                        outputArray(outputRowIndex, j) = LCase$(valueToWrite)
                    Else
                        outputArray(outputRowIndex, j) = UCase$(valueToWrite)
                    End If
                Else
                    outputArray(outputRowIndex, j) = valueToWrite
                End If
            Next j
        Else
            Debug.Print "    !!! VIRHE: dispRow.DisplayColumns ei palauttanut taulukkoa Task ID: " & dispRow.SourceRecordID
        End If

        Set dispRow = Nothing

NextIterationInLoop:
    Next i

    ' --- Vaihe 5: Kirjoita Koko Tulostaulukko Kerralla Exceliin ---
    Application.StatusBar = "Kirjoitetaan tietoja taulukkoon..."
    On Error Resume Next
    targetSheet.Range("G:G").NumberFormat = "@"
    targetSheet.Range("J:J").NumberFormat = "@"
    If Err.Number <> 0 Then Debug.Print "UpdateDisplay: Virhe asetettaessa Text-muotoilua sarakkeelle G: " & Err.Description: Err.Clear
    On Error GoTo ErrorHandler

    If outputRowIndex > 0 Then
        targetSheet.Cells(DATA_START_ROW, 1).Resize(outputRowIndex, DISPLAY_COLUMN_COUNT).value = outputArray
    End If

    ' --- Vaihe 6: Muotoilut ---
    Application.StatusBar = "Muotoillaan rivejä..."
    If outputRowIndex > 0 Then
        ApplyFormatting targetSheet, outputRowIndex + DATA_START_ROW - 1
    End If

    ' --- Piilota apusarakkeet ---
    If Not targetSheet Is Nothing Then
        On Error Resume Next
        'targetSheet.Range("AG:AK").EntireColumn.Hidden = True ' Poista tämä rivi tai muuta tarvittaessa
        On Error GoTo ErrorHandler
    End If
    
    ' --- Vaihe 6b: Piilota virheentarkistuskolmiot sarakkeilta G (M3m) ja J (Puhelin) ---
    If outputRowIndex > 0 Then
        Dim targetRange As Range
        Dim dataEndRow As Long
        Dim cell As Range
        Dim targetColumn As Variant

        dataEndRow = outputRowIndex + DATA_START_ROW - 1

        For Each targetColumn In Array("G", "J")
            Set targetRange = Nothing

            On Error Resume Next
            Set targetRange = targetSheet.Range(targetColumn & DATA_START_ROW & ":" & targetColumn & dataEndRow)

            If Not targetRange Is Nothing Then
                For Each cell In targetRange.Cells
                    If cell.Errors.Item(xlNumberAsText).value = True Then
                         cell.Errors.Item(xlNumberAsText).Ignore = True
                    End If
                Next cell
            Else
                Debug.Print "UpdateDisplay: Ei voitu määrittää aluetta sarakkeelle " & targetColumn & " virheiden ohitusta varten."
            End If
            On Error GoTo ErrorHandler

        Next targetColumn

        Set targetRange = Nothing
        Set cell = Nothing
    End If

    ' --- Vaihe 7: Siivous ---
CleanUp:
    Application.StatusBar = False
    Set targetSheet = Nothing
    Set displayRowsColl = Nothing
    Set taskItem = Nothing
    Set generatedRows = Nothing
    Set objToSort = Nothing
    Set dispRow = Nothing
    If IsArray(outputArray) Then Erase outputArray
    If IsArray(displayRowsArray) Then Erase displayRowsArray
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    debugMsg = "Virhe DisplayManager.UpdateDisplay: " & Err.Description & " (Virhe " & Err.Number & ")"
    Debug.Print "!!! " & debugMsg
    MsgBox debugMsg, vbCritical, "Virhe näytön päivityksessä"
    Application.ScreenUpdating = True
    Application.StatusBar = False
    GoTo CleanUp
End Sub


' Generoi tarvittavat DisplayRow-oliot (TaskType, AttentionType) yhdestä TaskItem-oliosta.
Private Function GenerateDisplayRowsForItem(ByVal taskItem As clsTaskItem) As Collection


    Dim generatedRows As New Collection
    Dim displayRow As clsDisplayRow
    Dim loopDate As Date
    Dim tempDateVariant As Variant

    ' Selkeämmät Date-muuttujat validoinnin jälkeen
    Dim validLoadStartDate As Date, validLoadEndDate As Date
    Dim validUnloadStartDate As Date, validUnloadEndDate As Date
    Dim hasValidUnloadStart As Boolean, hasValidUnloadEnd As Boolean
    Dim existingRow As clsDisplayRow

    Const RIVI_ALOITUS As Long = 1000 ' Käytetään vain Debug.Printissä

    On Error GoTo GenerationError_Handler
    
    If taskItem Is Nothing Then
        Debug.Print "GenerateDisplayRowsForItem: Sai Nothing taskItem -olion. Poistutaan."
        GoTo CleanExit_GenerateRows
    End If

    ' --- Käsittele Task-tyyppi ---
    If taskItem.RecordType = "Task" Then
        Dim isOneDayMove As Boolean

        tempDateVariant = mdlDateUtils.ConvertToDate(taskItem.lastausPaiva)
        If IsNull(tempDateVariant) Or Not IsDate(tempDateVariant) Then
             Debug.Print "--> Task ID: " & taskItem.ID & " (Rivi " & RIVI_ALOITUS + 1 & "), Virheellinen tai puuttuva Lastauspäivä ('" & IIf(IsNull(taskItem.lastausPaiva), "NULL_D", CStr(taskItem.lastausPaiva)) & "'). Ohitetaan Task-rivien generointi."
             GoTo CleanExit_GenerateRows
        End If
        On Error Resume Next
        validLoadStartDate = CDate(tempDateVariant)
        If Err.Number <> 0 Then
            Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 2 & ": CDate(lastausPaiva='" & IIf(IsNull(taskItem.lastausPaiva), "NULL_D", CStr(taskItem.lastausPaiva)) & "') epäonnistui - " & Err.Description
            Err.Clear
            GoTo CleanExit_GenerateRows
        End If
        On Error GoTo GenerationError_Handler

        tempDateVariant = mdlDateUtils.ConvertToDate(taskItem.LastausLoppuu)
        Dim tempLoadEndDate As Date
        Dim useDefaultLoadEndDate As Boolean: useDefaultLoadEndDate = True
        If Not IsNull(tempDateVariant) And IsDate(tempDateVariant) Then
            On Error Resume Next
            tempLoadEndDate = CDate(tempDateVariant)
            If Err.Number = 0 Then
                On Error GoTo GenerationError_Handler
                If tempLoadEndDate >= validLoadStartDate Then
                    validLoadEndDate = tempLoadEndDate
                    useDefaultLoadEndDate = False
                End If
            Else
                Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 3 & ": CDate(LastausLoppuu='" & IIf(IsNull(taskItem.LastausLoppuu), "NULL_D", CStr(taskItem.LastausLoppuu)) & "') epäonnistui - " & Err.Description
                Err.Clear
            End If
            On Error GoTo GenerationError_Handler
        End If
        If useDefaultLoadEndDate Then
            If IsDate(validLoadStartDate) Then ' Varmistus, ettei validLoadStartDate ole alustamaton
                 validLoadEndDate = validLoadStartDate
            Else
                 Debug.Print "--> Task ID: " & taskItem.ID & ": Ei voitu asettaa validLoadEndDate, koska validLoadStartDate ei ole validi Date."
                 GoTo CleanExit_GenerateRows ' Tai muu virheenkäsittely
            End If
        End If

        tempDateVariant = mdlDateUtils.ConvertToDate(taskItem.purkuPaiva)
        hasValidUnloadStart = False
        If Not IsNull(tempDateVariant) And IsDate(tempDateVariant) Then
            On Error Resume Next
            validUnloadStartDate = CDate(tempDateVariant)
            If Err.Number = 0 Then
                hasValidUnloadStart = True
            Else
                Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 5 & ": CDate(purkuPaiva='" & IIf(IsNull(taskItem.purkuPaiva), "NULL_D", CStr(taskItem.purkuPaiva)) & "') epäonnistui - " & Err.Description
                Err.Clear
            End If
            On Error GoTo GenerationError_Handler
        End If

        hasValidUnloadEnd = False
        If hasValidUnloadStart Then
            tempDateVariant = mdlDateUtils.ConvertToDate(taskItem.PurkuLoppuu)
            Dim tempUnloadEndDate As Date
            Dim useDefaultUnloadEndDate As Boolean: useDefaultUnloadEndDate = True
            If Not IsNull(tempDateVariant) And IsDate(tempDateVariant) Then
                On Error Resume Next
                tempUnloadEndDate = CDate(tempDateVariant)
                If Err.Number = 0 Then
                    On Error GoTo GenerationError_Handler
                    If tempUnloadEndDate >= validUnloadStartDate Then
                        validUnloadEndDate = tempUnloadEndDate
                        useDefaultUnloadEndDate = False
                        hasValidUnloadEnd = True
                    End If
                Else
                    Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 6 & ": CDate(PurkuLoppuu='" & IIf(IsNull(taskItem.PurkuLoppuu), "NULL_D", CStr(taskItem.PurkuLoppuu)) & "') epäonnistui - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo GenerationError_Handler
            End If
            If useDefaultUnloadEndDate Then
                If IsDate(validUnloadStartDate) Then ' Varmistus
                    validUnloadEndDate = validUnloadStartDate
                    hasValidUnloadEnd = True
                Else
                    Debug.Print "--> Task ID: " & taskItem.ID & ": Ei voitu asettaa validUnloadEndDate, koska validUnloadStartDate ei ole validi Date."
                    ' Ei välttämättä GoTo CleanExit, riippuu logiikasta
                End If
            End If
        End If
        '--- YHDEN PÄIVÄN MUUTTO ---
        isOneDayMove = False
        If Not hasValidUnloadStart Then
            isOneDayMove = True
        Else
            If IsDate(validLoadStartDate) And IsDate(validUnloadStartDate) Then
                On Error Resume Next
                Dim formatCheck1 As String, formatCheck2 As String
                formatCheck1 = Format$(validLoadStartDate, "YYYYMMDD")
                formatCheck2 = Format$(validUnloadStartDate, "YYYYMMDD")
                If Err.Number <> 0 Then
                    Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 8 & ": Format$ epäonnistui yhden päivän muuton tarkistuksessa (validLoadStartDate='" & IIf(IsNull(validLoadStartDate), "NULL_D", CStr(validLoadStartDate)) & "', validUnloadStartDate='" & IIf(IsNull(validUnloadStartDate), "NULL_D", CStr(validUnloadStartDate)) & "') - " & Err.Description
                    Err.Clear
                    isOneDayMove = True
                Else
                    If formatCheck1 = formatCheck2 Then isOneDayMove = True
                End If
                On Error GoTo GenerationError_Handler
            Else
                Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 8 & ": validLoadStartDate tai validUnloadStartDate ei ole validi Date Format$-kutsulle isOneDayMove-tarkistuksessa."
                isOneDayMove = True
            End If
        End If

        If IsDate(validLoadStartDate) And IsDate(validLoadEndDate) Then
            On Error Resume Next
            For loopDate = validLoadStartDate To validLoadEndDate
                If Err.Number <> 0 Then
                    Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 9 & ": Lastausrivi-silmukan aloitus epäonnistui (validLoadStartDate='" & IIf(IsNull(validLoadStartDate), "NULL_D", CStr(validLoadStartDate)) & "', validLoadEndDate='" & IIf(IsNull(validLoadEndDate), "NULL_D", CStr(validLoadEndDate)) & "') - " & Err.Description
                    Err.Clear
                    Exit For
                End If
                
                Set displayRow = New clsDisplayRow
                displayRow.RowType = TaskType
                displayRow.SortDate = loopDate
                displayRow.SourceRecordID = taskItem.ID
                displayRow.IsLoadingDay = True
                
                PopulateTaskDisplayColumns taskItem, displayRow
                
                generatedRows.Add displayRow
            Next loopDate
            On Error GoTo GenerationError_Handler
        Else
            Debug.Print "--> Task ID: " & taskItem.ID & ": Ei voitu generoida lastauspäiviä, koska validLoadStartDate (" & TypeName(validLoadStartDate) & ": " & CStr(IIf(IsDate(validLoadStartDate), validLoadStartDate, "NotDate")) & ") tai validLoadEndDate (" & TypeName(validLoadEndDate) & ": " & CStr(IIf(IsDate(validLoadEndDate), validLoadEndDate, "NotDate")) & ") ei ole validi Date."
        End If
        
        If Not isOneDayMove And hasValidUnloadStart And hasValidUnloadEnd Then
            If IsDate(validUnloadStartDate) And IsDate(validUnloadEndDate) Then
                Dim dateAlreadyAdded As Boolean
                On Error Resume Next
                For loopDate = validUnloadStartDate To validUnloadEndDate
                    If Err.Number <> 0 Then
                        Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 10 & ": Purkurivi-silmukan aloitus epäonnistui (validUnloadStartDate='" & IIf(IsNull(validUnloadStartDate), "NULL_D", CStr(validUnloadStartDate)) & "', validUnloadEndDate='" & IIf(IsNull(validUnloadEndDate), "NULL_D", CStr(validUnloadEndDate)) & "') - " & Err.Description
                        Err.Clear
                        Exit For
                    End If

                    Dim loopDateStrForFormat As String, existingDateStrForFormat As String
                    On Error Resume Next ' Suojaa Format$
                    loopDateStrForFormat = Format$(loopDate, "YYYYMMDD")
                    If Err.Number <> 0 Then loopDateStrForFormat = "ERR_LOOP_FMT": Err.Clear
                    On Error GoTo GenerationError_Handler

                    dateAlreadyAdded = False
                    For Each existingRow In generatedRows
                        On Error Resume Next ' Suojaa Format$
                        existingDateStrForFormat = Format$(existingRow.SortDate, "YYYYMMDD")
                        If Err.Number <> 0 Then existingDateStrForFormat = "ERR_EXIST_FMT": Err.Clear
                        On Error GoTo GenerationError_Handler

                        If existingDateStrForFormat = loopDateStrForFormat And loopDateStrForFormat <> "ERR_LOOP_FMT" And existingDateStrForFormat <> "ERR_EXIST_FMT" Then
                            dateAlreadyAdded = True
                            Exit For
                        End If
                    Next existingRow
                    
                    If Not dateAlreadyAdded Then
                        Set displayRow = New clsDisplayRow
                        displayRow.RowType = TaskType
                        displayRow.SortDate = loopDate
                        displayRow.SourceRecordID = taskItem.ID
                        displayRow.IsLoadingDay = False
                        PopulateTaskDisplayColumns taskItem, displayRow
                        generatedRows.Add displayRow
                    End If
                Next loopDate
                On Error GoTo GenerationError_Handler
            Else
                 Debug.Print "--> Task ID: " & taskItem.ID & ": Ei voitu generoida purkupäiviä, koska validUnloadStartDate (" & TypeName(validUnloadStartDate) & ": " & CStr(IIf(IsDate(validUnloadStartDate), validUnloadStartDate, "NotDate")) & ") tai validUnloadEndDate (" & TypeName(validUnloadEndDate) & ": " & CStr(IIf(IsDate(validUnloadEndDate), validUnloadEndDate, "NotDate")) & ") ei ole validi Date."
            End If
        End If

    ' --- Käsittele Attention-tyyppi ---
    ElseIf taskItem.RecordType = "Attention" Then
        tempDateVariant = mdlDateUtils.ConvertToDate(taskItem.AttentionSortDate)
        Dim validAttentionDate As Date
        If Not IsNull(tempDateVariant) And IsDate(tempDateVariant) Then
            On Error Resume Next
            validAttentionDate = CDate(tempDateVariant)
            If Err.Number <> 0 Then
                Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE RIVI " & RIVI_ALOITUS + 12 & ": CDate(AttentionSortDate='" & IIf(IsNull(taskItem.AttentionSortDate), "NULL_D", CStr(taskItem.AttentionSortDate)) & "') epäonnistui - " & Err.Description
                Err.Clear
                GoTo SkipAttention_GenerateRows
            End If
            On Error GoTo GenerationError_Handler

            Set displayRow = New clsDisplayRow
            displayRow.RowType = AttentionType
            displayRow.SortDate = validAttentionDate
            displayRow.SourceRecordID = taskItem.ID
            
            PopulateAttentionDisplayColumns taskItem, displayRow
            
            generatedRows.Add displayRow
        Else
SkipAttention_GenerateRows:
            Debug.Print "--> Task ID: " & taskItem.ID & " (Rivi " & RIVI_ALOITUS + 13 & "), Ohitetaan Attention-rivi, koska AttentionSortDate puuttuu tai on virheellinen: '" & IIf(IsNull(taskItem.AttentionSortDate), "NULL_D", CStr(taskItem.AttentionSortDate)) & "'"
        End If

    ' --- Käsittele Kontakti-tyyppi ---
    ElseIf taskItem.RecordType = "Kontakti" Then
        tempDateVariant = mdlDateUtils.ConvertToDate(taskItem.kontaktiPaiva)
        Dim validKontaktDate As Date
        If Not IsNull(tempDateVariant) And IsDate(tempDateVariant) Then
            On Error Resume Next
            validKontaktDate = CDate(tempDateVariant)
            If Err.Number <> 0 Then
                Debug.Print "!!! Task ID: " & taskItem.ID & " VIRHE: CDate(kontaktiPaiva='" & IIf(IsNull(taskItem.kontaktiPaiva), "NULL_D", CStr(taskItem.kontaktiPaiva)) & "') epäonnistui - " & Err.Description
                Err.Clear
                GoTo SkipKontakt_GenerateRows
            End If
            On Error GoTo GenerationError_Handler

            Set displayRow = New clsDisplayRow
            displayRow.RowType = KontaktType
            displayRow.SortDate = validKontaktDate
            displayRow.SourceRecordID = taskItem.ID
            
            PopulateKontaktDisplayColumns taskItem, displayRow
            
            generatedRows.Add displayRow
        Else
SkipKontakt_GenerateRows:
            Debug.Print "--> Task ID: " & taskItem.ID & ", Ohitetaan Kontakti-rivi, koska kontaktiPaiva puuttuu tai on virheellinen: '" & IIf(IsNull(taskItem.kontaktiPaiva), "NULL_D", CStr(taskItem.kontaktiPaiva)) & "'"
        End If
        
    ' --- Tuntematon tyyppi ---
    Else
        Debug.Print "--> Task ID: " & taskItem.ID & " (Rivi " & RIVI_ALOITUS + 14 & "), Tuntematon RecordType: '" & taskItem.RecordType & "'. Ohitetaan."
    End If

CleanExit_GenerateRows:
    Set GenerateDisplayRowsForItem = generatedRows
    Set displayRow = Nothing
    Set existingRow = Nothing
    Exit Function

GenerationError_Handler:
    Dim taskIDStr As String: Dim errNum As Long: Dim errDesc As String
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    taskIDStr = CStr(taskItem.ID)
    If Err.Number <> 0 Then taskIDStr = "[Tuntematon ID GenerateDisplayRowsForItem-virheessä]"
    Err.Clear ' Nollaa tämä sisempi virhe ennen kuin yritetään tulostaa errDesc
    On Error GoTo 0
    
    Dim descForPrint As String
    If IsNull(errDesc) Then
        descForPrint = "[Virheen kuvaus on Null]"
    ElseIf IsEmpty(errDesc) Then
        descForPrint = "[Virheen kuvaus on Empty]"
    Else
        descForPrint = errDesc
    End If
    Debug.Print "!!! VIRHE GenerateDisplayRowsForItem: Task ID " & taskIDStr & ", Virhe " & errNum & ": " & descForPrint
    
    Set generatedRows = New Collection
    Set GenerateDisplayRowsForItem = generatedRows
    Exit Function

End Function
' Täyttää DisplayRow.DisplayColumns -taulukon normaaleille tehtäväriveille (TaskType)
Private Sub PopulateTaskDisplayColumns(ByVal sourceTask As clsTaskItem, ByRef targetRow As clsDisplayRow)
    Dim tempStr As String
    Const DEFAULT_LIST_STR As String = "-"
    Dim tempDateVariant As Variant
    Dim valueToSet As Variant
    Dim currentCol As Long

    On Error GoTo PopulateColumnError

    ' Sarake 1 (A): Viikonpäivä (Tämä käyttää edelleen merkkijonoa)
    currentCol = 1
    If IsDate(targetRow.SortDate) Then valueToSet = mdlStringUtils.GetFinnishWeekdayName(targetRow.SortDate) Else valueToSet = "?"
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 2 (B): Asiakas
    currentCol = 2
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.asiakas)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 3 (C): Lastauspäivä
    currentCol = 3
    tempDateVariant = IIf(targetRow.IsLoadingDay, targetRow.SortDate, sourceTask.lastausPaiva)
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null ' Tai Empty
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 4 (D): Lastausmaa
    currentCol = 4
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.lastausMaa)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 5 (E): Purkumaa
    currentCol = 5
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.purkuMaa)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 6 (F): Purkupäivä
    currentCol = 6
    tempDateVariant = IIf(targetRow.IsLoadingDay, sourceTask.purkuPaiva, targetRow.SortDate)
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null ' Tai Empty
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 7 (G): M3m
    currentCol = 7
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.M3m)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 8 (H): Palvelu
    currentCol = 8
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.palvelu)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 9 (I): Huomioitavaa
    currentCol = 9
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Huomioitavaa)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 10 (M): Kuljettajat
    currentCol = 10
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Kuljettajat, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 11 (N): Autot
    currentCol = 11
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Autot, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 12 (O): Kontit
    currentCol = 12
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Kontit, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 13 (P): Apulaiset
    currentCol = 13
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Apulaiset, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 14 (Q): Rahtikirja
    currentCol = 14
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(sourceTask.Rahtikirja, "EI"))): Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK": valueToSet = "KYLLÄ": Case Else: valueToSet = "EI": End Select
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 15 (R): ApulaisetTilattu
    currentCol = 15
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.ApulaisetTilattu)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 16 (S): Pysakointilupa
    currentCol = 16
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Pysakointilupa)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 17 (T): Hissi
    currentCol = 17
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.hissi)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 18 (U): Laivalippu
    currentCol = 18
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Laivalippu)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 19 (J): Puhelin
    currentCol = 19
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.puhelin)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 20 (AK): Sahköposti
    currentCol = 20
    valueToSet = sourceTask.sahkoposti
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 21 (K): Lastausosoite
    currentCol = 21
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.lastausOsoite)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 22 (L): Purkuosoite
    currentCol = 22
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.purkuOsoite)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 23 (W): Vakuutus
    currentCol = 23
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Vakuutus)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 24 (X): Arvo
    currentCol = 24
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Arvo, "-")
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 25 (Y): hinta
    currentCol = 25
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.hinta, "-")
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 26 (Z): Muuttomaailma
    currentCol = 26
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(sourceTask.Muuttomaailma, "EI"))): Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK": valueToSet = "KYLLÄ": Case Else: valueToSet = "EI": End Select
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 27 (AA): M3t
    currentCol = 27
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.M3t)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 28 (AB): LastauspaivaVarmistunut
    currentCol = 28
    If sourceTask.LastauspaivaVarmistunut Then valueToSet = "KYLLÄ" Else valueToSet = "EI"
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 29 (AC): PurkupaivaVarmistunut
    currentCol = 29
    If sourceTask.PurkupaivaVarmistunut Then valueToSet = "KYLLÄ" Else valueToSet = "EI"
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 30 (AD): TarjousTehty
    currentCol = 30
    tempDateVariant = sourceTask.tarjousTehty
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 31 (AE): TarjousHyväksytty
    currentCol = 31
    tempDateVariant = sourceTask.TarjousHyvaksytty
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 32 (AF): TarjousHylatty
    currentCol = 32
    tempDateVariant = sourceTask.TarjousHylatty
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 33 (AG): Tila
    currentCol = 33
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Tila)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 34 (AH): LastausLoppuu
    currentCol = 34
    tempDateVariant = sourceTask.LastausLoppuu
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 35 (AI): PurkuLoppuu
    currentCol = 35
    tempDateVariant = sourceTask.PurkuLoppuu
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 38 (V): Laskutus
    currentCol = 38
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(sourceTask.Laskutus, "EI"))): Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK": valueToSet = "KYLLÄ": Case Else: valueToSet = "EI": End Select
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 100 (AJ): ID
    currentCol = 100
    valueToSet = sourceTask.ID
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 101 (AJ): ID
    currentCol = 101
    valueToSet = sourceTask.RecordType
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 102 (AK): SortDate
    currentCol = 102
    If IsDate(targetRow.SortDate) Then valueToSet = targetRow.SortDate Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Kaikki sarakkeet asetettu tälle riville
    On Error GoTo 0
    Exit Sub

PopulateColumnError:
    Debug.Print "!!! ERROR in PopulateTaskDisplayColumns for Task ID: " & sourceTask.ID & " while processing Column " & currentCol & " - Error " & Err.Number & ": " & Err.Description
    ' Yritä asettaa tyhjä arvo ja jatkaa
    targetRow.SetColumnValue currentCol, Null ' Tai Empty
    Resume Next
End Sub

Private Sub PopulateKontaktDisplayColumns(ByVal sourceTask As clsTaskItem, ByRef targetRow As clsDisplayRow)
    Dim combinedString As String
    Dim weekdayStr As String
    Dim displayDateStr As String
    Dim huomioStr As String
    Dim i As Long
    Dim validKontaktDate As Date
    Dim valueToSet As Variant
    Dim currentCol As Long
    Dim tempStr As String
    Dim tempDateVariant As Variant
    Const DEFAULT_LIST_STR As String = "-"

    ' Varmista, että SortDate (joka on kontaktiPaiva tässä tapauksessa) on validi
    If Not IsDate(targetRow.SortDate) Then
        Debug.Print "PopulateKontaktDisplayColumns: Invalid SortDate for Task ID: " & sourceTask.ID
        targetRow.SetColumnValue 1, "VIRHE PVM (KONTAKTI)"
        targetRow.SetColumnValue STATUS_COLUMN_INDEX, "KONTAKTI" ' AG, 33
        targetRow.SetColumnValue DISPLAY_ID_COL, sourceTask.ID ' 100
        targetRow.SetColumnValue DISPLAY_RECORD_TYPE_COL, "Kontakti" ' 101
        targetRow.SetColumnValue DISPLAY_ATTENTION_DATE_COL, Null ' 102 (Koska SortDate oli virheellinen)
        Exit Sub
    End If
    validKontaktDate = targetRow.SortDate

    ' --- Sarakkeiden täyttö "Kontakti"-riville ---

    ' Sarake 1 (A): Viikonpäivä, Päivämäärä ja tunniste "(KONTAKTI)"
    currentCol = 1
    weekdayStr = mdlStringUtils.GetFinnishWeekdayName(validKontaktDate)
    'displayDateStr = Format$(validKontaktDate, "dd.mm.yyyy")
    combinedString = UCase$(weekdayStr)
    targetRow.SetColumnValue currentCol, combinedString

    ' Sarake 2 (B): Asiakas
    currentCol = 2
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.asiakas)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 3 (C): Lastauspäivä (Jätetään tyhjäksi tai näytetään "-", koska kontaktiPaiva on pääpäivämäärä)
    currentCol = 3
    targetRow.SetColumnValue currentCol, "*" & validKontaktDate ' Tai vbNullString

    ' Sarake 4 (D): Lastausmaa (Voidaan näyttää, jos syötetty lomakkeella)
    currentCol = 4
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.lastausMaa)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 5 (E): Purkumaa (Voidaan näyttää, jos syötetty lomakkeella)
    currentCol = 5
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.purkuMaa)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 6 (F): Purkupäivä (Voidaan näyttää, jos syötetty lomakkeella, tai "-")
    currentCol = 6
    If IsDate(sourceTask.purkuPaiva) Then
        valueToSet = sourceTask.purkuPaiva
    Else
        valueToSet = "-" ' Tai vbNullString
    End If
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 7 (G): M3m (Jätetään tyhjäksi tai "-")
    currentCol = 7
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.M3m)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 8 (H): Palvelu (Voidaan näyttää, jos syötetty)
    currentCol = 8
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.palvelu)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 9 (I): Huomioitavaa (Tärkeä kontaktille)
    currentCol = 9
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Huomioitavaa)
    targetRow.SetColumnValue currentCol, valueToSet

  ' Sarake 10 (M): Kuljettajat
    currentCol = 10
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Kuljettajat, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 11 (N): Autot
    currentCol = 11
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Autot, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 12 (O): Kontit
    currentCol = 12
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Kontit, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 13 (P): Apulaiset
    currentCol = 13
    tempStr = mdlStringUtils.DefaultIfNull(sourceTask.Apulaiset, DEFAULT_LIST_STR): If tempStr <> DEFAULT_LIST_STR Then tempStr = Replace(tempStr, ";", ", ")
    valueToSet = tempStr
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 14 (Q): Rahtikirja
    currentCol = 14
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(sourceTask.Rahtikirja, "EI"))): Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK": valueToSet = "KYLLÄ": Case Else: valueToSet = "EI": End Select
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 15 (R): ApulaisetTilattu
    currentCol = 15
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.ApulaisetTilattu)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 16 (S): Pysakointilupa
    currentCol = 16
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Pysakointilupa)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 17 (T): Hissi
    currentCol = 17
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.hissi)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 18 (U): Laivalippu
    currentCol = 18
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Laivalippu)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 19 (J): Puhelin
    currentCol = 19
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.puhelin)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 20 (AK): Sahköposti
    currentCol = 20
    valueToSet = sourceTask.sahkoposti
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 21 (K): Lastausosoite
    currentCol = 21
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.lastausOsoite)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 22 (L): Purkuosoite
    currentCol = 22
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.purkuOsoite)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 23 (W): Vakuutus
    currentCol = 23
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Vakuutus)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 24 (X): Arvo
    currentCol = 24
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Arvo, "-")
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 25 (Y): hinta
    currentCol = 25
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.hinta, "-")
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 26 (Z): Muuttomaailma
    currentCol = 26
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(sourceTask.Muuttomaailma, "EI"))): Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK": valueToSet = "KYLLÄ": Case Else: valueToSet = "EI": End Select
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 27 (AA): M3t
    currentCol = 27
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.M3t)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 28 (AB): LastauspaivaVarmistunut
    currentCol = 28
    If sourceTask.LastauspaivaVarmistunut Then valueToSet = "KYLLÄ" Else valueToSet = "EI"
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 29 (AC): PurkupaivaVarmistunut
    currentCol = 29
    If sourceTask.PurkupaivaVarmistunut Then valueToSet = "KYLLÄ" Else valueToSet = "EI"
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 30 (AD): TarjousTehty
    currentCol = 30
    tempDateVariant = sourceTask.tarjousTehty
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 31 (AE): TarjousHyväksytty
    currentCol = 31
    tempDateVariant = sourceTask.TarjousHyvaksytty
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 32 (AF): TarjousHylatty
    currentCol = 32
    tempDateVariant = sourceTask.TarjousHylatty
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 33 (AG): Tila
    currentCol = 33
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.Tila)
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 34 (AH): LastausLoppuu
    currentCol = 34
    tempDateVariant = sourceTask.LastausLoppuu
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet

    ' Sarake 35 (AI): PurkuLoppuu
    currentCol = 35
    tempDateVariant = sourceTask.PurkuLoppuu
    If IsDate(tempDateVariant) Then valueToSet = tempDateVariant Else valueToSet = Null
    targetRow.SetColumnValue currentCol, valueToSet
    
    'Sarake 36 (AJ): Kontaktipäivä
    currentCol = 36
    targetRow.SetColumnValue currentCol, validKontaktDate
    
    ' Sarake 37 (AK): Tila
    currentCol = 37
    valueToSet = mdlStringUtils.DefaultIfNull(sourceTask.valimatka)
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' Sarake 38 (AL): Laskutus
    currentCol = 38
    Select Case UCase(Trim(mdlStringUtils.DefaultIfNull(sourceTask.Laskutus, "EI"))): Case "KYLLÄ", "K", "TRUE", "YES", "1", "-1", "OK": valueToSet = "KYLLÄ": Case Else: valueToSet = "EI": End Select
    targetRow.SetColumnValue currentCol, valueToSet
    
    ' --- Tärkeät piilotetut apusarakkeet ---
    ' Sarake AG (33): Tila
    targetRow.SetColumnValue STATUS_COLUMN_INDEX, "KONTAKTI"

    ' Sarake 100 (oli AJ): ID
    targetRow.SetColumnValue DISPLAY_ID_COL, sourceTask.ID

    ' Sarake 101 (oli AK): RecordType
    targetRow.SetColumnValue DISPLAY_RECORD_TYPE_COL, "Kontakti"

    ' Sarake 102 (oli AL): SortDate (tässä kontaktiPaiva)
    targetRow.SetColumnValue DISPLAY_ATTENTION_DATE_COL, validKontaktDate

End Sub

Private Sub PopulateAttentionDisplayColumns(ByVal sourceTask As clsTaskItem, ByRef targetRow As clsDisplayRow)
    Dim combinedString As String
    Dim weekdayStr As String
    ' Dim dateStr As String ' Ei enää tarvita täällä
    Dim huomioStr As String
    Dim kuljettajatStr As String
    Dim autotStr As String
    Dim kontitStr As String
    Dim i As Long
    Dim validSortDate As Date ' Oletetaan, että targetRow.SortDate on validi Date

    ' Varmista, että SortDate on validi ennen käyttöä
    If Not IsDate(targetRow.SortDate) Then
        Debug.Print "PopulateAttentionDisplayColumns: Invalid SortDate for Task ID: " & sourceTask.ID
        ' Aseta jotkin oletukset tai poistu
        targetRow.SetColumnValue 1, "VIRHE PVM"
        targetRow.SetColumnValue 33, "HUOMIO"
        targetRow.SetColumnValue 100, sourceTask.ID
        targetRow.SetColumnValue 102, Null ' Ei validia päivämäärää
        Exit Sub
    End If
    validSortDate = targetRow.SortDate

    ' --- 1. Hae tarvittavat tiedot ---
    weekdayStr = mdlStringUtils.GetFinnishWeekdayName(validSortDate) ' Haetaan suomeksi ja isolla
    ' *** Muotoillaan päivämäärä merkkijonoksi VAIN näyttöä varten TÄSSÄ ***
    Dim displayDateStr As String
    displayDateStr = Format$(validSortDate, "dd.mm.yyyy") ' Muotoiltu päivämäärä merkkijonona

    huomioStr = Trim$(mdlStringUtils.DefaultIfNull(sourceTask.Huomioitavaa, "")) ' Poista ylim. välilyönnit

    ' Muodosta merkkijonot listoista (ei muutosta tähän logiikkaan)
    If Trim$(sourceTask.Kuljettajat) <> "" Then kuljettajatStr = " KULJETTAJAT: " & Replace(sourceTask.Kuljettajat, ";", ", ") Else kuljettajatStr = ""
    If Trim$(sourceTask.Autot) <> "" Then autotStr = " AUTOT: " & Replace(sourceTask.Autot, ";", ", ") Else autotStr = ""
    If Trim$(sourceTask.Kontit) <> "" Then kontitStr = " KONTIT: " & Replace(sourceTask.Kontit, ";", ", ") Else kontitStr = ""

    ' --- 2. Rakenna yhdistetty merkkijono NÄYTTÖÄ varten ---
    combinedString = weekdayStr & " " & displayDateStr ' Käytä muotoiltua merkkijonoa

    If huomioStr <> "" Then combinedString = combinedString & ", " & huomioStr
    If kuljettajatStr <> "" Then combinedString = combinedString & " -" & kuljettajatStr
    If autotStr <> "" Then combinedString = combinedString & " -" & autotStr
    If kontitStr <> "" Then combinedString = combinedString & " -" & kontitStr

    ' --- 3. Aseta yhdistetty NÄYTTÖMERKKIJONO ensimmäiseen sarakkeeseen ---
    targetRow.SetColumnValue 1, combinedString

    ' --- 4. Tyhjennä muut NÄYTTÖsarakkeet (2-32) ---
    For i = 2 To 32 ' Vain näkyvät sarakkeet A-AF
        targetRow.SetColumnValue i, vbNullString
    Next i

    ' --- 5. Aseta arvot PIILOTETTUIHIN APUsarakkeisiin ---
    targetRow.SetColumnValue 33, "HUOMIO"            ' Tila
    targetRow.SetColumnValue 100, sourceTask.ID       ' ID
    targetRow.SetColumnValue 102, validSortDate       ' *** Tallenna oikea Date ***

End Sub

' --- ApplyFormatting-alimakro ---
Private Sub ApplyFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    Dim dataRowRange As Range
    Dim headerCell As Range
    Dim headerFormatRange As Range
    Dim heightSetRange As Range ' Added for setting height
    Dim statusValue As String
    ' Dim standardRowHeight As Double ' Not needed

    ' --- Constants ---
    Const DATA_START_ROW As Long = 2
    Const HEADER_TEXT_COL As Long = 1 ' A
    Const DISPLAY_COLUMN_COUNT As Long = 102 ' AL
    Const LAST_VISIBLE_DATA_COL As Long = 102 ' AL
    Const LOAD_DATE_COL As Long = 3, UNLOAD_DATE_COL As Long = 6, SORT_DATE_COL As Long = 102
    Const LOAD_COUNTRY_COL As Long = 4, UNLOAD_COUNTRY_COL As Long = 5
    Const LOAD_CONFIRMED_COL As Long = 28, UNLOAD_CONFIRMED_COL As Long = 29, STATUS_COL As Long = 33
    Const clrTarjous As Long = 11184895, clrWhite As Long = 16777215, clrEiVarmojaPaivia As Long = 7602175
    Const clrHuomio As Long = 6579300, clrViikko As Long = 18070
    Const clrGrey As Long = 6579300, clrBlack As Long = 0
    Const clrKontakt As Long = 11184895
    Const ALL_ROWS_FIXED_HEIGHT As Double = 12 ' <<< ADJUST THIS VALUE AS NEEDED

    Application.ScreenUpdating = False
    On Error GoTo ApplyFormatting_Error

    ' --- VAIHE 0: ASETA PÄIVÄMÄÄRÄSARAKKEIDEN MUOTOILU ---
    On Error Resume Next ' Ohita virhe, jos aluetta ei voi muotoilla
    ' Sarakkeet, joissa näytetään päivämääriä
    ' Indeksit: 3, 6, 30, 31, 32, 34, 35, 37
    ' Vastaavat kirjainviittaukset: C, F, AD, AE, AF, AH, AI, AK
    ws.Range("C:C,F:F,AD:AF,AH:AJ").NumberFormat = "dd.mm.yyyy"
    If Err.Number <> 0 Then
        Debug.Print "ApplyFormatting: Virhe asetettaessa päivämäärämuotoilua: " & Err.Description
        Err.Clear
    End If
    On Error GoTo ApplyFormatting_Error ' Palauta päävirheenkäsittelijä

    ' --- VAIHE 1: YLEISET MUOTOILUT ---
    If lastRow >= DATA_START_ROW Then
        Dim overallDataRange As Range
        Set overallDataRange = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, DISPLAY_COLUMN_COUNT))
        If Not overallDataRange Is Nothing Then
            With overallDataRange
                .Font.Size = 9
                .WrapText = False ' Keep True for potential manual resize, fixed height prevents auto-expand
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlLeft
                .Font.Italic = False
                .Font.Bold = False
                .Font.Color = clrBlack
                .Interior.ColorIndex = xlNone
            End With
        End If
        Set overallDataRange = Nothing
    End If

    ' --- VAIHE 2: RIVIKOHTAISET MUOTOILUT (ILMAN RowHeight-asetuksia) ---
    For r = DATA_START_ROW To lastRow
        Set dataRowRange = ws.Range(ws.Cells(r, 1), ws.Cells(r, DISPLAY_COLUMN_COUNT))
        Set headerCell = ws.Cells(r, HEADER_TEXT_COL)

        ' Otsikkorivien käsittely
        If Left$(CStr(headerCell.value), 7) = "VIIKKO " Then
             On Error Resume Next
             Set headerFormatRange = ws.Range(headerCell, ws.Cells(r, LAST_VISIBLE_DATA_COL))
             If Not headerFormatRange Is Nothing Then
                 With headerFormatRange
                     .Font.Bold = True
                     .Interior.Color = clrViikko
                     .Font.Color = vbWhite
                     .VerticalAlignment = xlCenter
                     .WrapText = False ' No wrap for header
                 End With
             End If
             ' NO RowHeight set here
             Set headerFormatRange = Nothing
             On Error GoTo ApplyFormatting_Error

        Else ' Data-rivien käsittely
            On Error Resume Next
            statusValue = CStr(ws.Cells(r, STATUS_COL).value)

            ' Taustavärit
            Select Case UCase(Trim(statusValue))
                 Case "TARJOUS": dataRowRange.Interior.Color = clrTarjous
                 Case "HYVÄKSYTTY"
                     Dim loadConfirmedStr As String, unloadConfirmedStr As String
                     loadConfirmedStr = UCase(CStr(ws.Cells(r, LOAD_CONFIRMED_COL).value))
                     unloadConfirmedStr = UCase(CStr(ws.Cells(r, UNLOAD_CONFIRMED_COL).value))
                     If loadConfirmedStr <> "KYLLÄ" Or unloadConfirmedStr <> "KYLLÄ" Then
                          dataRowRange.Interior.Color = clrEiVarmojaPaivia
                     Else: dataRowRange.Interior.Color = clrWhite
                     End If
                 Case "HUOMIO": dataRowRange.Interior.Color = clrHuomio
                 Case "KONTAKTI": dataRowRange.Interior.Color = clrKontakt
            End Select

            ' Fonttimuotoilut Huomio-rivlle
            If UCase(Trim(statusValue)) = "HUOMIO" Then
                dataRowRange.Font.Color = RGB(255, 255, 255)
                dataRowRange.Font.Bold = True
                ws.Cells(r, 1).WrapText = False ' No wrap for Attention cell A1
                ' NO RowHeight set here
            ElseIf UCase(Trim(statusValue)) = "KONTAKTI" Then ' Muutos 2: Fonttimuotoilut Kontakti-riveille
                ' Esimerkki: Normaali musta teksti, ensimmäinen sarake lihavoitu
                dataRowRange.Font.Color = clrBlack
                dataRowRange.Font.Bold = False ' Oletus
                ws.Cells(r, 1).Font.Bold = True ' Lihavoi vain ensimmäinen sarake (Viikonpäivä, PVM, (KONTAKTI))
                ws.Cells(r, 1).WrapText = False ' No wrap for Kontaktin A1 solu
            Else ' Task-rivit
                ' Päivämäärien fontin värit
                Dim dateInLoadingColStr As String, rowSpecificDateStr As String, isLoadingRow As Boolean
                dateInLoadingColStr = CStr(ws.Cells(r, LOAD_DATE_COL).value)
                rowSpecificDateStr = CStr(ws.Cells(r, SORT_DATE_COL).value)
                isLoadingRow = (dateInLoadingColStr = rowSpecificDateStr)
                If isLoadingRow Then
                    ws.Cells(r, 1).Font.Bold = True
                    ws.Cells(r, LOAD_DATE_COL).Font.Color = clrBlack
                    ws.Cells(r, LOAD_DATE_COL).Font.Bold = True
                    ws.Cells(r, LOAD_COUNTRY_COL).Font.Color = clrBlack
                    ws.Cells(r, LOAD_COUNTRY_COL).Font.Bold = True
                    ws.Cells(r, UNLOAD_DATE_COL).Font.Color = clrGrey
                    ws.Cells(r, UNLOAD_COUNTRY_COL).Font.Color = clrGrey
                Else
                    ws.Cells(r, 1).Font.Bold = True
                    ws.Cells(r, LOAD_DATE_COL).Font.Color = clrGrey
                    ws.Cells(r, LOAD_COUNTRY_COL).Font.Color = clrGrey
                    ws.Cells(r, UNLOAD_DATE_COL).Font.Color = clrBlack
                    ws.Cells(r, UNLOAD_DATE_COL).Font.Bold = True
                    ws.Cells(r, UNLOAD_COUNTRY_COL).Font.Color = clrBlack
                    ws.Cells(r, UNLOAD_COUNTRY_COL).Font.Bold = True
                End If
            End If
            On Error GoTo ApplyFormatting_Error
        End If
    Next r

    ' --- VAIHE 3: ASETA KIINTEÄ RIVIKORKEUS KOKO ALUEELLE LOPUKSI ---
    If lastRow >= DATA_START_ROW Then
        Set heightSetRange = ws.Range(ws.Cells(DATA_START_ROW, 1), ws.Cells(lastRow, 1)) ' Range for Rows property
        If Not heightSetRange Is Nothing Then
            On Error Resume Next
            heightSetRange.rows.RowHeight = ALL_ROWS_FIXED_HEIGHT ' Apply fixed height
            On Error GoTo ApplyFormatting_Error
        End If
        Set heightSetRange = Nothing
    End If

CleanExit_Format:
    Application.ScreenUpdating = True
    Exit Sub

ApplyFormatting_Error:
    MsgBox "Virhe muotoilussa (ApplyFormatting):" & vbCrLf & Err.Description & vbCrLf & "Rivillä: " & r, vbExclamation
    Resume CleanExit_Format

End Sub
