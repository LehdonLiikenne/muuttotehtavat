Attribute VB_Name = "mdlRegisterUtils"
' --- Standard Module: mdlRegisterUtils ---
Option Explicit

' --- Vakiot sarakenumeroille (muuta tarvittaessa) ---
Public Const ID_COL As Long = 1       ' Sarake A
Public Const NAME_COL As Long = 2     ' Sarake B
Public Const PHONE_COL As Long = 3    ' Sarake C
Public Const EMAIL_COL As Long = 4    ' Sarake D
Public Const ADDRESS_COL As Long = 5  ' Sarake E
Public Const FIRST_DATA_ROW As Long = 2 ' Oletetaan otsikkorivi 1



' --- Haetaan Worksheet-olio ---
Private Function GetWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0 ' Palauta normaali virheenk‰sittely
End Function

' --- Etsii viimeisen k‰ytetyn rivin annetusta sarakkeesta ---
Private Function GetLastUsedRow(ws As Worksheet, columnToCheck As Long) As Long
    If ws Is Nothing Then
        GetLastUsedRow = 0
        Exit Function
    End If
    On Error Resume Next
    GetLastUsedRow = ws.Cells(ws.rows.Count, columnToCheck).End(xlUp).row
    If Err.Number <> 0 Then GetLastUsedRow = 1 ' Jos virhe tai tyhj‰, palauta 1 (otsikkorivi)
    If GetLastUsedRow < FIRST_DATA_ROW - 1 Then GetLastUsedRow = FIRST_DATA_ROW - 1 ' Varmista, ett‰ palauttaa v‰h. otsikkorivin
    On Error GoTo 0
End Function

' --- Hakee seuraavan vapaan ID:n v‰lilehdelt‰ etsim‰ll‰ suurimman nykyisen ID:n ---
Public Function GetNextRegisterID(sheetName As String) As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxID As Long
    
    Set ws = GetWorksheet(sheetName)
    If ws Is Nothing Then
        MsgBox "V‰lilehte‰ '" & sheetName & "' ei lˆytynyt ID:n hakua varten.", vbCritical, "Virhe"
        GetNextRegisterID = 0 ' Virhekoodi
        Exit Function
    End If

    lastRow = GetLastUsedRow(ws, ID_COL)

    ' Jos dataa ei ole (vain otsikko), aloitetaan ID:st‰ 1
    If lastRow < FIRST_DATA_ROW Then
        GetNextRegisterID = 1
        Exit Function
    End If

    ' Etsi suurin ID ID-sarakkeesta (k‰yt‰ WorksheetFunction.Max)
    On Error Resume Next ' Jos sarakkeessa ei ole numeroita tai on virheit‰
    maxID = Application.WorksheetFunction.Max(ws.Columns(ID_COL))
    If Err.Number <> 0 Then
        ' Jos Max ep‰onnistuu (esim. sarake tyhj‰ tai vain teksti‰), yrit‰ manuaalisesti
        maxID = 0
        Dim i As Long
        For i = FIRST_DATA_ROW To lastRow
             If IsNumeric(ws.Cells(i, ID_COL).value) Then
                 If CLng(ws.Cells(i, ID_COL).value) > maxID Then
                     maxID = CLng(ws.Cells(i, ID_COL).value)
                 End If
             End If
        Next i
        Err.Clear
    End If
    On Error GoTo 0

    GetNextRegisterID = maxID + 1 ' Palauta suurin ID + 1

End Function

' --- Etsii rivinumeron annetun arvon perusteella tietyst‰ sarakkeesta (Case-Insensitive) ---
' Palauttaa rivinumeron tai 0, jos ei lˆydy.
Public Function FindRowByValue(sheetName As String, searchValue As String, searchColumn As Long) As Long
    Dim ws As Worksheet
    Dim foundCell As Range
    FindRowByValue = 0 ' Oletus: ei lˆytynyt

    Set ws = GetWorksheet(sheetName)
    If ws Is Nothing Then Exit Function

    On Error Resume Next ' Jos Find ei lˆyd‰ mit‰‰n
    Set foundCell = ws.Columns(searchColumn).Find(What:=searchValue, _
                                                 LookIn:=xlValues, _
                                                 LookAt:=xlWhole, _
                                                 SearchOrder:=xlByRows, _
                                                 SearchDirection:=xlNext, _
                                                 MatchCase:=False) ' False = Case-Insensitive
    On Error GoTo 0

    If Not foundCell Is Nothing Then
        ' Varmista ettei lˆytynyt otsikkorivilt‰ (jos etsint‰ kohdistuu myˆs siihen)
        If foundCell.row >= FIRST_DATA_ROW Then
             FindRowByValue = foundCell.row
        End If
    End If
    Set foundCell = Nothing
End Function

' --- Etsii rivinumeron annetun ID:n perusteella ID-sarakkeesta ---
' Palauttaa rivinumeron tai 0, jos ei lˆydy.
Public Function FindRowByID(sheetName As String, itemID As Long) As Long
    ' K‰ytet‰‰n FindRowByValue-funktiota ID-sarakkeelle
    FindRowByID = FindRowByValue(sheetName, CStr(itemID), ID_COL)
End Function


' --- Lataa rekisteritiedot v‰lilehdelt‰ annettuun ListBoxiin ---
' Tunnistaa v‰lilehden nimen perusteella, montako saraketta ladataan.
' Asettaa ListBoxin sarakkeet ja piilottaa ID-sarakkeen.
' --- Lataa rekisteritiedot v‰lilehdelt‰ annettuun ListBoxiin (VERSIO 3 - Korjattu 1D/2D -k‰sittely) ---
Public Sub LoadRegisterDataToListBox(lst As MSForms.ListBox, sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataArray As Variant
    Dim colCount As Long
    Dim i As Long, j As Long
    Dim readRange As Range

    On Error GoTo LoadError

    ' --- 1. Hae v‰lilehti ---
    Set ws = GetWorksheet(sheetName)
    If ws Is Nothing Then
        ' GetWorksheet hoitaa virheilmoituksen, jos tarpeen
        Exit Sub
    End If

    ' --- 2. Tyhjenn‰ ListBox ---
    lst.Clear

    ' --- 3. M‰‰rit‰ sarakkeiden m‰‰r‰ ja lue data ---
    lastRow = GetLastUsedRow(ws, ID_COL) ' Hae viimeinen rivi ID-sarakkeen perusteella

    ' Jos dataa ei ole (vain otsikko tai tyhj‰)
    If lastRow < FIRST_DATA_ROW Then
        lst.ColumnCount = 1 ' Varmista, ett‰ ainakin yksi sarake on m‰‰ritetty
        lst.AddItem "Ei tietoja" ' Voit lis‰t‰ viestin tyhj‰‰n listaan
        Exit Sub
    End If

    ' P‰‰ttele sarakkeiden m‰‰r‰ v‰lilehden nimen perusteella
    Select Case LCase(sheetName) ' K‰yt‰ LCase varmuuden vuoksi
        Case "kuljettajat", "apulaiset"
            colCount = 5 ' ID, Nimi, Puhelin, S‰hkˆposti, Osoite
        Case "palvelut", "autot", "kontit"
            colCount = 2 ' ID, Nimi
        Case Else
            ' Tuntematon v‰lilehti, oletetaan yksinkertainen
            colCount = 2
            Debug.Print "LoadRegisterDataToListBox: Tuntematon v‰lilehti '" & sheetName & "', oletetaan 2 saraketta."
    End Select

    ' --- Lue data (Varmistaa oikean leveyden ennen lukua) ---
    Set readRange = ws.Range(ws.Cells(FIRST_DATA_ROW, ID_COL), ws.Cells(lastRow, ID_COL)).Resize(, colCount)
    ' Debug.Print "LoadRegisterDataToListBox (" & sheetName & "): Luetaan alueelta: " & readRange.Address ' Voit poistaa t‰m‰n debugin
    dataArray = readRange.value
    Set readRange = Nothing

    ' --- 4. Konfiguroi ListBox ---
    lst.ColumnCount = colCount
    ' Aseta sarakeleveydet (piilota ID, s‰‰d‰ muita tarpeen mukaan)
    Select Case colCount
        Case 5 ' Kuljettajat, Apulaiset
             lst.ColumnWidths = "0 pt; 120 pt; 80 pt; 80 pt; 120 pt" ' Esimerkkileveydet
        Case 2 ' Palvelut, Autot, Kontit
             lst.ColumnWidths = "0 pt; 200 pt" ' Esimerkkileveys
        Case Else
             lst.ColumnWidths = "0 pt; 100 pt" ' Oletus
    End Select

    ' --- 5. T‰yt‰ ListBox datalla (KƒSITTELE AINA 2D-TAULUKKONA) ---
    If IsArray(dataArray) Then
        Dim actualColsInArray As Long
        Dim arrayLBound1 As Long, arrayUBound1 As Long
        Dim arrayLBound2 As Long, arrayUBound2 As Long

        ' M‰‰rit‰ todelliset rajat turvallisesti
        On Error Resume Next ' Varmista, ettei kaadu, jos ei ole 2D
        arrayLBound1 = LBound(dataArray, 1)
        arrayUBound1 = UBound(dataArray, 1)
        arrayLBound2 = LBound(dataArray, 2)
        arrayUBound2 = UBound(dataArray, 2)
        If Err.Number <> 0 Then
            ' Jos ei ollut 2D (ep‰todenn‰kˆist‰ nyt), ‰l‰ tee mit‰‰n t‰ss‰
             Debug.Print "LoadRegisterDataToListBox (" & sheetName & "): Varoitus - dataArray ei ollut 2D-taulukko?"
             Err.Clear
        Else
            ' Oli 2D-taulukko, jatketaan
            actualColsInArray = arrayUBound2 ' Todellinen sarakem‰‰r‰
            On Error GoTo LoadError ' Palauta normaali virheenk‰sittely

            For i = arrayLBound1 To arrayUBound1 ' K‰yt‰ todellisia rajoja
                lst.AddItem CStr(dataArray(i, arrayLBound2)) ' Lis‰‰ ID (1. sarake)
                For j = 2 To colCount ' K‰y l‰pi loput ODOTETUT sarakkeet
                    If j <= actualColsInArray Then
                        ' Lis‰‰ arvo, jos se on olemassa taulukossa
                        lst.List(lst.ListCount - 1, j - 1) = CStr(dataArray(i, j))
                    Else
                        ' Lis‰‰ tyhj‰, jos odotettua saraketta ei ollutkaan taulukossa
                        lst.List(lst.ListCount - 1, j - 1) = ""
                    End If
                Next j
            Next i
        End If
        On Error GoTo LoadError ' Palauta normaali virheenk‰sittely

    Else
        ' Jos dataArray ei ollut taulukko lainkaan (hyvin ep‰todenn‰kˆist‰)
        If colCount >= 1 Then lst.AddItem CStr(dataArray)
    End If

    ' Puhdistus
    Set ws = Nothing
    If IsArray(dataArray) Then Erase dataArray
    Exit Sub

LoadError:
    MsgBox "Virhe ladattaessa tietoja ListBoxiin '" & lst.Name & "' v‰lilehdelt‰ '" & sheetName & "':" & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Listan Latausvirhe"
    ' Yrit‰ puhdistaa
    Set ws = Nothing
    If IsArray(dataArray) Then Erase dataArray
    On Error Resume Next ' Varmista, ett‰ Clear toimii
    lst.Clear
    On Error GoTo 0
End Sub

' --- Lis‰‰ uuden rivin rekisteriv‰lilehdelle annetuilla tiedoilla ---
' dataArray sis‰lt‰‰ lis‰tt‰v‰t tiedot, mukaan lukien UUSI ID ensimm‰isen‰ alkiona.
' Palauttaa True, jos lis‰ys onnistui, muuten False.
Public Function AddRegisterItem(sheetName As String, dataArray As Variant) As Boolean
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim colCount As Long
    Dim i As Long

    AddRegisterItem = False ' Oletus: ep‰onnistui
    On Error GoTo AddError

    Set ws = GetWorksheet(sheetName)
    If ws Is Nothing Then
        MsgBox "V‰lilehte‰ '" & sheetName & "' ei lˆytynyt. Lis‰ys ep‰onnistui.", vbCritical, "Virhe"
        Exit Function
    End If

    ' Tarkista, ett‰ dataArray on validi taulukko
    If Not IsArray(dataArray) Then
         MsgBox "Lis‰tt‰v‰ data ei ole kelvollinen taulukko.", vbCritical, "Virhe"
         Exit Function
    End If
    colCount = UBound(dataArray) ' Sarakkeiden m‰‰r‰

    ' Hae seuraava vapaa rivi (ID-sarakkeen perusteella)
    nextRow = GetLastUsedRow(ws, ID_COL) + 1
    ' Jos edellinen oli otsikko (rivi 1), varmista ett‰ aloitetaan oikealta rivilt‰
    If nextRow < FIRST_DATA_ROW Then nextRow = FIRST_DATA_ROW

    ' Kirjoita data v‰lilehdelle
    Application.ScreenUpdating = False ' Nopeampi kirjoitus
    For i = 1 To colCount
        ws.Cells(nextRow, i).value = dataArray(i)
    Next i
    Application.ScreenUpdating = True

    AddRegisterItem = True ' Onnistui

CleanExit_Add:
    Set ws = Nothing
    Exit Function

AddError:
    Application.ScreenUpdating = True
    MsgBox "Virhe lis‰tt‰ess‰ tietoa v‰lilehdelle '" & sheetName & "':" & vbCrLf & Err.Description, vbCritical, "Lis‰ysvirhe"
    ' Palauttaa False (oletus)
    Resume CleanExit_Add
End Function

' --- P‰ivitt‰‰ olemassa olevan rivin tiedot ID:n perusteella ---
' dataArray sis‰lt‰‰ p‰ivitetyt tiedot (ID, Nimi, [Puhelin] jne.)
' Palauttaa True, jos p‰ivitys onnistui, muuten False.
Public Function UpdateRegisterItem(sheetName As String, itemID As Long, dataArray As Variant) As Boolean
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim colCount As Long
    Dim i As Long

    UpdateRegisterItem = False ' Oletus: ep‰onnistui
    On Error GoTo UpdateError

    Set ws = GetWorksheet(sheetName)
    If ws Is Nothing Then
        MsgBox "V‰lilehte‰ '" & sheetName & "' ei lˆytynyt. P‰ivitys ep‰onnistui.", vbCritical, "Virhe"
        Exit Function
    End If

    ' Tarkista dataArray
    If Not IsArray(dataArray) Then
         MsgBox "P‰ivitett‰v‰ data ei ole kelvollinen taulukko.", vbCritical, "Virhe"
         Exit Function
    End If
    colCount = UBound(dataArray)

    ' Etsi p‰ivitett‰v‰n rivin numero ID:n perusteella
    targetRow = FindRowByID(sheetName, itemID)
    If targetRow = 0 Then
        MsgBox "P‰ivitett‰v‰‰ tietoa ID:ll‰ " & itemID & " ei lˆytynyt v‰lilehdelt‰ '" & sheetName & "'.", vbExclamation, "Ei Lˆytynyt"
        Exit Function ' Ei lˆytynyt, palauta False
    End If

    ' Kirjoita p‰ivitetyt tiedot v‰lilehdelle lˆydetylle riville
    Application.ScreenUpdating = False
    For i = 1 To colCount
        ws.Cells(targetRow, i).value = dataArray(i)
    Next i
    Application.ScreenUpdating = True

    UpdateRegisterItem = True ' Onnistui

CleanExit_Update:
    Set ws = Nothing
    Exit Function

UpdateError:
    Application.ScreenUpdating = True
    MsgBox "Virhe p‰ivitett‰ess‰ tietoa (ID: " & itemID & ") v‰lilehdelle '" & sheetName & "':" & vbCrLf & Err.Description, vbCritical, "P‰ivitysvirhe"
    ' Palauttaa False (oletus)
    Resume CleanExit_Update
End Function

' --- Poistaa rivin rekisteriv‰lilehdelt‰ ID:n perusteella ---
' Palauttaa True, jos poisto onnistui, muuten False.
Public Function DeleteRegisterItem(sheetName As String, itemID As Long) As Boolean
    Dim ws As Worksheet
    Dim targetRow As Long

    DeleteRegisterItem = False ' Oletus: ep‰onnistui
    On Error GoTo DeleteError

    Set ws = GetWorksheet(sheetName)
    If ws Is Nothing Then
        MsgBox "V‰lilehte‰ '" & sheetName & "' ei lˆytynyt. Poisto ep‰onnistui.", vbCritical, "Virhe"
        Exit Function
    End If

    ' Etsi poistettavan rivin numero ID:n perusteella
    targetRow = FindRowByID(sheetName, itemID)
    If targetRow = 0 Then
        MsgBox "Poistettavaa tietoa ID:ll‰ " & itemID & " ei lˆytynyt v‰lilehdelt‰ '" & sheetName & "'.", vbExclamation, "Ei Lˆytynyt"
        Exit Function ' Ei lˆytynyt, palauta False
    End If

    ' Poista koko rivi
    Application.ScreenUpdating = False
    ws.rows(targetRow).Delete Shift:=xlUp
    Application.ScreenUpdating = True

    DeleteRegisterItem = True ' Onnistui

CleanExit_Delete:
    Set ws = Nothing
    Exit Function

DeleteError:
    Application.ScreenUpdating = True
    MsgBox "Virhe poistettaessa tietoa (ID: " & itemID & ") v‰lilehdelt‰ '" & sheetName & "':" & vbCrLf & Err.Description, vbCritical, "Poistovirhe"
    ' Palauttaa False (oletus)
    Resume CleanExit_Delete
End Function

