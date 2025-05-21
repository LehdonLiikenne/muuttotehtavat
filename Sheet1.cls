VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim taskID As Long
    Dim idCellValue As Variant
    Dim statusCellValue As Variant
    Dim clickedRow As Long
    Dim isAttentionRow As Boolean

    Const DISPLAY_ID_COLUMN As Long = 100     ' AJ = ID
    Const DISPLAY_STATUS_COLUMN As Long = 33 ' AG = Tila
    Const DISPLAY_DATA_START_ROW As Long = 2
    Const HEADER_TEXT_COL As Long = 1

    ' --- Poistetaan DEBUG-tulostukset (voit j‰tt‰‰, jos haluat) ---
    ' Debug.Print "--- DoubleClick Start ---"
    ' Debug.Print "Target Address: " & Target.Address

    On Error GoTo DoubleClickErrorHandler

    clickedRow = Target.row
    ' Debug.Print "Clicked Row: " & clickedRow

    ' --- Alkutarkistukset ---
    If clickedRow < DISPLAY_DATA_START_ROW Then GoTo CleanExit
    Dim headerCheckValue As Variant
    headerCheckValue = Me.Cells(clickedRow, HEADER_TEXT_COL).value
    If Not IsError(headerCheckValue) Then
        If Left$(CStr(headerCheckValue), 7) = "VIIKKO " Then GoTo CleanExit
    Else
        GoTo CleanExit ' Virhe luettaessa otsikkoa
    End If

    ' --- Hae Teht‰v‰n ID turvallisesti ---
    idCellValue = Me.Cells(clickedRow, DISPLAY_ID_COLUMN).value
    ' Debug.Print "Value from ID Cell (Column " & DISPLAY_ID_COLUMN & "): '" & CStr(idCellValue) & "'"

    taskID = 0
    If IsNumeric(idCellValue) Then
        If idCellValue > 0 Then
            taskID = CLng(idCellValue)
        End If
    End If
    ' Debug.Print "Task ID after conversion: " & taskID

    If taskID <= 0 Then
        GoTo CleanExit
    End If

    ' --- Tunnista Huomiorivi ---
    isAttentionRow = False
    statusCellValue = Me.Cells(clickedRow, DISPLAY_STATUS_COLUMN).value
    If Not IsError(statusCellValue) Then
        If UCase(Trim(CStr(statusCellValue))) = "HUOMIO" Then
            isAttentionRow = True
        End If
    End If

    ' --- Peruuta Excelin oletustoiminto ---
    Cancel = True

    ' --- Kutsu oikeaa toimintoa ---
    If isAttentionRow Then
        mdlMain.ShowAttentionForm_Edit taskID
    Else
        mdlMain.ShowTaskForm_Edit taskID
    End If

CleanExit:
    On Error GoTo 0
    Exit Sub

DoubleClickErrorHandler:
    ' Debug.Print "*** ERROR in DoubleClick: " & Err.Description & " ***"
    MsgBox "VIRHE DoubleClickiss‰: " & Err.Description, vbCritical, "Tuplaklikkausvirhe"
    Resume CleanExit

End Sub
Private Sub optKaikki_Click()
    Call mdlMain.UpdateMainView ' P‰ivit‰ n‰yttˆ, kun valinta muuttuu
End Sub

Private Sub optTarjoukset_Click()
    Call mdlMain.UpdateMainView ' P‰ivit‰ n‰yttˆ, kun valinta muuttuu
End Sub

Private Sub optVarmistuneet_Click()
    Call mdlMain.UpdateMainView ' P‰ivit‰ n‰yttˆ, kun valinta muuttuu
End Sub


Private Sub chkLastausOK_Click()
    Call mdlMain.UpdateMainView ' P‰ivit‰ n‰yttˆ, kun valinta muuttuu
End Sub

Private Sub chkPurkuOK_Click()
    Call mdlMain.UpdateMainView ' P‰ivit‰ n‰yttˆ, kun valinta muuttuu
End Sub

Private Sub chkLaskuttamatta_Click()
    Call mdlMain.UpdateMainView ' P‰ivit‰ n‰yttˆ, kun valinta muuttuu
End Sub

' --- APUFUNKTIO: Palauttaa suodatinasetukset taulukkona ---
' Palauttaa Variant-taulukon:
' Index 1: filterMode ("Kaikki", "Tarjoukset", "Varmistuneet") As String
' Index 2: showLastausOK As Boolean
' Index 3: showPurkuOK As Boolean
' Index 4: showLaskuttamatta As Boolean <<< UUSI
' --- Funktio Sheet("Teht‰v‰t") -moduulissa (Korjattu virheenk‰sittely) ---
Public Function GetFilterSettings() As Variant
    Dim settings(1 To 4) As Variant
    'Debug.Print "--- GetFilterSettings ALKAA ---"
    ' Aseta oletukset HETI silt‰ varalta, ett‰ tulee virhe
    settings(1) = "Kaikki"
    settings(2) = False
    settings(3) = False
    settings(4) = False

    On Error GoTo SettingsError ' Aktivoi virheenk‰sittely

    ' --- Yrit‰ lukea kontrollit ---
    'Debug.Print "  Yritet‰‰n lukea OptionButtons..."
    If Me.optKaikki.value = True Then
        settings(1) = "Kaikki"
    ElseIf Me.optTarjoukset.value = True Then
        settings(1) = "Tarjoukset"
    ElseIf Me.optVarmistuneet.value = True Then
        settings(1) = "Varmistuneet"
    Else: settings(1) = "Kaikki"
    End If
    'Debug.Print "  Modeksi luettu: " & settings(1)

    settings(2) = (Me.chkLastausOK.value = True)
    settings(3) = (Me.chkPurkuOK.value = True)
    settings(4) = (Me.chkLaskuttamatta.value = True)
    'Debug.Print "  Luettu LastausOK=" & settings(2) & ", PurkuOK=" & settings(3) & ", Laskuttamatta=" & settings(4)

    ' Jos t‰nne p‰‰stiin, kaikki OK -> Aseta palautusarvo
    GetFilterSettings = settings
    'Debug.Print "--- GetFilterSettings LOPPUU (OK) ---"
    Exit Function ' <<< TƒRKEƒƒ: Poistu onnistuneesti ENNEN virheenk‰sittelij‰‰

SettingsError: ' Virheenk‰sittelij‰ alkaa t‰st‰
     'Debug.Print "*** VIRHE GetFilterSettings-funktiossa: " & Err.Number & " - " & Err.Description & " ***"
     MsgBox "Virhe luettaessa suodatinasetuksia v‰lilehdelt‰ '" & Me.Name & "':" & vbCrLf & Err.Description & vbCrLf & _
            "Palautetaan oletusasetukset. Tarkista kontrollien nimet.", vbCritical, "Asetusvirhe"
     ' Palauta oletusasetukset (jotka asetettiin jo alussa)
     GetFilterSettings = settings
     'Debug.Print "--- GetFilterSettings LOPPUU (Error Return) ---"
     ' Ei Resume-komentoa, funktiosta poistutaan luonnollisesti t‰m‰n j‰lkeen
End Function

