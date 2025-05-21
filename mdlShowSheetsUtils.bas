Attribute VB_Name = "mdlShowSheetsUtils"
Public Sub HideSheets()
    Dim sheetNamesToHide As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet

    ' --- LISTAA TÄHÄN KAIKKI VÄLILEHDET, JOTKA HALUAT PIILOTTAA ---
    sheetNamesToHide = Array("Palvelut", "Kuljettajat", "Apulaiset", "Autot", "Kontit", "Config")

    On Error Resume Next ' Ohitetaan virheet, jos jokin lehti puuttuu listalta
    For Each sheetName In sheetNamesToHide
        Set ws = Nothing ' Nollaa ws joka kierroksella
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName)) ' Hae välilehti nimen perusteella

        If Not ws Is Nothing Then ' Jos välilehti löytyi
             If ws.Visible <> xlSheetVeryHidden Then
                ws.Visible = xlSheetVeryHidden
                Debug.Print "Välilehti '" & sheetName & "' piilotettu (VeryHidden)." ' Tulostaa viestin Immediate-ikkunaan (Ctrl+G)
             End If
        Else
            'Debug.Print "Varoitus: Välilehteä nimeltä '" & sheetName & "' ei löytynyt."
            ' Voit halutessasi näyttää MsgBoxin käyttäjälle:
            ' MsgBox "Varoitus: Välilehteä nimeltä '" & sheetName & "' ei löytynyt.", vbExclamation
        End If
    Next sheetName
    On Error GoTo 0 ' Palauta normaali virheenkäsittely

    'MsgBox "Määritetyt välilehdet on piilotettu (VeryHidden).", vbInformation
    Set ws = Nothing
End Sub

Public Sub ShowSheets()
    Dim ws As Worksheet
    Const LEHDEN_NIMI As String = "Tietovarasto" ' <-- MUUTA TÄHÄN SEN VÄLILEHDEN NIMI, JONKA HALUAT NÄYTTÄÄ

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEHDEN_NIMI)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            'MsgBox "Välilehti '" & LEHDEN_NIMI & "' on nyt näkyvissä.", vbInformation
        Else
             'MsgBox "Välilehti '" & LEHDEN_NIMI & "' oli jo näkyvissä.", vbInformation
        End If
    Else
        MsgBox "Virhe: Välilehteä nimeltä '" & LEHDEN_NIMI & "' ei löytynyt.", vbCritical, "Virhe"
    End If

    Set ws = Nothing
End Sub
