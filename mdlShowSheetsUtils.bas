Attribute VB_Name = "mdlShowSheetsUtils"
Public Sub HideSheets()
    Dim sheetNamesToHide As Variant
    Dim sheetName As Variant
    Dim ws As Worksheet

    ' --- LISTAA T�H�N KAIKKI V�LILEHDET, JOTKA HALUAT PIILOTTAA ---
    sheetNamesToHide = Array("Palvelut", "Kuljettajat", "Apulaiset", "Autot", "Kontit", "Config")

    On Error Resume Next ' Ohitetaan virheet, jos jokin lehti puuttuu listalta
    For Each sheetName In sheetNamesToHide
        Set ws = Nothing ' Nollaa ws joka kierroksella
        Set ws = ThisWorkbook.Worksheets(CStr(sheetName)) ' Hae v�lilehti nimen perusteella

        If Not ws Is Nothing Then ' Jos v�lilehti l�ytyi
             If ws.Visible <> xlSheetVeryHidden Then
                ws.Visible = xlSheetVeryHidden
                Debug.Print "V�lilehti '" & sheetName & "' piilotettu (VeryHidden)." ' Tulostaa viestin Immediate-ikkunaan (Ctrl+G)
             End If
        Else
            'Debug.Print "Varoitus: V�lilehte� nimelt� '" & sheetName & "' ei l�ytynyt."
            ' Voit halutessasi n�ytt�� MsgBoxin k�ytt�j�lle:
            ' MsgBox "Varoitus: V�lilehte� nimelt� '" & sheetName & "' ei l�ytynyt.", vbExclamation
        End If
    Next sheetName
    On Error GoTo 0 ' Palauta normaali virheenk�sittely

    'MsgBox "M��ritetyt v�lilehdet on piilotettu (VeryHidden).", vbInformation
    Set ws = Nothing
End Sub

Public Sub ShowSheets()
    Dim ws As Worksheet
    Const LEHDEN_NIMI As String = "Tietovarasto" ' <-- MUUTA T�H�N SEN V�LILEHDEN NIMI, JONKA HALUAT N�YTT��

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LEHDEN_NIMI)
    On Error GoTo 0

    If Not ws Is Nothing Then
        If ws.Visible <> xlSheetVisible Then
            ws.Visible = xlSheetVisible
            'MsgBox "V�lilehti '" & LEHDEN_NIMI & "' on nyt n�kyviss�.", vbInformation
        Else
             'MsgBox "V�lilehti '" & LEHDEN_NIMI & "' oli jo n�kyviss�.", vbInformation
        End If
    Else
        MsgBox "Virhe: V�lilehte� nimelt� '" & LEHDEN_NIMI & "' ei l�ytynyt.", vbCritical, "Virhe"
    End If

    Set ws = Nothing
End Sub
