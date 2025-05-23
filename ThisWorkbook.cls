VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    'Debug.Print Now & " ThisWorkbook_Open: Ladataan data ja päivitetään näkymä..." ' Aika mukaan lokiin
    
    Call HideSheets ' Kutsutaan ennen päävirheenkäsittelijää
    Call ShowSheets  ' Kutsutaan ennen päävirheenkäsittelijää
    
    On Error GoTo OpenError ' Aktivoi virheenkäsittely tälle rutiinille

    mdlMain.UpdateMainView ' Kutsu pääohjelmaa, joka sisältää oman virheenkäsittelynsä

    'Debug.Print Now & " ThisWorkbook_Open: Päivityskutsu valmis." ' Lokimerkintä onnistumisesta
    
    ' Kutsu varmuuskopiointirutiinia, kun työkirja avataan.
    On Error Resume Next ' Jatketaan, vaikka varmuuskopioinnissa tulisi virhe, jotta tiedosto aukeaa silti
    Call mdlBackups.TarkistaJaLuoVarmuuskopioAutomaattisesti
    If Err.Number <> 0 Then
        MsgBox "Automaattisessa varmuuskopioinnissa tapahtui virhe:" & vbCrLf & Err.Description, vbExclamation, "Varmuuskopiointivirhe"
        Err.Clear
    End If
    On Error GoTo 0 ' Palauta normaali virheenkäsittely
    
CleanExit: ' Poistutaan aina tämän kautta
    On Error GoTo 0 ' Nollaa virheenkäsittely ennen poistumista
    'Debug.Print Now & " ThisWorkbook_Open: Poistutaan rutiinista."
    Exit Sub ' Varmista poistuminen

OpenError: ' Virheenkäsittelijä tälle Workbook_Open -rutiinille
    MsgBox "Kriittinen virhe tiedostoa avattaessa tai dataa päivitettäessä:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Avausvirhe"
    Resume CleanExit ' Yritä siirtyä siistiin poistumiseen virheen jälkeen

End Sub

