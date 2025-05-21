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
    'Debug.Print Now & " ThisWorkbook_Open: Ladataan data ja p‰ivitet‰‰n n‰kym‰..." ' Aika mukaan lokiin
    
    Call HideSheets ' Kutsutaan ennen p‰‰virheenk‰sittelij‰‰
    Call ShowSheets  ' Kutsutaan ennen p‰‰virheenk‰sittelij‰‰
    
    On Error GoTo OpenError ' Aktivoi virheenk‰sittely t‰lle rutiinille

    mdlMain.UpdateMainView ' Kutsu p‰‰ohjelmaa, joka sis‰lt‰‰ oman virheenk‰sittelyns‰

    'Debug.Print Now & " ThisWorkbook_Open: P‰ivityskutsu valmis." ' Lokimerkint‰ onnistumisesta
    
    ' Kutsu varmuuskopiointirutiinia, kun tyˆkirja avataan.
    On Error Resume Next ' Jatketaan, vaikka varmuuskopioinnissa tulisi virhe, jotta tiedosto aukeaa silti
    Call mdlBackups.TarkistaJaLuoVarmuuskopioAutomaattisesti
    If Err.Number <> 0 Then
        MsgBox "Automaattisessa varmuuskopioinnissa tapahtui virhe:" & vbCrLf & Err.Description, vbExclamation, "Varmuuskopiointivirhe"
        Err.Clear
    End If
    On Error GoTo 0 ' Palauta normaali virheenk‰sittely
    
CleanExit: ' Poistutaan aina t‰m‰n kautta
    On Error GoTo 0 ' Nollaa virheenk‰sittely ennen poistumista
    'Debug.Print Now & " ThisWorkbook_Open: Poistutaan rutiinista."
    Exit Sub ' Varmista poistuminen

OpenError: ' Virheenk‰sittelij‰ t‰lle Workbook_Open -rutiinille
    MsgBox "Kriittinen virhe tiedostoa avattaessa tai dataa p‰ivitett‰ess‰:" & vbCrLf & vbCrLf & _
           "Virhe " & Err.Number & ": " & Err.Description, vbCritical, "Avausvirhe"
    Resume CleanExit ' Yrit‰ siirty‰ siistiin poistumiseen virheen j‰lkeen

End Sub

