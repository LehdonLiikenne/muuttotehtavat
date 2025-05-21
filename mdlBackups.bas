Attribute VB_Name = "mdlBackups"
Option Explicit ' Suositeltavaa: Pakottaa muuttujien esittelyn

' M‰‰rit‰ s‰ilytett‰vien varmuuskopioiden enimm‰ism‰‰r‰
Const MAX_BACKUPS As Long = 5

Sub TarkistaJaLuoVarmuuskopioAutomaattisesti()

    ' M‰‰ritell‰‰n tarvittavat muuttujat
    Dim FSO As Object             ' Tiedostoj‰rjestelm‰objekti
    Dim WShell As Object          ' Windows Script Shell -objekti
    Dim backupFolderObj As Object ' Varmuuskopiokansion objekti
    Dim fileItem As Object        ' Yksitt‰inen tiedosto kansiossa
    Dim oldestFile As Object      ' Viittaus vanhimpaan varmuuskopiotiedostoon
    Dim srcWb As Workbook         ' L‰hdetyˆkirja (joka avattiin)
    Dim destFolderPath As String  ' Kohdekansion polku
    Dim backupBaseFolder As String ' Varmuuskopioiden alikansion nimi
    Dim baseName As String        ' Tiedoston nimi ilman tarkenninta
    Dim fileExt As String         ' Tiedostotarkenne
    Dim timeStamp As String       ' Aikaleima tiedostonimeen
    Dim fullDestPath As String    ' Koko kohdepolku ja tiedostonimi
    Dim backupFiles As Object     ' Kokoelma relevanteista varmuuskopiotiedostoista (k‰ytet‰‰n Dictionarya)
    Dim currentBackupCount As Long ' Lˆydettyjen varmuuskopioiden m‰‰r‰
    Dim oldestDate As Date        ' Vanhimman tiedoston luontip‰iv‰m‰‰r‰
    Dim filePathKey As Variant    ' Avain Dictionaryssa (tiedoston polku)

    ' Virheenk‰sittely p‰‰lle
    On Error GoTo VirheKasittelija

    ' Aseta l‰hdetyˆkirjaksi t‰m‰ tyˆkirja
    Set srcWb = ThisWorkbook

    ' Luo tarvittavat COM-objektit
    Set WShell = CreateObject("WScript.Shell")
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' --- M‰‰rit‰ kohdekansio (sama kuin aiemmin, esim. Tyˆpˆyt‰) ---
    destFolderPath = WShell.SpecialFolders("USERPROFILE") ' Tai "MyDocuments", Environ("USERPROFILE") jne.
    backupBaseFolder = "Muuttoteht‰v‰t_Excel_Varmuuskopiot"          ' Varmuuskopioiden alikansio
    destFolderPath = destFolderPath & "\" & backupBaseFolder

    ' --- Varmista, ett‰ varmuuskopiokansio on olemassa ---
    If Not FSO.FolderExists(destFolderPath) Then
        FSO.CreateFolder destFolderPath
    End If

    ' --- Hae t‰m‰n tyˆkirjan perustiedot ---
    baseName = FSO.GetBaseName(srcWb.Name)
    fileExt = FSO.GetExtensionName(srcWb.Name)

    ' --- Etsi olemassa olevat varmuuskopiot t‰lle tiedostolle ---
    Set backupFolderObj = FSO.GetFolder(destFolderPath)
    Set backupFiles = CreateObject("Scripting.Dictionary") ' K‰ytet‰‰n Dictionarya polun ja p‰iv‰m‰‰r‰n tallentamiseen
    currentBackupCount = 0

    For Each fileItem In backupFolderObj.Files
        ' Tarkista, vastaako tiedoston nimi varmuuskopiointikaavaa TƒLLE tyˆkirjalle
        If FSO.GetBaseName(fileItem.Name) Like baseName & "_*" And FSO.GetExtensionName(fileItem.Name) = fileExt Then
            ' Lis‰‰ tiedoston polku ja luontip‰iv‰m‰‰r‰ Dictionaryyn
            If Not backupFiles.Exists(fileItem.Path) Then
                 backupFiles.Add fileItem.Path, fileItem.DateCreated
                 currentBackupCount = currentBackupCount + 1
            End If
        End If
    Next fileItem

    ' --- Jos varmuuskopioiden m‰‰r‰ ylitt‰‰ rajan, etsi ja poista vanhin ---
    If currentBackupCount >= MAX_BACKUPS Then
        ' Etsi vanhin tiedosto vertailemalla p‰iv‰m‰‰ri‰ Dictionaryssa
        Set oldestFile = Nothing
        oldestDate = Now ' Alustetaan nykyhetkell‰ (mik‰ tahansa tuleva p‰iv‰m‰‰r‰ k‰y)

        For Each filePathKey In backupFiles.Keys
            If backupFiles(filePathKey) < oldestDate Then
                oldestDate = backupFiles(filePathKey)
                ' Hae File-objekti FSO:n kautta polun perusteella
                 Set oldestFile = FSO.GetFile(filePathKey)
            End If
        Next filePathKey

        ' Poista vanhin lˆydetty tiedosto
        If Not oldestFile Is Nothing Then
            On Error Resume Next ' Ohita virheet v‰liaikaisesti poiston aikana
            FSO.DeleteFile oldestFile.Path, True ' True = pakota poisto (esim. jos read-only)
            If Err.Number <> 0 Then
                ' Virhe poistossa, ilmoita (valinnainen)
                MsgBox "Vanhimman varmuuskopion poistaminen ep‰onnistui:" & vbCrLf & oldestFile.Path & vbCrLf & Err.Description, vbExclamation, "Poistovirhe"
                Err.Clear
            'Else
                ' Poisto onnistui, voit lis‰t‰ lokimerkinn‰n tai Debug.Printin halutessasi
                'Debug.Print "Poistettu vanhin varmuuskopio: " & oldestFile.Path
            End If
            On Error GoTo VirheKasittelija ' Palauta normaali virheenk‰sittely
        Else
            ' T‰h‰n ei pit‰isi p‰‰ty‰, jos laskuri >= MAX_BACKUPS, mutta hyv‰ olla olemassa
             'Debug.Print "Varmuuskopioita oli " & currentBackupCount & " kpl, mutta vanhinta ei voitu tunnistaa poistettavaksi."
        End If
    End If

    ' --- Luo uusi varmuuskopio ---
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    fullDestPath = destFolderPath & "\" & baseName & "_" & timeStamp & "." & fileExt
    srcWb.SaveCopyAs fullDestPath


PuhdistusJaLopetus:
    ' Vapauta kaikki objektimuuttujat
    On Error Resume Next ' Varmista, ett‰ kaikki vapautetaan, vaikka jokin olisi jo Nothing
    Set FSO = Nothing
    Set WShell = Nothing
    Set backupFolderObj = Nothing
    Set fileItem = Nothing
    Set oldestFile = Nothing
    Set backupFiles = Nothing
    Set srcWb = Nothing
    On Error GoTo 0
    Exit Sub ' Poistu aliohjelmasta

VirheKasittelija:
    ' Keskitetty virheilmoitus k‰ytt‰j‰lle
    MsgBox "Varmuuskopioinnissa tapahtui virhe:" & vbCrLf & Err.Description, vbCritical, "Virhe"
    ' Siirry puhdistukseen virheen j‰lkeen
    Resume PuhdistusJaLopetus

End Sub

