Attribute VB_Name = "mdlConfigUtils"
' --- Module: mdlConfigUtils ---
Option Explicit

' Hakee seuraavan vapaan ID-numeron ja päivittää laskurin.
' Olettaa, että "Config"-välilehden solussa A1 on seuraava käytettävä ID.
Public Function GetNextID() As Long
    Dim ws As Worksheet
    Dim nextIDCell As Range
    Dim currentID As Long

    On Error GoTo ErrorHandler

    ' Yritä asettaa viittaus Config-välilehteen
    Set ws = Nothing ' Nollaa ensin
    On Error Resume Next ' Vaimenna virheet väliaikaisesti
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo ErrorHandler ' Palauta virheiden käsittely

    If ws Is Nothing Then
        MsgBox "Välilehteä 'Config' ei löytynyt! ID-numeron haku epäonnistui.", vbCritical, "Virhe"
        GetNextID = 0 ' Palauta 0 virheen merkiksi
        Exit Function
    End If

    Set nextIDCell = ws.Range("A1")

    ' Jos solu A1 on tyhjä tai ei-numeerinen, aloitetaan ID:stä 1
    If Not IsNumeric(nextIDCell.value) Or IsEmpty(nextIDCell.value) Then
        currentID = 1
        nextIDCell.value = currentID + 1 ' Aseta seuraava ID soluun
    Else
        currentID = CLng(nextIDCell.value) ' Ota nykyinen ID
        nextIDCell.value = currentID + 1    ' Päivitä seuraava ID soluun
    End If

    GetNextID = currentID ' Palauta käytettävä ID
    Exit Function

ErrorHandler:
    MsgBox "Virhe GetNextID-funktiossa: " & Err.Description, vbCritical, "Virhe"
    GetNextID = 0 ' Palauta 0 virheen merkiksi
End Function
