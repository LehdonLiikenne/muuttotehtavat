Attribute VB_Name = "mdlConfigUtils"
' --- Module: mdlConfigUtils ---
Option Explicit

' Hakee seuraavan vapaan ID-numeron ja p�ivitt�� laskurin.
' Olettaa, ett� "Config"-v�lilehden solussa A1 on seuraava k�ytett�v� ID.
Public Function GetNextID() As Long
    Dim ws As Worksheet
    Dim nextIDCell As Range
    Dim currentID As Long

    On Error GoTo ErrorHandler

    ' Yrit� asettaa viittaus Config-v�lilehteen
    Set ws = Nothing ' Nollaa ensin
    On Error Resume Next ' Vaimenna virheet v�liaikaisesti
    Set ws = ThisWorkbook.Worksheets("Config")
    On Error GoTo ErrorHandler ' Palauta virheiden k�sittely

    If ws Is Nothing Then
        MsgBox "V�lilehte� 'Config' ei l�ytynyt! ID-numeron haku ep�onnistui.", vbCritical, "Virhe"
        GetNextID = 0 ' Palauta 0 virheen merkiksi
        Exit Function
    End If

    Set nextIDCell = ws.Range("A1")

    ' Jos solu A1 on tyhj� tai ei-numeerinen, aloitetaan ID:st� 1
    If Not IsNumeric(nextIDCell.value) Or IsEmpty(nextIDCell.value) Then
        currentID = 1
        nextIDCell.value = currentID + 1 ' Aseta seuraava ID soluun
    Else
        currentID = CLng(nextIDCell.value) ' Ota nykyinen ID
        nextIDCell.value = currentID + 1    ' P�ivit� seuraava ID soluun
    End If

    GetNextID = currentID ' Palauta k�ytett�v� ID
    Exit Function

ErrorHandler:
    MsgBox "Virhe GetNextID-funktiossa: " & Err.Description, vbCritical, "Virhe"
    GetNextID = 0 ' Palauta 0 virheen merkiksi
End Function
