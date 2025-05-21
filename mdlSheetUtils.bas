Attribute VB_Name = "mdlSheetUtils"
' --- Module: mdlSheetUtils ---
Option Explicit

' Etsii viimeisen käytetyn rivin numeron annetulta välilehdeltä ja sarakkeesta.
' Oletuksena tarkistaa sarakkeen A (1).
Public Function GetLastRow(ByVal ws As Worksheet, Optional ByVal columnToCheck As Variant = 1) As Long
    On Error Resume Next ' Jos välilehti on tyhjä tai suojattu
    If ws Is Nothing Then
        GetLastRow = 1 ' Tai jokin muu oletus/virhearvo
        Exit Function
    End If

    GetLastRow = ws.Cells(ws.rows.Count, columnToCheck).End(xlUp).row
    ' Jos sarake on täysin tyhjä, End(xlUp) voi palauttaa 1 (otsikkorivi).
    ' Varmistetaan, ettei palauteta 0 tai negatiivista virhetilanteessa.
    If GetLastRow <= 0 Then GetLastRow = 1
    If IsEmpty(ws.Cells(GetLastRow, columnToCheck).value) And GetLastRow > 1 Then
       ' Jos viimeiseksi löytynyt solu on tyhjä ja se ei ole rivi 1, End(xlUp) on ehkä osunut tyhjään kohtaan.
       ' Tarkempi tarkistus voisi olla tarpeen, mutta tämä perusversio riittää usein.
       ' Joskus käytetään myös Find-metodia tarkempaan etsintään.
    End If

    On Error GoTo 0
End Function

' Tyhjentää annetun alueen sisällön.
Public Sub ClearRangeContents(ByVal targetRange As Range)
    On Error Resume Next ' Jos alue on suojattu tms.
    If Not targetRange Is Nothing Then
        targetRange.ClearContents
    End If
    On Error GoTo 0
End Sub

' Tyhjentää annetun alueen muotoilut.
Public Sub ClearRangeFormats(ByVal targetRange As Range)
    On Error Resume Next
    If Not targetRange Is Nothing Then
        targetRange.ClearFormats
    End If
    On Error GoTo 0
End Sub

' Voit lisätä tänne myöhemmin esim. keskitetyt Protect/Unprotect-funktiot,
' jotka ottavat salasanan argumenttina.

