Attribute VB_Name = "mdlSheetUtils"
' --- Module: mdlSheetUtils ---
Option Explicit

' Etsii viimeisen k�ytetyn rivin numeron annetulta v�lilehdelt� ja sarakkeesta.
' Oletuksena tarkistaa sarakkeen A (1).
Public Function GetLastRow(ByVal ws As Worksheet, Optional ByVal columnToCheck As Variant = 1) As Long
    On Error Resume Next ' Jos v�lilehti on tyhj� tai suojattu
    If ws Is Nothing Then
        GetLastRow = 1 ' Tai jokin muu oletus/virhearvo
        Exit Function
    End If

    GetLastRow = ws.Cells(ws.rows.Count, columnToCheck).End(xlUp).row
    ' Jos sarake on t�ysin tyhj�, End(xlUp) voi palauttaa 1 (otsikkorivi).
    ' Varmistetaan, ettei palauteta 0 tai negatiivista virhetilanteessa.
    If GetLastRow <= 0 Then GetLastRow = 1
    If IsEmpty(ws.Cells(GetLastRow, columnToCheck).value) And GetLastRow > 1 Then
       ' Jos viimeiseksi l�ytynyt solu on tyhj� ja se ei ole rivi 1, End(xlUp) on ehk� osunut tyhj��n kohtaan.
       ' Tarkempi tarkistus voisi olla tarpeen, mutta t�m� perusversio riitt�� usein.
       ' Joskus k�ytet��n my�s Find-metodia tarkempaan etsint��n.
    End If

    On Error GoTo 0
End Function

' Tyhjent�� annetun alueen sis�ll�n.
Public Sub ClearRangeContents(ByVal targetRange As Range)
    On Error Resume Next ' Jos alue on suojattu tms.
    If Not targetRange Is Nothing Then
        targetRange.ClearContents
    End If
    On Error GoTo 0
End Sub

' Tyhjent�� annetun alueen muotoilut.
Public Sub ClearRangeFormats(ByVal targetRange As Range)
    On Error Resume Next
    If Not targetRange Is Nothing Then
        targetRange.ClearFormats
    End If
    On Error GoTo 0
End Sub

' Voit lis�t� t�nne my�hemmin esim. keskitetyt Protect/Unprotect-funktiot,
' jotka ottavat salasanan argumenttina.

