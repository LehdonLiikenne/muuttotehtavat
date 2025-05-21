Attribute VB_Name = "mdlSortUtils"
' --- Standard Module: mdlSortUtils ---
Option Explicit

'===============================================================================
' Purpose: Tarjoaa apufunktioita taulukoiden lajitteluun, erityisesti
'          clsDisplayRow-olioiden lajitteluun SortDate-ominaisuuden perusteella.
'===============================================================================

' P‰‰funktio QuickSort-algoritmille clsDisplayRow-olioille
Public Sub QuickSortDisplayRowsBySortDate(ByRef rows() As clsDisplayRow, ByVal low As Long, ByVal high As Long)
    ' --- Lajittelee clsDisplayRow-taulukon osan (low -> high) paikallaan ---
    ' --- K‰ytt‰‰ SortDate-ominaisuutta lajitteluavaimena (nouseva) ---

    Dim pivotIndex As Long

    ' Jatketaan vain, jos alaraja on yl‰rajaa pienempi
    If low < high Then
        ' Jaetaan taulukko kahtia pivot-alkion ymp‰rilt‰
        pivotIndex = PartitionDisplayRows(rows, low, high)
        ' Lajitellaan rekursiivisesti vasen puoli (pienemm‰t)
        If low < pivotIndex - 1 Then QuickSortDisplayRowsBySortDate rows, low, pivotIndex - 1
        ' Lajitellaan rekursiivisesti oikea puoli (suuremmat)
        If pivotIndex + 1 < high Then QuickSortDisplayRowsBySortDate rows, pivotIndex + 1, high
    End If
End Sub

' QuickSortin jakofunktio (Partition)
Private Function PartitionDisplayRows(ByRef rows() As clsDisplayRow, ByVal low As Long, ByVal high As Long) As Long
    Dim pivotRow As clsDisplayRow
    Dim i As Long, j As Long
    Dim dateJ As Date, datePivot As Date

    ' Valitaan viimeinen alkio pivotiksi (yleinen tapa)
    Set pivotRow = rows(high)
    On Error Resume Next ' Ohita virhe, jos pivotRow.SortDate ei ole validi
    datePivot = pivotRow.SortDate
    If Err.Number <> 0 Then ' Jos pivotin p‰iv‰m‰‰r‰ on virheellinen, ei voida lajitella luotettavasti
        Debug.Print "Virheellinen pivot-p‰iv‰m‰‰r‰ QuickSortissa: " & pivotRow.SourceRecordID
        PartitionDisplayRows = high ' Palautetaan yl‰raja virheen merkiksi (ei optimaalinen, mutta est‰‰ kaatumisen)
        Exit Function
    End If
    On Error GoTo 0 ' Palauta normaali virheenk‰sittely

    i = low - 1 ' Indeksi alkioille, jotka ovat pivotia pienempi‰

    ' K‰yd‰‰n l‰pi alkiot alarajasta yl‰rajaan (pois lukien pivot itse)
    For j = low To high - 1
        ' Vertaa nykyist‰ alkiota (rows(j)) pivot-alkioon (pivotRow)
        On Error Resume Next ' Ohita virhe, jos rows(j).SortDate ei ole validi
        dateJ = rows(j).SortDate
        If Err.Number <> 0 Then ' Handle error reading current element's date
             Debug.Print "Virheellinen dateJ QuickSortissa: ID " & rows(j).SourceRecordID '<<< LISƒTTY DEBUG
             Err.Clear ' Clear error
        Else
            On Error GoTo 0 ' Palauta normaali virheenk‰sittely

            ' Suoritetaan vertailu ja mahdollinen vaihto
            If dateJ <= datePivot Then
                i = i + 1 ' Kasvatetaan pienempien alkioiden loppupisteen indeksi‰
                ' Vaihdetaan alkiot rows(i) ja rows(j) kesken‰‰n
                SwapDisplayRows rows, i, j

                ' --- Valinnainen: Toissijainen lajittelu ID:n mukaan, jos p‰iv‰m‰‰r‰t ovat samat ---
                ' T‰m‰ varmistaa, ett‰ saman p‰iv‰n rivit ovat AINA samassa j‰rjestyksess‰ (lajittelun vakaus)
                If dateJ = datePivot Then
                    ' Jos rows(i):n ID on suurempi kuin juuri vaihdetun rows(j):n ID (joka oli alunperin pivotia pienempi/yht‰suuri),
                    ' ja niiden pit‰isi olla ID-j‰rjestyksess‰, voidaan tarvita lis‰logiikkaa.
                    ' Yksinkertaisin tapa: Lajitellaan ID:n mukaan vain, jos dateJ < datePivot on ep‰tosi.
                    ' T‰ss‰ tapauksessa riitt‰‰ usein, ett‰ ID:t‰ ei k‰ytet‰, koska QuickSort ei ole vakaa.
                    ' Jos vakaus on t‰rke‰‰, MergeSort olisi parempi algoritmi. J‰tet‰‰n ID-vertailu pois t‰st‰.
                End If
            End If
         End If
         On Error GoTo 0 ' Varmista normaali virheenk‰sittely
    Next j

    ' Vaihdetaan pivot-alkio (joka on paikassa high) oikealle paikalleen (i + 1)
    SwapDisplayRows rows, i + 1, high

    PartitionDisplayRows = i + 1 ' Palauta pivot-alkion lopullinen indeksi
End Function

' Apufunktio kahden clsDisplayRow-olion paikan vaihtamiseksi taulukossa
Private Sub SwapDisplayRows(ByRef rows() As clsDisplayRow, ByVal index1 As Long, ByVal index2 As Long)
    ' Tarvitaan v‰liaikainen muuttuja vaihdon ajaksi
    Dim tempRow As clsDisplayRow
    ' Tarkistetaan indeksit varmuuden vuoksi (vaikka kutsujan pit‰isi hoitaa)
    If index1 = index2 Then Exit Sub
    On Error Resume Next ' Virheenk‰sittely, jos indeksit ovat virheellisi‰
    Set tempRow = rows(index1)
    Set rows(index1) = rows(index2)
    Set rows(index2) = tempRow
    On Error GoTo 0
End Sub
