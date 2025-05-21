Attribute VB_Name = "mdlStringUtils"
' --- Module: mdlStringUtils ---
Option Explicit

' Palauttaa defaultValue (oletuksena "-"), jos value on Null, Empty tai tyhjä merkkijono.
' Muuten palauttaa value:n merkkijonona.
Public Function DefaultIfNull(ByVal value As Variant, Optional ByVal defaultValue As String = "-") As String
    If IsNull(value) Then
        DefaultIfNull = defaultValue
    ElseIf IsEmpty(value) Then
        DefaultIfNull = defaultValue
    ElseIf Trim$(CStr(value)) = "" Then
        DefaultIfNull = defaultValue
    Else
        DefaultIfNull = CStr(value)
    End If
End Function

' Lukee monivalinta-ListBoxin valinnat ja palauttaa ne erotinmerkillä eroteltuna merkkijonona.
Public Function GetListBoxMultiSelection(ByVal lst As MSForms.ListBox, Optional ByVal Delimiter As String = ";") As String
    Dim selectedItems As String
    Dim i As Long
    selectedItems = ""

    ' Tarkistukset
    If lst Is Nothing Then Exit Function
    If lst.ListCount = 0 Then Exit Function
    If lst.MultiSelect = fmMultiSelectSingle Then ' Varmista, että on monivalinta
        If lst.listIndex > -1 Then GetListBoxMultiSelection = lst.List(lst.listIndex) ' Palauta yksittäinen valinta
        Exit Function
    End If

    On Error Resume Next ' Virheiden varalta
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then ' Tarkista, onko kohde valittu
            selectedItems = selectedItems & lst.List(i) & Delimiter
        End If
    Next i
    On Error GoTo 0

    ' Poista viimeinen erotinmerkki, jos jotain valittiin
    If Len(selectedItems) > 0 Then
        selectedItems = Left$(selectedItems, Len(selectedItems) - Len(Delimiter))
    End If

    GetListBoxMultiSelection = selectedItems
End Function ' <<< TÄMÄ ON OIKEA LOPPU TÄLLE FUNKTIOLLE

' Asettaa monivalinta-ListBoxin valinnat erotinmerkillä erotellun merkkijonon perusteella.
Public Sub SetListBoxMultiSelection(ByVal lst As MSForms.ListBox, ByVal delimitedString As String, Optional ByVal Delimiter As String = ";")
    Dim selectedArray() As String
    Dim i As Long
    Dim j As Long
    Dim found As Boolean

    ' Tarkistukset
    If lst Is Nothing Then Exit Sub
    If lst.ListCount = 0 Then Exit Sub

    ' Tyhjennä ensin kaikki valinnat
    On Error Resume Next ' Ohita virheet tyhjennyksessä
    For i = 0 To lst.ListCount - 1
        lst.Selected(i) = False
    Next i
    On Error GoTo 0

    ' Jos syötemerkkijono on tyhjä, ei tehdä muuta
    If Trim$(delimitedString) = "" Then Exit Sub

    ' Hajota merkkijono taulukoksi erotinmerkin kohdalta
    selectedArray = Split(delimitedString, Delimiter)

    ' Käy läpi ListBoxin kohteet ja aseta valinta, jos löytyy taulukosta
    On Error Resume Next ' Ohita virheet valintaa asetettaessa
    For i = 0 To lst.ListCount - 1
        found = False ' Oletus: ei löydy
        ' Käy läpi taulukkoon hajotetut arvot
        For j = LBound(selectedArray) To UBound(selectedArray)
            ' Vertaa ListBoxin kohteen tekstiä taulukon arvoon (poista ylimääräiset välilyönnit)
            If Trim$(lst.List(i)) = Trim$(selectedArray(j)) Then
                found = True ' Vastaavuus löytyi
                Exit For ' Siirry seuraavaan ListBoxin kohteeseen
            End If
        Next j
        ' Aseta valinta, jos vastaavuus löytyi
        lst.Selected(i) = found
    Next i
    On Error GoTo 0
End Sub

' --- Palauttaa viikonpäivän nimen suomeksi ja isoilla kirjaimilla ---
Public Function GetFinnishWeekdayName(ByVal inputDate As Date) As String
    Dim dayNum As Integer

    On Error GoTo ErrorHandler ' Virheenkäsittely, jos inputDate ei ole validi

    ' Hae viikonpäivän numero (vbMonday = 1=Maanantai, 7=Sunnuntai)
    dayNum = Weekday(inputDate, vbMonday)

    ' Palauta oikea nimi Select Case -rakenteella
    Select Case dayNum
        Case 1: GetFinnishWeekdayName = "MA"
        Case 2: GetFinnishWeekdayName = "TI"
        Case 3: GetFinnishWeekdayName = "KE"
        Case 4: GetFinnishWeekdayName = "TO"
        Case 5: GetFinnishWeekdayName = "PE"
        Case 6: GetFinnishWeekdayName = "LA"
        Case 7: GetFinnishWeekdayName = "SU"
        Case Else: GetFinnishWeekdayName = "VIRHE" ' Jos Weekday palauttaa jotain outoa
    End Select

CleanExit:
    Exit Function

ErrorHandler:
    GetFinnishWeekdayName = "PVM VIRHE" ' Palauta virheilmoitus
    Resume CleanExit
End Function

Public Function IsotAlkukirjaimet(syoteTeksti As String) As String
  IsotAlkukirjaimet = StrConv(syoteTeksti, vbProperCase)
End Function
