Attribute VB_Name = "mdlDateUtils"

' --- Module: mdlDateUtils ---
Option Explicit

' --- ParseDateString_Force_DDMMYYYY - Apufunktio (Sama kuin Versiossa 8) ---
' J�sent�� merkkijonon olettaen dd.mm.yyyy TAI dd/mm/yyyy -muodon
Private Function ParseDateString_Force_DDMMYYYY(ByVal dateString As String) As Variant
    Dim Parts() As String
    Dim separator As String
    Dim d As Integer, m As Integer, y As Integer
    ParseDateString_Force_DDMMYYYY = Null ' Oletus

    On Error GoTo ParseErrorHandler

    dateString = Trim$(dateString)
    If dateString = "" Then Exit Function

    ' Tunnista erotin (piste tai kauttaviiva)
    If InStr(1, dateString, ".") > 0 Then
        separator = "."
    ElseIf InStr(1, dateString, "/") > 0 Then
        separator = "/"
    Else
        'Debug.Print "ParseDateString_Force_DDMMYYYY: No '.' or '/' separator found in '" & dateString & "'"
        Exit Function
    End If

    Parts = Split(dateString, separator)
    If UBound(Parts) <> 2 Then Debug.Print "ParseDateString_Force_DDMMYYYY: Found " & (UBound(Parts) + 1) & " parts using '" & separator & "', expected 3 for '" & dateString & "'.": Exit Function

    ' Poista mahdollinen kellonaika vuoden per�st�
    Dim yearPart As String
    yearPart = Trim$(Parts(2))
    If InStr(1, yearPart, " ") > 0 Then
        yearPart = Trim$(Split(yearPart, " ")(0))
    End If

    If Not IsNumeric(Parts(0)) Or Not IsNumeric(Parts(1)) Or Not IsNumeric(yearPart) Then Debug.Print "ParseDateString_Force_DDMMYYYY: Non-numeric parts found ('" & Parts(0) & "', '" & Parts(1) & "', '" & yearPart & "').": Exit Function

    ' --- Oleta DD / MM / YYYY j�rjestys ---
    d = CInt(Parts(0)) ' Ensimm�inen on P�IV�
    m = CInt(Parts(1)) ' Toinen on KUUKAUSI
    y = CInt(yearPart)

    If d < 1 Or d > 31 Or m < 1 Or m > 12 Then Debug.Print "ParseDateString_Force_DDMMYYYY: Day (" & d & ") or Month (" & m & ") out of basic range.": Exit Function

    If y >= 0 And y <= 99 Then
        If y >= 30 Then y = 1900 + y Else y = 2000 + y
    End If
    If y < 1900 Or y > 2200 Then Debug.Print "ParseDateString_Force_DDMMYYYY: Year (" & y & ") out of range (1900-2200).": Exit Function

    Dim testDate As Date
    testDate = DateSerial(y, m, d) ' K�yt� y, m, d

    If Year(testDate) <> y Or Month(testDate) <> m Or Day(testDate) <> d Then
         'Debug.Print "ParseDateString_Force_DDMMYYYY: DateSerial rolled over invalid date parts for input: '" & dateString & "'"
         Exit Function
    End If

    ParseDateString_Force_DDMMYYYY = testDate ' Palauta oikea Date-arvo

Exit Function
ParseErrorHandler:
    Debug.Print "ParseDateString_Force_DDMMYYYY: ERROR parsing string '" & dateString & "': " & Err.Description
    ParseDateString_Force_DDMMYYYY = Null
End Function


Public Function ConvertToDate(ByVal inputDate As Variant) As Variant
    On Error GoTo ConversionError
    ConvertToDate = Null ' Oletus

    ' --- TARKISTA NULL JA EMPTY ENSIN ---
    If IsNull(inputDate) Or IsEmpty(inputDate) Then
        ' Debug.Print "ConvertToDate: Input was Null or Empty. Returning Null." ' Poista tai kommentoi tarvittaessa
        Exit Function ' Palauta Null (oletusarvo) turvallisesti
    End If

    ' --- Muunna merkkijonoksi ja kutsu AINOASTAAN omaa parseria ---
    Dim inputString As String
    ' Seuraava rivi on riskialtis, jos inputDate voi olla esim. Error-tyyppi�,
    ' mutta CStr(Null) on jo k�sitelty yll�.
    On Error Resume Next ' Lis�t��n turva CStr-kutsulle, jos inputDate olisi esim. Error Variant
    inputString = Trim$(VBA.CStr(inputDate))
    If Err.Number <> 0 Then
        ' Debug.Print "ConvertToDate: CStr failed for input type '" & TypeName(inputDate) & "'. Input: '" & inputDate & "'. Error: " & Err.Description
        Err.Clear
        Exit Function ' Palauta Null (oletusarvo)
    End If
    On Error GoTo ConversionError ' Palauta varsinainen virheenk�sittelij�

    If inputString = "" Then
        ' Debug.Print "ConvertToDate: Input string is empty after CStr and Trim. Returning Null." ' Poista tai kommentoi
        Exit Function
    End If

    Dim parsedDate As Variant
    parsedDate = ParseDateString_Force_DDMMYYYY(inputString)

    ConvertToDate = parsedDate

    'If IsNull(ConvertToDate) Then
         'Debug.Print "ConvertToDate: ParseDateString_Force_DDMMYYYY failed for '" & inputString & "'. Returning Null."
    'End If

Exit Function
ConversionError:
    Dim errInputValue As String
    On Error Resume Next ' V�lt� virhett� CStr-kutsussa virheenk�sittelij�ss� itsess��n
    errInputValue = CStr(inputDate)
    On Error GoTo 0
    Debug.Print "ConvertToDate: GENERAL Error for input (Type: " & TypeName(inputDate) & ", Value: '" & errInputValue & "') - " & Err.Description
    ConvertToDate = Null ' Varmista, ett� palautetaan Null virhetilanteessa
End Function

' --- FormatDateToString & GetWeekNumberISO8601 - EI MUUTOKSIA ---
Public Function FormatDateToString(ByVal value As Variant, Optional ByVal defaultValue As String = "-") As String
    Dim convertedDate As Variant
    Dim formattedString As String
    formattedString = defaultValue
    convertedDate = ConvertToDate(value) ' K�ytt�� t�t� Versio 11:sta
    If Not IsNull(convertedDate) Then
         If VBA.IsDate(convertedDate) Then
             On Error Resume Next
             formattedString = VBA.Format$(convertedDate, "dd.mm.yyyy")
             If Err.Number <> 0 Then
                  formattedString = defaultValue
                  Err.Clear
             End If
             On Error GoTo 0
         End If
    End If
    FormatDateToString = formattedString
End Function

Public Function GetWeekNumberISO8601(ByVal dt As Date) As Integer
    On Error Resume Next
    GetWeekNumberISO8601 = CInt(Format$(dt, "ww", vbMonday, vbFirstFourDays))
    If Err.Number <> 0 Then GetWeekNumberISO8601 = 0
    On Error GoTo 0
End Function

' Palauttaa kuukauden nimen suomeksi ISOLLA KIRJAIMILLA
Public Function GetFinnishMonthName(ByVal inputDate As Date) As String
    On Error GoTo ErrorHandler
    Dim monthNum As Integer
    monthNum = Month(inputDate)

    Select Case monthNum
        Case 1: GetFinnishMonthName = "TAMMIKUU"
        Case 2: GetFinnishMonthName = "HELMIKUU"
        Case 3: GetFinnishMonthName = "MAALISKUU"
        Case 4: GetFinnishMonthName = "HUHTIKUU"
        Case 5: GetFinnishMonthName = "TOUKOKUU"
        Case 6: GetFinnishMonthName = "KES�KUU"
        Case 7: GetFinnishMonthName = "HEIN�KUU"
        Case 8: GetFinnishMonthName = "ELOKUU"
        Case 9: GetFinnishMonthName = "SYYSKUU"
        Case 10: GetFinnishMonthName = "LOKAKUU"
        Case 11: GetFinnishMonthName = "MARRASKUU"
        Case 12: GetFinnishMonthName = "JOULUKUU"
        Case Else: GetFinnishMonthName = "VIRHE"
    End Select
CleanExit:
    Exit Function
ErrorHandler:
    GetFinnishMonthName = "KK VIRHE"
    Resume CleanExit
End Function

' Laskee ISO 8601 -viikon maanantain p�iv�m��r�n
' Perustuu kaavaan: https://en.wikipedia.org/wiki/ISO_week_date#Calculating_the_date_from_the_week_number,_year_and_day_of_the_week
Public Function GetFirstDayOfWeekISO(ByVal isoYear As Integer, ByVal isoWeek As Integer) As Date
    On Error GoTo GetFirstDayError
    Dim jan4 As Date
    Dim firstMonday As Date
    Dim weekDayJan4 As Integer ' vbMonday = 1...7

    ' Etsi vuoden Tammikuun 4. p�iv�
    jan4 = DateSerial(isoYear, 1, 4)
    ' Hae sen viikonp�iv� (Maanantai = 1)
    weekDayJan4 = Weekday(jan4, vbMonday)

    ' Laske viikon 1 maanantai
    firstMonday = DateAdd("d", 1 - weekDayJan4, jan4)

    ' Lis�� viikkojen m��r� (v�hennettyn� yhdell�)
    GetFirstDayOfWeekISO = DateAdd("ww", isoWeek - 1, firstMonday)

    Exit Function
GetFirstDayError:
    Debug.Print "Virhe GetFirstDayOfWeekISO: Year=" & isoYear & ", Week=" & isoWeek & " - " & Err.Description
    GetFirstDayOfWeekISO = DateSerial(isoYear, 1, 1) ' Palauta jokin oletus virhetilanteessa
End Function

' Laskee ISO 8601 -viikon sunnuntain p�iv�m��r�n
Public Function GetLastDayOfWeekISO(ByVal isoYear As Integer, ByVal isoWeek As Integer) As Date
    On Error Resume Next ' Olettaa, ett� GetFirstDayOfWeekISO k�sittelee virheet
    GetLastDayOfWeekISO = DateAdd("d", 6, GetFirstDayOfWeekISO(isoYear, isoWeek))
    If Err.Number <> 0 Then
         Debug.Print "Virhe GetLastDayOfWeekISO: Year=" & isoYear & ", Week=" & isoWeek & " - " & Err.Description
         GetLastDayOfWeekISO = DateSerial(isoYear, 1, 7) ' Palauta jokin oletus
    End If
End Function
