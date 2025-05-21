VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDisplayRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' --- Class Module: DisplayRow ---
Option Explicit

Public Enum RowContentType
    TaskType = 1
    HeaderType = 2
    AttentionType = 3
    KontaktType = 4
End Enum

' --- Yksityinen muuttuja näyttösarakkeille ---
Private mDisplayColumns(1 To 102) As Variant ' Käytetään Private ja m-etuliitettä

' --- Julkiset ominaisuudet ---
Public RowType As RowContentType
Public SortDate As Date
Public SourceRecordID As Long
Public WeekHeaderText As String
Public IsLoadingDay As Boolean


' --- Alustusmetodi ---
Public Sub Init()
    Dim i As Long
    ' Alusta YKSITYINEN taulukko
    For i = LBound(mDisplayColumns) To UBound(mDisplayColumns)
        mDisplayColumns(i) = vbNullString
    Next i
    ' Alusta muut tarvittaessa
End Sub

' --- Metodi arvon asettamiseksi tiettyyn sarakkeeseen ---
Public Sub SetColumnValue(ByVal columnIndex As Long, ByVal value As Variant)
    ' Tarkista, että indeksi on sallitulla alueella
    If columnIndex >= LBound(mDisplayColumns) And columnIndex <= UBound(mDisplayColumns) Then
        mDisplayColumns(columnIndex) = value
    Else
        ' Voit halutessasi lisätä virheilmoituksen tai lokimerkinnän
        Debug.Print "Virheellinen columnIndex (" & columnIndex & ") SetColumnValue-metodissa."
    End If
End Sub

' --- Property Get -proseduuri koko taulukon lukemiseen ---
Public Property Get DisplayColumns() As Variant
    ' Palauttaa kopion yksityisestä taulukosta
    DisplayColumns = mDisplayColumns
End Property
