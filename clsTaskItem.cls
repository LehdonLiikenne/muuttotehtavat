VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Class Module: TaskItem ---
Option Explicit

' T�m� luokka edustaa yht� rivi� "Tietovarasto"-v�lilehdelt�.
' Ominaisuudet vastaavat sarakkeita A-AL (1-38).


Public asiakas As String
Public tarjousTehty As Variant ' K�yt� Variantia, jotta voi olla Date/Null/Empty
Public lastausPaiva As Variant
Public lastausMaa As String
Public purkuMaa As String
Public purkuPaiva As Variant
Public M3m As String          ' Myydyt kuutiot (String joustavuuden vuoksi)


Public palvelu As String
Public Huomioitavaa As String ' K�ytet��n my�s Huomioriveill�
Public puhelin As String
Public lastausOsoite As String
Public purkuOsoite As String
Public Kuljettajat As String  ' K�ytet��n my�s Huomioriveill�
Public Autot As String        ' K�ytet��n my�s Huomioriveill�
Public Kontit As String       ' K�ytet��n my�s Huomioriveill�
Public Apulaiset As String
Public Rahtikirja As String
Public ApulaisetTilattu As String ' Esim. "Ok", "Tarvitaan", "Ei tarvita"
Public Pysakointilupa As String   ' Esim. "Ok", "Tarvitaan", "Ei tarvita"
Public hissi As String            ' Esim. "Ok", "Tarvitaan", "Ei tarvita"
Public Laivalippu As String       ' Esim. "Ok", "Tarvitaan", "Ei tarvita"
Public Laskutus As String
Public Vakuutus As String
Public Arvo As Variant          ' K�yt� Variantia (voi olla luku/teksti/tyhj�)
Public hinta As Variant         ' K�yt� Variantia
Public Muuttomaailma As String
Public M3t As String            ' Toteutuneet kuutiot
Public LastauspaivaVarmistunut As Boolean
Public PurkupaivaVarmistunut As Boolean
Public TarjousHyvaksytty As Variant
Public TarjousHylatty As Variant
Public Tila As String           ' Esim. "Tarjous", "Hyv�ksytty"
Public LastausLoppuu As Variant
Public PurkuLoppuu As Variant
Public sahkoposti As String
Public kontaktiPaiva As Variant
Public valimatka As String

' --- Tunnisteet ja tyyppitiedot AJ-AL (36-38) ---
Public ID As Long               ' Sarake AJ (36): Uniikki tunniste riville
Public RecordType As String     ' Sarake AK (37): "Task" tai "Attention"
Public AttentionSortDate As Variant ' Sarake AL (38): Huomiorivin oma lajittelupvm

' --- Alustusmetodi ---
Public Sub InitDefaults()
    ' Asettaa oletusarvoja uusille olioille
    Me.RecordType = "Task"
    Me.LastauspaivaVarmistunut = False
    Me.PurkupaivaVarmistunut = False

    ' Esimerkkej� lis�alustuksista:
    Me.asiakas = vbNullString
    Me.tarjousTehty = Empty ' Tai Null, jos halutaan erottaa tyhj�st� p�iv�st�
    Me.lastausPaiva = Empty
    Me.lastausMaa = vbNullString
    Me.purkuMaa = vbNullString
    Me.purkuPaiva = Empty
    Me.M3m = vbNullString
    Me.palvelu = vbNullString
    Me.Huomioitavaa = vbNullString
    Me.puhelin = vbNullString
    Me.lastausOsoite = vbNullString
    Me.purkuOsoite = vbNullString
    Me.Kuljettajat = vbNullString
    Me.Autot = vbNullString
    Me.Kontit = vbNullString
    Me.Apulaiset = vbNullString
    Me.Rahtikirja = vbNullString ' Tai False, jos Boolean on parempi
    Me.ApulaisetTilattu = vbNullString ' Tai jokin oletus kuten "EI TARVITA"
    Me.Pysakointilupa = vbNullString
    Me.hissi = vbNullString
    Me.Laivalippu = vbNullString
    Me.Laskutus = vbNullString ' Tai False
    Me.Vakuutus = vbNullString
    Me.Arvo = Empty
    Me.hinta = Empty
    Me.Muuttomaailma = vbNullString ' Tai False
    Me.M3t = vbNullString
    Me.valimatka = vbNullString
    Me.TarjousHyvaksytty = Empty
    Me.TarjousHylatty = Empty
    Me.Tila = "TARJOUS" ' Tai vbNullString, jos ei oletustilaa
    Me.LastausLoppuu = Empty
    Me.PurkuLoppuu = Empty
    Me.sahkoposti = vbNullString
    Me.kontaktiPaiva = Empty
    Me.AttentionSortDate = Empty
End Sub
