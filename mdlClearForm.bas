Attribute VB_Name = "mdlClearForm"
' --- Standard Module: mdlClearForm ---
Option Explicit

' Tyhjent‰‰ annetun UserForm-olion (frm) kontrollit oletusarvoihin.
' K‰y l‰pi myˆs Frame-kontrollien sis‰ll‰ olevat kontrollit.
Public Sub ClearForm(ByVal frm As Object)
    ' Varmistetaan, ett‰ annettu objekti on UserForm
    If Not TypeOf frm Is UserForm Then
        Debug.Print "ClearForm: Annettu objekti ei ole UserForm."
        Exit Sub
    End If

    Dim ctrl As MSForms.Control ' K‰ytet‰‰n tarkempaa tyyppi‰
    Dim subCtrl As MSForms.Control

    On Error Resume Next ' Ohita virheet, jos kontrollia ei voi k‰sitell‰

    ' K‰yd‰‰n l‰pi kaikki lomakkeen p‰‰kontrollit
    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            Case "TextBox"
                ctrl.Text = ""
            Case "ComboBox"
                ' Asetetaan arvo tyhj‰ksi ja poistetaan valinta
                ctrl.value = ""
                ctrl.listIndex = -1
            Case "ListBox"
                ' Poistetaan valinta (monivalinnalle t‰m‰ ei riit‰, vaatisi silmukan)
                ctrl.listIndex = -1
                 ' Jos ListBox sallii monivalinnan (MultiSelect), pit‰‰ k‰yd‰ l‰pi ja poistaa valinnat:
                 If ctrl.MultiSelect <> fmMultiSelectSingle Then ' fmMultiSelectSingle = 0
                     Dim i As Long
                     For i = 0 To ctrl.ListCount - 1
                         ctrl.Selected(i) = False
                     Next i
                 End If
            Case "CheckBox"
                ctrl.value = False
            Case "OptionButton"
                ctrl.value = False
            Case "Frame"
                ' Jos kontrolli on Frame, k‰yd‰‰n sen sis‰ll‰ olevat kontrollit l‰pi
                For Each subCtrl In ctrl.Controls
                    Select Case TypeName(subCtrl)
                        Case "TextBox"
                            subCtrl.Text = ""
                        Case "ComboBox"
                            subCtrl.value = ""
                            subCtrl.listIndex = -1
                        Case "ListBox"
                             subCtrl.listIndex = -1
                             If subCtrl.MultiSelect <> fmMultiSelectSingle Then
                                 Dim j As Long
                                 For j = 0 To subCtrl.ListCount - 1
                                     subCtrl.Selected(j) = False
                                 Next j
                             End If
                        Case "CheckBox"
                            subCtrl.value = False
                        Case "OptionButton"
                            ' Huom: OptionButtoneita Framessa ei yleens‰ nollata yksitellen,
                            ' vaan asetetaan jokin niist‰ valituksi Initialize-vaiheessa.
                            ' T‰m‰ rivi poistaa yksitt‰isen napin valinnan, jos se oli p‰‰ll‰.
                             subCtrl.value = False
                    End Select
                Next subCtrl
            ' Lis‰‰ Case-lausekkeita muille kontrollityypeille tarvittaessa
            ' Case "ToggleButton"
            '    ctrl.Value = False
            ' Case "ScrollBar", "SpinButton"
            '    ctrl.Value = ctrl.Min ' Tai jokin muu oletusarvo
        End Select
    Next ctrl

    On Error GoTo 0 ' Palauta normaali virheenk‰sittely
End Sub

