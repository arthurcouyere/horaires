Attribute VB_Name = "PutHoraires"
Option Explicit

'----------------------------------------------------------------
' But : Affiche l'heure courante
' Entr�es : Un contr�le Label
' Sorties : tableau gTabHeureVisite mis � jour
' Suppositions :
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub PutHeureCouranteDansLabel(lbl As Label)
    Dim cTime As String
    
    cTime = Format(Time, "hh:nn:ss")
    
    If lbl.Caption <> cTime Then lbl.Caption = cTime

End Sub

'----------------------------------------------------------------
' But : Affiche les horaires des visites
' Entr�es : gTabHeureVisite
' Sorties :
' Suppositions :
'       Le tableau gTabHeureVisite a �t� mis � jour
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub PutVisitesDansForm()
    Dim i As Long
    Dim cTime As String
    
    With frmHoraires
        For i = 1 To gOptions.NbVisites
            cTime = Format(gTabHeureVisite(i), "hh:nn")
            .lblHeureVisite(i).Caption = cTime
            .lblNoVisite(i).Caption = CStr(gTabNoVisite(i))
        Next
    End With
    
End Sub
