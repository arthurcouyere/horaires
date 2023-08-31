Attribute VB_Name = "BasHoraires"
Option Explicit

'----------------------------------------------------------------
' But : Transforme une chaine de couleur RGB au format "RRGGBB"
'       en un long représentant le code RGB
' Entrées : Chaine
' Sorties : Code couleur
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Function StringToRGB(CRGB As String) As Long
    
    On Error GoTo TraiteErreur
    
    StringToRGB = RGB(CInt("&H" & Mid(CRGB, 1, 2)), _
                      CInt("&H" & Mid(CRGB, 3, 2)), _
                      CInt("&H" & Mid(CRGB, 5, 2)))
                      
    Exit Function

TraiteErreur:
    StringToRGB = 0
    
End Function
'----------------------------------------------------------------
' But : Récupère les paramètres généraux de l'application
' Entrées :
' Sorties : variable gOptions mise à jour
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub GetOptions()

    With gOptions
    
        'Couleurs
        .Couleurs.Fond = StringToRGB(GetSetting(App.Title, "Couleurs", _
                            "Fond", "C0C0C0"))
        .Couleurs.Titre = StringToRGB(GetSetting(App.Title, "Couleurs", _
                            "Titre", "000000"))
        .Couleurs.HeureCourante = StringToRGB(GetSetting(App.Title, "Couleurs", _
                            "HeureCourante", "000000"))
        .Couleurs.ProchaineVisite = StringToRGB(GetSetting(App.Title, "Couleurs", _
                            "ProchaineVisite", "000000"))
        .Couleurs.AutreVisite = StringToRGB(GetSetting(App.Title, "Couleurs", _
                            "AutreVisite", "000000"))
        
        ' Timer
        .IntervalleMAJHeure = GetSetting(App.Title, "Général", _
                            "IntervalleMAJHeure", 50)
        
        ' Visites
        .DureeVisite = GetSetting(App.Title, "Général", "DuréeVisite", 6)
        .NbVisites = GetSetting(App.Title, "Général", "NbVisites", 10)
        
    End With
    
End Sub

'----------------------------------------------------------------
' But : Decale une visite de x minutes
'       Si x est négatif, vérifie que la nouvelle heure n'est pas
'       antérieure à l'heure courante
' Entrées : Minutes, optionel Booléen: Force (True) le décalage
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub DecaleVisite(Minute As Long, Optional Force As Boolean = False)
    Dim i As Long
    
    ' Vérifie que la nouvelle heure n'est pas antérieure à l'heure courante
    If DateAdd("n", Minute, gTabHeureVisite(1)) < Time And Not Force Then Exit Sub
    
    For i = LBound(gTabHeureVisite) To UBound(gTabHeureVisite)
       gTabHeureVisite(i) = DateAdd("n", Minute, gTabHeureVisite(i))
    Next
    
End Sub

