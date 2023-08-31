Attribute VB_Name = "BasHoraires"
Option Explicit

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4

Declare Function ExitWindowsEx Lib "user32" _
        (ByVal uFlags As Long, _
        ByVal dwReserved As Long) As Long

'----------------------------------------------------------------
' But : Arreter l'ordinateur
' Entrées :
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Function ArreterWindows() As Long

    Call ExitWindowsEx(EWX_SHUTDOWN + EWX_FORCE, 0)

End Function

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
        .Couleurs.Fond = RGB(255, 255, 255)
        .Couleurs.Titre = RGB(32, 57, 148) ' RGB(28, 119, 29)
        .Couleurs.HeureCourante = RGB(100, 61, 125)
        .Couleurs.ProchaineVisite = RGB(243, 41, 31)
        .Couleurs.AutreVisite = RGB(0, 0, 0)
        .Couleurs.EtatVisite = RGB(30, 56, 140)
        
        ' Timer
        .IntervalleMAJHeure = GetSetting(App.Title, "Général", _
                                         "IntervalleMAJHeure", 50)
        .IntervalleClignote = GetSetting(App.Title, "Général", _
                                         "IntervalleClignote", 1000)
        
        ' Visites
        .DureeVisite = GetSetting(App.Title, "Général", "DuréeVisite", 6)
        .NbVisites = NB_VISITES
        
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
    
    For i = 1 To gOptions.NbVisites
       gTabHeureVisite(i) = DateAdd("n", Minute, gTabHeureVisite(i))
    Next
    
End Sub

'----------------------------------------------------------------
' But : Decale les numéros de visite de x
'       Si x est négatif, vérifie que le nouveau numéro n'est
'       pas inférieur à 1
' Entrées : Nombre
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub DecaleNoVisite(Nb As Long)
    Dim i As Long
    
    ' Vérifie que le nouveau numéro n'est pas inférieur à 1
    If gTabNoVisite(1) + Nb < 1 Then Exit Sub
    
    For i = 1 To gOptions.NbVisites
       gTabNoVisite(i) = gTabNoVisite(i) + Nb
    Next
    
End Sub

'----------------------------------------------------------------
' But : Decale les etats de visite (une visite vient de commencer)
' Entrées :
' Sorties :
' Suppositions :
' Effets de bord :
'----------------------------------------------------------------
Public Sub DecaleEtatVisite()
    Dim i As Long
    
    For i = 1 To gOptions.NbVisites - 1
       gTabEtatVisite(i) = gTabEtatVisite(i + 1)
    Next
    
    ' Etat par défaut de la dernière visite (nouvellement ajoutée)
    gTabEtatVisite(gOptions.NbVisites) = Ouvert
    
End Sub

'----------------------------------------------------------------
' But : Modifie l'état de visite selon le cycle suivant:
'
'      +--> ouvert --> complet --> fermé --+
'      |                                   |
'      +----<----------<----------<--------+
'
' Entrées : No de la visite
' Sorties :
' Suppositions :
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub ModifEtatVisite(NumVisite As Long)

    ' Vérifie la validité du numéro de visite
    If NumVisite < 1 Or NumVisite > gOptions.NbVisites Then Exit Sub
    
    gTabEtatVisite(NumVisite) = (gTabEtatVisite(NumVisite)) Mod 3 + 1
End Sub

'----------------------------------------------------------------
' But : Met à jour le tableau des visites (cas de la modif de la
'       durée des visites)
' Entrées :
' Sorties : tableau gTabHeureVisite mis à jour
' Suppositions :
'       Les paramètres généraux ont été chargés
' Effets de bord : IHM
'----------------------------------------------------------------
Public Sub MajTabHeureVisite()
    Dim i As Long
    Dim dTime As Date
    
    dTime = gTabHeureVisite(1)
    
    For i = 1 + 1 To gOptions.NbVisites
        gTabHeureVisite(i) = DateAdd("n", gOptions.DureeVisite * (i - 1), dTime)
    Next
    
End Sub

'----------------------------------------------------------------
' But : Renvoi le libellé d'un état de visite (ouvert / complet /fermé)
' Entrées : Etat
' Sorties : Libellé
' Suppositions :
' Effets de bord : IHM
'----------------------------------------------------------------
Public Function GetLibelleEtatVisite(etat As Etat_Enum) As String

    Select Case etat
    Case Complet
        GetLibelleEtatVisite = LIBELLE_COMPLET
    Case Ferme
        GetLibelleEtatVisite = LIBELLE_FERME
    Case Else
        GetLibelleEtatVisite = ""
    End Select
    
End Function

